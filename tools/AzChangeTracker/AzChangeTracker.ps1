#Requires -Version 5.1
<#
.SYNOPSIS
    Azure Change Tracker - Track Azure resource changes across subscriptions.
.DESCRIPTION
    PowerShell/WPF tool that snapshots Azure subscription resource state
    as JSON, diffs snapshots to detect adds/removes/modifications,
    classifies changes by severity, and exports reports.
    Modeled after WinGetManifestManager and AIBPipelineCreator.
.NOTES
    Author : Anton Romanyuk
    Version: 0.1.0-alpha
    Date   : 2026-03-19
#>

# ===============================================================================
# SECTION 1: PRE-LOAD & INITIALIZATION
# Loads WPF assemblies, sets DPI awareness, creates storage directories,
# strips OneDrive module paths to prevent Az version conflicts, and
# initialises named constants for timers, toast durations, and page sizes.
# ===============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'


# Strip OneDrive user-profile module path to prevent old Az.Accounts versions
# (2.2.6, 2.1.0, 1.9.5) from loading before the current v5.x in Program Files.
# Must happen before any Az cmdlet is invoked to avoid assembly conflicts.
$env:PSModulePath = ($env:PSModulePath -split ';' |
    Where-Object { $_ -notlike '*OneDrive*' }) -join ';'
$Global:Root = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Global:Root)) { $Global:Root = $PWD.Path }

$Global:AppVersion   = "0.1.0-alpha"
$Global:AppTitle     = "Azure Change Tracker v$($Global:AppVersion)"
$Global:PrefsPath    = Join-Path $Global:Root "user_prefs.json"
$Global:AchievementsFile = Join-Path $Global:Root "achievements.json"
$Global:SnapshotDir  = Join-Path $Global:Root "snapshots"
$Global:DiffDir      = Join-Path $Global:Root "diffs"
$Global:ReportDir    = Join-Path $Global:Root "reports"
$Global:BaselineDir  = Join-Path $Global:Root "baselines"

# Ensure storage directories exist
foreach ($dir in @($Global:SnapshotDir, $Global:DiffDir, $Global:ReportDir, $Global:BaselineDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
        Write-Host "[INIT] Created directory: $dir" -ForegroundColor DarkGray
    }
}

# Named constants
$Script:LOG_MAX_LINES       = 500
$Script:TIMER_INTERVAL_MS   = 50
$Script:TOAST_DURATION_MS   = 4000
$Script:CONFETTI_COUNT      = 60
$Script:CLEANUP_DELAY_MS    = 5000
$Script:GRAPH_PAGE_SIZE     = 1000

# Load WPF assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName System.Windows.Forms

# DPI Awareness
try {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class DpiHelper {
    [DllImport("shcore.dll")]
    public static extern int SetProcessDpiAwareness(int value);
}
"@ -ErrorAction SilentlyContinue
    [DpiHelper]::SetProcessDpiAwareness(2) | Out-Null
    Write-Host "[INIT] DPI awareness set to PerMonitorV2" -ForegroundColor DarkGray
} catch {
    Write-Host "[INIT] DPI awareness already set or OS doesn't support" -ForegroundColor DarkGray
}

$Global:CachedBC = [System.Windows.Media.BrushConverter]::new()
Write-Host "[INIT] WPF assemblies loaded, BrushConverter cached" -ForegroundColor DarkGray

# ===============================================================================
# SECTION 2: PREREQUISITE MODULE CHECK
# ===============================================================================

$RequiredModules = @(
    @{ Name = 'Az.Accounts';       MinVersion = '2.7.5';  Purpose = 'Azure authentication (Connect-AzAccount)' }
    @{ Name = 'Az.ResourceGraph';  MinVersion = '0.13.0'; Purpose = 'Fast resource queries (Search-AzGraph)' }
    @{ Name = 'Az.Resources';      MinVersion = '6.0.0';  Purpose = 'Fallback resource queries, RBAC' }
)

    <#
    .SYNOPSIS
        Validates that all required PowerShell modules are installed.
    .DESCRIPTION
        Iterates through $RequiredModules and checks whether each module is available
        locally using Get-Module -ListAvailable. Returns an array of result hashtables
        with Name, Required version, Purpose, Installed version, and Available boolean.
        The splash screen and debug log are updated with the status of each module.
    .OUTPUTS
        [hashtable[]] One entry per required module with keys: Name, Required, Purpose,
        Installed, Available.
    #>
function Test-Prerequisites {
    Write-Host "[PREREQ] Checking $($RequiredModules.Count) required modules..." -ForegroundColor DarkGray
    $Results = @()
    foreach ($Mod in $RequiredModules) {
        $Installed = Get-Module -ListAvailable -Name $Mod.Name -ErrorAction SilentlyContinue |
                     Sort-Object Version -Descending | Select-Object -First 1
        $Status = if ($Installed) { "v$($Installed.Version) OK" } else { "MISSING" }
        Write-Host "[PREREQ] $($Mod.Name) >= $($Mod.MinVersion): $Status" -ForegroundColor DarkGray
        $Results += @{
            Name      = $Mod.Name
            Required  = $Mod.MinVersion
            Purpose   = $Mod.Purpose
            Installed = if ($Installed) { $Installed.Version.ToString() } else { $null }
            Available = [bool]$Installed
        }
    }
    return $Results
}

    <#
    .SYNOPSIS
        Generates an Install-Module command string for any missing prerequisite modules.
    .DESCRIPTION
        Filters the module check results for entries where Available is $false, then
        builds a single Install-Module command that installs all missing modules at once
        with -Scope CurrentUser -Force -AllowClobber.
    .PARAMETER ModuleResults
        Array of result hashtables returned by Test-Prerequisites.
    .OUTPUTS
        [string] A ready-to-run Install-Module command, or empty string if all modules
        are present.
    #>
function Get-RemediationCommand {
    param([array]$ModuleResults)
    $Missing = $ModuleResults | Where-Object { -not $_.Available }
    if ($Missing.Count -eq 0) { return "" }
    $Names = ($Missing | ForEach-Object { $_.Name }) -join "', '"
    return "Install-Module -Name '$Names' -Scope CurrentUser -Force -AllowClobber"
}

# ===============================================================================
# SECTION 3: THREAD SYNCHRONIZATION BRIDGE
# Provides Start-BackgroundWork which runs scriptblocks in STA runspaces
# and delivers results back to the UI thread via a DispatcherTimer poll
# loop (Section 20). All Azure API calls use this to keep the UI responsive.
# ===============================================================================

$Global:SyncHash = [Hashtable]::Synchronized(@{
    StatusQueue  = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())
    LogQueue     = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())
    StopFlag     = $false
})

$Global:BgJobs          = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
$Global:TimerProcessing = $false

    <#
    .SYNOPSIS
        Launches a PowerShell scriptblock in a background runspace for non-blocking execution.
    .DESCRIPTION
        Creates a new STA-mode runspace with clean PSModulePath (OneDrive paths stripped),
        injects the supplied variables, and starts the scriptblock asynchronously. The
        completion callback (OnComplete) runs on the UI thread when the DispatcherTimer
        detects that the async operation has finished. This pattern keeps the WPF UI
        responsive during long-running Azure operations (Connect-AzAccount, Search-AzGraph).
    .PARAMETER Work
        The scriptblock to execute in the background runspace. Has access to all
        variables injected via the Variables parameter.
    .PARAMETER OnComplete
        Scriptblock invoked on the UI thread when Work completes. Receives two parameters:
        $Results (output objects) and $Errors (ErrorRecord objects from the runspace).
    .PARAMETER Variables
        Hashtable of variables to inject into the runspace session state. Keys become
        variable names; values are serialised across the runspace boundary.
    .PARAMETER Context
        Optional hashtable of context data passed through to OnComplete for correlation.
    #>
function Start-BackgroundWork {
    param(
        [ScriptBlock]$Work,
        [ScriptBlock]$OnComplete,
        [hashtable]$Variables = @{},
        [hashtable]$Context   = @{}
    )
    $ISS = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $ISS.ExecutionPolicy = [Microsoft.PowerShell.ExecutionPolicy]::Bypass
    $RS  = [RunspaceFactory]::CreateRunspace($ISS)
    $RS.ApartmentState = 'STA'
    $RS.ThreadOptions  = 'ReuseThread'
    $RS.Open()

    # Strip OneDrive user-profile module path to avoid old Az version conflicts
    $CleanModulePath = ($env:PSModulePath -split ';' |
        Where-Object { $_ -notlike '*OneDrive*' }) -join ';'

    $PS = [PowerShell]::Create()
    $PS.Runspace = $RS
    foreach ($k in $Variables.Keys) {
        $PS.Runspace.SessionStateProxy.SetVariable($k, $Variables[$k])
    }
    $PS.Runspace.SessionStateProxy.SetVariable('__CleanModulePath', $CleanModulePath)

    # Prepend PSModulePath cleanup before user Work script
    $CombinedScript = '$env:PSModulePath = $__CleanModulePath' + "`n" + $Work.ToString()
    $PS.AddScript($CombinedScript) | Out-Null

    Write-DebugLog "BgWork: launching runspace (vars=$($Variables.Keys -join ','))" -Level 'DEBUG'
    $Async = $PS.BeginInvoke()
    $Global:BgJobs.Add(@{
        PS          = $PS
        Runspace    = $RS
        AsyncResult = $Async
        OnComplete  = $OnComplete
        Context     = $Context
        StartedAt   = (Get-Date)
    }) | Out-Null
    Write-DebugLog "BgWork: queued job #$($Global:BgJobs.Count)" -Level 'DEBUG'
}

# ===============================================================================
# SECTION 4: XAML GUI LOAD
# ===============================================================================

$XamlPath = Join-Path $Global:Root "AzChangeTracker_UI.xaml"
if (-not (Test-Path $XamlPath)) {
    [System.Windows.MessageBox]::Show(
        "CRITICAL: 'AzChangeTracker_UI.xaml' not found in:`n$Global:Root",
        "Azure Change Tracker", 'OK', 'Error') | Out-Null
    exit 1
}

$XamlContent = Get-Content $XamlPath -Raw -Encoding UTF8
Write-Host "[XAML] Loaded XAML: $([math]::Round($XamlContent.Length / 1024, 1)) KB" -ForegroundColor DarkGray
$XamlContent = $XamlContent -replace 'Title="Azure Change Tracker"', "Title=`"$Global:AppTitle`""

try {
    $Window = [Windows.Markup.XamlReader]::Parse($XamlContent)
    Write-Host "[XAML] Window parsed successfully" -ForegroundColor DarkGray
} catch {
    [System.Windows.MessageBox]::Show(
        "XAML Parse Error:`n$($_.Exception.Message)",
        "Azure Change Tracker", 'OK', 'Error') | Out-Null
    exit 1
}

# ===============================================================================
# SECTION 5: ELEMENT REFERENCES
# Binds ~80 named XAML elements to PowerShell variables using FindName().
# Grouped by UI area: title bar, auth bar, icon rail, sidebar panels,
# splash overlay, toolbar, tab headers/panels, dashboard, changes,
# timeline, resources, settings, debug log, status bar, confetti/toast.
# ===============================================================================

# Title bar
$btnThemeToggle    = $Window.FindName("btnThemeToggle")
$btnMinimize       = $Window.FindName("btnMinimize")
$btnMaximize       = $Window.FindName("btnMaximize")
$btnClose          = $Window.FindName("btnClose")
$btnHelp           = $Window.FindName("btnHelp")
$lblTitle          = $Window.FindName("lblTitle")
$lblTitleVersion   = $Window.FindName("lblTitleVersion")

# Auth bar
$authDot           = $Window.FindName("authDot")
$lblAuthStatus     = $Window.FindName("lblAuthStatus")
$cmbSubscription   = $Window.FindName("cmbSubscription")
$btnLogin          = $Window.FindName("btnLogin")
$btnLogout         = $Window.FindName("btnLogout")

# Icon rail
$btnHamburger      = $Window.FindName("btnHamburger")
$railDashboard     = $Window.FindName("railDashboard")
$railChanges       = $Window.FindName("railChanges")
$railTimeline      = $Window.FindName("railTimeline")
$railResources     = $Window.FindName("railResources")
$railSettings      = $Window.FindName("railSettings")
$railOutput        = $Window.FindName("railOutput")
$railDashboardIndicator  = $Window.FindName("railDashboardIndicator")
$railChangesIndicator    = $Window.FindName("railChangesIndicator")
$railTimelineIndicator   = $Window.FindName("railTimelineIndicator")
$railResourcesIndicator  = $Window.FindName("railResourcesIndicator")
$railSettingsIndicator   = $Window.FindName("railSettingsIndicator")

# Sidebar panels
$colLeftPanel      = $Window.FindName("colLeftPanel")
$pnlSidebar        = $Window.FindName("pnlSidebar")
$pnlSidebarDashboard  = $Window.FindName("pnlSidebarDashboard")
$pnlSidebarChanges    = $Window.FindName("pnlSidebarChanges")
$pnlSidebarTimeline   = $Window.FindName("pnlSidebarTimeline")
$pnlSidebarResources  = $Window.FindName("pnlSidebarResources")
$pnlSidebarSettings   = $Window.FindName("pnlSidebarSettings")

# Sidebar - Dashboard context
$lblSidebarResourceCount = $Window.FindName("lblSidebarResourceCount")
$lblSidebarLastSnapshot  = $Window.FindName("lblSidebarLastSnapshot")
$lblSidebarSnapshotCount = $Window.FindName("lblSidebarSnapshotCount")
$lstSnapshots            = $Window.FindName("lstSnapshots")

# Sidebar - Changes context
$cmbDiffFrom       = $Window.FindName("cmbDiffFrom")
$cmbDiffTo         = $Window.FindName("cmbDiffTo")
$btnRunCompare     = $Window.FindName("btnRunCompare")
$pnlResourceTypeFilters = $Window.FindName("pnlResourceTypeFilters")
$cmbFilterChangeType   = $Window.FindName("cmbFilterChangeType")
$txtFilterSearch       = $Window.FindName("txtFilterSearch")
$pnlHideVolatileToggle = $Window.FindName("pnlHideVolatileToggle")
$chkHideVolatileTrack  = $Window.FindName("chkHideVolatileTrack")
$chkHideVolatileThumb  = $Window.FindName("chkHideVolatileThumb")
$Global:HideVolatileEnabled = $true   # ON by default

# Sidebar - Timeline context
$txtTimelineSearch = $Window.FindName("txtTimelineSearch")
$cmbTimelineType   = $Window.FindName("cmbTimelineType")

# Sidebar - Resources context
$cmbResourceType   = $Window.FindName("cmbResourceType")
$cmbResourceGroup  = $Window.FindName("cmbResourceGroup")
$txtResourceSearch = $Window.FindName("txtResourceSearch")
$lblResourceInventoryCount = $Window.FindName("lblResourceInventoryCount")

# Sidebar - Achievements
$lblAchievementCount = $Window.FindName("lblAchievementCount")
$pnlAchievements     = $Window.FindName("pnlAchievements")

# Splash overlay
$pnlSplash         = $Window.FindName("pnlSplash")
$lblSplashVersion  = $Window.FindName("lblSplashVersion")
$lblSplashStatus   = $Window.FindName("lblSplashStatus")
$dotStep1          = $Window.FindName("dotStep1")
$dotStep2          = $Window.FindName("dotStep2")
$dotStep3          = $Window.FindName("dotStep3")
$lblStep1          = $Window.FindName("lblStep1")
$lblStep2          = $Window.FindName("lblStep2")
$lblStep3          = $Window.FindName("lblStep3")

# Background layers
$bdrDotGrid        = $Window.FindName("bdrDotGrid")
$bdrGradientGlow   = $Window.FindName("bdrGradientGlow")
$bdrShimmer        = $Window.FindName("bdrShimmer")

# Toolbar
$btnSnapshotNow    = $Window.FindName("btnSnapshotNow")
$btnExportReport   = $Window.FindName("btnExportReport")
$btnSaveBaseline   = $Window.FindName("btnSaveBaseline")

# Tab headers
$tabDashboard      = $Window.FindName("tabDashboard")
$tabChanges        = $Window.FindName("tabChanges")
$tabTimeline       = $Window.FindName("tabTimeline")
$tabResources      = $Window.FindName("tabResources")
$tabSettings       = $Window.FindName("tabSettings")

# Tab content panels
$pnlTabDashboard   = $Window.FindName("pnlTabDashboard")
$pnlTabChanges     = $Window.FindName("pnlTabChanges")
$pnlTabTimeline    = $Window.FindName("pnlTabTimeline")
$pnlTabResources   = $Window.FindName("pnlTabResources")
$pnlTabSettings    = $Window.FindName("pnlTabSettings")

# Dashboard panel
$pnlDashboardCards     = $Window.FindName("pnlDashboardCards")
$lblDashTotalResources = $Window.FindName("lblDashTotalResources")
$lblDashChanges        = $Window.FindName("lblDashChanges")
$lblDashAdded          = $Window.FindName("lblDashAdded")
$lblDashRemoved        = $Window.FindName("lblDashRemoved")
$pnlSeverityBars       = $Window.FindName("pnlSeverityBars")
$pnlRecentChanges      = $Window.FindName("pnlRecentChanges")

# Changes panel
$lblDiffHeader     = $Window.FindName("lblDiffHeader")
$badgeAdded        = $Window.FindName("badgeAdded")
$badgeRemoved      = $Window.FindName("badgeRemoved")
$badgeModified     = $Window.FindName("badgeModified")
$lblBadgeAdded     = $Window.FindName("lblBadgeAdded")
$lblBadgeRemoved   = $Window.FindName("lblBadgeRemoved")
$lblBadgeModified  = $Window.FindName("lblBadgeModified")
$lstChanges        = $Window.FindName("lstChanges")

# Timeline
$pnlTimeline       = $Window.FindName("pnlTimeline")

# Resources
$treeResources      = $Window.FindName("treeResources")

# Settings
$chkDarkMode       = $Window.FindName("chkDarkMode")
$chkAutoSnapshot   = $Window.FindName("chkAutoSnapshot")
$txtRetentionDays  = $Window.FindName("txtRetentionDays")
$chkRetentionEnabled = $Window.FindName("chkRetentionEnabled")
$pnlRetentionDays   = $Window.FindName("pnlRetentionDays")
$txtExcludedTypes  = $Window.FindName("txtExcludedTypes")

# Debug log
$rtbActivityLog    = $Window.FindName("rtbActivityLog")
$docActivityLog    = $Window.FindName("docActivityLog")
$paraLog           = $Window.FindName("paraLog")
$logScroller       = $Window.FindName("logScroller")
$btnClearLog       = $Window.FindName("btnClearLog")
$btnToggleBottomSize = $Window.FindName("btnToggleBottomSize")
$icoToggleBottomSize = $Window.FindName("icoToggleBottomSize")
$btnHideBottom     = $Window.FindName("btnHideBottom")
$rowBottomPanel    = $Window.FindName("rowBottomPanel")

# Status bar
$statusDot         = $Window.FindName("statusDot")
$lblStatus         = $Window.FindName("lblStatus")
$lblStatusResourceCount = $Window.FindName("lblStatusResourceCount")
$lblStatusSub      = $Window.FindName("lblStatusSub")

# Confetti & Toast
$cnvConfetti       = $Window.FindName("cnvConfetti")
$pnlToastHost      = $Window.FindName("pnlToastHost")

# Verify critical element bindings
$_criticalElements = @(
    @('Window',           $Window),           @('btnLogin',         $btnLogin)
    @('btnClose',         $btnClose),          @('cmbSubscription',  $cmbSubscription)
    @('paraLog',          $paraLog),           @('pnlToastHost',     $pnlToastHost)
    @('bdrDotGrid',       $bdrDotGrid),        @('statusDot',        $statusDot)
    @('btnSnapshotNow',   $btnSnapshotNow),    @('lstChanges',       $lstChanges)
)
$_nullCount = 0
foreach ($_el in $_criticalElements) {
    if ($null -eq $_el[1]) { Write-Host "[XAML] WARNING: Element '$($_el[0])' is NULL" -ForegroundColor Yellow; $_nullCount++ }
}
Write-Host "[XAML] Element binding complete: $_nullCount null references" -ForegroundColor DarkGray

# ===============================================================================
# SECTION 6: GLOBAL STATE
# Runtime variables: theme mode, debug log counters, active tab/rail,
# sidebar/bottom panel visibility, subscription list, current subscription,
# latest snapshot/diff caches, and achievement data. All prefixed $Global:
# for visibility from background-work OnComplete callbacks.
# ===============================================================================

$Global:IsLightMode       = $false
$Global:DebugLogFile      = Join-Path $env:TEMP "AzChangeTracker_debug.log"
$Global:DebugLineCount    = 0
$Global:DebugMaxLines     = $Script:LOG_MAX_LINES
$Global:DebugOverlayEnabled = $false
$Global:FullLogSB         = [System.Text.StringBuilder]::new()
$Global:FullLogMaxLines   = 1000
$Global:FullLogLineCount  = 0
Write-Host "[INIT] Disk log: $Global:DebugLogFile" -ForegroundColor DarkGray
$Global:ActiveTabName     = 'Dashboard'
$Global:ActiveRailName    = 'Dashboard'
$Global:BottomExpanded    = $true
$Global:BottomSavedHeight = 160
$Global:LeftPanelVisible  = $true

# Runtime data
$Global:Subscriptions     = @()
$Global:CurrentSubId      = ''
$Global:CurrentSubName    = ''
$Global:LatestSnapshot    = $null
$Global:LatestDiff        = $null
$Global:SnapshotList      = @()
$Global:Achievements      = @{}

# ===============================================================================
# SECTION 7: WRITE-DEBUGLOG (3-destination output)
# ===============================================================================

    <#
    .SYNOPSIS
        Writes a timestamped log message to three destinations: console, disk, and UI.
    .DESCRIPTION
        Every log entry is written to:
          1. PowerShell console (Write-Host, always)
          2. Disk log file ($env:TEMP\AzChangeTracker_debug.log, auto-rotates at 2 MB)
          3. WPF RichTextBox activity log panel (colour-coded by level)
        A ring buffer (1000 lines) in memory ($Global:FullLogSB) holds the full history.
        DEBUG-level messages are suppressed from the UI unless $Global:DebugOverlayEnabled
        is true. Log format: [HH:mm:ss.fff] [LEVEL] Message.
    .PARAMETER Message
        The log message text. Supports string interpolation.
    .PARAMETER Level
        Severity level: INFO, DEBUG, WARN, ERROR, or SUCCESS. Controls colour coding
        in both light and dark themes.
    #>
function Write-DebugLog {
    param([string]$Message, [string]$Level = 'INFO')

    $ts   = (Get-Date).ToString('HH:mm:ss.fff')
    $line = "[$ts] [$Level] $Message"

    # 1. Console output (always)
    Write-Host $line -ForegroundColor DarkGray

    # 2. Disk log (always, rotate at 2 MB)
    try {
        if ((Test-Path $Global:DebugLogFile) -and (Get-Item $Global:DebugLogFile).Length -gt 2MB) {
            $Rotated = $Global:DebugLogFile + '.old'
            if (Test-Path $Rotated) { Remove-Item $Rotated -Force }
            Rename-Item $Global:DebugLogFile $Rotated -Force
            Write-Host "[LOG] Rotated disk log (>2 MB)" -ForegroundColor DarkGray
        }
        [System.IO.File]::AppendAllText($Global:DebugLogFile, $line + "`r`n")
    } catch { }

    # 2b. Full log ring buffer (memory, 1000 lines)
    $Global:FullLogSB.AppendLine($line) | Out-Null
    $Global:FullLogLineCount++
    if ($Global:FullLogLineCount -gt $Global:FullLogMaxLines) {
        $idx = $Global:FullLogSB.ToString().IndexOf("`n")
        if ($idx -ge 0) { $Global:FullLogSB.Remove(0, $idx + 1) | Out-Null }
        $Global:FullLogLineCount--
    }

    # 3. RichTextBox (UI) - skip DEBUG level unless overlay enabled
    if ($Level -eq 'DEBUG' -and -not $Global:DebugOverlayEnabled) { return }
    if ($null -eq $paraLog) { return }

    $Color = switch ($Level) {
        'ERROR'   { if ($Global:IsLightMode) { '#CC0000' } else { '#FF4040' } }
        'WARN'    { if ($Global:IsLightMode) { '#B86E00' } else { '#FF9100' } }
        'SUCCESS' { if ($Global:IsLightMode) { '#008A2E' } else { '#16C60C' } }
        'DEBUG'   { if ($Global:IsLightMode) { '#8888AA' } else { '#B8860B' } }
        default   { if ($Global:IsLightMode) { '#444444' } else { '#888888' } }
    }

    try {
        $Run = New-Object System.Windows.Documents.Run($line)
        $Run.Foreground = $Global:CachedBC.ConvertFromString($Color)
        if ($paraLog.Inlines.Count -gt 0) {
            $paraLog.Inlines.Add([System.Windows.Documents.LineBreak]::new())
        }
        $paraLog.Inlines.Add($Run)

        # Ring buffer cleanup
        $Global:DebugLineCount++
        if ($Global:DebugLineCount -gt $Global:DebugMaxLines) {
            $paraLog.Inlines.Remove($paraLog.Inlines.FirstInline)  # old Run
            if ($paraLog.Inlines.Count -gt 0) {
                $paraLog.Inlines.Remove($paraLog.Inlines.FirstInline)  # old LineBreak
            }
            $Global:DebugLineCount--
        }
        $logScroller.ScrollToEnd()
    } catch { }
}

# ===============================================================================
# SECTION 8: THEME SYSTEM
# ===============================================================================

$Global:ThemeDark = @{
    ThemeAppBg         = "#111113";  ThemePanelBg       = "#18181B"
    ThemeCardBg        = "#1E1E1E";  ThemeCardAltBg     = "#1A1A1A"
    ThemeInputBg       = "#141414";  ThemeDeepBg        = "#0D0D0D"
    ThemeOutputBg      = "#0A0A0A";  ThemeSurfaceBg     = "#1F1F23"
    ThemeHoverBg       = "#27272B";  ThemeSelectedBg    = "#2A2A2A"
    ThemePressedBg     = "#1A1A1A"
    ThemeAccent        = "#0078D4";  ThemeAccentHover   = "#1A8AD4"
    ThemeAccentLight   = "#60CDFF";  ThemeAccentDim     = "#1A0078D4"
    ThemeGreenAccent   = "#00C853";  ThemeAccentText    = "#FFFFFF"
    ThemeTextPrimary   = "#FFFFFF";  ThemeTextBody      = "#E0E0E0"
    ThemeTextSecondary = "#A1A1AA";  ThemeTextMuted     = "#71717A"
    ThemeTextDim       = "#8B8B93";  ThemeTextDisabled  = "#383838"
    ThemeTextFaintest  = "#6B6B73"
    ThemeBorder        = "#0FFFFFFF"; ThemeBorderCard   = "#0FFFFFFF"
    ThemeBorderElevated = "#333333";  ThemeBorderHover  = "#444444"
    ThemeBorderMedium  = "#1AFFFFFF"; ThemeBorderSubtle = "#333333"
    ThemeScrollThumb   = "#999999";  ThemeScrollTrack   = "#22FFFFFF"
    ThemeSuccess       = "#00C853";  ThemeWarning       = "#F59E0B"
    ThemeError         = "#FF5000";  ThemeErrorDim      = "#20FF5000"
    ThemeProgressEdge  = "#18FFFFFF"
    ThemeSidebarBg     = "#111113";  ThemeSidebarBorder = "#00000000"
}

$Global:ThemeLight = @{
    ThemeAppBg         = "#F5F5F5";  ThemePanelBg       = "#FAFAFA"
    ThemeCardBg        = "#FFFFFF";  ThemeCardAltBg     = "#F8F8F8"
    ThemeInputBg       = "#F0F0F0";  ThemeDeepBg        = "#EEEEEE"
    ThemeOutputBg      = "#F0F0F0";  ThemeSurfaceBg     = "#F2F2F5"
    ThemeHoverBg       = "#E8E8EC";  ThemeSelectedBg    = "#E0E0E4"
    ThemePressedBg     = "#D8D8DC"
    ThemeAccent        = "#0078D4";  ThemeAccentHover   = "#106EBE"
    ThemeAccentLight   = "#0063B1";  ThemeAccentDim     = "#200078D4"
    ThemeGreenAccent   = "#107C10";  ThemeAccentText    = "#FFFFFF"
    ThemeTextPrimary   = "#111111";  ThemeTextBody      = "#222222"
    ThemeTextSecondary = "#555555";  ThemeTextMuted     = "#666666"
    ThemeTextDim       = "#808080";  ThemeTextDisabled  = "#999999"
    ThemeTextFaintest  = "#AAAAAA"
    ThemeBorder        = "#E0E0E0";  ThemeBorderCard   = "#D0D0D0"
    ThemeBorderElevated = "#CCCCCC"; ThemeBorderHover  = "#AAAAAA"
    ThemeBorderMedium  = "#1A000000"; ThemeBorderSubtle = "#CCCCCC"
    ThemeScrollThumb   = "#888888";  ThemeScrollTrack   = "#18000000"
    ThemeSuccess       = "#16A34A";  ThemeWarning       = "#EA580C"
    ThemeError         = "#DC2626";  ThemeErrorDim      = "#20DC2626"
    ThemeProgressEdge  = "#18000000"
    ThemeSidebarBg     = "#EEEEEE";  ThemeSidebarBorder = "#D0D0D0"
}

$ApplyTheme = {
    param([bool]$IsLight)
    $Palette = if ($IsLight) { $Global:ThemeLight } else { $Global:ThemeDark }

    foreach ($Key in $Palette.Keys) {
        try {
            $NewColor = [System.Windows.Media.ColorConverter]::ConvertFromString($Palette[$Key])
            $NewBrush = [System.Windows.Media.SolidColorBrush]::new($NewColor)
            $Window.Resources[$Key] = $NewBrush
        } catch { }
    }
    $Global:IsLightMode = $IsLight

    # Repaint dot grid
    if ($bdrDotGrid) {
        $DotColor = if ($IsLight) { '#14000000' } else { '#1EFFFFFF' }
        $Brush = [System.Windows.Media.DrawingBrush]::new()
        $Brush.TileMode   = 'Tile'
        $Brush.Viewport   = [System.Windows.Rect]::new(0,0,44,44)
        $Brush.ViewportUnits = 'Absolute'
        $Brush.Viewbox    = [System.Windows.Rect]::new(0,0,44,44)
        $Brush.ViewboxUnits  = 'Absolute'
        $DotGeo = [System.Windows.Media.EllipseGeometry]::new(
            [System.Windows.Point]::new(22,22), 1.3, 1.3)
        $DotBrush = $Global:CachedBC.ConvertFromString($DotColor)
        $Drawing  = [System.Windows.Media.GeometryDrawing]::new($DotBrush, $null, $DotGeo)
        $Brush.Drawing = $Drawing
        $bdrDotGrid.Background = $Brush
    }

    # Repaint gradient glow
    if ($bdrGradientGlow) {
        $AccentHex = if ($IsLight) { '#0078D4' } else { '#0078D4' }
        $lgb = [System.Windows.Media.LinearGradientBrush]::new()
        $lgb.StartPoint = [System.Windows.Point]::new(0.5, 0)
        $lgb.EndPoint   = [System.Windows.Point]::new(0.5, 1)
        $lgb.GradientStops.Add([System.Windows.Media.GradientStop]::new(
            [System.Windows.Media.ColorConverter]::ConvertFromString(
                $(if ($IsLight) { '#100078D4' } else { '#200078D4' })), 0.0))
        $lgb.GradientStops.Add([System.Windows.Media.GradientStop]::new(
            [System.Windows.Media.ColorConverter]::ConvertFromString(
                $(if ($IsLight) { '#080078D4' } else { '#100078D4' })), 0.35))
        $lgb.GradientStops.Add([System.Windows.Media.GradientStop]::new(
            [System.Windows.Media.ColorConverter]::ConvertFromString(
                $(if ($IsLight) { '#040078D4' } else { '#080078D4' })), 0.6))
        $lgb.GradientStops.Add([System.Windows.Media.GradientStop]::new(
            [System.Windows.Media.ColorConverter]::ConvertFromString('#00000000'), 1.0))
        $bdrGradientGlow.Background = $lgb
    }

    Write-DebugLog "[THEME] Applied $(if ($IsLight) { 'Light' } else { 'Dark' }) palette ($($Palette.Keys.Count) brush keys)" -Level 'INFO'
    Write-DebugLog "[THEME] Dot grid repainted: dotColor=$DotColor" -Level 'DEBUG'
    Write-DebugLog "[THEME] Gradient glow repainted: 4 stops" -Level 'DEBUG'
}

# ===============================================================================
# SECTION 9: TOAST NOTIFICATIONS
# ===============================================================================

    <#
    .SYNOPSIS
        Displays a themed, auto-dismissing toast notification in the bottom-right corner.
    .DESCRIPTION
        Programmatically builds a WPF Border with icon and message text, inserts it into
        the pnlToastHost StackPanel, fades it in over 200 ms, then starts a two-timer
        dismiss chain (fade-out at DurationMs, remove element 250 ms later). Colours adapt
        to the current dark/light theme. Multiple toasts stack vertically.
    .PARAMETER Message
        The notification text to display.
    .PARAMETER Type
        Visual style: Success (green), Error (red), Warning (amber), or Info (blue).
    .PARAMETER DurationMs
        Time in milliseconds before the toast auto-dismisses. Default: 4000.
    #>
function Show-Toast {
    param(
        [string]$Message,
        [ValidateSet('Success','Error','Warning','Info')][string]$Type = 'Info',
        [int]$DurationMs = $Script:TOAST_DURATION_MS
    )

    $Colors = if ($Global:IsLightMode) {
        switch ($Type) {
            'Success' { @{ Bg='#F0FDF4';  Border='#16A34A'; Text='#166534'; Icon=[char]0xE73E } }
            'Error'   { @{ Bg='#FEF2F2';  Border='#DC2626'; Text='#991B1B'; Icon=[char]0xEA39 } }
            'Warning' { @{ Bg='#FFF7ED';  Border='#EA580C'; Text='#9A3412'; Icon=[char]0xE7BA } }
            default   { @{ Bg='#EFF6FF';  Border='#0078D4'; Text='#1E40AF'; Icon=[char]0xE946 } }
        }
    } else {
        switch ($Type) {
            'Success' { @{ Bg='#0A3D1A';  Border='#00C853'; Text='#FFFFFF'; Icon=[char]0xE73E } }
            'Error'   { @{ Bg='#3D0A0A';  Border='#D13438'; Text='#FFFFFF'; Icon=[char]0xEA39 } }
            'Warning' { @{ Bg='#3D2D0A';  Border='#FFB900'; Text='#FFFFFF'; Icon=[char]0xE7BA } }
            default   { @{ Bg='#0A1E3D';  Border='#0078D4'; Text='#FFFFFF'; Icon=[char]0xE946 } }
        }
    }

    $Toast = [System.Windows.Controls.Border]::new()
    $Toast.CornerRadius    = [System.Windows.CornerRadius]::new(12)
    $Toast.BorderThickness = [System.Windows.Thickness]::new(1)
    $Toast.Padding         = [System.Windows.Thickness]::new(12,8,12,8)
    $Toast.Margin          = [System.Windows.Thickness]::new(0,0,0,6)
    $Toast.Background      = $Global:CachedBC.ConvertFromString($Colors.Bg)
    $Toast.BorderBrush     = $Global:CachedBC.ConvertFromString($Colors.Border)
    $Toast.Opacity         = 0

    $SP = [System.Windows.Controls.StackPanel]::new()
    $SP.Orientation = 'Horizontal'

    $IconTB = [System.Windows.Controls.TextBlock]::new()
    $IconTB.Text       = $Colors.Icon
    $IconTB.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
    $IconTB.FontSize   = 13
    $IconTB.Foreground = $Global:CachedBC.ConvertFromString($Colors.Border)
    $IconTB.VerticalAlignment = 'Center'
    $IconTB.Margin     = [System.Windows.Thickness]::new(0,0,8,0)
    $SP.Children.Add($IconTB)

    $MsgTB = [System.Windows.Controls.TextBlock]::new()
    $MsgTB.Text       = $Message
    $MsgTB.FontSize   = 12
    $MsgTB.Foreground = $Global:CachedBC.ConvertFromString($Colors.Text)
    $MsgTB.VerticalAlignment = 'Center'
    $MsgTB.TextWrapping = 'Wrap'
    $MsgTB.MaxWidth   = 400
    $SP.Children.Add($MsgTB)

    $Toast.Child = $SP
    $pnlToastHost.Children.Insert(0, $Toast)

    # Fade in
    $FadeIn = [System.Windows.Media.Animation.DoubleAnimation]::new()
    $FadeIn.From     = 0; $FadeIn.To = 1
    $FadeIn.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(200))
    $Toast.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $FadeIn)

    # Auto-dismiss timer (two-timer pattern for PS 5.1 safety)
    $ToastRef = $Toast
    $DismissTimer = New-Object System.Windows.Threading.DispatcherTimer
    $DismissTimer.Interval = [TimeSpan]::FromMilliseconds($DurationMs)
    $DismissTimer.Add_Tick({
        $DismissTimer.Stop()
        $FadeOut = [System.Windows.Media.Animation.DoubleAnimation]::new()
        $FadeOut.From     = 1; $FadeOut.To = 0
        $FadeOut.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(200))
        $ToastRef.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $FadeOut)

        $RemoveTimer = New-Object System.Windows.Threading.DispatcherTimer
        $RemoveTimer.Interval = [TimeSpan]::FromMilliseconds(250)
        $RemoveTimer.Add_Tick({
            $RemoveTimer.Stop()
            try { $pnlToastHost.Children.Remove($ToastRef) } catch { }
        }.GetNewClosure())
        $RemoveTimer.Start()
    }.GetNewClosure())
    $DismissTimer.Start()

    Write-DebugLog "[TOAST] [$Type] $Message (duration=${DurationMs}ms, queue=$($pnlToastHost.Children.Count))" -Level 'DEBUG'
}

# ===============================================================================
# SECTION 10: SHOW-THEMEDDIALOG
# ===============================================================================

    <#
    .SYNOPSIS
        Shows a modal dialog window that inherits the current theme from the main window.
    .DESCRIPTION
        Programmatically constructs a chromeless WPF Window with rounded corners, drop
        shadow, icon badge, title, message, optional text input field, and configurable
        buttons. The dialog inherits all DynamicResource brushes from the main window so
        it matches the active dark/light theme. Returns the clicked button's Result string,
        or a hashtable with Result and Input if an input field was requested. The dialog
        supports drag-move on the outer border area.
    .PARAMETER Title
        Dialog header text.
    .PARAMETER Message
        Body text displayed below the header separator.
    .PARAMETER Icon
        Segoe Fluent Icons character for the header icon badge.
    .PARAMETER IconColor
        Theme resource key for the icon foreground colour (e.g. 'ThemeAccentLight').
    .PARAMETER Buttons
        Array of hashtables, each with Text (label), IsAccent (bool), and Result (string).
    .PARAMETER InputPrompt
        If non-empty, adds a labelled text input field above the buttons.
    .PARAMETER InputMatch
        Reserved for future input validation (regex pattern).
    .OUTPUTS
        [string] The Result value of the clicked button, or 'Cancel' if closed.
        [hashtable] If InputPrompt was specified: @{ Result; Input }.
    #>
function Show-ThemedDialog {
    param(
        [string]$Title    = 'Confirm',
        [string]$Message  = '',
        [string]$Icon     = [string]([char]0xE897),
        [string]$IconColor = 'ThemeAccentLight',
        [array]$Buttons   = @( @{ Text='OK'; IsAccent=$true; Result='OK' } ),
        [string]$InputPrompt = '',
        [string]$InputMatch  = ''
    )

    Write-DebugLog "[DIALOG] Opening: Title='$Title', Buttons=$($Buttons.Count), HasInput=$(![string]::IsNullOrEmpty($InputPrompt))" -Level 'DEBUG'
    $DlgResult = 'Cancel'

    $Dlg = [System.Windows.Window]::new()
    $Dlg.WindowStyle             = 'None'
    $Dlg.AllowsTransparency      = $true
    $Dlg.Background              = [System.Windows.Media.Brushes]::Transparent
    $Dlg.SizeToContent           = 'WidthAndHeight'
    $Dlg.WindowStartupLocation   = 'CenterOwner'
    $Dlg.Owner                   = $Window
    $Dlg.MinWidth                = 380
    $Dlg.MaxWidth                = 520
    $Dlg.ShowInTaskbar           = $false

    # Inherit theme resources from main window
    foreach ($rk in $Window.Resources.Keys) {
        $Dlg.Resources[$rk] = $Window.Resources[$rk]
    }

    # Outer border with shadow
    $OuterBorder = [System.Windows.Controls.Border]::new()
    $OuterBorder.CornerRadius    = [System.Windows.CornerRadius]::new(12)
    $OuterBorder.BorderThickness = [System.Windows.Thickness]::new(1)
    $OuterBorder.Padding         = [System.Windows.Thickness]::new(24,20,24,20)
    $OuterBorder.Margin          = [System.Windows.Thickness]::new(20)
    $OuterBorder.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeCardBg')
    $OuterBorder.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderElevated')
    $Shadow = [System.Windows.Media.Effects.DropShadowEffect]::new()
    $Shadow.Color       = [System.Windows.Media.Colors]::Black
    $Shadow.BlurRadius  = 40
    $Shadow.ShadowDepth = 8
    $Shadow.Opacity     = 0.5
    $OuterBorder.Effect = $Shadow

    $Dlg.Content = $OuterBorder

    $MainStack = [System.Windows.Controls.StackPanel]::new()
    $OuterBorder.Child = $MainStack

    # Header
    $Header = [System.Windows.Controls.StackPanel]::new()
    $Header.Orientation = 'Horizontal'
    $Header.Margin = [System.Windows.Thickness]::new(0,0,0,16)

    $IconBadge = [System.Windows.Controls.Border]::new()
    $IconBadge.Width  = 36; $IconBadge.Height = 36
    $IconBadge.CornerRadius = [System.Windows.CornerRadius]::new(8)
    $IconBadge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeAccentDim')
    $IconBadge.Margin = [System.Windows.Thickness]::new(0,0,12,0)
    $IconTB2 = [System.Windows.Controls.TextBlock]::new()
    $IconTB2.Text = $Icon
    $IconTB2.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
    $IconTB2.FontSize = 17
    $IconTB2.HorizontalAlignment = 'Center'
    $IconTB2.VerticalAlignment   = 'Center'
    $IconTB2.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $IconColor)
    $IconBadge.Child = $IconTB2
    $Header.Children.Add($IconBadge)

    $TitleTB = [System.Windows.Controls.TextBlock]::new()
    $TitleTB.Text = $Title
    $TitleTB.FontSize   = 15
    $TitleTB.FontWeight = [System.Windows.FontWeights]::Bold
    $TitleTB.VerticalAlignment = 'Center'
    $TitleTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextPrimary')
    $Header.Children.Add($TitleTB)
    $MainStack.Children.Add($Header)

    # Separator
    $Sep = [System.Windows.Controls.Border]::new()
    $Sep.Height = 1
    $Sep.Margin = [System.Windows.Thickness]::new(0,0,0,16)
    $Sep.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeBorder')
    $MainStack.Children.Add($Sep)

    # Message
    if ($Message) {
        $MsgBlock = [System.Windows.Controls.TextBlock]::new()
        $MsgBlock.Text = $Message
        $MsgBlock.FontSize   = 12
        $MsgBlock.TextWrapping = 'Wrap'
        $MsgBlock.Margin     = [System.Windows.Thickness]::new(0,0,0, $(if ($InputPrompt) { 14 } else { 24 }))
        $MsgBlock.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextSecondary')
        $MainStack.Children.Add($MsgBlock)
    }

    # Optional input field
    $InputBox = $null
    if ($InputPrompt) {
        $InputLabel = [System.Windows.Controls.TextBlock]::new()
        $InputLabel.Text = $InputPrompt
        $InputLabel.FontSize = 11
        $InputLabel.Margin   = [System.Windows.Thickness]::new(0,0,0,6)
        $InputLabel.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextSecondary')
        $MainStack.Children.Add($InputLabel)

        $InputBorder = [System.Windows.Controls.Border]::new()
        $InputBorder.CornerRadius    = [System.Windows.CornerRadius]::new(6)
        $InputBorder.BorderThickness = [System.Windows.Thickness]::new(1)
        $InputBorder.Margin          = [System.Windows.Thickness]::new(0,0,0,24)
        $InputBorder.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeInputBg')
        $InputBorder.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderSubtle')

        $InputBox = [System.Windows.Controls.TextBox]::new()
        $InputBox.FontSize       = 13
        $InputBox.FontFamily     = [System.Windows.Media.FontFamily]::new("Cascadia Code, Consolas, Segoe UI")
        $InputBox.Padding        = [System.Windows.Thickness]::new(10,8,10,8)
        $InputBox.Background     = [System.Windows.Media.Brushes]::Transparent
        $InputBox.BorderThickness = [System.Windows.Thickness]::new(0)
        $InputBox.SetResourceReference([System.Windows.Controls.TextBox]::ForegroundProperty, 'ThemeTextBody')
        $InputBox.SetResourceReference([System.Windows.Controls.TextBox]::CaretBrushProperty, 'ThemeTextBody')
        $InputBorder.Child = $InputBox
        $MainStack.Children.Add($InputBorder)
    }

    # Button row
    $BtnRow = [System.Windows.Controls.StackPanel]::new()
    $BtnRow.Orientation = 'Horizontal'
    $BtnRow.HorizontalAlignment = 'Right'

    # Chromeless button template (removes default WPF chrome for clean rendering)
    $ChromelessTpl = [System.Windows.Markup.XamlReader]::Parse(
        '<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" TargetType="Button"><Border Padding="{TemplateBinding Padding}" Background="{TemplateBinding Background}"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border></ControlTemplate>')

    foreach ($BtnDef in $Buttons) {
        $Btn = [System.Windows.Controls.Button]::new()
        $Btn.Content = $BtnDef.Text
        $Btn.Tag     = $BtnDef.Result
        $Btn.Padding = [System.Windows.Thickness]::new(20,9,20,9)
        $Btn.FontSize = 13
        $Btn.FontWeight = [System.Windows.FontWeights]::SemiBold
        $Btn.Cursor  = [System.Windows.Input.Cursors]::Hand
        $Btn.Template = $ChromelessTpl
        $Btn.Background = [System.Windows.Media.Brushes]::Transparent

        $BtnBorder = [System.Windows.Controls.Border]::new()
        $BtnBorder.CornerRadius = [System.Windows.CornerRadius]::new(8)
        $BtnBorder.Margin = [System.Windows.Thickness]::new(6,0,0,0)

        if ($BtnDef.IsAccent) {
            $BtnBorder.Background = $Global:CachedBC.ConvertFromString('#0078D4')
            $Btn.Foreground = [System.Windows.Media.Brushes]::White
        } else {
            $BtnBorder.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeSurfaceBg')
            $BtnBorder.BorderThickness = [System.Windows.Thickness]::new(1)
            $BtnBorder.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderSubtle')
            $Btn.SetResourceReference([System.Windows.Controls.Button]::ForegroundProperty, 'ThemeTextSecondary')
        }

        $BtnBorder.Child = $Btn
        $BtnRow.Children.Add($BtnBorder)

        # Build closure in an isolated child scope so each button captures its own $tag/$dlg
        $handler = & {
            param($tag, $dlg)
            return {
                $dlg.Tag = $tag
                $dlg.DialogResult = $true
                $dlg.Close()
            }.GetNewClosure()
        } $BtnDef.Result $Dlg
        $Btn.Add_Click($handler)
    }

    $MainStack.Children.Add($BtnRow)

    $OuterBorder.Add_MouseLeftButtonDown({
        try { $Dlg.DragMove() } catch { }
    }.GetNewClosure())

    $Dlg.ShowDialog() | Out-Null
    $DlgResult = if ($Dlg.Tag) { $Dlg.Tag } else { 'Cancel' }
    Write-DebugLog "[DIALOG] Closed: Result='$DlgResult'$(if ($InputBox) { ", Input='$($InputBox.Text)'" })" -Level 'DEBUG'

    if ($InputBox) {
        return @{ Result = $DlgResult; Input = $InputBox.Text }
    }
    return $DlgResult
}

# ===============================================================================
# SECTION 11: SHIMMER PROGRESS ANIMATION
# ===============================================================================

$Global:ShimmerRunning = $false

    <#
    .SYNOPSIS
        Starts the indeterminate shimmer progress animation in the toolbar.
    .DESCRIPTION
        Makes the shimmer border visible and begins a repeating DoubleAnimation on the
        ShimmerTranslate TranslateTransform, sweeping from -500 to 1200 over 2 seconds.
        Guards against double-start via $Global:ShimmerRunning flag.
    #>
function Start-Shimmer {
    if ($Global:ShimmerRunning) { return }
    Write-DebugLog "[SHIMMER] Start" -Level 'DEBUG'
    $Global:ShimmerRunning = $true
    $bdrShimmer.Visibility = 'Visible'
    $ShimmerAnim = [System.Windows.Media.Animation.DoubleAnimation]::new()
    $ShimmerAnim.From           = -500
    $ShimmerAnim.To             = 1200
    $ShimmerAnim.Duration       = [System.Windows.Duration]::new([TimeSpan]::FromSeconds(2))
    $ShimmerAnim.RepeatBehavior = [System.Windows.Media.Animation.RepeatBehavior]::Forever
    $ShimmerTranslate = $Window.FindName("ShimmerTranslate")
    if ($ShimmerTranslate) {
        $ShimmerTranslate.BeginAnimation(
            [System.Windows.Media.TranslateTransform]::XProperty, $ShimmerAnim)
    }
}

    <#
    .SYNOPSIS
        Stops the shimmer progress animation and hides the shimmer border.
    #>
function Stop-Shimmer {
    Write-DebugLog "[SHIMMER] Stop" -Level 'DEBUG'
    $Global:ShimmerRunning = $false
    $bdrShimmer.Visibility = 'Collapsed'
    $ShimmerTranslate = $Window.FindName("ShimmerTranslate")
    if ($ShimmerTranslate) {
        $ShimmerTranslate.BeginAnimation(
            [System.Windows.Media.TranslateTransform]::XProperty, $null)
    }
}

# ===============================================================================
# SECTION 11b: SPLASH STEP HELPERS
# ===============================================================================

    <#
    .SYNOPSIS
        Updates a numbered step indicator on the splash overlay.
    .DESCRIPTION
        Sets the fill colour of the step dot and the foreground of the step label based
        on the Status value: pending (dim), running (accent blue), done (green), error
        (red), or skipped (dim). Optionally updates the label text.
    .PARAMETER Step
        Step number (1, 2, or 3) corresponding to dotStep{N} and lblStep{N} XAML elements.
    .PARAMETER Status
        Visual state: pending, running, done, error, or skipped.
    .PARAMETER Text
        Optional replacement text for the step label.
    #>
function Set-SplashStep {
    param([int]$Step, [string]$Status, [string]$Text)
    # Status: pending, running, done, error, skipped
    $Dot = Get-Variable -Name "dotStep$Step" -ValueOnly -Scope Script
    $Lbl = Get-Variable -Name "lblStep$Step" -ValueOnly -Scope Script
    if ($Text) { $Lbl.Text = $Text }
    switch ($Status) {
        'pending' { $Dot.Fill = $Window.Resources['ThemeTextDim'] }
        'running' {
            $Dot.Fill = $Global:CachedBC.ConvertFromString('#0078D4')
            $Lbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextPrimary')
        }
        'done' {
            $Dot.Fill = $Global:CachedBC.ConvertFromString(
                $(if ($Global:IsLightMode) { '#16A34A' } else { '#00C853' }))
            $Lbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextPrimary')
        }
        'error' {
            $Dot.Fill = $Window.Resources['ThemeError']
            $Lbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeError')
        }
        'skipped' {
            $Dot.Fill = $Window.Resources['ThemeTextDim']
            $Lbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
        }
    }
}

    <#
    .SYNOPSIS
        Collapses the splash overlay, revealing the main application content beneath.
    #>
function Hide-Splash {
    Write-DebugLog "[SPLASH] Dismissing splash overlay" -Level 'DEBUG'
    $pnlSplash.Visibility = 'Collapsed'
}

    <#
    .SYNOPSIS
        Schedules automatic dismissal of the splash overlay after a delay.
    .DESCRIPTION
        If the window has already rendered, starts a DispatcherTimer immediately. If not,
        sets a flag so the Window.ContentRendered handler starts the timer. This ensures
        the splash stays visible until both the delay has elapsed and the window is painted.
    .PARAMETER DelayMs
        Milliseconds to wait before dismissing the splash. Default: 3000.
    #>
function Start-SplashDismissTimer {
    param([int]$DelayMs = 3000)
    # Defer to ContentRendered so the timer doesn't fire before the window is visible
    $Script:SplashDismissDelayMs = $DelayMs
    if ($Global:WindowRendered) {
        # Already rendered - start timer immediately
        $ST = [System.Windows.Threading.DispatcherTimer]::new()
        $ST.Interval = [TimeSpan]::FromMilliseconds($DelayMs)
        $STRef = $ST
        $HideSplashFn = { Hide-Splash }
        $ST.Add_Tick({ $STRef.Stop(); & $HideSplashFn }.GetNewClosure())
        $ST.Start()
    } else {
        # Will be started by ContentRendered handler
        $Global:SplashDismissQueued = $true
    }
}

# ===============================================================================
# SECTION 12: TAB & NAVIGATION MANAGEMENT
# ===============================================================================

$AllTabs = @('Dashboard','Changes','Timeline','Resources','Settings')
$AllTabHeaders = @{
    Dashboard  = $tabDashboard
    Changes    = $tabChanges
    Timeline   = $tabTimeline
    Resources  = $tabResources
    Settings   = $tabSettings
}
$AllTabPanels = @{
    Dashboard  = $pnlTabDashboard
    Changes    = $pnlTabChanges
    Timeline   = $pnlTabTimeline
    Resources  = $pnlTabResources
    Settings   = $pnlTabSettings
}
$AllRailIndicators = @{
    Dashboard  = $railDashboardIndicator
    Changes    = $railChangesIndicator
    Timeline   = $railTimelineIndicator
    Resources  = $railResourcesIndicator
    Settings   = $railSettingsIndicator
}
$AllSidebarPanels = @{
    Dashboard  = $pnlSidebarDashboard
    Changes    = $pnlSidebarChanges
    Timeline   = $pnlSidebarTimeline
    Resources  = $pnlSidebarResources
    Settings   = $pnlSidebarSettings
}

    <#
    .SYNOPSIS
        Navigates to the specified tab, updating all visual indicators and sidebar panels.
    .DESCRIPTION
        Iterates through all five tabs (Dashboard, Changes, Timeline, Resources, Settings)
        and for each: styles the tab header (accent border + hover background for active,
        transparent for inactive), toggles content panel Visibility, updates the icon rail
        indicator colour, and shows/hides the corresponding sidebar panel. Applies a 150 ms
        fade-in animation to the newly active panel. Auto-refreshes Timeline data when that
        tab is selected.
    .PARAMETER TabName
        Target tab: Dashboard, Changes, Timeline, Resources, or Settings.
    #>
function Switch-Tab {
    param([string]$TabName)
    $Global:ActiveTabName = $TabName
    $Global:ActiveRailName = $TabName

    foreach ($t in $AllTabs) {
        # Tab header styling
        $hdr = $AllTabHeaders[$t]
        if ($hdr) {
            if ($t -eq $TabName) {
                $hdr.BorderBrush = $Global:CachedBC.ConvertFromString(
                    $(if ($Global:IsLightMode) { '#0078D4' } else { '#0078D4' }))
                $hdr.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeHoverBg')
            } else {
                $hdr.BorderBrush = [System.Windows.Media.Brushes]::Transparent
                $hdr.Background  = [System.Windows.Media.Brushes]::Transparent
            }
        }

        # Tab panel visibility
        $pnl = $AllTabPanels[$t]
        if ($pnl) {
            $pnl.Visibility = if ($t -eq $TabName) { 'Visible' } else { 'Collapsed' }
        }

        # Rail indicator
        $ind = $AllRailIndicators[$t]
        if ($ind) {
            if ($t -eq $TabName) {
                $ind.SetResourceReference([System.Windows.Shapes.Rectangle]::FillProperty, 'ThemeAccentLight')
            } else {
                $ind.Fill = [System.Windows.Media.Brushes]::Transparent
            }
        }

        # Sidebar panel
        $sb = $AllSidebarPanels[$t]
        if ($sb) {
            $sb.Visibility = if ($t -eq $TabName) { 'Visible' } else { 'Collapsed' }
        }
    }

    # Tab fade-in animation
    $ActivePanel = $AllTabPanels[$TabName]
    if ($ActivePanel) {
        $Fade = [System.Windows.Media.Animation.DoubleAnimation]::new()
        $Fade.From     = 0; $Fade.To = 1
        $Fade.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(150))
        $ActivePanel.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $Fade)
    }

    # Auto-refresh data for specific tabs
    if ($TabName -eq 'Timeline') { Populate-Timeline }

    Write-DebugLog "[NAV] Switched to tab: $TabName (sidebar=$($AllSidebarPanels[$TabName] -ne $null))" -Level 'DEBUG'
}

# ===============================================================================
# SECTION 13: USER PREFERENCES
# ===============================================================================

    <#
    .SYNOPSIS
        Loads persisted user preferences from user_prefs.json.
    .DESCRIPTION
        Reads the JSON preferences file and applies saved values to UI controls:
        theme (dark/light), auto-snapshot toggle, retention days/enabled, excluded
        resource types, default subscription ID, and volatile change filter state.
        Silently skips missing or corrupt files.
    #>
function Load-UserPrefs {
    Write-DebugLog "[PREFS] Loading from: $Global:PrefsPath" -Level 'DEBUG'
    if (Test-Path $Global:PrefsPath) {
        try {
            $prefs = Get-Content $Global:PrefsPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($null -ne $prefs.theme -and $prefs.theme -eq 'light') {
                $chkDarkMode.IsChecked = $false
                & $ApplyTheme $true
            }
            if ($null -ne $prefs.autoSnapshotOnLaunch) {
                $chkAutoSnapshot.IsChecked = [bool]$prefs.autoSnapshotOnLaunch
            }
            if ($null -ne $prefs.snapshotRetentionDays) {
                $txtRetentionDays.Text = $prefs.snapshotRetentionDays.ToString()
            }
            if ($null -ne $prefs.retentionEnabled) {
                $chkRetentionEnabled.IsChecked = [bool]$prefs.retentionEnabled
                $vis = if ([bool]$prefs.retentionEnabled) { 'Visible' } else { 'Collapsed' }
                $pnlRetentionDays.Visibility = $vis
            }
            if ($null -ne $prefs.excludedResourceTypes) {
                $txtExcludedTypes.Text = ($prefs.excludedResourceTypes -join "`r`n")
            }
            if ($null -ne $prefs.defaultSubscriptionId) {
                $Global:CurrentSubId = $prefs.defaultSubscriptionId
            }
            if ($null -ne $prefs.hideVolatileChanges) {
                $Global:HideVolatileEnabled = [bool]$prefs.hideVolatileChanges
                Update-VolatileToggleVisual
            }
            Write-DebugLog "[PREFS] Loaded: theme=$($prefs.theme), autoSnapshot=$($prefs.autoSnapshotOnLaunch), retention=$($prefs.snapshotRetentionDays)d, excludedTypes=$($prefs.excludedResourceTypes.Count), savedSub=$($prefs.defaultSubscriptionId)" -Level 'SUCCESS'
        } catch {
            Write-DebugLog "[PREFS] Failed to load: $($_.Exception.Message)" -Level 'WARN'
            Write-DebugLog "[PREFS] StackTrace: $($_.ScriptStackTrace)" -Level 'DEBUG'
        }
    }
}

    <#
    .SYNOPSIS
        Persists current UI settings to user_prefs.json.
    .DESCRIPTION
        Serialises the current theme, auto-snapshot toggle, retention settings,
        excluded resource types, selected subscription ID, and volatile filter state
        into a JSON file alongside the script. Called on close, theme change, setting
        change, and subscription switch.
    #>
function Save-UserPrefs {
    $prefs = @{
        theme                 = if ($Global:IsLightMode) { 'light' } else { 'dark' }
        autoSnapshotOnLaunch  = [bool]$chkAutoSnapshot.IsChecked
        retentionEnabled      = [bool]$chkRetentionEnabled.IsChecked
        snapshotRetentionDays = [int]$txtRetentionDays.Text
        excludedResourceTypes = @($txtExcludedTypes.Text -split "`r?`n" | Where-Object { $_.Trim() })
        defaultSubscriptionId = $Global:CurrentSubId
        hideVolatileChanges   = [bool]$Global:HideVolatileEnabled
    }
    try {
        $prefs | ConvertTo-Json -Depth 4 | Set-Content $Global:PrefsPath -Encoding UTF8
        Write-DebugLog "[PREFS] Saved: theme=$($prefs.theme), sub=$($prefs.defaultSubscriptionId)" -Level 'DEBUG'
    } catch {
        Write-DebugLog "[PREFS] Save failed: $($_.Exception.Message)" -Level 'WARN'
        Write-DebugLog "[PREFS] StackTrace: $($_.ScriptStackTrace)" -Level 'DEBUG'
    }
}

    <#
    .SYNOPSIS
        Deletes snapshot and diff files older than the configured retention period.
    .DESCRIPTION
        When retention is enabled, scans the snapshots/ and diffs/ directories for files
        whose LastWriteTime is older than txtRetentionDays days. Each matching file is
        removed with -Force. Logs the count of deleted files. Defaults to 90 days if the
        retention days text box contains an invalid value.
    #>
function Invoke-RetentionCleanup {
    if (-not [bool]$chkRetentionEnabled.IsChecked) {
        Write-DebugLog "[RETENTION] Cleanup disabled -- skipping" -Level 'DEBUG'
        return
    }
    $days = 90
    if ([int]::TryParse($txtRetentionDays.Text, [ref]$days) -and $days -gt 0) {
        # keep as parsed
    } else {
        $days = 90
    }
    $cutoff = (Get-Date).AddDays(-$days)
    $removed = 0
    foreach ($dir in @($Global:SnapshotDir, $Global:DiffDir)) {
        if (-not (Test-Path $dir)) { continue }
        Get-ChildItem $dir -File | Where-Object { $_.LastWriteTime -lt $cutoff } | ForEach-Object {
            try {
                Remove-Item $_.FullName -Force
                $removed++
                Write-DebugLog "[RETENTION] Deleted: $($_.Name)" -Level 'DEBUG'
            } catch {
                Write-DebugLog "[RETENTION] Failed to delete $($_.Name): $($_.Exception.Message)" -Level 'WARN'
            }
        }
    }
    if ($removed -gt 0) {
        Write-DebugLog "[RETENTION] Cleaned up $removed file(s) older than $days days" -Level 'INFO'
    } else {
        Write-DebugLog "[RETENTION] No files older than $days days found" -Level 'DEBUG'
    }
}
# ===============================================================================
# SECTION 14: SNAPSHOT ENGINE
# Captures a full resource inventory from the current Azure subscription
# via Search-AzGraph, computes SHA-256 state hashes excluding volatile
# properties, and saves the result as a dated JSON file. Uses a background
# runspace for non-blocking operation.
# ===============================================================================

$Global:VolatilePropertyPaths = @(
    'properties.instanceView'
    'properties.statusTimestamp'
    'properties.lastModifiedTime'
    'properties.lastModifiedTimeUtc'
    'properties.lastModifiedBy'
    'properties.lastModifiedByType'
    'properties.lastModifiedAt'
    'properties.uniqueId'
    'changedTime'
)
function New-ResourceSnapshot {
    <#
    .SYNOPSIS
        Captures all resources in the current subscription via Search-AzGraph.
    .DESCRIPTION
        Launches a background runspace that pages through all resources using Search-AzGraph
        with a configurable page size ($Script:GRAPH_PAGE_SIZE, default 1000). Each resource
        is enriched with a SHA-256 state hash computed from its stable properties (volatile
        fields like instanceView, timestamps, and modification metadata are excluded).
        The snapshot is saved as a dated JSON file under snapshots/sub_{id}/ and includes
        metadata (snapshot ID, subscription, capture timestamp, resource/RG counts, tool
        version). The OnComplete callback updates the UI, dashboard cards, resource filters,
        sidebar counts, and triggers snapshot-related achievements (first_snapshot,
        five_snapshots, ten_snapshots, over_100, over_500).
    .PARAMETER SubscriptionId
        The Azure subscription GUID to snapshot.
    .PARAMETER SubscriptionName
        Display name of the subscription (stored in snapshot metadata).
    #>
    param([string]$SubscriptionId, [string]$SubscriptionName)

    $SubId   = $SubscriptionId
    $SubName = $SubscriptionName
    $ExcludedTypes = @($txtExcludedTypes.Text -split "`r?`n" | Where-Object { $_.Trim() })
    $VolatilePaths = $Global:VolatilePropertyPaths
    $SnapDir = $Global:SnapshotDir
    $PageSize = $Script:GRAPH_PAGE_SIZE

    Write-DebugLog "[SNAPSHOT] === Snapshot button clicked ===" -Level 'DEBUG'
    Write-DebugLog "[SNAPSHOT] Subscription: $SubName ($SubId)" -Level 'INFO'
    Write-DebugLog "[SNAPSHOT] PageSize=$PageSize, ExcludedTypes=$($ExcludedTypes.Count), VolatilePaths=$($VolatilePaths.Count)" -Level 'DEBUG'
    Write-DebugLog "[SNAPSHOT] Target dir: $SnapDir" -Level 'DEBUG'
    Start-Shimmer



    Start-BackgroundWork -Variables @{
        SubId          = $SubId
        SubName        = $SubName
        ExcludedTypes  = $ExcludedTypes
        VolatilePaths  = $VolatilePaths
        SnapDir        = $SnapDir
        PageSize       = $PageSize
    } -Work {
        $ProgressPreference = 'SilentlyContinue'
        Import-Module Az.Accounts -MinimumVersion 2.7.5 -ErrorAction Stop
        Import-Module Az.ResourceGraph -ErrorAction Stop
        $AllResources = [System.Collections.ArrayList]::new()
        $Query = "resources | extend provisioningState = tostring(properties.provisioningState), createdTime = tostring(properties.createdTime), changedTime = tostring(properties.changedTime) | project id, name, type, resourceGroup, location, sku, tags, properties, provisioningState, createdTime, changedTime"
        $Skip  = $null

        do {
            $Params = @{
                Query        = $Query
                Subscription = @($SubId)
                First        = $PageSize
            }
            if ($Skip) { $Params['SkipToken'] = $Skip }

            $Result = Search-AzGraph @Params
            foreach ($r in $Result.Data) {
                # Skip excluded types
                if ($ExcludedTypes -contains $r.type) { continue }

                # Remove volatile properties for hashing
                $HashObj = $r | Select-Object -Property * -ExcludeProperty instanceView, statusTimestamp, lastModifiedTime, uniqueId

                $SHA = [System.Security.Cryptography.SHA256]::Create()
                $Json  = $HashObj | ConvertTo-Json -Depth 20 -Compress
                $Bytes = [System.Text.Encoding]::UTF8.GetBytes($Json)
                $Hash  = $SHA.ComputeHash($Bytes)
                $HashStr = [BitConverter]::ToString($Hash).Replace('-','').ToLowerInvariant()

                $AllResources.Add(@{
                    id               = $r.id
                    name             = $r.name
                    type             = $r.type
                    resourceGroup    = $r.resourceGroup
                    location         = $r.location
                    sku              = $r.sku
                    tags             = $r.tags
                    properties       = $r.properties
                    provisioningState = $r.provisioningState
                    createdTime      = $r.createdTime
                    changedTime      = $r.changedTime
                    stateHash        = $HashStr
                }) | Out-Null
            }
            $Skip = $Result.SkipToken
        } while ($Skip)

        # Build snapshot
        $Context = Get-AzContext
        $Snapshot = @{
            metadata = @{
                snapshotId         = [guid]::NewGuid().ToString()
                subscriptionId     = $SubId
                subscriptionName   = $SubName
                capturedAt         = (Get-Date).ToUniversalTime().ToString('o')
                capturedBy         = $Context.Account.Id
                resourceCount      = $AllResources.Count
                resourceGroupCount = ($AllResources | Select-Object -ExpandProperty resourceGroup -Unique).Count
                toolVersion        = '0.1.0-alpha'
            }
            resources = $AllResources
        }

        # Save
        $SubDir = Join-Path $SnapDir "sub_$($SubId -replace '-','')"
        if (-not (Test-Path $SubDir)) { New-Item -Path $SubDir -ItemType Directory -Force | Out-Null }
        $FileName = (Get-Date).ToString('yyyy-MM-dd') + '.json'
        $FilePath = Join-Path $SubDir $FileName
        $Snapshot | ConvertTo-Json -Depth 30 | Set-Content $FilePath -Encoding UTF8

        return @{
            Snapshot = $Snapshot
            FilePath = $FilePath
        }
    } -OnComplete {
        param($Results, $Errors)
        Stop-Shimmer

        Write-DebugLog "[SnapshotCB] OnComplete ENTERED (Results=$($Results.Count), Errors=$($Errors.Count))" -Level 'DEBUG'
        if ($Errors.Count -gt 0) {
            Write-DebugLog "[SnapshotCB] FAILED: $($Errors[0])" -Level 'ERROR'
            if ($Errors[0].Exception) {
                Write-DebugLog "[SnapshotCB] InnerException: $($Errors[0].Exception.InnerException)" -Level 'ERROR'
            }
            Show-Toast -Message "Snapshot failed: $($Errors[0].ToString())" -Type 'Error'
            return
        }

        $Data = $Results | Select-Object -Last 1
        if ($Data -and $Data.Snapshot) {
            $Global:LatestSnapshot = $Data.Snapshot
            $Count = $Data.Snapshot.metadata.resourceCount
            $RgCount = $Data.Snapshot.metadata.resourceGroupCount
            Write-DebugLog "[SnapshotCB] Complete: $Count resources, $RgCount resource groups" -Level 'SUCCESS'
            Write-DebugLog "[SnapshotCB] SnapshotId: $($Data.Snapshot.metadata.snapshotId)" -Level 'DEBUG'
            Write-DebugLog "[SnapshotCB] File: $($Data.FilePath)" -Level 'DEBUG'
            Show-Toast -Message "Snapshot captured: $Count resources" -Type 'Success'

            # Update UI
            Update-SnapshotList
            Update-DashboardCards
            Update-ResourceFilters -Snapshot $Data.Snapshot
            Populate-ResourcesList -Snapshot $Data.Snapshot
            $lblSidebarResourceCount.Text = $Count.ToString()
            $lblSidebarLastSnapshot.Text  = (Get-Date).ToString('MMM dd, HH:mm')
            $lblStatus.Text               = "Snapshot complete ($Count resources)"
            $lblStatusResourceCount.Text  = "$Count resources"
            $lblDashTotalResources.Text   = $Count.ToString()

            Save-UserPrefs

            # Achievement triggers
            Unlock-Achievement 'first_snapshot'
            $SnapCount = @(Get-ChildItem (Join-Path $Global:Root 'snapshots') -Recurse -Filter '*.json' -ErrorAction SilentlyContinue).Count
            if ($SnapCount -ge 5)  { Unlock-Achievement 'five_snapshots' }
            if ($SnapCount -ge 10) { Unlock-Achievement 'ten_snapshots' }
            if ($Count -ge 100)    { Unlock-Achievement 'over_100' }
            if ($Count -ge 500)    { Unlock-Achievement 'over_500' }
        }
    }
}

# ===============================================================================
# SECTION 15: DIFF ENGINE
# Three-pass comparison (Added/Removed/Modified) between two snapshots.
# Uses canonical JSON serialisation with sorted keys and normalised
# timestamps for stable hashing. Classifies each change by severity
# (Critical/High/Medium/Low) based on resource type and changed properties.
# Flags volatile-only changes for optional filtering.
# ===============================================================================

$SeverityRules = @{
    Critical = @(
        'Microsoft.Authorization/*'
        'Microsoft.Network/networkSecurityGroups*'
        'Microsoft.KeyVault/vaults/accessPolicies'
        'Microsoft.ManagedIdentity/*'
    )
    High     = @(
        'sku.name'
        'sku.tier'
        'location'
        'Microsoft.Network/publicIPAddresses'
        'Microsoft.Network/virtualNetworks/virtualNetworkPeerings'
    )
}

    <#
    .SYNOPSIS
        Classifies a change as Critical, High, Medium, or Low severity.
    .DESCRIPTION
        Applies a cascade of rules:
        1. All Removed resources are Critical.
        2. Resource types matching $SeverityRules.Critical patterns (authorization, NSGs,
           KeyVault access policies, managed identities) are Critical.
        3. Property changes matching $SeverityRules.High patterns (sku, location, public IPs,
           VNet peerings) are High.
        4. Tag-only changes are Medium.
        5. Added resources are Medium.
        6. Everything else is Low.
    .PARAMETER ChangeType
        The change classification: Added, Removed, or Modified.
    .PARAMETER ResourceType
        Full Azure resource type string (e.g. 'Microsoft.Network/networkSecurityGroups').
    .PARAMETER Details
        Hashtable of property-level diffs (key = property path, value = @{from;to}).
    .OUTPUTS
        [string] Severity label: Critical, High, Medium, or Low.
    #>
function Get-ChangeSeverity {
    param([string]$ChangeType, [string]$ResourceType, [hashtable]$Details)

    if ($ChangeType -eq 'Removed') { return 'Critical' }

    foreach ($pattern in $SeverityRules.Critical) {
        if ($ResourceType -like $pattern) { return 'Critical' }
    }

    if ($Details) {
        foreach ($prop in $Details.Keys) {
            foreach ($pattern in $SeverityRules.High) {
                if ($prop -like $pattern) { return 'High' }
            }
        }
        if ($Details.Keys -match '^tags\.') { return 'Medium' }
    }

    if ($ChangeType -eq 'Added') { return 'Medium' }

    return 'Low'
}

    <#
    .SYNOPSIS
        Recursively computes property-level diffs between two objects.
    .DESCRIPTION
        Walks nested PSCustomObject/hashtable trees, building dot-notation property paths.
        Leaf values are compared using canonical JSON serialisation (ConvertTo-CanonicalJson)
        to avoid false positives from key ordering or timestamp precision differences. The
        'stateHash' property is excluded from comparison. Returns a flat hashtable where keys
        are dot-notation paths and values are @{ from = oldValue; to = newValue }.
    .PARAMETER Old
        The baseline object or value.
    .PARAMETER New
        The current object or value.
    .PARAMETER Prefix
        Accumulated dot-notation path (used internally during recursion).
    .OUTPUTS
        [hashtable] Flat map of changed property paths to @{ from; to } pairs.
    #>
function Compare-PropertyDiff {
    param([object]$Old, [object]$New, [string]$Prefix = '')
    $Diff = @{}
    
    $OldIsObj = ($null -ne $Old) -and ($Old -is [hashtable] -or $Old.GetType().Name -eq 'PSCustomObject')
    $NewIsObj = ($null -ne $New) -and ($New -is [hashtable] -or $New.GetType().Name -eq 'PSCustomObject')

    if (($OldIsObj -or $NewIsObj) -and ($OldIsObj -or $null -eq $Old) -and ($NewIsObj -or $null -eq $New)) {
        $OldProps = if ($Old -is [hashtable]) { $Old.Keys } elseif ($OldIsObj) { $Old.PSObject.Properties.Name } else { @() }
        $NewProps = if ($New -is [hashtable]) { $New.Keys } elseif ($NewIsObj) { $New.PSObject.Properties.Name } else { @() }
        $AllKeys = @($OldProps) + @($NewProps) | Select-Object -Unique

        foreach ($k in $AllKeys) {
            if ($Prefix -eq '' -and $k -eq 'stateHash') { continue }
            
            $OldVal = if ($null -eq $Old) { $null } elseif ($Old -is [hashtable]) { $Old[$k] } else { $Old.$k }
            $NewVal = if ($null -eq $New) { $null } elseif ($New -is [hashtable]) { $New[$k] } else { $New.$k }
            
            $nextPrefix = if ($Prefix -eq '') { $k } else { "$Prefix.$k" }
            $subDiff = Compare-PropertyDiff -Old $OldVal -New $NewVal -Prefix $nextPrefix
            foreach ($subK in $subDiff.Keys) {
                $Diff[$subK] = $subDiff[$subK]
            }
        }
    } else {
        # Use canonical JSON (sorted keys, normalized timestamps) to avoid false positives
        $OldJson = if ($null -ne $Old) { ConvertTo-CanonicalJson $Old } else { 'null' }
        $NewJson = if ($null -ne $New) { ConvertTo-CanonicalJson $New } else { 'null' }
        
        if ($OldJson -ne $NewJson) {
            $fmtOld = if ($null -eq $Old) { "null" } elseif ($Old -is [array] -or $OldIsObj) { $OldJson } else { "$Old" }
            $fmtNew = if ($null -eq $New) { "null" } elseif ($New -is [array] -or $NewIsObj) { $NewJson } else { "$New" }
            
            if ([string]::IsNullOrEmpty($fmtOld)) { $fmtOld = "null" }
            if ([string]::IsNullOrEmpty($fmtNew)) { $fmtNew = "null" }
            
            $Diff[$Prefix] = @{ from = $fmtOld; to = $fmtNew }
        }
    }
    
    return $Diff
}

# Canonical JSON serialization (sorted keys, recursive) for stable hashing
    <#
    .SYNOPSIS
        Serialises an object to deterministic JSON with sorted keys and normalised timestamps.
    .DESCRIPTION
        Produces a canonical JSON representation suitable for stable hashing. Object keys are
        sorted alphabetically at every nesting level. ISO 8601 timestamp strings are normalised
        to 3-digit millisecond 'Z' format to avoid false positives when Azure returns varying
        sub-second precision across snapshots. Handles strings, booleans, numbers, datetimes,
        arrays, hashtables, and PSCustomObjects.
    .PARAMETER Obj
        The object to serialise. Can be any combination of primitive types, arrays,
        hashtables, and PSCustomObjects.
    .OUTPUTS
        [string] Deterministic JSON string.
    #>
function ConvertTo-CanonicalJson {
    param([object]$Obj)
    if ($null -eq $Obj) { return 'null' }
    if ($Obj -is [string]) {
        # Normalize ISO 8601 timestamps to 3-digit ms to prevent false positives
        # from Azure returning varying sub-second precision across snapshots
        if ($Obj -match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?Z?$') {
            $dt = [datetime]::MinValue
            if ([datetime]::TryParse($Obj, [ref]$dt)) {
                $Obj = $dt.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
            }
        }
        return '"' + $Obj.Replace('\','\\').Replace('"','\"') + '"'
    }
    if ($Obj -is [bool]) { if ($Obj) { return 'true' } else { return 'false' } }
    if ($Obj -is [int] -or $Obj -is [long] -or $Obj -is [double] -or $Obj -is [decimal] -or $Obj -is [float]) { return $Obj.ToString() }
    if ($Obj -is [datetime]) { return '"' + $Obj.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ') + '"' }
    if ($Obj -is [array] -or $Obj -is [System.Collections.IList]) {
        $items = @(); foreach ($item in $Obj) { $items += (ConvertTo-CanonicalJson $item) }
        return '[' + ($items -join ',') + ']'
    }
    # Object/hashtable - sort keys
    $keys = @()
    if ($Obj -is [hashtable]) { $keys = $Obj.Keys | Sort-Object }
    elseif ($Obj.PSObject) { $keys = $Obj.PSObject.Properties.Name | Sort-Object }
    else { return '"' + $Obj.ToString() + '"' }
    $pairs = @()
    foreach ($k in $keys) {
        $v = if ($Obj -is [hashtable]) { $Obj[$k] } else { $Obj.$k }
        $pairs += '"' + $k + '":' + (ConvertTo-CanonicalJson $v)
    }
    return '{' + ($pairs -join ',') + '}'
}

    <#
    .SYNOPSIS
        Computes a SHA-256 hash of a resource's stable properties for change detection.
    .DESCRIPTION
        Deep-clones the resource properties, strips all volatile metadata paths (instance
        views, timestamps, modification info), builds a canonical JSON representation with
        sorted keys, and returns the SHA-256 hex digest. Two resources with the same
        canonical hash are considered unchanged between snapshots.
    .PARAMETER Resource
        A resource object from a snapshot, containing id, name, type, resourceGroup,
        location, sku, tags, properties, and provisioningState.
    .OUTPUTS
        [string] Lowercase hex SHA-256 hash (64 characters).
    #>
function Get-CanonicalHash {
    param([object]$Resource)
    # Hash only the stable, meaningful fields — strip volatile metadata first
    $Props = $Resource.properties
    if ($null -ne $Props) {
        # Deep-clone properties to avoid mutating the snapshot, then remove volatile keys
        $Props = $Props | ConvertTo-Json -Depth 20 -Compress | ConvertFrom-Json
        foreach ($vp in $Global:VolatilePropertyPaths) {
            if ($vp -like 'properties.*') {
                $leaf = $vp.Substring('properties.'.Length)
                # Support one level of nesting (e.g. 'lastModifiedTimeUtc')
                if ($null -ne $Props.$leaf) { $Props.PSObject.Properties.Remove($leaf) }
            }
        }
    }
    $Canonical = @{
        id = $Resource.id; name = $Resource.name; type = $Resource.type
        resourceGroup = $Resource.resourceGroup; location = $Resource.location
        sku = $Resource.sku; tags = $Resource.tags; properties = $Props
        provisioningState = $Resource.provisioningState
    }
    # Also remove top-level volatile keys (e.g. 'changedTime')
    foreach ($vp in $Global:VolatilePropertyPaths) {
        if ($vp -notlike '*.*') { $Canonical.Remove($vp) }
    }
    $Json = ConvertTo-CanonicalJson $Canonical
    $SHA = [System.Security.Cryptography.SHA256]::Create()
    $Hash = $SHA.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Json))
    return [BitConverter]::ToString($Hash).Replace('-','').ToLowerInvariant()
}

    <#
    .SYNOPSIS
        Compares two resource snapshots and produces a structured diff with severity classification.
    .DESCRIPTION
        Performs a three-pass comparison:
          Pass 1 (Added): Resources in ToSnapshot not present in FromSnapshot.
          Pass 2 (Removed): Resources in FromSnapshot not present in ToSnapshot.
          Pass 3 (Modified): Resources present in both where the canonical hash differs;
            runs Compare-PropertyDiff to identify specific property changes and flags
            changes that are volatile-only (metadata timestamps).
        Each change is classified by Get-ChangeSeverity. Returns a diff object with metadata
        (summary counts, snapshot dates, diff ID) and an ArrayList of change records. Also
        starts/stops the shimmer progress animation during processing.
    .PARAMETER FromSnapshot
        The baseline (older) snapshot object with metadata and resources arrays.
    .PARAMETER ToSnapshot
        The current (newer) snapshot object.
    .OUTPUTS
        [hashtable] Diff object with .metadata.summary (added/removed/modified/unchanged
        counts) and .changes ArrayList of change records.
    #>
function Compare-Snapshots {
    param([object]$FromSnapshot, [object]$ToSnapshot)

    $DiffStart = Get-Date
    Write-DebugLog "[DIFF] === Compare-Snapshots ENTERED ===" -Level 'DEBUG'
    Write-DebugLog "[DIFF] FROM: $($FromSnapshot.metadata.capturedAt) ($($FromSnapshot.resources.Count) resources)" -Level 'INFO'
    Write-DebugLog "[DIFF] TO:   $($ToSnapshot.metadata.capturedAt) ($($ToSnapshot.resources.Count) resources)" -Level 'INFO'
    Start-Shimmer

    # Build lookup tables
    $OldMap = @{}
    foreach ($r in $FromSnapshot.resources) {
        $OldMap[$r.id] = $r
    }
    $NewMap = @{}
    foreach ($r in $ToSnapshot.resources) {
        $NewMap[$r.id] = $r
    }
    Write-DebugLog "[DIFF] Lookup tables built: Old=$($OldMap.Count), New=$($NewMap.Count)" -Level 'DEBUG'

    $Changes = [System.Collections.ArrayList]::new()

    # Pass 1: Added (in New, not in Old)
    foreach ($id in $NewMap.Keys) {
        if (-not $OldMap.ContainsKey($id)) {
            $r = $NewMap[$id]
            $sev = Get-ChangeSeverity -ChangeType 'Added' -ResourceType $r.type -Details $null
            $Changes.Add(@{
                changeType    = 'Added'
                resourceId    = $id
                resourceName  = $r.name
                resourceType  = $r.type
                resourceGroup = $r.resourceGroup
                severity      = $sev
                details       = @{}
            }) | Out-Null
        }
    }

    # Pass 2: Removed (in Old, not in New)
    foreach ($id in $OldMap.Keys) {
        if (-not $NewMap.ContainsKey($id)) {
            $r = $OldMap[$id]
            $sev = Get-ChangeSeverity -ChangeType 'Removed' -ResourceType $r.type -Details $null
            $Changes.Add(@{
                changeType    = 'Removed'
                resourceId    = $id
                resourceName  = $r.name
                resourceType  = $r.type
                resourceGroup = $r.resourceGroup
                severity      = $sev
                details       = @{}
            }) | Out-Null
        }
    }

    # Pass 3: Modified (in both, canonical hash differs)
    foreach ($id in $NewMap.Keys) {
        if ($OldMap.ContainsKey($id)) {
            $OldR = $OldMap[$id]
            $NewR = $NewMap[$id]
            $OldHash = Get-CanonicalHash $OldR
            $NewHash = Get-CanonicalHash $NewR
            if ($OldHash -ne $NewHash) {
                $details = Compare-PropertyDiff -Old $OldR -New $NewR
                # Check if all changes are volatile/metadata-only
                $nonVolatileKeys = @($details.Keys | Where-Object { $_ -notin $Global:VolatilePropertyPaths })
                $isVolatileOnly = ($details.Count -gt 0 -and $nonVolatileKeys.Count -eq 0)
                $sev = Get-ChangeSeverity -ChangeType 'Modified' -ResourceType $NewR.type -Details $details
                $Changes.Add(@{
                    changeType    = 'Modified'
                    resourceId    = $id
                    resourceName  = $NewR.name
                    resourceType  = $NewR.type
                    resourceGroup = $NewR.resourceGroup
                    severity      = $sev
                    details       = $details
                    volatileOnly  = $isVolatileOnly
                }) | Out-Null
            }
        }
    }

    Stop-Shimmer

    $Diff = @{
        metadata = @{
            diffId         = [guid]::NewGuid().ToString()
            subscriptionId = $ToSnapshot.metadata.subscriptionId
            fromSnapshot   = $FromSnapshot.metadata.capturedAt
            toSnapshot     = $ToSnapshot.metadata.capturedAt
            summary        = @{
                added     = ($Changes | Where-Object { $_.changeType -eq 'Added' }).Count
                removed   = ($Changes | Where-Object { $_.changeType -eq 'Removed' }).Count
                modified  = ($Changes | Where-Object { $_.changeType -eq 'Modified' }).Count
                unchanged = $ToSnapshot.metadata.resourceCount - $Changes.Count
            }
        }
        changes = $Changes
    }

    $Global:LatestDiff = $Diff
    $DiffElapsed = [math]::Round(((Get-Date) - $DiffStart).TotalMilliseconds)
    Write-DebugLog "[DIFF] Complete in ${DiffElapsed}ms: +$($Diff.metadata.summary.added) -$($Diff.metadata.summary.removed) ~$($Diff.metadata.summary.modified) =$($Diff.metadata.summary.unchanged)" -Level 'SUCCESS'
    Write-DebugLog "[DIFF] DiffId: $($Diff.metadata.diffId)" -Level 'DEBUG'

    return $Diff
}

# ===============================================================================
# SECTION 16: UI UPDATE HELPERS
# ===============================================================================

    <#
    .SYNOPSIS
        Refreshes the sidebar snapshot list and comparison combo boxes.
    .DESCRIPTION
        Scans the subscription-specific snapshot directory for JSON files,
        adds baseline files from the baselines/ directory, populates the sidebar
        ListBox and the From/To comparison ComboBoxes. Auto-selects the two most
        recent snapshots for quick comparison. Updates the sidebar snapshot count label.
    #>
function Update-SnapshotList {
    Write-DebugLog "[UI] Update-SnapshotList: sub=$Global:CurrentSubId" -Level 'DEBUG'
    $lstSnapshots.Items.Clear()
    $cmbDiffFrom.Items.Clear()
    $cmbDiffTo.Items.Clear()
    $Global:SnapshotList = @()

    if (-not $Global:CurrentSubId) { Write-DebugLog "[UI] No subscription selected, skipping" -Level 'DEBUG'; return }
    $SubDir = Join-Path $Global:SnapshotDir "sub_$($Global:CurrentSubId -replace '-','')"
    if (-not (Test-Path $SubDir)) { New-Item -Path $SubDir -ItemType Directory -Force | Out-Null }

    $Files = Get-ChildItem $SubDir -Filter '*.json' | Sort-Object Name -Descending
    $Global:SnapshotList = @($Files)

    # Include baselines (subscription-independent) so they appear in compare dropdowns
    $BaselineFiles = @()
    if (Test-Path $Global:BaselineDir) {
        $BaselineFiles = Get-ChildItem $Global:BaselineDir -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object Name -Descending
    }

    $lblSidebarSnapshotCount.Text = $Files.Count.ToString()

    foreach ($f in $Files) {
        $DateStr = $f.BaseName  # YYYY-MM-DD
        $lstSnapshots.Items.Add($DateStr) | Out-Null
        $cmbDiffFrom.Items.Add($DateStr) | Out-Null
        $cmbDiffTo.Items.Add($DateStr) | Out-Null
    }

    foreach ($bf in $BaselineFiles) {
        $Label = "[B] $($bf.BaseName)"
        $Global:SnapshotList += $bf
        $cmbDiffFrom.Items.Add($Label) | Out-Null
        $cmbDiffTo.Items.Add($Label) | Out-Null
    }

    if ($Files.Count -ge 2) {
        $cmbDiffFrom.SelectedIndex = 1  # older
        $cmbDiffTo.SelectedIndex   = 0  # newer
    } elseif ($Files.Count -eq 1) {
        $cmbDiffFrom.SelectedIndex = 0
        $cmbDiffTo.SelectedIndex   = 0
    }
    Write-DebugLog "[UI] Snapshot list refreshed: $($Files.Count) snapshots, $($BaselineFiles.Count) baselines" -Level 'DEBUG'
}

    <#
    .SYNOPSIS
        Refreshes the Dashboard tab stat cards and Changes tab summary badges.
    .DESCRIPTION
        Updates Total Resources from the latest snapshot, and Changes/Added/Removed
        counts plus badge labels from the latest diff. Called after snapshot capture
        and comparison operations.
    #>
function Update-DashboardCards {
    Write-DebugLog "[UI] Update-DashboardCards: hasSnapshot=$($null -ne $Global:LatestSnapshot), hasDiff=$($null -ne $Global:LatestDiff)" -Level 'DEBUG'
    if ($Global:LatestSnapshot) {
        $lblDashTotalResources.Text = $Global:LatestSnapshot.metadata.resourceCount.ToString()
    }
    if ($Global:LatestDiff) {
        $s = $Global:LatestDiff.metadata.summary
        $lblDashChanges.Text  = ($s.added + $s.removed + $s.modified).ToString()
        $lblDashAdded.Text    = $s.added.ToString()
        $lblDashRemoved.Text  = $s.removed.ToString()

        $lblBadgeAdded.Text    = "$($s.added) Added"
        $lblBadgeRemoved.Text  = "$($s.removed) Removed"
        $lblBadgeModified.Text = "$($s.modified) Modified"
    }
}

# --- Timeline functions --------------------------------------------------------

    <#
    .SYNOPSIS
        Builds the Timeline tab UI from persisted diff files.
    .DESCRIPTION
        Reads all diff JSON files from the subscription's diffs/ directory, applies the
        current timeline search text and change type filter, and constructs a vertical
        timeline with connector lines, date range headers, mini stat pills (added/removed/
        modified counts), and expandable detail panels showing individual resource changes
        grouped by change type. Each timeline entry is a clickable card that toggles its
        detail panel with a chevron indicator.
    #>
function Populate-Timeline {
    $pnlTimeline.Children.Clear()
    if (-not $Global:CurrentSubId) { return }

    $DiffSubDir = Join-Path $Global:DiffDir "sub_$($Global:CurrentSubId -replace '-','')"
    $DiffFiles = Get-ChildItem $DiffSubDir -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object Name -Descending
    $FilterText = $txtTimelineSearch.Text.Trim()
    $FilterType = if ($cmbTimelineType.SelectedItem) { $cmbTimelineType.SelectedItem.ToString() } else { 'All' }

    if (-not $DiffFiles -or $DiffFiles.Count -eq 0) {
        $Msg = [System.Windows.Controls.TextBlock]::new()
        $Msg.Text = "No diffs yet -- compare two snapshots to build a timeline."
        $Msg.FontSize = 11; $Msg.FontStyle = [System.Windows.FontStyles]::Italic
        $Msg.HorizontalAlignment = 'Center'; $Msg.Margin = [System.Windows.Thickness]::new(0,40,0,0)
        $Msg.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
        $pnlTimeline.Children.Add($Msg) | Out-Null
        return
    }

    $TotalEntries = 0

    foreach ($df in $DiffFiles) {
        try { $Diff = Get-Content $df.FullName -Raw | ConvertFrom-Json } catch { continue }
        if (-not $Diff.changes -or $Diff.changes.Count -eq 0) { continue }

        # Apply filters to get matching changes for this diff
        $Matched = @($Diff.changes | Where-Object {
            ($FilterType -eq 'All' -or $_.changeType -eq $FilterType) -and
            (-not $FilterText -or $_.resourceName -like "*$FilterText*" -or $_.resourceType -like "*$FilterText*" -or $_.resourceGroup -like "*$FilterText*")
        })
        if ($Matched.Count -eq 0) { continue }

        # Extract dates from metadata
        $_from = $Diff.metadata.fromSnapshot
        $_to   = $Diff.metadata.toSnapshot
        $fStr  = if ($_from -is [datetime]) { $_from.ToString('yyyy-MM-dd') } else { ([string]$_from).Substring(0,10) }
        $tStr  = if ($_to -is [datetime]) { $_to.ToString('yyyy-MM-dd') } else { ([string]$_to).Substring(0,10) }

        # Counts for this diff
        $cAdded    = @($Matched | Where-Object { $_.changeType -eq 'Added' }).Count
        $cRemoved  = @($Matched | Where-Object { $_.changeType -eq 'Removed' }).Count
        $cModified = @($Matched | Where-Object { $_.changeType -eq 'Modified' }).Count

        # ── Timeline entry: horizontal layout with connector ──
        $EntryGrid = [System.Windows.Controls.Grid]::new()
        $EntryGrid.Margin = [System.Windows.Thickness]::new(0,0,0,0)
        # Column 0: timeline connector (fixed 32px)
        $col0 = [System.Windows.Controls.ColumnDefinition]::new()
        $col0.Width = [System.Windows.GridLength]::new(32)
        # Column 1: card (star)
        $col1 = [System.Windows.Controls.ColumnDefinition]::new()
        $col1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $EntryGrid.ColumnDefinitions.Add($col0) | Out-Null
        $EntryGrid.ColumnDefinitions.Add($col1) | Out-Null

        # ── Connector column: vertical line + dot ──
        $ConnectorCanvas = [System.Windows.Controls.Canvas]::new()
        $ConnectorCanvas.Width = 32
        $ConnectorCanvas.ClipToBounds = $false
        [System.Windows.Controls.Grid]::SetColumn($ConnectorCanvas, 0)

        # Vertical line (full height, centered at x=15)
        $VLine = [System.Windows.Shapes.Rectangle]::new()
        $VLine.Width = 2
        $VLine.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, 'ThemeBorder')
        [System.Windows.Controls.Canvas]::SetLeft($VLine, 15)
        [System.Windows.Controls.Canvas]::SetTop($VLine, 0)
        $VLine.Height = 400
        $ConnectorCanvas.Children.Add($VLine) | Out-Null

        # Dot at top of line
        $Dot = [System.Windows.Shapes.Ellipse]::new()
        $Dot.Width = 10; $Dot.Height = 10
        $Dot.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, 'ThemeAccent')
        [System.Windows.Controls.Canvas]::SetLeft($Dot, 11)
        [System.Windows.Controls.Canvas]::SetTop($Dot, 14)
        $ConnectorCanvas.Children.Add($Dot) | Out-Null

        $EntryGrid.Children.Add($ConnectorCanvas) | Out-Null

        # ── Card column: summary + expandable details ──
        $CardOuter = [System.Windows.Controls.StackPanel]::new()
        $CardOuter.Margin = [System.Windows.Thickness]::new(4,0,0,16)
        [System.Windows.Controls.Grid]::SetColumn($CardOuter, 1)

        # Summary card (always visible)
        $SummaryCard = [System.Windows.Controls.Border]::new()
        $SummaryCard.CornerRadius = [System.Windows.CornerRadius]::new(8)
        $SummaryCard.Padding = [System.Windows.Thickness]::new(14,10,14,10)
        $SummaryCard.BorderThickness = [System.Windows.Thickness]::new(1)
        $SummaryCard.Cursor = [System.Windows.Input.Cursors]::Hand
        $SummaryCard.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeCardBg')
        $SummaryCard.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderCard')

        $SumSP = [System.Windows.Controls.StackPanel]::new()
        $SummaryCard.Child = $SumSP

        # Row 1: date range + total changes
        $Row1 = [System.Windows.Controls.DockPanel]::new()
        $DateSP = [System.Windows.Controls.StackPanel]::new()
        $DateSP.Orientation = 'Horizontal'
        $DateSP.VerticalAlignment = 'Center'
        $FromTB = [System.Windows.Controls.TextBlock]::new()
        $FromTB.Text = $fStr; $FromTB.FontSize = 11; $FromTB.FontWeight = [System.Windows.FontWeights]::SemiBold
        $FromTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextSecondary')
        $ArrowTB = [System.Windows.Controls.TextBlock]::new()
        $ArrowTB.Text = ' -> '; $ArrowTB.FontSize = 11
        $ArrowTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
        $ToTB = [System.Windows.Controls.TextBlock]::new()
        $ToTB.Text = $tStr; $ToTB.FontSize = 11; $ToTB.FontWeight = [System.Windows.FontWeights]::SemiBold
        $ToTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextPrimary')
        $DateSP.Children.Add($FromTB) | Out-Null
        $DateSP.Children.Add($ArrowTB) | Out-Null
        $DateSP.Children.Add($ToTB) | Out-Null
        [System.Windows.Controls.DockPanel]::SetDock($DateSP, 'Left')
        $Row1.Children.Add($DateSP) | Out-Null

        # Total count + chevron
        $RightSP = [System.Windows.Controls.StackPanel]::new()
        $RightSP.Orientation = 'Horizontal'; $RightSP.HorizontalAlignment = 'Right'
        $TotalBadge = [System.Windows.Controls.Border]::new()
        $TotalBadge.CornerRadius = [System.Windows.CornerRadius]::new(10)
        $TotalBadge.Padding = [System.Windows.Thickness]::new(7,1,7,1)
        $TotalBadge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeHoverBg')
        $TotalBadgeTB = [System.Windows.Controls.TextBlock]::new()
        $TotalBadgeTB.Text = "$($Matched.Count) changes"; $TotalBadgeTB.FontSize = 10
        $TotalBadgeTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
        $TotalBadge.Child = $TotalBadgeTB
        $RightSP.Children.Add($TotalBadge) | Out-Null
        $ChevronTB = [System.Windows.Controls.TextBlock]::new()
        $ChevronTB.Text = [char]0xE76C  # ChevronDown
        $ChevronTB.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
        $ChevronTB.FontSize = 10; $ChevronTB.Margin = [System.Windows.Thickness]::new(8,0,0,0)
        $ChevronTB.VerticalAlignment = 'Center'
        $ChevronTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
        $RightSP.Children.Add($ChevronTB) | Out-Null
        [System.Windows.Controls.DockPanel]::SetDock($RightSP, 'Right')
        $Row1.Children.Add($RightSP) | Out-Null
        $SumSP.Children.Add($Row1) | Out-Null

        # Row 2: mini stat pills (added / removed / modified)
        $Row2 = [System.Windows.Controls.StackPanel]::new()
        $Row2.Orientation = 'Horizontal'; $Row2.Margin = [System.Windows.Thickness]::new(0,6,0,0)

        $PillData = @(
            @{ Count = $cAdded;    Label = 'added';    FgKey = 'ThemeSuccess' }
            @{ Count = $cRemoved;  Label = 'removed';  FgKey = 'ThemeError' }
            @{ Count = $cModified; Label = 'modified'; FgKey = 'ThemeWarning' }
        )
        foreach ($pd in $PillData) {
            if ($pd.Count -eq 0) { continue }
            $Pill = [System.Windows.Controls.Border]::new()
            $Pill.CornerRadius = [System.Windows.CornerRadius]::new(4)
            $Pill.Padding = [System.Windows.Thickness]::new(6,1,6,1)
            $Pill.Margin = [System.Windows.Thickness]::new(0,0,6,0)
            $Pill.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeHoverBg')
            $PillSP = [System.Windows.Controls.StackPanel]::new()
            $PillSP.Orientation = 'Horizontal'
            $PillNum = [System.Windows.Controls.TextBlock]::new()
            $PillNum.Text = "$($pd.Count) "; $PillNum.FontSize = 10; $PillNum.FontWeight = [System.Windows.FontWeights]::Bold
            $PillNum.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $pd.FgKey)
            $PillLbl = [System.Windows.Controls.TextBlock]::new()
            $PillLbl.Text = $pd.Label; $PillLbl.FontSize = 10
            $PillLbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
            $PillSP.Children.Add($PillNum) | Out-Null
            $PillSP.Children.Add($PillLbl) | Out-Null
            $Pill.Child = $PillSP
            $Row2.Children.Add($Pill) | Out-Null
        }
        $SumSP.Children.Add($Row2) | Out-Null
        $CardOuter.Children.Add($SummaryCard) | Out-Null

        # ── Detail panel (collapsed) with individual changes ──
        $DetailPanel = [System.Windows.Controls.StackPanel]::new()
        $DetailPanel.Margin = [System.Windows.Thickness]::new(0,4,0,0)
        $DetailPanel.Visibility = 'Collapsed'

        # Group matched changes by changeType for visual clarity
        $TypeOrder = @('Removed','Added','Modified')
        foreach ($ct in $TypeOrder) {
            $TypeItems = @($Matched | Where-Object { $_.changeType -eq $ct })
            if ($TypeItems.Count -eq 0) { continue }

            foreach ($c in ($TypeItems | Sort-Object { $_.resourceName })) {
                $ItemCard = [System.Windows.Controls.Border]::new()
                $ItemCard.CornerRadius = [System.Windows.CornerRadius]::new(4)
                $ItemCard.Padding = [System.Windows.Thickness]::new(10,6,10,6)
                $ItemCard.Margin = [System.Windows.Thickness]::new(0,0,0,3)
                $ItemCard.BorderThickness = [System.Windows.Thickness]::new(1)
                $ItemCard.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeSurfaceBg')
                $ItemCard.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorder')

                $ItemGrid = [System.Windows.Controls.Grid]::new()
                $ig0 = [System.Windows.Controls.ColumnDefinition]::new()
                $ig0.Width = [System.Windows.GridLength]::new(8)
                $ig1 = [System.Windows.Controls.ColumnDefinition]::new()
                $ig1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
                $ig2 = [System.Windows.Controls.ColumnDefinition]::new()
                $ig2.Width = [System.Windows.GridLength]::Auto
                $ItemGrid.ColumnDefinitions.Add($ig0) | Out-Null
                $ItemGrid.ColumnDefinitions.Add($ig1) | Out-Null
                $ItemGrid.ColumnDefinitions.Add($ig2) | Out-Null

                # Change dot
                $CDot = [System.Windows.Shapes.Ellipse]::new()
                $CDot.Width = 6; $CDot.Height = 6; $CDot.VerticalAlignment = 'Center'
                $DotColor = switch ($c.changeType) { 'Added' { 'ThemeSuccess' } 'Removed' { 'ThemeError' } default { 'ThemeWarning' } }
                $CDot.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, $DotColor)
                [System.Windows.Controls.Grid]::SetColumn($CDot, 0)

                # Name + type stacked
                $InfoSP = [System.Windows.Controls.StackPanel]::new()
                $InfoSP.Margin = [System.Windows.Thickness]::new(8,0,0,0)
                $ItemName = [System.Windows.Controls.TextBlock]::new()
                $ItemName.Text = $c.resourceName; $ItemName.FontSize = 11; $ItemName.FontWeight = [System.Windows.FontWeights]::SemiBold
                $ItemName.TextTrimming = 'CharacterEllipsis'
                $ItemName.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
                $ItemType = [System.Windows.Controls.TextBlock]::new()
                $ShortType = ($c.resourceType -split '/')[-1]
                $ItemType.Text = $ShortType; $ItemType.FontSize = 9
                $ItemType.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
                $InfoSP.Children.Add($ItemName) | Out-Null
                $InfoSP.Children.Add($ItemType) | Out-Null
                [System.Windows.Controls.Grid]::SetColumn($InfoSP, 1)

                # Badge
                $ItemBadge = [System.Windows.Controls.Border]::new()
                $ItemBadge.CornerRadius = [System.Windows.CornerRadius]::new(4)
                $ItemBadge.Padding = [System.Windows.Thickness]::new(5,1,5,1)
                $ItemBadge.VerticalAlignment = 'Center'
                $ItemBadge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeHoverBg')
                $ItemBadgeTB = [System.Windows.Controls.TextBlock]::new()
                $ItemBadgeTB.Text = $c.changeType; $ItemBadgeTB.FontSize = 9
                $ItemBadgeTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $DotColor)
                $ItemBadge.Child = $ItemBadgeTB
                [System.Windows.Controls.Grid]::SetColumn($ItemBadge, 2)

                $ItemGrid.Children.Add($CDot) | Out-Null
                $ItemGrid.Children.Add($InfoSP) | Out-Null
                $ItemGrid.Children.Add($ItemBadge) | Out-Null
                $ItemCard.Child = $ItemGrid

                $DetailPanel.Children.Add($ItemCard) | Out-Null
            }
        }
        $CardOuter.Children.Add($DetailPanel) | Out-Null

        # ── Toggle click handler ──
        $SummaryCard.Tag = @{ Detail = $DetailPanel; Chevron = $ChevronTB }
        $SummaryCard.Add_MouseLeftButtonUp({
            $info = $this.Tag
            if ($info.Detail.Visibility -eq 'Visible') {
                $info.Detail.Visibility = 'Collapsed'
                $info.Chevron.Text = [char]0xE76C   # ChevronDown
            } else {
                $info.Detail.Visibility = 'Visible'
                $info.Chevron.Text = [char]0xE70D   # ChevronUp (ChevronUpMed)
            }
        })

        $EntryGrid.Children.Add($CardOuter) | Out-Null
        $pnlTimeline.Children.Add($EntryGrid) | Out-Null
        $TotalEntries++
    }

    # If filters hid everything
    if ($TotalEntries -eq 0) {
        $NoMatch = [System.Windows.Controls.TextBlock]::new()
        $NoMatch.Text = "No timeline entries match the current filters."
        $NoMatch.FontSize = 11; $NoMatch.FontStyle = [System.Windows.FontStyles]::Italic
        $NoMatch.HorizontalAlignment = 'Center'; $NoMatch.Margin = [System.Windows.Thickness]::new(0,40,0,0)
        $NoMatch.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
        $pnlTimeline.Children.Add($NoMatch) | Out-Null
    }

    Write-DebugLog "[UI] Timeline populated: $TotalEntries diff entries, $($pnlTimeline.Children.Count) UI elements" -Level 'DEBUG'
}

    <#
    .SYNOPSIS
        Populates the Changes tab list with styled change cards from a diff result.
    .DESCRIPTION
        Clears and rebuilds the lstChanges ListBox with one card per resource change. Each
        card features a coloured left accent bar, change type icon circle, resource name,
        type/resource group metadata row with severity indicator, and a staggered fade-in
        animation (opacity 0→1 + translateY 8→0). Respects the current filters: resource
        type checkboxes, change type combo, search text, and volatile change toggle. Cards
        for modified resources include an expandable property diff detail section.
    .PARAMETER Diff
        The diff object returned by Compare-Snapshots, containing .changes ArrayList.
    #>
function Populate-ChangesList {
    param([object]$Diff)
    $lstChanges.Items.Clear()
    if (-not $Diff -or -not $Diff.changes) { return }

    $BC = [System.Windows.Media.BrushConverter]::new()

    $SevIcon = @{
        Critical = [char]0xEA39; High = [char]0xE7BA
        Medium   = [char]0xE946; Low  = [char]0xE8CE
    }
    $SevColor = @{
        Critical = 'ThemeError';       High = 'ThemeWarning'
        Medium   = 'ThemeAccentLight'; Low  = 'ThemeTextDim'
    }
    $ChangeIcon = @{
        Added    = [char]0xE710   # +
        Removed  = [char]0xE74D   # x
        Modified = [char]0xE70F   # edit
    }
    $ChangeAccent = @{
        Added    = @{ Solid = '#00C853'; Dim = '#1A00C853' }
        Removed  = @{ Solid = '#FF5000'; Dim = '#1AFF5000' }
        Modified = @{ Solid = '#F59E0B'; Dim = '#1AF59E0B' }
    }

    # Collect selected resource types from checkboxes
    $SelectedTypes = @()
    foreach ($child in $pnlResourceTypeFilters.Children) {
        if ($child -is [System.Windows.Controls.CheckBox] -and $child.IsChecked -and $child.Tag -ne '__ALL__') {
            $SelectedTypes += $child.Tag
        }
    }
    $AllTypesCheck = $pnlResourceTypeFilters.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_.Tag -eq '__ALL__' }
    $FilterAllTypes = (-not $AllTypesCheck) -or $AllTypesCheck.IsChecked

    $FilterType = if ($cmbFilterChangeType.SelectedItem) { $cmbFilterChangeType.SelectedItem.ToString() } else { 'All' }
    $FilterText = $txtFilterSearch.Text.Trim()
    $HideVolatile = $Global:HideVolatileEnabled

    $CardIndex = 0
    foreach ($c in $Diff.changes) {
        # Filter out volatile-only modified resources when toggle is on
        if ($HideVolatile -and $c.volatileOnly) { continue }
        # Apply filters
        if (-not $FilterAllTypes -and $SelectedTypes.Count -gt 0) {
            $shortType = ($c.resourceType -split '/')[-1]
            if ($c.resourceType -notin $SelectedTypes -and $shortType -notin $SelectedTypes) { continue }
        }
        if ($FilterType -ne 'All' -and $c.changeType -ne $FilterType) { continue }
        if ($FilterText -and $c.resourceName -notlike "*$FilterText*" -and $c.resourceType -notlike "*$FilterText*" -and $c.resourceGroup -notlike "*$FilterText*") { continue }

        $Accent = $ChangeAccent[$c.changeType]
        if (-not $Accent) { $Accent = @{ Solid = '#8B8B93'; Dim = '#1A8B8B93' } }

        # -- Outer card with colored left accent stripe --
        $Card = [System.Windows.Controls.Border]::new()
        $Card.CornerRadius    = [System.Windows.CornerRadius]::new(8)
        $Card.BorderThickness = [System.Windows.Thickness]::new(1)
        $Card.Padding         = [System.Windows.Thickness]::new(0)
        $Card.Margin          = [System.Windows.Thickness]::new(0,0,0,6)
        $Card.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeCardBg')
        $Card.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderCard')
        # Start invisible for staggered fade-in
        $Card.Opacity = 0
        $Card.RenderTransform = [System.Windows.Media.TranslateTransform]::new(0, 8)

        # Inner DockPanel: left accent bar + content
        $Dock = [System.Windows.Controls.DockPanel]::new()

        # -- Left accent bar --
        $AccentBar = [System.Windows.Controls.Border]::new()
        $AccentBar.Width = 3
        $AccentBar.CornerRadius = [System.Windows.CornerRadius]::new(8,0,0,8)
        $AccentBar.Background = $BC.ConvertFromString($Accent.Solid)
        [System.Windows.Controls.DockPanel]::SetDock($AccentBar, 'Left')
        $Dock.Children.Add($AccentBar) | Out-Null

        # -- Content area --
        $Content = [System.Windows.Controls.StackPanel]::new()
        $Content.Margin = [System.Windows.Thickness]::new(12,10,12,10)

        # -- Row 1: Icon + Name + Badge --
        $Row1 = [System.Windows.Controls.DockPanel]::new()
        $Row1.Margin = [System.Windows.Thickness]::new(0,0,0,6)

        # Change type icon circle
        $IconCircle = [System.Windows.Controls.Border]::new()
        $IconCircle.Width = 28; $IconCircle.Height = 28
        $IconCircle.CornerRadius = [System.Windows.CornerRadius]::new(14)
        $IconCircle.Background = $BC.ConvertFromString($Accent.Dim)
        $IconCircle.Margin = [System.Windows.Thickness]::new(0,0,10,0)
        $IconCircle.VerticalAlignment = 'Center'
        $IconTB = [System.Windows.Controls.TextBlock]::new()
        $IconTB.Text = $ChangeIcon[$c.changeType]
        $IconTB.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
        $IconTB.FontSize = 12
        $IconTB.Foreground = $BC.ConvertFromString($Accent.Solid)
        $IconTB.HorizontalAlignment = 'Center'; $IconTB.VerticalAlignment = 'Center'
        $IconCircle.Child = $IconTB
        [System.Windows.Controls.DockPanel]::SetDock($IconCircle, 'Left')
        $Row1.Children.Add($IconCircle) | Out-Null

        # Change type badge (right-aligned)
        $Badge = [System.Windows.Controls.Border]::new()
        $Badge.CornerRadius = [System.Windows.CornerRadius]::new(10)
        $Badge.Padding = [System.Windows.Thickness]::new(10,3,10,3)
        $Badge.VerticalAlignment = 'Center'
        $Badge.HorizontalAlignment = 'Right'
        $Badge.Background = $BC.ConvertFromString($Accent.Dim)
        $BadgeTB = [System.Windows.Controls.TextBlock]::new()
        $BadgeTB.Text = $c.changeType.ToUpper()
        $BadgeTB.FontSize = 9; $BadgeTB.FontWeight = [System.Windows.FontWeights]::Bold
        $BadgeTB.Foreground = $BC.ConvertFromString($Accent.Solid)
        $Badge.Child = $BadgeTB
        [System.Windows.Controls.DockPanel]::SetDock($Badge, 'Right')
        $Row1.Children.Add($Badge) | Out-Null

        # Resource name (fills center)
        $NameTB = [System.Windows.Controls.TextBlock]::new()
        $NameTB.Text = $c.resourceName
        $NameTB.FontSize = 13; $NameTB.FontWeight = [System.Windows.FontWeights]::SemiBold
        $NameTB.VerticalAlignment = 'Center'
        $NameTB.TextTrimming = 'CharacterEllipsis'
        $NameTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextPrimary')
        $Row1.Children.Add($NameTB) | Out-Null

        $Content.Children.Add($Row1) | Out-Null

        # -- Row 2: Type pill . Resource Group . Severity --
        $Row2 = [System.Windows.Controls.WrapPanel]::new()
        $Row2.Orientation = 'Horizontal'

        # Resource type icon
        $ResIcon = [System.Windows.Controls.TextBlock]::new()
        $ResIcon.Text = (Get-ResourceTypeIcon -ResourceType $c.resourceType)
        $ResIcon.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
        $ResIcon.FontSize = 10; $ResIcon.Margin = [System.Windows.Thickness]::new(0,0,4,0)
        $ResIcon.VerticalAlignment = 'Center'
        $ResIcon.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
        $Row2.Children.Add($ResIcon) | Out-Null

        # Resource type text
        $TypeTB = [System.Windows.Controls.TextBlock]::new()
        $TypeTB.Text = ($c.resourceType -split '/')[-1]
        $TypeTB.FontSize = 10; $TypeTB.VerticalAlignment = 'Center'
        $TypeTB.TextTrimming = 'CharacterEllipsis'
        $TypeTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
        $Row2.Children.Add($TypeTB) | Out-Null

        # Dot separator
        $Dot1 = [System.Windows.Controls.TextBlock]::new()
        $Dot1.Text = " . "; $Dot1.FontSize = 10; $Dot1.VerticalAlignment = 'Center'
        $Dot1.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
        $Row2.Children.Add($Dot1) | Out-Null

        # Resource group with folder icon
        $RgIcon = [System.Windows.Controls.TextBlock]::new()
        $RgIcon.Text = [string]([char]0xE8B7)
        $RgIcon.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
        $RgIcon.FontSize = 9; $RgIcon.Margin = [System.Windows.Thickness]::new(0,0,3,0)
        $RgIcon.VerticalAlignment = 'Center'
        $RgIcon.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
        $Row2.Children.Add($RgIcon) | Out-Null

        $RgTB = [System.Windows.Controls.TextBlock]::new()
        $RgTB.Text = $c.resourceGroup; $RgTB.FontSize = 10
        $RgTB.VerticalAlignment = 'Center'; $RgTB.TextTrimming = 'CharacterEllipsis'
        $RgTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
        $Row2.Children.Add($RgTB) | Out-Null

        # Severity indicator
        if ($c.severity) {
            $Dot2 = [System.Windows.Controls.TextBlock]::new()
            $Dot2.Text = " . "; $Dot2.FontSize = 10; $Dot2.VerticalAlignment = 'Center'
            $Dot2.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
            $Row2.Children.Add($Dot2) | Out-Null

            $SevIconTB = [System.Windows.Controls.TextBlock]::new()
            $SevIconTB.Text = $SevIcon[$c.severity]
            $SevIconTB.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
            $SevIconTB.FontSize = 10; $SevIconTB.VerticalAlignment = 'Center'
            $SevIconTB.Margin = [System.Windows.Thickness]::new(0,0,4,0)
            $SevIconTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $SevColor[$c.severity])
            $Row2.Children.Add($SevIconTB) | Out-Null

            $SevLabelTB = [System.Windows.Controls.TextBlock]::new()
            $SevLabelTB.Text = $c.severity
            $SevLabelTB.FontSize = 10; $SevLabelTB.VerticalAlignment = 'Center'
            $SevLabelTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $SevColor[$c.severity])
            $Row2.Children.Add($SevLabelTB) | Out-Null
        }

        $Content.Children.Add($Row2) | Out-Null

        $Dock.Children.Add($Content) | Out-Null
        $Card.Child = $Dock

        # -- Hover effect --
        $CardRef = $Card
        $Card.Add_MouseEnter([System.Windows.Input.MouseEventHandler]{
            param($s,$e)
            $s.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeHoverBg')
            $s.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderHover')
        }.GetNewClosure())
        $Card.Add_MouseLeave([System.Windows.Input.MouseEventHandler]{
            param($s,$e)
            $s.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeCardBg')
            $s.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderCard')
        }.GetNewClosure())

        $lstChanges.Items.Add($Card) | Out-Null

        # -- Staggered fade-in + slide-up animation --
        $Delay = [TimeSpan]::FromMilliseconds([math]::Min($CardIndex * 30, 600))
        $FadeIn = [System.Windows.Media.Animation.DoubleAnimation]::new()
        $FadeIn.From = 0; $FadeIn.To = 1
        $FadeIn.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(250))
        $FadeIn.BeginTime = $Delay
        $FadeIn.EasingFunction = [System.Windows.Media.Animation.QuadraticEase]::new()
        $Card.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $FadeIn)

        $SlideUp = [System.Windows.Media.Animation.DoubleAnimation]::new()
        $SlideUp.From = 8; $SlideUp.To = 0
        $SlideUp.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(250))
        $SlideUp.BeginTime = $Delay
        $SlideUp.EasingFunction = [System.Windows.Media.Animation.QuadraticEase]::new()
        $Card.RenderTransform.BeginAnimation([System.Windows.Media.TranslateTransform]::YProperty, $SlideUp)

        $CardIndex++
    }

    $_from = $Diff.metadata.fromSnapshot; $_fStr = $(if ($_from -is [datetime]) { $_from.ToString('yyyy-MM-dd') } else { ([string]$_from).Substring(0,10) }); $_to = $Diff.metadata.toSnapshot; $_tStr = $(if ($_to -is [datetime]) { $_to.ToString('yyyy-MM-dd') } else { ([string]$_to).Substring(0,10) }); $lblDiffHeader.Text = "Comparing $_fStr -> $_tStr"
    Write-DebugLog "[UI] Changes list populated: $($lstChanges.Items.Count)/$($Diff.changes.Count) items (filters: type=$FilterType, text='$FilterText', resTypes=$($SelectedTypes.Count))" -Level 'DEBUG'
}

# --- Resource Type Filter Checkboxes ------------------------------------------
    <#
    .SYNOPSIS
        Dynamically generates resource type filter checkboxes from the current diff.
    .DESCRIPTION
        Extracts unique resource types from the diff changes, creates an "All Types"
        master checkbox plus individual type checkboxes with change counts and coloured
        icons. Each checkbox triggers a re-filter of the changes list. The "All Types"
        checkbox toggles all individual checkboxes on/off.
    .PARAMETER Diff
        The diff object containing .changes with resourceType properties.
    #>
function Populate-ResourceTypeFilters {
    param([object]$Diff)
    $pnlResourceTypeFilters.Children.Clear()
    if (-not $Diff -or -not $Diff.changes) { return }

    $Types = $Diff.changes | ForEach-Object { $_.resourceType } | Sort-Object -Unique

    # --- "All Types" checkbox with Grid layout ---
    $AllCB = [System.Windows.Controls.CheckBox]::new()
    $AllCB.Tag = '__ALL__'
    $AllCB.IsChecked = $true
    $AllCB.Margin = [System.Windows.Thickness]::new(0,0,0,2)
    $AllCB.Padding = [System.Windows.Thickness]::new(0,2,0,2)
    $AllCB.VerticalContentAlignment = 'Center'
    $AllCB.HorizontalContentAlignment = 'Stretch'
    $AllCB.HorizontalAlignment = 'Stretch'

    $AllGrid = [System.Windows.Controls.Grid]::new()
    $ac0 = [System.Windows.Controls.ColumnDefinition]::new()
    $ac0.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $ac1 = [System.Windows.Controls.ColumnDefinition]::new()
    $ac1.Width = [System.Windows.GridLength]::Auto
    $AllGrid.ColumnDefinitions.Add($ac0) | Out-Null
    $AllGrid.ColumnDefinitions.Add($ac1) | Out-Null

    $AllTB = [System.Windows.Controls.TextBlock]::new()
    $AllTB.Text = "All Types"
    $AllTB.FontSize = 11
    $AllTB.FontWeight = [System.Windows.FontWeights]::SemiBold
    $AllTB.VerticalAlignment = 'Center'
    $AllTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextPrimary')
    [System.Windows.Controls.Grid]::SetColumn($AllTB, 0)

    $AllBadge = [System.Windows.Controls.Border]::new()
    $AllBadge.CornerRadius = [System.Windows.CornerRadius]::new(8)
    $AllBadge.Padding = [System.Windows.Thickness]::new(6,1,6,1)
    $AllBadge.VerticalAlignment = 'Center'
    $AllBadge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeHoverBg')
    $AllBadgeTB = [System.Windows.Controls.TextBlock]::new()
    $AllBadgeTB.Text = "$($Types.Count)"
    $AllBadgeTB.FontSize = 9
    $AllBadgeTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
    $AllBadge.Child = $AllBadgeTB
    [System.Windows.Controls.Grid]::SetColumn($AllBadge, 1)

    $AllGrid.Children.Add($AllTB) | Out-Null
    $AllGrid.Children.Add($AllBadge) | Out-Null
    $AllCB.Content = $AllGrid

    # Toggle all/none behavior
    $AllCB.Add_Checked({
        foreach ($ch in $pnlResourceTypeFilters.Children) {
            if ($ch -is [System.Windows.Controls.CheckBox] -and $ch.Tag -ne '__ALL__') {
                $ch.IsChecked = $false
            }
        }
        if ($Global:LatestDiff) { Populate-ChangesList -Diff $Global:LatestDiff }
    })
    $AllCB.Add_Unchecked({ })

    $pnlResourceTypeFilters.Children.Add($AllCB) | Out-Null

    # Separator
    $Sep = [System.Windows.Controls.Border]::new()
    $Sep.Height = 1
    $Sep.Margin = [System.Windows.Thickness]::new(0,4,0,6)
    $Sep.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeBorder')
    $pnlResourceTypeFilters.Children.Add($Sep) | Out-Null

    # Group by provider namespace
    $Grouped = @($Types | Group-Object { ($_ -split '/')[0] } | Sort-Object Name)

    foreach ($grp in $Grouped) {
        $ShortNS = ($grp.Name -replace '^Microsoft\.', '')

        # Provider section header (only when 2+ providers)
        if ($Grouped.Count -gt 1) {
            $Hdr = [System.Windows.Controls.TextBlock]::new()
            $Hdr.Text = $ShortNS.ToUpper()
            $Hdr.FontSize = 9
            $Hdr.FontWeight = [System.Windows.FontWeights]::SemiBold
            $Hdr.Margin = [System.Windows.Thickness]::new(2,6,0,3)
            $Hdr.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
            $pnlResourceTypeFilters.Children.Add($Hdr) | Out-Null
        }

        foreach ($rt in ($grp.Group | Sort-Object { ($_ -split '/')[-1] })) {
            $ShortName = ($rt -split '/')[-1]
            $TypeCount = @($Diff.changes | Where-Object { $_.resourceType -eq $rt }).Count

            $CB = [System.Windows.Controls.CheckBox]::new()
            $CB.Tag = $rt
            $CB.IsChecked = $false
            $CB.Margin = [System.Windows.Thickness]::new(0,0,0,1)
            $CB.Padding = [System.Windows.Thickness]::new(0,2,0,2)
            $CB.VerticalContentAlignment = 'Center'
            $CB.HorizontalContentAlignment = 'Stretch'
            $CB.HorizontalAlignment = 'Stretch'

            $CBGrid = [System.Windows.Controls.Grid]::new()
            $gc0 = [System.Windows.Controls.ColumnDefinition]::new()
            $gc0.Width = [System.Windows.GridLength]::Auto
            $gc1 = [System.Windows.Controls.ColumnDefinition]::new()
            $gc1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
            $gc2 = [System.Windows.Controls.ColumnDefinition]::new()
            $gc2.Width = [System.Windows.GridLength]::Auto
            $CBGrid.ColumnDefinitions.Add($gc0) | Out-Null
            $CBGrid.ColumnDefinitions.Add($gc1) | Out-Null
            $CBGrid.ColumnDefinitions.Add($gc2) | Out-Null

            # Type icon
            $ICO = [System.Windows.Controls.TextBlock]::new()
            $ICO.Text = (Get-ResourceTypeIcon -ResourceType $rt)
            $ICO.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
            $ICO.FontSize = 12
            $ICO.Margin = [System.Windows.Thickness]::new(0,0,6,0)
            $ICO.VerticalAlignment = 'Center'
            $ICO.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
            [System.Windows.Controls.Grid]::SetColumn($ICO, 0)

            # Short name
            $LBL = [System.Windows.Controls.TextBlock]::new()
            $LBL.Text = $ShortName
            $LBL.FontSize = 11
            $LBL.VerticalAlignment = 'Center'
            $LBL.TextTrimming = 'CharacterEllipsis'
            $LBL.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextSecondary')
            [System.Windows.Controls.Grid]::SetColumn($LBL, 1)

            # Count badge pill
            $Badge = [System.Windows.Controls.Border]::new()
            $Badge.CornerRadius = [System.Windows.CornerRadius]::new(8)
            $Badge.Padding = [System.Windows.Thickness]::new(5,0,5,0)
            $Badge.Margin = [System.Windows.Thickness]::new(4,0,0,0)
            $Badge.VerticalAlignment = 'Center'
            $Badge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeHoverBg')
            $BadgeTB = [System.Windows.Controls.TextBlock]::new()
            $BadgeTB.Text = "$TypeCount"
            $BadgeTB.FontSize = 9
            $BadgeTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
            $Badge.Child = $BadgeTB
            [System.Windows.Controls.Grid]::SetColumn($Badge, 2)

            $CBGrid.Children.Add($ICO) | Out-Null
            $CBGrid.Children.Add($LBL) | Out-Null
            $CBGrid.Children.Add($Badge) | Out-Null
            $CB.Content = $CBGrid

            # Filter events (no .GetNewClosure -- need script-scope access)
            $CB.Add_Checked({
                $AllItem = $pnlResourceTypeFilters.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_.Tag -eq '__ALL__' }
                if ($AllItem -and $AllItem.IsChecked) { $AllItem.IsChecked = $false }
                if ($Global:LatestDiff) { Populate-ChangesList -Diff $Global:LatestDiff }
            })
            $CB.Add_Unchecked({
                $AnyChecked = $false
                foreach ($ch in $pnlResourceTypeFilters.Children) {
                    if ($ch -is [System.Windows.Controls.CheckBox] -and $ch.Tag -ne '__ALL__' -and $ch.IsChecked) { $AnyChecked = $true; break }
                }
                if (-not $AnyChecked) {
                    $AllItem = $pnlResourceTypeFilters.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_.Tag -eq '__ALL__' }
                    if ($AllItem) { $AllItem.IsChecked = $true }
                }
                if ($Global:LatestDiff) { Populate-ChangesList -Diff $Global:LatestDiff }
            })

            $pnlResourceTypeFilters.Children.Add($CB) | Out-Null
        }
    }

    Write-DebugLog "[FILTER] Resource type filters populated: $($Types.Count) types ($($Grouped.Count) providers)" -Level 'DEBUG'
}
    <#
    .SYNOPSIS
        Returns a Segoe Fluent Icons character for a given Azure resource type.
    .DESCRIPTION
        Maps common Azure resource type strings to icon characters for visual display.
        Falls back to a generic resource icon for unrecognised types.
    .PARAMETER ResourceType
        Full Azure resource type string (e.g. 'Microsoft.Compute/virtualMachines').
    .OUTPUTS
        [string] Single icon character from Segoe Fluent Icons / MDL2 Assets.
    #>
function Get-ResourceTypeIcon {
    param([string]$ResourceType)
    $ns = ($ResourceType -split '/')[0].ToLowerInvariant()
    switch ($ns) {
        'microsoft.compute'               { return [char]0xE7F4 }  # Monitor
        'microsoft.storage'               { return [char]0xEDA2 }  # HardDrive
        'microsoft.network'               { return [char]0xE839 }  # Network
        'microsoft.web'                   { return [char]0xE774 }  # Globe
        'microsoft.keyvault'              { return [char]0xE72E }  # Lock
        'microsoft.desktopvirtualization' { return [char]0xE7F4 }  # Monitor
        'microsoft.managedidentity'       { return [char]0xE77B }  # Contact
        'microsoft.operationalinsights'   { return [char]0xE9D9 }  # Health
        'microsoft.insights'              { return [char]0xE9D9 }  # Health
        'microsoft.maintenance'           { return [char]0xE90F }  # Repair
        'microsoft.automation'            { return [char]0xE711 }  # Gear
        'microsoft.containerregistry'     { return [char]0xE7C1 }  # Page
        default                           { return [char]0xECAA }  # Cube
    }
}

    <#
    .SYNOPSIS
        Returns a theme resource key name for a provisioning state colour.
    .DESCRIPTION
        Maps common Azure provisioning states (Succeeded, Failed, Creating, Deleting,
        etc.) to theme colour resource keys (ThemeSuccess, ThemeError, ThemeWarning,
        ThemeAccentLight, ThemeTextMuted) for visual display in the Resources tree.
    .PARAMETER Status
        Provisioning state string from the Azure resource.
    .OUTPUTS
        [string] Theme resource key name.
    #>
function Get-StatusColor {
    param([string]$State)
    if (-not $State) { return $null }
    switch ($State.ToLowerInvariant()) {
        'succeeded'  { if ($Global:IsLightMode) { '#16A34A' } else { '#00C853' } }
        'failed'     { if ($Global:IsLightMode) { '#DC2626' } else { '#FF5000' } }
        'creating'   { if ($Global:IsLightMode) { '#0078D4' } else { '#60CDFF' } }
        'updating'   { if ($Global:IsLightMode) { '#EA580C' } else { '#F59E0B' } }
        'deleting'   { if ($Global:IsLightMode) { '#EA580C' } else { '#F59E0B' } }
        default      { if ($Global:IsLightMode) { '#808080' } else { '#8B8B93' } }
    }
}

    <#
    .SYNOPSIS
        Creates a styled TreeViewItem header for a resource type group.
    .DESCRIPTION
        Builds a WPF StackPanel containing the resource type icon, short type name,
        and a count badge. Used as the header of TreeViewItem groups in the Resources
        tab tree view.
    .PARAMETER ResourceType
        Full Azure resource type string.
    .PARAMETER Count
        Number of resources of this type.
    .OUTPUTS
        [System.Windows.Controls.StackPanel] The composed header element.
    #>
function New-ResourceTreeHeader {
    param(
        [string]$Name,
        [string]$TypeLabel,
        [string]$Location,
        [string]$ResourceType = '',
        [string]$ProvisioningState = '',
        [int]$Count = 0,
        [bool]$IsGroup = $false,
        [bool]$HasChildren = $false
    )

    if ($IsGroup) {
        # -- Resource Group: card-style header with count badge --
        $Card = [System.Windows.Controls.Border]::new()
        $Card.CornerRadius = [System.Windows.CornerRadius]::new(6)
        $Card.Padding = [System.Windows.Thickness]::new(10,7,12,7)
        $Card.Margin = [System.Windows.Thickness]::new(0,2,0,2)
        $Card.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeHoverBg')

        $Dock = [System.Windows.Controls.DockPanel]::new()
        $Dock.LastChildFill = $true
        $Card.Child = $Dock

        # Folder icon
        $Icon = [System.Windows.Controls.TextBlock]::new()
        $Icon.FontFamily = [System.Windows.Media.FontFamily]::new('Segoe MDL2 Assets')
        $Icon.FontSize = 15
        $Icon.Text = [char]0xF168
        $Icon.VerticalAlignment = 'Center'
        $Icon.Margin = [System.Windows.Thickness]::new(0,0,10,0)
        $Icon.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeAccent')
        [System.Windows.Controls.DockPanel]::SetDock($Icon, 'Left')
        $Dock.Children.Add($Icon) | Out-Null

        # Count badge pill (right-aligned)
        if ($Count -gt 0) {
            $Badge = [System.Windows.Controls.Border]::new()
            $Badge.CornerRadius = [System.Windows.CornerRadius]::new(10)
            $Badge.Padding = [System.Windows.Thickness]::new(8,2,8,2)
            $Badge.Margin = [System.Windows.Thickness]::new(8,0,0,0)
            $Badge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeAccentDim')
            $Badge.VerticalAlignment = 'Center'
            [System.Windows.Controls.DockPanel]::SetDock($Badge, 'Right')
            $CountTB = [System.Windows.Controls.TextBlock]::new()
            $CountTB.Text = "$Count"
            $CountTB.FontSize = 10.5
            $CountTB.FontWeight = [System.Windows.FontWeights]::SemiBold
            $CountTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeAccent')
            $Badge.Child = $CountTB
            $Dock.Children.Add($Badge) | Out-Null
        }

        # RG name (fill remaining space)
        $NameTB = [System.Windows.Controls.TextBlock]::new()
        $NameTB.Text = $Name
        $NameTB.FontSize = 12.5
        $NameTB.FontWeight = [System.Windows.FontWeights]::SemiBold
        $NameTB.VerticalAlignment = 'Center'
        $NameTB.TextTrimming = 'CharacterEllipsis'
        $NameTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextPrimary')
        $Dock.Children.Add($NameTB) | Out-Null

        return $Card
    }

    # -- Resource item: Robinhood-style row --
    $SP = [System.Windows.Controls.StackPanel]::new()
    $SP.Orientation = 'Horizontal'
    $SP.Margin = [System.Windows.Thickness]::new(0,3,0,3)

    # Type-specific icon
    $Icon = [System.Windows.Controls.TextBlock]::new()
    $Icon.FontFamily = [System.Windows.Media.FontFamily]::new('Segoe MDL2 Assets')
    $Icon.FontSize = if ($HasChildren) { 13 } else { 12 }
    $Icon.Text = Get-ResourceTypeIcon -ResourceType $ResourceType
    $Icon.VerticalAlignment = 'Center'
    $Icon.Margin = [System.Windows.Thickness]::new(0,0,8,0)
    $Icon.Width = 16
    $Icon.TextAlignment = 'Center'
    $IconFgKey = if ($HasChildren) { 'ThemeAccentLight' } else { 'ThemeTextMuted' }
    $Icon.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $IconFgKey)
    $SP.Children.Add($Icon) | Out-Null

    # Resource name
    $NameTB = [System.Windows.Controls.TextBlock]::new()
    $NameTB.Text = $Name
    $NameTB.FontSize = if ($HasChildren) { 12 } else { 11.5 }
    $NameTB.FontWeight = if ($HasChildren) { [System.Windows.FontWeights]::SemiBold } else { [System.Windows.FontWeights]::Normal }
    $NameTB.VerticalAlignment = 'Center'
    $NameTB.TextTrimming = 'CharacterEllipsis'
    $NameTB.MaxWidth = 280
    $NameTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextPrimary')
    $SP.Children.Add($NameTB) | Out-Null

    # Type pill badge
    if ($TypeLabel) {
        $TypePill = [System.Windows.Controls.Border]::new()
        $TypePill.CornerRadius = [System.Windows.CornerRadius]::new(4)
        $TypePill.Padding = [System.Windows.Thickness]::new(6,1,6,1)
        $TypePill.Margin = [System.Windows.Thickness]::new(8,0,0,0)
        $TypePill.VerticalAlignment = 'Center'
        $TypePill.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeHoverBg')
        $TypeTB = [System.Windows.Controls.TextBlock]::new()
        $TypeTB.Text = $TypeLabel
        $TypeTB.FontSize = 9.5
        $TypeTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
        $TypePill.Child = $TypeTB
        $SP.Children.Add($TypePill) | Out-Null
    }

    # Location
    if ($Location) {
        $LocTB = [System.Windows.Controls.TextBlock]::new()
        $LocTB.Text = $Location
        $LocTB.FontSize = 9.5
        $LocTB.Margin = [System.Windows.Thickness]::new(8,0,0,0)
        $LocTB.VerticalAlignment = 'Center'
        $LocTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
        $SP.Children.Add($LocTB) | Out-Null
    }

    # Provisioning state dot + label
    if ($ProvisioningState) {
        $StatColor = Get-StatusColor -State $ProvisioningState
        if ($StatColor) {
            $Dot = [System.Windows.Shapes.Ellipse]::new()
            $Dot.Width = 6; $Dot.Height = 6
            $Dot.Margin = [System.Windows.Thickness]::new(10,0,4,0)
            $Dot.VerticalAlignment = 'Center'
            $Dot.Fill = $Global:CachedBC.ConvertFromString($StatColor)
            $SP.Children.Add($Dot) | Out-Null

            $StatTB = [System.Windows.Controls.TextBlock]::new()
            $StatTB.Text = $ProvisioningState
            $StatTB.FontSize = 9.5
            $StatTB.VerticalAlignment = 'Center'
            $StatTB.Foreground = $Global:CachedBC.ConvertFromString($StatColor)
            $SP.Children.Add($StatTB) | Out-Null
        }
    }

    return $SP
}

    <#
    .SYNOPSIS
        Creates a TreeViewItem for an individual Azure resource.
    .DESCRIPTION
        Builds a header with the resource name and a provisioning state colour dot.
        Used as a leaf node inside a resource type group in the Resources tab tree view.
    .PARAMETER Resource
        A resource object from the snapshot with name and provisioningState properties.
    .OUTPUTS
        [System.Windows.Controls.TreeViewItem] Styled tree node for the resource.
    #>
function New-ResourceNode {
    param($Resource, $Children, $IdMap)
    $rid = $Resource.id.ToLowerInvariant()
    $TypeShort = ($Resource.type -split '/')[-1]
    $HasKids = $Children.ContainsKey($rid)

    $Node = [System.Windows.Controls.TreeViewItem]::new()
    $Node.Header = New-ResourceTreeHeader -Name $Resource.name -TypeLabel $TypeShort -Location $Resource.location `
        -ResourceType $Resource.type -ProvisioningState $Resource.provisioningState -HasChildren $HasKids
    $Node.IsExpanded = $HasKids
    $Node.Padding = [System.Windows.Thickness]::new(1)
    $Node.Tag = $Resource

    if ($HasKids) {
        foreach ($child in ($Children[$rid] | Sort-Object { ($_.type -split '/').Count }, name)) {
            $ChildNode = New-ResourceNode -Resource $child -Children $Children -IdMap $IdMap
            $Node.Items.Add($ChildNode) | Out-Null
        }
    }
    return $Node
}
    <#
    .SYNOPSIS
        Populates the Resources tab tree view with a hierarchical resource inventory.
    .DESCRIPTION
        Groups snapshot resources by type, applies the current resource type, resource
        group, and search text filters, then builds a TreeView with type group headers
        (containing icon, short type name, and count badge) and individual resource leaf
        nodes (with name and provisioning state dot). Updates the inventory count labels
        in both the sidebar and status bar.
    .PARAMETER Snapshot
        The snapshot object containing .resources array of Azure resource records.
    #>
function Populate-ResourcesList {
    param([object]$Snapshot)
    $treeResources.Items.Clear()
    if (-not $Snapshot -or -not $Snapshot.resources) { return }

    $FilterType = if ($cmbResourceType.SelectedItem) { $cmbResourceType.SelectedItem.ToString() } else { 'All' }
    $FilterRG   = if ($cmbResourceGroup.SelectedItem) { $cmbResourceGroup.SelectedItem.ToString() } else { 'All' }
    $FilterText = $txtResourceSearch.Text.Trim()

    # Apply filters
    $Filtered = @($Snapshot.resources | Where-Object {
        ($FilterType -eq 'All' -or $_.type -eq $FilterType) -and
        ($FilterRG   -eq 'All' -or $_.resourceGroup -eq $FilterRG) -and
        (-not $FilterText -or $_.name -like "*$FilterText*")
    })

    # Build lookup by lowercase resource ID
    $IdMap = @{}
    foreach ($r in $Filtered) { $IdMap[$r.id.ToLowerInvariant()] = $r }

    # Group by resource group
    $Groups = $Filtered | Group-Object -Property resourceGroup | Sort-Object Name

    $ShownCount = 0
    foreach ($grp in $Groups) {
        $RgNode = [System.Windows.Controls.TreeViewItem]::new()
        $RgNode.Header = New-ResourceTreeHeader -Name $grp.Name -Count $grp.Count -IsGroup $true
        $RgNode.IsExpanded = $true
        $RgNode.Padding = [System.Windows.Thickness]::new(2)

        # Sort resources: parents first (fewer type segments)
        $Sorted = $grp.Group | Sort-Object { ($_.type -split '/').Count }, name

        # Track which resource IDs have been placed as children
        $Placed = [System.Collections.Generic.HashSet[string]]::new()

        # Build parent-child map based on resource ID containment
        $Children = @{}
        foreach ($r in $Sorted) {
            $rid = $r.id.ToLowerInvariant()
            $parts = $rid -split '/'
            $foundParent = $false
            for ($i = $parts.Count - 3; $i -ge 7; $i -= 2) {
                $candidate = ($parts[0..$i]) -join '/'
                if ($IdMap.ContainsKey($candidate) -and $candidate -ne $rid) {
                    if (-not $Children.ContainsKey($candidate)) { $Children[$candidate] = [System.Collections.ArrayList]::new() }
                    $Children[$candidate].Add($r) | Out-Null
                    $Placed.Add($rid) | Out-Null
                    $foundParent = $true
                    break
                }
            }
        }

        # Add top-level resources (not placed as children)
        foreach ($r in ($Sorted | Where-Object { -not $Placed.Contains($_.id.ToLowerInvariant()) })) {
            $Node = New-ResourceNode -Resource $r -Children $Children -IdMap $IdMap
            $RgNode.Items.Add($Node) | Out-Null
            $ShownCount++
        }

        $treeResources.Items.Add($RgNode) | Out-Null
    }
    $lblResourceInventoryCount.Text = "$ShownCount / $($Snapshot.resources.Count) resources"
    Write-DebugLog "[UI] Resource tree populated: $ShownCount top-level, $($Filtered.Count) total (filters: type=$FilterType, rg=$FilterRG, text='$FilterText')" -Level 'DEBUG'
}

    <#
    .SYNOPSIS
        Refreshes the Resources tab filter combo boxes from a snapshot.
    .DESCRIPTION
        Extracts unique resource types and resource groups from the snapshot, sorts them
        alphabetically, and populates the cmbResourceType and cmbResourceGroup combos
        with an "All" option prepended. Preserves any previously selected filter value
        if it still exists in the new data.
    .PARAMETER Snapshot
        The snapshot object containing .resources array.
    #>
function Update-ResourceFilters {
    param([object]$Snapshot)
    Write-DebugLog "[UI] Update-ResourceFilters: hasSnapshot=$($null -ne $Snapshot)" -Level 'DEBUG'
    $cmbResourceType.Items.Clear()
    $cmbResourceGroup.Items.Clear()
    $cmbResourceType.Items.Add('All') | Out-Null
    $cmbResourceGroup.Items.Add('All') | Out-Null
    $cmbResourceType.SelectedIndex  = 0
    $cmbResourceGroup.SelectedIndex = 0

    if (-not $Snapshot -or -not $Snapshot.resources) { return }

    $Types = $Snapshot.resources | Select-Object -ExpandProperty type -Unique | Sort-Object
    foreach ($t in $Types) { $cmbResourceType.Items.Add($t) | Out-Null }

    $RGs = $Snapshot.resources | Select-Object -ExpandProperty resourceGroup -Unique | Sort-Object
    foreach ($rg in $RGs) { $cmbResourceGroup.Items.Add($rg) | Out-Null }
    Write-DebugLog "[UI] Resource filters updated: $($Types.Count) types, $($RGs.Count) resource groups" -Level 'DEBUG'
}

# ===============================================================================
# SECTION 17: EXPORT ENGINE
# Exports diff results as CSV (one row per changed property) or as a
# self-contained HTML report with dark/light theme toggle, sticky toolbar,
# summary stat cards, proportional progress bar, collapsible resource group
# sections, property-level diff tables, global search, and print CSS.
# ===============================================================================

    <#
    .SYNOPSIS
        Exports a diff result as a CSV file.
    .DESCRIPTION
        Generates a CSV with columns: ResourceName, ResourceType, ResourceGroup,
        ChangeType, Severity, Property, OldValue, NewValue. Modified resources emit
        one row per changed property. Added/Removed resources emit a single row with
        empty property/value columns. Values are double-quote escaped. UTF-8 encoded.
    .PARAMETER Diff
        The diff object from Compare-Snapshots.
    .PARAMETER Path
        Absolute file path for the output CSV.
    #>
function Export-ChangeReportCSV {
    param([object]$Diff, [string]$Path)
    Write-DebugLog "[EXPORT] CSV: $($Diff.changes.Count) changes -> $Path" -Level 'INFO'
    $Rows = @()
    $Rows += "ResourceName,ResourceType,ResourceGroup,ChangeType,Severity,Property,OldValue,NewValue"
    foreach ($c in $Diff.changes) {
        if ($c.details -and $c.details.Count -gt 0) {
            foreach ($prop in $c.details.Keys) {
                $old = ($c.details[$prop].from | ConvertTo-Json -Compress) -replace '"','""'
                $new = ($c.details[$prop].to   | ConvertTo-Json -Compress) -replace '"','""'
                $Rows += "`"$($c.resourceName)`",`"$($c.resourceType)`",`"$($c.resourceGroup)`",`"$($c.changeType)`",`"$($c.severity)`",`"$prop`",`"$old`",`"$new`""
            }
        } else {
            $Rows += "`"$($c.resourceName)`",`"$($c.resourceType)`",`"$($c.resourceGroup)`",`"$($c.changeType)`",`"$($c.severity)`",`"`",`"`",`"`""
        }
    }
    $Rows -join "`r`n" | Set-Content $Path -Encoding UTF8
    $SizeKB = [math]::Round((Get-Item $Path).Length / 1024, 1)
    Write-DebugLog "[EXPORT] CSV complete: $($Rows.Count - 1) rows, ${SizeKB} KB -> $Path" -Level 'SUCCESS'
}

    <#
    .SYNOPSIS
        Exports a diff result as a self-contained HTML report.
    .DESCRIPTION
        Generates a polished, single-file HTML report with embedded CSS and JavaScript.
        Features include:
        - Dark/light theme toggle with localStorage persistence
        - Sticky toolbar with subscription badge, date range, and search box
        - Summary stat cards (Total, Added, Removed, Modified, Unchanged)
        - Proportional progress bar showing the change distribution
        - Collapsible sections per change type, sub-grouped by resource group
        - Property-level diff tables for modified resources
        - Global search filtering across resource name, type, and group
        - Expand/collapse all buttons and print-friendly CSS
        HTML-encodes all user-supplied values to prevent XSS.
    .PARAMETER Diff
        The diff object from Compare-Snapshots.
    .PARAMETER Path
        Absolute file path for the output HTML.
    #>
function Export-ChangeReportHTML {
    param([object]$Diff, [string]$Path)
    Write-DebugLog "[EXPORT] HTML: $($Diff.changes.Count) changes -> $Path" -Level 'INFO'
    $s = $Diff.metadata.summary
    $_from = $Diff.metadata.fromSnapshot; $_fStr = $(if ($_from -is [datetime]) { $_from.ToString('yyyy-MM-dd') } else { ([string]$_from).Substring(0,10) })
    $_to = $Diff.metadata.toSnapshot; $_tStr = $(if ($_to -is [datetime]) { $_to.ToString('yyyy-MM-dd') } else { ([string]$_to).Substring(0,10) })
    $Total = $s.added + $s.removed + $s.modified
    $TotalResources = $Total + $s.unchanged
    $Added   = @($Diff.changes | Where-Object { $_.changeType -eq 'Added' })
    $Removed = @($Diff.changes | Where-Object { $_.changeType -eq 'Removed' })
    $Modified = @($Diff.changes | Where-Object { $_.changeType -eq 'Modified' })
    $enc = { param($v) [System.Net.WebUtility]::HtmlEncode("$v") }
    $subEnc = [System.Net.WebUtility]::HtmlEncode($Diff.metadata.subscriptionId)
    $html = [System.Text.StringBuilder]::new(32768)

    # ── Head + CSS ──
    [void]$html.Append(@"
<!DOCTYPE html>
<html lang="en" data-theme="dark"><head><meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Azure Change Report - $_fStr -> $_tStr</title>
<style>
:root {
  --font: 'Inter','Segoe UI Variable Display','Segoe UI',-apple-system,sans-serif;
  --bg: #0a0a0c; --bg2: #111113; --surface: #141416; --card: #1a1a1e; --card-hover: #222226;
  --card-border: #2a2a2e; --border: rgba(255,255,255,0.06); --border-strong: rgba(255,255,255,0.12);
  --text: #e4e4e7; --text-bright: #ffffff; --text-secondary: #a1a1aa; --text-dim: #71717a; --text-faint: #52525b;
  --accent: #60cdff; --accent-dim: rgba(96,205,255,0.1); --accent-text: #60cdff;
  --green: #00c853; --green-dim: rgba(0,200,83,0.12); --green-text: #4ade80;
  --red: #ff5000; --red-dim: rgba(255,80,0,0.12); --red-text: #fb923c;
  --orange: #f59e0b; --orange-dim: rgba(245,158,11,0.12); --orange-text: #fbbf24;
  --radius: 12px; --radius-sm: 8px; --radius-xs: 6px;
  --shadow: 0 1px 3px rgba(0,0,0,0.3);
}
[data-theme="light"] {
  --bg: #f8f9fa; --bg2: #ffffff; --surface: #f4f4f5; --card: #ffffff; --card-hover: #f0f0f2;
  --card-border: #e4e4e7; --border: rgba(0,0,0,0.08); --border-strong: rgba(0,0,0,0.15);
  --text: #18181b; --text-bright: #09090b; --text-secondary: #52525b; --text-dim: #71717a; --text-faint: #a1a1aa;
  --accent: #0078d4; --accent-dim: rgba(0,120,212,0.08); --accent-text: #0066b8;
  --green: #16a34a; --green-dim: rgba(22,163,74,0.08); --green-text: #15803d;
  --red: #dc2626; --red-dim: rgba(220,38,38,0.08); --red-text: #b91c1c;
  --orange: #ea580c; --orange-dim: rgba(234,88,12,0.08); --orange-text: #c2410c;
  --shadow: 0 1px 3px rgba(0,0,0,0.06), 0 4px 12px rgba(0,0,0,0.04);
}
*,*::before,*::after { margin:0; padding:0; box-sizing:border-box; }
html { scroll-behavior:smooth; }
body { font-family:var(--font); background:var(--bg); color:var(--text); line-height:1.5; }

/* Sticky toolbar */
.toolbar { position:sticky; top:0; z-index:100; background:var(--bg2); border-bottom:1px solid var(--border);
  padding:12px 32px; display:flex; align-items:center; gap:14px; backdrop-filter:blur(12px); }
.toolbar h1 { font-size:16px; font-weight:600; color:var(--text-bright); white-space:nowrap; display:flex; align-items:center; gap:8px; }
.toolbar .sub-badge { background:var(--accent-dim); color:var(--accent-text); padding:3px 10px;
  border-radius:20px; font-size:10px; font-weight:600; max-width:240px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.toolbar .date-badge { background:var(--card); border:1px solid var(--card-border); border-radius:var(--radius-xs);
  padding:4px 12px; font-size:11px; color:var(--text-secondary); display:flex; align-items:center; gap:6px; }
.toolbar .date-badge .arrow { color:var(--accent); font-weight:700; }
.toolbar .spacer { flex:1; }
.search-box { background:var(--card); border:1px solid var(--card-border); border-radius:var(--radius-sm);
  padding:6px 12px; color:var(--text); font-size:12px; width:220px; outline:none; font-family:var(--font); }
.search-box:focus { border-color:var(--accent); box-shadow:0 0 0 2px var(--accent-dim); }
.search-box::placeholder { color:var(--text-faint); }
.btn { background:var(--card); border:1px solid var(--card-border); border-radius:var(--radius-sm);
  padding:5px 12px; color:var(--text-secondary); font-size:11px; cursor:pointer; font-family:var(--font); transition:all 0.15s; }
.btn:hover { background:var(--card-hover); border-color:var(--border-strong); color:var(--text); }
.btn-icon { padding:5px 8px; font-size:15px; line-height:1; }

/* Content */
.container { max-width:1200px; margin:0 auto; padding:28px 40px 60px; }

/* Summary cards */
.summary { display:grid; grid-template-columns:repeat(5,1fr); gap:14px; margin-bottom:32px; }
.stat-card { background:var(--card); border:1px solid var(--card-border); border-radius:var(--radius); padding:18px 22px; position:relative; overflow:hidden; transition:border-color 0.2s; }
.stat-card:hover { border-color:var(--border-strong); }
.stat-card .accent-bar { position:absolute; top:0; left:0; right:0; height:3px; border-radius:var(--radius) var(--radius) 0 0; }
.stat-card .num { font-size:32px; font-weight:800; letter-spacing:-1px; line-height:1; }
.stat-card .label { font-size:10px; font-weight:600; text-transform:uppercase; letter-spacing:0.5px; color:var(--text-dim); margin-top:5px; }
.stat-card.total .num { color:var(--accent); }  .stat-card.total .accent-bar { background:var(--accent); }
.stat-card.added .num { color:var(--green); }    .stat-card.added .accent-bar { background:var(--green); }
.stat-card.removed .num { color:var(--red); }    .stat-card.removed .accent-bar { background:var(--red); }
.stat-card.modified .num { color:var(--orange); } .stat-card.modified .accent-bar { background:var(--orange); }
.stat-card.unchanged .num { color:var(--text-dim); } .stat-card.unchanged .accent-bar { background:var(--text-faint); }

/* Progress bar */
.progress-bar { display:flex; height:6px; border-radius:3px; overflow:hidden; margin-bottom:32px; background:var(--card); }
.progress-bar .seg { transition:width 0.6s ease; }
.seg-added { background:var(--green); } .seg-removed { background:var(--red); }
.seg-modified { background:var(--orange); } .seg-unchanged { background:var(--card-border); }

/* Collapsible sections */
details.section { margin-bottom:12px; border:1px solid var(--card-border); border-radius:var(--radius); overflow:hidden; }
details.section[open] { border-color:var(--border-strong); }
details.section > summary { display:flex; align-items:center; gap:10px; padding:12px 18px; cursor:pointer;
  background:var(--card); font-size:14px; font-weight:600; color:var(--text); user-select:none; list-style:none; }
details.section > summary::-webkit-details-marker { display:none; }
details.section > summary::marker { display:none; content:''; }
details.section > summary .chevron { font-size:10px; color:var(--text-dim); transition:transform 0.2s; margin-left:auto; }
details.section[open] > summary .chevron { transform:rotate(90deg); }
details.section > summary .section-dot { width:9px; height:9px; border-radius:50%; flex-shrink:0; }
details.section > summary .section-count { font-size:11px; font-weight:500; color:var(--text-dim); background:var(--surface);
  border:1px solid var(--card-border); border-radius:12px; padding:1px 9px; }
.section-body { padding:0; }

/* Change table */
.change-table { width:100%; border-collapse:separate; border-spacing:0; }
.change-table thead th { padding:8px 16px; font-size:10px; font-weight:600; text-transform:uppercase; letter-spacing:0.8px;
  color:var(--text-faint); border-bottom:1px solid var(--border); text-align:left; background:var(--surface); position:sticky; top:0; }
.change-table tbody tr { transition:background 0.1s; }
.change-table tbody tr:hover { background:var(--card-hover); }
.change-table td { padding:10px 16px; font-size:12px; border-bottom:1px solid var(--border); vertical-align:middle; }
.change-table td:first-child { width:36px; text-align:center; }

/* RG sub-group */
details.rg-group { border:none; margin:0; border-radius:0; }
details.rg-group > summary { display:flex; align-items:center; gap:8px; padding:7px 16px; cursor:pointer;
  background:var(--surface); font-size:11px; font-weight:600; color:var(--text-secondary); list-style:none;
  border-bottom:1px solid var(--border); user-select:none; }
details.rg-group > summary::-webkit-details-marker { display:none; }
details.rg-group > summary::marker { display:none; content:''; }
details.rg-group > summary .rg-chevron { font-size:9px; color:var(--text-dim); transition:transform 0.15s; }
details.rg-group[open] > summary .rg-chevron { transform:rotate(90deg); }
details.rg-group > summary .rg-icon { opacity:0.5; }
details.rg-group > summary .rg-count { font-size:10px; color:var(--text-faint); margin-left:auto; }

/* Row indicators */
.row-dot { width:7px; height:7px; border-radius:50%; display:inline-block; }
.row-dot.added { background:var(--green); box-shadow:0 0 6px rgba(0,200,83,0.4); }
.row-dot.removed { background:var(--red); box-shadow:0 0 6px rgba(255,80,0,0.4); }
.row-dot.modified { background:var(--orange); box-shadow:0 0 6px rgba(245,158,11,0.4); }
.res-name { font-weight:600; color:var(--text); }
.res-type { font-size:10px; color:var(--text-dim); margin-top:1px; }

/* Badge */
.badge { display:inline-flex; align-items:center; gap:4px; padding:2px 9px; border-radius:20px;
  font-size:9px; font-weight:700; letter-spacing:0.3px; text-transform:uppercase; }
.badge-added { background:var(--green-dim); color:var(--green-text); }
.badge-removed { background:var(--red-dim); color:var(--red-text); }
.badge-modified { background:var(--orange-dim); color:var(--orange-text); }

/* Details for modified resources */
.detail-row td { padding:0; }
.detail-pane { background:var(--surface); padding:10px 20px 10px 52px; }
.detail-pane table { width:100%; font-size:11px; }
.detail-pane th { color:var(--text-faint); font-size:9px; text-transform:uppercase; letter-spacing:0.5px; padding:3px 8px; border-bottom:1px solid var(--border); }
.detail-pane td { padding:4px 8px; border-bottom:1px solid var(--border); font-family:'Cascadia Mono','Consolas',monospace; font-size:11px; color:var(--text-secondary); word-break:break-all; }
.detail-pane .old-val { color:var(--red-text); }
.detail-pane .new-val { color:var(--green-text); }

/* Empty / no-results */
.empty { padding:32px; text-align:center; color:var(--text-dim); font-size:13px; }
.no-results { display:none; padding:24px; text-align:center; color:var(--text-dim); font-size:13px; }
.no-results.visible { display:block; }

/* Footer */
.footer { margin-top:40px; padding-top:20px; border-top:1px solid var(--border); display:flex; justify-content:space-between; align-items:center; color:var(--text-faint); font-size:11px; }
.footer-brand { display:flex; align-items:center; gap:8px; }
.footer-dot { width:6px; height:6px; border-radius:50%; background:var(--accent); }

/* Responsive */
@media (max-width:800px) {
  .summary { grid-template-columns:repeat(2,1fr); }
  .container { padding:16px; }
  .toolbar { padding:10px 16px; flex-wrap:wrap; }
  .search-box { width:100%; }
}
/* Print */
@media print {
  .toolbar { position:static; }
  .btn, .search-box { display:none; }
  body { background:#fff; color:#000; }
  .stat-card, details.section > summary, .section-body, th { background:#f5f5f5; border-color:#ddd; }
  .stat-card .num, details.section > summary { color:#000; }
  td { border-color:#ddd; }
  details { open:true; }
  details > .section-body { display:block !important; }
  details.rg-group > summary + * { display:block !important; }
}
</style></head><body>

<div class="toolbar">
  <h1>&#x1F50D; Azure Change Report</h1>
  <span class="sub-badge" title="$subEnc">$subEnc</span>
  <span class="date-badge"><span>$_fStr</span><span class="arrow">&#8594;</span><span>$_tStr</span></span>
  <span class="spacer"></span>
  <input type="text" class="search-box" id="globalSearch" placeholder="Search resources, types, groups..." />
  <button class="btn" onclick="expandAll()" title="Expand all sections">&#x25BC; Expand</button>
  <button class="btn" onclick="collapseAll()" title="Collapse all sections">&#x25B6; Collapse</button>
  <button class="btn btn-icon" id="themeToggle" onclick="toggleTheme()" title="Toggle dark/light mode">&#x263E;</button>
  <button class="btn" onclick="window.print()" title="Print report">&#x1F5A8;</button>
</div>

<div class="container">

  <div class="summary">
    <div class="stat-card total"><div class="accent-bar"></div><div class="num">$Total</div><div class="label">Total Changes</div></div>
    <div class="stat-card added"><div class="accent-bar"></div><div class="num">$($s.added)</div><div class="label">Added</div></div>
    <div class="stat-card removed"><div class="accent-bar"></div><div class="num">$($s.removed)</div><div class="label">Removed</div></div>
    <div class="stat-card modified"><div class="accent-bar"></div><div class="num">$($s.modified)</div><div class="label">Modified</div></div>
    <div class="stat-card unchanged"><div class="accent-bar"></div><div class="num">$($s.unchanged)</div><div class="label">Unchanged</div></div>
  </div>

  <div class="progress-bar">
"@)

    # Progress bar segments
    if ($TotalResources -gt 0) {
        $pctAdded    = [math]::Round(($s.added / $TotalResources) * 100, 1)
        $pctRemoved  = [math]::Round(($s.removed / $TotalResources) * 100, 1)
        $pctModified = [math]::Round(($s.modified / $TotalResources) * 100, 1)
        $pctUnchanged = 100 - $pctAdded - $pctRemoved - $pctModified
    } else { $pctAdded = 0; $pctRemoved = 0; $pctModified = 0; $pctUnchanged = 100 }

    [void]$html.Append("    <div class='seg seg-added' style='width:${pctAdded}%' title='Added: $($s.added)'></div>`n")
    [void]$html.Append("    <div class='seg seg-removed' style='width:${pctRemoved}%' title='Removed: $($s.removed)'></div>`n")
    [void]$html.Append("    <div class='seg seg-modified' style='width:${pctModified}%' title='Modified: $($s.modified)'></div>`n")
    [void]$html.Append("    <div class='seg seg-unchanged' style='width:${pctUnchanged}%' title='Unchanged: $($s.unchanged)'></div>`n")
    [void]$html.Append("  </div>`n`n")

    # ── Change sections ──
    $Sections = @(
        @{ Label = 'Removed';  Items = $Removed;  Color = 'var(--red)';    Class = 'removed' }
        @{ Label = 'Added';    Items = $Added;     Color = 'var(--green)';  Class = 'added' }
        @{ Label = 'Modified'; Items = $Modified;  Color = 'var(--orange)'; Class = 'modified' }
    )

    [void]$html.Append("  <div id='sectionsContainer'>`n")

    foreach ($sec in $Sections) {
        if ($sec.Items.Count -eq 0) { continue }

        [void]$html.Append(@"
  <details class="section" open>
  <summary>
    <span class="section-dot" style="background:$($sec.Color)"></span>
    $($sec.Label)
    <span class="section-count">$($sec.Items.Count)</span>
    <span class="chevron">&#x25B6;</span>
  </summary>
  <div class="section-body">
"@)

        # Group by resource group
        $SecByRG = @{}
        foreach ($c in $sec.Items) {
            $rg = $c.resourceGroup
            if (-not $SecByRG.ContainsKey($rg)) { $SecByRG[$rg] = @() }
            $SecByRG[$rg] += $c
        }

        foreach ($rg in ($SecByRG.Keys | Sort-Object)) {
            $rgEnc = & $enc $rg
            $rgCount = $SecByRG[$rg].Count

            [void]$html.Append(@"
    <details class="rg-group" open>
    <summary>
      <span class="rg-chevron">&#x25B6;</span>
      <span class="rg-icon">&#x1F4C1;</span>
      $rgEnc
      <span class="rg-count">$rgCount</span>
    </summary>
    <table class="change-table">
    <thead><tr><th></th><th>Resource</th><th>Type</th><th>Change</th></tr></thead>
    <tbody>
"@)

            foreach ($c in $SecByRG[$rg]) {
                $nameEnc = & $enc $c.resourceName
                $shortType = & $enc (($c.resourceType -split '/')[-1])
                $fullType = & $enc $c.resourceType
                [void]$html.Append("      <tr data-name=`"$nameEnc`" data-type=`"$fullType`" data-rg=`"$rgEnc`"><td><span class='row-dot $($sec.Class)'></span></td><td><div class='res-name'>$nameEnc</div></td><td><div class='res-type' title='$fullType'>$shortType</div></td><td><span class='badge badge-$($sec.Class)'>$($sec.Label)</span></td></tr>`n")

                # Property diff details for modified resources
                if ($sec.Class -eq 'modified' -and $c.details -and $c.details.Count -gt 0) {
                    [void]$html.Append("      <tr class='detail-row'><td colspan='4'><div class='detail-pane'><table><thead><tr><th>Property</th><th>Previous</th><th>Current</th></tr></thead><tbody>`n")
                    foreach ($propKey in ($c.details.Keys | Sort-Object)) {
                        $d = $c.details[$propKey]
                        $propEnc = & $enc $propKey
                        $oldVal = & $enc "$($d.from)"
                        $newVal = & $enc "$($d.to)"
                        [void]$html.Append("        <tr><td>$propEnc</td><td class='old-val'>$oldVal</td><td class='new-val'>$newVal</td></tr>`n")
                    }
                    [void]$html.Append("      </tbody></table></div></td></tr>`n")
                }
            }

            [void]$html.Append("    </tbody></table>`n    </details>`n")
        }

        [void]$html.Append("  </div>`n  </details>`n`n")
    }

    [void]$html.Append("  </div>`n")
    [void]$html.Append("  <div class='no-results' id='noResults'>No resources match your search.</div>`n")

    # Footer + JS
    [void]$html.Append(@"

  <div class="footer">
    <div class="footer-brand"><div class="footer-dot"></div>Azure Change Tracker v$($Global:AppVersion)</div>
    <div>Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm')</div>
  </div>
</div>

<script>
function toggleTheme() {
  var html = document.documentElement;
  var isDark = html.getAttribute('data-theme') === 'dark';
  html.setAttribute('data-theme', isDark ? 'light' : 'dark');
  document.getElementById('themeToggle').textContent = isDark ? '\u2600' : '\u263E';
  try { localStorage.setItem('azct-theme', isDark ? 'light' : 'dark'); } catch(e) {}
}
(function() {
  try {
    var saved = localStorage.getItem('azct-theme');
    if (saved === 'light') { document.documentElement.setAttribute('data-theme','light'); document.getElementById('themeToggle').textContent = '\u2600'; }
  } catch(e) {}
})();

function expandAll() { document.querySelectorAll('details').forEach(function(d) { d.open = true; }); }
function collapseAll() { document.querySelectorAll('details.section').forEach(function(d) { d.open = false; }); }

document.getElementById('globalSearch').addEventListener('input', function() {
  var q = this.value.toLowerCase().trim();
  var anyVisible = false;
  document.querySelectorAll('.change-table tbody tr:not(.detail-row)').forEach(function(row) {
    if (!q) { row.style.display = ''; var det = row.nextElementSibling; if (det && det.classList.contains('detail-row')) det.style.display = ''; anyVisible = true; return; }
    var txt = (row.dataset.name || '') + ' ' + (row.dataset.type || '') + ' ' + (row.dataset.rg || '');
    var match = txt.toLowerCase().indexOf(q) !== -1;
    row.style.display = match ? '' : 'none';
    var det = row.nextElementSibling;
    if (det && det.classList.contains('detail-row')) det.style.display = match ? '' : 'none';
    if (match) anyVisible = true;
  });
  document.querySelectorAll('details.rg-group').forEach(function(rg) {
    var rows = rg.querySelectorAll('tbody tr:not(.detail-row):not([style*="display: none"])');
    var cnt = rg.querySelector('.rg-count');
    if (cnt) cnt.textContent = rows.length;
    rg.style.display = rows.length > 0 ? '' : 'none';
    if (q && rows.length > 0) rg.open = true;
  });
  document.querySelectorAll('details.section').forEach(function(sec) {
    var rows = sec.querySelectorAll('tbody tr:not(.detail-row):not([style*="display: none"])');
    var cnt = sec.querySelector('.section-count');
    if (cnt) cnt.textContent = rows.length;
    sec.style.display = rows.length > 0 ? '' : 'none';
    if (q && rows.length > 0) sec.open = true;
  });
  var noRes = document.getElementById('noResults');
  if (noRes) noRes.className = (q && !anyVisible) ? 'no-results visible' : 'no-results';
});
</script>

</body></html>
"@)

    $html.ToString() | Set-Content $Path -Encoding UTF8
    $SizeKB = [math]::Round((Get-Item $Path).Length / 1024, 1)
    Write-DebugLog "[EXPORT] HTML complete: ${SizeKB} KB -> $Path" -Level 'SUCCESS'
}

# ===============================================================================
# ===============================================================================
# SECTION 17b: AUTH UI HELPERS
# ===============================================================================

    <#
    .SYNOPSIS
        Updates the auth bar and subscription combo for a successful Azure connection.
    .DESCRIPTION
        Populates the subscription ComboBox with all enabled subscriptions, sets the auth
        status dot to green, updates the status label with the account ID, shows the
        logout button, and auto-selects the previously saved subscription (or the current
        context subscription). Unlocks the 'first_login' achievement.
    .PARAMETER AccountId
        The Azure account email or service principal ID.
    .PARAMETER SubName
        Display name of the current subscription.
    .PARAMETER SubId
        GUID of the current subscription.
    .PARAMETER Subscriptions
        Array of subscription objects with Id and Name properties.
    #>
function Set-AuthUIConnected {
    param(
        [string]$AccountId,
        [string]$SubName,
        [string]$SubId,
        [array]$Subscriptions
    )
    Write-DebugLog "[AUTH] Set-AuthUIConnected: $AccountId ($($Subscriptions.Count) subs)" -Level 'DEBUG'

    $Global:Subscriptions = @($Subscriptions)
    $cmbSubscription.Items.Clear()
    foreach ($sub in $Global:Subscriptions) {
        $Item = New-Object System.Windows.Controls.ComboBoxItem
        $Item.Content = [string]$sub.Name
        $Item.Tag     = [string]$sub.Id
        $cmbSubscription.Items.Add($Item) | Out-Null
    }

    $authDot.Fill = $Global:CachedBC.ConvertFromString(
        $(if ($Global:IsLightMode) { '#16A34A' } else { '#00C853' }))
    $lblAuthStatus.Text = "Connected: $AccountId"
    $lblAuthStatus.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextPrimary')
    $lblStatus.Text = "Connected"

    $cmbSubscription.Visibility = 'Visible'
    $btnLogin.Visibility  = 'Collapsed'
    $btnLogout.Visibility = 'Visible'

    if ($Global:Subscriptions.Count -gt 0) {
        $Idx = 0
        if ($Global:CurrentSubId) {
            for ($i = 0; $i -lt $Global:Subscriptions.Count; $i++) {
                if ($Global:Subscriptions[$i].Id -eq $Global:CurrentSubId) { $Idx = $i; break }
            }
        } elseif ($SubId) {
            for ($i = 0; $i -lt $Global:Subscriptions.Count; $i++) {
                if ($Global:Subscriptions[$i].Id -eq $SubId) { $Idx = $i; break }
            }
        }
        $cmbSubscription.SelectedIndex = $Idx
    }

    Write-DebugLog "[AUTH] $($Global:Subscriptions.Count) subscriptions available" -Level 'SUCCESS'
    Unlock-Achievement 'first_login'
}

    <#
    .SYNOPSIS
        Resets the auth bar to the disconnected state.
    .DESCRIPTION
        Clears the subscription combo and global subscription state, sets the auth dot
        to red, updates status labels, and shows the login button.
    .PARAMETER Reason
        Status text to display (default: 'Not connected').
    #>
function Set-AuthUIDisconnected {
    param([string]$Reason = 'Not connected')
    Write-DebugLog "[AUTH] Set-AuthUIDisconnected: $Reason" -Level 'DEBUG'
    $cmbSubscription.Items.Clear()
    $Global:Subscriptions  = @()
    $Global:CurrentSubId   = ''
    $Global:CurrentSubName = ''
    $authDot.SetResourceReference([System.Windows.Shapes.Ellipse]::FillProperty, 'ThemeError')
    $lblAuthStatus.Text = $Reason
    $lblAuthStatus.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
    $lblStatus.Text     = "Ready"
    $lblStatusSub.Text  = ""
    $cmbSubscription.Visibility = 'Collapsed'
    $btnLogin.Visibility  = 'Visible'
    $btnLogout.Visibility = 'Collapsed'
}

# ===============================================================================
# SECTION 18: EVENT WIRING
# Connects XAML elements to PowerShell event handlers: title bar buttons,
# theme toggle, settings controls, icon rail navigation, tab headers,
# help dialog, sidebar toggle, debug log panel, Azure auth (login/logout),
# subscription selector, snapshot/compare/export/baseline buttons, filters,
# and the volatile change toggle.
# ===============================================================================

# --- Title bar -----------------------------------------------------------------
$btnMinimize.Add_Click({
    Write-DebugLog "[UI] Minimize clicked" -Level 'DEBUG'
    $Window.WindowState = 'Minimized'
})
$btnMaximize.Add_Click({
    $NewState = if ($Window.WindowState -eq 'Maximized') { 'Normal' } else { 'Maximized' }
    Write-DebugLog "[UI] Maximize toggled -> $NewState" -Level 'DEBUG'
    $Window.WindowState = $NewState
})
$btnClose.Add_Click({
    Write-DebugLog "[UI] Close clicked - saving prefs and shutting down" -Level 'INFO'
    Save-UserPrefs
    $Window.Close()
})

# --- Theme toggle --------------------------------------------------------------
$btnThemeToggle.Add_Click({
    $SetLight = -not $Global:IsLightMode
    Write-DebugLog "[UI] Theme toggle clicked -> $(if ($SetLight) { 'Light' } else { 'Dark' })" -Level 'DEBUG'
    $chkDarkMode.IsChecked = (-not $SetLight)
    & $ApplyTheme $SetLight
    Save-UserPrefs
    Unlock-Achievement 'theme_toggle'
})
$chkDarkMode.Add_Checked({
    Write-DebugLog "[UI] Dark mode checkbox CHECKED" -Level 'DEBUG'
    & $ApplyTheme $false; Save-UserPrefs
})
$chkDarkMode.Add_Unchecked({
    Write-DebugLog "[UI] Dark mode checkbox UNCHECKED" -Level 'DEBUG'
    & $ApplyTheme $true;  Save-UserPrefs
})


# --- Settings field immediate-save handlers ----------------------------------
$chkAutoSnapshot.Add_Checked({   Write-DebugLog "[SETTINGS] Auto-snapshot ENABLED" -Level 'DEBUG'; Save-UserPrefs })
$chkAutoSnapshot.Add_Unchecked({ Write-DebugLog "[SETTINGS] Auto-snapshot DISABLED" -Level 'DEBUG'; Save-UserPrefs })
$txtRetentionDays.Add_LostFocus({ Write-DebugLog "[SETTINGS] Retention days changed: $($txtRetentionDays.Text)" -Level 'DEBUG'; Save-UserPrefs })
$txtExcludedTypes.Add_LostFocus({ Write-DebugLog "[SETTINGS] Excluded types changed" -Level 'DEBUG'; Save-UserPrefs })
$chkRetentionEnabled.Add_Checked({
    Write-DebugLog "[SETTINGS] Retention cleanup ENABLED" -Level 'DEBUG'
    $pnlRetentionDays.Visibility = 'Visible'
    Save-UserPrefs
})
$chkRetentionEnabled.Add_Unchecked({
    Write-DebugLog "[SETTINGS] Retention cleanup DISABLED" -Level 'DEBUG'
    $pnlRetentionDays.Visibility = 'Collapsed'
    Save-UserPrefs
})

# --- Icon rail navigation -----------------------------------------------------
$railDashboard.Add_Click({ Write-DebugLog "[NAV] Rail: Dashboard clicked" -Level 'DEBUG'; Switch-Tab 'Dashboard' })
$railChanges.Add_Click({   Write-DebugLog "[NAV] Rail: Changes clicked" -Level 'DEBUG'; Switch-Tab 'Changes' })
$railTimeline.Add_Click({  Write-DebugLog "[NAV] Rail: Timeline clicked" -Level 'DEBUG'; Switch-Tab 'Timeline' })
$railResources.Add_Click({ Write-DebugLog "[NAV] Rail: Resources clicked" -Level 'DEBUG'; Switch-Tab 'Resources' })
$railSettings.Add_Click({  Write-DebugLog "[NAV] Rail: Settings clicked" -Level 'DEBUG'; Switch-Tab 'Settings' })

# --- Tab headers ---------------------------------------------------------------
$tabDashboard.Add_MouseLeftButtonUp({ Switch-Tab 'Dashboard' })
$tabChanges.Add_MouseLeftButtonUp({   Switch-Tab 'Changes' })
$tabTimeline.Add_MouseLeftButtonUp({  Switch-Tab 'Timeline' })
$tabResources.Add_MouseLeftButtonUp({ Switch-Tab 'Resources' })
$tabSettings.Add_MouseLeftButtonUp({  Switch-Tab 'Settings' })


# --- Help button --------------------------------------------------------------
$btnHelp.Add_Click({
    Write-DebugLog "[UI] Help button clicked" -Level 'DEBUG'
    Show-ThemedDialog -Title "About Azure Change Tracker" \
        -Message "Azure Change Tracker v0.1.0`n`nTrack Azure resource changes with daily JSON snapshots, visual diff engine, and export reports.`n`nBuilt with PowerShell + WPF." \
        -Icon ([string]([char]0xE897)) \
        -Buttons @( @{ Text='Close'; IsAccent=$true; Result='OK' } )
})
# --- Hamburger (sidebar toggle) -----------------------------------------------
$btnHamburger.Add_Click({
    if ($Global:LeftPanelVisible) {
        Write-DebugLog "[UI] Hamburger: collapsing sidebar" -Level 'DEBUG'
        $colLeftPanel.Width = [System.Windows.GridLength]::new(0)
        $Global:LeftPanelVisible = $false
    } else {
        Write-DebugLog "[UI] Hamburger: expanding sidebar (260px)" -Level 'DEBUG'
        $colLeftPanel.Width = [System.Windows.GridLength]::new(260)
        $Global:LeftPanelVisible = $true
    }
    Unlock-Achievement 'sidebar_toggle'
})

# --- Debug log panel ----------------------------------------------------------
$railOutput.Add_Click({
    if ($rowBottomPanel.Height.Value -gt 0) {
        $Global:BottomSavedHeight = $rowBottomPanel.Height.Value
        Write-DebugLog "[UI] Output panel collapsed (saved height=$($Global:BottomSavedHeight))" -Level 'DEBUG'
        $rowBottomPanel.Height = [System.Windows.GridLength]::new(0)
    } else {
        Write-DebugLog "[UI] Output panel restored (height=$($Global:BottomSavedHeight))" -Level 'DEBUG'
        $rowBottomPanel.Height = [System.Windows.GridLength]::new($Global:BottomSavedHeight)
    }
})
$btnClearLog.Add_Click({
    $paraLog.Inlines.Clear()
    $Global:DebugLineCount = 0
    Write-DebugLog "[UI] Log cleared (was $($Global:DebugLineCount) lines)" -Level 'INFO'
    Unlock-Achievement 'log_clear'
})
$btnToggleBottomSize.Add_Click({
    if ($Global:BottomExpanded) {
        Write-DebugLog "[UI] Log panel expanded to 500px" -Level 'DEBUG'
        $rowBottomPanel.Height = [System.Windows.GridLength]::new(500)
        $icoToggleBottomSize.Text = [string]([char]0xE70D)
        $Global:BottomExpanded = $false
    } else {
        Write-DebugLog "[UI] Log panel shrunk to 160px" -Level 'DEBUG'
        $rowBottomPanel.Height = [System.Windows.GridLength]::new(160)
        $icoToggleBottomSize.Text = [string]([char]0xE70E)
        $Global:BottomExpanded = $true
    }
})
$btnHideBottom.Add_Click({
    Write-DebugLog "[UI] Hide bottom panel clicked" -Level 'DEBUG'
    $Global:BottomSavedHeight = $rowBottomPanel.Height.Value
    $rowBottomPanel.Height = [System.Windows.GridLength]::new(0)
})

# --- Auth ----------------------------------------------------------------------
$btnLogin.Add_Click({
    Write-DebugLog "[AUTH] === Login button clicked ===" -Level 'DEBUG'
    $lblAuthStatus.Text = "Signing in..."
    $lblStatus.Text     = "Signing in..."
    $btnLogin.IsEnabled = $false
    Start-Shimmer

    # Check for cached context first (on UI thread - fast)
    $CachedCtx = Get-AzContext -ErrorAction SilentlyContinue
    $HasCache  = ($CachedCtx -and $CachedCtx.Account)

    Start-BackgroundWork -Variables @{
        HasCachedContext = $HasCache
        CachedSubId     = if ($HasCache) { $CachedCtx.Subscription.Id } else { '' }
        CachedTenantId  = if ($HasCache) { $CachedCtx.Tenant.Id } else { '' }
    } -Work {
        $ProgressPreference = 'SilentlyContinue'
        $NeedLogin = $true

        if ($HasCachedContext) {
            try {
                # Validate cached token with a lightweight call
                Get-AzSubscription -SubscriptionId $CachedSubId -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
                $NeedLogin = $false
            } catch {
                # Token expired - fall through to interactive
            }
        }

        if ($NeedLogin) {
            try { Update-AzConfig -EnableLoginByWam $false -ErrorAction SilentlyContinue | Out-Null } catch { }
            Connect-AzAccount -ErrorAction Stop | Out-Null
        }

        $Ctx = Get-AzContext -ErrorAction SilentlyContinue
        $TenantFilter = if ($Ctx.Tenant.Id) { $Ctx.Tenant.Id } elseif ($CachedTenantId) { $CachedTenantId } else { $null }
        $SubParams = @{ ErrorAction = 'SilentlyContinue'; WarningAction = 'SilentlyContinue' }
        if ($TenantFilter) { $SubParams['TenantId'] = $TenantFilter }
        $Subs = Get-AzSubscription @SubParams |
                Where-Object { $_.State -eq 'Enabled' } |
                Sort-Object Name
        return $Subs
    } -OnComplete {
        param($Results, $Errors)
        Stop-Shimmer
        $btnLogin.IsEnabled = $true
        Write-DebugLog "[AuthCB] OnComplete ENTERED (Results=$($Results.Count), Errors=$($Errors.Count))" -Level 'DEBUG'
        if ($Errors.Count -gt 0) {
            Write-DebugLog "[AuthCB] Sign-in FAILED: $($Errors[0])" -Level 'ERROR'
            if ($Errors[0].Exception) {
                Write-DebugLog "[AuthCB] InnerException: $($Errors[0].Exception.InnerException)" -Level 'ERROR'
                Write-DebugLog "[AuthCB] StackTrace: $($Errors[0].ScriptStackTrace)" -Level 'DEBUG'
            }
            Show-Toast -Message "Sign-in failed: $($Errors[0].ToString())" -Type 'Error'
            Set-AuthUIDisconnected -Reason "Not connected"
            return
        }

        $Ctx = Get-AzContext
        Set-AuthUIConnected -AccountId $Ctx.Account.Id -SubName $Ctx.Subscription.Name -SubId $Ctx.Subscription.Id -Subscriptions @($Results)
        Write-DebugLog "[AuthCB] Signed in: $($Ctx.Account.Id), tenant=$($Ctx.Tenant.Id)" -Level 'DEBUG'
        Show-Toast -Message "Signed in with $($Global:Subscriptions.Count) subscriptions" -Type 'Success'
    }
})

$btnLogout.Add_Click({
    Write-DebugLog "[AUTH] === Logout button clicked ===" -Level 'DEBUG'
    Start-BackgroundWork -Variables @{} -Work {
        try { Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null } catch { }
    } -OnComplete {
        Write-DebugLog "[AUTH] Disconnect-AzAccount completed" -Level 'DEBUG'
    }
    Set-AuthUIDisconnected -Reason "Not connected"
    Write-DebugLog "[AUTH] Signed out" -Level 'INFO'
    Show-Toast -Message "Disconnected from Azure" -Type 'Info'
})

# --- Subscription change ------------------------------------------------------
$cmbSubscription.Add_SelectionChanged({
    $idx = $cmbSubscription.SelectedIndex
    if ($idx -lt 0 -or $idx -ge $Global:Subscriptions.Count) { return }
    $Sub = $Global:Subscriptions[$idx]
    $Global:CurrentSubId   = $Sub.Id
    $Global:CurrentSubName = $Sub.Name
    $lblStatusSub.Text = $Sub.Name
    $lblStatusResourceCount.Text = ''

    Write-DebugLog "[SUB] Selected: $($Sub.Name) ($($Sub.Id))" -Level 'INFO'

    # Track unique subscriptions for multi_sub achievement
    if (-not $Global:UsedSubIds) { $Global:UsedSubIds = [System.Collections.Generic.HashSet[string]]::new() }
    $Global:UsedSubIds.Add($Sub.Id) | Out-Null
    if ($Global:UsedSubIds.Count -ge 2) { Unlock-Achievement 'multi_sub' }

    Write-DebugLog "[SUB] Setting Az context in background..." -Level 'DEBUG'

    # Set Az context in background
    Start-BackgroundWork -Variables @{ SubId = $Sub.Id } -Work {
        Set-AzContext -SubscriptionId $SubId -ErrorAction Stop | Out-Null
    } -OnComplete {
        param($R, $E)
        if ($E.Count -gt 0) {
            Write-DebugLog "[SubCB] Set-AzContext FAILED: $($E[0])" -Level 'ERROR'
        } else {
            Write-DebugLog "[SubCB] Az context set to: $($Global:CurrentSubName)" -Level 'SUCCESS'
        }
    }

    Update-SnapshotList
    if ($Global:SnapshotList.Count -gt 0) {
        Write-DebugLog "[SUB] Loading latest snapshot: $($Global:SnapshotList[0].Name)" -Level 'DEBUG'
        $Latest = Get-Content $Global:SnapshotList[0].FullName -Raw -Encoding UTF8 | ConvertFrom-Json
        $Global:LatestSnapshot = $Latest
        Update-DashboardCards
        Update-ResourceFilters -Snapshot $Latest
        Populate-ResourcesList -Snapshot $Latest
        $lblSidebarResourceCount.Text = $Latest.metadata.resourceCount.ToString()
        Write-DebugLog "[SUB] Loaded $($Latest.metadata.resourceCount) resources from latest snapshot" -Level 'DEBUG'
    } else {
        Write-DebugLog "[SUB] No snapshots found for this subscription" -Level 'DEBUG'
    }
})

# --- Snapshot Now --------------------------------------------------------------
$btnSnapshotNow.Add_Click({
    Write-DebugLog "[UI] === Snapshot Now button clicked ===" -Level 'DEBUG'
    if (-not $Global:CurrentSubId) {
        Write-DebugLog "[SNAPSHOT] Blocked: no subscription selected" -Level 'WARN'
        Show-Toast -Message "Please sign in and select a subscription first" -Type 'Warning'
        return
    }

    # Check if today's snapshot already exists
    $TodayFile = Join-Path (Join-Path $Global:SnapshotDir "sub_$($Global:CurrentSubId -replace '-','')") ((Get-Date).ToString('yyyy-MM-dd') + '.json')
    if (Test-Path $TodayFile) {
        Write-DebugLog "[SNAPSHOT] Today's snapshot already exists: $TodayFile" -Level 'WARN'
        $Confirm = Show-ThemedDialog -Title 'Snapshot Exists' `
            -Message "A snapshot for today ($((Get-Date).ToString('yyyy-MM-dd'))) already exists.`n`nDo you want to overwrite it?" `
            -Icon ([string]([char]0xE7BA)) `
            -IconColor 'ThemeWarning' `
            -Buttons @(
                @{ Text = 'Cancel'; IsAccent = $false; Result = 'Cancel' },
                @{ Text = 'Overwrite'; IsAccent = $true; Result = 'OK' }
            )
        if ($Confirm -ne 'OK') {
            Write-DebugLog "[SNAPSHOT] User cancelled overwrite" -Level 'DEBUG'
            return
        }
        Write-DebugLog "[SNAPSHOT] User confirmed overwrite" -Level 'DEBUG'
    }

    New-ResourceSnapshot -SubscriptionId $Global:CurrentSubId -SubscriptionName $Global:CurrentSubName
})

# --- Run Compare --------------------------------------------------------------
$btnRunCompare.Add_Click({
    Write-DebugLog "[UI] === Run Compare button clicked ===" -Level 'DEBUG'
    if ($cmbDiffFrom.SelectedIndex -lt 0 -or $cmbDiffTo.SelectedIndex -lt 0) {
        Write-DebugLog "[DIFF] Blocked: no snapshots selected (from=$($cmbDiffFrom.SelectedIndex), to=$($cmbDiffTo.SelectedIndex))" -Level 'WARN'
        Show-Toast -Message "Select two snapshots to compare" -Type 'Warning'
        return
    }
    $FromFile = $Global:SnapshotList[$cmbDiffFrom.SelectedIndex]
    $ToFile   = $Global:SnapshotList[$cmbDiffTo.SelectedIndex]

    if ($FromFile.FullName -eq $ToFile.FullName) {
        Show-Toast -Message "Cannot compare a snapshot with itself" -Type 'Warning'
        return
    }

    Write-DebugLog "[DIFF] Loading FROM: $($FromFile.Name) ($([math]::Round($FromFile.Length/1024))KB)" -Level 'INFO'
    Write-DebugLog "[DIFF] Loading TO:   $($ToFile.Name) ($([math]::Round($ToFile.Length/1024))KB)" -Level 'INFO'
    $FromSnap = Get-Content $FromFile.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
    $ToSnap   = Get-Content $ToFile.FullName   -Raw -Encoding UTF8 | ConvertFrom-Json
    Write-DebugLog "[DIFF] Parsed: FROM=$($FromSnap.resources.Count) resources, TO=$($ToSnap.resources.Count) resources" -Level 'DEBUG'

    $Diff = Compare-Snapshots -FromSnapshot $FromSnap -ToSnapshot $ToSnap

    # Save diff to disk so Timeline can read it later
    $DiffSubDir = Join-Path $Global:DiffDir "sub_$($Global:CurrentSubId -replace '-','')"
    if (-not (Test-Path $DiffSubDir)) { New-Item -Path $DiffSubDir -ItemType Directory -Force | Out-Null }
    $_fName = $FromFile.BaseName; $_tName = $ToFile.BaseName
    $DiffFileName = "${_fName}_vs_${_tName}.json"
    $DiffFilePath = Join-Path $DiffSubDir $DiffFileName
    $Diff | ConvertTo-Json -Depth 20 | Set-Content $DiffFilePath -Encoding UTF8
    Write-DebugLog "[DIFF] Saved diff to: $DiffFilePath" -Level 'SUCCESS'

    Update-DashboardCards
    Populate-ResourceTypeFilters -Diff $Diff
    Populate-ChangesList -Diff $Diff
    Populate-Timeline
    Switch-Tab 'Changes'

    $Total = $Diff.metadata.summary.added + $Diff.metadata.summary.removed + $Diff.metadata.summary.modified
    $ToastType = if ($Total -eq 0) { 'Success' } else { 'Info' }
    Show-Toast -Message "Comparison complete: $Total changes found" -Type $ToastType
    Unlock-Achievement 'first_compare'
    if ($Total -eq 0) { Unlock-Achievement 'zero_changes' }
    $CriticalCount = @($Diff.changes | Where-Object { $_.severity -eq 'Critical' }).Count
    if ($CriticalCount -gt 0) { Unlock-Achievement 'critical_found' }
})

# --- Filters ------------------------------------------------------------------
# Resource type filter is driven by checkboxes populated dynamically after compare
$cmbFilterChangeType.Add_SelectionChanged({
    Write-DebugLog "[FILTER] ChangeType changed: $($cmbFilterChangeType.SelectedItem)" -Level 'DEBUG'
    if ($Global:LatestDiff) { Populate-ChangesList -Diff $Global:LatestDiff }
})
$txtFilterSearch.Add_TextChanged({
    Write-DebugLog "[FILTER] Search text: '$($txtFilterSearch.Text)'" -Level 'DEBUG'
    if ($Global:LatestDiff) { Populate-ChangesList -Diff $Global:LatestDiff }
})
    <#
    .SYNOPSIS
        Syncs the volatile-change toggle visual state with $Global:HideVolatileEnabled.
    .DESCRIPTION
        Sets the custom toggle track background (accent when ON, subtle when OFF) and
        moves the thumb dot to the right (ON) or left (OFF) position.
    #>
function Update-VolatileToggleVisual {
    if ($Global:HideVolatileEnabled) {
        $chkHideVolatileTrack.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeAccentLight')
        $chkHideVolatileThumb.HorizontalAlignment = 'Right'
    } else {
        $chkHideVolatileTrack.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeBorderSubtle')
        $chkHideVolatileThumb.HorizontalAlignment = 'Left'
    }
}
$pnlHideVolatileToggle.Add_MouseLeftButtonDown({
    $Global:HideVolatileEnabled = -not $Global:HideVolatileEnabled
    Write-DebugLog "[FILTER] Hide volatile: $( if ($Global:HideVolatileEnabled) { 'ON' } else { 'OFF' } )" -Level 'DEBUG'
    Update-VolatileToggleVisual
    if ($Global:LatestDiff) { Populate-ChangesList -Diff $Global:LatestDiff }
    Save-UserPrefs
})

$cmbResourceType.Add_SelectionChanged({
    Write-DebugLog "[FILTER] ResourceType changed: $($cmbResourceType.SelectedItem)" -Level 'DEBUG'
    if ($Global:LatestSnapshot) { Populate-ResourcesList -Snapshot $Global:LatestSnapshot }
})
$cmbResourceGroup.Add_SelectionChanged({
    Write-DebugLog "[FILTER] ResourceGroup changed: $($cmbResourceGroup.SelectedItem)" -Level 'DEBUG'
    if ($Global:LatestSnapshot) { Populate-ResourcesList -Snapshot $Global:LatestSnapshot }
})
$txtResourceSearch.Add_TextChanged({
    Write-DebugLog "[FILTER] Resource search: '$($txtResourceSearch.Text)'" -Level 'DEBUG'
    if ($Global:LatestSnapshot) { Populate-ResourcesList -Snapshot $Global:LatestSnapshot }
})

# --- Export -------------------------------------------------------------------
$btnExportReport.Add_Click({
    Write-DebugLog "[UI] === Export Report button clicked ===" -Level 'DEBUG'
    if (-not $Global:LatestDiff) {
        Write-DebugLog "[EXPORT] Blocked: no diff available" -Level 'WARN'
        Show-Toast -Message "Run a comparison first before exporting" -Type 'Warning'
        return
    }

    $SaveDlg = [Microsoft.Win32.SaveFileDialog]::new()
    $SaveDlg.InitialDirectory = $Global:ReportDir
    $SaveDlg.Filter = "HTML Report (*.html)|*.html|CSV Report (*.csv)|*.csv|JSON (*.json)|*.json"
    $SaveDlg.FileName = "change-report_$(Get-Date -Format 'yyyy-MM-dd')"

    if ($SaveDlg.ShowDialog()) {
        $ext = [System.IO.Path]::GetExtension($SaveDlg.FileName).ToLower()
        Write-DebugLog "[EXPORT] User selected: $($SaveDlg.FileName) (format=$ext)" -Level 'DEBUG'
        switch ($ext) {
            '.html' { Export-ChangeReportHTML -Diff $Global:LatestDiff -Path $SaveDlg.FileName }
            '.csv'  { Export-ChangeReportCSV  -Diff $Global:LatestDiff -Path $SaveDlg.FileName }
            '.json' {
                $Global:LatestDiff | ConvertTo-Json -Depth 20 | Set-Content $SaveDlg.FileName -Encoding UTF8
                $JsonSizeKB = [math]::Round((Get-Item $SaveDlg.FileName).Length / 1024, 1)
                Write-DebugLog "[EXPORT] JSON complete: ${JsonSizeKB} KB -> $($SaveDlg.FileName)" -Level 'SUCCESS'
            }
        }
        Show-Toast -Message "Report exported: $([System.IO.Path]::GetFileName($SaveDlg.FileName))" -Type 'Success'
        Unlock-Achievement 'first_export'
    }
})

# --- Save Baseline ------------------------------------------------------------
$btnSaveBaseline.Add_Click({
    Write-DebugLog "[UI] === Save Baseline button clicked ===" -Level 'DEBUG'

    # Allow saving from LatestSnapshot (in-memory) or the most recent snapshot file on disk
    $SnapshotToSave = $Global:LatestSnapshot
    if (-not $SnapshotToSave -and $Global:SnapshotList.Count -gt 0) {
        $MostRecent = $Global:SnapshotList[0]
        Write-DebugLog "[BASELINE] No in-memory snapshot, loading most recent from disk: $($MostRecent.Name)" -Level 'DEBUG'
        try { $SnapshotToSave = Get-Content $MostRecent.FullName -Raw -Encoding UTF8 | ConvertFrom-Json }
        catch { Write-DebugLog "[BASELINE] Failed to load: $($_.Exception.Message)" -Level 'WARN' }
    }
    if (-not $SnapshotToSave) {
        Write-DebugLog "[BASELINE] Blocked: no snapshot available" -Level 'WARN'
        Show-Toast -Message "Take a snapshot first" -Type 'Warning'
        return
    }
    $Result = Show-ThemedDialog -Title "Save Baseline" `
                  -Message "Enter a name for this baseline snapshot:" `
                  -Icon ([string]([char]0xE74E)) `
                  -InputPrompt "Baseline Name" `
                  -Buttons @(
                      @{ Text='Save'; IsAccent=$true; Result='Save' }
                      @{ Text='Cancel'; IsAccent=$false; Result='Cancel' }
                  )

    if ($Result.Result -eq 'Save' -and $Result.Input.Trim()) {
        $SafeName = $Result.Input.Trim() -replace '[^\w\-]', '_'
        $BaselinePath = Join-Path $Global:BaselineDir "$SafeName.json"
        Write-DebugLog "[BASELINE] Saving: '$SafeName' -> $BaselinePath" -Level 'DEBUG'
        $SnapshotToSave | ConvertTo-Json -Depth 30 | Set-Content $BaselinePath -Encoding UTF8
        $BaselineSizeKB = [math]::Round((Get-Item $BaselinePath).Length / 1024, 1)
        Write-DebugLog "[BASELINE] Saved: $SafeName (${BaselineSizeKB} KB, $($SnapshotToSave.metadata.resourceCount) resources)" -Level 'SUCCESS'
        Show-Toast -Message "Baseline saved: $SafeName" -Type 'Success'
        Update-SnapshotList
        Unlock-Achievement 'first_baseline'
    }
})

# ===============================================================================
# SECTION 19: FILTER COMBOBOX INITIALIZATION
# ===============================================================================

# Resource type filter checkboxes are populated dynamically after each compare

$cmbFilterChangeType.Items.Add('All')      | Out-Null
$cmbFilterChangeType.Items.Add('Added')    | Out-Null
$cmbFilterChangeType.Items.Add('Removed')  | Out-Null
$cmbFilterChangeType.Items.Add('Modified') | Out-Null
$cmbFilterChangeType.SelectedIndex = 0

# Timeline type filter
$cmbTimelineType.Items.Add('All')      | Out-Null
$cmbTimelineType.Items.Add('Added')    | Out-Null
$cmbTimelineType.Items.Add('Removed')  | Out-Null
$cmbTimelineType.Items.Add('Modified') | Out-Null
$cmbTimelineType.SelectedIndex = 0

# Timeline filter event handlers
$txtTimelineSearch.Add_TextChanged({ Populate-Timeline })
$cmbTimelineType.Add_SelectionChanged({ Populate-Timeline })

# ===============================================================================
# SECTION 19b: ACHIEVEMENT SYSTEM
# ===============================================================================

$Global:AchievementDefs = @(
    @{ Id='first_login';      Icon=[char]0xE72E;  Label='First Login';         Tip='Sign in to Azure for the first time' }
    @{ Id='first_snapshot';   Icon=[char]0xE786;  Label='First Snapshot';      Tip='Take your first resource snapshot' }
    @{ Id='first_compare';    Icon=[char]0xE8F1;  Label='First Compare';       Tip='Run your first snapshot comparison' }
    @{ Id='first_export';     Icon=[char]0xE74E;  Label='First Export';        Tip='Export a change report' }
    @{ Id='first_baseline';   Icon=[char]0xE735;  Label='First Baseline';      Tip='Save a baseline snapshot' }
    @{ Id='five_snapshots';   Icon=[char]0xE780;  Label='Collector';           Tip='Take 5 snapshots' }
    @{ Id='ten_snapshots';    Icon=[char]0xE787;  Label='Archivist';           Tip='Take 10 snapshots' }
    @{ Id='theme_toggle';     Icon=[char]0xE793;  Label='Painter';             Tip='Toggle between dark and light themes' }
    @{ Id='zero_changes';     Icon=[char]0xE73E;  Label='All Clear';           Tip='Run a comparison with zero changes' }
    @{ Id='over_100';         Icon=[char]0xE7B8;  Label='Resource Rich';       Tip='Snapshot 100+ resources' }
    @{ Id='over_500';         Icon=[char]0xE838;  Label='Enterprise Scale';    Tip='Snapshot 500+ resources' }
    @{ Id='critical_found';   Icon=[char]0xEA39;  Label='Watchdog';            Tip='Detect a critical severity change' }
    @{ Id='multi_sub';        Icon=[char]0xE774;  Label='Multi-Tenant';        Tip='Switch between 2+ subscriptions' }
    @{ Id='sidebar_toggle';   Icon=[char]0xE700;  Label='Space Maker';         Tip='Toggle the sidebar' }
    @{ Id='log_clear';        Icon=[char]0xE74D;  Label='Clean Slate';         Tip='Clear the debug log' }
)

    <#
    .SYNOPSIS
        Renders all achievement badges in the sidebar panel.
    .DESCRIPTION
        Loads achievements from achievements.json, then creates a 32x32 badge Border
        for each of the 15 defined achievements. Earned badges get an accent background
        and coloured icon; unearned badges are dimmed with 50% opacity. Updates the
        achievement count label (e.g. "7/15").
    #>
function Initialize-AchievementBadges {
    Write-DebugLog "[ACHIEVE] Initializing $($Global:AchievementDefs.Count) achievement badges" -Level 'DEBUG'
    $pnlAchievements.Children.Clear()

    # Load saved achievements
    if (Test-Path $Global:AchievementsFile) {
        try {
            $Global:Achievements = @{}
            $Saved = Get-Content $Global:AchievementsFile -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($prop in $Saved.PSObject.Properties) {
                $Global:Achievements[$prop.Name] = $prop.Value
            }
            Write-DebugLog "[ACHIEVE] Loaded $($Global:Achievements.Count) earned achievements" -Level 'DEBUG'
        } catch {
            Write-DebugLog "[ACHIEVE] Failed to load: $($_.Exception.Message)" -Level 'WARN'
        }
    }

    foreach ($Def in $Global:AchievementDefs) {
        $Earned = $Global:Achievements.ContainsKey($Def.Id)

        $Badge = [System.Windows.Controls.Border]::new()
        $Badge.Width  = 32; $Badge.Height = 32
        $Badge.CornerRadius = [System.Windows.CornerRadius]::new(6)
        $Badge.Margin = [System.Windows.Thickness]::new(0,0,4,4)
        $Badge.ToolTip = "$($Def.Label) - $($Def.Tip)"
        $Badge.Tag = $Def.Id

        $IconTB = [System.Windows.Controls.TextBlock]::new()
        $IconTB.Text = $Def.Icon
        $IconTB.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
        $IconTB.FontSize = 14
        $IconTB.HorizontalAlignment = 'Center'
        $IconTB.VerticalAlignment = 'Center'

        if ($Earned) {
            $Badge.Background = $Global:CachedBC.ConvertFromString('#200078D4')
            $Badge.BorderBrush = $Global:CachedBC.ConvertFromString('#0078D4')
            $Badge.BorderThickness = [System.Windows.Thickness]::new(1)
            $IconTB.Foreground = $Global:CachedBC.ConvertFromString('#60CDFF')
        } else {
            $Badge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeSurfaceBg')
            $Badge.BorderThickness = [System.Windows.Thickness]::new(0)
            $IconTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDisabled')
            $Badge.Opacity = 0.5
        }

        $Badge.Child = $IconTB
        $pnlAchievements.Children.Add($Badge) | Out-Null
    }

    $EarnedCount = $Global:Achievements.Count
    $lblAchievementCount.Text = "$EarnedCount/$($Global:AchievementDefs.Count)"
    Write-DebugLog "[ACHIEVE] Badges rendered: $EarnedCount/$($Global:AchievementDefs.Count) earned" -Level 'DEBUG'
}

    <#
    .SYNOPSIS
        Unlocks an achievement by ID, persists it, and shows a success toast.
    .DESCRIPTION
        Short-circuits if the achievement is already earned. Otherwise, records the
        current ISO 8601 timestamp, merges with on-disk data to avoid overwriting
        achievements from previous sessions, saves to achievements.json, shows a
        "Achievement Unlocked" success toast, and re-renders all badges.
    .PARAMETER AchievementId
        The unique achievement identifier (e.g. 'first_snapshot', 'critical_found').
    #>
function Unlock-Achievement {
    param([string]$AchievementId)

    # Ensure in-memory hashtable has disk data (prevents overwriting on first call)
    if ($Global:Achievements.Count -eq 0 -and (Test-Path $Global:AchievementsFile)) {
        try {
            $Saved = Get-Content $Global:AchievementsFile -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($prop in $Saved.PSObject.Properties) {
                $Global:Achievements[$prop.Name] = $prop.Value
            }
            Write-DebugLog "[ACHIEVE] Lazy-loaded $($Global:Achievements.Count) achievements from disk" -Level 'DEBUG'
        } catch {
            Write-DebugLog "[ACHIEVE] Lazy-load failed: $($_.Exception.Message)" -Level 'WARN'
        }
    }

    if ($Global:Achievements.ContainsKey($AchievementId)) { return }

    $Def = $Global:AchievementDefs | Where-Object { $_.Id -eq $AchievementId }
    if (-not $Def) { return }

    $Global:Achievements[$AchievementId] = (Get-Date).ToString('o')
    Write-DebugLog "[ACHIEVE] UNLOCKED: $($Def.Label) ($AchievementId)" -Level 'SUCCESS'
    Show-Toast -Message "Achievement Unlocked: $($Def.Label)!" -Type 'Success'

    # Save -- merge with disk to avoid overwriting achievements from a previous session
    try {
        if (Test-Path $Global:AchievementsFile) {
            $OnDisk = Get-Content $Global:AchievementsFile -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($prop in $OnDisk.PSObject.Properties) {
                if (-not $Global:Achievements.ContainsKey($prop.Name)) {
                    $Global:Achievements[$prop.Name] = $prop.Value
                }
            }
        }
        $Global:Achievements | ConvertTo-Json -Depth 2 | Set-Content $Global:AchievementsFile -Encoding UTF8
    } catch {
        Write-DebugLog "[ACHIEVE] Save failed: $($_.Exception.Message)" -Level 'WARN'
    }

    # Re-render badges
    Initialize-AchievementBadges
}

# ===============================================================================
# SECTION 20: DISPATCHER TIMER (Queue -> UI Bridge)
# A 50 ms DispatcherTimer polls $Global:BgJobs for completed runspaces.
# When a job's AsyncResult.IsCompleted is true, it calls EndInvoke to
# collect results and errors, then executes the OnComplete callback on
# the UI thread. Also logs periodic "still running" heartbeats.
# ===============================================================================

$Timer = New-Object System.Windows.Threading.DispatcherTimer
$Timer.Interval = [TimeSpan]::FromMilliseconds($Script:TIMER_INTERVAL_MS)
$Timer.Add_Tick({
    if ($Global:TimerProcessing) { return }
    $Global:TimerProcessing = $true
    try {
        $TickNow = Get-Date

        # Background job completion polling
        for ($bi = $Global:BgJobs.Count - 1; $bi -ge 0; $bi--) {
            $Job = $Global:BgJobs[$bi]
            $ElapsedSec = ($TickNow - $Job.StartedAt).TotalSeconds
            if ($Job.AsyncResult.IsCompleted) {
                Write-DebugLog "BgJob[$bi]: COMPLETED after $([math]::Round($ElapsedSec,1))s" -Level 'DEBUG'
                try {
                    $BgResult = $Job.PS.EndInvoke($Job.AsyncResult)
                    $BgErrors = @($Job.PS.Streams.Error)
                    if ($BgErrors.Count -gt 0) {
                        foreach ($e in $BgErrors) {
                            Write-DebugLog "BgJob[$bi]: ERROR: $($e.ToString())" -Level 'ERROR'
                        }
                    }
                    & $Job.OnComplete $BgResult $BgErrors $Job.Context
                } catch {
                    Write-DebugLog "BgJob[$bi]: callback EXCEPTION: $($_.Exception.Message)" -Level 'ERROR'
                    Write-DebugLog "BgJob[$bi]: StackTrace: $($_.ScriptStackTrace)" -Level 'ERROR'
                }
                try { $Job.PS.Dispose() } catch { }
                try { $Job.Runspace.Dispose() } catch { }
                $Global:BgJobs.RemoveAt($bi)
            } else {
                $Bucket = [math]::Floor($ElapsedSec / 5)
                if ($Bucket -ge 1 -and -not $Job.LastBucket) { $Job.LastBucket = 0 }
                if ($Bucket -ge 1 -and $Bucket -ne $Job.LastBucket) {
                    $Job.LastBucket = $Bucket
                    Write-DebugLog "BgJob[$bi]: still running ($([math]::Round($ElapsedSec,0))s)..." -Level 'DEBUG'
                }
            }
        }

    } finally {
        $Global:TimerProcessing = $false
    }
})
$Timer.Start()

# ===============================================================================
# SECTION 21: STARTUP
# Three-step initialisation sequence shown on the splash overlay:
#   Step 1: Verify prerequisite modules (Az.Accounts, Az.ResourceGraph, Az.Resources)
#   Step 2: Validate cached Azure token and auto-connect if valid
#   Step 3: Load achievements, apply preferences, run retention cleanup
# The splash auto-dismisses after 3-4 seconds via Start-SplashDismissTimer.
# ===============================================================================

Write-DebugLog "=======================================================" -Level 'INFO'
Write-DebugLog "  Azure Change Tracker v$($Global:AppVersion)" -Level 'INFO'
Write-DebugLog "  Root: $Global:Root" -Level 'INFO'
Write-DebugLog "=======================================================" -Level 'INFO'

# Show splash with version
$lblSplashVersion.Text = "v$($Global:AppVersion)"
$lblSplashStatus.Text = "Starting up..."
$pnlSplash.Visibility  = 'Visible'
$Global:WindowRendered = $false
$Global:SplashDismissQueued = $false
$Script:SplashDismissDelayMs = 3000

$Window.Add_ContentRendered({
    $Global:WindowRendered = $true
    Write-DebugLog "[SPLASH] Window ContentRendered fired" -Level 'DEBUG'
    if ($Global:SplashDismissQueued) {
        $ST = [System.Windows.Threading.DispatcherTimer]::new()
        $ST.Interval = [TimeSpan]::FromMilliseconds($Script:SplashDismissDelayMs)
        $STRef = $ST
        $HideSplashFn = { Hide-Splash }
        $ST.Add_Tick({ $STRef.Stop(); & $HideSplashFn }.GetNewClosure())
        $ST.Start()
    }
})

# Initialize default tab (hidden behind splash)
Switch-Tab 'Dashboard'

# Load user preferences early (for theme)
Load-UserPrefs
Invoke-RetentionCleanup

# Apply initial theme (before splash is visible)
& $ApplyTheme $Global:IsLightMode

# --- Step 1: Check prerequisites ----------------------------------------------
Set-SplashStep -Step 1 -Status 'running' -Text "Checking prerequisites..."
$lblSplashStatus.Text = "Verifying required PowerShell modules..."
Write-DebugLog "[STARTUP] Step 1: Checking prerequisites..." -Level 'INFO'

$ModResults = Test-Prerequisites
$AllOK = ($ModResults | Where-Object { -not $_.Available }).Count -eq 0

foreach ($M in $ModResults) {
    if ($M.Available) {
        Write-DebugLog "  [OK] $($M.Name) v$($M.Installed)" -Level 'SUCCESS'
    } else {
        Write-DebugLog "  [MISSING] $($M.Name) (need >= $($M.Required)) - $($M.Purpose)" -Level 'WARN'
    }
}

if ($AllOK) {
    Set-SplashStep -Step 1 -Status 'done' -Text "Prerequisites OK"
    Write-DebugLog "[STARTUP] Step 1 complete: all modules present" -Level 'SUCCESS'

} else {
    Set-SplashStep -Step 1 -Status 'error' -Text "Missing modules detected"
    $Missing = $ModResults | Where-Object { -not $_.Available }
    $Cmd = Get-RemediationCommand -ModuleResults $ModResults
    Write-DebugLog "[STARTUP] Step 1 FAILED - install: $Cmd" -Level 'WARN'
    Show-Toast -Message "Missing modules: $($Missing.Name -join ', '). See debug log." -Type 'Warning' -DurationMs 8000
}

# --- Step 2: Validate Azure credentials ---------------------------------------
if ($AllOK) {
    Set-SplashStep -Step 2 -Status 'running' -Text "Validating Azure credentials..."
    $lblSplashStatus.Text = "Looking for cached Azure session..."
    Write-DebugLog "[STARTUP] Step 2: Checking cached Azure context..." -Level 'INFO'

    $CachedCtx = Get-AzContext -ErrorAction SilentlyContinue
    if ($CachedCtx -and $CachedCtx.Account) {
        $lblSplashStatus.Text = "Validating token for $($CachedCtx.Account.Id)..."
        Write-DebugLog "[STARTUP] Cached context: $($CachedCtx.Account.Id), validating in background..." -Level 'INFO'

        Start-BackgroundWork -Variables @{
            AccountId = $CachedCtx.Account.Id
            SubName   = $CachedCtx.Subscription.Name
            SubId     = $CachedCtx.Subscription.Id
            TenantId  = $CachedCtx.Tenant.Id
        } -Work {
            $R = @{ TokenValid = $false; AccountId = $AccountId; SubName = $SubName; SubId = $SubId; Subs = @(); Error = '' }
            try {
                $ProgressPreference = 'SilentlyContinue'
                Get-AzSubscription -SubscriptionId $SubId -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
                $R.TokenValid = $true
                $R.Subs = @(Get-AzSubscription -TenantId $TenantId -ErrorAction SilentlyContinue -WarningAction SilentlyContinue |
                            Where-Object { $_.State -eq 'Enabled' } |
                            Sort-Object Name)
            } catch {
                $R.Error = $_.Exception.Message
            }
            return $R
        } -OnComplete {
            param($Results, $Errors)
            $R = if ($Results -and $Results.Count -gt 0) { $Results[-1] } else { @{ TokenValid = $false; Error = 'No result' } }
            if ($R.TokenValid) {
                Set-SplashStep -Step 2 -Status 'done' -Text "Connected: $($R.AccountId)"
                Write-DebugLog "[STARTUP] Step 2 complete: auto-connected as $($R.AccountId)" -Level 'SUCCESS'
                Set-AuthUIConnected -AccountId $R.AccountId -SubName $R.SubName -SubId $R.SubId -Subscriptions $R.Subs

                # --- Step 3: Initialize workspace --------------------------
                Set-SplashStep -Step 3 -Status 'running' -Text "Initializing workspace..."
                $lblSplashStatus.Text = "Loading achievements and workspace data..."
                Write-DebugLog "[STARTUP] Step 3: Initializing workspace..." -Level 'INFO'
                Initialize-AchievementBadges
                Set-SplashStep -Step 3 -Status 'done' -Text "Workspace ready"
                Write-DebugLog "[STARTUP] Step 3 complete" -Level 'SUCCESS'

                $lblSplashStatus.Text = "Ready - welcome back!"
                # Auto-dismiss splash after a brief pause

                # Auto-snapshot on launch if enabled
                if ($chkAutoSnapshot.IsChecked) {
                    Write-DebugLog "[STARTUP] Auto-snapshot enabled - triggering snapshot" -Level 'INFO'
                    $lblSplashStatus.Text = "Taking auto-snapshot..."
                    New-ResourceSnapshot
                }

                # Purge old snapshots if retention is set
                $RetDays = 0
                if ([int]::TryParse($txtRetentionDays.Text, [ref]$RetDays) -and $RetDays -gt 0) {
                    $SubDir = Join-Path $Global:SnapshotDir "sub_$($Global:CurrentSubId -replace '-','')"
                    if (Test-Path $SubDir) {
                        $Cutoff = (Get-Date).AddDays(-$RetDays)
                        $Old = Get-ChildItem $SubDir -Filter '*.json' | Where-Object { $_.LastWriteTime -lt $Cutoff }
                        if ($Old) {
                            Write-DebugLog "[STARTUP] Purging $($Old.Count) snapshots older than $RetDays days" -Level 'INFO'
                            $Old | Remove-Item -Force
                        }
                    }
                }
                Start-SplashDismissTimer -DelayMs 3000
            } else {
                Set-SplashStep -Step 2 -Status 'error' -Text "Token expired - sign in required"
                Write-DebugLog "[STARTUP] Step 2: Token expired - $($R.Error)" -Level 'WARN'
                $lblAuthStatus.Text = "Token expired - click Sign In"

                # Still run Step 3
                Set-SplashStep -Step 3 -Status 'running' -Text "Initializing workspace..."
                $lblSplashStatus.Text = "Loading workspace..."
                Initialize-AchievementBadges
                Set-SplashStep -Step 3 -Status 'done' -Text "Workspace ready"

                $lblSplashStatus.Text = "Sign in to get started"
                Start-SplashDismissTimer -DelayMs 3500
            }
        }
    } else {
        Set-SplashStep -Step 2 -Status 'skipped' -Text "No cached session - sign in required"
        Write-DebugLog "[STARTUP] Step 2: No cached context found" -Level 'DEBUG'

        # --- Step 3: Initialize workspace (no auth) -----------------------
        Set-SplashStep -Step 3 -Status 'running' -Text "Initializing workspace..."
        $lblSplashStatus.Text = "Loading workspace..."
        Write-DebugLog "[STARTUP] Step 3: Initializing workspace..." -Level 'INFO'
        Initialize-AchievementBadges
        Set-SplashStep -Step 3 -Status 'done' -Text "Workspace ready"
        Write-DebugLog "[STARTUP] Step 3 complete" -Level 'SUCCESS'

        $lblSplashStatus.Text = "Sign in to get started"
        Start-SplashDismissTimer -DelayMs 3500
    }
} else {
    # Prerequisites failed - skip auth, but still init workspace
    Set-SplashStep -Step 2 -Status 'skipped' -Text "Skipped (missing modules)"
    Set-SplashStep -Step 3 -Status 'running' -Text "Initializing workspace..."
    $lblSplashStatus.Text = "Loading workspace..."
    Initialize-AchievementBadges
    Set-SplashStep -Step 3 -Status 'done' -Text "Workspace ready"

    $lblSplashStatus.Text = "Install missing modules to continue"
    Start-SplashDismissTimer -DelayMs 4000
}

# Show window

# Save prefs on any close method (Alt+F4, taskbar, etc.)
$Window.Add_Closing({
    Write-DebugLog "[UI] Window closing - saving prefs" -Level 'INFO'
    Save-UserPrefs
})

$Window.ShowDialog() | Out-Null

# Cleanup
$Timer.Stop()
Write-DebugLog "Application closed" -Level 'INFO'
