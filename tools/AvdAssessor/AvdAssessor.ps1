#Requires -Version 5.1
<#
.SYNOPSIS
    AVD Assessor - CAF/WAF/LZA readiness assessment for Azure Virtual Desktop.
.DESCRIPTION
    PowerShell/WPF tool for conducting workshop-style assessments of Azure
    Virtual Desktop environments. Supports manual checklist, automated
    discovery import, and rich HTML report export. Follows CAF for AVD,
    Well-Architected Framework, and Landing Zone Accelerator best practices.
.NOTES
    Author : Anton Romanyuk
    Version: 0.1.0
    Date   : 2026-03-26
#>

# ===============================================================================
# SECTION 1: PRE-LOAD & INITIALIZATION
# Console encoding, OneDrive module path cleanup, global variables, storage directory
# bootstrap, named constants, WPF/WinForms assembly loading, DPI awareness, BrushConverter cache.
# ===============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'

# Strip OneDrive user-profile module path to prevent old Az.Accounts versions
$env:PSModulePath = ($env:PSModulePath -split ';' |
    Where-Object { $_ -notlike '*OneDrive*' }) -join ';'

$Global:Root = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Global:Root)) { $Global:Root = $PWD.Path }

$Global:AppVersion       = "0.1.0-alpha"
$Global:AppTitle         = "AVD Assessor v$($Global:AppVersion)"
$Global:PrefsPath        = Join-Path $Global:Root "user_prefs.json"
$Global:AssessmentDir    = Join-Path $Global:Root "assessments"
$Global:ReportDir        = Join-Path $Global:Root "reports"
$Global:DebugLogFile     = Join-Path $env:TEMP "AvdAssessor_debug.log"

$Global:AutoSaveFile      = Join-Path $Global:Root "_autosave.json"
$Global:AutoSaveEnabled   = $false
$Global:AutoSaveInterval  = 60
$Global:OpenAfterExport   = $true
$Global:VerboseLogging    = $false
$Global:MaxBackups        = 10

# Ensure storage directories exist
foreach ($dir in @($Global:AssessmentDir, $Global:ReportDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
        Write-Host "[INIT] Created directory: $dir" -ForegroundColor DarkGray
    }
}

# Named constants
$Script:LOG_MAX_LINES       = 500
$Script:TIMER_INTERVAL_MS   = 50
$Script:TOAST_DURATION_MS   = 4000

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
public class ForegroundHelper {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
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
# SECTION 2: THEME PALETTES
# Dark and Light palettes with 40+ color keys each (backgrounds, accents, borders, text, status).
# $ApplyTheme scriptblock swaps all Window.Resources brushes for instant theme switching.
# ===============================================================================

$Global:ThemeDark = @{
    ThemeAppBg         = "#111113"; ThemePanelBg       = "#18181B"
    ThemeCardBg        = "#1E1E1E"; ThemeCardAltBg     = "#1A1A1A"
    ThemeInputBg       = "#141414"; ThemeDeepBg        = "#0D0D0D"
    ThemeOutputBg      = "#0A0A0A"; ThemeSurfaceBg     = "#1F1F23"
    ThemeHoverBg       = "#27272B"; ThemeSelectedBg    = "#2A2A2A"
    ThemePressedBg     = "#1A1A1A"
    ThemeAccent        = "#0078D4"; ThemeAccentHover    = "#1A8AD4"
    ThemeAccentLight   = "#60CDFF"; ThemeAccentDim     = "#1A0078D4"
    ThemeGreenAccent   = "#00C853"
    ThemeAccentText    = "#FFFFFF"
    ThemeTextPrimary   = "#FFFFFF"; ThemeTextBody      = "#E0E0E0"
    ThemeTextSecondary = "#B0B0B8"; ThemeTextMuted     = "#8E8E98"
    ThemeTextDim       = "#9A9AA3"; ThemeTextDisabled  = "#383838"
    ThemeTextFaintest  = "#82828C"
    ThemeBorder        = "#0FFFFFFF"; ThemeBorderCard  = "#0FFFFFFF"
    ThemeBorderElevated = "#333333"; ThemeBorderHover   = "#444444"
    ThemeBorderMedium  = "#1AFFFFFF"
    ThemeBorderSubtle  = "#333333"
    ThemeErrorDim      = "#20FF5000"
    ThemeProgressEdge  = "#18FFFFFF"
    ThemeScrollThumb   = "#999999"
    ThemeScrollTrack   = "#22FFFFFF"
    ThemeSuccess       = "#00C853"; ThemeWarning       = "#F59E0B"
    ThemeError         = "#FF5000"; ThemeSidebarBg     = "#111113"
    ThemeSidebarBorder = "#00000000"
}

$Global:ThemeLight = @{
    ThemeAppBg         = "#F5F5F5"; ThemePanelBg       = "#FAFAFA"
    ThemeCardBg        = "#FFFFFF"; ThemeCardAltBg     = "#FAFAFA"
    ThemeInputBg       = "#F0F0F0"; ThemeDeepBg        = "#EEEEEE"
    ThemeOutputBg      = "#FAFAFA"; ThemeSurfaceBg     = "#EEEEEE"
    ThemeHoverBg       = "#E8E8E8"; ThemeSelectedBg    = "#E0E0E0"
    ThemePressedBg     = "#D5D5D5"
    ThemeAccent        = "#0078D4"; ThemeAccentHover    = "#106EBE"
    ThemeAccentLight   = "#0078D4"; ThemeAccentDim     = "#1A0078D4"
    ThemeGreenAccent   = "#15803D"
    ThemeAccentText    = "#FFFFFF"
    ThemeTextPrimary   = "#111111"; ThemeTextBody      = "#222222"
    ThemeTextSecondary = "#484848"; ThemeTextMuted     = "#555555"
    ThemeTextDim       = "#5A5A5A"; ThemeTextDisabled  = "#CCCCCC"
    ThemeTextFaintest  = "#595959"
    ThemeBorder        = "#0A000000"; ThemeBorderCard  = "#0A000000"
    ThemeBorderElevated = "#0A000000"; ThemeBorderHover = "#BBBBBB"
    ThemeBorderMedium  = "#1A000000"
    ThemeBorderSubtle  = "#D0D0D0"
    ThemeErrorDim      = "#20DC2626"
    ThemeProgressEdge  = "#18000000"
    ThemeScrollThumb   = "#C0C0C0"
    ThemeScrollTrack   = "#0A000000"
    ThemeSuccess       = "#15803D"; ThemeWarning       = "#C2410C"
    ThemeError         = "#DC2626"; ThemeSidebarBg     = "#EBEBEB"
    ThemeSidebarBorder = "#00000000"
}

$Global:IsLightMode         = $false
$Global:AnimationsDisabled  = $false
$Global:SuppressThemeHandler = $false

# ===============================================================================
# SECTION 3: THREAD SYNCHRONIZATION BRIDGE
# Synchronized hashtable with StatusQueue and LogQueue for cross-thread communication.
# Start-BackgroundWork launches STA runspaces tracked in $Global:BgJobs ArrayList.
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
    Launches a PowerShell script block in a separate STA runspace for non-blocking execution.
.DESCRIPTION
    Creates and opens a new runspace with Bypass execution policy, injects specified variables
    and the SyncHash for cross-thread communication, and begins asynchronous invocation. The
    job is tracked in $Global:BgJobs and polled by the dispatcher timer for completion.
.PARAMETER Name
    Descriptive label for the background job (used in debug logging).
.PARAMETER ScriptBlock
    The work to execute in the background runspace.
.PARAMETER OnComplete
    Callback script block invoked on the UI thread when the job finishes.
.PARAMETER Arguments
    Positional arguments passed to the script block.
.PARAMETER Variables
    Hashtable of named variables injected into the runspace session state.
.PARAMETER Context
    Arbitrary hashtable stored with the job for use in OnComplete.
#>
function Start-BackgroundWork {
    param(
        [string]$Name = 'BgWork',
        [ScriptBlock]$ScriptBlock,
        [ScriptBlock]$OnComplete,
        [array]$Arguments = @(),
        [hashtable]$Variables = @{},
        [hashtable]$Context   = @{}
    )
    $Work = $ScriptBlock
    $ISS = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $ISS.ExecutionPolicy = [Microsoft.PowerShell.ExecutionPolicy]::Bypass
    $RS  = [RunspaceFactory]::CreateRunspace($ISS)
    $RS.ApartmentState = 'STA'
    $RS.ThreadOptions  = 'ReuseThread'
    $RS.Open()

    $CleanModulePath = ($env:PSModulePath -split ';' |
        Where-Object { $_ -notlike '*OneDrive*' }) -join ';'

    $PS = [PowerShell]::Create()
    $PS.Runspace = $RS

    $RS.SessionStateProxy.SetVariable('SyncHash', $Global:SyncHash)
    $RS.SessionStateProxy.SetVariable('PSModulePath_Clean', $CleanModulePath)

    foreach ($Key in $Variables.Keys) {
        $RS.SessionStateProxy.SetVariable($Key, $Variables[$Key])
    }

    [void]$PS.AddScript($Work)
    foreach ($Arg in $Arguments) { [void]$PS.AddArgument($Arg) }

    $Handle = $PS.BeginInvoke()

    [void]$Global:BgJobs.Add([PSCustomObject]@{
        Name       = $Name
        PS         = $PS
        Handle     = $Handle
        RS         = $RS
        OnComplete = $OnComplete
        Context    = $Context
        StartTime  = [DateTime]::Now
    })

    Write-DebugLog "BgWork: launched '$Name'" -Level 'DEBUG'
}

# ===============================================================================
# SECTION 4: XAML GUI LOAD
# Reads AvdAssessor_UI.xaml, strips designer-only attributes, parses with XamlReader,
# and creates the main WPF Window. Fatal errors show a MessageBox and exit.
# ===============================================================================

$XamlPath = Join-Path $Global:Root "AvdAssessor_UI.xaml"
if (-not (Test-Path $XamlPath)) {
    [System.Windows.MessageBox]::Show(
        "CRITICAL: 'AvdAssessor_UI.xaml' not found in:`n$Global:Root",
        "AVD Assessor", 'OK', 'Error') | Out-Null
    exit 1
}

$XamlContent = Get-Content $XamlPath -Raw -Encoding UTF8
Write-Host "[XAML] Loaded XAML: $([math]::Round($XamlContent.Length / 1024, 1)) KB" -ForegroundColor DarkGray
$XamlContent = $XamlContent -replace 'Title="AVD Assessor"', "Title=`"$Global:AppTitle`""

try {
    $Window = [Windows.Markup.XamlReader]::Parse($XamlContent)
    Write-Host "[XAML] Window parsed successfully" -ForegroundColor DarkGray
} catch {
    [System.Windows.MessageBox]::Show(
        "XAML Parse Error:`n$($_.Exception.Message)",
        "AVD Assessor", 'OK', 'Error') | Out-Null
    exit 1
}

# ===============================================================================
# SECTION 5: ELEMENT REFERENCES
# Binds 100+ named XAML elements (title bar, auth bar, icon rail, sidebar panels, tab panels,
# dashboard cards, assessment checklist, findings view, report preview, settings controls,
# splash overlay, debug log, status bar) to PowerShell variables via FindName().
# ===============================================================================

# Title bar
$btnThemeToggle    = $Window.FindName("btnThemeToggle")
$btnMinimize       = $Window.FindName("btnMinimize")
$btnMaximize       = $Window.FindName("btnMaximize")
$btnClose          = $Window.FindName("btnClose")
$lblTitle          = $Window.FindName("lblTitle")
$lblTitleVersion   = $Window.FindName("lblTitleVersion")
$dotDirty          = $Window.FindName("dotDirty")
$lblDirtyText      = $Window.FindName("lblDirtyText")
$btnQuickSave      = $Window.FindName("btnQuickSave")

$Global:IsDirty = $false
$Global:ActiveFilePath = $null   # Path to currently loaded/saved assessment file

# Icon rail
$railDashboard     = $Window.FindName("railDashboard")
$railAssessment    = $Window.FindName("railAssessment")
$railFindings      = $Window.FindName("railFindings")
$railReport        = $Window.FindName("railReport")
$railSettings      = $Window.FindName("railSettings")
$railDashboardIndicator   = $Window.FindName("railDashboardIndicator")
$railAssessmentIndicator  = $Window.FindName("railAssessmentIndicator")
$railFindingsIndicator    = $Window.FindName("railFindingsIndicator")
$railReportIndicator      = $Window.FindName("railReportIndicator")
$railSettingsIndicator    = $Window.FindName("railSettingsIndicator")

# Sidebar panels
$pnlSidebar                = $Window.FindName("pnlSidebar")
$pnlSidebarDashboard       = $Window.FindName("pnlSidebarDashboard")
$pnlSidebarAssessment      = $Window.FindName("pnlSidebarAssessment")
$pnlSidebarFindings        = $Window.FindName("pnlSidebarFindings")
$pnlSidebarReport          = $Window.FindName("pnlSidebarReport")
$pnlSidebarSettings        = $Window.FindName("pnlSidebarSettings")

# Sidebar - Dashboard
$txtCustomerName   = $Window.FindName("txtCustomerName")
$txtAssessorName   = $Window.FindName("txtAssessorName")
$txtAssessmentDate = $Window.FindName("txtAssessmentDate")
$lblStatHostPools    = $Window.FindName("lblStatHostPools")
$lblStatSessionHosts = $Window.FindName("lblStatSessionHosts")
$lblStatAppGroups    = $Window.FindName("lblStatAppGroups")
$lblStatWorkspaces   = $Window.FindName("lblStatWorkspaces")
$lblStatScalingPlans = $Window.FindName("lblStatScalingPlans")
$btnSaveAssessment = $Window.FindName("btnSaveAssessment")
$btnLoadAssessment = $Window.FindName("btnLoadAssessment")
$lstSavedAssessments = $Window.FindName("lstSavedAssessments")
$lblSavedCount     = $Window.FindName("lblSavedCount")
$lblSavedEmpty     = $Window.FindName("lblSavedEmpty")

# Sidebar - Assessment
$pnlCategoryList   = $Window.FindName("pnlCategoryList")
$lblProgressChecked = $Window.FindName("lblProgressChecked")
$barProgress       = $Window.FindName("barProgress")

# Sidebar - Findings filter
$cmbFilterSeverity = $Window.FindName("cmbFilterSeverity")
$cmbFilterStatus   = $Window.FindName("cmbFilterStatus")
$cmbFilterCategory = $Window.FindName("cmbFilterCategory")
$txtFindingsSearch = $Window.FindName("txtFindingsSearch")

# Sidebar - Report
$btnExportHtml     = $Window.FindName("btnExportHtml")
$btnExportCsv      = $Window.FindName("btnExportCsv")
$btnExportJson     = $Window.FindName("btnExportJson")
$btnCopySummary    = $Window.FindName("btnCopySummary")
$lblOverallScore   = $Window.FindName("lblOverallScore")
$lblOverallLabel   = $Window.FindName("lblOverallLabel")

# Settings
$chkDarkMode         = $Window.FindName("chkDarkMode")
$txtDefaultAssessor  = $Window.FindName("txtDefaultAssessor")
$txtNotes            = $Window.FindName("txtNotes")
$chkAutoSave         = $Window.FindName("chkAutoSave")
$cmbAutoSaveInterval = $Window.FindName("cmbAutoSaveInterval")
$txtExportPath       = $Window.FindName("txtExportPath")
$btnBrowseExportPath = $Window.FindName("btnBrowseExportPath")
$chkOpenAfterExport  = $Window.FindName("chkOpenAfterExport")
$chkShowActivityLog  = $Window.FindName("chkShowActivityLog")
$chkVerboseLog       = $Window.FindName("chkVerboseLog")
$cmbMaxBackups       = $Window.FindName("cmbMaxBackups")
$btnResetPrefs       = $Window.FindName("btnResetPrefs")

# Content panels
$pnlTabDashboard   = $Window.FindName("pnlTabDashboard")
$pnlTabAssessment  = $Window.FindName("pnlTabAssessment")
$pnlTabFindings    = $Window.FindName("pnlTabFindings")
$pnlTabReport      = $Window.FindName("pnlTabReport")
$pnlTabSettings    = $Window.FindName("pnlTabSettings")

# Dashboard content
$pnlScoreCards          = $Window.FindName("pnlScoreCards")
$pnlCategoryBars        = $Window.FindName("pnlCategoryBars")
$pnlMaturityAlignment   = $Window.FindName("pnlMaturityAlignment")
$pnlDiscoveredResources = $Window.FindName("pnlDiscoveredResources")

# Assessment content
$pnlAssessmentChecks = $Window.FindName("pnlAssessmentChecks")

# Findings content
$pnlFindingsList   = $Window.FindName("pnlFindingsList")
$lblFindingsCount  = $Window.FindName("lblFindingsCount")

# Report
$rtbReportSummary  = $Window.FindName("rtbReportSummary")
$docReportSummary  = $Window.FindName("docReportSummary")
$paraReportSummary = $Window.FindName("paraReportSummary")

# Activity log
$rtbActivityLog    = $Window.FindName("rtbActivityLog")
$docActivityLog    = $Window.FindName("docActivityLog")
$paraLog           = $Window.FindName("paraLog")
$logScroller       = $Window.FindName("logScroller")
$btnClearLog       = $Window.FindName("btnClearLog")
$btnHideLog        = $Window.FindName("btnHideLog")
$pnlActivityLog    = $Window.FindName("pnlActivityLog")
$railActivityLog   = $Window.FindName("railActivityLog")

# Splash
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

# ===============================================================================
# SECTION 6: LOGGING
# Write-DebugLog: centralized 3-destination logging (console, rotating disk file, UI RichTextBox)
# with level-based color coding (INFO/DEBUG/WARN/ERROR/SUCCESS) and ring buffer management.
# ===============================================================================

$Global:FullLogSB        = [System.Text.StringBuilder]::new()
$Global:FullLogLineCount = 0
$Global:FullLogMaxLines  = 1000
$Global:DebugLineCount   = 0
$Global:DebugMaxLines    = $Script:LOG_MAX_LINES
$Global:DebugOverlayEnabled = $false

<#
.SYNOPSIS
    Writes a timestamped log entry to console, disk file, ring buffer, and UI RichTextBox.
.DESCRIPTION
    Formats [HH:mm:ss.fff] [LEVEL] Message and outputs to: (1) PowerShell console,
    (2) rotating disk file at %TEMP%\AvdAssessor_debug.log (2 MB rotation), (3) in-memory
    ring buffer (1000 lines), and (4) an activity log RichTextBox with level-based color coding.
    DEBUG-level messages are suppressed in the UI unless verbose logging is enabled.
.PARAMETER Message
    The log message text.
.PARAMETER Level
    Log severity: INFO, DEBUG, WARN, ERROR, or SUCCESS. Defaults to INFO.
#>
function Write-DebugLog {
    param([string]$Message, [string]$Level = 'INFO')

    $ts   = (Get-Date).ToString('HH:mm:ss.fff')
    $line = "[$ts] [$Level] $Message"

    Write-Host $line -ForegroundColor DarkGray

    # Disk log (rotate at 2 MB)
    try {
        if ((Test-Path $Global:DebugLogFile) -and (Get-Item $Global:DebugLogFile).Length -gt 2MB) {
            $Rotated = $Global:DebugLogFile + '.old'
            if (Test-Path $Rotated) { [System.IO.File]::Delete($Rotated) }
            [System.IO.File]::Move($Global:DebugLogFile, $Rotated)
        }
        [System.IO.File]::AppendAllText($Global:DebugLogFile, $line + "`r`n")
    } catch { }

    # Ring buffer
    [void]$Global:FullLogSB.AppendLine($line)
    $Global:FullLogLineCount++
    if ($Global:FullLogLineCount -gt $Global:FullLogMaxLines) {
        $idx = $Global:FullLogSB.ToString().IndexOf("`n")
        if ($idx -ge 0) { [void]$Global:FullLogSB.Remove(0, $idx + 1) }
        $Global:FullLogLineCount--
    }

    # UI RichTextBox
    if ($Level -eq 'DEBUG' -and -not $Global:VerboseLogging) { return }
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

        $Global:DebugLineCount++
        if ($Global:DebugLineCount -gt $Global:DebugMaxLines) {
            $paraLog.Inlines.Remove($paraLog.Inlines.FirstInline)
            if ($paraLog.Inlines.Count -gt 0) {
                $paraLog.Inlines.Remove($paraLog.Inlines.FirstInline)
            }
            $Global:DebugLineCount--
        }
        $logScroller.ScrollToEnd()
    } catch { }
}

# ===============================================================================
# SECTION 7: THEME ENGINE
# $ApplyTheme scriptblock processes all palette keys, creates SolidColorBrush resources,
# generates procedural dot-grid and gradient-glow backgrounds, and updates scroll thumb colors.
# ===============================================================================

$ApplyTheme = {
    param([bool]$IsLight)
    $Palette = if ($IsLight) { $Global:ThemeLight } else { $Global:ThemeDark }

    foreach ($Key in $Palette.Keys) {
        try {
            $NewColor = [System.Windows.Media.ColorConverter]::ConvertFromString($Palette[$Key])
            # Replace the resource entry — this triggers DynamicResource re-evaluation
            # Mutating .Color does NOT trigger DynamicResource/SetResourceReference updates
            $Window.Resources[$Key] = [System.Windows.Media.SolidColorBrush]::new($NewColor)
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
}

# ===============================================================================
# SECTION 8: TOAST NOTIFICATIONS
# Auto-dismissing overlay notifications with fade-in/fade-out animations. Four visual types:
# Success (green), Error (red), Warning (amber), Info (blue) with theme-aware color palettes.
# ===============================================================================

<#
.SYNOPSIS
    Displays an auto-dismissing toast notification with theme-aware colors and fade animations.
.PARAMETER Message
    The notification text to display.
.PARAMETER Type
    Visual style: Success (green), Error (red), Warning (amber), or Info (blue).
.PARAMETER DurationMs
    Auto-dismiss delay in milliseconds. Defaults to TOAST_DURATION_MS (4000).
#>
function Show-Toast {
    param(
        [string]$Message,
        [ValidateSet('Success','Error','Warning','Info')][string]$Type = 'Info',
        [int]$DurationMs = $Script:TOAST_DURATION_MS
    )
    # Toast host — create if not present
    if (-not $Script:ToastHost) {
        $Script:ToastHost = New-Object System.Windows.Controls.StackPanel
        $Script:ToastHost.VerticalAlignment = 'Bottom'
        $Script:ToastHost.HorizontalAlignment = 'Right'
        $Script:ToastHost.Margin = [System.Windows.Thickness]::new(0,0,24,24)
        [System.Windows.Controls.Panel]::SetZIndex($Script:ToastHost, 200)
        $ContentGrid = $Window.Content
        if ($ContentGrid -is [System.Windows.Controls.Grid]) {
            $ContentGrid.Children.Add($Script:ToastHost)
        }
    }
    $pnlToastHost = $Script:ToastHost

    if ($Global:IsLightMode) {
        $Colors = @{
            Success = @{ Bg='#E8F5E9'; Border='#15803D'; TextColor='#14532D'; Icon=[char]0xE73E }
            Error   = @{ Bg='#FEE2E2'; Border='#DC2626'; TextColor='#7F1D1D'; Icon=[char]0xEA39 }
            Warning = @{ Bg='#FFF7ED'; Border='#C2410C'; TextColor='#7C2D12'; Icon=[char]0xE7BA }
            Info    = @{ Bg='#EFF6FF'; Border='#0078D4'; TextColor='#1E3A5F'; Icon=[char]0xE946 }
        }
    } else {
        $Colors = @{
            Success = @{ Bg='#0A3D1A'; Border='#00C853'; TextColor='#FFFFFF'; Icon=[char]0xE73E }
            Error   = @{ Bg='#3D0A0A'; Border='#D13438'; TextColor='#FFFFFF'; Icon=[char]0xEA39 }
            Warning = @{ Bg='#3D2D0A'; Border='#FFB900'; TextColor='#FFFFFF'; Icon=[char]0xE7BA }
            Info    = @{ Bg='#0A1E3D'; Border='#0078D4'; TextColor='#FFFFFF'; Icon=[char]0xE946 }
        }
    }
    $C = $Colors[$Type]

    $Toast = New-Object System.Windows.Controls.Border
    $Toast.Background      = $Global:CachedBC.ConvertFromString($C.Bg)
    $Toast.BorderBrush     = $Global:CachedBC.ConvertFromString($C.Border)
    $Toast.BorderThickness = [System.Windows.Thickness]::new(1)
    $Toast.CornerRadius    = [System.Windows.CornerRadius]::new(12)
    $Toast.Padding         = [System.Windows.Thickness]::new(12,8,12,8)
    $Toast.Margin          = [System.Windows.Thickness]::new(0,0,0,6)
    $Toast.Opacity         = 0

    $SP = New-Object System.Windows.Controls.StackPanel
    $SP.Orientation = 'Horizontal'
    $IconTB = New-Object System.Windows.Controls.TextBlock
    $IconTB.Text       = $C.Icon
    $IconTB.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe Fluent Icons, Segoe MDL2 Assets")
    $IconTB.FontSize   = 14
    $IconTB.Foreground = $Global:CachedBC.ConvertFromString($C.Border)
    $IconTB.VerticalAlignment = 'Center'
    $IconTB.Margin     = [System.Windows.Thickness]::new(0,0,8,0)
    $MsgTB = New-Object System.Windows.Controls.TextBlock
    $MsgTB.Text        = $Message
    $MsgTB.FontFamily  = New-Object System.Windows.Media.FontFamily("Segoe UI Emoji, Segoe UI")
    $MsgTB.FontSize    = 11
    $MsgTB.Foreground  = $Global:CachedBC.ConvertFromString($C.TextColor)
    $MsgTB.TextWrapping = 'Wrap'
    $MsgTB.MaxWidth    = 340
    $MsgTB.VerticalAlignment = 'Center'
    [void]$SP.Children.Add($IconTB)
    [void]$SP.Children.Add($MsgTB)
    $Toast.Child = $SP

    $pnlToastHost.Children.Insert(0, $Toast)

    # Fade In
    if ($Global:AnimationsDisabled) {
        $Toast.Opacity = 1
    } else {
        $FadeIn = New-Object System.Windows.Media.Animation.DoubleAnimation
        $FadeIn.From     = 0; $FadeIn.To = 1
        $FadeIn.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(200))
        $Toast.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $FadeIn)
    }

    # Auto-dismiss
    $DismissTimer = New-Object System.Windows.Threading.DispatcherTimer
    $DismissTimer.Interval = [TimeSpan]::FromMilliseconds($DurationMs)
    $ToastRef = $Toast; $HostRef = $pnlToastHost; $TimerRef = $DismissTimer
    $AnimDisabledRef = $Global:AnimationsDisabled

    $RemoveTimer = New-Object System.Windows.Threading.DispatcherTimer
    $RemoveTimer.Interval = [TimeSpan]::FromMilliseconds(250)
    $RemoveTimerRef = $RemoveTimer
    $RemoveTimer.Add_Tick({
        $RemoveTimerRef.Stop()
        if ($HostRef -and $ToastRef) { $HostRef.Children.Remove($ToastRef) }
    }.GetNewClosure())

    $DismissTimer.Add_Tick({
        try {
            $TimerRef.Stop()
            if ($AnimDisabledRef) {
                if ($HostRef -and $ToastRef) { $HostRef.Children.Remove($ToastRef) }
            } else {
                $FadeOut = New-Object System.Windows.Media.Animation.DoubleAnimation
                $FadeOut.From = 1; $FadeOut.To = 0
                $FadeOut.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(200))
                $ToastRef.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $FadeOut)
                $RemoveTimerRef.Start()
            }
        } catch { }
    }.GetNewClosure())
    $DismissTimer.Start()
    Write-DebugLog "Toast [$Type]: $Message"
}

# ===============================================================================
# SECTION 8B: THEMED DIALOG
# Chromeless modal dialog inheriting parent theme resources. Supports icon, title, message body,
# configurable button array with accent and danger styles. (AvdRewind pattern — proven working)
# ===============================================================================

<#
.SYNOPSIS
    Shows a modal themed dialog with icon, title, message, and custom buttons.
.DESCRIPTION
    Creates a chromeless WPF dialog that inherits the parent window's theme resources.
    Supports configurable icon, title, message body, and an array of button definitions.
.PARAMETER Title
    Dialog header text. Defaults to 'AVD Assessor'.
.PARAMETER Message
    Body text below the header.
.PARAMETER Icon
    Segoe Fluent Icons character for the header icon.
.PARAMETER IconColor
    Theme resource key for the icon foreground.
.PARAMETER Buttons
    Array of hashtables with Text, IsAccent, IsDanger, and Result keys.
.OUTPUTS
    String result tag of the clicked button, or 'Cancel' if dismissed.
#>
function Show-ThemedDialog {
    param(
        [string]$Title    = 'AVD Assessor',
        [string]$Message  = '',
        [string]$Icon     = [string]([char]0xE897),
        [string]$IconColor = 'ThemeAccentLight',
        [array]$Buttons   = @( @{ Text='OK'; IsAccent=$true; Result='OK' } )
    )

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

    # CRITICAL: Copy theme resources from parent window to dialog
    foreach ($rk in $Window.Resources.Keys) {
        $Dlg.Resources[$rk] = $Window.Resources[$rk]
    }

    $OuterBorder = [System.Windows.Controls.Border]::new()
    $OuterBorder.CornerRadius    = [System.Windows.CornerRadius]::new(12)
    $OuterBorder.BorderThickness = [System.Windows.Thickness]::new(1)
    $OuterBorder.Padding         = [System.Windows.Thickness]::new(24,20,24,20)
    $OuterBorder.Margin          = [System.Windows.Thickness]::new(20,20,20,20)
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

    # Header: icon badge + title
    $Header = [System.Windows.Controls.StackPanel]::new()
    $Header.Orientation = 'Horizontal'
    $Header.Margin = [System.Windows.Thickness]::new(0,0,0,16)

    $IconBadge = [System.Windows.Controls.Border]::new()
    $IconBadge.Width  = 36; $IconBadge.Height = 36
    $IconBadge.CornerRadius = [System.Windows.CornerRadius]::new(8)
    $IconBadge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeAccentDim')
    $IconBadge.Margin = [System.Windows.Thickness]::new(0,0,12,0)
    $IconTB = [System.Windows.Controls.TextBlock]::new()
    $IconTB.Text = $Icon
    $IconTB.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
    $IconTB.FontSize = 17
    $IconTB.HorizontalAlignment = 'Center'
    $IconTB.VerticalAlignment   = 'Center'
    $IconTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $IconColor)
    $IconBadge.Child = $IconTB
    [void]$Header.Children.Add($IconBadge)

    $TitleTB = [System.Windows.Controls.TextBlock]::new()
    $TitleTB.Text = $Title
    $TitleTB.FontSize   = 15
    $TitleTB.FontWeight = [System.Windows.FontWeights]::Bold
    $TitleTB.VerticalAlignment = 'Center'
    $TitleTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextPrimary')
    [void]$Header.Children.Add($TitleTB)
    [void]$MainStack.Children.Add($Header)

    # Separator
    $Sep = [System.Windows.Controls.Border]::new()
    $Sep.Height = 1
    $Sep.Margin = [System.Windows.Thickness]::new(0,0,0,16)
    $Sep.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeBorder')
    [void]$MainStack.Children.Add($Sep)

    # Message
    if ($Message) {
        $MsgBlock = [System.Windows.Controls.TextBlock]::new()
        $MsgBlock.Text = $Message
        $MsgBlock.FontSize   = 12
        $MsgBlock.TextWrapping = 'Wrap'
        $MsgBlock.Margin     = [System.Windows.Thickness]::new(0,0,0,24)
        $MsgBlock.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextSecondary')
        [void]$MainStack.Children.Add($MsgBlock)
    }

    # Button row
    $BtnRow = [System.Windows.Controls.StackPanel]::new()
    $BtnRow.Orientation = 'Horizontal'
    $BtnRow.HorizontalAlignment = 'Right'

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
        } elseif ($BtnDef.IsDanger) {
            $BtnBorder.Background = $Global:CachedBC.ConvertFromString('#D13438')
            $Btn.Foreground = [System.Windows.Media.Brushes]::White
        } else {
            $BtnBorder.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeSurfaceBg')
            $BtnBorder.BorderThickness = [System.Windows.Thickness]::new(1)
            $BtnBorder.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderSubtle')
            $Btn.SetResourceReference([System.Windows.Controls.Button]::ForegroundProperty, 'ThemeTextSecondary')
        }

        $BtnBorder.Child = $Btn
        [void]$BtnRow.Children.Add($BtnBorder)

        # Click handler — uses .GetNewClosure() with explicit param capture
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

    [void]$MainStack.Children.Add($BtnRow)

    # Drag support
    $OuterBorder.Add_MouseLeftButtonDown({
        try { $Dlg.DragMove() } catch { }
    }.GetNewClosure())

    $Dlg.ShowDialog() | Out-Null
    $DlgResult = if ($Dlg.Tag) { $Dlg.Tag } else { 'Cancel' }
    return $DlgResult
}

# ===============================================================================
# SECTION 9: TAB SWITCHING
# Five-tab switcher (Dashboard, Assessment, Findings, Report, Settings) with icon rail indicators,
# sidebar panel swaps, fade-in transitions, and a collapsible activity log sidebar.
# ===============================================================================

<#
.SYNOPSIS
    Applies a 150ms fade-in animation to a tab content panel on tab switch.
.PARAMETER Panel
    The WPF UIElement to animate (opacity 0 to 1).
#>
function Invoke-TabFade {
    param([System.Windows.UIElement]$Panel)
    $Panel.Visibility = 'Visible'
    if ($Global:AnimationsDisabled) { $Panel.Opacity = 1; return }
    $Panel.Opacity = 0
    $Fade = New-Object System.Windows.Media.Animation.DoubleAnimation
    $Fade.From = 0; $Fade.To = 1
    $Fade.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(150))
    $Panel.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $Fade)
}

$Global:ActiveTabName = 'Dashboard'

# ── Activity Log collapse/expand ──
$Global:ActivityLogOpen       = $true
$Global:ActivityLogSavedHeight = 160

<#
.SYNOPSIS
    Toggles the activity log sidebar panel between visible and collapsed states.
#>
function Toggle-ActivityLog {
    param([bool]$Show)
    $Global:ActivityLogOpen = $Show
    if ($Show) {
        $h = if ($Global:ActivityLogSavedHeight -gt 30) { $Global:ActivityLogSavedHeight } else { 160 }
        $pnlActivityLog.Height     = $h
        $pnlActivityLog.Visibility = 'Visible'
        $railActivityLog.Opacity   = 1.0
    } else {
        if ($pnlActivityLog.ActualHeight -gt 30) {
            $Global:ActivityLogSavedHeight = $pnlActivityLog.ActualHeight
        }
        $pnlActivityLog.Height     = 0
        $pnlActivityLog.Visibility = 'Collapsed'
        $railActivityLog.Opacity   = 0.4
    }
    # Keep settings checkbox in sync (guard against recursion)
    if ($chkShowActivityLog -and $chkShowActivityLog.IsChecked -ne $Show) {
        $chkShowActivityLog.IsChecked = $Show
    }
}

<#
.SYNOPSIS
    Switches the active tab in the main content area with sidebar and icon rail updates.
.DESCRIPTION
    Updates tab header highlight, content panel visibility, icon rail indicator, and sidebar
    panel for the selected tab. Applies a fade-in animation to the newly visible panel.
.PARAMETER TabName
    Target tab: Dashboard, Assessment, Findings, Report, or Settings.
#>
function Switch-Tab {
    param([string]$TabName)
    $Global:ActiveTabName = $TabName

    # Hide all content panels
    $pnlTabDashboard.Visibility   = 'Collapsed'
    $pnlTabAssessment.Visibility  = 'Collapsed'
    $pnlTabFindings.Visibility    = 'Collapsed'
    $pnlTabReport.Visibility      = 'Collapsed'
    $pnlTabSettings.Visibility    = 'Collapsed'

    # Hide all sidebar panels
    $pnlSidebarDashboard.Visibility   = 'Collapsed'
    $pnlSidebarAssessment.Visibility  = 'Collapsed'
    $pnlSidebarFindings.Visibility    = 'Collapsed'
    $pnlSidebarReport.Visibility      = 'Collapsed'
    $pnlSidebarSettings.Visibility    = 'Collapsed'

    # Reset rail indicators
    foreach ($ind in @($railDashboardIndicator, $railAssessmentIndicator, $railFindingsIndicator,
                       $railReportIndicator, $railSettingsIndicator)) {
        if ($ind) { $ind.Background = [System.Windows.Media.Brushes]::Transparent }
    }

    $AccentBrush = $Window.Resources['ThemeAccent']

    switch ($TabName) {
        'Dashboard' {
            Invoke-TabFade $pnlTabDashboard
            $pnlSidebarDashboard.Visibility = 'Visible'
            if ($railDashboardIndicator) { $railDashboardIndicator.Background = $AccentBrush }
            Update-Dashboard
        }
        'Assessment' {
            Invoke-TabFade $pnlTabAssessment
            $pnlSidebarAssessment.Visibility = 'Visible'
            if ($railAssessmentIndicator) { $railAssessmentIndicator.Background = $AccentBrush }
            Render-AssessmentChecks
        }
        'Findings' {
            Invoke-TabFade $pnlTabFindings
            $pnlSidebarFindings.Visibility = 'Visible'
            if ($railFindingsIndicator) { $railFindingsIndicator.Background = $AccentBrush }
            Render-Findings
        }
        'Report' {
            Invoke-TabFade $pnlTabReport
            $pnlSidebarReport.Visibility = 'Visible'
            if ($railReportIndicator) { $railReportIndicator.Background = $AccentBrush }
            Update-ReportPreview
        }
        'Settings' {
            Invoke-TabFade $pnlTabSettings
            $pnlSidebarSettings.Visibility = 'Visible'
            if ($railSettingsIndicator) { $railSettingsIndicator.Background = $AccentBrush }
        }
    }
    Write-DebugLog "Switched to tab: $TabName"
}

# ===============================================================================
# SECTION 10: ASSESSMENT DATA MODEL
# Loads checks.json definitions, initializes the $Global:Assessment object with CustomerName,
# AssessorName, Date, Discovery, and a Checks ArrayList of 145 check objects with Status,
# Excluded, Details, Notes, Source, Severity, Weight, and Origin fields.
# ===============================================================================

$Global:Assessment = [PSCustomObject]@{
    CustomerName  = ''
    AssessorName  = ''
    Date          = (Get-Date -Format 'yyyy-MM-dd')
    Notes         = ''
    Discovery     = $null   # Imported discovery JSON
    Checks        = [System.Collections.ArrayList]::new()
    ManualOverrides = @{}
}
$Global:CatScoreRefs = @{}  # Category -> @{ ScoreTB; CountTB; DotEllipse; SbScoreTB; SbDot }

# Load check definitions from external JSON
$Global:ChecksJsonPath = Join-Path $Global:Root 'checks.json'
if (-not (Test-Path $Global:ChecksJsonPath)) {
    [System.Windows.MessageBox]::Show(
        "CRITICAL: 'checks.json' not found in:`n$Global:Root",
        'AVD Assessor', 'OK', 'Error') | Out-Null
    exit 1
}

try {
    $ChecksFile = Get-Content $Global:ChecksJsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $Global:CheckDefinitions = $ChecksFile.checks
    Write-Host "[INIT] Loaded $($Global:CheckDefinitions.Count) check definitions from checks.json" -ForegroundColor DarkGray
} catch {
    [System.Windows.MessageBox]::Show(
        "Failed to parse checks.json:`n$($_.Exception.Message)",
        'AVD Assessor', 'OK', 'Error') | Out-Null
    exit 1
}

# Initialize checks in assessment
foreach ($Def in $Global:CheckDefinitions) {
    [void]$Global:Assessment.Checks.Add([PSCustomObject]@{
        Id             = $Def.id
        Category       = $Def.category
        Name           = $Def.name
        Description    = $Def.description
        Severity       = $Def.severity
        Type           = $Def.type
        Weight         = [int]$Def.weight    # 1-5 per-check weight
        Reference      = $Def.reference
        Origin         = $Def.origin         # AVD, WAF, CAF, LZA, FSL, SEC
        Status         = 'Not Assessed'      # Pass, Fail, Warning, N/A, Not Assessed
        Excluded       = $false              # Excluded from scoring (still visible in report)
        Details        = ''
        Notes          = ''
        Source         = $Def.type           # Auto or Manual
        Effort         = $Def.effort         # Quick Win, Some Effort, Major Effort
    })
}

# ===============================================================================
# SECTION 11: ASSESSMENT LOGIC
# Scoring engine: weighted category/overall/dimension scores (Pass=100, Warning=50, Fail=0).
# Six maturity dimensions (Security, Operations, Networking, Resiliency, Profiles, Monitoring).
# Maturity levels: Initial (0-34), Developing (35-54), Defined (55-74), Managed (75-89), Optimized (90+).
# ===============================================================================

<#
.SYNOPSIS
    Returns the list of unique assessment category names from the current check set.
.OUTPUTS
    String array of category names.
#>
function Get-Categories {
    return @($Global:Assessment.Checks | Select-Object -ExpandProperty Category -Unique)
}

<#
.SYNOPSIS
    Calculates the weighted readiness score (0-100) for a single assessment category.
.DESCRIPTION
    Scores all non-excluded checks in the category using weighted points: Pass=100,
    Warning=50, Fail=0, Not Assessed=0. N/A scores the same as Pass. Excluded checks are removed from
    scoring entirely. Returns -1 if no checks are assessable.
.PARAMETER Category
    The category name to score.
.OUTPUTS
    Integer score (0-100) or -1 if not scorable.
#>
function Get-CategoryScore {
    param([string]$Category)
    # Score ALL non-excluded checks in category. "Not Assessed" counts as 0 points
    # but IS included in the denominator so partial completion shows honest scores.
    # N/A scores the same as Pass (100 points). Excluded are removed from scoring.
    $Checks = @($Global:Assessment.Checks | Where-Object {
        $_.Category -eq $Category -and -not $_.Excluded
    })
    if ($Checks.Count -eq 0) { return -1 }
    # If nothing has been assessed at all, return -1 (not scored)
    $AnyAssessed = @($Checks | Where-Object { $_.Status -ne 'Not Assessed' }).Count
    if ($AnyAssessed -eq 0) { return -1 }
    $WeightedScore = 0; $TotalWeight = 0
    foreach ($C in $Checks) {
        $W = [math]::Max(1, [int]$C.Weight)
        $Points = switch ($C.Status) {
            'Pass'         { 100 }
            'N/A'          { 100 }  # N/A = not applicable, scores same as Pass
            'Warning'      { 50 }
            'Fail'         { 0 }
            'Not Assessed' { 0 }   # Unreviewed = 0 points, still counts in denominator
            default        { 0 }
        }
        $WeightedScore += $Points * $W
        $TotalWeight += $W
    }
    if ($TotalWeight -eq 0) { return -1 }
    return [math]::Round($WeightedScore / $TotalWeight, 0)
}

<#
.SYNOPSIS
    Calculates the weighted overall readiness score across all assessment categories.
.DESCRIPTION
    Same scoring logic as Get-CategoryScore but applied to all non-excluded checks.
    Returns -1 if no checks have been assessed.
.OUTPUTS
    Integer score (0-100) or -1 if not scorable.
#>
function Get-OverallScore {
    # Score across ALL non-excluded checks. "Not Assessed" = 0 points in denominator.
    # N/A scores the same as Pass (100 points).
    $AllChecks = @($Global:Assessment.Checks | Where-Object {
        -not $_.Excluded
    })
    if ($AllChecks.Count -eq 0) { return -1 }
    $AnyAssessed = @($AllChecks | Where-Object { $_.Status -ne 'Not Assessed' }).Count
    if ($AnyAssessed -eq 0) { return -1 }
    $WeightedSum = 0; $TotalWeight = 0
    foreach ($C in $AllChecks) {
        $W = [math]::Max(1, [int]$C.Weight)
        $Points = switch ($C.Status) {
            'Pass'         { 100 }
            'N/A'          { 100 }
            'Warning'      { 50 }
            'Fail'         { 0 }
            'Not Assessed' { 0 }
            default        { 0 }
        }
        $WeightedSum += $Points * $W
        $TotalWeight += $W
    }
    if ($TotalWeight -eq 0) { return -1 }
    return [math]::Round($WeightedSum / $TotalWeight, 0)
}

# ═══════════════════════════════════════════════════════════════════════════
# MATURITY DIMENSION SCORING
# ═══════════════════════════════════════════════════════════════════════════

$Global:MaturityDimensions = [ordered]@{
    Security   = @{ Prefixes = @('SEC-','IAM-');           Label = 'Security & Identity' }
    Operations = @{ Prefixes = @('OPS-','GOV-','SH-');     Label = 'Operations & Hosts' }
    Networking = @{ Prefixes = @('NET-');                   Label = 'Networking' }
    Resiliency = @{ Prefixes = @('BCDR-');                  Label = 'Resiliency & BCDR' }
    Profiles   = @{ Prefixes = @('PROF-');                  Label = 'Profiles & Storage' }
    Monitoring = @{ Prefixes = @('MON-');                   Label = 'Monitoring' }
}

<#
.SYNOPSIS
    Calculates the weighted score for a maturity dimension (Security, Operations, Networking, etc.).
.DESCRIPTION
    Filters checks by ID prefix matching the dimension's prefix list and computes a weighted
    percentage. Used for the six-dimension maturity radar display.
.PARAMETER DimensionKey
    Maturity dimension key: Security, Operations, Networking, Resiliency, Profiles, or Monitoring.
.OUTPUTS
    Integer score (0-100) or -1 if not scorable.
#>
function Get-DimensionScore {
    param([string]$DimensionKey)
    $Dim = $Global:MaturityDimensions[$DimensionKey]
    if (-not $Dim) { return -1 }
    $DimChecks = @($Global:Assessment.Checks | Where-Object {
        $Id = $_.Id; -not $_.Excluded -and
        ($Dim.Prefixes | Where-Object { $Id -like "$_*" })
    })
    $Assessed = @($DimChecks | Where-Object { $_.Status -ne 'Not Assessed' })
    if ($Assessed.Count -eq 0) { return -1 }
    $WSum = 0; $WMax = 0
    foreach ($C in $DimChecks) {
        $W = [math]::Max(1, [int]$C.Weight)
        $Pts = switch ($C.Status) {
            'Pass'         { 100 }
            'N/A'          { 100 }
            'Warning'      { 50 }
            'Fail'         { 0 }
            'Not Assessed' { 0 }
            default        { 0 }
        }
        $WSum += $Pts * $W; $WMax += 100 * $W
    }
    if ($WMax -eq 0) { return -1 }
    return [math]::Round($WSum / $WMax * 100, 0)
}

<#
.SYNOPSIS
    Maps a numeric score to a maturity level label (Initial/Developing/Defined/Managed/Optimized).
.PARAMETER Score
    Numeric score (0-100).
.OUTPUTS
    String maturity level name.
#>
function Get-MaturityLevel {
    param([int]$Score)
    switch ($true) {
        ($Score -ge 90) { return 'Optimized'  }
        ($Score -ge 75) { return 'Managed'    }
        ($Score -ge 55) { return 'Defined'    }
        ($Score -ge 35) { return 'Developing' }
        ($Score -ge 0)  { return 'Initial'    }
        default         { return 'Not Scored' }
    }
}

<#
.SYNOPSIS
    Calculates the average maturity score across all six dimensions.
.DESCRIPTION
    Averages the dimension scores (excluding unscorable dimensions) to produce a single
    composite maturity indicator.
.OUTPUTS
    Integer composite score (0-100) or -1 if no dimensions are scorable.
#>
function Get-CompositeMaturityScore {
    $Scores = @()
    foreach ($Key in $Global:MaturityDimensions.Keys) {
        $S = Get-DimensionScore $Key
        if ($S -ge 0) { $Scores += $S }
    }
    if ($Scores.Count -eq 0) { return -1 }
    return [math]::Round(($Scores | Measure-Object -Average).Average, 0)
}

<#
.SYNOPSIS
    Resets the current assessment to a blank state, optionally saving first.
.DESCRIPTION
    Prompts to save if work in progress, then clears all check statuses/notes/details,
    resets metadata, clears discovery data, and switches to the Dashboard tab.
#>
function Reset-Assessment {
    if (Test-AssessmentDirty) {
        $Result = Show-ThemedDialog `
            -Title 'New Assessment' `
            -Message "This will clear all check statuses, notes, and discovery data.`n`nSave current assessment first?" `
            -Icon ([char]0xE7BA) `
            -IconColor 'ThemeWarning' `
            -Buttons @(
                @{ Text = 'Save & Reset'; IsAccent = $true; Result = 'SaveReset' }
                @{ Text = 'Reset'; Result = 'Reset' }
                @{ Text = 'Cancel'; Result = 'Cancel' }
            )
        if ($Result -eq 'Cancel') { return }
        if ($Result -eq 'SaveReset') { Save-Assessment; Refresh-AssessmentList }
    }

    # Reset all checks to default state
    foreach ($Chk in $Global:Assessment.Checks) {
        $Chk.Status   = 'Not Assessed'
        $Chk.Excluded = $false
        $Chk.Details  = ''
        $Chk.Notes    = ''
        $Chk.Source   = $Chk.Type
    }

    # Reset metadata
    $Global:Assessment.CustomerName = ''
    $Global:Assessment.AssessorName = ''
    $Global:Assessment.Date         = (Get-Date -Format 'yyyy-MM-dd')
    $Global:Assessment.Discovery    = $null
    $Global:ActiveFilePath          = $null

    # Reset UI fields
    $txtCustomerName.Text = ''
    if ($Window.FindName('txtAssessorName')) { $Window.FindName('txtAssessorName').Text = '' }

    # Re-render
    Render-AssessmentChecks
    Update-Dashboard
    Switch-Tab 'Dashboard'
    Show-Toast 'New assessment started' -Type 'Info'
    Write-DebugLog 'Assessment reset — new blank assessment' -Level 'INFO'
}

<#
.SYNOPSIS
    Imports automated check results from an Invoke-AvdDiscovery JSON file into the assessment.
.DESCRIPTION
    Reads discovery JSON, merges automated check results by ID into the current assessment,
    preserves manual overrides and notes, populates resource inventory data, computes maturity
    scores, and refreshes the dashboard and assessment views. Prompts before overwriting
    existing automated results.
.PARAMETER Path
    File path to the discovery JSON file.
#>
function Import-DiscoveryJson {
    param([string]$Path)

    # Guard: discovery import overwrites automated check results
    $AutoAssessed = @($Global:Assessment.Checks | Where-Object { $_.Source -eq 'Auto' -and $_.Status -ne 'Not Assessed' }).Count
    if ($AutoAssessed -gt 0) {
        $Result = Show-ThemedDialog `
            -Title 'Re-import Discovery' `
            -Message "This will overwrite $AutoAssessed automated check result(s).`nManual overrides and notes will be preserved." `
            -Icon ([char]0xE946) `
            -IconColor 'ThemeAccentLight' `
            -Buttons @(
                @{ Text = 'Continue'; Result = 'Yes'; IsAccent = $true }
                @{ Text = 'Cancel';   Result = 'No' }
            )
        if ($Result -ne 'Yes') { return }
    }

    Write-DebugLog "Importing discovery: $Path" -Level 'INFO'
    try {
        $Json = Get-Content $Path -Raw -Encoding UTF8 | ConvertFrom-Json
        if (-not $Json.SchemaVersion) {
            Show-Toast "Invalid discovery file — missing SchemaVersion" -Type 'Error'
            return
        }

        $Global:Assessment.Discovery = $Json

        # Update inventory stats
        $lblStatHostPools.Text    = "$($Json.Inventory.HostPools.Count)"
        $lblStatSessionHosts.Text = "$($Json.Inventory.SessionHosts.Count)"
        $lblStatAppGroups.Text    = "$($Json.Inventory.AppGroups.Count)"
        $lblStatWorkspaces.Text   = "$($Json.Inventory.Workspaces.Count)"
        $lblStatScalingPlans.Text = "$($Json.Inventory.ScalingPlans.Count)"

        # Map discovery check results to assessment checks
        # Phase 1: collect all per-object results grouped by assessment check ID
        $CheckBuckets = @{}  # checkId → [list of {Status, Details, ObjectName}]
        $StatusRank = @{ 'Fail' = 0; 'Error' = 0; 'Warning' = 1; 'Pass' = 2; 'N/A' = 3; 'Not Assessed' = 4 }
        $Mapped = 0
        foreach ($DiscCheck in $Json.CheckResults) {
            # Find matching assessment check by category pattern
            $Match = $null
            switch -Wildcard ($DiscCheck.Id) {
                'SEC-TL-*'     { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SH-001' }
                'SH-002-*'     { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SH-002' }
                'SEC-RDP-*'    { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SH-011' }
                'NET-PIP-*'    { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'NET-003' }
                'NET-RDP-*'    { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'NET-004' }
                'NET-AN-*'     { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'NET-009' }
                'NET-DNS-*'    { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'NET-010' }
                'MON-DIAG-*'   { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'MON-001' }
                'SEC-RBAC-*'   { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'IAM-004' }
                'GOV-001-*'    { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'GOV-001' }
                'GOV-TAG-*'    { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'GOV-002' }
                'OPS-001-*'    { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'GOV-003' }
                'GOV-SCALE-*'  { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'BCDR-001' }
                'PROF-PE-*'    { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'PROF-002' }
                'PROF-TIER-*'  { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'PROF-003' }
                'PROF-HTTPS-*' { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'PROF-004' }
                'PROF-TLS-*'   { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'PROF-005' }
                'BCDR-STOR-*'  { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'BCDR-002' }
                'BCDR-AZ-*'    { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SH-010' }
                # New mappings
                'NET-NSG-*'    { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'NET-002' }
                'APP-CFG-*'    { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'APP-001' }
                'SEC-DISK-*'   { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SH-013' }
                'NET-SHORTPATH-*' { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'NET-011' }
                # New auto checks
                'SH-LB-*'     { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SH-017' }
                'SH-SSD-*'    { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SH-019' }
                'OPS-HB-*'    { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'OPS-002' }
                'NET-UDR-*'   { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'NET-013' }
                'PROF-FW-*'   { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'PROF-024' }
                'PROF-SD-*'   { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'PROF-025' }
                # New from full automation pass
                'IAM-SSO-*'   { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'IAM-008' }
                'IAM-JOIN-*'  { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'IAM-001' }
                'SEC-WM-*'    { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SEC-002' }
                'SEC-SCP-*'   { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SH-007' }
                'SH-IMG-*'    { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SH-004' }
                'SH-BSERIES-*' { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SH-003' }
                'NET-NATGW-*' { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'NET-012' }
                'GOV-SPACTIVE-*' { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'GOV-009' }
                'BCDR-MULTIREGION' { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'BCDR-008' }
                'MON-WSDIAG-*' { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'MON-001' }
                'MON-AGDIAG-*' { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'MON-001' }
                'MON-DEFENDER' { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'MON-011' }
                'PROF-SMBVER-*' { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'PROF-020' }
                'PROF-KERB-*' { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'PROF-022' }
                'PROF-AUTH-*' { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'PROF-023' }
                'PROF-SMBENC-*' { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'PROF-021' }
                # Phase 2 new checks
                'SH-SECBOOT-*' { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SH-025' }
                'SH-VTPM-*'    { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SH-026' }
                'SEC-OSDISK-*' { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SEC-020' }
                'SEC-KEYVAULT' { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SEC-022' }
                'NET-PRIVDNS'  { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'NET-018' }
                'NET-NETWATCHER' { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'NET-019' }
                'GOV-ORPHDISK' { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'GOV-011' }
                'GOV-ORPHNIC'  { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'GOV-012' }
                'GOV-POLICY'   { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'GOV-013' }
                'MON-ALERTS'   { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'MON-015' }
                'BCDR-SPSCHED-*' { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'BCDR-012' }
                # Phase 3 new checks
                'SH-IMGFRESH-*'  { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SH-027' }
                'GOV-QUOTA'      { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'GOV-014' }
                'GOV-CAPRESERV'  { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'GOV-015' }
                'GOV-BUDGET'     { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'GOV-016' }
                'GOV-RI'         { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'GOV-017' }
                'GOV-TAGALL'     { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'GOV-018' }
                'NET-HUBFW'      { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'NET-020' }
                'NET-HUBGW'      { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'NET-021' }
                'MON-DIAGCAT-*'  { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'MON-016' }
                'SEC-KVPE'       { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SEC-023' }
                'SEC-HPPL-*'     { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SEC-024' }
                'GOV-DISKSKU-*'  { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SH-022' }
                'MON-AMA-*'      { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'MON-004' }
                'NET-AVDOUT-*'   { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'NET-016' }
                'NET-PEER-*'     { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'NET-017' }
                'NET-SSH-*'      { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'NET-004' }
                'NET-SUBCAP-*'   { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'NET-015' }
                'OPS-AGENT-*'    { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'OPS-005' }
                'SEC-ATTEST-*'   { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SEC-019' }
                'SEC-AUDIO-*'    { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SH-011' }
                'SEC-CAM-*'      { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SH-011' }
                'SEC-CLIP-*'     { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SEC-015' }
                'SEC-COM-*'      { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SH-011' }
                'SEC-DRIVE-*'    { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SEC-014' }
                'SEC-MDE-*'      { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SH-006' }
                'SEC-PRINT-*'    { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SEC-017' }
                'SEC-USB-*'      { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SEC-016' }
                'SH-EPHEMERAL-*' { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SH-021' }
                'SH-STATUS-*'    { $Match = $Global:Assessment.Checks | Where-Object Id -eq 'SH-024' }

            }

            if ($Match) {
                $CheckId = $Match.Id
                if (-not $CheckBuckets.ContainsKey($CheckId)) {
                    $CheckBuckets[$CheckId] = [System.Collections.Generic.List[object]]::new()
                }
                # Extract object name from discovery ID (strip prefix up to first AVD- or use Evidence.HostPool)
                $ObjName = if ($DiscCheck.Evidence -and $DiscCheck.Evidence.HostPool) {
                    $DiscCheck.Evidence.HostPool
                } elseif ($DiscCheck.Id -match '-AVD-(.+)$') {
                    $Matches[1]
                } elseif ($DiscCheck.Id -match '^[A-Z]+-[A-Z]+-(.+)$') {
                    $Matches[1]
                } else {
                    $DiscCheck.Id
                }
                [void]$CheckBuckets[$CheckId].Add(@{
                    Status   = $DiscCheck.Status
                    Details  = $DiscCheck.Details
                    Object   = $ObjName
                })
                $Mapped++
            }
        }

        # Phase 2: aggregate per-object results into each assessment check
        foreach ($CheckId in $CheckBuckets.Keys) {
            $Results = $CheckBuckets[$CheckId]
            $Match = $Global:Assessment.Checks | Where-Object Id -eq $CheckId
            if (-not $Match) { continue }

            # Determine worst status across all objects
            $WorstRank = 4
            foreach ($R in $Results) {
                $Rank = if ($StatusRank.ContainsKey($R.Status)) { $StatusRank[$R.Status] } else { 4 }
                if ($Rank -lt $WorstRank) { $WorstRank = $Rank }
            }
            $WorstStatus = switch ($WorstRank) { 0 { 'Fail' } 1 { 'Warning' } 2 { 'Pass' } 3 { 'N/A' } default { 'Not Assessed' } }
            $Match.Status = $WorstStatus
            $Match.Source  = 'Auto'

            # Build aggregated details with per-object breakdown
            $Total = $Results.Count
            if ($Total -eq 1) {
                # Single object: show details directly with object name
                $R = $Results[0]
                $Match.Details = "[$($R.Object)] $($R.Details)"
            } else {
                # Multiple objects: show summary + list affected
                $FailList    = @($Results | Where-Object { $_.Status -eq 'Fail' -or $_.Status -eq 'Error' })
                $WarnList    = @($Results | Where-Object { $_.Status -eq 'Warning' })
                $PassList    = @($Results | Where-Object { $_.Status -eq 'Pass' })
                $NaList      = @($Results | Where-Object { $_.Status -eq 'N/A' })

                $Parts = [System.Collections.Generic.List[string]]::new()
                if ($PassList.Count -gt 0) { [void]$Parts.Add("$([char]0x2713) $($PassList.Count) pass") }
                if ($WarnList.Count -gt 0) { [void]$Parts.Add("$([char]0x26A0) $($WarnList.Count) warning") }
                if ($FailList.Count -gt 0) { [void]$Parts.Add("$([char]0x2717) $($FailList.Count) fail") }
                if ($NaList.Count -gt 0)   { [void]$Parts.Add("$($NaList.Count) N/A") }
                $Summary = "$Total objects: $($Parts -join ', ')"

                # List non-passing objects with their details
                $IssueItems = @($Results | Where-Object { $_.Status -eq 'Fail' -or $_.Status -eq 'Error' -or $_.Status -eq 'Warning' })
                if ($IssueItems.Count -gt 0) {
                    $ObjLines = $IssueItems | ForEach-Object {
                        $StatusIcon = if ($_.Status -eq 'Fail' -or $_.Status -eq 'Error') { [char]0x2717 } else { [char]0x26A0 }
                        "$StatusIcon $($_.Object): $($_.Details)"
                    }
                    # Limit displayed lines to avoid overwhelming the UI
                    if ($ObjLines.Count -gt 8) {
                        $Remaining = $ObjLines.Count - 6
                        $ObjLines = @($ObjLines | Select-Object -First 6) + @("  ...and $Remaining more")
                    }
                    $Match.Details = "$Summary`n$($ObjLines -join "`n")"
                } else {
                    $Match.Details = $Summary
                }
            }
        }

        Write-DebugLog "Mapped $Mapped discovery checks to assessment" -Level 'SUCCESS'
        Show-Toast "Imported discovery: $($Json.Inventory.HostPools.Count) host pools, $($Json.Inventory.SessionHosts.Count) session hosts, $Mapped checks mapped" -Type 'Success'

        Unlock-Achievement 'first_discovery'

        Update-Dashboard
        Update-Progress

    } catch {
        Write-DebugLog "Import failed: $($_.Exception.Message)" -Level 'ERROR'
        Show-Toast "Import failed: $($_.Exception.Message)" -Type 'Error'
    }
}

<#
.SYNOPSIS
    Enriches loaded assessment checks with latest fields from checks.json definitions.
.DESCRIPTION
    After loading a saved assessment, syncs impact, remediation, and effort fields from the
    current checks.json definitions without overwriting user-entered data. Also updates
    description, severity, and weight if the check definition has been updated.
#>
function Sync-CheckDefinitions {
    $DefLookup = @{}
    foreach ($Def in $Global:CheckDefinitions) { $DefLookup[$Def.id] = $Def }
    foreach ($Check in $Global:Assessment.Checks) {
        $Def = $DefLookup[$Check.Id]
        if ($Def) {
            # Enrich with new optional fields from checks.json
            foreach ($Field in @('impact','remediation','effort')) {
                $Val = $Def.$Field
                if ($Val) {
                    if (-not ($Check.PSObject.Properties.Name -contains $Field)) {
                        $Check | Add-Member -NotePropertyName $Field -NotePropertyValue $Val -Force
                    } elseif (-not $Check.$Field) {
                        $Check.$Field = $Val
                    }
                }
            }
            # Update description/severity/weight if check definition was updated
            $Check.Description = $Def.description
            $Check.Severity    = $Def.severity
            $Check.Weight      = [int]$Def.weight
        }
    }
}

# ===============================================================================
# SECTION 12: UI RENDERING
# Dashboard (score cards, category bars, maturity alignment, resource inventory), assessment
# checklist (145 checks grouped by category with status dropdowns and notes), filtered findings
# view (Fail/Warning only), and text-based report preview generation.
# ===============================================================================

<#
.SYNOPSIS
    Creates a themed score card Border element for the dashboard (title, big score, subtitle).
.PARAMETER Title
    Card header text (e.g. 'Overall Readiness' or category name).
.PARAMETER Score
    Numeric score (0-100), or -1 to show a placeholder dash.
.PARAMETER Subtitle
    Optional secondary text below the score.
.OUTPUTS
    System.Windows.Controls.Border containing the score card layout.
#>
function New-ScoreCard {
    param([string]$Title, [int]$Score, [string]$Subtitle = '')
    $Card = New-Object System.Windows.Controls.Border
    $Card.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeCardBg')
    $Card.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderCard')
    $Card.BorderThickness = [System.Windows.Thickness]::new(1)
    $Card.CornerRadius    = [System.Windows.CornerRadius]::new(10)
    $Card.Padding         = [System.Windows.Thickness]::new(16,14,16,14)
    $Card.Margin          = [System.Windows.Thickness]::new(0,0,12,12)
    $Card.MinWidth        = 180

    $SP = New-Object System.Windows.Controls.StackPanel
    $TitleTB = New-Object System.Windows.Controls.TextBlock
    $TitleTB.Text = $Title
    $TitleTB.FontSize = 12
    $TitleTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
    $TitleTB.Margin = [System.Windows.Thickness]::new(0,0,0,4)

    $ScoreTB = New-Object System.Windows.Controls.TextBlock
    $ScoreTB.FontSize = 28
    $ScoreTB.FontWeight = [System.Windows.FontWeights]::Bold
    if ($Score -lt 0) {
        $ScoreTB.Text = '—'
        $ScoreTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
    } else {
        $ScoreTB.Text = "$Score%"
        $ScoreTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $(if ($Score -ge 80) { 'ThemeSuccess' } elseif ($Score -ge 50) { 'ThemeWarning' } else { 'ThemeError' }))
    }

    [void]$SP.Children.Add($TitleTB)
    [void]$SP.Children.Add($ScoreTB)

    if ($Subtitle) {
        $SubTB = New-Object System.Windows.Controls.TextBlock
        $SubTB.Text = $Subtitle
        $SubTB.FontSize = 12
        $SubTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextFaintest')
        [void]$SP.Children.Add($SubTB)
    }

    $Card.Child = $SP
    return $Card
}

<#
.SYNOPSIS
    Creates a themed category bar card with segmented progress bar and pass/warn/fail stats.
.PARAMETER Category
    Category name displayed as the card header.
.PARAMETER Score
    Category score (0-100) or -1.
.PARAMETER Pass
    Number of checks with Pass status.
.PARAMETER Warn
    Number of checks with Warning status.
.PARAMETER Fail
    Number of checks with Fail status.
.PARAMETER Total
    Total number of checks in the category.
.OUTPUTS
    System.Windows.Controls.Border containing the category bar layout.
#>
function New-CategoryBar {
    param([string]$Category, [int]$Score, [int]$Pass, [int]$Warn, [int]$Fail, [int]$Total)

    # Card container
    $Card = New-Object System.Windows.Controls.Border
    $Card.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeCardBg')
    $Card.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderCard')
    $Card.BorderThickness = [System.Windows.Thickness]::new(1)
    $Card.CornerRadius    = [System.Windows.CornerRadius]::new(10)
    $Card.Padding         = [System.Windows.Thickness]::new(16,12,16,12)
    $Card.Margin          = [System.Windows.Thickness]::new(0,0,0,8)

    $Root = New-Object System.Windows.Controls.Grid
    $ColL = New-Object System.Windows.Controls.ColumnDefinition; $ColL.Width = [System.Windows.GridLength]::new(1, 'Star')
    $ColR = New-Object System.Windows.Controls.ColumnDefinition; $ColR.Width = [System.Windows.GridLength]::new(1, 'Auto')
    [void]$Root.ColumnDefinitions.Add($ColL)
    [void]$Root.ColumnDefinitions.Add($ColR)

    # Left side: name + progress bar + stats
    $LeftSP = New-Object System.Windows.Controls.StackPanel

    $CatName = New-Object System.Windows.Controls.TextBlock
    $CatName.Text = $Category
    $CatName.FontSize = 14
    $CatName.FontWeight = [System.Windows.FontWeights]::SemiBold
    $CatName.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
    $CatName.Margin = [System.Windows.Thickness]::new(0,0,0,8)
    [void]$LeftSP.Children.Add($CatName)

    # Stacked progress bar (pass + warn + fail segments)
    $BarGrid = New-Object System.Windows.Controls.Grid
    $BarGrid.Height = 6
    $BarGrid.Margin = [System.Windows.Thickness]::new(0,0,0,0)

    $BarBg = New-Object System.Windows.Controls.Border
    $BarBg.CornerRadius = [System.Windows.CornerRadius]::new(3)
    $BarBg.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeDeepBg')

    [void]$BarGrid.Children.Add($BarBg)

    if ($Total -gt 0 -and $Score -ge 0) {
        $TotalAssessed = $Pass + $Warn + $Fail
        if ($TotalAssessed -gt 0) {
            $BarInner = New-Object System.Windows.Controls.Grid
            $BarInner.Height = 6
            $BarInner.HorizontalAlignment = 'Stretch'
            $BarInner.ClipToBounds = $true

            # Use proportional columns for each segment
            $ColIdx = 0
            if ($Pass -gt 0) {
                $Col = New-Object System.Windows.Controls.ColumnDefinition
                $Col.Width = [System.Windows.GridLength]::new($Pass, 'Star')
                [void]$BarInner.ColumnDefinitions.Add($Col)
                $PassBar = New-Object System.Windows.Controls.Border
                $PassBar.Background = $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#15803D' } else { '#00C853' }))
                $PassBar.CornerRadius = [System.Windows.CornerRadius]::new($(if ($ColIdx -eq 0) { 3 } else { 0 }),0,0,$(if ($ColIdx -eq 0) { 3 } else { 0 }))
                [System.Windows.Controls.Grid]::SetColumn($PassBar, $ColIdx)
                [void]$BarInner.Children.Add($PassBar)
                $ColIdx++
            }
            if ($Warn -gt 0) {
                $Col = New-Object System.Windows.Controls.ColumnDefinition
                $Col.Width = [System.Windows.GridLength]::new($Warn, 'Star')
                [void]$BarInner.ColumnDefinitions.Add($Col)
                $WarnBar = New-Object System.Windows.Controls.Border
                $WarnBar.Background = $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#C2410C' } else { '#F59E0B' }))
                [System.Windows.Controls.Grid]::SetColumn($WarnBar, $ColIdx)
                [void]$BarInner.Children.Add($WarnBar)
                $ColIdx++
            }
            if ($Fail -gt 0) {
                $Col = New-Object System.Windows.Controls.ColumnDefinition
                $Col.Width = [System.Windows.GridLength]::new($Fail, 'Star')
                [void]$BarInner.ColumnDefinitions.Add($Col)
                $FailBar = New-Object System.Windows.Controls.Border
                $FailBar.Background = $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#DC2626' } else { '#FF5000' }))
                $FailBar.CornerRadius = [System.Windows.CornerRadius]::new(0,3,3,0)
                [System.Windows.Controls.Grid]::SetColumn($FailBar, $ColIdx)
                [void]$BarInner.Children.Add($FailBar)
                $ColIdx++
            }
            # Remaining unassessed portion
            $Unassessed = $Total - $TotalAssessed
            if ($Unassessed -gt 0) {
                $Col = New-Object System.Windows.Controls.ColumnDefinition
                $Col.Width = [System.Windows.GridLength]::new($Unassessed, 'Star')
                [void]$BarInner.ColumnDefinitions.Add($Col)
            }

            $BarClip = New-Object System.Windows.Controls.Border
            $BarClip.CornerRadius = [System.Windows.CornerRadius]::new(3)
            $BarClip.ClipToBounds = $true
            $BarClip.Child = $BarInner
            [void]$BarGrid.Children.Add($BarClip)
        }
    }

    [void]$LeftSP.Children.Add($BarGrid)

    # Stats row: pass/warn/fail counts
    $StatsSP = New-Object System.Windows.Controls.StackPanel
    $StatsSP.Orientation = 'Horizontal'
    $StatsSP.Margin = [System.Windows.Thickness]::new(0,6,0,0)

    $Counts = @(
        @{ Label = "$Pass pass"; Color = $(if ($Global:IsLightMode) { '#15803D' } else { '#4ADE80' }) }
        @{ Label = "$Warn warn"; Color = $(if ($Global:IsLightMode) { '#C2410C' } else { '#FBBF24' }) }
        @{ Label = "$Fail fail"; Color = $(if ($Global:IsLightMode) { '#DC2626' } else { '#FB923C' }) }
    )
    foreach ($Ct in $Counts) {
        $DotSP = New-Object System.Windows.Controls.StackPanel
        $DotSP.Orientation = 'Horizontal'
        $DotSP.Margin = [System.Windows.Thickness]::new(0,0,14,0)

        $Dot = New-Object System.Windows.Shapes.Ellipse
        $Dot.Width = 6; $Dot.Height = 6
        $Dot.Fill = $Global:CachedBC.ConvertFromString($Ct.Color)
        $Dot.VerticalAlignment = 'Center'
        $Dot.Margin = [System.Windows.Thickness]::new(0,0,4,0)

        $LblTB = New-Object System.Windows.Controls.TextBlock
        $LblTB.Text = $Ct.Label
        $LblTB.FontSize = 12
        $LblTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')

        [void]$DotSP.Children.Add($Dot)
        [void]$DotSP.Children.Add($LblTB)
        [void]$StatsSP.Children.Add($DotSP)
    }
    [void]$LeftSP.Children.Add($StatsSP)

    [System.Windows.Controls.Grid]::SetColumn($LeftSP, 0)
    [void]$Root.Children.Add($LeftSP)

    # Right side: big score
    $ScoreSP = New-Object System.Windows.Controls.StackPanel
    $ScoreSP.VerticalAlignment = 'Center'
    $ScoreSP.HorizontalAlignment = 'Right'
    $ScoreSP.Margin = [System.Windows.Thickness]::new(12,0,0,0)

    $ScoreTB = New-Object System.Windows.Controls.TextBlock
    $ScoreTB.FontSize = 24
    $ScoreTB.FontWeight = [System.Windows.FontWeights]::Bold
    $ScoreTB.HorizontalAlignment = 'Right'
    if ($Score -lt 0) {
        $ScoreTB.Text = '—'
        $ScoreTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
    } else {
        $ScoreTB.Text = "$Score%"
        $ScoreTB.Foreground = if ($Score -ge 80) {
            $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#15803D' } else { '#4ADE80' }))
        } elseif ($Score -ge 50) {
            $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#C2410C' } else { '#FBBF24' }))
        } else {
            $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#DC2626' } else { '#FB923C' }))
        }
    }
    [void]$ScoreSP.Children.Add($ScoreTB)

    if ($Total -gt 0) {
        $OfTB = New-Object System.Windows.Controls.TextBlock
        $OfTB.Text = "$Total checks"
        $OfTB.FontSize = 11
        $OfTB.HorizontalAlignment = 'Right'
        $OfTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextFaintest')
        [void]$ScoreSP.Children.Add($OfTB)
    }

    [System.Windows.Controls.Grid]::SetColumn($ScoreSP, 1)
    [void]$Root.Children.Add($ScoreSP)

    $Card.Child = $Root
    return $Card
}

<#
.SYNOPSIS
    Rebuilds the entire Dashboard tab: score cards, category bars, maturity alignment, and resource inventory.
.DESCRIPTION
    Calculates overall and per-category scores, renders score cards and progress bars,
    computes the six-dimension maturity composite with zone strip visualization, and
    displays discovered resource inventory if available from a discovery import.
#>
function Update-Dashboard {
    $pnlScoreCards.Children.Clear()
    $pnlCategoryBars.Children.Clear()
    $pnlMaturityAlignment.Children.Clear()
    $pnlDiscoveredResources.Children.Clear()

    $Overall = Get-OverallScore
    [void]$pnlScoreCards.Children.Add((New-ScoreCard -Title 'Overall Readiness' -Score $Overall -Subtitle 'Weighted average across all categories'))
    $lblOverallScore.Text = if ($Overall -ge 0) { "$Overall%" } else { '—%' }

    $Categories = Get-Categories
    foreach ($Cat in $Categories) {
        $Score = Get-CategoryScore $Cat
        $CatChecks = @($Global:Assessment.Checks | Where-Object Category -eq $Cat)
        $Assessed  = @($CatChecks | Where-Object { $_.Status -ne 'Not Assessed' -and -not $_.Excluded })
        $Pass = @($Assessed | Where-Object { $_.Status -eq 'Pass' -or $_.Status -eq 'N/A' }).Count
        $Warn = @($Assessed | Where-Object Status -eq 'Warning').Count
        $Fail = @($Assessed | Where-Object Status -eq 'Fail').Count
        [void]$pnlScoreCards.Children.Add((New-ScoreCard -Title $Cat -Score $Score -Subtitle "$($CatChecks.Count) checks"))
        [void]$pnlCategoryBars.Children.Add((New-CategoryBar -Category $Cat -Score $Score -Pass $Pass -Warn $Warn -Fail $Fail -Total $CatChecks.Count))
    }

    # Maturity alignment
    $pnlMaturityAlignment.Children.Clear()
    $CompositeScore = Get-CompositeMaturityScore
    $CompositeLevel = if ($CompositeScore -ge 0) { Get-MaturityLevel $CompositeScore } else { 'Not Scored' }

    # Zone definitions: threshold, label, color
    $Zones = @(
        @{ Min = 0;  Max = 34; Label = 'Initial';    Color = $(if ($Global:IsLightMode) { '#DC2626' } else { '#FF5000' }) }
        @{ Min = 35; Max = 54; Label = 'Developing'; Color = $(if ($Global:IsLightMode) { '#C2410C' } else { '#F59E0B' }) }
        @{ Min = 55; Max = 74; Label = 'Defined';    Color = $(if ($Global:IsLightMode) { '#CA8A04' } else { '#FBBF24' }) }
        @{ Min = 75; Max = 89; Label = 'Managed';    Color = $(if ($Global:IsLightMode) { '#0D9488' } else { '#2DD4BF' }) }
        @{ Min = 90; Max = 100; Label = 'Optimized'; Color = $(if ($Global:IsLightMode) { '#15803D' } else { '#00C853' }) }
    )

    # Composite header card
    $MatCard = New-Object System.Windows.Controls.Border
    $MatCard.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeCardBg')
    $MatCard.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderCard')
    $MatCard.BorderThickness = [System.Windows.Thickness]::new(1)
    $MatCard.CornerRadius    = [System.Windows.CornerRadius]::new(10)
    $MatCard.Padding         = [System.Windows.Thickness]::new(16,14,16,14)

    $MatRoot = New-Object System.Windows.Controls.StackPanel

    # Composite score header
    $MatHeader = New-Object System.Windows.Controls.DockPanel
    $MatHeader.Margin = [System.Windows.Thickness]::new(0,0,0,14)
    $MatHeaderLeft = New-Object System.Windows.Controls.StackPanel
    $MatHeaderLeft.Orientation = 'Horizontal'
    $MatTitleTB = New-Object System.Windows.Controls.TextBlock
    $MatTitleTB.Text = 'Composite Maturity'
    $MatTitleTB.FontSize = 14; $MatTitleTB.FontWeight = [System.Windows.FontWeights]::SemiBold
    $MatTitleTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
    $MatTitleTB.VerticalAlignment = 'Center'
    [void]$MatHeaderLeft.Children.Add($MatTitleTB)

    if ($CompositeScore -ge 0) {
        $MatLevelBadge = New-Object System.Windows.Controls.Border
        $LevelZone = $Zones | Where-Object { $CompositeScore -ge $_.Min -and $CompositeScore -le $_.Max } | Select-Object -First 1
        $MatLevelBadge.Background = $Global:CachedBC.ConvertFromString(($LevelZone.Color + '33'))
        $MatLevelBadge.CornerRadius = [System.Windows.CornerRadius]::new(4)
        $MatLevelBadge.Padding = [System.Windows.Thickness]::new(8,2,8,2)
        $MatLevelBadge.Margin = [System.Windows.Thickness]::new(10,0,0,0)
        $MatLevelBadge.VerticalAlignment = 'Center'
        $MatLevelTB = New-Object System.Windows.Controls.TextBlock
        $MatLevelTB.Text = $CompositeLevel
        $MatLevelTB.FontSize = 12; $MatLevelTB.FontWeight = [System.Windows.FontWeights]::SemiBold
        $MatLevelTB.Foreground = $Global:CachedBC.ConvertFromString($LevelZone.Color)
        $MatLevelBadge.Child = $MatLevelTB
        [void]$MatHeaderLeft.Children.Add($MatLevelBadge)
    }
    [void]$MatHeader.Children.Add($MatHeaderLeft)

    $MatScoreTB = New-Object System.Windows.Controls.TextBlock
    $MatScoreTB.Text = if ($CompositeScore -ge 0) { "$CompositeScore%" } else { '—' }
    $MatScoreTB.FontSize = 20; $MatScoreTB.FontWeight = [System.Windows.FontWeights]::Bold
    $MatScoreTB.HorizontalAlignment = 'Right'; $MatScoreTB.VerticalAlignment = 'Center'
    if ($CompositeScore -ge 0 -and $LevelZone) {
        $MatScoreTB.Foreground = $Global:CachedBC.ConvertFromString($LevelZone.Color)
    } else {
        $MatScoreTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
    }
    [System.Windows.Controls.DockPanel]::SetDock($MatScoreTB, 'Right')
    [void]$MatHeader.Children.Add($MatScoreTB)
    [void]$MatRoot.Children.Add($MatHeader)

    # Zone legend strip (horizontal segments)
    $ZoneStrip = New-Object System.Windows.Controls.Grid
    $ZoneStrip.Height = 8
    $ZoneStrip.Margin = [System.Windows.Thickness]::new(0,0,0,4)
    $ZoneClip = New-Object System.Windows.Controls.Border
    $ZoneClip.CornerRadius = [System.Windows.CornerRadius]::new(4)
    $ZoneClip.ClipToBounds = $true
    $ZoneInner = New-Object System.Windows.Controls.Grid
    for ($zi = 0; $zi -lt $Zones.Count; $zi++) {
        $ZCol = New-Object System.Windows.Controls.ColumnDefinition
        $ZCol.Width = [System.Windows.GridLength]::new(($Zones[$zi].Max - $Zones[$zi].Min + 1), 'Star')
        [void]$ZoneInner.ColumnDefinitions.Add($ZCol)
        $ZSeg = New-Object System.Windows.Controls.Border
        $ZSeg.Background = $Global:CachedBC.ConvertFromString(($Zones[$zi].Color + '55'))
        $ZSeg.Margin = [System.Windows.Thickness]::new($(if ($zi -gt 0) { 1 } else { 0 }),0,0,0)
        [System.Windows.Controls.Grid]::SetColumn($ZSeg, $zi)
        [void]$ZoneInner.Children.Add($ZSeg)
    }
    $ZoneClip.Child = $ZoneInner
    [void]$ZoneStrip.Children.Add($ZoneClip)

    # Composite score marker on strip
    if ($CompositeScore -ge 0) {
        $Marker = New-Object System.Windows.Shapes.Ellipse
        $Marker.Width = 12; $Marker.Height = 12
        $Marker.Fill = $Global:CachedBC.ConvertFromString($LevelZone.Color)
        $Marker.Stroke = $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#FFFFFF' } else { '#1E1E2E' }))
        $Marker.StrokeThickness = 2
        $Marker.HorizontalAlignment = 'Left'
        $Marker.VerticalAlignment = 'Center'
        $Marker.Margin = [System.Windows.Thickness]::new(-6,-2,0,0)
        $Marker.SetValue([System.Windows.Controls.Canvas]::ZIndexProperty, 10)
        # Position using a canvas overlay
        $MarkerCanvas = New-Object System.Windows.Controls.Canvas
        $MarkerCanvas.Height = 12
        $MarkerCanvas.HorizontalAlignment = 'Stretch'
        # We position the marker after render using SizeChanged
        $MarkerRef = $Marker
        $ScoreRef = $CompositeScore
        $MarkerCanvas.Add_SizeChanged({
            param($sender, $e)
            $ActualW = $sender.ActualWidth
            if ($ActualW -gt 0) {
                $Pct = [math]::Min(100, [math]::Max(0, $ScoreRef)) / 100
                [System.Windows.Controls.Canvas]::SetLeft($MarkerRef, ($ActualW * $Pct) - 6)
            }
        }.GetNewClosure())
        [void]$MarkerCanvas.Children.Add($Marker)
        [void]$ZoneStrip.Children.Add($MarkerCanvas)
    }
    [void]$MatRoot.Children.Add($ZoneStrip)

    # Zone labels
    $ZoneLabelGrid = New-Object System.Windows.Controls.Grid
    $ZoneLabelGrid.Margin = [System.Windows.Thickness]::new(0,0,0,16)
    for ($zi = 0; $zi -lt $Zones.Count; $zi++) {
        $ZCol = New-Object System.Windows.Controls.ColumnDefinition
        $ZCol.Width = [System.Windows.GridLength]::new(($Zones[$zi].Max - $Zones[$zi].Min + 1), 'Star')
        [void]$ZoneLabelGrid.ColumnDefinitions.Add($ZCol)
        $ZLbl = New-Object System.Windows.Controls.TextBlock
        $ZLbl.Text = $Zones[$zi].Label
        $ZLbl.FontSize = 9
        $ZLbl.HorizontalAlignment = 'Center'
        $ZLbl.Foreground = $Global:CachedBC.ConvertFromString(($Zones[$zi].Color + 'BB'))
        [System.Windows.Controls.Grid]::SetColumn($ZLbl, $zi)
        [void]$ZoneLabelGrid.Children.Add($ZLbl)
    }
    [void]$MatRoot.Children.Add($ZoneLabelGrid)

    # Per-dimension rows
    foreach ($Key in $Global:MaturityDimensions.Keys) {
        $DimLabel = $Global:MaturityDimensions[$Key].Label
        $DimScore = Get-DimensionScore $Key
        $DimLevel = if ($DimScore -ge 0) { Get-MaturityLevel $DimScore } else { 'Not Scored' }

        $DimRow = New-Object System.Windows.Controls.Grid
        $DimRow.Margin = [System.Windows.Thickness]::new(0,0,0,8)
        $DRCol1 = New-Object System.Windows.Controls.ColumnDefinition; $DRCol1.Width = [System.Windows.GridLength]::new(130, 'Pixel')
        $DRCol2 = New-Object System.Windows.Controls.ColumnDefinition; $DRCol2.Width = [System.Windows.GridLength]::new(1, 'Star')
        $DRCol3 = New-Object System.Windows.Controls.ColumnDefinition; $DRCol3.Width = [System.Windows.GridLength]::new(48, 'Pixel')
        [void]$DimRow.ColumnDefinitions.Add($DRCol1)
        [void]$DimRow.ColumnDefinitions.Add($DRCol2)
        [void]$DimRow.ColumnDefinitions.Add($DRCol3)

        # Dimension label
        $DLblTB = New-Object System.Windows.Controls.TextBlock
        $DLblTB.Text = $DimLabel
        $DLblTB.FontSize = 12
        $DLblTB.VerticalAlignment = 'Center'
        $DLblTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextSecondary')
        [System.Windows.Controls.Grid]::SetColumn($DLblTB, 0)
        [void]$DimRow.Children.Add($DLblTB)

        # Gauge bar with zone coloring
        $DGauge = New-Object System.Windows.Controls.Grid
        $DGauge.Height = 8
        $DGauge.VerticalAlignment = 'Center'
        $DGaugeBg = New-Object System.Windows.Controls.Border
        $DGaugeBg.CornerRadius = [System.Windows.CornerRadius]::new(4)
        $DGaugeBg.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeDeepBg')
        [void]$DGauge.Children.Add($DGaugeBg)

        if ($DimScore -ge 0) {
            $DFillClip = New-Object System.Windows.Controls.Border
            $DFillClip.CornerRadius = [System.Windows.CornerRadius]::new(4)
            $DFillClip.ClipToBounds = $true
            $DFillClip.HorizontalAlignment = 'Stretch'

            $DFillInner = New-Object System.Windows.Controls.Grid
            $DFillPct = [math]::Min(100, [math]::Max(0, $DimScore))
            if ($DFillPct -gt 0) {
                $FCol = New-Object System.Windows.Controls.ColumnDefinition
                $FCol.Width = [System.Windows.GridLength]::new($DFillPct, 'Star')
                [void]$DFillInner.ColumnDefinitions.Add($FCol)
                $DimZone = $Zones | Where-Object { $DimScore -ge $_.Min -and $DimScore -le $_.Max } | Select-Object -First 1
                $DFill = New-Object System.Windows.Controls.Border
                $DFill.Background = $Global:CachedBC.ConvertFromString($DimZone.Color)
                $DFill.CornerRadius = [System.Windows.CornerRadius]::new(4)
                [System.Windows.Controls.Grid]::SetColumn($DFill, 0)
                [void]$DFillInner.Children.Add($DFill)
            }
            if ($DFillPct -lt 100) {
                $ECol = New-Object System.Windows.Controls.ColumnDefinition
                $ECol.Width = [System.Windows.GridLength]::new((100 - $DFillPct), 'Star')
                [void]$DFillInner.ColumnDefinitions.Add($ECol)
            }

            $DFillClip.Child = $DFillInner
            [void]$DGauge.Children.Add($DFillClip)
        }

        [System.Windows.Controls.Grid]::SetColumn($DGauge, 1)
        [void]$DimRow.Children.Add($DGauge)

        # Score text
        $DScoreTB = New-Object System.Windows.Controls.TextBlock
        $DScoreTB.FontSize = 12; $DScoreTB.FontWeight = [System.Windows.FontWeights]::SemiBold
        $DScoreTB.HorizontalAlignment = 'Right'; $DScoreTB.VerticalAlignment = 'Center'
        if ($DimScore -ge 0) {
            $DimZone = $Zones | Where-Object { $DimScore -ge $_.Min -and $DimScore -le $_.Max } | Select-Object -First 1
            $DScoreTB.Text = "$DimScore%"
            $DScoreTB.Foreground = $Global:CachedBC.ConvertFromString($DimZone.Color)
        } else {
            $DScoreTB.Text = '—'
            $DScoreTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
        }
        [System.Windows.Controls.Grid]::SetColumn($DScoreTB, 2)
        [void]$DimRow.Children.Add($DScoreTB)

        [void]$MatRoot.Children.Add($DimRow)
    }

    $MatCard.Child = $MatRoot
    [void]$pnlMaturityAlignment.Children.Add($MatCard)

    # Discovered resources
    if ($Global:Assessment.Discovery) {
        $D = $Global:Assessment.Discovery
        $Items = @(
            @{ Label = 'Host Pools';       Count = $D.Inventory.HostPools.Count;       Icon = [char]0xE770 }
            @{ Label = 'Session Hosts';    Count = $D.Inventory.SessionHosts.Count;    Icon = [char]0xE7F8 }
            @{ Label = 'App Groups';       Count = $D.Inventory.AppGroups.Count;       Icon = [char]0xE74C }
            @{ Label = 'Workspaces';       Count = $D.Inventory.Workspaces.Count;      Icon = [char]0xE8F1 }
            @{ Label = 'Scaling Plans';    Count = $D.Inventory.ScalingPlans.Count;    Icon = [char]0xE9D9 }
            @{ Label = 'Virtual Networks'; Count = $D.Inventory.VNets.Count;           Icon = [char]0xE968 }
            @{ Label = 'Storage Accounts'; Count = $D.Inventory.StorageAccounts.Count; Icon = [char]0xEDA2 }
        )

        $ResWrap = New-Object System.Windows.Controls.WrapPanel
        $ResWrap.Orientation = 'Horizontal'

        foreach ($Item in $Items) {
            $Tile = New-Object System.Windows.Controls.Border
            $Tile.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeCardBg')
            $Tile.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderCard')
            $Tile.BorderThickness = [System.Windows.Thickness]::new(1)
            $Tile.CornerRadius    = [System.Windows.CornerRadius]::new(10)
            $Tile.Padding         = [System.Windows.Thickness]::new(14,12,14,12)
            $Tile.Margin          = [System.Windows.Thickness]::new(0,0,10,10)
            $Tile.MinWidth        = 150

            $TileSP = New-Object System.Windows.Controls.StackPanel

            # Icon + count row
            $TopRow = New-Object System.Windows.Controls.DockPanel
            $IconTB = New-Object System.Windows.Controls.TextBlock
            $IconTB.Text = $Item.Icon
            $IconTB.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
            $IconTB.FontSize = 18
            $IconTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeAccent')
            $IconTB.VerticalAlignment = 'Center'

            $CountTB = New-Object System.Windows.Controls.TextBlock
            $CountTB.Text = "$($Item.Count)"
            $CountTB.FontSize = 22
            $CountTB.FontWeight = [System.Windows.FontWeights]::Bold
            $CountTB.HorizontalAlignment = 'Right'
            $CountTB.VerticalAlignment = 'Center'
            if ($Item.Count -gt 0) {
                $CountTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextPrimary')
            } else {
                $CountTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextFaintest')
            }
            [System.Windows.Controls.DockPanel]::SetDock($CountTB, 'Right')
            [void]$TopRow.Children.Add($CountTB)
            [void]$TopRow.Children.Add($IconTB)
            [void]$TileSP.Children.Add($TopRow)

            # Label
            $LblTB = New-Object System.Windows.Controls.TextBlock
            $LblTB.Text = $Item.Label
            $LblTB.FontSize = 12
            $LblTB.Margin = [System.Windows.Thickness]::new(0,6,0,0)
            $LblTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
            [void]$TileSP.Children.Add($LblTB)

            $Tile.Child = $TileSP
            [void]$ResWrap.Children.Add($Tile)
        }
        [void]$pnlDiscoveredResources.Children.Add($ResWrap)
    } else {
        $Msg = New-Object System.Windows.Controls.TextBlock
        $Msg.Text = "No discovery data loaded. Use 'Load Assessment' to import a discovery JSON, or run Invoke-AvdDiscovery.ps1."
        $Msg.FontSize = 11; $Msg.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted'); $Msg.TextWrapping = 'Wrap'
        [void]$pnlDiscoveredResources.Children.Add($Msg)
    }
}

<#
.SYNOPSIS
    Updates the header progress bar and completion statistics based on current check statuses.
.DESCRIPTION
    Counts assessed, passed, warned, failed, excluded, and N/A checks. Updates the progress
    bar value, overall score label, and status text. Triggers the dirty flag for auto-save.
#>
function Update-Progress {
    # Single-pass stats collection
    $Total = @($Global:Assessment.Checks).Count
    $Excluded = 0; $NA = 0; $Checked = 0
    $CatStats = @{}
    foreach ($Chk in $Global:Assessment.Checks) {
        if ($Chk.Excluded) { $Excluded++; continue }
        if ($Chk.Status -ne 'Not Assessed') { $Checked++ }
        if ($Chk.Status -eq 'N/A') { $NA++ }
        # Build per-category stats
        $Cat = $Chk.Category
        if (-not $CatStats.ContainsKey($Cat)) { $CatStats[$Cat] = @{ Assessed = 0; Pass = 0; Warn = 0; Fail = 0; Total = 0 } }
        $CatStats[$Cat].Total++
        if ($Chk.Status -ne 'Not Assessed') {
            $CatStats[$Cat].Assessed++
            switch ($Chk.Status) {
                'Pass'    { $CatStats[$Cat].Pass++ }
                'N/A'     { $CatStats[$Cat].Pass++ }  # N/A counts as Pass
                'Warning' { $CatStats[$Cat].Warn++ }
                'Fail'    { $CatStats[$Cat].Fail++ }
            }
        }
    }
    $Scorable = $Total - $Excluded
    $lblProgressChecked.Text = "$Checked / $Scorable scored$(if ($Excluded -gt 0) { " ($Excluded excluded)" })$(if ($NA -gt 0) { " · $NA N/A" })"
    $barProgress.Value  = if ($Scorable -gt 0) { [math]::Round(($Checked / $Scorable) * 100, 0) } else { 0 }

    # Refresh category scores in cards + sidebar (using cached CatStats)
    foreach ($Cat in $Global:CatScoreRefs.Keys) {
        $Refs = $Global:CatScoreRefs[$Cat]
        $CatScore = Get-CategoryScore $Cat
        $CS = if ($CatStats.ContainsKey($Cat)) { $CatStats[$Cat] } else { @{ Assessed = 0; Pass = 0; Warn = 0; Fail = 0; Total = 0 } }
        $CatPass = $CS.Pass
        $CatTotalCount = $CS.Total

        $DotColor = if ($CatScore -ge 80) {
            if ($Global:IsLightMode) { '#15803D' } else { '#00C853' }
        } elseif ($CatScore -ge 50) {
            if ($Global:IsLightMode) { '#C2410C' } else { '#F59E0B' }
        } elseif ($CatScore -ge 0) {
            if ($Global:IsLightMode) { '#DC2626' } else { '#FF5000' }
        } else { $null }

        # Update main card score
        if ($Refs.ScoreTB) {
            if ($CatScore -ge 0) {
                $Refs.ScoreTB.Text = "$CatScore%"
                $Refs.ScoreTB.Foreground = if ($CatScore -ge 80) {
                    $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#15803D' } else { '#4ADE80' }))
                } elseif ($CatScore -ge 50) {
                    $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#C2410C' } else { '#FBBF24' }))
                } else {
                    $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#DC2626' } else { '#FB923C' }))
                }
            } else {
                $Refs.ScoreTB.Text = [string][char]0x2014
                $Refs.ScoreTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
            }
        }
        if ($Refs.CountTB) { $Refs.CountTB.Text = "$CatPass/$CatTotalCount" }
        if ($Refs.DotEllipse -and $DotColor) { $Refs.DotEllipse.Fill = $Global:CachedBC.ConvertFromString($DotColor) }
        elseif ($Refs.DotEllipse) { $Refs.DotEllipse.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, 'ThemeTextFaintest') }

        # Update sidebar score
        if ($Refs.SbScoreTB) {
            if ($CatScore -ge 0) {
                $Refs.SbScoreTB.Text = "$CatScore%"
                if ($DotColor) { $Refs.SbScoreTB.Foreground = $Global:CachedBC.ConvertFromString($DotColor) } else { $Refs.SbScoreTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted') }
            } else {
                $Refs.SbScoreTB.Text = [string][char]0x2014
                $Refs.SbScoreTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
            }
        }
        if ($Refs.SbDot -and $DotColor) { $Refs.SbDot.Fill = $Global:CachedBC.ConvertFromString($DotColor) }
        elseif ($Refs.SbDot) { $Refs.SbDot.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, 'ThemeTextFaintest') }
    }
}

<#
.SYNOPSIS
    Renders the full 145-check assessment checklist in the Assessment tab.
.DESCRIPTION
    Groups checks by category, renders each with a status dropdown (Pass/Fail/Warning/N/A/
    Not Assessed), severity badge, weight indicator, origin tag, reference link, details/notes
    input fields, and an exclude checkbox. Wires change handlers that update scoring in real-time.
#>
function Render-AssessmentChecks {
    # Clear existing dynamic content (keep header TextBlocks)
    $KeepCount = 2  # Title + subtitle
    while ($pnlAssessmentChecks.Children.Count -gt $KeepCount) {
        $pnlAssessmentChecks.Children.RemoveAt($pnlAssessmentChecks.Children.Count - 1)
    }

    # Category sidebar
    $pnlCategoryList.Children.Clear()
    $Global:CatScoreRefs = @{}
    $Categories = Get-Categories

    foreach ($Cat in $Categories) {
        $CatChecks = @($Global:Assessment.Checks | Where-Object Category -eq $Cat)
        $CatScore = Get-CategoryScore $Cat
        $CatAssessed = @($CatChecks | Where-Object { $_.Status -ne 'Not Assessed' -and -not $_.Excluded })
        $CatPass = @($CatAssessed | Where-Object { $_.Status -eq 'Pass' -or $_.Status -eq 'N/A' }).Count

        # Checks container (collapsed by default, toggled via category card click)
        $ChecksContainer = New-Object System.Windows.Controls.StackPanel
        $ChecksContainer.Tag = "checks_$Cat"
        $ChecksContainer.Visibility = 'Collapsed'

        # Category dot color
        $SidebarDotColor = if ($CatScore -ge 80) {
            if ($Global:IsLightMode) { '#15803D' } else { '#00C853' }
        } elseif ($CatScore -ge 50) {
            if ($Global:IsLightMode) { '#C2410C' } else { '#F59E0B' }
        } elseif ($CatScore -ge 0) {
            if ($Global:IsLightMode) { '#DC2626' } else { '#FF5000' }
        } else { $null }

        # Category sidebar button with dot + name + score
        $CatBtn = New-Object System.Windows.Controls.Button
        $CatBtn.Style = $Window.FindResource('GhostButton')
        $CatBtn.HorizontalAlignment = 'Stretch'
        $CatBtn.HorizontalContentAlignment = 'Stretch'
        $CatBtn.FontSize = 12
        $CatBtn.Margin = [System.Windows.Thickness]::new(0,0,0,0)
        $CatBtn.Padding = [System.Windows.Thickness]::new(4,8,4,8)
        $SbDP = New-Object System.Windows.Controls.DockPanel
        $SbDot = New-Object System.Windows.Shapes.Ellipse
        $SbDot.Width = 8; $SbDot.Height = 8; $SbDot.Margin = [System.Windows.Thickness]::new(0,0,10,0); $SbDot.VerticalAlignment = 'Center'
        if ($SidebarDotColor) { $SbDot.Fill = $Global:CachedBC.ConvertFromString($SidebarDotColor) } else { $SbDot.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, 'ThemeTextFaintest') }
        $SbScoreTB = New-Object System.Windows.Controls.TextBlock
        $SbScoreTB.VerticalAlignment = 'Center'; $SbScoreTB.FontSize = 12
        [System.Windows.Controls.DockPanel]::SetDock($SbScoreTB, 'Right')
        if ($CatScore -ge 0) {
            $SbScoreTB.Text = "$CatScore%"
            if ($SidebarDotColor) { $SbScoreTB.Foreground = $Global:CachedBC.ConvertFromString($SidebarDotColor) } else { $SbScoreTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted') }
        } else {
            $SbScoreTB.Text = [char]0x2014
            $SbScoreTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
        }
        $SbNameTB = New-Object System.Windows.Controls.TextBlock
        $SbNameTB.Text = $Cat; $SbNameTB.VerticalAlignment = 'Center'; $SbNameTB.TextTrimming = 'CharacterEllipsis'
        $SbNameTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextSecondary')
        [void]$SbDP.Children.Add($SbDot)
        [void]$SbDP.Children.Add($SbScoreTB)
        [void]$SbDP.Children.Add($SbNameTB)
        $CatBtn.Content = $SbDP
        [void]$pnlCategoryList.Children.Add($CatBtn)

        # Horizontal separator between categories
        $SbSep = New-Object System.Windows.Controls.Border
        $SbSep.Height = 1
        $SbSep.Margin = [System.Windows.Thickness]::new(18,2,0,2)
        $SbSep.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeBorder')
        [void]$pnlCategoryList.Children.Add($SbSep)

        # Category header — styled card with chevron + dot + score
        $CatCard = New-Object System.Windows.Controls.Border
        $CatCard.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeSurfaceBg')
        $CatCard.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderCard')
        $CatCard.BorderThickness = [System.Windows.Thickness]::new(1)
        $CatCard.CornerRadius    = [System.Windows.CornerRadius]::new(10)
        $CatCard.Padding         = [System.Windows.Thickness]::new(16,10,16,10)
        $CatCard.Margin          = [System.Windows.Thickness]::new(0,20,0,6)

        $CatDP = New-Object System.Windows.Controls.DockPanel

        # Right: score
        $CatScoreTB = New-Object System.Windows.Controls.TextBlock
        $CatScoreTB.FontSize = 16; $CatScoreTB.FontWeight = [System.Windows.FontWeights]::Bold
        $CatScoreTB.VerticalAlignment = 'Center'
        [System.Windows.Controls.DockPanel]::SetDock($CatScoreTB, 'Right')
        if ($CatScore -ge 0) {
            $CatScoreTB.Text = "$CatScore%"
            $CatScoreTB.Foreground = if ($CatScore -ge 80) {
                $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#15803D' } else { '#4ADE80' }))
            } elseif ($CatScore -ge 50) {
                $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#C2410C' } else { '#FBBF24' }))
            } else {
                $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#DC2626' } else { '#FB923C' }))
            }
        } else {
            $CatScoreTB.Text = '—'
            $CatScoreTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
        }
        [void]$CatDP.Children.Add($CatScoreTB)

        # Right: count badge
        $CountBadge = New-Object System.Windows.Controls.Border
        $CountBadge.CornerRadius = [System.Windows.CornerRadius]::new(10)
        $CountBadge.Padding = [System.Windows.Thickness]::new(8,2,8,2)
        $CountBadge.Margin  = [System.Windows.Thickness]::new(0,0,10,0)
        $CountBadge.VerticalAlignment = 'Center'
        $CountBadge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeDeepBg')
        $CountTB = New-Object System.Windows.Controls.TextBlock
        $CountTB.Text = "$CatPass/$($CatChecks.Count)"
        $CountTB.FontSize = 12
        $CountTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
        $CountBadge.Child = $CountTB
        [System.Windows.Controls.DockPanel]::SetDock($CountBadge, 'Right')
        [void]$CatDP.Children.Add($CountBadge)

        # Left: chevron + dot + name
        $CatLeftSP = New-Object System.Windows.Controls.StackPanel
        $CatLeftSP.Orientation = 'Horizontal'
        $CatLeftSP.VerticalAlignment = 'Center'

        $ChevronTB = New-Object System.Windows.Controls.TextBlock
        $ChevronTB.Text = [char]0xE76C  # ChevronRight = collapsed
        $ChevronTB.FontFamily = New-Object System.Windows.Media.FontFamily('Segoe Fluent Icons, Segoe MDL2 Assets')
        $ChevronTB.FontSize = 12
        $ChevronTB.VerticalAlignment = 'Center'
        $ChevronTB.Margin = [System.Windows.Thickness]::new(0,0,8,0)
        $ChevronTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')

        $CatDot = New-Object System.Windows.Shapes.Ellipse
        $CatDot.Width = 10; $CatDot.Height = 10
        $CatDot.Margin = [System.Windows.Thickness]::new(0,0,10,0)
        $CatDot.Fill = if ($CatScore -ge 80) {
            $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#15803D' } else { '#00C853' }))
        } elseif ($CatScore -ge 50) {
            $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#C2410C' } else { '#F59E0B' }))
        } elseif ($CatScore -ge 0) {
            $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#DC2626' } else { '#FF5000' }))
        } else {
            $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#595959' } else { '#82828C' }))
        }

        $CatNameTB = New-Object System.Windows.Controls.TextBlock
        $CatNameTB.Text = $Cat
        $CatNameTB.FontSize = 15; $CatNameTB.FontWeight = [System.Windows.FontWeights]::SemiBold
        $CatNameTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextPrimary')

        [void]$CatLeftSP.Children.Add($ChevronTB)
        [void]$CatLeftSP.Children.Add($CatDot)
        [void]$CatLeftSP.Children.Add($CatNameTB)
        [void]$CatDP.Children.Add($CatLeftSP)

        $CatCard.Child = $CatDP
        $CatCard.Cursor = [System.Windows.Input.Cursors]::Hand
        $CatCard.Tag = @{ Chevron = $ChevronTB; Container = $ChecksContainer }
        $CatCard.Add_MouseLeftButtonUp({
            $Info = $this.Tag
            if ($Info.Container.Visibility -eq 'Collapsed') {
                $Info.Container.Visibility = 'Visible'
                $Info.Chevron.Text = [char]0xE70D  # ChevronDown
            } else {
                $Info.Container.Visibility = 'Collapsed'
                $Info.Chevron.Text = [char]0xE76C  # ChevronRight
            }
        })
        [void]$pnlAssessmentChecks.Children.Add($CatCard)

        # Register refs for live score updates
        $Global:CatScoreRefs[$Cat] = @{
            ScoreTB    = $CatScoreTB
            CountTB    = $CountTB
            DotEllipse = $CatDot
            SbScoreTB  = $SbScoreTB
            SbDot      = $SbDot
        }

        # Wire sidebar button click: expand + scroll
        $CatBtn.Tag = @{ CatCard = $CatCard; Chevron = $ChevronTB; Container = $ChecksContainer }
        $CatBtn.Add_Click({
            $Info = $this.Tag
            if ($Info.Container.Visibility -eq 'Collapsed') {
                $Info.Container.Visibility = 'Visible'
                $Info.Chevron.Text = [char]0xE70D
            }
            $Info.CatCard.BringIntoView()
        })

        foreach ($Check in $CatChecks) {
            # Determine accent color for left bar based on current status
            $AccentColor = switch ($Check.Status) {
                'Pass'    { if ($Global:IsLightMode) { '#15803D' } else { '#00C853' } }
                'Warning' { if ($Global:IsLightMode) { '#C2410C' } else { '#F59E0B' } }
                'Fail'    { if ($Global:IsLightMode) { '#DC2626' } else { '#FF5000' } }
                default   { if ($Global:IsLightMode) { '#C0C0C4' } else { '#444448' } }
            }

            # Outer container with left accent bar
            $OuterGrid = New-Object System.Windows.Controls.Grid
            $OuterGrid.Margin = [System.Windows.Thickness]::new(0,0,0,4)
            $ColAccent = New-Object System.Windows.Controls.ColumnDefinition; $ColAccent.Width = [System.Windows.GridLength]::new(4)
            $ColContent = New-Object System.Windows.Controls.ColumnDefinition; $ColContent.Width = [System.Windows.GridLength]::new(1, 'Star')
            [void]$OuterGrid.ColumnDefinitions.Add($ColAccent)
            [void]$OuterGrid.ColumnDefinitions.Add($ColContent)

            # Left accent bar
            $AccentBar = New-Object System.Windows.Controls.Border
            $AccentBar.CornerRadius = [System.Windows.CornerRadius]::new(2,0,0,2)
            $AccentBar.Background = $Global:CachedBC.ConvertFromString($AccentColor)
            [System.Windows.Controls.Grid]::SetColumn($AccentBar, 0)
            [void]$OuterGrid.Children.Add($AccentBar)

            $CheckCard = New-Object System.Windows.Controls.Border
            $CheckCard.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeCardBg')
            $CheckCard.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderCard')
            $CheckCard.BorderThickness = [System.Windows.Thickness]::new(0,1,1,1)
            $CheckCard.CornerRadius    = [System.Windows.CornerRadius]::new(0,8,8,0)
            $CheckCard.Padding         = [System.Windows.Thickness]::new(14,10,14,10)

            $Grid = New-Object System.Windows.Controls.Grid
            $Col0 = New-Object System.Windows.Controls.ColumnDefinition; $Col0.Width = [System.Windows.GridLength]::new(1, 'Star')
            $Col1 = New-Object System.Windows.Controls.ColumnDefinition; $Col1.Width = [System.Windows.GridLength]::new(130)
            $Col2 = New-Object System.Windows.Controls.ColumnDefinition; $Col2.Width = [System.Windows.GridLength]::new(50)
            $Col3 = New-Object System.Windows.Controls.ColumnDefinition; $Col3.Width = [System.Windows.GridLength]::new(60)
            [void]$Grid.ColumnDefinitions.Add($Col0)
            [void]$Grid.ColumnDefinitions.Add($Col1)
            [void]$Grid.ColumnDefinitions.Add($Col2)
            [void]$Grid.ColumnDefinitions.Add($Col3)

            # Left: Check info
            $InfoSP = New-Object System.Windows.Controls.StackPanel
            $NameTB = New-Object System.Windows.Controls.TextBlock
            $NameTB.Text = "$($Check.Id): $($Check.Name)"
            $NameTB.FontSize = 14; $NameTB.FontWeight = [System.Windows.FontWeights]::SemiBold
            $NameTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
            $DescTB = New-Object System.Windows.Controls.TextBlock
            $DescTB.Text = $Check.Description
            $DescTB.FontSize = 12; $DescTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
            $DescTB.TextWrapping = 'Wrap'; $DescTB.Margin = [System.Windows.Thickness]::new(0,2,0,0)
            [void]$InfoSP.Children.Add($NameTB)
            [void]$InfoSP.Children.Add($DescTB)

            # Origin badges + Auto/Manual badge
            $TagSP = New-Object System.Windows.Controls.StackPanel
            $TagSP.Orientation = 'Horizontal'
            $TagSP.Margin = [System.Windows.Thickness]::new(0,3,0,0)
            if ($Check.Origin) {
                foreach ($Tag in ($Check.Origin -split ',')) {
                    $Badge = New-Object System.Windows.Controls.Border
                    $Badge.CornerRadius = [System.Windows.CornerRadius]::new(3)
                    $Badge.Padding = [System.Windows.Thickness]::new(5,1,5,1)
                    $Badge.Margin = [System.Windows.Thickness]::new(0,0,4,0)
                    $Badge.Background = $Global:CachedBC.ConvertFromString($(switch ($Tag.Trim()) {
                        'WAF' { if ($Global:IsLightMode) { '#E0F0FF' } else { '#0A1E3D' } }
                        'CAF' { if ($Global:IsLightMode) { '#E8F5E9' } else { '#0A3D1A' } }
                        'LZA' { if ($Global:IsLightMode) { '#FFF3E0' } else { '#3D2D0A' } }
                        'SEC' { if ($Global:IsLightMode) { '#FEE2E2' } else { '#3D0A0A' } }
                        'FSL' { if ($Global:IsLightMode) { '#F3E5F5' } else { '#2D0A3D' } }
                        'AVD' { if ($Global:IsLightMode) { '#E0E0E0' } else { '#1F1F23' } }
                        default { if ($Global:IsLightMode) { '#E0E0E0' } else { '#1F1F23' } }
                    }))
                    $Badge.ToolTip = switch ($Tag.Trim()) {
                        'WAF' { 'Well-Architected Framework for AVD' }
                        'CAF' { 'Cloud Adoption Framework for AVD' }
                        'LZA' { 'AVD Landing Zone Accelerator' }
                        'SEC' { 'AVD Security Recommendations' }
                        'FSL' { 'FSLogix Documentation' }
                        'AVD' { 'Azure Virtual Desktop Documentation' }
                        default { $Tag.Trim() }
                    }
                    $BadgeTB = New-Object System.Windows.Controls.TextBlock
                    $BadgeTB.Text = $Tag.Trim()
                    $BadgeTB.FontSize = 11; $BadgeTB.FontWeight = [System.Windows.FontWeights]::SemiBold
                    $BadgeTB.Foreground = $Global:CachedBC.ConvertFromString($(switch ($Tag.Trim()) {
                        'WAF' { if ($Global:IsLightMode) { '#0078D4' } else { '#5BB8F5' } }
                        'CAF' { if ($Global:IsLightMode) { '#00873D' } else { '#34D399' } }
                        'LZA' { if ($Global:IsLightMode) { '#C2410C' } else { '#FB923C' } }
                        'SEC' { if ($Global:IsLightMode) { '#DC2626' } else { '#F87171' } }
                        'FSL' { if ($Global:IsLightMode) { '#7B1FA2' } else { '#C084FC' } }
                        'AVD' { if ($Global:IsLightMode) { '#555' } else { '#A1A1AA' } }
                        default { if ($Global:IsLightMode) { '#555' } else { '#A1A1AA' } }
                    }))
                    $Badge.Child = $BadgeTB
                    [void]$TagSP.Children.Add($Badge)
                }
            }
            # Auto / Manual type badge
            $TypeLabel = if ($Check.Type -eq 'Auto') { 'AUTO' } else { 'MANUAL' }
            $TypeBadge = New-Object System.Windows.Controls.Border
            $TypeBadge.CornerRadius = [System.Windows.CornerRadius]::new(3)
            $TypeBadge.Padding = [System.Windows.Thickness]::new(5,1,5,1)
            $TypeBadge.Margin = [System.Windows.Thickness]::new(0,0,4,0)
            $TypeBadge.Background = $Global:CachedBC.ConvertFromString($(if ($Check.Type -eq 'Auto') {
                if ($Global:IsLightMode) { '#E0F7FA' } else { '#0A2E3D' }
            } else {
                if ($Global:IsLightMode) { '#FFF8E1' } else { '#3D350A' }
            }))
            $TypeBadge.ToolTip = if ($Check.Type -eq 'Auto') {
                'Automated check — evaluated automatically via discovery data'
            } else {
                'Manual check — requires manual verification and status update'
            }
            $TypeTB = New-Object System.Windows.Controls.TextBlock
            $TypeTB.Text = $TypeLabel
            $TypeTB.FontSize = 11; $TypeTB.FontWeight = [System.Windows.FontWeights]::SemiBold
            $TypeTB.Foreground = $Global:CachedBC.ConvertFromString($(if ($Check.Type -eq 'Auto') {
                if ($Global:IsLightMode) { '#00838F' } else { '#4DD0E1' }
            } else {
                if ($Global:IsLightMode) { '#F9A825' } else { '#FFD54F' }
            }))
            $TypeBadge.Child = $TypeTB
            [void]$TagSP.Children.Add($TypeBadge)
            [void]$InfoSP.Children.Add($TagSP)

            # Reference URL link
            if ($Check.Reference) {
                $RefSP = New-Object System.Windows.Controls.StackPanel
                $RefSP.Orientation = 'Horizontal'
                $RefSP.Margin = [System.Windows.Thickness]::new(0,3,0,0)
                $RefIcon = New-Object System.Windows.Controls.TextBlock
                $RefIcon.Text = [char]0xE71B
                $RefIcon.FontFamily = New-Object System.Windows.Media.FontFamily('Segoe Fluent Icons, Segoe MDL2 Assets')
                $RefIcon.FontSize = 11; $RefIcon.VerticalAlignment = 'Center'
                $RefIcon.Foreground = $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#0078D4' } else { '#60CDFF' }))
                $RefIcon.Margin = [System.Windows.Thickness]::new(0,0,5,0)
                $RefLink = New-Object System.Windows.Controls.TextBlock
                $RefLink.Text = ($Check.Reference -replace 'https://learn.microsoft.com/en-us/', '' -replace 'https://', '')
                $RefLink.FontSize = 11
                $RefLink.Foreground = $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#0078D4' } else { '#60CDFF' }))
                $RefLink.TextDecorations = [System.Windows.TextDecorations]::Underline
                $RefLink.Cursor = [System.Windows.Input.Cursors]::Hand
                $RefLink.TextTrimming = 'CharacterEllipsis'
                $RefLink.ToolTip = $Check.Reference
                $RefLink.Tag = $Check.Reference
                $RefLink.Add_MouseLeftButtonUp({
                    Start-Process $this.Tag
                })
                [void]$RefSP.Children.Add($RefIcon)
                [void]$RefSP.Children.Add($RefLink)
                [void]$InfoSP.Children.Add($RefSP)
            }
            if ($Check.Details) {
                $DetTB = New-Object System.Windows.Controls.TextBlock
                $DetTB.Text = $Check.Details
                $DetTB.FontSize = 12; $DetTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextFaintest')
                $DetTB.TextWrapping = 'Wrap'; $DetTB.Margin = [System.Windows.Thickness]::new(0,2,0,0)
                [void]$InfoSP.Children.Add($DetTB)
            }

            # Per-check notes input
            $NotesTB = New-Object System.Windows.Controls.TextBox
            $NotesTB.Text = if ($Check.Notes) { $Check.Notes } else { '' }
            $NotesTB.FontSize = 12
            $NotesTB.TextWrapping = 'Wrap'
            $NotesTB.AcceptsReturn = $false
            $NotesTB.Height = 22
            $NotesTB.Margin = [System.Windows.Thickness]::new(0,4,0,0)
            $NotesTB.Padding = [System.Windows.Thickness]::new(6,2,6,2)
            $NotesTB.ToolTip = 'Add notes (e.g., why N/A, observations, evidence)'
            $NotesTB.Tag = $Check.Id
            $NotesTB.SetResourceReference([System.Windows.Controls.TextBox]::BackgroundProperty, 'ThemeInputBg')
            $NotesTB.SetResourceReference([System.Windows.Controls.TextBox]::ForegroundProperty, 'ThemeTextSecondary')
            $NotesTB.SetResourceReference([System.Windows.Controls.TextBox]::BorderBrushProperty, 'ThemeBorderSubtle')
            $NotesTB.BorderThickness = [System.Windows.Thickness]::new(1)
            # Placeholder via Tag trick — set muted text if empty
            if (-not $NotesTB.Text) {
                $NotesTB.SetResourceReference([System.Windows.Controls.TextBox]::ForegroundProperty, 'ThemeTextFaintest')
                $NotesTB.Text = 'Add notes...'
            }
            $NotesTB.Add_GotFocus({
                if ($this.Text -eq 'Add notes...') {
                    $this.Text = ''
                    $this.SetResourceReference([System.Windows.Controls.TextBox]::ForegroundProperty, 'ThemeTextSecondary')
                }
            })
            $NotesTB.Add_LostFocus({
                $CheckId = $this.Tag
                $Target = $Global:Assessment.Checks | Where-Object Id -eq $CheckId
                if ($Target) {
                    $Val = if ($this.Text -eq 'Add notes...') { '' } else { $this.Text }
                    $Target.Notes = $Val
                    if ($Val) { AutoSave-Assessment }
                }
                if (-not $this.Text) {
                    $this.SetResourceReference([System.Windows.Controls.TextBox]::ForegroundProperty, 'ThemeTextFaintest')
                    $this.Text = 'Add notes...'
                }
            })
            [void]$InfoSP.Children.Add($NotesTB)

            [System.Windows.Controls.Grid]::SetColumn($InfoSP, 0)
            [void]$Grid.Children.Add($InfoSP)

            # Middle: Status dropdown
            $StatusCmb = New-Object System.Windows.Controls.ComboBox
            $StatusCmb.Style = $Window.FindResource('DarkComboBox')
            $StatusCmb.Height = 30; $StatusCmb.FontSize = 12; $StatusCmb.VerticalContentAlignment = 'Center'
            $StatusCmb.Margin = [System.Windows.Thickness]::new(8,0,8,0)
            foreach ($S in @('Not Assessed','Pass','Warning','Fail','N/A')) {
                $Item = New-Object System.Windows.Controls.ComboBoxItem
                $Item.Content = $S
                [void]$StatusCmb.Items.Add($Item)
            }
            # Set current value
            for ($i = 0; $i -lt $StatusCmb.Items.Count; $i++) {
                if ($StatusCmb.Items[$i].Content -eq $Check.Status) {
                    $StatusCmb.SelectedIndex = $i; break
                }
            }
            $StatusCmb.Tag = $Check.Id
            $StatusCmb.Add_SelectionChanged({
                $CheckId = $this.Tag
                $NewStatus = $this.SelectedItem.Content
                $Target = $Global:Assessment.Checks | Where-Object Id -eq $CheckId
                if ($Target) {
                    $Target.Status = $NewStatus
                    $Target.Source = 'Manual'
                    # Repaint the left accent bar to reflect new status
                    $Bar = $this.Parent.Parent.Parent.Children[0]
                    $BarColor = switch ($NewStatus) {
                        'Pass'    { if ($Global:IsLightMode) { '#15803D' } else { '#00C853' } }
                        'Warning' { if ($Global:IsLightMode) { '#C2410C' } else { '#F59E0B' } }
                        'Fail'    { if ($Global:IsLightMode) { '#DC2626' } else { '#FF5000' } }
                        default   { if ($Global:IsLightMode) { '#C0C0C4' } else { '#444448' } }
                    }
                    $Bar.Background = $Global:CachedBC.ConvertFromString($BarColor)
                    Update-Progress
                    AutoSave-Assessment
                }
            })
            [System.Windows.Controls.Grid]::SetColumn($StatusCmb, 1)
            [void]$Grid.Children.Add($StatusCmb)

            # Col 2: Weight indicator (dots)
            $WeightTB = New-Object System.Windows.Controls.TextBlock
            $WeightVal = [math]::Max(1, [math]::Min(5, [int]$Check.Weight))
            $WeightTB.Text = ("$([char]0x25CF)" * $WeightVal) + ("$([char]0x25CB)" * (5 - $WeightVal))
            $WeightTB.FontSize = 7; $WeightTB.ToolTip = "Weight: $WeightVal/5"
            $WeightTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
            $WeightTB.HorizontalAlignment = 'Center'
            $WeightTB.VerticalAlignment = 'Center'
            [System.Windows.Controls.Grid]::SetColumn($WeightTB, 2)
            [void]$Grid.Children.Add($WeightTB)

            # Col 3: Severity pill badge
            $SevPill = New-Object System.Windows.Controls.Border
            $SevPill.CornerRadius = [System.Windows.CornerRadius]::new(10)
            $SevPill.Padding = [System.Windows.Thickness]::new(8,2,8,2)
            $SevPill.HorizontalAlignment = 'Center'; $SevPill.VerticalAlignment = 'Center'
            $SevColors = switch ($Check.Severity) {
                'Critical' { @{ Bg = $(if ($Global:IsLightMode) { '#FEE2E2' } else { '#3D0A0A' }); Fg = $(if ($Global:IsLightMode) { '#DC2626' } else { '#FB923C' }) } }
                'High'     { @{ Bg = $(if ($Global:IsLightMode) { '#FFF7ED' } else { '#3D2D0A' }); Fg = $(if ($Global:IsLightMode) { '#C2410C' } else { '#FBBF24' }) } }
                'Medium'   { @{ Bg = $(if ($Global:IsLightMode) { '#F4F4F5' } else { '#1F1F23' }); Fg = $(if ($Global:IsLightMode) { '#3F3F46' } else { '#A1A1AA' }) } }
                'Low'      { @{ Bg = $(if ($Global:IsLightMode) { '#F4F4F5' } else { '#1F1F23' }); Fg = $(if ($Global:IsLightMode) { '#52525B' } else { '#71717A' }) } }
                default    { @{ Bg = $(if ($Global:IsLightMode) { '#F4F4F5' } else { '#1F1F23' }); Fg = $(if ($Global:IsLightMode) { '#52525B' } else { '#71717A' }) } }
            }
            $SevPill.Background = $Global:CachedBC.ConvertFromString($SevColors.Bg)
            $SevBadge = New-Object System.Windows.Controls.TextBlock
            $SevBadge.Text = $Check.Severity
            $SevBadge.FontSize = 11; $SevBadge.FontWeight = [System.Windows.FontWeights]::SemiBold
            $SevBadge.Foreground = $Global:CachedBC.ConvertFromString($SevColors.Fg)
            $SevPill.Child = $SevBadge
            [System.Windows.Controls.Grid]::SetColumn($SevPill, 3)
            [void]$Grid.Children.Add($SevPill)

            $CheckCard.Child = $Grid
            [System.Windows.Controls.Grid]::SetColumn($CheckCard, 1)
            [void]$OuterGrid.Children.Add($CheckCard)
            [void]$ChecksContainer.Children.Add($OuterGrid)
        }
        [void]$pnlAssessmentChecks.Children.Add($ChecksContainer)
    }
    Update-Progress
}

<#
.SYNOPSIS
    Renders the filtered findings view showing only Fail and Warning checks.
.DESCRIPTION
    Applies severity, status, category, and search text filters to display actionable findings
    with severity badges, status indicators, and direct links to remediation references.
#>
function Render-Findings {
    # Clear dynamic content
    $KeepCount = 2
    while ($pnlFindingsList.Children.Count -gt $KeepCount) {
        $pnlFindingsList.Children.RemoveAt($pnlFindingsList.Children.Count - 1)
    }

    # Get filter values
    $SevFilter = if ($cmbFilterSeverity.SelectedIndex -gt 0) { $cmbFilterSeverity.SelectedItem.Content } else { $null }
    $StatFilter = if ($cmbFilterStatus.SelectedIndex -gt 0) { $cmbFilterStatus.SelectedItem.Content } else { $null }
    $CatFilter = if ($cmbFilterCategory.SelectedIndex -gt 0) { $cmbFilterCategory.SelectedItem.Content } else { $null }
    $SearchText = $txtFindingsSearch.Text

    $Filtered = @($Global:Assessment.Checks | Where-Object {
        $_.Status -ne 'Not Assessed' -and
        (-not $SevFilter -or $_.Severity -eq $SevFilter) -and
        (-not $StatFilter -or $_.Status -eq $StatFilter) -and
        (-not $CatFilter -or $_.Category -eq $CatFilter) -and
        (-not $SearchText -or $_.Name -like "*$SearchText*" -or $_.Description -like "*$SearchText*")
    })

    $lblFindingsCount.Text = "$($Filtered.Count) findings"

    foreach ($Check in $Filtered) {
        # Accent color for left bar
        $AccColor = switch ($Check.Status) {
            'Pass'    { if ($Global:IsLightMode) { '#15803D' } else { '#00C853' } }
            'Warning' { if ($Global:IsLightMode) { '#C2410C' } else { '#F59E0B' } }
            'Fail'    { if ($Global:IsLightMode) { '#DC2626' } else { '#FF5000' } }
            default   { if ($Global:IsLightMode) { '#C0C0C4' } else { '#444448' } }
        }

        $OuterGrid = New-Object System.Windows.Controls.Grid
        $OuterGrid.Margin = [System.Windows.Thickness]::new(0,0,0,6)
        $ColAcc = New-Object System.Windows.Controls.ColumnDefinition; $ColAcc.Width = [System.Windows.GridLength]::new(4)
        $ColCnt = New-Object System.Windows.Controls.ColumnDefinition; $ColCnt.Width = [System.Windows.GridLength]::new(1, 'Star')
        [void]$OuterGrid.ColumnDefinitions.Add($ColAcc)
        [void]$OuterGrid.ColumnDefinitions.Add($ColCnt)

        $AccBar = New-Object System.Windows.Controls.Border
        $AccBar.CornerRadius = [System.Windows.CornerRadius]::new(2,0,0,2)
        $AccBar.Background = $Global:CachedBC.ConvertFromString($AccColor)
        [System.Windows.Controls.Grid]::SetColumn($AccBar, 0)
        [void]$OuterGrid.Children.Add($AccBar)

        $Card = New-Object System.Windows.Controls.Border
        $Card.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeCardBg')
        $Card.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderCard')
        $Card.BorderThickness = [System.Windows.Thickness]::new(0,1,1,1)
        $Card.CornerRadius    = [System.Windows.CornerRadius]::new(0,8,8,0)
        $Card.Padding         = [System.Windows.Thickness]::new(14,10,14,10)

        $SP = New-Object System.Windows.Controls.StackPanel

        # Row 1: Status badge + category + name
        $HeaderDP = New-Object System.Windows.Controls.DockPanel

        # Status pill
        $StatusPill = New-Object System.Windows.Controls.Border
        $StatusPill.CornerRadius = [System.Windows.CornerRadius]::new(10)
        $StatusPill.Padding = [System.Windows.Thickness]::new(8,2,8,2)
        $StatusPill.VerticalAlignment = 'Center'
        $StatusPill.Background = $Global:CachedBC.ConvertFromString($(switch ($Check.Status) {
            'Fail'    { if ($Global:IsLightMode) { '#FEE2E2' } else { '#3D0A0A' } }
            'Warning' { if ($Global:IsLightMode) { '#FFF7ED' } else { '#3D2D0A' } }
            'Pass'    { if ($Global:IsLightMode) { '#E8F5E9' } else { '#0A3D1A' } }
            default   { if ($Global:IsLightMode) { '#F4F4F5' } else { '#1F1F23' } }
        }))
        $StatusTB = New-Object System.Windows.Controls.TextBlock
        $StatusTB.Text = $Check.Status
        $StatusTB.FontSize = 11; $StatusTB.FontWeight = [System.Windows.FontWeights]::Bold
        $StatusTB.Foreground = $Global:CachedBC.ConvertFromString($(switch ($Check.Status) {
            'Fail'    { if ($Global:IsLightMode) { '#DC2626' } else { '#FB923C' } }
            'Warning' { if ($Global:IsLightMode) { '#C2410C' } else { '#FBBF24' } }
            'Pass'    { if ($Global:IsLightMode) { '#15803D' } else { '#4ADE80' } }
            default   { if ($Global:IsLightMode) { '#71717A' } else { '#A1A1AA' } }
        }))
        $StatusPill.Child = $StatusTB
        [System.Windows.Controls.DockPanel]::SetDock($StatusPill, 'Right')
        [void]$HeaderDP.Children.Add($StatusPill)

        $NameTB = New-Object System.Windows.Controls.TextBlock
        $NameTB.Text = "[$($Check.Category)] $($Check.Id): $($Check.Name)"
        $NameTB.FontSize = 13; $NameTB.FontWeight = [System.Windows.FontWeights]::SemiBold
        $NameTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
        $NameTB.TextTrimming = 'CharacterEllipsis'
        [void]$HeaderDP.Children.Add($NameTB)
        [void]$SP.Children.Add($HeaderDP)

        # Row 2: Description
        $DescTB = New-Object System.Windows.Controls.TextBlock
        $DescTB.Text = $Check.Description
        $DescTB.FontSize = 12; $DescTB.TextWrapping = 'Wrap'
        $DescTB.Margin = [System.Windows.Thickness]::new(0,4,0,0)
        $DescTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
        [void]$SP.Children.Add($DescTB)

        # Row 3: Details if present
        if ($Check.Details) {
            $DetTB = New-Object System.Windows.Controls.TextBlock
            $DetTB.Text = $Check.Details; $DetTB.FontSize = 12
            $DetTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextFaintest')
            $DetTB.TextWrapping = 'Wrap'; $DetTB.Margin = [System.Windows.Thickness]::new(0,2,0,0)
            [void]$SP.Children.Add($DetTB)
        }

        # Row 4: Reference URL if present
        if ($Check.Reference) {
            $RefTB = New-Object System.Windows.Controls.TextBlock
            $RefTB.Text = $Check.Reference; $RefTB.FontSize = 11
            $RefTB.Foreground = $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#0078D4' } else { '#60CDFF' }))
            $RefTB.Margin = [System.Windows.Thickness]::new(0,2,0,0)
            $RefTB.TextTrimming = 'CharacterEllipsis'
            [void]$SP.Children.Add($RefTB)
        }

        $Card.Child = $SP
        [System.Windows.Controls.Grid]::SetColumn($Card, 1)
        [void]$OuterGrid.Children.Add($Card)
        [void]$pnlFindingsList.Children.Add($OuterGrid)
    }
}

<#
.SYNOPSIS
    Generates a text-based report preview in the Report tab for quick review before export.
.DESCRIPTION
    Builds a structured text summary with overall score, category breakdowns, pass/warn/fail
    counts, and individual check details for non-passing items.
#>
function Update-ReportPreview {
    <# Builds a text-based report summary in the RichTextBox #>
    if ($null -eq $paraReportSummary) { return }
    $paraReportSummary.Inlines.Clear()

    $Overall = Get-OverallScore
    $Categories = Get-Categories

    # Helper to add colored text
    $AddText = { param($Text, $Color)
        $Run = New-Object System.Windows.Documents.Run($Text)
        $Run.Foreground = $Global:CachedBC.ConvertFromString($Color)
        $paraReportSummary.Inlines.Add($Run)
    }
    $AddLine = {
        $paraReportSummary.Inlines.Add([System.Windows.Documents.LineBreak]::new())
    }

    $Fg = if ($Global:IsLightMode) { '#222222' } else { '#E0E0E0' }
    $Dim = if ($Global:IsLightMode) { '#666666' } else { '#71717A' }
    $Green = if ($Global:IsLightMode) { '#15803D' } else { '#4ADE80' }
    $Orange = if ($Global:IsLightMode) { '#C2410C' } else { '#FBBF24' }
    $Red = if ($Global:IsLightMode) { '#DC2626' } else { '#FB923C' }
    $Blue = if ($Global:IsLightMode) { '#0078D4' } else { '#60CDFF' }

    # Title
    & $AddText 'AVD ASSESSMENT REPORT SUMMARY' $Blue
    & $AddLine
    & $AddText ('=' * 50) $Dim
    & $AddLine; & $AddLine

    # Metadata
    & $AddText "Customer:  $($Global:Assessment.CustomerName)" $Fg; & $AddLine
    & $AddText "Assessor:  $($Global:Assessment.AssessorName)" $Fg; & $AddLine
    & $AddText "Date:      $($Global:Assessment.Date)" $Fg; & $AddLine
    & $AddText "Checks:    $($Global:Assessment.Checks.Count) total" $Dim; & $AddLine
    & $AddLine

    # Overall score
    $ScoreColor = if ($Overall -ge 80) { $Green } elseif ($Overall -ge 50) { $Orange } else { $Red }
    & $AddText "OVERALL READINESS: " $Fg
    & $AddText "$(if ($Overall -ge 0) { "$Overall%" } else { 'Not scored' })" $ScoreColor
    & $AddLine; & $AddLine

    # Totals
    $Assessed = @($Global:Assessment.Checks | Where-Object { $_.Status -ne 'Not Assessed' -and -not $_.Excluded })
    $Pass = @($Assessed | Where-Object { $_.Status -eq 'Pass' -or $_.Status -eq 'N/A' }).Count
    $Warn = @($Assessed | Where-Object Status -eq 'Warning').Count
    $Fail = @($Assessed | Where-Object Status -eq 'Fail').Count
    $Excl = @($Global:Assessment.Checks | Where-Object { $_.Excluded }).Count

    & $AddText "Pass: $Pass  " $Green
    & $AddText "Warning: $Warn  " $Orange
    & $AddText "Fail: $Fail  " $Red
    & $AddText "Excluded: $Excl" $Dim
    & $AddLine
    & $AddText ('-' * 50) $Dim
    & $AddLine; & $AddLine

    # Category breakdown
    & $AddText 'CATEGORY BREAKDOWN' $Blue
    & $AddLine; & $AddLine

    foreach ($Cat in $Categories) {
        $Score = Get-CategoryScore $Cat
        $CatChecks = @($Global:Assessment.Checks | Where-Object Category -eq $Cat)
        $CatAssessed = @($CatChecks | Where-Object { $_.Status -ne 'Not Assessed' -and -not $_.Excluded })
        $CatPass = @($CatAssessed | Where-Object { $_.Status -eq 'Pass' -or $_.Status -eq 'N/A' }).Count
        $CatColor = if ($Score -ge 80) { $Green } elseif ($Score -ge 50) { $Orange } elseif ($Score -ge 0) { $Red } else { $Dim }

        & $AddText "  $($Cat.PadRight(25))" $Fg
        & $AddText "$(if ($Score -ge 0) { "$Score%" } else { '---' })" $CatColor
        & $AddText "  ($CatPass/$($CatAssessed.Count) pass, $($CatChecks.Count) total)" $Dim
        & $AddLine
    }

    & $AddLine
    & $AddText ('-' * 50) $Dim
    & $AddLine; & $AddLine

    # Findings (Fail + Warning)
    $Findings = @($Global:Assessment.Checks | Where-Object { $_.Status -eq 'Fail' -or $_.Status -eq 'Warning' } | Sort-Object @{E={switch($_.Status){'Fail'{0}'Warning'{1}default{2}}}}, Category)

    if ($Findings.Count -gt 0) {
        & $AddText "FINDINGS ($($Findings.Count))" $Blue
        & $AddLine; & $AddLine

        foreach ($F in $Findings) {
            $Icon = if ($F.Status -eq 'Fail') { '[FAIL]   ' } else { '[WARN]   ' }
            $FColor = if ($F.Status -eq 'Fail') { $Red } else { $Orange }
            & $AddText $Icon $FColor
            & $AddText "$($F.Id): $($F.Name)" $Fg
            & $AddLine
            if ($F.Details) {
                & $AddText "           $($F.Details)" $Dim
                & $AddLine
            }
            if ($F.Reference) {
                & $AddText "           $($F.Reference)" $Blue
                & $AddLine
            }
        }
    } else {
        & $AddText 'No findings (all checks pass or not assessed).' $Green
        & $AddLine
    }

    & $AddLine
    & $AddText "Export to HTML for the full styled report with collapsible sections." $Dim
    & $AddLine
}

# ===============================================================================
# SECTION 13: HTML REPORT GENERATION
# Comprehensive standalone HTML report with embedded CSS/JS: score hero, summary stat cards,
# segmented progress bar, maturity hexagonal radar SVG, zone strip, per-dimension gauges,
# effort matrix, collapsible category sections with check tables, print support, theme toggle.
# ===============================================================================

<#
.SYNOPSIS
    Generates a comprehensive standalone HTML assessment report with embedded CSS and JavaScript.
.DESCRIPTION
    Builds a responsive HTML report with: score hero, summary stat cards, segmented progress bar,
    maturity hexagonal radar SVG, maturity zone strip, per-dimension gauge cards, effort matrix,
    collapsible category sections with check tables, severity badges, weight dots, origin tags,
    impact/remediation boxes, notes callouts, category data summaries, and field notes. Includes
    print support, theme toggle, and expand/collapse controls via embedded JavaScript.
.OUTPUTS
    String containing the complete HTML document.
#>
function Build-HtmlReport {
    $html = [System.Text.StringBuilder]::new(65536)
    $enc  = { param($s) [System.Net.WebUtility]::HtmlEncode($s) }
    $Overall = Get-OverallScore

    $Assessed  = @($Global:Assessment.Checks | Where-Object { $_.Status -ne 'Not Assessed' -and -not $_.Excluded })
    $PassCount = @($Assessed | Where-Object { $_.Status -eq 'Pass' }).Count
    $NACount   = @($Assessed | Where-Object { $_.Status -eq 'N/A' }).Count
    $WarnCount = @($Assessed | Where-Object Status -eq 'Warning').Count
    $FailCount = @($Assessed | Where-Object Status -eq 'Fail').Count
    $ExclCount = @($Global:Assessment.Checks | Where-Object { $_.Excluded }).Count
    $OverallClass = if ($Overall -ge 80) { 'green' } elseif ($Overall -ge 50) { 'orange' } else { 'red' }

    [void]$html.Append(@"
<!DOCTYPE html>
<html lang="en" data-theme="dark">
<head>
<meta http-equiv="X-UA-Compatible" content="IE=edge"/>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AVD Assessment Report$(if($Global:Assessment.CustomerName){" — $(& $enc $Global:Assessment.CustomerName)"})</title>
<style>
:root {
  --font:'Inter','Segoe UI Variable Display','Segoe UI',-apple-system,sans-serif;
  --bg:#0a0a0c;--bg2:#111113;--surface:#141416;--card:#1a1a1e;--card-hover:#222226;
  --card-border:#2a2a2e;--border:rgba(255,255,255,0.06);--border-strong:rgba(255,255,255,0.12);
  --text:#e4e4e7;--text-bright:#fff;--text-secondary:#a1a1aa;--text-dim:#71717a;--text-faint:#52525b;
  --accent:#60cdff;--accent-dim:rgba(96,205,255,0.1);--accent-text:#60cdff;
  --green:#00c853;--green-dim:rgba(0,200,83,0.12);--green-text:#4ade80;
  --red:#ff5000;--red-dim:rgba(255,80,0,0.12);--red-text:#fb923c;
  --orange:#f59e0b;--orange-dim:rgba(245,158,11,0.12);--orange-text:#fbbf24;
  --radius:12px;--radius-sm:8px;--radius-xs:6px;
  --shadow:0 1px 3px rgba(0,0,0,0.3);
}
[data-theme="light"]{
  --bg:#f8f9fa;--bg2:#fff;--surface:#f4f4f5;--card:#fff;--card-hover:#f0f0f2;
  --card-border:#e4e4e7;--border:rgba(0,0,0,0.08);--border-strong:rgba(0,0,0,0.15);
  --text:#18181b;--text-bright:#09090b;--text-secondary:#52525b;--text-dim:#71717a;--text-faint:#a1a1aa;
  --accent:#0078d4;--accent-dim:rgba(0,120,212,0.08);--accent-text:#0066b8;
  --green:#16a34a;--green-dim:rgba(22,163,74,0.08);--green-text:#15803d;
  --red:#dc2626;--red-dim:rgba(220,38,38,0.08);--red-text:#b91c1c;
  --orange:#ea580c;--orange-dim:rgba(234,88,12,0.08);--orange-text:#c2410c;
  --shadow:0 1px 3px rgba(0,0,0,0.06),0 4px 12px rgba(0,0,0,0.04);
}
*,*::before,*::after{margin:0;padding:0;box-sizing:border-box}
html{scroll-behavior:smooth}
body{font-family:var(--font);background:var(--bg);color:var(--text);line-height:1.5}

/* Toolbar */
.toolbar{position:sticky;top:0;z-index:100;background:var(--bg2);border-bottom:1px solid var(--border);
  padding:12px 32px;display:flex;align-items:center;gap:14px;backdrop-filter:blur(12px)}
.toolbar h1{font-size:16px;font-weight:600;color:var(--text-bright);white-space:nowrap;display:flex;align-items:center;gap:8px}
.toolbar .spacer{flex:1}
.btn{background:var(--card);border:1px solid var(--card-border);border-radius:var(--radius-sm);
  padding:5px 12px;color:var(--text-secondary);font-size:11px;cursor:pointer;font-family:var(--font);transition:all 0.15s}
.btn:hover{background:var(--card-hover);border-color:var(--border-strong);color:var(--text)}

/* Container */
.container{max-width:1200px;margin:0 auto;padding:28px 40px 60px}

/* Meta */
.meta-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:6px;margin-bottom:28px;
  background:var(--card);border:1px solid var(--card-border);border-radius:var(--radius);padding:16px 20px}
.meta-item{font-size:11px;color:var(--text-dim)}
.meta-item span{font-weight:600;color:var(--text)}

/* Score Hero */
.score-hero{text-align:center;padding:32px 0 28px;margin-bottom:28px;border-bottom:1px solid var(--border)}
.score-hero .score-value{font-size:72px;font-weight:800;letter-spacing:-3px;line-height:1}
.score-hero .score-value.green{color:var(--green)} .score-hero .score-value.orange{color:var(--orange)} .score-hero .score-value.red{color:var(--red)}
.score-hero .score-label{font-size:13px;color:var(--text-dim);margin-top:4px;text-transform:uppercase;letter-spacing:1px}

/* Summary cards */
.summary{display:grid;grid-template-columns:repeat(5,1fr);gap:14px;margin-bottom:32px}
.stat-card{background:var(--card);border:1px solid var(--card-border);border-radius:var(--radius);padding:18px 22px;position:relative;overflow:hidden;transition:border-color 0.2s}
.stat-card:hover{border-color:var(--border-strong)}
.stat-card .accent-bar{position:absolute;top:0;left:0;right:0;height:3px;border-radius:var(--radius) var(--radius) 0 0}
.stat-card .num{font-size:32px;font-weight:800;letter-spacing:-1px;line-height:1}
.stat-card .label{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:0.5px;color:var(--text-dim);margin-top:5px}
.stat-card.pass .num{color:var(--green)} .stat-card.pass .accent-bar{background:var(--green)}
.stat-card.warn .num{color:var(--orange)} .stat-card.warn .accent-bar{background:var(--orange)}
.stat-card.fail .num{color:var(--red)} .stat-card.fail .accent-bar{background:var(--red)}
.stat-card.total .num{color:var(--accent)} .stat-card.total .accent-bar{background:var(--accent)}
.stat-card.excluded .num{color:var(--text-dim)} .stat-card.excluded .accent-bar{background:var(--text-faint)}
.stat-card.na .num{color:var(--text-secondary)} .stat-card.na .accent-bar{background:var(--text-secondary)}

/* Progress bar */
.progress-bar{display:flex;height:8px;border-radius:4px;overflow:hidden;margin-bottom:32px;background:var(--card);border:1px solid var(--card-border)}
.progress-bar .seg{transition:width 0.6s ease}
.seg-pass{background:var(--green)} .seg-warn{background:var(--orange)} .seg-fail{background:var(--red)} .seg-na{background:var(--card-border)}

/* Sections */
.section-header{display:flex;align-items:center;gap:10px;padding:14px 20px;cursor:pointer;
  background:var(--card);border:1px solid var(--card-border);border-radius:var(--radius);margin-bottom:2px;
  font-size:14px;font-weight:600;color:var(--text);user-select:none;transition:border-color 0.2s}
.section-header:hover{border-color:var(--border-strong)}
.section-header .chevron{font-size:10px;color:var(--text-dim);transition:transform 0.2s}
.section-header.open .chevron{transform:rotate(90deg)}
.section-header .section-dot{width:9px;height:9px;border-radius:50%;flex-shrink:0}
.section-header .section-score{font-size:12px;font-weight:500;margin-left:auto}
.section-header .section-count{font-size:10px;color:var(--text-faint);background:var(--surface);border:1px solid var(--card-border);border-radius:12px;padding:1px 9px}
.section-content{display:none;margin-bottom:12px}
.section-content.open{display:block}

/* Cat progress */
.cat-progress{height:4px;border-radius:2px;background:var(--surface);margin:8px 20px 0;overflow:hidden}
.cat-progress-fill{height:100%;border-radius:2px;transition:width 0.4s ease}

/* Table */
.check-table{width:100%;border-collapse:separate;border-spacing:0;table-layout:fixed}
.check-table thead th{padding:8px 14px;font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:0.8px;
  color:var(--text-faint);border-bottom:1px solid var(--border);text-align:left;background:var(--surface);position:sticky;top:0}
.check-table thead th:nth-child(1){width:70px} /* ID */
.check-table thead th:nth-child(2){width:auto} /* Check */
.check-table thead th:nth-child(3){width:100px} /* Status */
.check-table thead th:nth-child(4){width:70px} /* Severity */
.check-table thead th:nth-child(5){width:60px} /* Weight */
.check-table thead th:nth-child(6){width:80px} /* Origin */
.check-table thead th:nth-child(7){width:75px} /* Source */
.check-table thead th:nth-child(8){width:auto} /* Details */
.check-table tbody tr{transition:background 0.1s}
.check-table tbody tr:hover{background:var(--card-hover)}
.check-table td{padding:10px 14px;font-size:12px;border-bottom:1px solid var(--border);vertical-align:top;word-wrap:break-word;overflow-wrap:break-word}

/* Status badges */
.status-badge{display:inline-flex;align-items:center;gap:4px;padding:2px 9px;border-radius:20px;font-size:10px;font-weight:700;letter-spacing:0.3px;text-transform:uppercase}
.status-badge.pass{background:var(--green-dim);color:var(--green-text)}
.status-badge.warning{background:var(--orange-dim);color:var(--orange-text)}
.status-badge.fail{background:var(--red-dim);color:var(--red-text)}
.status-badge.na{background:var(--surface);color:var(--text-faint)}
.status-badge.notassessed{background:var(--surface);color:var(--text-faint)}

/* Severity */
.sev{font-size:10px;font-weight:600}
.sev-critical{color:var(--red-text)} .sev-high{color:var(--orange-text)} .sev-medium{color:var(--text-secondary)} .sev-low{color:var(--text-faint)}

/* Weight dots */
.wt{font-size:8px;letter-spacing:1px;color:var(--text-dim);white-space:nowrap}

/* Origin badges */
.origin-badge{display:inline-block;padding:1px 6px;border-radius:3px;font-size:9px;font-weight:600;margin-right:2px}
.origin-badge.waf{background:var(--accent-dim);color:var(--accent-text)}
.origin-badge.caf{background:var(--green-dim);color:var(--green-text)}
.origin-badge.lza{background:var(--orange-dim);color:var(--orange-text)}
.origin-badge.sec{background:var(--red-dim);color:var(--red-text)}
.origin-badge.fsl{background:rgba(123,31,162,0.12);color:#ce93d8}
.origin-badge.avd{background:var(--surface);color:var(--text-dim)}

/* Excluded row */
.excluded-row{opacity:0.4}
.excluded-tag{font-style:italic;font-size:9px;color:var(--text-faint);text-decoration:none}

/* Notes section */
.notes-section{background:var(--card);border:1px solid var(--card-border);border-radius:var(--radius);padding:20px;margin-top:28px;white-space:pre-wrap;font-size:12px;color:var(--text-secondary)}

/* Footer */
.footer{margin-top:40px;padding-top:20px;border-top:1px solid var(--border);display:flex;justify-content:space-between;align-items:center;color:var(--text-faint);font-size:11px}
.footer a{color:var(--accent-text);text-decoration:none}
.footer a:hover{text-decoration:underline}
.footer-dot{width:6px;height:6px;border-radius:50%;background:var(--accent);display:inline-block;margin-right:8px}

@media(max-width:900px){.summary{grid-template-columns:repeat(2,1fr)}.container{padding:16px}.toolbar{padding:10px 16px;flex-wrap:wrap}.maturity-header{flex-direction:column;align-items:stretch;text-align:center}.maturity-dims{grid-template-columns:repeat(2,1fr)}}
@media(max-width:600px){.maturity-dims{grid-template-columns:1fr}}
@media print{.toolbar{position:static}.btn{display:none}body{background:#fff;color:#000}.stat-card,.section-header,.check-table th{background:#f5f5f5;border-color:#ddd}.stat-card .num,.section-header{color:#000}td{border-color:#ddd}.section-content{display:block!important}}

/* Maturity Section */
.maturity-section{margin-bottom:36px}
.maturity-title{font-size:18px;font-weight:700;color:var(--text-bright);margin-bottom:4px}
.maturity-subtitle{font-size:12px;color:var(--text-dim);margin-bottom:20px}
.maturity-header{display:flex;align-items:center;gap:32px;background:var(--card);border:1px solid var(--card-border);border-radius:var(--radius);padding:24px 32px;margin-bottom:16px}
.maturity-radar{flex-shrink:0}
.maturity-radar svg{display:block}
.maturity-composite{display:flex;flex-direction:column;align-items:center;gap:4px;flex-shrink:0}
.maturity-composite .big-score{font-size:42px;font-weight:800;letter-spacing:-2px;line-height:1}
.maturity-composite .level-label{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:1px;color:var(--text-dim);margin-top:2px}
.maturity-zone-strip{flex:1;display:flex;flex-direction:column;gap:6px;min-width:0}
.maturity-zone-bar{display:flex;height:10px;border-radius:5px;overflow:hidden}
.maturity-zone-bar .zone{transition:flex 0.3s ease}
.maturity-zone-labels{display:flex;font-size:9px;color:var(--text-faint)}
.maturity-zone-labels span{text-align:center}
.maturity-zone-marker{position:relative;height:0}
.maturity-zone-marker .pin{position:absolute;top:-18px;width:2px;height:14px;background:var(--text-bright);border-radius:1px}
.maturity-dims{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}
.dim-card{background:var(--card);border:1px solid var(--card-border);border-radius:var(--radius-sm);padding:16px 20px;transition:border-color 0.2s}
.dim-card:hover{border-color:var(--border-strong)}
.dim-card .dim-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:8px}
.dim-card .dim-label{font-size:12px;font-weight:600;color:var(--text)}
.dim-card .dim-score{font-size:20px;font-weight:800;letter-spacing:-1px}
.dim-card .dim-bar{height:4px;border-radius:2px;background:var(--surface);overflow:hidden}
.dim-card .dim-bar-fill{height:100%;border-radius:2px;transition:width 0.4s ease}
.dim-card .dim-meta{display:flex;gap:10px;margin-top:6px;font-size:10px;color:var(--text-faint)}

/* Field Notes */
.field-note{background:var(--card);border:1px solid var(--card-border);border-left:3px solid var(--accent);border-radius:0 var(--radius-sm) var(--radius-sm) 0;padding:16px 20px;margin:8px 20px 16px}
.field-note-title{font-size:12px;font-weight:700;color:var(--accent-text);margin-bottom:6px;display:flex;align-items:center;gap:6px}
.field-note-title::before{content:'💡';font-size:14px}
.field-note p{font-size:12px;color:var(--text-secondary);line-height:1.6;margin:4px 0}
.field-note a{color:var(--accent-text);text-decoration:none;border-bottom:1px dotted var(--accent-text)}
.field-note a:hover{text-decoration:underline}
.field-note ul{margin:6px 0 6px 16px;font-size:11px;color:var(--text-secondary)}
.field-note ul li{margin:3px 0}

/* Effort Matrix */
.effort-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:32px}
.effort-card{background:var(--card);border:1px solid var(--card-border);border-radius:var(--radius);padding:18px;text-align:center}
.effort-card .effort-icon{font-size:24px;margin-bottom:4px}
.effort-card .effort-count{font-size:28px;font-weight:800;letter-spacing:-1px}
.effort-card .effort-label{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:0.5px;color:var(--text-dim);margin-top:2px}
.effort-card.quick .effort-count{color:var(--green-text)} .effort-card.some .effort-count{color:var(--orange-text)} .effort-card.major .effort-count{color:var(--red-text)}

/* Impact & Remediation callouts */
.impact-box{margin-top:6px;padding:5px 8px;background:var(--accent-dim);border-left:3px solid var(--orange-text);border-radius:var(--radius-xs);font-size:11px;color:var(--text-secondary)}
.impact-box strong{color:var(--orange-text)}
.remediation-box{margin-top:4px;padding:5px 8px;background:var(--card);border-left:3px solid var(--green-text);border-radius:var(--radius-xs);font-family:'Cascadia Mono','Consolas',monospace;font-size:10px;color:var(--text-dim);overflow-x:auto;white-space:pre-wrap;word-break:break-all}
.remediation-box strong{color:var(--green-text);font-family:inherit}

/* Category data summary */
.cat-summary{padding:10px 14px;margin-bottom:12px;background:var(--card);border:1px solid var(--card-border);border-radius:var(--radius);font-size:12px;color:var(--text-secondary);line-height:1.6}
.cat-summary strong{color:var(--text-primary)}

/* Comparison table */
.comparison-table{width:100%;border-collapse:collapse;margin:16px 0;font-size:12px}
.comparison-table th{background:var(--card);border:1px solid var(--card-border);padding:8px 10px;text-align:left;font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:0.3px}
.comparison-table td{border:1px solid var(--card-border);padding:6px 10px}
.comparison-table tr:nth-child(even){background:var(--card)}
</style>
</head>
<body>
<div class="toolbar">
  <h1><span class="footer-dot"></span>AVD Assessment Report</h1>
  <div class="spacer"></div>
  <button class="btn" onclick="var t=document.documentElement.getAttribute('data-theme');document.documentElement.setAttribute('data-theme',t==='dark'?'light':'dark')">Theme</button>
  <button class="btn" onclick="var s=document.getElementsByClassName('section-content');for(var i=0;i<s.length;i++){s[i].className='section-content open'};var h=document.getElementsByClassName('section-header');for(var j=0;j<h.length;j++){h[j].className='section-header open'}">Expand</button>
  <button class="btn" onclick="var s=document.getElementsByClassName('section-content open');while(s.length>0){s[0].className='section-content'};var h=document.getElementsByClassName('section-header open');while(h.length>0){h[0].className='section-header'}">Collapse</button>
  <button class="btn" onclick="window.print()">Print</button>
</div>
<div class="container">

<!-- Meta -->
<div class="meta-grid">
  <div class="meta-item">Customer <span>$(& $enc $Global:Assessment.CustomerName)</span></div>
  <div class="meta-item">Assessor <span>$(& $enc $Global:Assessment.AssessorName)</span></div>
  <div class="meta-item">Date <span>$($Global:Assessment.Date)</span></div>
  <div class="meta-item">Tool <span>AVD Assessor v$Global:AppVersion</span></div>
</div>

<!-- Score Hero -->
<div class="score-hero">
  <div class="score-value $OverallClass">$(if ($Overall -ge 0) { "$Overall%" } else { '—' })</div>
  <div class="score-label">Overall Readiness Score</div>
</div>

<!-- Summary cards -->
<div class="summary">
  <div class="stat-card pass"><div class="accent-bar"></div><div class="num">$PassCount</div><div class="label">Pass</div></div>
  <div class="stat-card warn"><div class="accent-bar"></div><div class="num">$WarnCount</div><div class="label">Warning</div></div>
  <div class="stat-card fail"><div class="accent-bar"></div><div class="num">$FailCount</div><div class="label">Fail</div></div>
  <div class="stat-card na"><div class="accent-bar"></div><div class="num">$NACount</div><div class="label">N/A</div></div>
  <div class="stat-card excluded"><div class="accent-bar"></div><div class="num">$ExclCount</div><div class="label">Excluded</div></div>
  <div class="stat-card total"><div class="accent-bar"></div><div class="num">$($Global:Assessment.Checks.Count)</div><div class="label">Total</div></div>
</div>

<!-- Progress bar -->
"@)

    $Total = [math]::Max(1, $PassCount + $NACount + $WarnCount + $FailCount)
    $PctPass = [math]::Round(($PassCount + $NACount) / $Total * 100, 1)
    $PctWarn = [math]::Round($WarnCount / $Total * 100, 1)
    $PctFail = [math]::Round($FailCount / $Total * 100, 1)
    [void]$html.Append("<div class='progress-bar'><div class='seg seg-pass' style='width:${PctPass}%'></div><div class='seg seg-warn' style='width:${PctWarn}%'></div><div class='seg seg-fail' style='width:${PctFail}%'></div></div>`n")

    # ─── MATURITY RADAR ────────────────────────────────────────────────
    $Composite = Get-CompositeMaturityScore
    $MatLevel  = Get-MaturityLevel $Composite
    $CompClass = if ($Composite -ge 80) { 'green' } elseif ($Composite -ge 50) { 'orange' } else { 'red' }
    $DimKeys   = @($Global:MaturityDimensions.Keys)
    $DimScores = @{}
    foreach ($K in $DimKeys) { $DimScores[$K] = Get-DimensionScore $K }

    # Build SVG radar (hexagonal)
    $CX = 150; $CY = 150; $MaxR = 95
    $N = $DimKeys.Count
    $SvgPolys = [System.Text.StringBuilder]::new()

    # Grid rings at 25%, 50%, 75%, 100%
    foreach ($Ring in @(0.25, 0.50, 0.75, 1.0)) {
        $RingPts = @()
        for ($i = 0; $i -lt $N; $i++) {
            $Angle = [math]::PI * 2 * $i / $N - [math]::PI / 2
            $R = $MaxR * $Ring
            $RingPts += "$([math]::Round($CX + $R * [math]::Cos($Angle),1)),$([math]::Round($CY + $R * [math]::Sin($Angle),1))"
        }
        [void]$SvgPolys.Append("<polygon points='$($RingPts -join ' ')' fill='none' stroke='var(--border)' stroke-width='0.5'/>`n")
    }

    # Axis lines
    for ($i = 0; $i -lt $N; $i++) {
        $Angle = [math]::PI * 2 * $i / $N - [math]::PI / 2
        $EX = [math]::Round($CX + $MaxR * [math]::Cos($Angle), 1)
        $EY = [math]::Round($CY + $MaxR * [math]::Sin($Angle), 1)
        [void]$SvgPolys.Append("<line x1='$CX' y1='$CY' x2='$EX' y2='$EY' stroke='var(--border)' stroke-width='0.5'/>`n")
    }

    # Data polygon
    $DataPts = @()
    for ($i = 0; $i -lt $N; $i++) {
        $Angle = [math]::PI * 2 * $i / $N - [math]::PI / 2
        $Score = [math]::Max(0, $DimScores[$DimKeys[$i]])
        $R = $MaxR * $Score / 100
        $DataPts += "$([math]::Round($CX + $R * [math]::Cos($Angle),1)),$([math]::Round($CY + $R * [math]::Sin($Angle),1))"
    }
    [void]$SvgPolys.Append("<polygon points='$($DataPts -join ' ')' fill='var(--accent-dim)' stroke='var(--accent)' stroke-width='2' opacity='0.7'/>`n")

    # Labels
    for ($i = 0; $i -lt $N; $i++) {
        $Angle = [math]::PI * 2 * $i / $N - [math]::PI / 2
        $LR = $MaxR + 20
        $LX = [math]::Round($CX + $LR * [math]::Cos($Angle), 1)
        $LY = [math]::Round($CY + $LR * [math]::Sin($Angle), 1)
        $Anchor = if ([math]::Abs($LX - $CX) -lt 5) { 'middle' } elseif ($LX -gt $CX) { 'start' } else { 'end' }
        $DimLabel = $Global:MaturityDimensions[$DimKeys[$i]].Label -replace ' & .*',''
        $DimSc = if ($DimScores[$DimKeys[$i]] -ge 0) { "$($DimScores[$DimKeys[$i]])%" } else { '—' }
        [void]$SvgPolys.Append("<text x='$LX' y='$LY' text-anchor='$Anchor' dominant-baseline='middle' font-size='9' font-weight='600' fill='var(--text-secondary)'>$DimLabel $DimSc</text>`n")
    }

    [void]$html.Append(@"
<!-- Maturity Section -->
<div class="maturity-section">
<div class="maturity-title">Environment Maturity</div>
<div class="maturity-subtitle">Composite score across 6 dimensions — aligned to WAF, CAF, LZA, and Security best practices</div>
<div class="maturity-header">
  <div class="maturity-radar">
    <svg width="240" height="240" viewBox="0 0 300 300">
    $($SvgPolys.ToString())
    </svg>
  </div>
  <div class="maturity-composite">
    <div class="big-score $CompClass">$(if ($Composite -ge 0) { "$Composite%" } else { '—' })</div>
    <div class="level-label">$MatLevel</div>
  </div>
  <div class="maturity-zone-strip">
    <div class="maturity-zone-bar">
      <div class="zone" style="flex:35;background:var(--red);opacity:0.25"></div>
      <div class="zone" style="flex:20;background:var(--orange);opacity:0.3"></div>
      <div class="zone" style="flex:20;background:#EAB308;opacity:0.3"></div>
      <div class="zone" style="flex:15;background:#14B8A6;opacity:0.3"></div>
      <div class="zone" style="flex:11;background:var(--green);opacity:0.3"></div>
    </div>
    <div class="maturity-zone-marker" style="width:100%">
      <div class="pin" style="left:clamp(0%,$(if ($Composite -ge 0) { $Composite } else { 0 })%,100%)"></div>
    </div>
    <div class="maturity-zone-labels">
      <span style="flex:35">Initial</span>
      <span style="flex:20">Developing</span>
      <span style="flex:20">Defined</span>
      <span style="flex:15">Managed</span>
      <span style="flex:11">Optimized</span>
    </div>
  </div>
</div>
<div class="maturity-dims">
"@)

    foreach ($K in $DimKeys) {
        $DS = $DimScores[$K]
        $DLabel = $Global:MaturityDimensions[$K].Label
        $DColor = if ($DS -ge 80) { 'var(--green)' } elseif ($DS -ge 50) { 'var(--orange)' } elseif ($DS -ge 0) { 'var(--red)' } else { 'var(--text-faint)' }
        $DWidth = if ($DS -ge 0) { $DS } else { 0 }
        # Count pass/warn/fail in dimension
        $DimPrefixes = $Global:MaturityDimensions[$K].Prefixes
        $DChecks = @($Global:Assessment.Checks | Where-Object {
            $Id = $_.Id; -not $_.Excluded -and
            ($DimPrefixes | Where-Object { $Id -like "$_*" })
        })
        $DPass = @($DChecks | Where-Object { $_.Status -eq 'Pass' -or $_.Status -eq 'N/A' }).Count
        $DWarn = @($DChecks | Where-Object Status -eq 'Warning').Count
        $DFail = @($DChecks | Where-Object Status -eq 'Fail').Count

        [void]$html.Append(@"
<div class="dim-card">
  <div class="dim-header">
    <span class="dim-label">$(& $enc $DLabel)</span>
    <span class="dim-score" style="color:$DColor">$(if ($DS -ge 0) { "$DS%" } else { '—' })</span>
  </div>
  <div class="dim-bar"><div class="dim-bar-fill" style="width:${DWidth}%;background:$DColor"></div></div>
  <div class="dim-meta"><span style="color:var(--green-text)">$DPass pass</span><span style="color:var(--orange-text)">$DWarn warn</span><span style="color:var(--red-text)">$DFail fail</span></div>
</div>
"@)
    }

    [void]$html.Append("</div></div>`n")

    # ─── EFFORT BREAKDOWN ──────────────────────────────────────────────
    $EffortQuick = @($Global:Assessment.Checks | Where-Object { $_.Effort -eq 'Quick Win' -and $_.Status -in @('Fail','Warning') }).Count
    $EffortSome  = @($Global:Assessment.Checks | Where-Object { $_.Effort -eq 'Some Effort' -and $_.Status -in @('Fail','Warning') }).Count
    $EffortMajor = @($Global:Assessment.Checks | Where-Object { $_.Effort -eq 'Major Effort' -and $_.Status -in @('Fail','Warning') }).Count
    if ($EffortQuick + $EffortSome + $EffortMajor -gt 0) {
        [void]$html.Append(@"
<div style="margin-bottom:32px">
<div style="font-size:16px;font-weight:700;color:var(--text-bright);margin-bottom:4px">Priority Matrix</div>
<div style="font-size:12px;color:var(--text-dim);margin-bottom:14px">Findings that need attention, grouped by remediation effort</div>
<div class="effort-grid">
  <div class="effort-card quick"><div class="effort-icon">&#9889;</div><div class="effort-count">$EffortQuick</div><div class="effort-label">Quick Wins</div></div>
  <div class="effort-card some"><div class="effort-icon">&#128736;</div><div class="effort-count">$EffortSome</div><div class="effort-label">Some Effort</div></div>
  <div class="effort-card major"><div class="effort-icon">&#127959;</div><div class="effort-count">$EffortMajor</div><div class="effort-label">Major Effort</div></div>
</div>
</div>
"@)
    }

    # Category sections
    $Categories = Get-Categories
    foreach ($Cat in $Categories) {
        $Score = Get-CategoryScore $Cat
        $CatChecks = @($Global:Assessment.Checks | Where-Object Category -eq $Cat)
        $CatAssessed = @($CatChecks | Where-Object { $_.Status -ne 'Not Assessed' -and -not $_.Excluded })
        $CatPass = @($CatAssessed | Where-Object { $_.Status -eq 'Pass' -or $_.Status -eq 'N/A' }).Count
        $ScoreColor = if ($Score -ge 80) { 'var(--green)' } elseif ($Score -ge 50) { 'var(--orange)' } else { 'var(--red)' }
        $DotColor = if ($Score -ge 80) { 'background:var(--green)' } elseif ($Score -ge 50) { 'background:var(--orange)' } elseif ($Score -ge 0) { 'background:var(--red)' } else { 'background:var(--text-faint)' }
        $BarWidth = if ($Score -ge 0) { $Score } else { 0 }

        [void]$html.Append(@"
<div class="section-header" onclick="this.className=this.className.indexOf('open')>=0?'section-header':'section-header open';var c=this.nextElementSibling;c.className=c.className.indexOf('open')>=0?'section-content':'section-content open'">
  <span class="chevron">&#9658;</span>
  <span class="section-dot" style="$DotColor"></span>
  $(& $enc $Cat)
  <span class="section-score" style="color:$ScoreColor">$(if ($Score -ge 0) { "$Score%" } else { '—' })</span>
  <span class="section-count">$CatPass/$($CatAssessed.Count) pass</span>
</div>
<div class="section-content">
<div class="cat-progress"><div class="cat-progress-fill" style="width:${BarWidth}%;background:$ScoreColor"></div></div>
"@)
        # ─── CATEGORY DATA SUMMARY ────────────────────────────────────
        $CatSummary = $null
        $Inv = $Global:Assessment.Discovery.Inventory
        if ($Inv) {
            switch -Wildcard ($Cat) {
                'Session Host*' {
                    $SHCount = @($Inv.SessionHosts).Count
                    $HPCount = @($Inv.HostPools).Count
                    if ($SHCount -gt 0) {
                        $Pooled = @($Inv.HostPools | Where-Object { $_.HostPoolType -eq 'Pooled' }).Count
                        $Personal = $HPCount - $Pooled
                        $SKUs = @($Inv.SessionHosts | Group-Object VMSize | Sort-Object Count -Descending | ForEach-Object { "$($_.Name) ($($_.Count))" }) -join ', '
                        $Unhealthy = @($Inv.SessionHosts | Where-Object { $_.Status -eq 'NeedsAssistance' }).Count
                        $CatSummary = "Found <strong>$SHCount session host(s)</strong> across <strong>$HPCount host pool(s)</strong> ($Pooled pooled, $Personal personal). VM sizes: $SKUs.$(if ($Unhealthy -gt 0) { " <strong>$Unhealthy host(s)</strong> report NeedsAssistance status." })"
                    }
                }
                'Network*' {
                    $VNetCount = @($Inv.VNets).Count
                    $NSGCount = @($Inv.NSGs).Count
                    if ($VNetCount -gt 0) {
                        $SubnetCount = @($Inv.VNets | ForEach-Object { $_.Subnets } | Where-Object { $_ }).Count
                        $CatSummary = "Discovered <strong>$VNetCount VNet(s)</strong> with <strong>$SubnetCount subnet(s)</strong> and <strong>$NSGCount NSG(s)</strong>."
                        $Watchers = @($Inv.NetworkWatchers).Count
                        if ($Watchers -gt 0) { $CatSummary += " Network Watcher in <strong>$Watchers region(s)</strong>." }
                        $DnsZones = @($Inv.PrivateDnsZones).Count
                        if ($DnsZones -gt 0) { $CatSummary += " <strong>$DnsZones private DNS zone(s)</strong> configured." }
                        $FwCount = @($Inv.Firewalls).Count
                        if ($FwCount -gt 0) { $CatSummary += " <strong>$FwCount Azure Firewall(s)</strong>." }
                        $GwCount = @($Inv.VPNGateways).Count
                        if ($GwCount -gt 0) {
                            $GwTypes = @($Inv.VPNGateways | ForEach-Object { $_.GatewayType } | Sort-Object -Unique) -join '/'
                            $CatSummary += " <strong>$GwCount $GwTypes gateway(s)</strong>."
                        }
                    }
                }
                'Security*' {
                    $TLCount = @($Inv.SessionHosts | Where-Object { $_.SecurityType -eq 'TrustedLaunch' }).Count
                    $TotalSH = @($Inv.SessionHosts).Count
                    $KVCount = @($Inv.KeyVaults).Count
                    if ($TotalSH -gt 0) {
                        $CatSummary = "<strong>$TLCount/$TotalSH</strong> session hosts use Trusted Launch."
                        if ($KVCount -gt 0) {
                            $KVWithPE = @($Inv.KeyVaults | Where-Object { $_.HasPrivateEndpoint }).Count
                            $CatSummary += " <strong>$KVCount Key Vault(s)</strong> found ($KVWithPE with private endpoints)."
                        }
                    }
                }
                'FSLogix*' {
                    $SACount = @($Inv.StorageAccounts).Count
                    if ($SACount -gt 0) {
                        $PECount = @($Inv.StorageAccounts | Where-Object { $_.HasPrivateEndpoint }).Count
                        $TierSummary = @($Inv.StorageAccounts | Group-Object Tier | ForEach-Object { "$($_.Name) ($($_.Count))" }) -join ', '
                        $CatSummary = "Found <strong>$SACount storage account(s)</strong>: $TierSummary. <strong>$PECount/$SACount</strong> have private endpoints."
                    }
                }
                'Governance*' {
                    $SPCount = @($Inv.ScalingPlans).Count
                    $OrphDisks = @($Inv.OrphanedDisks).Count
                    $OrphNICs = @($Inv.OrphanedNICs).Count
                    $PolicyCount = @($Inv.PolicyAssignments).Count
                    $BudgetCount = @($Inv.Budgets).Count
                    $RICount = @($Inv.Reservations).Count
                    $QuotaHigh = @($Inv.Quotas | Where-Object { $_.UsagePct -ge 80 }).Count
                    $parts = @()
                    if ($SPCount -gt 0) { $parts += "<strong>$SPCount scaling plan(s)</strong>" }
                    if ($PolicyCount -gt 0) { $parts += "<strong>$PolicyCount policy assignment(s)</strong>" }
                    if ($BudgetCount -gt 0) { $parts += "<strong>$BudgetCount cost budget(s)</strong>" }
                    if ($RICount -gt 0) { $parts += "<strong>$RICount VM reservation(s)</strong>" }
                    if ($OrphDisks -gt 0) { $parts += "<strong>$OrphDisks orphaned disk(s)</strong>" }
                    if ($OrphNICs -gt 0) { $parts += "<strong>$OrphNICs orphaned NIC(s)</strong>" }
                    if ($QuotaHigh -gt 0) { $parts += "<strong>$QuotaHigh quota(s) &gt;80%</strong>" }
                    if ($parts.Count -gt 0) { $CatSummary = ($parts -join ', ') + ' detected.' }
                }
                'Monitor*' {
                    $AlertCount = @($Inv.AlertRules).Count
                    $DiagCount = @($Global:Assessment.Checks | Where-Object { $_.Id -match '^MON-DIAG' -and $_.Status -eq 'Pass' }).Count
                    $DiagTotal = @($Global:Assessment.Checks | Where-Object { $_.Id -match '^MON-DIAG' }).Count
                    $parts = @()
                    if ($DiagTotal -gt 0) { $parts += "<strong>$DiagCount/$DiagTotal</strong> resources have diagnostic settings" }
                    if ($AlertCount -gt 0) { $parts += "<strong>$AlertCount alert rule(s)</strong> configured" }
                    if ($parts.Count -gt 0) { $CatSummary = ($parts -join '. ') + '.' }
                }
                'BCDR*' {
                    $SPCount = @($Inv.ScalingPlans).Count
                    $MultiRegion = @($Inv.HostPools | ForEach-Object { $_.Location } | Sort-Object -Unique).Count
                    $CatSummary = "Host pools in <strong>$MultiRegion region(s)</strong>."
                    if ($SPCount -gt 0) { $CatSummary += " <strong>$SPCount scaling plan(s)</strong> with schedule coverage." }
                }
            }
        }
        if ($CatSummary) {
            [void]$html.Append("<div class='cat-summary'>$CatSummary</div>`n")
        }

        [void]$html.Append(@"
<table class="check-table">
<thead><tr><th>ID</th><th>Check</th><th>Status</th><th>Sev</th><th>Wt</th><th>Origin</th><th>Source</th><th>Details</th><th>Reference</th></tr></thead>
<tbody>
"@)
        foreach ($Check in $CatChecks) {
            $StatusClass = switch ($Check.Status) { 'Pass' { 'pass' } 'Warning' { 'warning' } 'Fail' { 'fail' } 'N/A' { 'na' } default { 'notassessed' } }
            $SevClass = "sev-$($Check.Severity.ToLower())"
            $RowClass = if ($Check.Excluded) { " class='excluded-row'" } else { '' }
            $ExclTag = if ($Check.Excluded) { ' <span class="excluded-tag">[excluded]</span>' } else { '' }
            $WVal = [math]::Max(1,[math]::Min(5,[int]$Check.Weight))
            $WDots = ('&#9679;' * $WVal) + ('&#9675;' * (5 - $WVal))
            $OriginBadges = if ($Check.Origin) {
                ($Check.Origin -split ',' | ForEach-Object {
                    $t = $_.Trim().ToLower()
                    "<span class='origin-badge $t'>$($_.Trim().ToUpper())</span>"
                }) -join ''
            } else { '' }

            [void]$html.Append("<tr$RowClass>")
            [void]$html.Append("<td>$($Check.Id)</td>")
            [void]$html.Append("<td><strong>$(& $enc $Check.Name)</strong>$ExclTag<br><small style='color:var(--text-dim)'>$(& $enc $Check.Description)</small></td>")
            [void]$html.Append("<td><span class='status-badge $StatusClass'>$($Check.Status)</span></td>")
            [void]$html.Append("<td><span class='sev $SevClass'>$($Check.Severity)</span></td>")
            [void]$html.Append("<td><span class='wt' title='Weight $WVal/5'>$WDots</span></td>")
            [void]$html.Append("<td>$OriginBadges</td>")
            [void]$html.Append("<td>$($Check.Source)</td>")
            [void]$html.Append("<td>$(& $enc $Check.Details)$(if($Check.Notes){"`n<div style='margin-top:6px;padding:4px 8px;background:var(--accent-dim);border-left:2px solid var(--accent);border-radius:var(--radius-xs);font-size:11px;color:var(--accent-text)'><strong>Note:</strong> $(& $enc $Check.Notes)</div>"})$(if($Check.Impact -and $Check.Status -in @('Fail','Warning')){"`n<div class='impact-box'><strong>Impact:</strong> $(& $enc $Check.Impact)</div>"})$(if($Check.Remediation -and $Check.Status -in @('Fail','Warning')){"`n<div class='remediation-box'><strong>Fix:</strong> $(& $enc $Check.Remediation)</div>"})</td>")
            $RefHtml = if ($Check.Reference) { "<a href='$(& $enc $Check.Reference)' target='_blank' style='color:var(--accent);font-size:11px;word-break:break-all'>$(& $enc $Check.Reference)</a>" } else { '' }
            [void]$html.Append("<td>$RefHtml</td>")
            [void]$html.Append("</tr>`n")
        }
        [void]$html.Append("</tbody></table>`n")

        # ─── COMPARISON TABLES (resource inventory per category) ──────
        if ($Inv) {
            switch -Wildcard ($Cat) {
                'Session Host*' {
                    $SHs = @($Inv.SessionHosts)
                    if ($SHs.Count -gt 0) {
                        $SKUGroups = @($SHs | Group-Object VMSize | Sort-Object Count -Descending)
                        [void]$html.Append("<div style='margin-top:12px;font-size:13px;font-weight:600;color:var(--text-bright)'>VM SKU Distribution</div>`n")
                        [void]$html.Append("<table class='comparison-table'><thead><tr><th>VM Size</th><th>Count</th><th>% of Fleet</th></tr></thead><tbody>`n")
                        foreach ($G in $SKUGroups) {
                            $Pct = [math]::Round($G.Count / $SHs.Count * 100, 0)
                            [void]$html.Append("<tr><td>$($G.Name)</td><td>$($G.Count)</td><td>$Pct%</td></tr>`n")
                        }
                        [void]$html.Append("</tbody></table>`n")
                    }
                }
                'FSLogix*' {
                    $SAs = @($Inv.StorageAccounts)
                    if ($SAs.Count -gt 0) {
                        [void]$html.Append("<div style='margin-top:12px;font-size:13px;font-weight:600;color:var(--text-bright)'>Storage Account Comparison</div>`n")
                        [void]$html.Append("<table class='comparison-table'><thead><tr><th>Name</th><th>Tier</th><th>Replication</th><th>TLS</th><th>Private EP</th><th>Firewall</th></tr></thead><tbody>`n")
                        foreach ($SA in $SAs) {
                            $TlsBadge = if ($SA.Tls -match '1_2|1.2') { "<span style='color:var(--green-text)'>TLS 1.2</span>" } else { "<span style='color:var(--red-text)'>$($SA.Tls)</span>" }
                            $PEBadge = if ($SA.HasPrivateEndpoint) { "<span style='color:var(--green-text)'>Yes</span>" } else { "<span style='color:var(--orange-text)'>No</span>" }
                            $FwBadge = if ($SA.FirewallDefaultAction -eq 'Deny') { "<span style='color:var(--green-text)'>Deny</span>" } else { "<span style='color:var(--orange-text)'>Allow</span>" }
                            [void]$html.Append("<tr><td>$(& $enc $SA.Name)</td><td>$($SA.Tier)</td><td>$($SA.Replication)</td><td>$TlsBadge</td><td>$PEBadge</td><td>$FwBadge</td></tr>`n")
                        }
                        [void]$html.Append("</tbody></table>`n")
                    }
                }
                'Governance*' {
                    $OrphDisks = @($Inv.OrphanedDisks)
                    if ($OrphDisks.Count -gt 0) {
                        [void]$html.Append("<div style='margin-top:12px;font-size:13px;font-weight:600;color:var(--text-bright)'>Orphaned Disks</div>`n")
                        [void]$html.Append("<table class='comparison-table'><thead><tr><th>Name</th><th>Resource Group</th><th>Size (GB)</th><th>SKU</th></tr></thead><tbody>`n")
                        foreach ($D in $OrphDisks) {
                            [void]$html.Append("<tr><td>$(& $enc $D.Name)</td><td>$(& $enc $D.ResourceGroup)</td><td>$($D.SizeGB)</td><td>$($D.Sku)</td></tr>`n")
                        }
                        [void]$html.Append("</tbody></table>`n")
                    }
                    $Quotas = @($Inv.Quotas)
                    if ($Quotas.Count -gt 0) {
                        [void]$html.Append("<div style='margin-top:12px;font-size:13px;font-weight:600;color:var(--text-bright)'>VM Quota Usage</div>`n")
                        [void]$html.Append("<table class='comparison-table'><thead><tr><th>Region</th><th>Quota</th><th>Used</th><th>Limit</th><th>Usage %</th></tr></thead><tbody>`n")
                        foreach ($Q in ($Quotas | Sort-Object UsagePct -Descending)) {
                            $QColor = if ($Q.UsagePct -ge 80) { "color:var(--red-text)" } elseif ($Q.UsagePct -ge 60) { "color:var(--orange-text)" } else { "color:var(--green-text)" }
                            [void]$html.Append("<tr><td>$(& $enc $Q.Region)</td><td>$(& $enc $Q.Name)</td><td>$($Q.Current)</td><td>$($Q.Limit)</td><td><span style='$QColor;font-weight:600'>$($Q.UsagePct)%</span></td></tr>`n")
                        }
                        [void]$html.Append("</tbody></table>`n")
                    }
                }
            }
        }

        # ─── FIELD NOTES (expert knowledge per category) ──────────────
        $FieldNote = $null
        switch -Wildcard ($Cat) {
            'Security*' {
                $FieldNote = @"
<div class="field-note">
<div class="field-note-title">Field Notes — Security & Identity</div>
<p>RDP property security is one of the most overlooked areas in AVD deployments. The default RDP template allows <strong>all device redirections</strong> — drives, clipboard, printers, USB — creating data exfiltration vectors. Always parse <code>CustomRdpProperty</code> and explicitly restrict each channel.</p>
<ul>
<li><strong>Clipboard</strong> — Use <a href="https://learn.microsoft.com/en-us/azure/virtual-desktop/clipboard-transfer-direction-data-types">clipboard transfer direction policies</a> for granular control instead of binary enable/disable</li>
<li><strong>Screen Capture Protection</strong> — Blocks PrintScreen, Snipping Tool, AND screen sharing in Teams/Zoom. Level 2 blocks remote content only; Level 1 blocks local too</li>
<li><strong>Watermarking</strong> — QR codes contain session ID, enabling forensic tracing of photos taken of screens</li>
<li><strong>Trusted Launch</strong> — vTPM + Secure Boot prevents rootkits. Pair with <a href="https://learn.microsoft.com/en-us/azure/virtual-machines/trusted-launch#microsoft-defender-for-cloud-integration">Guest Attestation extension</a> for Defender for Cloud integrity monitoring</li>
<li><strong>MDE via VM Extension</strong> — If MDE.Windows extension is not found, the agent may still be deployed via Intune or GPO. The extension check is a proxy, not definitive</li>
<li><strong>Disk Encryption</strong> — Azure platform-managed keys (PMK) encrypt all managed disks by default, but ADE or encryption-at-host provides customer-controlled keys. For regulated workloads, use <a href="https://learn.microsoft.com/en-us/azure/virtual-machines/disk-encryption-overview">DES with customer-managed keys</a></li>
<li><strong>Key Vault</strong> — Store AVD deployment secrets, certificates, and encryption keys centrally. Enable soft delete + purge protection. Use RBAC (not access policies) for granular control. Always configure private endpoints — Key Vault without PE exposes secrets over the public internet</li>
<li><strong>Host Pool Private Link</strong> — AVD Private Link keeps control-plane traffic (session brokering, diagnostics) off the public internet. Requires private endpoints on host pools and workspaces with DNS configuration for <code>wvd.microsoft.com</code></li>
</ul>
</div>
"@
            }
            'Identity*' {
                $FieldNote = @"
<div class="field-note">
<div class="field-note-title">Field Notes — Identity & Access</div>
<p>SSO (<code>enablerdsaadauth:i:1</code>) is <strong>strongly recommended</strong> for Entra ID joined pools — without it, users face a double authentication prompt. For Hybrid-joined hosts, SSO requires specific <a href="https://learn.microsoft.com/en-us/azure/virtual-desktop/configure-single-sign-on">Kerberos configuration</a>.</p>
<ul>
<li><strong>Entra ID Join</strong> — Preferred for new deployments. Eliminates AD DS dependency and simplifies CA policies</li>
<li><strong>Conditional Access</strong> — Create separate policies for the "Windows 365 / Azure Virtual Desktop" cloud app with device compliance, MFA, and session controls</li>
<li><strong>RBAC least privilege</strong> — Use <em>Desktop Virtualization User</em> (not Contributor) for end users, <em>Desktop Virtualization Host Pool Contributor</em> for operators</li>
</ul>
</div>
"@
            }
            'Network*' {
                $FieldNote = @"
<div class="field-note">
<div class="field-note-title">Field Notes — Networking</div>
<p>AVD requires very specific outbound connectivity. If NSGs use <strong>deny-all outbound</strong>, you must explicitly allow the <a href="https://learn.microsoft.com/en-us/azure/virtual-desktop/required-fqdn-endpoint">required FQDNs and service tags</a>: <code>WindowsVirtualDesktop</code>, <code>AzureActiveDirectory</code>, <code>AzureResourceManager</code> on port 443.</p>
<ul>
<li><strong>Subnet capacity</strong> — Each session host needs 1 IP. Azure reserves 5 per subnet. A /24 gives 251 usable IPs. Plan for 30% headroom for scaling and maintenance replacements</li>
<li><strong>RDP Shortpath</strong> — UDP-based transport with STUN/TURN significantly reduces latency. Requires UDP 3478 outbound for public networks</li>
<li><strong>VNet peering</strong> — Disconnected peerings silently break connectivity. Monitor peering state in your alerting strategy</li>
<li><strong>NAT Gateway</strong> — Recommended over Azure default outbound for explicit, scalable SNAT. Critical when scaling beyond ~30 VMs per subnet</li>
<li><strong>Private DNS Zones</strong> — Essential for private endpoint resolution. <code>privatelink.file.core.windows.net</code> must be linked to AVD VNets or PE DNS resolution falls back to public IP, bypassing network isolation entirely</li>
<li><strong>Network Watcher</strong> — Enable in every AVD region. Provides packet capture, connection troubleshooting, and NSG flow logs for security audit trails</li>
<li><strong>Hub Firewall</strong> — Azure Firewall (or NVA) in the hub VNet provides centralized egress filtering with threat intelligence. Essential for deny-all-outbound with explicit allow rules using AVD service tags</li>
<li><strong>VPN/ExpressRoute</strong> — Required for hybrid identity (AD DS), on-premises file shares, and LOB apps. ExpressRoute provides private, low-latency connectivity; VPN is cost-effective for smaller deployments</li>
</ul>
</div>
"@
            }
            'Session*' {
                $FieldNote = @"
<div class="field-note">
<div class="field-note-title">Field Notes — Session Hosts</div>
<p>Session host sizing is the single biggest cost lever. D-series (general purpose) is the default recommendation, but measure actual workloads before right-sizing.</p>
<ul>
<li><strong>Ephemeral OS disks</strong> — Pooled hosts should use ephemeral disks when VM SKU supports it. Eliminates storage cost and provides faster reimage (~2 min vs ~10 min)</li>
<li><strong>B-series warning</strong> — Burstable VMs exhaust CPU credits under sustained load, causing severe performance degradation for pooled desktops</li>
<li><strong>Agent version</strong> — AVD agent auto-updates when the VM is running. VMs that are deallocated for weeks may fall behind and eventually fail registration</li>
<li><strong>Image freshness</strong> — Gallery image versions older than 90 days are a patching risk. Automate image builds with Azure Image Builder on a monthly cadence. Check <code>PublishedDate</code> on <code>Get-AzGalleryImageVersion</code> to verify currency</li>
<li><strong>Premium vs Standard SSD</strong> — For pooled hosts with ephemeral disks, the OS disk type is moot. For persistent disks, Standard SSD is usually sufficient — Premium only helps with high-IOPS workloads</li>
<li><strong>NeedsAssistance status</strong> — Typically means the agent crashed or the VM lost connectivity. Check the RDAgent and Geneva logs on the VM</li>
</ul>
</div>
"@
            }
            'FSLogix*' {
                $FieldNote = @"
<div class="field-note">
<div class="field-note-title">Field Notes — FSLogix & Profiles</div>
<p>Profile storage is often the hidden bottleneck in AVD deployments. Premium FileStorage with private endpoints is the gold standard for production.</p>
<ul>
<li><strong>SMB security</strong> — Disable SMB 2.1 and NTLMv2 on storage accounts. Use <code>AES-256-GCM</code> channel encryption and <code>AES-256</code> Kerberos ticket encryption</li>
<li><strong>Storage firewall</strong> — Set default action to <em>Deny</em> and allowlist via private endpoint or VNet rules. Never leave profile storage open to all networks</li>
<li><strong>ZRS vs GRS</strong> — ZRS protects against zone failure within a region (most common). GRS protects against full region failure but has higher cost and potential for <a href="https://learn.microsoft.com/en-us/fslogix/concepts-container-recovery-business-continuity">split-brain scenarios</a></li>
<li><strong>Soft delete</strong> — Always enable. Default 7-day retention. Saved us from accidental deletion dozens of times in production</li>
<li><strong>AV exclusions</strong> — FSLogix VHD/VHDX files MUST be excluded from antivirus scanning. Without this, profile mount times can jump from ~2s to 30s+</li>
</ul>
</div>
"@
            }
            'Monitor*' {
                $FieldNote = @"
<div class="field-note">
<div class="field-note-title">Field Notes — Monitoring & Telemetry</div>
<p>Without diagnostic settings on host pools, you lose visibility into connection quality, session errors, and scaling events. <a href="https://learn.microsoft.com/en-us/azure/virtual-desktop/insights">AVD Insights</a> is free but requires AMA on every session host.</p>
<ul>. Ensure all 6 recommended log categories are enabled: Checkpoint, Error, Management, Connection, HostRegistration, AgentHealthStatus
<li><strong>Diagnostic coverage</strong> — Enable diagnostics on host pools, workspaces, AND app groups. Missing any one creates blind spots in AVD Insights</li>
<li><strong>AMA vs MMA</strong> — Log Analytics agent (MMA) is deprecated. Migrate to <a href="https://learn.microsoft.com/en-us/azure/azure-monitor/agents/agents-overview">Azure Monitor Agent</a> with Data Collection Rules</li>
<li><strong>Defender for Cloud</strong> — Standard tier enables vulnerability assessment, adaptive application controls, and just-in-time VM access — all applicable to AVD</li>
<li><strong>Performance counters</strong> — For AVD Insights: CPU, Available Memory, Disk Queue Length, and Input Delay (under RemoteFX Graphics)</li>
</ul>
</div>
"@
            }
            'BCDR*' {
                $FieldNote = @"
<div class="field-note">
<div class="field-note-title">Field Notes — Resiliency & BCDR</div>
<p>AVD control plane is Microsoft-managed and geo-redundant. Your BCDR focus should be on session hosts, images, and profile storage.</p>
<ul>
<li><strong>Availability Zones</strong> — Spread session hosts across 3 zones. Use <code>FaultDomainMode</code> or zone-pinned deployments</li>
<li><strong>Scaling plans</strong> — Not just for cost — they're your capacity management tool. Without one, host pools can't auto-scale during outages</li>
<li><strong>Single-region risk</strong> — Acceptable for many workloads if documented. DR via secondary region is expensive — evaluate if RTO justifies cost</li>
<li><strong>Image replication</strong> — Use Compute Gallery with multi-region replication so DR region can spin up VMs without copying the image first</li>
</ul>
</div>
"@
            }
            'Governance*' {
                $FieldNote = @"
<div class="field-note">
<div class="field-note-title">Field Notes — Governance & Cost</div>
<p>AVD costs are primarily compute (VMs) and storage (disks + profiles). Scaling plans typically save 40-70% on compute for pooled host pools.</p>
<ul>
<li><strong>Tag strategy</strong> — At minimum: <code>Environment</code>, <code>Owner</code>, <code>CostCenter</code>, <code>Application</code>. Tags enable cost allocation and automated governance</li>
<li><strong>Start VM on Connect</strong> — Pairs with scaling plans. VMs start only when users connect, then auto-deallocate after scaling plan off-peak schedule</li>
<li><strong>Validation host pool</strong> — Receives AVD service updates before production. Catches breaking changes before they affect all users. Zero cost beyond the VMs</li>
<li><strong>Orphaned resources</strong> — Decommissioned session hosts often leave behind unattached disks and NICs. These incur storage costs and may contain cached credentials. Run regular cleanup sweeps</li>
<li><strong>Azure Policy</strong> — Deploy policy assignments on AVD resource groups to enforce guardrails: allowed VM SKUs, required tags, mandatory encryption, network restrictions. Use <a href="https://learn.microsoft.com/en-us/azure/cloud-adoption-framework/ready/landing-zone/design-area/resource-org-subscriptions">management group hierarchy</a> for inheritance</li>
<li><strong>Cost budgets</strong> — Create Azure budgets with alert thresholds at 80% and 100%. Without budget alerts, misconfigured scaling plans or orphaned resources can cause surprise bills. Budget alerts are free to configure</li>
<li><strong>Reserved Instances</strong> — For personal (always-on) host pools, evaluate 1-year or 3-year Reserved Instances to save 40-72% on compute. For pooled host pools with scaling plans, Savings Plans may be more flexible</li>
<li><strong>Quota headroom</strong> — Monitor vCPU quota usage per region. If quota is near limits, scaling plans silently fail to start new hosts. Request increases proactively before peak seasons</li>
<li><strong>Cross-resource tagging</strong> — Don't just tag host pools — tag VMs, storage, VNets, and NSGs consistently. Use Azure Policy <em>'Inherit a tag from the resource group'</em> for automatic propagation</li>
</ul>
</div>
"@
            }
            'Landing*' {
                $FieldNote = @"
<div class="field-note">
<div class="field-note-title">Field Notes — Landing Zone</div>
<p>The <a href="https://github.com/Azure/avdaccelerator">AVD Landing Zone Accelerator</a> provides Bicep/Terraform templates that encode Microsoft's opinionated best practices.</p>
<ul>
<li><strong>Application vs Platform LZ</strong> — AVD should be in an application landing zone, consuming shared services (DNS, firewall, ExpressRoute) from the platform</li>
<li><strong>Azure Policy</strong> — Enforce guardrails: require Trusted Launch, block public IPs on VMs, enforce NSGs — these are preventive, not detective controls</li>
</ul>
</div>
"@
            }
            'Application*' {
                $FieldNote = @"
<div class="field-note">
<div class="field-note-title">Field Notes — Application Delivery</div>
<p>App Attach enables dynamic app delivery without baking apps into the golden image — decouples app lifecycle from image lifecycle.</p>
<ul>
<li><strong>Desktop vs RemoteApp</strong> — RemoteApp provides individual apps in their own window. Better for task workers. Desktop gives full Windows experience</li>
<li><strong>Teams optimization</strong> — New Teams (Chromium-based) with <a href="https://learn.microsoft.com/en-us/azure/virtual-desktop/teams-on-avd">media optimization</a> is the recommended path. Classic Teams with WebRTC redirector is end-of-life</li>
</ul>
</div>
"@
            }
            'Operations*' {
                $FieldNote = @"
<div class="field-note">
<div class="field-note-title">Field Notes — Operations</div>
<p>Image management and agent currency are the most common operational gaps. Stale agents cause registration failures and connectivity issues.</p>
<ul>
<li><strong>Agent auto-update</strong> — Works only when VM is running. VMs deallocated >90 days may need manual re-registration</li>
<li><strong>Golden image pipeline</strong> — Automate via <a href="https://learn.microsoft.com/en-us/azure/virtual-machines/image-builder-overview">Azure Image Builder</a> + Compute Gallery. Monthly cadence for security updates</li>
</ul>
</div>
"@
            }
        }
        if ($FieldNote) {
            [void]$html.Append($FieldNote)
        }
        [void]$html.Append("</div>`n")
    }

    # Notes
    if ($Global:Assessment.Notes) {
        [void]$html.Append("<h2 style='font-size:16px;font-weight:600;margin:28px 0 12px;color:var(--text-bright)'>Notes</h2>`n")
        [void]$html.Append("<div class='notes-section'>$(& $enc $Global:Assessment.Notes)</div>`n")
    }

    # Footer
    [void]$html.Append(@"
<div class="footer">
  <div><span class="footer-dot"></span>Generated by AVD Assessor v$Global:AppVersion on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') &middot; Maturity: <strong>$MatLevel</strong> ($Composite%)</div>
  <div>
    <a href="https://learn.microsoft.com/en-us/azure/well-architected/azure-virtual-desktop/">WAF</a> &middot;
    <a href="https://learn.microsoft.com/en-us/azure/cloud-adoption-framework/scenarios/azure-virtual-desktop/">CAF</a> &middot;
    <a href="https://github.com/Azure/avdaccelerator">LZA</a> &middot;
    <a href="https://learn.microsoft.com/en-us/azure/architecture/guide/virtual-desktop/start-here">Architecture</a>
  </div>
</div>
</div>
<script>
var cols = document.getElementsByClassName('section-header');
for (var i = 0; i < cols.length; i++) {
    cols[i].onclick = function() {
        var cl = this.className;
        this.className = cl.indexOf('open') >= 0 ? 'section-header' : 'section-header open';
        var next = this.nextElementSibling;
        if (next) { next.className = next.className.indexOf('open') >= 0 ? 'section-content' : 'section-content open'; }
    };
}
</script>
</body></html>
"@)

    return $html.ToString()
}

# ===============================================================================
# SECTION 12b: EXECUTIVE SUMMARY (Clipboard)
# Generates plain-text and richly formatted RTF executive summaries from the
# current assessment data. The RTF version pastes beautifully into Outlook,
# Word, Teams, and OneNote. Both formats are placed on the clipboard via a
# DataObject so the receiving app picks the richest format it supports.
# ===============================================================================

function Get-ExecutiveSummary {
    <#
    .SYNOPSIS  Generates a plain-text executive summary from the current assessment.
    #>
    $Checks   = $Global:Assessment.Checks
    $Score    = Get-OverallScore
    $Maturity = Get-MaturityLevel $Score

    $Total    = $Checks.Count
    $Assessed = @($Checks | Where-Object { $_.Status -ne 'Not Assessed' }).Count
    $Pass     = @($Checks | Where-Object { $_.Status -eq 'Pass' }).Count
    $Fail     = @($Checks | Where-Object { $_.Status -eq 'Fail' }).Count
    $Warn     = @($Checks | Where-Object { $_.Status -eq 'Warning' }).Count
    $NA       = @($Checks | Where-Object { $_.Status -eq 'N/A' }).Count

    # Build definition lookup for effort/remediation (not always on check objects)
    $DefLookup = @{}
    foreach ($Def in $Global:CheckDefinitions) { $DefLookup[$Def.id] = $Def }

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine('AVD Assessor Executive Summary')
    [void]$sb.AppendLine('==============================')
    [void]$sb.AppendLine("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
    [void]$sb.AppendLine('')

    # ── Assessment Information ──
    [void]$sb.AppendLine('ASSESSMENT INFORMATION')
    [void]$sb.AppendLine('----------------------')
    [void]$sb.AppendLine("  Customer:      $($Global:Assessment.CustomerName)")
    [void]$sb.AppendLine("  Assessor:      $($Global:Assessment.AssessorName)")
    [void]$sb.AppendLine("  Date:          $($Global:Assessment.Date)")
    if ($Global:Assessment.Discovery) {
        $D = $Global:Assessment.Discovery
        if ($D.Subscriptions) {
            [void]$sb.AppendLine("  Subscriptions: $($D.Subscriptions.Count)")
            foreach ($Sub in $D.Subscriptions) {
                [void]$sb.AppendLine("                 $($Sub.Name) ($($Sub.Id))")
            }
        }
        if ($D.Inventory) {
            [void]$sb.AppendLine("  Host Pools:    $($D.Inventory.HostPools.Count)")
            [void]$sb.AppendLine("  Session Hosts: $($D.Inventory.SessionHosts.Count)")
        }
    }
    [void]$sb.AppendLine('')

    # ── Score Summary ──
    [void]$sb.AppendLine('SCORE SUMMARY')
    [void]$sb.AppendLine('-------------')
    [void]$sb.AppendLine("  Overall Score:    $(if ($Score -ge 0) { "$Score/100" } else { 'N/A' })")
    [void]$sb.AppendLine("  Maturity Level:   $Maturity")
    [void]$sb.AppendLine("  Total Checks:     $Total ($Assessed assessed)")
    [void]$sb.AppendLine("  Pass: $Pass  |  Fail: $Fail  |  Warning: $Warn  |  N/A: $NA")
    [void]$sb.AppendLine('')

    # ── Category Scores ──
    [void]$sb.AppendLine('CATEGORY SCORES')
    [void]$sb.AppendLine('---------------')
    foreach ($Cat in (Get-Categories)) {
        $CatScore  = Get-CategoryScore $Cat
        $CatChecks = @($Checks | Where-Object { $_.Category -eq $Cat })
        $CatFail   = @($CatChecks | Where-Object { $_.Status -eq 'Fail' }).Count
        $CatPass   = @($CatChecks | Where-Object { $_.Status -eq 'Pass' }).Count
        $ScoreStr  = if ($CatScore -ge 0) { "$CatScore%" } else { 'N/A' }
        [void]$sb.AppendLine("  $($Cat.PadRight(32)) $($ScoreStr.PadLeft(5))   ($CatPass pass, $CatFail fail)")
    }
    [void]$sb.AppendLine('')

    # ── Maturity Dimensions ──
    [void]$sb.AppendLine('MATURITY DIMENSIONS')
    [void]$sb.AppendLine('-------------------')
    foreach ($Key in $Global:MaturityDimensions.Keys) {
        $Dim   = $Global:MaturityDimensions[$Key]
        $DScore = Get-DimensionScore $Key
        $DStr   = if ($DScore -ge 0) { "$DScore%" } else { 'N/A' }
        [void]$sb.AppendLine("  $($Dim.Label.PadRight(28)) $($DStr.PadLeft(5))")
    }
    $Composite = Get-CompositeMaturityScore
    if ($Composite -ge 0) {
        [void]$sb.AppendLine("  $('Composite'.PadRight(28)) $("$Composite%".PadLeft(5))")
    }
    [void]$sb.AppendLine('')

    # ── Key Passes (Critical/High) ──
    $strongChecks = @($Checks | Where-Object { $_.Status -eq 'Pass' -and $_.Severity -in @('Critical','High') })
    if ($strongChecks.Count -gt 0) {
        [void]$sb.AppendLine('KEY PASSES (Critical/High checks met)')
        [void]$sb.AppendLine('--------------------------------------')
        $grouped = $strongChecks | Group-Object Category
        foreach ($g in ($grouped | Sort-Object Count -Descending)) {
            $items = @($g.Group | Select-Object -First 5 | ForEach-Object { $_.Name })
            [void]$sb.AppendLine("  $($g.Name) ($($g.Count)):")
            foreach ($item in $items) {
                [void]$sb.AppendLine("    + $item")
            }
            if ($g.Count -gt 5) { [void]$sb.AppendLine("    + ... and $($g.Count - 5) more") }
        }
        [void]$sb.AppendLine('')
    }

    # ── Critical & High Failures ──
    $critFails = @($Checks | Where-Object {
        $_.Status -eq 'Fail' -and $_.Severity -in @('Critical','High')
    } | Sort-Object @{E={if($_.Severity -eq 'Critical'){0}else{1}}}, Id)
    if ($critFails.Count -gt 0) {
        [void]$sb.AppendLine("CRITICAL & HIGH FAILURES ($($critFails.Count))")
        [void]$sb.AppendLine('---------------------------------------------')
        foreach ($f in $critFails) {
            [void]$sb.AppendLine("  [$($f.Severity.ToUpper().PadRight(8))] $($f.Id)  $($f.Name)")
            if ($f.Details) {
                $detailStr = if ($f.Details.Length -gt 80) { $f.Details.Substring(0,80) + '...' } else { $f.Details }
                [void]$sb.AppendLine("             $detailStr")
            }
        }
        [void]$sb.AppendLine('')
    }

    # ── Medium/Low Failures (count only) ──
    $medLowFails = @($Checks | Where-Object { $_.Status -eq 'Fail' -and $_.Severity -in @('Medium','Low') })
    if ($medLowFails.Count -gt 0) {
        $medCount = @($medLowFails | Where-Object { $_.Severity -eq 'Medium' }).Count
        $lowCount = @($medLowFails | Where-Object { $_.Severity -eq 'Low' }).Count
        [void]$sb.AppendLine("OTHER FAILURES: $medCount Medium, $lowCount Low")
        [void]$sb.AppendLine('')
    }

    # ── Quick Win Remediation Priority ──
    $quickWins = @($Checks | Where-Object {
        $_.Status -eq 'Fail' -and ($DefLookup[$_.Id].effort -eq 'Quick Win' -or $_.effort -eq 'Quick Win')
    } | Sort-Object @{E={switch($_.Severity){'Critical'{0}'High'{1}'Medium'{2}default{3}}}}, Id | Select-Object -First 10)
    if ($quickWins.Count -gt 0) {
        [void]$sb.AppendLine("QUICK WIN REMEDIATION (top $($quickWins.Count))")
        [void]$sb.AppendLine('----------------------------------------------')
        $qi = 0
        foreach ($qw in $quickWins) {
            $qi++
            [void]$sb.AppendLine("  $qi. [$($qw.Severity)] $($qw.Name) ($($qw.Id))")
            $rem = if ($qw.remediation) { $qw.remediation } elseif ($DefLookup[$qw.Id].remediation) { $DefLookup[$qw.Id].remediation } else { $null }
            if ($rem) {
                $remText = if ($rem.Length -gt 100) { $rem.Substring(0,100) + '...' } else { $rem }
                [void]$sb.AppendLine("     $remText")
            }
        }
        [void]$sb.AppendLine('')
    }

    # ── Scoring Methodology ──
    [void]$sb.AppendLine('SCORING METHODOLOGY')
    [void]$sb.AppendLine('-------------------')
    [void]$sb.AppendLine('  Score = weighted average: Critical(5x), High(4x), Medium(3x), Low(2x)')
    [void]$sb.AppendLine('  Pass/N/A = 100pts, Warning = 50pts, Fail/Not Assessed = 0pts')
    [void]$sb.AppendLine('  Maturity: Initial(0-34), Developing(35-54), Defined(55-74), Managed(75-89), Optimized(90+)')

    return $sb.ToString()
}

function Get-ExecutiveSummaryRtf {
    <#
    .SYNOPSIS  Generates a richly formatted RTF executive summary; pastes into Outlook/Word/Teams.
    #>
    $Checks   = $Global:Assessment.Checks
    $Score    = Get-OverallScore
    $Maturity = Get-MaturityLevel $Score

    $Total    = $Checks.Count
    $Assessed = @($Checks | Where-Object { $_.Status -ne 'Not Assessed' }).Count
    $Pass     = @($Checks | Where-Object { $_.Status -eq 'Pass' }).Count
    $Fail     = @($Checks | Where-Object { $_.Status -eq 'Fail' }).Count
    $Warn     = @($Checks | Where-Object { $_.Status -eq 'Warning' }).Count
    $NA       = @($Checks | Where-Object { $_.Status -eq 'N/A' }).Count

    $DefLookup = @{}
    foreach ($Def in $Global:CheckDefinitions) { $DefLookup[$Def.id] = $Def }

    # RTF escape helper
    $esc = { param($t) if (-not $t) { return '' }; "$t".Replace('\','\\').Replace('{','\{').Replace('}','\}') }

    # Color table (1-based): 1=accent blue, 2=green, 3=red, 4=orange, 5=gray,
    # 6=light bg, 7=white, 8=dark blue, 9=amber, 10=dark text, 11=green bg, 12=red bg, 13=orange bg, 14=blue bg
    $rtf = [System.Text.StringBuilder]::new()
    [void]$rtf.Append('{\rtf1\ansi\deff0')
    [void]$rtf.Append('{\fonttbl{\f0\fswiss\fcharset0 Segoe UI;}{\f1\fmodern\fcharset0 Cascadia Mono;}}')
    [void]$rtf.Append('{\colortbl;')
    [void]$rtf.Append('\red0\green120\blue212;')    # 1 accent blue
    [void]$rtf.Append('\red16\green124\blue16;')     # 2 green
    [void]$rtf.Append('\red209\green52\blue56;')     # 3 red
    [void]$rtf.Append('\red202\green80\blue16;')     # 4 orange
    [void]$rtf.Append('\red138\green136\blue134;')   # 5 gray
    [void]$rtf.Append('\red243\green242\blue241;')   # 6 light bg
    [void]$rtf.Append('\red255\green255\blue255;')   # 7 white
    [void]$rtf.Append('\red0\green99\blue177;')      # 8 dark blue
    [void]$rtf.Append('\red212\green140\blue0;')     # 9 amber
    [void]$rtf.Append('\red50\green50\blue50;')      # 10 dark text
    [void]$rtf.Append('\red232\green245\blue233;')   # 11 green bg
    [void]$rtf.Append('\red253\green232\blue232;')   # 12 red bg
    [void]$rtf.Append('\red255\green243\blue224;')   # 13 orange bg
    [void]$rtf.Append('\red227\green242\blue253;')   # 14 blue bg
    [void]$rtf.Append('}')
    [void]$rtf.Append('\f0\fs20\cf10 ')

    # ── Title ──
    $custName = if ($Global:Assessment.CustomerName) { " \u8212 $(& $esc $Global:Assessment.CustomerName)" } else { '' }
    [void]$rtf.Append("\pard\sb0\sa80\qc{\f0\fs36\b\cf1 AVD Assessment Summary$custName}\par")
    [void]$rtf.Append("{\f0\fs18\cf5 Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm')  \u8226  Azure Virtual Desktop}\par")
    [void]$rtf.Append('\pard\sb20\sa100{\f0\fs4\cf1\brdrb\brdrs\brdrw10\brsp40 \par}')

    # ── Score Hero (big number + maturity + status cards in a table row) ──
    $scoreStr = if ($Score -ge 0) { "$Score" } else { '\u8212' }
    $scoreClr = if ($Score -ge 75) { '\cf2' } elseif ($Score -ge 50) { '\cf4' } elseif ($Score -ge 0) { '\cf3' } else { '\cf5' }

    # Score + maturity on one line
    [void]$rtf.Append('\pard\sb0\sa20\qc')
    [void]$rtf.Append("{{\f0\fs72\b$scoreClr $scoreStr}}")
    [void]$rtf.Append("{\f0\fs28\cf5 /100}")
    [void]$rtf.Append('{\f0\fs20  }')
    [void]$rtf.Append("{{\f0\fs28\b\cf8 $Maturity}}")
    [void]$rtf.Append('\par')

    # Status summary as colored table cells
    [void]$rtf.Append('\pard\sb0\sa0\qc{\trowd\trgaph60\trqc')
    [void]$rtf.Append('\clcbpat11\clbrdrt\brdrs\brdrw5\brdrcf2\clbrdrb\brdrs\brdrw5\brdrcf2\clbrdrl\brdrs\brdrw5\brdrcf2\clbrdrr\brdrs\brdrw5\brdrcf2\cellx2000')
    [void]$rtf.Append('\clcbpat12\clbrdrt\brdrs\brdrw5\brdrcf3\clbrdrb\brdrs\brdrw5\brdrcf3\clbrdrl\brdrs\brdrw5\brdrcf3\clbrdrr\brdrs\brdrw5\brdrcf3\cellx4000')
    [void]$rtf.Append('\clcbpat13\clbrdrt\brdrs\brdrw5\brdrcf4\clbrdrb\brdrs\brdrw5\brdrcf4\clbrdrl\brdrs\brdrw5\brdrcf4\clbrdrr\brdrs\brdrw5\brdrcf4\cellx6000')
    [void]$rtf.Append('\clcbpat14\clbrdrt\brdrs\brdrw5\brdrcf1\clbrdrb\brdrs\brdrw5\brdrcf1\clbrdrl\brdrs\brdrw5\brdrcf1\clbrdrr\brdrs\brdrw5\brdrcf1\cellx8000')
    [void]$rtf.Append("\pard\intbl\sb40\sa40\qc{\f0\fs24\b\cf2 $Pass}\line{\f0\fs15\cf10 Pass}\cell")
    [void]$rtf.Append("\pard\intbl\sb40\sa40\qc{\f0\fs24\b\cf3 $Fail}\line{\f0\fs15\cf10 Fail}\cell")
    [void]$rtf.Append("\pard\intbl\sb40\sa40\qc{\f0\fs24\b\cf4 $Warn}\line{\f0\fs15\cf10 Warning}\cell")
    [void]$rtf.Append("\pard\intbl\sb40\sa40\qc{\f0\fs24\b\cf1 $NA}\line{\f0\fs15\cf10 N/A}\cell")
    [void]$rtf.Append('\row}')
    [void]$rtf.Append("\pard\sb20\sa80\qc{\f0\fs16\cf5 $Assessed of $Total checks assessed}\par")

    # Section header helper
    $SectionHeader = {
        param($title, $icon)
        [void]$rtf.Append("\pard\sb200\sa80\keepn{\f0\fs24\b\cf1 $icon  $(& $esc $title)}\par")
        [void]$rtf.Append('\pard\sb0\sa60{\f0\fs2\cf5\brdrb\brdrs\brdrw5\brsp20 \par}')
    }

    # ═══════════════ ASSESSMENT INFORMATION ═══════════════
    & $SectionHeader 'Assessment Information' '\u9889'
    [void]$rtf.Append('\pard\sb0\sa0{')
    $infoRows = [System.Collections.ArrayList]::new()
    [void]$infoRows.Add(@('Customer',  $Global:Assessment.CustomerName))
    [void]$infoRows.Add(@('Assessor',  $Global:Assessment.AssessorName))
    [void]$infoRows.Add(@('Date',      $Global:Assessment.Date))
    if ($Global:Assessment.Discovery) {
        $D = $Global:Assessment.Discovery
        if ($D.Subscriptions) {
            $subNames = ($D.Subscriptions | ForEach-Object { $_.Name }) -join ', '
            [void]$infoRows.Add(@('Subscriptions', "$($D.Subscriptions.Count) ($subNames)"))
        }
        if ($D.Inventory) {
            [void]$infoRows.Add(@('Host Pools',    "$($D.Inventory.HostPools.Count)"))
            [void]$infoRows.Add(@('Session Hosts', "$($D.Inventory.SessionHosts.Count)"))
            [void]$infoRows.Add(@('App Groups',    "$($D.Inventory.AppGroups.Count)"))
            [void]$infoRows.Add(@('Workspaces',    "$($D.Inventory.Workspaces.Count)"))
        }
    }
    $ri = 0
    foreach ($row in $infoRows) {
        $bgPat = if ($ri % 2 -eq 0) { '\clcbpat6' } else { '' }
        [void]$rtf.Append("\trowd\trgaph80${bgPat}\cellx2600${bgPat}\cellx8500")
        [void]$rtf.Append("\pard\intbl\sb20\sa20{\f0\fs18\b\cf8  $(& $esc $row[0])}\cell")
        [void]$rtf.Append("\pard\intbl\sb20\sa20{\f0\fs18\cf10  $(& $esc $row[1])}\cell")
        [void]$rtf.Append('\row')
        $ri++
    }
    [void]$rtf.Append('}')

    # ═══════════════ CATEGORY SCORES ═══════════════
    & $SectionHeader 'Category Scores' '\u9733'
    [void]$rtf.Append('\pard\sb0\sa0{\trowd\trgaph80')
    [void]$rtf.Append('\clcbpat1\cellx3800\clcbpat1\cellx4800\clcbpat1\cellx5600\clcbpat1\cellx6400\clcbpat1\cellx7200\clcbpat1\cellx8000\clcbpat1\cellx8800')
    [void]$rtf.Append('\pard\intbl\sb30\sa30{\f0\fs16\b\cf7  Category}\cell')
    [void]$rtf.Append('\pard\intbl\sb30\sa30\qc{\f0\fs16\b\cf7 Score}\cell')
    [void]$rtf.Append('\pard\intbl\sb30\sa30\qc{\f0\fs16\b\cf7 Pass}\cell')
    [void]$rtf.Append('\pard\intbl\sb30\sa30\qc{\f0\fs16\b\cf7 Warn}\cell')
    [void]$rtf.Append('\pard\intbl\sb30\sa30\qc{\f0\fs16\b\cf7 Fail}\cell')
    [void]$rtf.Append('\pard\intbl\sb30\sa30\qc{\f0\fs16\b\cf7 N/A}\cell')
    [void]$rtf.Append('\pard\intbl\sb30\sa30\qc{\f0\fs16\b\cf7 Total}\cell')
    [void]$rtf.Append('\row')
    $catIdx = 0
    foreach ($Cat in (Get-Categories)) {
        $CatScore  = Get-CategoryScore $Cat
        $CatChecks = @($Checks | Where-Object { $_.Category -eq $Cat })
        $CatTotal  = $CatChecks.Count
        $CatPass   = @($CatChecks | Where-Object { $_.Status -eq 'Pass' }).Count
        $CatWarn   = @($CatChecks | Where-Object { $_.Status -eq 'Warning' }).Count
        $CatFail   = @($CatChecks | Where-Object { $_.Status -eq 'Fail' }).Count
        $CatNA     = @($CatChecks | Where-Object { $_.Status -eq 'N/A' }).Count
        $ScoreStr  = if ($CatScore -ge 0) { "$CatScore%" } else { 'N/A' }
        $sClr      = if ($CatScore -ge 80) { '\cf2' } elseif ($CatScore -ge 50) { '\cf4' } elseif ($CatScore -ge 0) { '\cf3' } else { '\cf5' }
        $bgPat     = if ($catIdx % 2 -eq 0) { '\clcbpat6' } else { '' }
        [void]$rtf.Append("\trowd\trgaph80${bgPat}\cellx3800${bgPat}\cellx4800${bgPat}\cellx5600${bgPat}\cellx6400${bgPat}\cellx7200${bgPat}\cellx8000${bgPat}\cellx8800")
        [void]$rtf.Append("\pard\intbl\sb20\sa20{\f0\fs17\cf10  $(& $esc $Cat)}\cell")
        [void]$rtf.Append("\pard\intbl\sb20\sa20\qc{\f0\fs17\b$sClr $ScoreStr}\cell")
        [void]$rtf.Append("\pard\intbl\sb20\sa20\qc{\f0\fs17\cf2 $CatPass}\cell")
        [void]$rtf.Append("\pard\intbl\sb20\sa20\qc{\f0\fs17\cf4 $CatWarn}\cell")
        [void]$rtf.Append("\pard\intbl\sb20\sa20\qc{\f0\fs17\cf3 $CatFail}\cell")
        [void]$rtf.Append("\pard\intbl\sb20\sa20\qc{\f0\fs17\cf5 $CatNA}\cell")
        [void]$rtf.Append("\pard\intbl\sb20\sa20\qc{\f0\fs17\cf5 $CatTotal}\cell")
        [void]$rtf.Append('\row')
        $catIdx++
    }
    [void]$rtf.Append('}')

    # ═══════════════ MATURITY DIMENSIONS ═══════════════
    & $SectionHeader 'Maturity Dimensions' '\u9881'
    [void]$rtf.Append('\pard\sb0\sa0{\trowd\trgaph80')
    [void]$rtf.Append('\clcbpat1\cellx5400\clcbpat1\cellx6600\clcbpat1\cellx8500')
    [void]$rtf.Append('\pard\intbl\sb30\sa30{\f0\fs17\b\cf7  Dimension}\cell')
    [void]$rtf.Append('\pard\intbl\sb30\sa30\qc{\f0\fs17\b\cf7 Score}\cell')
    [void]$rtf.Append('\pard\intbl\sb30\sa30\qc{\f0\fs17\b\cf7 Level}\cell')
    [void]$rtf.Append('\row')
    $dimIdx = 0
    foreach ($Key in $Global:MaturityDimensions.Keys) {
        $Dim   = $Global:MaturityDimensions[$Key]
        $DScore = Get-DimensionScore $Key
        $DStr   = if ($DScore -ge 0) { "$DScore%" } else { 'N/A' }
        $DLevel = if ($DScore -ge 0) { Get-MaturityLevel $DScore } else { '' }
        $dClr   = if ($DScore -ge 80) { '\cf2' } elseif ($DScore -ge 50) { '\cf4' } elseif ($DScore -ge 0) { '\cf3' } else { '\cf5' }
        $bgPat  = if ($dimIdx % 2 -eq 0) { '\clcbpat6' } else { '' }
        [void]$rtf.Append("\trowd\trgaph80${bgPat}\cellx5400${bgPat}\cellx6600${bgPat}\cellx8500")
        [void]$rtf.Append("\pard\intbl\sb20\sa20{\f0\fs18\cf10  $(& $esc $Dim.Label)}\cell")
        [void]$rtf.Append("\pard\intbl\sb20\sa20\qc{\f0\fs18\b$dClr $DStr}\cell")
        [void]$rtf.Append("\pard\intbl\sb20\sa20\qc{\f0\fs17\cf8 $DLevel}\cell")
        [void]$rtf.Append('\row')
        $dimIdx++
    }
    $Composite = Get-CompositeMaturityScore
    if ($Composite -ge 0) {
        $cClr = if ($Composite -ge 80) { '\cf2' } elseif ($Composite -ge 50) { '\cf4' } else { '\cf3' }
        [void]$rtf.Append("\trowd\trgaph80\clcbpat6\cellx5400\clcbpat6\cellx6600\clcbpat6\cellx8500")
        [void]$rtf.Append("\pard\intbl\sb20\sa20{\f0\fs18\b\cf1  Composite}\cell")
        [void]$rtf.Append("\pard\intbl\sb20\sa20\qc{\f0\fs18\b$cClr $Composite%}\cell")
        [void]$rtf.Append("\pard\intbl\sb20\sa20\qc{\f0\fs17\b\cf8 $(Get-MaturityLevel $Composite)}\cell")
        [void]$rtf.Append('\row')
    }
    [void]$rtf.Append('}')

    # ═══════════════ KEY PASSES ═══════════════
    $strongChecks = @($Checks | Where-Object { $_.Status -eq 'Pass' -and $_.Severity -in @('Critical','High') })
    if ($strongChecks.Count -gt 0) {
        & $SectionHeader "Key Passes ($($strongChecks.Count) Critical/High)" '\u10004'
        $grouped = $strongChecks | Group-Object Category | Sort-Object Count -Descending
        foreach ($g in $grouped) {
            [void]$rtf.Append("\pard\sb60\sa20\li200{\f0\fs18\b\cf8 $(& $esc $g.Name) ($($g.Count))}\par")
            $items = @($g.Group | Select-Object -First 5)
            foreach ($item in $items) {
                [void]$rtf.Append("\pard\sb0\sa0\li400{\f0\fs17\cf2 \u10003  }{\f0\fs17\cf10 $(& $esc $item.Name)}\par")
            }
            if ($g.Count -gt 5) {
                [void]$rtf.Append("\pard\sb0\sa0\li400{\f0\fs17\i\cf5 \u8230  and $($g.Count - 5) more}\par")
            }
        }
    }

    # ═══════════════ CRITICAL & HIGH FAILURES ═══════════════
    $critFails = @($Checks | Where-Object {
        $_.Status -eq 'Fail' -and $_.Severity -in @('Critical','High')
    } | Sort-Object @{E={if($_.Severity -eq 'Critical'){0}else{1}}}, Id)
    if ($critFails.Count -gt 0) {
        & $SectionHeader "Critical & High Failures ($($critFails.Count))" '\u9888'
        [void]$rtf.Append('\pard\sb0\sa0{\trowd\trgaph80')
        [void]$rtf.Append('\clcbpat1\cellx1200\clcbpat1\cellx2200\clcbpat1\cellx5600\clcbpat1\cellx8500')
        [void]$rtf.Append('\pard\intbl\sb30\sa30{\f0\fs17\b\cf7  Sev}\cell')
        [void]$rtf.Append('\pard\intbl\sb30\sa30{\f0\fs17\b\cf7  ID}\cell')
        [void]$rtf.Append('\pard\intbl\sb30\sa30{\f0\fs17\b\cf7  Check Name}\cell')
        [void]$rtf.Append('\pard\intbl\sb30\sa30{\f0\fs17\b\cf7  Details}\cell')
        [void]$rtf.Append('\row')
        $fi = 0
        foreach ($f in $critFails) {
            $sevClr  = if ($f.Severity -eq 'Critical') { '\cf3' } else { '\cf4' }
            $bgPat   = if ($fi % 2 -eq 0) { '\clcbpat6' } else { '' }
            $detVal  = if ($f.Details) { $d = & $esc "$($f.Details)"; if ($d.Length -gt 60) { $d.Substring(0,60) + '...' } else { $d } } else { '\u8212' }
            [void]$rtf.Append("\trowd\trgaph80${bgPat}\cellx1200${bgPat}\cellx2200${bgPat}\cellx5600${bgPat}\cellx8500")
            [void]$rtf.Append("\pard\intbl\sb20\sa20{\f0\fs17\b$sevClr  $($f.Severity)}\cell")
            [void]$rtf.Append("\pard\intbl\sb20\sa20{\f1\fs16\cf8  $($f.Id)}\cell")
            [void]$rtf.Append("\pard\intbl\sb20\sa20{\f0\fs17\cf10  $(& $esc $f.Name)}\cell")
            [void]$rtf.Append("\pard\intbl\sb20\sa20{\f1\fs15\cf5  $detVal}\cell")
            [void]$rtf.Append('\row')
            $fi++
        }
        [void]$rtf.Append('}')
    }

    # ═══════════════ OTHER FAILURES ═══════════════
    $medLowFails = @($Checks | Where-Object { $_.Status -eq 'Fail' -and $_.Severity -in @('Medium','Low') })
    if ($medLowFails.Count -gt 0) {
        $medCount = @($medLowFails | Where-Object { $_.Severity -eq 'Medium' }).Count
        $lowCount = @($medLowFails | Where-Object { $_.Severity -eq 'Low' }).Count
        [void]$rtf.Append("\pard\sb120\sa60{\f0\fs18\cf5 Additional failures: {\b\cf4 $medCount} Medium, {\b\cf5 $lowCount} Low}\par")
    }

    # ═══════════════ QUICK WIN REMEDIATION ═══════════════
    $quickWins = @($Checks | Where-Object {
        $_.Status -eq 'Fail' -and ($DefLookup[$_.Id].effort -eq 'Quick Win' -or $_.effort -eq 'Quick Win')
    } | Sort-Object @{E={switch($_.Severity){'Critical'{0}'High'{1}'Medium'{2}default{3}}}}, Id | Select-Object -First 10)
    if ($quickWins.Count -gt 0) {
        & $SectionHeader "Quick Win Remediation (top $($quickWins.Count))" '\u9889'
        $qi = 0
        foreach ($qw in $quickWins) {
            $qi++
            $sevClr = switch ($qw.Severity) { 'Critical' { '\cf3' } 'High' { '\cf4' } default { '\cf5' } }
            [void]$rtf.Append("\pard\sb40\sa0\li200{\f0\fs18\b\cf10 ${qi}. }{\f0\fs17$sevClr [$($qw.Severity)]} {\f0\fs18\cf10 $(& $esc $qw.Name)} {\f1\fs15\cf5 ($($qw.Id))}\par")
            $rem = if ($qw.remediation) { $qw.remediation } elseif ($DefLookup[$qw.Id].remediation) { $DefLookup[$qw.Id].remediation } else { $null }
            if ($rem) {
                $remText = if ($rem.Length -gt 120) { $rem.Substring(0,120) + '...' } else { $rem }
                [void]$rtf.Append("\pard\sb0\sa20\li400{\f0\fs16\i\cf5 $(& $esc $remText)}\par")
            }
        }
    }

    # ═══════════════ SCORING METHODOLOGY ═══════════════
    [void]$rtf.Append('\pard\sb200\sa60{\f0\fs2\cf5\brdrb\brdrs\brdrw5\brsp20 \par}')
    [void]$rtf.Append('\pard\sb40\sa20{\f0\fs16\i\cf5 Scoring: Weighted average \u8212 Critical(5\u215), High(4\u215), Medium(3\u215), Low(2\u215). ')
    [void]$rtf.Append('Pass/N\u8725A = 100pts, Warning = 50pts, Fail = 0pts. ')
    [void]$rtf.Append('Maturity: Initial(0\u821234), Developing(35\u821254), Defined(55\u821274), Managed(75\u821289), Optimized(90+).}\par')

    [void]$rtf.Append('}')
    return $rtf.ToString()
}

<#
.SYNOPSIS
    Exports the assessment as a standalone HTML report via SaveFileDialog.
.DESCRIPTION
    Calls Build-HtmlReport, writes to the user-chosen path, unlocks the export_html
    achievement, and optionally opens the file in the default browser.
#>
function Export-HtmlReport {
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'HTML Files (*.html)|*.html'
    $dlg.FileName = "AVD_Assessment_$(& { $n = $Global:Assessment.CustomerName -replace '[^\w]','_'; if ($n) { "${n}_" } })$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    $dlg.InitialDirectory = $Global:ReportDir

    if ($dlg.ShowDialog() -eq $true) {
        try {
            $Html = Build-HtmlReport
            [System.IO.File]::WriteAllText($dlg.FileName, $Html, [System.Text.Encoding]::UTF8)
            Write-DebugLog "HTML report exported: $($dlg.FileName)" -Level 'SUCCESS'
            Show-Toast "Report exported: $(Split-Path $dlg.FileName -Leaf)" -Type 'Success'
            Unlock-Achievement 'export_html'
            if ($Global:OpenAfterExport) { Start-Process $dlg.FileName }
        } catch {
            Write-DebugLog "Export failed: $($_.Exception.Message)" -Level 'ERROR'
            Show-Toast "Export failed: $($_.Exception.Message)" -Type 'Error'
        }
    }
}

<#
.SYNOPSIS
    Exports the assessment check data as a CSV file via SaveFileDialog.
.DESCRIPTION
    Exports all check fields (ID, Category, Name, Status, Severity, Weight, etc.) to CSV
    format. Unlocks the export_csv achievement on success.
#>
function Export-CsvReport {
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'CSV Files (*.csv)|*.csv'
    $dlg.FileName = "AVD_Assessment_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $dlg.InitialDirectory = $Global:ReportDir

    if ($dlg.ShowDialog() -eq $true) {
        try {
            $Global:Assessment.Checks | Select-Object Id, Category, Name, Description, Status, Severity, Weight, Excluded, Origin, Source, Details, Notes |
                Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
            Write-DebugLog "CSV exported: $($dlg.FileName)" -Level 'SUCCESS'
            Unlock-Achievement 'export_csv'
            Show-Toast "CSV exported: $(Split-Path $dlg.FileName -Leaf)" -Type 'Success'
            if ($Global:OpenAfterExport) { Start-Process $dlg.FileName }
        } catch {
            Show-Toast "CSV export failed: $($_.Exception.Message)" -Type 'Error'
        }
    }
}

<#
.SYNOPSIS
    Exports the full assessment object as a JSON file via SaveFileDialog.
.DESCRIPTION
    Serializes the complete assessment (checks, metadata, discovery data) to JSON at depth 10.
#>
function Export-JsonAssessment {
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'JSON Files (*.json)|*.json'
    $dlg.FileName = "AVD_Assessment_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    $dlg.InitialDirectory = $Global:ReportDir

    if ($dlg.ShowDialog() -eq $true) {
        try {
            $Global:Assessment | ConvertTo-Json -Depth 10 | Set-Content $dlg.FileName -Encoding UTF8 -Force
            Write-DebugLog "JSON exported: $($dlg.FileName)" -Level 'SUCCESS'
            Show-Toast "JSON exported: $(Split-Path $dlg.FileName -Leaf)" -Type 'Success'
            if ($Global:OpenAfterExport) { Start-Process $dlg.FileName }
        } catch {
            Show-Toast "JSON export failed: $($_.Exception.Message)" -Type 'Error'
        }
    }
}

# ===============================================================================
# SECTION 14: SAVE / LOAD ASSESSMENT
# Assessment persistence: auto-save to _autosave.json with rolling backups (max 10), manual
# save with customer-slug filename, load with discovery vs. assessment file detection,
# dirty-flag tracking, and save/discard/cancel confirmation dialogs.
# ===============================================================================

<#
.SYNOPSIS
    Captures current UI field values (customer name, assessor, date, notes) into the assessment object.
#>
function Sync-AssessmentFromUI {
    $Global:Assessment.CustomerName = $txtCustomerName.Text
    $Global:Assessment.AssessorName = $txtAssessorName.Text
    $Global:Assessment.Date         = $txtAssessmentDate.Text
    $Global:Assessment.Notes        = if ($txtNotes) { $txtNotes.Text } else { '' }
}

<#
.SYNOPSIS
    Marks the assessment as having unsaved changes and shows the dirty indicator dot.
#>
function Set-Dirty {
    if (-not $Global:IsDirty) {
        $Global:IsDirty = $true
        if ($dotDirty)    { $dotDirty.Visibility = 'Visible' }
        if ($lblDirtyText) { $lblDirtyText.Visibility = 'Visible' }
    }
}

<#
.SYNOPSIS
    Clears the unsaved-changes flag and hides the dirty indicator dot.
#>
function Clear-Dirty {
    $Global:IsDirty = $false
    if ($dotDirty)    { $dotDirty.Visibility = 'Collapsed' }
    if ($lblDirtyText) { $lblDirtyText.Visibility = 'Collapsed' }
}

<#
.SYNOPSIS
    Silently auto-saves the assessment to _autosave.json with rolling backup (max 10).
.DESCRIPTION
    When auto-save is enabled, writes the current assessment to the primary autosave file
    and creates a timestamped rolling backup in _backups/, purging old backups beyond the
    configured maximum.
#>
function AutoSave-Assessment {
    if (-not $Global:AutoSaveEnabled) { return }
    Set-Dirty
    try {
        Sync-AssessmentFromUI
        $JsonStr = $Global:Assessment | ConvertTo-Json -Depth 10

        # Write primary autosave
        [System.IO.File]::WriteAllText($Global:AutoSaveFile, $JsonStr, [System.Text.Encoding]::UTF8)

        # Rolling backup — _backups/ folder, max 10 files
        $BackupDir = Join-Path $Global:Root '_backups'
        if (-not (Test-Path $BackupDir)) {
            New-Item -Path $BackupDir -ItemType Directory -Force | Out-Null
        }
        $BackupPath = Join-Path $BackupDir "backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
        [System.IO.File]::WriteAllText($BackupPath, $JsonStr, [System.Text.Encoding]::UTF8)

        # Purge old backups beyond limit
        $OldBackups = @(Get-ChildItem $BackupDir -Filter 'backup_*.json' -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -Skip $Global:MaxBackups)
        foreach ($Old in $OldBackups) {
            try { [System.IO.File]::Delete($Old.FullName) } catch { }
        }
    } catch { }
}

<#
.SYNOPSIS
    Saves the current assessment to a JSON file in the assessments directory.
.DESCRIPTION
    Syncs UI state, validates that a customer name is set, determines the save path
    (overwrite existing or create new slug-based filename), writes JSON, clears the dirty
    flag, and triggers achievement checks.
#>
function Save-Assessment {
    Sync-AssessmentFromUI

    # Prompt for customer name if missing
    if (-not $Global:Assessment.CustomerName.Trim()) {
        $Result = Show-ThemedDialog `
            -Title 'Customer Name Required' `
            -Message 'Please enter a customer name before saving.' `
            -Icon ([char]0xE77B) `
            -IconColor 'ThemeWarning' `
            -Buttons @(
                @{ Text = 'OK'; IsAccent = $true; Result = 'OK' }
            )
        # Focus the customer name field and switch to dashboard
        Switch-Tab 'Dashboard'
        $txtCustomerName.Focus()
        return
    }

    # If we have an active file path, overwrite it. Otherwise create new.
    # Also regenerate path if current file has 'untitled' but customer name is now set.
    $NeedNewPath = (-not $Global:ActiveFilePath) -or (-not (Test-Path $Global:ActiveFilePath))
    if (-not $NeedNewPath -and $Global:Assessment.CustomerName.Trim() -and
        (Split-Path $Global:ActiveFilePath -Leaf) -match '^untitled_') {
        $NeedNewPath = $true
    }

    if ($NeedNewPath) {
        $CustSlug = if ($Global:Assessment.CustomerName) {
            ($Global:Assessment.CustomerName -replace '[^\w\-]','_').Trim('_')
        } else { 'untitled' }
        $Path = Join-Path $Global:AssessmentDir "${CustSlug}_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
        # Remove the old untitled file if renaming
        if ($Global:ActiveFilePath -and (Test-Path $Global:ActiveFilePath) -and
            (Split-Path $Global:ActiveFilePath -Leaf) -match '^untitled_') {
            try { [System.IO.File]::Delete($Global:ActiveFilePath) } catch { }
        }
        $Global:ActiveFilePath = $Path
    } else {
        $Path = $Global:ActiveFilePath
    }
    try {
        $Global:Assessment | ConvertTo-Json -Depth 10 | Set-Content $Path -Encoding UTF8 -Force
        Clear-Dirty
        Write-DebugLog "Assessment saved: $Path" -Level 'SUCCESS'
        Show-Toast "Saved: $(Split-Path $Path -Leaf)" -Type 'Success'
        Check-AssessmentAchievements
    } catch {
        Write-DebugLog "Save failed: $($_.Exception.Message)" -Level 'ERROR'
        Show-Toast "Save failed: $($_.Exception.Message)" -Type 'Error'
    }
}

<#
.SYNOPSIS
    Returns true if the current assessment has any assessed checks or notes entered.
#>
function Test-AssessmentDirty {
    $Assessed = @($Global:Assessment.Checks | Where-Object { $_.Status -ne 'Not Assessed' }).Count
    $HasNotes = @($Global:Assessment.Checks | Where-Object { $_.Notes }).Count
    return ($Assessed -gt 0 -or $HasNotes -gt 0)
}

<#
.SYNOPSIS
    Shows a save/discard/cancel dialog if unsaved work exists, returns true if safe to proceed.
.PARAMETER Action
    Description of the action about to be taken (e.g. 'import', 'load').
.OUTPUTS
    Boolean indicating whether the caller should proceed.
#>
function Confirm-OverwriteAssessment {
    param([string]$Action = 'import')
    if (-not (Test-AssessmentDirty)) { return $true }

    $Assessed = @($Global:Assessment.Checks | Where-Object { $_.Status -ne 'Not Assessed' }).Count
    $Result = Show-ThemedDialog `
        -Title 'Unsaved Work' `
        -Message "You have $Assessed check(s) assessed with notes or status changes.`n`nSave before ${Action}?" `
        -Icon ([char]0xE7BA) `
        -IconColor 'ThemeWarning' `
        -Buttons @(
            @{ Text = 'Save First'; Result = 'Save'; IsAccent = $true }
            @{ Text = 'Discard';    Result = 'Discard' }
            @{ Text = 'Cancel';     Result = 'Cancel' }
        )

    switch ($Result) {
        'Save'    { Save-Assessment; return $true }
        'Discard' { return $true }
        default   { return $false }
    }
}

<#
.SYNOPSIS
    Loads a saved assessment or discovery JSON via OpenFileDialog.
.DESCRIPTION
    Detects whether the selected file is a discovery JSON (has SchemaVersion + Inventory)
    or a saved assessment. For discovery files, creates a fresh assessment and imports
    automated results. For assessments, loads check state, syncs definitions, and restores
    the assessment view. Prompts to save unsaved work before loading.
#>
function Load-Assessment {
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'JSON Files (*.json)|*.json'
    $dlg.InitialDirectory = $Global:AssessmentDir
    $dlg.Title = 'Load Assessment or Discovery JSON'

    if ($dlg.ShowDialog() -eq $true) {
        try {
            $Json = Get-Content $dlg.FileName -Raw -Encoding UTF8 | ConvertFrom-Json

            # Determine if this is a discovery file or assessment file
            if ($Json.SchemaVersion -and $Json.Inventory) {
                # Discovery JSON — create a fresh assessment, then import discovery into it
                $Global:Assessment = [PSCustomObject]@{
                    CustomerName  = $txtCustomerName.Text
                    AssessorName  = $txtAssessorName.Text
                    Date          = (Get-Date -Format 'yyyy-MM-dd')
                    Notes         = $txtNotes.Text
                    Discovery     = $null
                    Checks        = [System.Collections.ArrayList]::new()
                    ManualOverrides = @{}
                }
                foreach ($Def in $Global:CheckDefinitions) {
                    [void]$Global:Assessment.Checks.Add([PSCustomObject]@{
                        Id=$Def.id;Category=$Def.category;Name=$Def.name;Description=$Def.description
                        Severity=$Def.severity;Type=$Def.type;Weight=[int]$Def.weight;Reference=$Def.reference
                        Origin=$Def.origin;Status='Not Assessed';Excluded=$false;Details='';Notes='';Source=$Def.type;Effort=$Def.effort
                    })
                }
                Sync-CheckDefinitions
                $Global:ActiveFilePath = $null
                $txtAssessmentDate.Text = $Global:Assessment.Date
                Import-DiscoveryJson -Path $dlg.FileName
                Render-AssessmentChecks
            } elseif ($Json.Checks) {
                # Assessment JSON — restore full state
                $Global:Assessment = $Json
                $Global:Assessment.Checks = [System.Collections.ArrayList]@($Json.Checks)
                # Merge new checks from checks.json
                $ExIds = @($Json.Checks | ForEach-Object { $_.Id })
                foreach ($Def in $Global:CheckDefinitions) {
                    if ($Def.id -notin $ExIds) {
                        [void]$Global:Assessment.Checks.Add([PSCustomObject]@{
                            Id=$Def.id;Category=$Def.category;Name=$Def.name;Description=$Def.description
                            Severity=$Def.severity;Type=$Def.type;Weight=[int]$Def.weight;Reference=$Def.reference
                            Origin=$Def.origin;Status='Not Assessed';Excluded=$false;Details='';Notes='';Source=$Def.type;Effort=$Def.effort
                        })
                    }
                }
                Sync-CheckDefinitions
                $txtCustomerName.Text  = $Json.CustomerName
                $txtAssessorName.Text  = $Json.AssessorName
                $txtAssessmentDate.Text = $Json.Date
                $txtNotes.Text         = $Json.Notes

                if ($Json.Discovery) {
                    $lblStatHostPools.Text    = "$($Json.Discovery.Inventory.HostPools.Count)"
                    $lblStatSessionHosts.Text = "$($Json.Discovery.Inventory.SessionHosts.Count)"
                    $lblStatAppGroups.Text    = "$($Json.Discovery.Inventory.AppGroups.Count)"
                    $lblStatWorkspaces.Text   = "$($Json.Discovery.Inventory.Workspaces.Count)"
                    $lblStatScalingPlans.Text = "$($Json.Discovery.Inventory.ScalingPlans.Count)"
                }

                Update-Dashboard
                Update-Progress
                Render-AssessmentChecks
                $Global:ActiveFilePath = $dlg.FileName
                Clear-Dirty
                Write-DebugLog "Assessment loaded: $($dlg.FileName)" -Level 'SUCCESS'
                Show-Toast "Assessment loaded: $($Json.Checks.Count) checks" -Type 'Success'
            } else {
                Show-Toast "Unrecognized JSON format" -Type 'Error'
            }
        } catch {
            Write-DebugLog "Load failed: $($_.Exception.Message)" -Level 'ERROR'
            Show-Toast "Load failed: $($_.Exception.Message)" -Type 'Error'
        }
    }
}

# ===============================================================================
# SECTION 14B: SAVED ASSESSMENTS LIST
# ===============================================================================

<#
.SYNOPSIS
    Populates the saved assessments ListView in the sidebar from the assessments directory.
.DESCRIPTION
    Scans for JSON files (excluding autosave and discovery files), parses metadata, and
    renders each as a clickable list item with customer name, date, and check status counts.
#>
function Refresh-AssessmentList {
    <# Scans assessments/ dir and populates sidebar. Shows assessments first, then discoveries. #>
    $lstSavedAssessments.Items.Clear()

    if (-not (Test-Path $Global:AssessmentDir)) {
        $lblSavedEmpty.Visibility = 'Visible'
        $lblSavedCount.Text = ''
        return
    }

    # Get all JSON files, exclude _autosave.json
    $AllFiles = @(Get-ChildItem $Global:AssessmentDir -Filter '*.json' -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -ne '_autosave.json' } |
        Sort-Object LastWriteTime -Descending)

    if ($AllFiles.Count -eq 0) {
        $lblSavedEmpty.Visibility = 'Visible'
        $lblSavedCount.Text = ''
        return
    }
    $lblSavedEmpty.Visibility = 'Collapsed'

    foreach ($F in $AllFiles) {
        $IsDiscovery   = $false
        $DisplayName   = $F.BaseName
        $Subtitle      = $F.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
        $ProgressText  = ''
        $ScorePercent  = -1

        # Peek metadata
        try {
            $Peek = Get-Content $F.FullName -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json
            if ($Peek.SchemaVersion -and $Peek.Inventory) {
                # Skip discovery JSON files — only show assessments
                continue
            } elseif ($Peek.Checks) {
                # Count using merged total (saved checks + any new checks from checks.json)
                $SavedIds = @($Peek.Checks | ForEach-Object { $_.Id })
                $NewCount = @($Global:CheckDefinitions | Where-Object { $_.id -notin $SavedIds }).Count
                $Total    = @($Peek.Checks).Count + $NewCount
                $Assessed = @($Peek.Checks | Where-Object { $_.Status -ne 'Not Assessed' }).Count
                $DisplayName = if ($Peek.CustomerName) { $Peek.CustomerName } else { 'Untitled Assessment' }
                $ProgressText = "$Assessed/$Total"
                if ($Assessed -gt 0) {
                    # Honest scoring matching dashboard: Not Assessed = 0 pts, included in denominator, weighted
                    $ScoringChecks = @($Peek.Checks)
                    $WeightedSum = 0; $TotalWeight = 0
                    foreach ($SC in $ScoringChecks) {
                        $W = [math]::Max(1, [int]$SC.Weight)
                        $Pts = switch ($SC.Status) { 'Pass' { 100 } 'N/A' { 100 } 'Warning' { 50 } default { 0 } }
                        $WeightedSum += $Pts * $W; $TotalWeight += $W
                    }
                    if ($TotalWeight -gt 0) { $ScorePercent = [math]::Round($WeightedSum / $TotalWeight, 0) }
                }
            }
        } catch { }

        $Item = New-Object System.Windows.Controls.ListViewItem
        $Item.Tag = $F.FullName
        $Item.Cursor = [System.Windows.Input.Cursors]::Hand

        $SP = New-Object System.Windows.Controls.StackPanel

        # Row 1: Icon badge + display name
        $Row1 = New-Object System.Windows.Controls.DockPanel

        $IconBdr = New-Object System.Windows.Controls.Border
        $IconBdr.Width = 28; $IconBdr.Height = 28
        $IconBdr.CornerRadius = [System.Windows.CornerRadius]::new(6)
        $IconBdr.Margin = [System.Windows.Thickness]::new(0,0,8,0)
        $IconBdr.VerticalAlignment = 'Center'
        $IconTB = New-Object System.Windows.Controls.TextBlock
        $IconTB.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe Fluent Icons, Segoe MDL2 Assets")
        $IconTB.FontSize = 13
        $IconTB.HorizontalAlignment = 'Center'
        $IconTB.VerticalAlignment = 'Center'

        if ($IsDiscovery) {
            $IconBdr.Background = $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#E0F0FF' } else { '#0A1E3D' }))
            $IconTB.Text = [char]0xE9D9
            $IconTB.Foreground = $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#0078D4' } else { '#60CDFF' }))
        } else {
            $IconBdr.Background = $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#E8F5E9' } else { '#0A3D1A' }))
            $IconTB.Text = [char]0xE9F9
            $IconTB.Foreground = $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#15803D' } else { '#4ADE80' }))
        }
        $IconBdr.Child = $IconTB

        # Name + score on right
        $NameSP = New-Object System.Windows.Controls.StackPanel
        $NameTB = New-Object System.Windows.Controls.TextBlock
        $NameTB.Text = $DisplayName
        $NameTB.FontSize = 12; $NameTB.FontWeight = [System.Windows.FontWeights]::SemiBold
        $NameTB.TextTrimming = 'CharacterEllipsis'; $NameTB.MaxWidth = 160
        $NameTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')

        $SubTB = New-Object System.Windows.Controls.TextBlock
        $SubTB.Text = $Subtitle
        $SubTB.FontSize = 11
        $SubTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextFaintest')

        [void]$NameSP.Children.Add($NameTB)
        [void]$NameSP.Children.Add($SubTB)

        # Score badge on right (for assessments with scores)
        if ($ScorePercent -ge 0) {
            $ScoreTB = New-Object System.Windows.Controls.TextBlock
            $ScoreTB.Text = "$ScorePercent%"
            $ScoreTB.FontSize = 12; $ScoreTB.FontWeight = [System.Windows.FontWeights]::Bold
            $ScoreTB.VerticalAlignment = 'Center'
            $ScoreTB.Foreground = if ($ScorePercent -ge 80) { $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#15803D' } else { '#4ADE80' })) }
                                  elseif ($ScorePercent -ge 50) { $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#C2410C' } else { '#FBBF24' })) }
                                  else { $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#DC2626' } else { '#FB923C' })) }
            [System.Windows.Controls.DockPanel]::SetDock($ScoreTB, 'Right')
            [void]$Row1.Children.Add($ScoreTB)
        } elseif ($ProgressText) {
            $ProgTB = New-Object System.Windows.Controls.TextBlock
            $ProgTB.Text = $ProgressText
            $ProgTB.FontSize = 12
            $ProgTB.VerticalAlignment = 'Center'
            $ProgTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
            [System.Windows.Controls.DockPanel]::SetDock($ProgTB, 'Right')
            [void]$Row1.Children.Add($ProgTB)
        }

        [void]$Row1.Children.Add($IconBdr)
        [void]$Row1.Children.Add($NameSP)
        [void]$SP.Children.Add($Row1)

        $Item.Content = $SP
        [void]$lstSavedAssessments.Items.Add($Item)
    }

    $lblSavedCount.Text = "$($AllFiles.Count)"
}

# ===============================================================================
# SECTION 15: USER PREFERENCES
# ===============================================================================
# SECTION 15B: ACHIEVEMENTS
# ===============================================================================

$Global:AchievementDefs = @(
    @{ Id='first_assessment';   Name='First Steps';      Icon='🎯'; Desc='Complete your first assessment' }
    @{ Id='five_assessments';   Name='Repeat Auditor';   Icon='📋'; Desc='Complete 5 assessments' }
    @{ Id='ten_assessments';    Name='Assessment Pro';   Icon='🏆'; Desc='Complete 10 assessments' }
    @{ Id='first_discovery';    Name='Explorer';         Icon='🔍'; Desc='Import your first discovery file' }
    @{ Id='full_sweep';         Name='Full Sweep';       Icon='🧹'; Desc='Assess all checks in a single run' }
    @{ Id='perfect_category';   Name='Flawless';         Icon='✨'; Desc='100% pass rate in any category' }
    @{ Id='maturity_managed';   Name='Well Managed';     Icon='📈'; Desc='Reach Managed maturity level (75%+)' }
    @{ Id='maturity_optimized'; Name='Peak Performance'; Icon='💎'; Desc='Reach Optimized maturity level (90%+)' }
    @{ Id='export_html';        Name='Reporter';         Icon='📄'; Desc='Export your first HTML report' }
    @{ Id='export_csv';         Name='Data Wrangler';    Icon='📊'; Desc='Export a CSV report' }
    @{ Id='night_owl';          Name='Night Owl';        Icon='🦉'; Desc='Run an assessment between 00:00-05:00' }
    @{ Id='early_bird';         Name='Early Bird';       Icon='🌅'; Desc='Run an assessment between 05:00-07:00' }
    @{ Id='weekend_warrior';    Name='Weekend Warrior';  Icon='🛡'; Desc='Save an assessment on a weekend' }
    @{ Id='theme_toggle';       Name='Chameleon';        Icon='🎨'; Desc='Toggle the theme for the first time' }
    @{ Id='note_taker';         Name='Note Taker';       Icon='📝'; Desc='Add notes to 5 different checks' }
    @{ Id='exclusion_expert';   Name='Scope Master';     Icon='🎯'; Desc='Exclude a check from scoring' }
    @{ Id='zero_critical';      Name='Zero Critical';    Icon='🛡'; Desc='No critical-severity failures' }
    @{ Id='fifty_pass';         Name='Half Way';         Icon='⭐'; Desc='Pass 50 checks in a single assessment' }
    @{ Id='hundred_pass';       Name='Century Club';     Icon='💯'; Desc='Pass 100 checks in a single assessment' }
    @{ Id='speed_demon';        Name='Speed Demon';      Icon='⚡'; Desc='Complete assessment in under 5 minutes' }
)

$Global:Achievements = @{}

<#
.SYNOPSIS
    Unlocks a specific achievement if not already unlocked, shows a toast, and saves preferences.
.PARAMETER Id
    The achievement identifier (e.g. 'first_assessment', 'full_sweep').
#>
function Unlock-Achievement {
    param([string]$Id)
    if ($Global:Achievements.ContainsKey($Id)) { return }
    $Def = $Global:AchievementDefs | Where-Object { $_.Id -eq $Id }
    if (-not $Def) { return }
    $Global:Achievements[$Id] = (Get-Date).ToString('o')
    Write-DebugLog "Achievement unlocked: $($Def.Name)" -Level 'SUCCESS'
    Show-Toast "$($Def.Icon) Achievement Unlocked: $($Def.Name) — $($Def.Desc)" -Type 'Success' -DurationMs 5000
    Update-AchievementBadges
    Save-UserPrefs
}

<#
.SYNOPSIS
    Renders the achievement badge grid in the Settings tab showing locked and unlocked achievements.
.DESCRIPTION
    Iterates all 20 achievement definitions, renders each as a themed badge with its icon
    (or locked placeholder), tooltip with name and description, and updates the unlock counter.
#>
function Update-AchievementBadges {
    $pnl = $Window.FindName('pnlAchievements')
    $lbl = $Window.FindName('lblAchievementCount')
    if (-not $pnl) { return }
    $pnl.Children.Clear()
    $Unlocked = 0
    foreach ($Def in $Global:AchievementDefs) {
        $IsUnlocked = $Global:Achievements.ContainsKey($Def.Id)
        if ($IsUnlocked) { $Unlocked++ }
        $Badge = New-Object System.Windows.Controls.Border
        $Badge.Width = 32; $Badge.Height = 32
        $Badge.CornerRadius = [System.Windows.CornerRadius]::new(6)
        $Badge.Margin = [System.Windows.Thickness]::new(2)
        if ($IsUnlocked) {
            $Badge.Background = $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#E8F5E9' } else { '#1A3D2A' }))
            $Badge.BorderBrush = $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#15803D' } else { '#00C853' }))
        } else {
            $Badge.Background = $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#F5F5F5' } else { '#2A2A2A' }))
            $Badge.BorderBrush = $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#E0E0E0' } else { '#404040' }))
        }
        $Badge.BorderThickness = [System.Windows.Thickness]::new(1)
        $Badge.ToolTip = "$($Def.Name): $($Def.Desc)$(if ($IsUnlocked) { "`nUnlocked: $($Global:Achievements[$Def.Id])" })"
        $BadgeText = New-Object System.Windows.Controls.TextBlock
        $BadgeText.Text = if ($IsUnlocked) { $Def.Icon } else { '?' }
        $BadgeText.FontSize = 16
        $BadgeText.HorizontalAlignment = 'Center'
        $BadgeText.VerticalAlignment = 'Center'
        $Badge.Child = $BadgeText
        [void]$pnl.Children.Add($Badge)
    }
    if ($lbl) { $lbl.Text = "$Unlocked/$($Global:AchievementDefs.Count)" }
}

<#
.SYNOPSIS
    Evaluates all achievement criteria and unlocks any newly earned achievements.
.DESCRIPTION
    Called after saves, imports, and exports to check milestone triggers: assessed counts,
    saved file counts, full sweep, perfect category, maturity thresholds, zero critical
    failures, pass milestones, note-taking, exclusions, and time-based achievements.
#>
function Check-AssessmentAchievements {
    $Checks = $Global:Assessment.Checks
    $Assessed = @($Checks | Where-Object { $_.Status -ne 'Not Assessed' })
    $Passed = @($Assessed | Where-Object { $_.Status -eq 'Pass' -or $_.Status -eq 'N/A' })
    $Failed = @($Assessed | Where-Object { $_.Status -eq 'Fail' })

    # Milestone: first assessment (any check assessed)
    if ($Assessed.Count -gt 0) { Unlock-Achievement 'first_assessment' }

    # Count saved assessments
    $SavedFiles = @(Get-ChildItem -Path $Global:AssessmentDir -Filter '*.json' -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notmatch '_autosave|^discovery_' })
    if ($SavedFiles.Count -ge 5)  { Unlock-Achievement 'five_assessments' }
    if ($SavedFiles.Count -ge 10) { Unlock-Achievement 'ten_assessments' }

    # Full sweep — all checks assessed
    $NotAssessed = @($Checks | Where-Object { $_.Status -eq 'Not Assessed' -and -not $_.Excluded })
    if ($NotAssessed.Count -eq 0 -and $Checks.Count -gt 0) { Unlock-Achievement 'full_sweep' }

    # Perfect category — every non-excluded check in the category must be Pass or N/A
    $Categories = @($Checks | ForEach-Object { $_.Category } | Sort-Object -Unique)
    foreach ($Cat in $Categories) {
        $CatChecks = @($Checks | Where-Object { $_.Category -eq $Cat -and -not $_.Excluded })
        $CatNonPerfect = @($CatChecks | Where-Object { $_.Status -ne 'Pass' -and $_.Status -ne 'N/A' })
        if ($CatChecks.Count -ge 3 -and $CatNonPerfect.Count -eq 0) {
            Unlock-Achievement 'perfect_category'
            break
        }
    }

    # Maturity levels
    if ($Global:Assessment.Discovery -and $Global:Assessment.Discovery.Maturity) {
        $Score = $Global:Assessment.Discovery.Maturity.CompositeScore
        if ($Score -ge 75) { Unlock-Achievement 'maturity_managed' }
        if ($Score -ge 90) { Unlock-Achievement 'maturity_optimized' }
    }

    # Zero critical failures
    $CritFail = @($Failed | Where-Object { $_.Severity -eq 'Critical' })
    if ($Assessed.Count -ge 10 -and $CritFail.Count -eq 0) { Unlock-Achievement 'zero_critical' }

    # Pass milestones
    if ($Passed.Count -ge 50)  { Unlock-Achievement 'fifty_pass' }
    if ($Passed.Count -ge 100) { Unlock-Achievement 'hundred_pass' }

    # Notes
    $WithNotes = @($Checks | Where-Object { $_.Notes }).Count
    if ($WithNotes -ge 5) { Unlock-Achievement 'note_taker' }

    # Exclusions
    $Excluded = @($Checks | Where-Object { $_.Excluded }).Count
    if ($Excluded -ge 1) { Unlock-Achievement 'exclusion_expert' }

    # Time checks
    $Hour = (Get-Date).Hour
    if ($Hour -ge 0 -and $Hour -lt 5) { Unlock-Achievement 'night_owl' }
    if ($Hour -ge 5 -and $Hour -lt 7) { Unlock-Achievement 'early_bird' }
    if ((Get-Date).DayOfWeek -in @('Saturday','Sunday')) { Unlock-Achievement 'weekend_warrior' }
}

# ===============================================================================

<#
.SYNOPSIS
    Persists user preferences (theme, window state, auto-save, achievements) to user_prefs.json.
#>
function Save-UserPrefs {
    try {
        $P = [PSCustomObject]@{
            IsLightMode        = $Global:IsLightMode
            DefaultAssessor    = $txtDefaultAssessor.Text
            LastActiveFile     = $Global:ActiveFilePath
            WindowState        = $Window.WindowState.ToString()
            WindowLeft         = $Window.Left
            WindowTop          = $Window.Top
            WindowWidth        = $Window.Width
            WindowHeight       = $Window.Height
            ActivityLogOpen    = $Global:ActivityLogOpen
            AutoSaveEnabled    = $Global:AutoSaveEnabled
            AutoSaveInterval   = $Global:AutoSaveInterval
            ExportPath         = $Global:ReportDir
            OpenAfterExport    = $Global:OpenAfterExport
            VerboseLogging     = $Global:VerboseLogging
            MaxBackups         = $Global:MaxBackups
            Achievements       = $Global:Achievements
            AchievementsCollapsed = ($Window.FindName('pnlAchievements').Visibility -eq 'Collapsed')
        }
        $P | ConvertTo-Json | Set-Content $Global:PrefsPath -Force
        Write-DebugLog "User preferences saved" -Level 'DEBUG'
    } catch {
        Write-DebugLog "Failed to save preferences: $($_.Exception.Message)" -Level 'WARN'
    }
}

<#
.SYNOPSIS
    Loads user preferences from user_prefs.json and applies them to UI controls.
.DESCRIPTION
    Restores theme, window position/size, default assessor name, auto-save settings,
    export preferences, verbose logging, achievement state, and last active assessment path.
#>
function Load-UserPrefs {
    if (-not (Test-Path $Global:PrefsPath)) { return }
    try {
        $P = Get-Content $Global:PrefsPath -Raw | ConvertFrom-Json

        if ($null -ne $P.IsLightMode) {
            $Global:SuppressThemeHandler = $true
            $Global:IsLightMode = $P.IsLightMode
            & $ApplyTheme $P.IsLightMode
            $btnThemeToggle.Content = if ($P.IsLightMode) { [char]0x263E } else { [char]0x2600 }
            if ($chkDarkMode) { $chkDarkMode.IsChecked = -not $P.IsLightMode }
            $Global:SuppressThemeHandler = $false
        }
        if ($null -ne $P.DefaultAssessor -and $P.DefaultAssessor) {
            $txtDefaultAssessor.Text = $P.DefaultAssessor
            $txtAssessorName.Text = $P.DefaultAssessor
        }

        # Window position and state
        if ($null -ne $P.WindowState) {
            if ($P.WindowState -eq 'Maximized') {
                $Window.WindowState = 'Maximized'
            } elseif ($null -ne $P.WindowLeft) {
                $VW = [System.Windows.SystemParameters]::VirtualScreenWidth
                $VH = [System.Windows.SystemParameters]::VirtualScreenHeight
                $VL = [System.Windows.SystemParameters]::VirtualScreenLeft
                $VT = [System.Windows.SystemParameters]::VirtualScreenTop
                if ($P.WindowLeft -gt ($VL + $VW - 100) -or ($P.WindowLeft + $P.WindowWidth) -lt ($VL + 100) -or
                    $P.WindowTop -gt ($VT + $VH - 100) -or ($P.WindowTop + $P.WindowHeight) -lt ($VT + 100)) {
                    $Window.WindowStartupLocation = 'CenterScreen'
                } else {
                    $Window.WindowState = 'Normal'
                    $Window.Left   = $P.WindowLeft
                    $Window.Top    = $P.WindowTop
                    $Window.Width  = $P.WindowWidth
                    $Window.Height = $P.WindowHeight
                }
            }
        }

        # Restore last active assessment file path
        if ($null -ne $P.LastActiveFile -and $P.LastActiveFile -and (Test-Path $P.LastActiveFile)) {
            $Global:ActiveFilePath = $P.LastActiveFile
        }

        # Restore activity log state
        if ($null -ne $P.ActivityLogOpen -and $P.ActivityLogOpen -eq $false) {
            Toggle-ActivityLog $false
            $chkShowActivityLog.IsChecked = $false
        }

        # Auto-save settings
        if ($null -ne $P.AutoSaveEnabled) {
            $Global:AutoSaveEnabled = $P.AutoSaveEnabled
            $chkAutoSave.IsChecked = $P.AutoSaveEnabled
        }
        if ($null -ne $P.AutoSaveInterval -and $P.AutoSaveInterval -gt 0) {
            $Global:AutoSaveInterval = [int]$P.AutoSaveInterval
            # Select matching combo item
            for ($i = 0; $i -lt $cmbAutoSaveInterval.Items.Count; $i++) {
                if ($cmbAutoSaveInterval.Items[$i].Content -eq "$($Global:AutoSaveInterval)s") {
                    $cmbAutoSaveInterval.SelectedIndex = $i; break
                }
            }
        }

        # Export settings
        if ($null -ne $P.ExportPath -and $P.ExportPath -and (Test-Path $P.ExportPath)) {
            $Global:ReportDir = $P.ExportPath
            $txtExportPath.Text = $P.ExportPath
        }
        if ($null -ne $P.OpenAfterExport) {
            $Global:OpenAfterExport = $P.OpenAfterExport
            $chkOpenAfterExport.IsChecked = $P.OpenAfterExport
        }

        # Verbose logging
        if ($null -ne $P.VerboseLogging) {
            $Global:VerboseLogging = $P.VerboseLogging
            $chkVerboseLog.IsChecked = $P.VerboseLogging
        }

        # Max backups
        if ($null -ne $P.MaxBackups -and $P.MaxBackups -gt 0) {
            $Global:MaxBackups = [int]$P.MaxBackups
            for ($i = 0; $i -lt $cmbMaxBackups.Items.Count; $i++) {
                if ($cmbMaxBackups.Items[$i].Content -eq "$($Global:MaxBackups)") {
                    $cmbMaxBackups.SelectedIndex = $i; break
                }
            }
        }

        # Achievements
        if ($null -ne $P.Achievements) {
            $Global:Achievements = @{}
            $P.Achievements.PSObject.Properties | ForEach-Object {
                $Global:Achievements[$_.Name] = $_.Value
            }
            Update-AchievementBadges
        }

        # Achievement sidebar collapsed state
        if ($null -ne $P.AchievementsCollapsed -and $P.AchievementsCollapsed -eq $true) {
            $Window.FindName('pnlAchievements').Visibility = 'Collapsed'
            $Window.FindName('lblAchievementChevron').Text = [char]0xE76C  # right chevron
        } else {
            $Window.FindName('pnlAchievements').Visibility = 'Visible'
            $Window.FindName('lblAchievementChevron').Text = [char]0xE70D  # down chevron
        }

        Write-DebugLog "User preferences loaded" -Level 'DEBUG'
    } catch {
        Write-DebugLog "Failed to load preferences: $($_.Exception.Message)" -Level 'WARN'
    }
}

# ===============================================================================
# SECTION 16: EVENT WIRING
# ===============================================================================

# Achievement sidebar collapse/expand toggle
$Window.FindName('pnlAchievementHeader').Add_MouseLeftButtonUp({
    $pnlAch = $Window.FindName('pnlAchievements')
    $chevron = $Window.FindName('lblAchievementChevron')
    if ($pnlAch.Visibility -eq 'Visible') {
        $pnlAch.Visibility = 'Collapsed'
        $chevron.Text = [char]0xE76C  # right chevron
    } else {
        $pnlAch.Visibility = 'Visible'
        $chevron.Text = [char]0xE70D  # down chevron
    }
    Save-UserPrefs
})

# Title bar
$btnMinimize.Add_Click({ $Window.WindowState = 'Minimized' })
$btnMaximize.Add_Click({
    $Window.WindowState = if ($Window.WindowState -eq 'Maximized') { 'Normal' } else { 'Maximized' }
})
$btnClose.Add_Click({ $Window.Close() })

# Theme toggle
$btnThemeToggle.Add_Click({
    if ($Global:SuppressThemeHandler) { return }
    $Global:IsLightMode = -not $Global:IsLightMode
    & $ApplyTheme $Global:IsLightMode
    $btnThemeToggle.Content = if ($Global:IsLightMode) { [char]0x263E } else { [char]0x2600 }
    if ($chkDarkMode) {
        $Global:SuppressThemeHandler = $true
        $chkDarkMode.IsChecked = -not $Global:IsLightMode
        $Global:SuppressThemeHandler = $false
    }
    # Re-render all dynamic content with hardcoded theme-conditional colors
    Update-AchievementBadges
    Unlock-Achievement 'theme_toggle'
    Render-AssessmentChecks
    Render-Findings
    Update-Dashboard
    Update-ReportPreview
    Refresh-AssessmentList
})

if ($chkDarkMode) {
    $chkDarkMode.Add_Checked({
        if ($Global:SuppressThemeHandler) { return }
        $Global:IsLightMode = $false
        & $ApplyTheme $false
        $btnThemeToggle.Content = [char]0x2600
        Render-AssessmentChecks
        Render-Findings
        Update-Dashboard
        Update-ReportPreview
        Refresh-AssessmentList
    })
    $chkDarkMode.Add_Unchecked({
        if ($Global:SuppressThemeHandler) { return }
        $Global:IsLightMode = $true
        & $ApplyTheme $true
        $btnThemeToggle.Content = [char]0x263E
        Render-AssessmentChecks
        Render-Findings
        Update-Dashboard
        Update-ReportPreview
        Refresh-AssessmentList
    })
}

# ── Settings controls ──

# Auto-save toggle
$chkAutoSave.Add_Checked({
    $Global:AutoSaveEnabled = $true
    $Global:AutoSaveTimer.Start()
    Write-DebugLog "Auto-save enabled" -Level 'INFO'
})
$chkAutoSave.Add_Unchecked({
    $Global:AutoSaveEnabled = $false
    $Global:AutoSaveTimer.Stop()
    Write-DebugLog "Auto-save disabled" -Level 'INFO'
})

# Auto-save interval
$cmbAutoSaveInterval.Add_SelectionChanged({
    $Sel = $cmbAutoSaveInterval.SelectedItem
    if ($Sel) {
        $Val = [int]($Sel.Content -replace 's','')
        $Global:AutoSaveInterval = $Val
        $Global:AutoSaveTimer.Interval = [TimeSpan]::FromSeconds($Val)
        Write-DebugLog "Auto-save interval set to ${Val}s" -Level 'INFO'
    }
})

# Export path browse
$txtExportPath.Text = $Global:ReportDir
$btnBrowseExportPath.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Select default export folder'
    $dlg.SelectedPath = $Global:ReportDir
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $Global:ReportDir = $dlg.SelectedPath
        $txtExportPath.Text = $dlg.SelectedPath
        Write-DebugLog "Export path set to: $($dlg.SelectedPath)" -Level 'INFO'
    }
})

# Open after export
$chkOpenAfterExport.Add_Checked({  $Global:OpenAfterExport = $true })
$chkOpenAfterExport.Add_Unchecked({ $Global:OpenAfterExport = $false })

# Show activity log (syncs with Toggle-ActivityLog)
$chkShowActivityLog.Add_Checked({  Toggle-ActivityLog $true })
$chkShowActivityLog.Add_Unchecked({ Toggle-ActivityLog $false })

# Verbose logging
$chkVerboseLog.Add_Checked({
    $Global:VerboseLogging = $true
    Write-DebugLog "Verbose logging enabled" -Level 'INFO'
})
$chkVerboseLog.Add_Unchecked({
    $Global:VerboseLogging = $false
    Write-DebugLog "Verbose logging disabled" -Level 'INFO'
})

# Max backups
$cmbMaxBackups.Add_SelectionChanged({
    $Sel = $cmbMaxBackups.SelectedItem
    if ($Sel) {
        $Global:MaxBackups = [int]$Sel.Content
        Write-DebugLog "Max backups set to $($Global:MaxBackups)" -Level 'INFO'
    }
})

# Reset all settings
$btnResetPrefs.Add_Click({
    $Result = Show-ThemedDialog `
        -Title 'Reset Settings' `
        -Message "This will reset all settings to their default values.`nYour assessment data will not be affected." `
        -Icon ([char]0xE7BA) `
        -IconColor 'ThemeWarning' `
        -Buttons @(
            @{ Text = 'Reset'; Result = 'Yes'; IsAccent = $true }
            @{ Text = 'Cancel'; Result = 'No' }
        )
    if ($Result -ne 'Yes') { return }

    # Apply defaults
    $Global:SuppressThemeHandler = $true
    $Global:IsLightMode = $false; & $ApplyTheme $false
    $btnThemeToggle.Content = [char]0x2600
    $chkDarkMode.IsChecked = $true
    $Global:SuppressThemeHandler = $false

    $txtDefaultAssessor.Text = ''
    $txtNotes.Text = ''
    $Global:AutoSaveEnabled = $true; $chkAutoSave.IsChecked = $true
    $Global:AutoSaveInterval = 60; $cmbAutoSaveInterval.SelectedIndex = 1
    $Global:AutoSaveTimer.Interval = [TimeSpan]::FromSeconds(60); $Global:AutoSaveTimer.Start()
    $Global:ReportDir = Join-Path $Global:Root 'reports'; $txtExportPath.Text = $Global:ReportDir
    $Global:OpenAfterExport = $true; $chkOpenAfterExport.IsChecked = $true
    $Global:VerboseLogging = $false; $chkVerboseLog.IsChecked = $false
    $Global:MaxBackups = 10; $cmbMaxBackups.SelectedIndex = 1
    Toggle-ActivityLog $true; $chkShowActivityLog.IsChecked = $true

    Render-AssessmentChecks; Render-Findings; Update-Dashboard; Update-ReportPreview; Refresh-AssessmentList
    Write-DebugLog "All settings reset to defaults" -Level 'INFO'
    Show-Toast 'Settings restored to defaults' -Type 'Info'
})

# Icon rail navigation
$railDashboard.Add_Click({  Switch-Tab 'Dashboard' })
$railAssessment.Add_Click({ Switch-Tab 'Assessment' })
$railFindings.Add_Click({   Switch-Tab 'Findings' })
$railReport.Add_Click({     Switch-Tab 'Report' })
$railSettings.Add_Click({   Switch-Tab 'Settings' })

# Sidebar buttons
$btnSaveAssessment.Add_Click({ Save-Assessment; Refresh-AssessmentList })
$btnLoadAssessment.Add_Click({ Load-Assessment; Refresh-AssessmentList })
$btnNewAssessment = $Window.FindName("btnNewAssessment")
$btnNewAssessment.Add_Click({ Reset-Assessment; Refresh-AssessmentList })

# Quick save in title bar
$btnQuickSave.Add_Click({ Save-Assessment; Refresh-AssessmentList })

# Saved assessments list — double-click to load
$lstSavedAssessments.Add_MouseDoubleClick({
    $Sel = $lstSavedAssessments.SelectedItem
    if ($Sel -and $Sel.Tag) {
        if (-not (Confirm-OverwriteAssessment -Action 'loading a saved file')) { return }

        try {
            $Json = Get-Content $Sel.Tag -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($Json.SchemaVersion -and $Json.Inventory) {
                Import-DiscoveryJson -Path $Sel.Tag
            } elseif ($Json.Checks) {
                $Global:Assessment = $Json
                $Global:Assessment.Checks = [System.Collections.ArrayList]@($Json.Checks)
                # Merge new checks
                $ExIds = @($Json.Checks | ForEach-Object { $_.Id })
                foreach ($Def in $Global:CheckDefinitions) {
                    if ($Def.id -notin $ExIds) {
                        [void]$Global:Assessment.Checks.Add([PSCustomObject]@{
                            Id=$Def.id;Category=$Def.category;Name=$Def.name;Description=$Def.description
                            Severity=$Def.severity;Type=$Def.type;Weight=[int]$Def.weight;Reference=$Def.reference
                            Origin=$Def.origin;Status='Not Assessed';Excluded=$false;Details='';Notes='';Source=$Def.type
                        })
                    }
                }
                Sync-CheckDefinitions
                $txtCustomerName.Text   = $Json.CustomerName
                $txtAssessorName.Text   = $Json.AssessorName
                $txtAssessmentDate.Text = $Json.Date
                $txtNotes.Text          = $Json.Notes
                if ($Json.Discovery) {
                    $lblStatHostPools.Text    = "$($Json.Discovery.Inventory.HostPools.Count)"
                    $lblStatSessionHosts.Text = "$($Json.Discovery.Inventory.SessionHosts.Count)"
                    $lblStatAppGroups.Text    = "$($Json.Discovery.Inventory.AppGroups.Count)"
                    $lblStatWorkspaces.Text   = "$($Json.Discovery.Inventory.Workspaces.Count)"
                    $lblStatScalingPlans.Text = "$($Json.Discovery.Inventory.ScalingPlans.Count)"
                }
                Update-Dashboard
                Update-Progress
                $Global:ActiveFilePath = $Sel.Tag
                Clear-Dirty
                Show-Toast "Loaded: $(Split-Path $Sel.Tag -Leaf)" -Type 'Success'
            }
        } catch {
            Show-Toast "Load failed: $($_.Exception.Message)" -Type 'Error'
        }
    }
})

# Export buttons
$btnExportHtml.Add_Click({ Export-HtmlReport })
$btnExportCsv.Add_Click({ Export-CsvReport })
$btnExportJson.Add_Click({ Export-JsonAssessment })

# Copy summary to clipboard (RTF + plain text)
$btnCopySummary.Add_Click({
    $Assessed = @($Global:Assessment.Checks | Where-Object { $_.Status -ne 'Not Assessed' }).Count
    if ($Assessed -eq 0) {
        Show-Toast 'No assessed checks — import discovery or assess checks first' -Type 'Warning'
        return
    }
    try {
        $plainText = Get-ExecutiveSummary
        $rtfText   = Get-ExecutiveSummaryRtf
        $dataObj   = New-Object System.Windows.DataObject
        $dataObj.SetData([System.Windows.DataFormats]::Rtf, $rtfText)
        $dataObj.SetData([System.Windows.DataFormats]::UnicodeText, $plainText)
        [System.Windows.Clipboard]::SetDataObject($dataObj, $true)
        Show-Toast 'Executive summary copied to clipboard (rich text)' -Type 'Success'
    } catch {
        Show-Toast "Copy failed: $($_.Exception.Message)" -Type 'Error'
    }
})

# Filter changes trigger re-render
$cmbFilterSeverity.Add_SelectionChanged({ Render-Findings })
$cmbFilterStatus.Add_SelectionChanged({ Render-Findings })
$cmbFilterCategory.Add_SelectionChanged({ Render-Findings })
$txtFindingsSearch.Add_TextChanged({ Render-Findings })

# Clear log
$btnClearLog.Add_Click({
    $paraLog.Inlines.Clear()
    $Global:DebugLineCount = 0
    Write-DebugLog "Log cleared" -Level 'INFO'
})

# Activity log toggle
$btnHideLog.Add_Click({ Toggle-ActivityLog $false })
$railActivityLog.Add_Click({ Toggle-ActivityLog (-not $Global:ActivityLogOpen) })

# Populate category filter in findings sidebar
$cmbFilterCategory.Items.Clear()
$AllItem = New-Object System.Windows.Controls.ComboBoxItem; $AllItem.Content = 'All'
[void]$cmbFilterCategory.Items.Add($AllItem)
foreach ($Cat in (Get-Categories)) {
    $Item = New-Object System.Windows.Controls.ComboBoxItem; $Item.Content = $Cat
    [void]$cmbFilterCategory.Items.Add($Item)
}
$cmbFilterCategory.SelectedIndex = 0

# ===============================================================================
# SECTION 17: BACKGROUND JOB POLLING TIMER
# ===============================================================================

$Timer = New-Object System.Windows.Threading.DispatcherTimer
$Timer.Interval = [TimeSpan]::FromMilliseconds($Script:TIMER_INTERVAL_MS)
$Timer.Add_Tick({
    if ($Global:TimerProcessing) { return }
    $Global:TimerProcessing = $true
    try {
        # Process completed background jobs
        $Completed = @($Global:BgJobs | Where-Object { $_.Handle.IsCompleted })
        foreach ($Job in $Completed) {
            try {
                $Results = $Job.PS.EndInvoke($Job.Handle)
                $Errors  = $Job.PS.Streams.Error
                if ($Job.OnComplete) {
                    & $Job.OnComplete $Results $Errors $Job.Context
                }
            } catch {
                Write-DebugLog "BgJob [$($Job.Name)] error: $($_.Exception.Message)" -Level 'ERROR'
            } finally {
                $Job.PS.Dispose()
                $Job.RS.Dispose()
                [void]$Global:BgJobs.Remove($Job)
            }
        }

        # Process sync queue messages
        while ($Global:SyncHash.LogQueue.Count -gt 0) {
            $Msg = $Global:SyncHash.LogQueue.Dequeue()
            Write-DebugLog $Msg.Message -Level $Msg.Level
        }
        while ($Global:SyncHash.StatusQueue.Count -gt 0) {
            $null = $Global:SyncHash.StatusQueue.Dequeue()
        }
    } finally {
        $Global:TimerProcessing = $false
    }
})
$Timer.Start()

# ===============================================================================
# SECTION 18: INITIALIZATION
# ===============================================================================

Write-DebugLog "$Global:AppTitle starting..." -Level 'INFO'
Write-DebugLog "[INIT] Root: $Global:Root" -Level 'DEBUG'

# Set initial values
$txtAssessmentDate.Text = (Get-Date -Format 'yyyy-MM-dd')
$lblSplashVersion.Text = "v$Global:AppVersion"

# Splash step 
Update-AchievementBadges
$dotStep1.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, 'ThemeSuccess')
$lblStep1.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
$lblSplashStatus.Text = "Loading assemblies..."

# Load preferences
Load-UserPrefs

# Splash step 2
$dotStep2.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, 'ThemeSuccess')
$lblStep2.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
$lblSplashStatus.Text = "Checking prerequisites..."

# Splash step 3
$dotStep3.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, 'ThemeSuccess')
$lblStep3.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
$lblSplashStatus.Text = "Ready!"

# Dismiss splash immediately (initialization is fast, no need for timer)
$pnlSplash.Visibility = 'Collapsed'
Write-DebugLog "Splash dismissed, ready for assessment" -Level 'INFO'

# Restore last session: prefer active file from prefs, fall back to _autosave.json
$RestorePath = $null
if ($Global:ActiveFilePath -and (Test-Path $Global:ActiveFilePath)) {
    $RestorePath = $Global:ActiveFilePath
    Write-DebugLog "Restoring last active profile: $(Split-Path $RestorePath -Leaf)" -Level 'DEBUG'
} elseif (Test-Path $Global:AutoSaveFile) {
    $RestorePath = $Global:AutoSaveFile
    Write-DebugLog "Restoring from autosave" -Level 'DEBUG'
}

if ($RestorePath) {
    try {
        $RestoredData = Get-Content $RestorePath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($RestoredData.Checks -and $RestoredData.Checks.Count -gt 0) {
            $HasWork = @($RestoredData.Checks | Where-Object { $_.Status -ne 'Not Assessed' }).Count
            if ($HasWork -gt 0) {
                $Global:Assessment = $RestoredData
                $Global:Assessment.Checks = [System.Collections.ArrayList]@($RestoredData.Checks)

                # Merge: add any new checks from checks.json that don't exist in restored data
                $ExistingIds = @($Global:Assessment.Checks | ForEach-Object { $_.Id })
                $Added = 0
                foreach ($Def in $Global:CheckDefinitions) {
                    if ($Def.id -notin $ExistingIds) {
                        [void]$Global:Assessment.Checks.Add([PSCustomObject]@{
                            Id=$Def.id; Category=$Def.category; Name=$Def.name; Description=$Def.description
                            Severity=$Def.severity; Type=$Def.type; Weight=[int]$Def.weight; Reference=$Def.reference
                            Origin=$Def.origin; Status='Not Assessed'; Excluded=$false; Details=''; Notes=''; Source=$Def.type; Effort=$Def.effort
                        })
                        $Added++
                    }
                }
                Sync-CheckDefinitions

                if ($RestoredData.CustomerName) { $txtCustomerName.Text = $RestoredData.CustomerName }
                if ($RestoredData.AssessorName) { $txtAssessorName.Text = $RestoredData.AssessorName }
                if ($RestoredData.Date)         { $txtAssessmentDate.Text = $RestoredData.Date }
                if ($RestoredData.Discovery) {
                    $lblStatHostPools.Text    = "$($RestoredData.Discovery.Inventory.HostPools.Count)"
                    $lblStatSessionHosts.Text = "$($RestoredData.Discovery.Inventory.SessionHosts.Count)"
                    $lblStatAppGroups.Text    = "$($RestoredData.Discovery.Inventory.AppGroups.Count)"
                    $lblStatWorkspaces.Text   = "$($RestoredData.Discovery.Inventory.Workspaces.Count)"
                    $lblStatScalingPlans.Text = "$($RestoredData.Discovery.Inventory.ScalingPlans.Count)"
                }
                $SourceLabel = if ($RestorePath -eq $Global:AutoSaveFile) { 'autosave' } else { Split-Path $RestorePath -Leaf }
                Write-DebugLog "Restored: $SourceLabel ($HasWork checks$(if ($Added -gt 0) { ", +$Added new" }))" -Level 'SUCCESS'
                Show-Toast "Restored: $SourceLabel" -Type 'Info'
            }
        }
    } catch {
        Write-DebugLog "Restore failed: $($_.Exception.Message)" -Level 'WARN'
    }
}

# Enable autosave now that startup is complete
$Global:AutoSaveEnabled = $true

# Periodic autosave timer
$Global:AutoSaveTimer = New-Object System.Windows.Threading.DispatcherTimer
$Global:AutoSaveTimer.Interval = [TimeSpan]::FromSeconds($Global:AutoSaveInterval)
$Global:AutoSaveTimer.Add_Tick({ AutoSave-Assessment })
if ($Global:AutoSaveEnabled) { $Global:AutoSaveTimer.Start() }
Write-DebugLog "Periodic autosave enabled ($($Global:AutoSaveInterval)s interval)" -Level 'DEBUG'

# Initialize dashboard and saved assessments list
Refresh-AssessmentList
Switch-Tab 'Dashboard'

# Dispatcher exception handler
$Window.Dispatcher.Add_UnhandledException({
    param($sender, $e)
    $Ex = $e.Exception
    $Inner = if ($Ex.InnerException) { $Ex.InnerException } else { $Ex }
    Write-DebugLog "DISPATCHER UNHANDLED: $($Inner.GetType().FullName): $($Inner.Message)" -Level 'ERROR'
    Write-DebugLog "  Stack: $($Inner.StackTrace)" -Level 'ERROR'
    $e.Handled = $false
})

# Window close — save to active profile
$Window.Add_Closing({
    param($sender, $e)

    # Always autosave to _autosave.json (safety net)
    AutoSave-Assessment

    # If dirty and there's an active file, save silently to it
    if ($Global:IsDirty -and $Global:ActiveFilePath) {
        try {
            Sync-AssessmentFromUI
            $Global:Assessment | ConvertTo-Json -Depth 10 | Set-Content $Global:ActiveFilePath -Encoding UTF8 -Force
            Write-DebugLog "Auto-saved to active profile: $Global:ActiveFilePath" -Level 'SUCCESS'
        } catch { }
    }
    # If dirty but no active file, warn
    elseif ($Global:IsDirty -and (Test-AssessmentDirty)) {
        $Result = Show-ThemedDialog `
            -Title 'Unsaved Assessment' `
            -Message "You have unsaved changes. Save before closing?" `
            -Icon ([char]0xE7BA) `
            -IconColor 'ThemeWarning' `
            -Buttons @(
                @{ Text = 'Save'; Result = 'Save'; IsAccent = $true }
                @{ Text = 'Discard'; Result = 'Discard' }
                @{ Text = 'Cancel'; Result = 'Cancel' }
            )
        switch ($Result) {
            'Save'    { Save-Assessment }
            'Cancel'  { $e.Cancel = $true; return }
        }
    }

    Write-DebugLog "[UI] Window closing" -Level 'INFO'
    Save-UserPrefs
})

# Bring window to front
$Window.Topmost = $true
$Window.Add_ContentRendered({
    $Window.Topmost = $false
    $Window.Activate()
    # P/Invoke SetForegroundWindow for reliable foreground grab
    try {
        $WIH = New-Object System.Windows.Interop.WindowInteropHelper($Window)
        [ForegroundHelper]::SetForegroundWindow($WIH.Handle) | Out-Null
    } catch { }
})

# Show window (blocks until closed)
$Window.ShowDialog() | Out-Null

# Cleanup
$Timer.Stop()
$Global:AutoSaveTimer.Stop()
Write-DebugLog "Application closed" -Level 'INFO'



