<#
.SYNOPSIS
    AIB Packer Log Monitor - Real-time Azure Image Builder log viewer.

.DESCRIPTION
    A WPF-based GUI tool that authenticates against Azure using Entra ID,
    discovers running Azure Image Builder (AIB) builds, and streams the
    Packer container log in real-time with color-coded output and auto-scroll.

    ARCHITECTURE:
    Uses the same proven threading model as Bagel Commander:
    1. GUI Thread (STA): Renders WPF, handles user input.
    2. Worker Runspace (Background): Polls Azure APIs for log updates.
    3. Synchronized Queue: Thread-safe bridge between worker and GUI.
    4. DispatcherTimer (50ms): Drains the queue and updates the RichTextBox.

.NOTES
    Version:        1.0.0
    Author:         Anton Romanyuk
    Creation Date:  03/06/2026
    Dependencies:   Az.Accounts, Az.Resources, Az.ContainerInstance, Az.Storage

.EXAMPLE
    .\AIBLogMonitor.ps1
    # Launches the GUI. Sign in with your Entra ID account, select a build, and stream logs.
#>

# ==============================================================================
# SECTION 1: PRE-LOAD & INITIALIZATION
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'

$Global:Root = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Global:Root)) { $Global:Root = $PWD.Path }

$Global:AppVersion = "0.1.0-alpha"
$Global:AppTitle = "AIB Packer Log Monitor v$($Global:AppVersion)"

# Load WPF assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# DPI Awareness: Enable Per-Monitor V2 for crisp text on high-DPI displays
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
} catch { <# Already set or OS doesn't support #> }

# ==============================================================================
# SECTION 2: PREREQUISITE MODULE CHECK (No Auto-Install)
# ==============================================================================

$RequiredModules = @(
    @{ Name = 'Az.Accounts';          MinVersion = '2.0.0'; Purpose = 'Azure authentication (Connect-AzAccount)' }
    @{ Name = 'Az.Resources';         MinVersion = '6.0.0'; Purpose = 'Resource group discovery (Get-AzResourceGroup)' }
    @{ Name = 'Az.ContainerInstance'; MinVersion = '3.0.0'; Purpose = 'Container logs (Get-AzContainerInstanceLog)' }
    @{ Name = 'Az.Storage';           MinVersion = '5.0.0'; Purpose = 'Full Packer log from blob storage (packerlogs)' }
)

function Test-Prerequisites {
    <#
    .SYNOPSIS
        Checks if all required Az modules are installed. Returns a results array.
    #>
    $Results = @()
    foreach ($Mod in $RequiredModules) {
        $Installed = Get-Module -ListAvailable -Name $Mod.Name -ErrorAction SilentlyContinue |
                     Sort-Object Version -Descending | Select-Object -First 1
        $Results += @{
            Name       = $Mod.Name
            Required   = $Mod.MinVersion
            Purpose    = $Mod.Purpose
            Installed  = if ($Installed) { $Installed.Version.ToString() } else { $null }
            Available  = [bool]$Installed
        }
    }
    return $Results
}

function Get-RemediationCommand {
    <#
    .SYNOPSIS
        Generates a copy-pasteable Install-Module command for all missing modules.
    #>
    param([array]$ModuleResults)
    $Missing = $ModuleResults | Where-Object { -not $_.Available }
    if ($Missing.Count -eq 0) { return "" }
    $Names = ($Missing | ForEach-Object { $_.Name }) -join "', '"
    return "Install-Module -Name '$Names' -Scope CurrentUser -Force -AllowClobber"
}

# ==============================================================================
# SECTION 3: THREAD SYNCHRONIZATION BRIDGE
# ==============================================================================

$Global:SyncHash = [Hashtable]::Synchronized(@{
    StopFlag        = $false
    IsStreaming      = $false
    LogQueue        = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())
    BuildQueue      = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())
    SelectedBuild   = $null
    PollIntervalSec = 5
    LastLogLength   = 0
    TotalLines      = 0
    BuildStartTime  = $null
})

# Background job tracker — polled by the DispatcherTimer to avoid UI-thread hangs
$Global:BgJobs = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())

function Start-BackgroundWork {
    <#
    .SYNOPSIS
        Runs a scriptblock in a background STA runspace and tracks it for timer-based completion polling.
    .DESCRIPTION
        The DispatcherTimer tick handler checks $Global:BgJobs every 50ms.
        When the async result is complete, it invokes OnComplete on the UI thread
        with the results and any error stream, then disposes the runspace.
    #>
    param(
        [ScriptBlock]$Work,
        [ScriptBlock]$OnComplete,
        [hashtable]$Variables = @{}
    )
    $ISS = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $ISS.ExecutionPolicy = [Microsoft.PowerShell.ExecutionPolicy]::Bypass
    $RS  = [RunspaceFactory]::CreateRunspace($ISS)
    $RS.ApartmentState = 'STA'
    $RS.ThreadOptions  = 'ReuseThread'
    $RS.Open()

    $PS = [PowerShell]::Create()
    $PS.Runspace = $RS
    foreach ($k in $Variables.Keys) {
        $PS.Runspace.SessionStateProxy.SetVariable($k, $Variables[$k])
    }
    $PS.AddScript($Work) | Out-Null

    Write-DebugLog "BgWork: launching runspace (state=$($RS.RunspaceStateInfo.State), vars=$($Variables.Keys -join ','))"
    $Async = $PS.BeginInvoke()
    $Global:BgJobs.Add(@{
        PS          = $PS
        Runspace    = $RS
        AsyncResult = $Async
        OnComplete  = $OnComplete
        StartedAt   = (Get-Date)
    }) | Out-Null
    Write-DebugLog "BgWork: queued job #$($Global:BgJobs.Count)"
}

# ==============================================================================
# SECTION 4: XAML GUI LOAD
# ==============================================================================

$XamlPath = Join-Path $Global:Root "AIBLogMonitor_UI.xaml"
if (-not (Test-Path $XamlPath)) {
    [System.Windows.MessageBox]::Show(
        "CRITICAL: 'AIBLogMonitor_UI.xaml' not found in:`n$Global:Root",
        "AIB Log Monitor", 'OK', 'Error') | Out-Null
    exit 1
}

$XamlContent = Get-Content $XamlPath -Raw -Encoding UTF8
$XamlContent = $XamlContent -replace 'Title="AIB Packer Log Monitor"', "Title=`"$Global:AppTitle`""

# Parse XAML
try {
    $Window = [Windows.Markup.XamlReader]::Parse($XamlContent)
} catch {
    [System.Windows.MessageBox]::Show(
        "XAML Parse Error:`n$($_.Exception.Message)",
        "AIB Log Monitor", 'OK', 'Error') | Out-Null
    exit 1
}

# ==============================================================================
# SECTION 5: ELEMENT REFERENCES
# ==============================================================================

# Title bar
$btnThemeToggle  = $Window.FindName("btnThemeToggle")
$btnMinimize     = $Window.FindName("btnMinimize")
$btnMaximize     = $Window.FindName("btnMaximize")
$btnClose        = $Window.FindName("btnClose")
$lblTitle        = $Window.FindName("lblTitle")
$lblTitleVersion = $Window.FindName("lblTitleVersion")

# Auth bar
$authDot         = $Window.FindName("authDot")
$lblAuthStatus   = $Window.FindName("lblAuthStatus")
$cmbSubscription = $Window.FindName("cmbSubscription")
$btnLogin        = $Window.FindName("btnLogin")
$btnLogout       = $Window.FindName("btnLogout")

# Build list
$lstBuilds         = $Window.FindName("lstBuilds")
$lblBuildCount     = $Window.FindName("lblBuildCount")
$btnRefreshBuilds  = $Window.FindName("btnRefreshBuilds")
$buildScanDot      = $Window.FindName("buildScanDot")
$pnlEmptyState     = $Window.FindName("pnlEmptyState")
$lblEmptyTitle     = $Window.FindName("lblEmptyTitle")
$lblEmptySubtitle  = $Window.FindName("lblEmptySubtitle")
$chkAutoRefresh    = $Window.FindName("chkAutoRefresh")
$txtPollInterval   = $Window.FindName("txtPollInterval")

# Log viewer
$rtbOutput          = $Window.FindName("rtbOutput")
$lblLogTitle        = $Window.FindName("lblLogTitle")
$lblLogTemplateName = $Window.FindName("lblLogTemplateName")
$txtSearch          = $Window.FindName("txtSearch")
$lblSearchCount     = $Window.FindName("lblSearchCount")
$btnSearchPrev      = $Window.FindName("btnSearchPrev")
$btnSearchNext      = $Window.FindName("btnSearchNext")
$chkWordWrap        = $Window.FindName("chkWordWrap")
$btnClearLog        = $Window.FindName("btnClearLog")
$btnSaveLog         = $Window.FindName("btnSaveLog")
$btnCopyLog         = $Window.FindName("btnCopyLog")

# Phase tracker
$pnlPhaseTracker = $Window.FindName("pnlPhaseTracker")
$lblPhaseEmpty   = $Window.FindName("lblPhaseEmpty")

# Auto-scan builds
$chkAutoScanBuilds = $Window.FindName("chkAutoScanBuilds")

# Stats dashboard (#9)
$pnlStatsDashboard = $Window.FindName("pnlStatsDashboard")
$lblStatsErrors    = $Window.FindName("lblStatsErrors")
$lblStatsWarnings  = $Window.FindName("lblStatsWarnings")
$lblStatsSuccess   = $Window.FindName("lblStatsSuccess")
$lblStatsLines     = $Window.FindName("lblStatsLines")
$lblStatsLps       = $Window.FindName("lblStatsLps")

# Tail-follow indicator (#10)
$lblFollowIndicator = $Window.FindName("lblFollowIndicator")

# Build history (#16)
$pnlBuildHistory = $Window.FindName("pnlBuildHistory")

# Flame chart flyout (#7)
$pnlFlameChart       = $Window.FindName("pnlFlameChart")
$btnFlameChartToggle = $Window.FindName("btnFlameChartToggle")
$popFlameChart       = $Window.FindName("popFlameChart")
$lblFlameCount       = $Window.FindName("lblFlameCount")

# Minimap (#20)
$cnvMinimap = $Window.FindName("cnvMinimap")

# Sound toggle (#18)
$chkSoundAlerts = $Window.FindName("chkSoundAlerts")

# Diff button (#12)
$btnDiffLog = $Window.FindName("btnDiffLog")

# Health score (#R2-1)
$lblHealthGrade = $Window.FindName("lblHealthGrade")
$lblHealthScore = $Window.FindName("lblHealthScore")

# ETA & progress (#R2-4)
$lblETAPercent   = $Window.FindName("lblETAPercent")
$lblETARemaining = $Window.FindName("lblETARemaining")
$prgETA          = $Window.FindName("prgETA")

# Package tracker (#R2-5)
$pnlPackageTracker = $Window.FindName("pnlPackageTracker")

# Error clusters (#R2-6)
$pnlErrorClusters = $Window.FindName("pnlErrorClusters")

# Known issues (#R2-8)
$pnlKnownIssues = $Window.FindName("pnlKnownIssues")

# Build streak (#R2-9)
$lblStreak   = $Window.FindName("lblStreak")
$cnvConfetti = $Window.FindName("cnvConfetti")

# Achievements (#R2-10)
$pnlAchievements = $Window.FindName("pnlAchievements")

# Script profiler (#R2-11)
$pnlScriptProfiler = $Window.FindName("pnlScriptProfiler")

# Error heatmap (#R2-14)
$cnvHeatmap = $Window.FindName("cnvHeatmap")

# Cost estimator (#R2-15)
$lblCostEstimate = $Window.FindName("lblCostEstimate")
$lblCostRate     = $Window.FindName("lblCostRate")
$lblCostSku      = $Window.FindName("lblCostSku")

# Prereq error panel
$pnlPrereqError   = $Window.FindName("pnlPrereqError")
$pnlModuleStatus  = $Window.FindName("pnlModuleStatus")
$txtPrereqCommand  = $Window.FindName("txtPrereqCommand")
$btnCopyCommand    = $Window.FindName("btnCopyCommand")
$btnRetryPrereq    = $Window.FindName("btnRetryPrereq")
$lblPrereqMessage  = $Window.FindName("lblPrereqMessage")

# Status bar
$statusDot   = $Window.FindName("statusDot")
$lblStatus   = $Window.FindName("lblStatus")
$lblLineCount = $Window.FindName("lblLineCount")
$lblElapsed  = $Window.FindName("lblElapsed")
$lblVersion  = $Window.FindName("lblVersion")

# Build status ticker strip
$lblTickerBuilds     = $Window.FindName("lblTickerBuilds")
$badgeRunning        = $Window.FindName("badgeRunning")

# Background layers (WinGet MM design language)
$bdrDotGrid      = $Window.FindName("bdrDotGrid")
$bdrGradientGlow = $Window.FindName("bdrGradientGlow")
$lblTickerRunning    = $Window.FindName("lblTickerRunning")
$badgeFailed         = $Window.FindName("badgeFailed")
$lblTickerFailed     = $Window.FindName("lblTickerFailed")
$badgeSucceeded      = $Window.FindName("badgeSucceeded")
$lblTickerSucceeded  = $Window.FindName("lblTickerSucceeded")
$lblTickerSub        = $Window.FindName("lblTickerSub")
$lblTickerRegion     = $Window.FindName("lblTickerRegion")
$badgeConnection     = $Window.FindName("badgeConnection")
$lblTickerConnection = $Window.FindName("lblTickerConnection")
$btnScanBuilds       = $null  # Removed — redundant with btnRefreshBuilds

# Settings panel (inline collapsible)
$pnlSettingsHeader  = $Window.FindName("pnlSettingsHeader")
$pnlSettingsBody    = $Window.FindName("pnlSettingsBody")
$lblSettingsChevron = $Window.FindName("lblSettingsChevron")
$chkDarkMode        = $Window.FindName("chkDarkMode")

# Debug overlay
$chkDebug       = $Window.FindName("chkDebug")
$pnlDebugOverlay = $Window.FindName("pnlDebugOverlay")
$txtDebugLog    = $Window.FindName("txtDebugLog")
$lblDebugCount  = $Window.FindName("lblDebugCount")
$btnDebugClear  = $Window.FindName("btnDebugClear")
$btnDebugClose  = $Window.FindName("btnDebugClose")
$debugScroller  = $Window.FindName("debugScroller")

# Right analytics panel & icon rail
$pnlRightSidebar    = $Window.FindName("pnlRightSidebar")
$colRightPanel      = $Window.FindName("colRightPanel")
$colRightSplitter   = $Window.FindName("colRightSplitter")
$splitterRight      = $Window.FindName("splitterRight")
$btnRightHamburger  = $Window.FindName("btnRightHamburger")
$railRStats         = $Window.FindName("railRStats")
$railRProfiler      = $Window.FindName("railRProfiler")
$railRHistory       = $Window.FindName("railRHistory")
$railRStatsIndicator    = $Window.FindName("railRStatsIndicator")
$railRProfilerIndicator = $Window.FindName("railRProfilerIndicator")
$railRHistoryIndicator  = $Window.FindName("railRHistoryIndicator")
$railRStatsIcon     = $Window.FindName("railRStatsIcon")
$railRProfilerIcon  = $Window.FindName("railRProfilerIcon")
$railRHistoryIcon   = $Window.FindName("railRHistoryIcon")
$grpRightStats      = $Window.FindName("grpRightStats")
$grpRightProfiler   = $Window.FindName("grpRightProfiler")
$grpRightHistory    = $Window.FindName("grpRightHistory")

# Left icon rail & collapsible sidebar
$pnlLeftSidebar     = $Window.FindName("pnlLeftSidebar")
$colLeftPanel       = $Window.FindName("colLeftPanel")
$colLeftSplitter    = $Window.FindName("colLeftSplitter")
$splitterLeft       = $Window.FindName("splitterLeft")
$btnHamburger       = $Window.FindName("btnHamburger")
$railBuilds         = $Window.FindName("railBuilds")
$railInfo           = $Window.FindName("railInfo")
$railSettings       = $Window.FindName("railSettings")
$railBuildsIndicator   = $Window.FindName("railBuildsIndicator")
$railInfoIndicator     = $Window.FindName("railInfoIndicator")
$railSettingsIndicator = $Window.FindName("railSettingsIndicator")
$railBuildsIcon     = $Window.FindName("railBuildsIcon")
$railInfoIcon       = $Window.FindName("railInfoIcon")
$railSettingsIcon   = $Window.FindName("railSettingsIcon")
$grpLeftBuildsHeader = $Window.FindName("grpLeftBuildsHeader")
$grpLeftBottomBorder = $Window.FindName("grpLeftBottomBorder")
$grpLeftInfo         = $Window.FindName("grpLeftInfo")
$grpLeftSettings     = $Window.FindName("grpLeftSettings")
$rowLeftHeader       = $Window.FindName("rowLeftHeader")
$rowLeftBuilds       = $Window.FindName("rowLeftBuilds")
$rowLeftBottom       = $Window.FindName("rowLeftBottom")

$lblVersion.Text = "v$($Global:AppVersion)"
$lblTitleVersion.Text = "v$($Global:AppVersion)"

# ==============================================================================
# SECTION 5B: DEBUG LOG FUNCTION
# ==============================================================================

$Global:DebugLineCount = 0
$Global:DebugMaxLines  = 500
$Global:DebugSB        = [System.Text.StringBuilder]::new(4096)

function Write-DebugLog {
    <#
    .SYNOPSIS
        Writes a timestamped line to both the PS console and the debug overlay panel.
    .DESCRIPTION
        Always writes to the PowerShell host console (Write-Host) so hang diagnostics
        are visible even when the GUI is frozen. Also appends to the debug overlay
        when the toggle is checked. Ring-buffers at 500 lines.
    #>
    param([string]$Message)

    $ts   = (Get-Date).ToString('HH:mm:ss.fff')
    $line = "[$ts] $Message"

    # Always write to PS console — visible even when GUI hangs
    Write-Host $line -ForegroundColor DarkGray

    # Overlay panel — only when toggle is on
    if (-not $chkDebug.IsChecked) { return }

    $Global:DebugLineCount++
    if ($Global:DebugLineCount -gt $Global:DebugMaxLines) {
        # Trim the oldest line
        $text = $Global:DebugSB.ToString()
        $nl   = $text.IndexOf("`n")
        if ($nl -ge 0) {
            $Global:DebugSB.Clear()
            $Global:DebugSB.Append($text.Substring($nl + 1)) | Out-Null
            $Global:DebugLineCount--
        }
    }
    $Global:DebugSB.AppendLine($line) | Out-Null
    $txtDebugLog.Text     = $Global:DebugSB.ToString()
    $lblDebugCount.Text   = "$($Global:DebugLineCount) lines"
    $debugScroller.ScrollToEnd()
}

# ==============================================================================
# SECTION 6: THEME ENGINE (Microsoft Color Scheme)
# ==============================================================================

$Global:IsLightMode = $false
$Global:SuppressThemeHandler = $false
$Global:IsAuthenticated = $false

$Global:ThemeDark = @{
    ThemeAppBg         = "#111111"; ThemePanelBg       = "#1A1A1A"
    ThemeCardBg        = "#222222"; ThemeCardAltBg     = "#252528"
    ThemeInputBg       = "#141414"; ThemeDeepBg        = "#0D0D0D"
    ThemeOutputBg      = "#080808"; ThemeSurfaceBg     = "#2A2A2A"
    ThemeHoverBg       = "#333333"; ThemeSelectedBg    = "#2E2E32"
    ThemePressedBg     = "#181818"
    ThemeAccent        = "#0078D4"; ThemeAccentHover    = "#1A8AD4"
    ThemeAccentLight   = "#60CDFF"; ThemeAccentDim     = "#004578"
    ThemeGreenAccent   = "#00C853"
    ThemeTextPrimary   = "#FFFFFF"; ThemeTextBody      = "#E0E0E0"
    ThemeTextSecondary = "#CCCCCC"; ThemeTextMuted     = "#AAAAAA"
    ThemeTextDim       = "#888888"; ThemeTextDisabled  = "#555555"
    ThemeTextFaintest  = "#3A3A3A"
    ThemeBorder        = "#2A2A2A"; ThemeBorderCard    = "#333333"
    ThemeBorderElevated = "#3E3E3E"; ThemeBorderHover   = "#505050"
    ThemeScrollThumb   = "#383838"
    ThemeSuccess       = "#00C853"; ThemeWarning       = "#FFB900"
    ThemeError         = "#D13438"; ThemeSidebarBg     = "#0D0D0D"
    ThemeSidebarBorder = "#222222"
}

$Global:ThemeLight = @{
    ThemeAppBg         = "#F3F3F3"; ThemePanelBg       = "#FAFAFA"
    ThemeCardBg        = "#FFFFFF"; ThemeCardAltBg     = "#F5F5F5"
    ThemeInputBg       = "#FFFFFF"; ThemeDeepBg        = "#EEEEEE"
    ThemeOutputBg      = "#FAFAFA"; ThemeSurfaceBg     = "#E8E8E8"
    ThemeHoverBg       = "#E0E0E0"; ThemeSelectedBg    = "#D8D8D8"
    ThemePressedBg     = "#CCCCCC"
    ThemeAccent        = "#0078D4"; ThemeAccentHover    = "#106EBE"
    ThemeAccentLight   = "#0078D4"; ThemeAccentDim     = "#B3D7F2"
    ThemeGreenAccent   = "#00873D"
    ThemeTextPrimary   = "#1A1A1A"; ThemeTextBody      = "#333333"
    ThemeTextSecondary = "#5C5C5C"; ThemeTextMuted     = "#707070"
    ThemeTextDim       = "#8C8C8C"; ThemeTextDisabled  = "#A0A0A0"
    ThemeTextFaintest  = "#D0D0D0"
    ThemeBorder        = "#E0E0E0"; ThemeBorderCard    = "#D8D8D8"
    ThemeBorderElevated = "#C8C8C8"; ThemeBorderHover   = "#B0B0B0"
    ThemeScrollThumb   = "#C0C0C0"
    ThemeSuccess       = "#0B6A0B"; ThemeWarning       = "#9E6B00"
    ThemeError         = "#A80000"; ThemeSidebarBg     = "#EBEBEB"
    ThemeSidebarBorder = "#D8D8D8"
}

# Remap log colors for light mode readability
$Global:RemapLight = {
    param([string]$C)
    if (-not $Global:IsLightMode) { return $C }
    switch ($C) {
        { $_ -eq "White" -or $_ -eq "#FFFFFF" -or $_ -eq "#E0E0E0" }   { return "#1A1A1A" }
        { $_ -eq "#60CDFF" -or $_ -eq "#50E6FF" -or $_ -eq "Cyan" }   { return "#0066AA" }
        { $_ -eq "#0078D4" }                                             { return "#005A9E" }
        { $_ -eq "#FFB900" -or $_ -eq "Gold" }                          { return "#9E6B00" }
        { $_ -eq "#888888" -or $_ -eq "Gray" }                          { return "#555555" }
        { $_ -eq "#00C853" }                                             { return "#00873D" }
        { $_ -eq "#107C10" -or $_ -eq "Green" }                         { return "#0B6A0B" }
        { $_ -eq "#D13438" -or $_ -eq "Red" }                           { return "#A80000" }
        { $_ -eq "#E74856" }                                             { return "#C42B1C" }
        { $_ -eq "#FF8C00" -or $_ -eq "Orange" }                        { return "#B85C00" }
        { $_ -eq "#3B7D23" }                                             { return "#1A7D1A" }
        { $_ -eq "#8764B8" -or $_ -eq "Purple" }                        { return "#5C2D91" }
        default { return $C }
    }
}

$ApplyTheme = {
    param([bool]$IsLight)
    $Palette = if ($IsLight) { $Global:ThemeLight } else { $Global:ThemeDark }
    $BC = $Global:CachedBC
    
    foreach ($Key in $Palette.Keys) {
        $NewColor = [System.Windows.Media.ColorConverter]::ConvertFromString($Palette[$Key])
        $NewBrush = [System.Windows.Media.SolidColorBrush]::new($NewColor)
        $Window.Resources[$Key] = $NewBrush
    }
    $Global:IsLightMode = $IsLight

    # Rebuild dot grid pattern for theme
    if ($bdrDotGrid) {
        $DotColor = if ($IsLight) { '#14000000' } else { '#1EFFFFFF' }
        $Brush = [System.Windows.Media.DrawingBrush]::new()
        $Brush.TileMode = 'Tile'
        $Brush.Viewport = [System.Windows.Rect]::new(0,0,44,44)
        $Brush.ViewportUnits = 'Absolute'
        $Brush.Viewbox = [System.Windows.Rect]::new(0,0,44,44)
        $Brush.ViewboxUnits = 'Absolute'
        $GD = [System.Windows.Media.GeometryDrawing]::new()
        $GD.Brush = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.ColorConverter]::ConvertFromString($DotColor))
        $GD.Geometry = [System.Windows.Media.EllipseGeometry]::new(
            [System.Windows.Point]::new(22,22), 1.3, 1.3)
        $Brush.Drawing = $GD
        $bdrDotGrid.Background = $Brush
    }
    if ($bdrGradientGlow) {
        $Grad = [System.Windows.Media.LinearGradientBrush]::new()
        $Grad.StartPoint = [System.Windows.Point]::new(0.5, 0)
        $Grad.EndPoint   = [System.Windows.Point]::new(0.5, 1)
        if ($IsLight) {
            $Grad.GradientStops.Add([System.Windows.Media.GradientStop]::new(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#280078D4'), 0.0))
            $Grad.GradientStops.Add([System.Windows.Media.GradientStop]::new(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#180078D4'), 0.3))
            $Grad.GradientStops.Add([System.Windows.Media.GradientStop]::new(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#080078D4'), 0.6))
            $Grad.GradientStops.Add([System.Windows.Media.GradientStop]::new(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#00FFFFFF'), 1.0))
        } else {
            $Grad.GradientStops.Add([System.Windows.Media.GradientStop]::new(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#350078D4'), 0.0))
            $Grad.GradientStops.Add([System.Windows.Media.GradientStop]::new(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#200078D4'), 0.3))
            $Grad.GradientStops.Add([System.Windows.Media.GradientStop]::new(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#100078D4'), 0.6))
            $Grad.GradientStops.Add([System.Windows.Media.GradientStop]::new(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#00000000'), 1.0))
        }
        $bdrGradientGlow.Background = $Grad
        $bdrGradientGlow.Opacity = 1.0
    }

    # Update semi-transparent ticker badge backgrounds (can't use DynamicResource for alpha colors)
    if ($IsLight) {
        if ($badgeRunning)    { $badgeRunning.Background    = $BC.ConvertFromString('#1500873D') }
        if ($badgeFailed)     { $badgeFailed.Background     = $BC.ConvertFromString('#15A80000') }
        if ($badgeSucceeded)  { $badgeSucceeded.Background  = $BC.ConvertFromString('#159E6B00') }
        if ($badgeConnection) {
            if ($Global:IsAuthenticated) {
                $badgeConnection.Background = $BC.ConvertFromString('#1500873D')
            } else {
                $badgeConnection.Background = $BC.ConvertFromString('#15A80000')
            }
        }
    } else {
        if ($badgeRunning)    { $badgeRunning.Background    = $BC.ConvertFromString('#1500C853') }
        if ($badgeFailed)     { $badgeFailed.Background     = $BC.ConvertFromString('#15D13438') }
        if ($badgeSucceeded)  { $badgeSucceeded.Background  = $BC.ConvertFromString('#15FFB900') }
        if ($badgeConnection) {
            if ($Global:IsAuthenticated) {
                $badgeConnection.Background = $BC.ConvertFromString('#1500C853')
            } else {
                $badgeConnection.Background = $BC.ConvertFromString('#15D13438')
            }
        }
    }

    # Repaint existing terminal text
    if ($rtbOutput -and $rtbOutput.Document) {
        if ($IsLight) {
            $TermMap = @{
                '#FFFFFFFF'='#1A1A1A'; '#FFE0E0E0'='#333333'
                '#FF60CDFF'='#0066AA'; '#FF50E6FF'='#0066AA'; '#FF0078D4'='#005A9E'
                '#FFFFB900'='#9E6B00'; '#FF888888'='#555555'
                '#FF107C10'='#0B6A0B'; '#FF00C853'='#00873D'; '#FFD13438'='#A80000'
                '#FFFF8C00'='#B85C00'; '#FF8764B8'='#5C2D91'
            }
        } else {
            $TermMap = @{
                '#FF1A1A1A'='#FFFFFF'; '#FF333333'='#E0E0E0'
                '#FF0066AA'='#60CDFF'; '#FF005A9E'='#0078D4'
                '#FF9E6B00'='#FFB900'; '#FF555555'='#888888'
                '#FF0B6A0B'='#107C10'; '#FF00873D'='#00C853'; '#FFA80000'='#D13438'
                '#FFB85C00'='#FF8C00'; '#FF5C2D91'='#8764B8'
            }
        }
        foreach ($Block in $rtbOutput.Document.Blocks) {
            if ($Block -is [System.Windows.Documents.Paragraph]) {
                foreach ($Inline in $Block.Inlines) {
                    if ($Inline -is [System.Windows.Documents.Run] -and $Inline.Foreground -is [System.Windows.Media.SolidColorBrush]) {
                        $FgKey = $Inline.Foreground.Color.ToString().ToUpper()
                        if ($TermMap.ContainsKey($FgKey)) { $Inline.Foreground = $BC.ConvertFromString($TermMap[$FgKey]) }
                    }
                }
            }
        }
    }

    # Re-render right-panel widgets so code-behind color assignments pick up the new theme
    if (Get-Command Render-ScriptProfiler -ErrorAction SilentlyContinue) { Render-ScriptProfiler }
    if (Get-Command Render-PackageTracker  -ErrorAction SilentlyContinue) { Render-PackageTracker }
    if (Get-Command Render-BuildHistory    -ErrorAction SilentlyContinue) { Render-BuildHistory }
    if (Get-Command Render-KnownIssues     -ErrorAction SilentlyContinue) { Render-KnownIssues }
    if (Get-Command Render-PhaseTracker    -ErrorAction SilentlyContinue) { Render-PhaseTracker }
    if (Get-Command Render-FlameChart      -ErrorAction SilentlyContinue) { Render-FlameChart }
    # Minimap & Heatmap use incremental rendering — force full rebuild by clearing children + resetting index
    if ($cnvMinimap -and $Global:MinimapDots.Count -gt 0) {
        $cnvMinimap.Children.Clear(); $Global:MinimapLastRenderedIndex = 0
        if (Get-Command Render-Minimap -ErrorAction SilentlyContinue) { Render-Minimap }
    }
    if ($cnvHeatmap -and $Global:HeatmapData.Count -gt 0) {
        $cnvHeatmap.Children.Clear(); $Global:HeatmapLastRenderedIndex = 0
        if (Get-Command Render-Heatmap -ErrorAction SilentlyContinue) { Render-Heatmap }
    }
    if (Get-Command Update-StreakDisplay   -ErrorAction SilentlyContinue) { Update-StreakDisplay }
    if (Get-Command Render-Achievements    -ErrorAction SilentlyContinue) { Render-Achievements }
    if (Get-Command Render-ErrorClusters   -ErrorAction SilentlyContinue) { Render-ErrorClusters }
}

# ==============================================================================
# SECTION 7: PACKER LOG COLOR RULES
# ==============================================================================

# Pre-compiled color-rule regexes — compiled once, tested per line via .IsMatch()
$Global:RxError     = [regex]::new('(?i)(error|FAILED|fatal|panic|CRITICAL|Build .* errored|provisioner .* failed)', 'Compiled')
$Global:RxWarning   = [regex]::new('(?i)(warn(ing)?|DEPRECAT|retry|retrying|timeout)', 'Compiled')
$Global:RxSuccess   = [regex]::new('(?i)(Build .* finished|[Ss]uccessfully|[Cc]ompleted|SUCCESS\.|Artifacts were created|output will be|ManagedImageId|SharedImageGallery|Found .+\] Version)', 'Compiled')
$Global:RxSpinner   = [regex]::new('^(?:==>|   )\s+[a-zA-Z_-]+:\s*(?:[-\\|/]\s*|[\u2588\u2591\u2592\u2593].*|)$', 'Compiled')
$Global:RxPackage   = [regex]::new('(?i)(Installing[: ]|Installed[: ]|Upgrading[: ]|Removing[: ]|Downloading https?://|Refreshing metadata|Total installed size|Total download size|Package.*is already installed|Nothing to do|winget install|choco install|Found .+\[.+\])', 'Compiled')
$Global:RxProvision = [regex]::new('(?i)(Provisioning with|Communicating with|Connected to|Waiting for SSH|Waiting for WinRM|Uploading|Pausing|Restarting Machine|Running powershell|Running command)', 'Compiled')
$Global:RxScript    = [regex]::new('(?i)(customizer|BuildArtifacts|Invoke-Expression|Start-Transcript|PSVersion|LogDir|Apply-Customizations|AdminSysPrep)', 'Compiled')
$Global:RxStep      = [regex]::new('^==> [a-zA-Z-]+:\s*(Provisioning|Creating|Querying|Deleting|Capturing|Generalizing|Preparing|Building|Using|Trying|Starting|Waiting for|Looking for|Registering|Cleaning up|Skipping|Tagging|Publishing|Validating|Getting image|Power off|Running sysprep|OS disk|Attempting copy|Creating managed|Getting the VM|Obtaining|Comparing)', 'Compiled')
$Global:RxTimestamp = [regex]::new('^\d{4}/\d{2}/\d{2}|^\d{4}-\d{2}-\d{2}|^Loaded plugin', 'Compiled')
$Global:RxPrefix    = [regex]::new('^(?:==>|   )\s+[a-zA-Z_-]+:', 'Compiled')

function Get-PackerLineColor {
    <#
    .SYNOPSIS
        Determines the display color for a Packer log line based on pattern matching.
    .DESCRIPTION
        Returns a hashtable with Color (hex) and Bold (bool) for the given log line.
        Uses pre-compiled regexes for performance. Rule order is critical.
    #>
    param([string]$Line)

    if ($Global:RxError.IsMatch($Line))     { return @{ Color = "#D13438"; Bold = $true  } }
    if ($Global:RxWarning.IsMatch($Line))   { return @{ Color = "#FFB900"; Bold = $false } }
    if ($Global:RxSuccess.IsMatch($Line))   { return @{ Color = "#107C10"; Bold = $true  } }
    if ($Global:RxSpinner.IsMatch($Line))   { return @{ Color = "#444444"; Bold = $false } }
    if ($Global:RxPackage.IsMatch($Line))   { return @{ Color = "#8764B8"; Bold = $false } }
    if ($Global:RxProvision.IsMatch($Line)) { return @{ Color = "#60CDFF"; Bold = $false } }
    if ($Global:RxScript.IsMatch($Line))    { return @{ Color = "#FF8C00"; Bold = $false } }
    if ($Global:RxStep.IsMatch($Line))      { return @{ Color = "#0078D4"; Bold = $true  } }
    if ($Global:RxTimestamp.IsMatch($Line))  { return @{ Color = "#888888"; Bold = $false } }
    if ($Global:RxPrefix.IsMatch($Line))    { return @{ Color = "#C0C0C0"; Bold = $false } }
    return @{ Color = "#E0E0E0"; Bold = $false }
}

# ==============================================================================
# SECTION 7B: ANSI ESCAPE CODE STRIPPING
# ==============================================================================

$Global:AnsiRegex = [regex]::new('\x1b\[[0-9;]*m', 'Compiled')

function Strip-AnsiCodes {
    <#
    .SYNOPSIS
        Removes ANSI escape sequences (color/style) from a log line.
    #>
    param([string]$Text)
    if (-not $Text) { return $Text }
    return $Global:AnsiRegex.Replace($Text, '')
}

# ==============================================================================
# SECTION 7C: TIMESTAMP PARSING & RELATIVE TIME
# ==============================================================================

$Global:LastTimestamp = $null
$Global:TimestampRegex = [regex]::new(
    '(?:^\[?(\d{1,2}:\d{2}:\d{2}(?:\.\d+)?)\]?|^(\d{4}[/-]\d{2}[/-]\d{2}[T ]\d{2}:\d{2}:\d{2}))',
    'Compiled'
)

function Get-RelativeTimestamp {
    <#
    .SYNOPSIS
        Parses a timestamp from a log line and returns a delta string from the previous line.
    #>
    param([string]$Line)
    $M = $Global:TimestampRegex.Match($Line)
    if (-not $M.Success) { return $null }
    $Raw = if ($M.Groups[1].Success) { $M.Groups[1].Value } else { $M.Groups[2].Value }
    try {
        $Parsed = [DateTime]::Parse($Raw, [System.Globalization.CultureInfo]::InvariantCulture)
        $Delta = $null
        if ($Global:LastTimestamp) {
            $Span = $Parsed - $Global:LastTimestamp
            if ($Span.TotalSeconds -ge 0) {
                if ($Span.TotalMinutes -ge 1) {
                    $Delta = "+{0:0}m{1:00}s" -f [int]$Span.TotalMinutes, $Span.Seconds
                } else {
                    $Delta = "+{0:0.0}s" -f $Span.TotalSeconds
                }
            }
        }
        $Global:LastTimestamp = $Parsed
        return $Delta
    } catch { return $null }
}

# ==============================================================================
# SECTION 7D: SEARCH HIGHLIGHT & NAVIGATION
# ==============================================================================

$Global:SearchMatches = [System.Collections.Generic.List[System.Windows.Documents.TextRange]]::new()
$Global:SearchIndex   = -1
$Global:HighlightBrush = [System.Windows.Media.SolidColorBrush]::new(
    [System.Windows.Media.Color]::FromArgb(100, 255, 185, 0))  # Semi-transparent amber
$Global:ActiveHighlight = [System.Windows.Media.SolidColorBrush]::new(
    [System.Windows.Media.Color]::FromArgb(180, 255, 185, 0))  # Brighter amber for current match

function Search-LogHighlight {
    <#
    .SYNOPSIS
        Scans all paragraphs in the RichTextBox for the search term, highlights matches, and returns count.
    #>
    param([string]$Term)

    # Clear previous highlights
    foreach ($Range in $Global:SearchMatches) {
        try { $Range.ApplyPropertyValue([System.Windows.Documents.TextElement]::BackgroundProperty, $null) } catch {}
    }
    $Global:SearchMatches.Clear()
    $Global:SearchIndex = -1

    if ([string]::IsNullOrWhiteSpace($Term)) {
        $lblSearchCount.Text = ""
        return
    }

    $Doc = $rtbOutput.Document
    $Start = $Doc.ContentStart
    $End   = $Doc.ContentEnd

    # Use TextRange to get full text, then find occurrences
    $FullRange = New-Object System.Windows.Documents.TextRange($Start, $End)
    $FullText  = $FullRange.Text

    $EscapedTerm = [regex]::Escape($Term)
    $Matches = [regex]::Matches($FullText, $EscapedTerm, 'IgnoreCase')

    if ($Matches.Count -eq 0) {
        $lblSearchCount.Text = "0"
        return
    }

    # Walk through the document to find TextPointer positions for each match
    $TextPos = $Start
    $CharOffset = 0

    foreach ($M in $Matches) {
        # Advance to the match start
        while ($CharOffset -lt $M.Index -and $TextPos -ne $null) {
            $NextCtx = $TextPos.GetNextInsertionPosition([System.Windows.Documents.LogicalDirection]::Forward)
            if ($NextCtx -eq $null) { break }
            $RunBetween = New-Object System.Windows.Documents.TextRange($TextPos, $NextCtx)
            $Len = $RunBetween.Text.Length
            if ($CharOffset + $Len -gt $M.Index) { break }
            $CharOffset += $Len
            $TextPos = $NextCtx
        }

        # Find match end
        $MatchStart = $TextPos
        $MatchEnd = $MatchStart
        for ($ci = 0; $ci -lt $Term.Length; $ci++) {
            $Next = $MatchEnd.GetNextInsertionPosition([System.Windows.Documents.LogicalDirection]::Forward)
            if ($Next -eq $null) { break }
            $MatchEnd = $Next
        }

        $HighlightRange = New-Object System.Windows.Documents.TextRange($MatchStart, $MatchEnd)
        $HighlightRange.ApplyPropertyValue(
            [System.Windows.Documents.TextElement]::BackgroundProperty,
            $Global:HighlightBrush)
        $Global:SearchMatches.Add($HighlightRange)
    }

    $lblSearchCount.Text = "$($Global:SearchMatches.Count)"
    if ($Global:SearchMatches.Count -gt 0) {
        $Global:SearchIndex = 0
        Set-SearchActiveMatch 0
    }
}

function Set-SearchActiveMatch {
    <#
    .SYNOPSIS
        Scrolls to and visually emphasises the Nth match.
    #>
    param([int]$Index)
    if ($Global:SearchMatches.Count -eq 0) { return }
    # Reset previous active to normal highlight
    foreach ($Range in $Global:SearchMatches) {
        try { $Range.ApplyPropertyValue([System.Windows.Documents.TextElement]::BackgroundProperty, $Global:HighlightBrush) } catch {}
    }
    $CurRange = $Global:SearchMatches[$Index]
    $CurRange.ApplyPropertyValue([System.Windows.Documents.TextElement]::BackgroundProperty, $Global:ActiveHighlight)
    # Scroll to the match
    $Rect = $CurRange.Start.GetCharacterRect([System.Windows.Documents.LogicalDirection]::Forward)
    $rtbOutput.ScrollToVerticalOffset($rtbOutput.VerticalOffset + $Rect.Top - ($rtbOutput.ViewportHeight / 3))
    $lblSearchCount.Text = "$($Index + 1)/$($Global:SearchMatches.Count)"
}

function Search-Next {
    if ($Global:SearchMatches.Count -eq 0) { return }
    $Global:SearchIndex = ($Global:SearchIndex + 1) % $Global:SearchMatches.Count
    Set-SearchActiveMatch $Global:SearchIndex
}

function Search-Prev {
    if ($Global:SearchMatches.Count -eq 0) { return }
    $Global:SearchIndex = ($Global:SearchIndex - 1 + $Global:SearchMatches.Count) % $Global:SearchMatches.Count
    Set-SearchActiveMatch $Global:SearchIndex
}

# ==============================================================================
# SECTION 7E: PHASE PROGRESS TRACKER
# ==============================================================================

$Global:PhaseDefinitions = @(
    @{ Pattern = 'Preparing Environment|Blocking Updates|Disable.*Auto.*Update|Provisioning with';                    Name = 'Prepare Environment'; Icon = [char]0xE713 }
    @{ Pattern = 'UpdateWinGet|Winget|WinGet|DesktopAppInstaller|winget install|winget upgrade|Found .+\[.+\]';     Name = 'Install WinGet';       Icon = [char]0xE896 }
    @{ Pattern = 'Install.*Language|LanguagePack|Language.*Online|lpksetup|Install-Language';                          Name = 'Language Packs';       Icon = [char]0xE774 }
    @{ Pattern = 'Install.*BasePkg|Install.*Base.*Package|CUSTOMIZER PHASE.*Base';                                     Name = 'Base Packages';        Icon = [char]0xE7B8 }
    @{ Pattern = 'Install.*KrPkg|Install.*Kr.*Package|CUSTOMIZER PHASE.*Kr|CUSTOMIZER PHASE.*Custom';                  Name = 'KR/Custom Packages';   Icon = [char]0xE7B8 }
    @{ Pattern = 'Install.*TradersPkg|Traders|CUSTOMIZER PHASE.*Trader';                                               Name = 'Traders Packages';     Icon = [char]0xE7B8 }
    @{ Pattern = 'WindowsOptimization|Optimization|CUSTOMIZER PHASE.*Optim';                                          Name = 'Windows Optimization'; Icon = [char]0xE9F5 }
    @{ Pattern = 'RemoveAppx|Remove.*Appx|ProvisionedAppx|CUSTOMIZER PHASE.*Remove|CUSTOMIZER PHASE.*Cleanup';        Name = 'Remove AppX';          Icon = [char]0xE74D }
    @{ Pattern = 'TimezoneRedirection|Timezone|CUSTOMIZER PHASE.*Timezone';                                            Name = 'Timezone Config';      Icon = [char]0xE916 }
    @{ Pattern = 'SysPrep|Sysprep|AdminSysPrep|Generalize|CUSTOMIZER PHASE.*SysPrep|deprovision\+user';               Name = 'SysPrep/Generalize';   Icon = [char]0xE912 }
    @{ Pattern = 'Artifacts were created|SharedImageGallery|Creating image version|Published to|Capturing image';      Name = 'Publish Image';        Icon = [char]0xE73E }
)
$Global:PhaseStatus = @{}

function Reset-PhaseTracker {
    <# Resets phase tracker to initial state #>
    $Global:PhaseStatus = @{}
    $Global:LastTimestamp = $null
    $pnlPhaseTracker.Children.Clear()
    $pnlPhaseTracker.Children.Add($lblPhaseEmpty) | Out-Null
    $lblPhaseEmpty.Visibility = 'Visible'
}

function Update-PhaseFromLine {
    <#
    .SYNOPSIS
        Checks a log line against phase patterns and marks matching phases as active.
        Also updates flame chart phase timing (merged from Update-PhaseTiming to avoid double regex matching).
    #>
    param([string]$Line)
    $Changed = $false
    $MatchedPhase = $null
    foreach ($Phase in $Global:PhaseDefinitions) {
        $Key = $Phase.Name
        if ($Global:PhaseStatus[$Key] -eq 'done') { continue }
        if ($Line -match $Phase.Pattern) {
            # Mark previous in-progress phase as done
            foreach ($PK in @($Global:PhaseStatus.Keys)) {
                if ($Global:PhaseStatus[$PK] -eq 'active') {
                    $Global:PhaseStatus[$PK] = 'done'
                }
            }
            $Global:PhaseStatus[$Key] = 'active'
            $MatchedPhase = $Phase
            $Changed = $true
            break
        }
    }
    if ($Changed) {
        Render-PhaseTracker

        # Flame chart timing (merged from Update-PhaseTiming)
        $LogTime = Get-Date
        $TsMatch = $Global:TimestampRegex.Match($Line)
        if ($TsMatch.Success) {
            $TsRaw = if ($TsMatch.Groups[1].Success) { $TsMatch.Groups[1].Value } else { $TsMatch.Groups[2].Value }
            try { $LogTime = [DateTime]::Parse($TsRaw, [System.Globalization.CultureInfo]::InvariantCulture) } catch { }
        }
        # Close previous phase
        if ($Global:FlameChartLastPhase) {
            $Prev = $Global:PhaseTimings | Where-Object { $_.Name -eq $Global:FlameChartLastPhase -and -not $_.EndTime }
            if ($Prev) {
                $Prev | ForEach-Object { $_.EndTime = $LogTime }
            }
        }
        # Open new phase
        $Global:PhaseTimings.Add([PSCustomObject]@{
            Name      = $MatchedPhase.Name
            StartTime = $LogTime
            EndTime   = $null
            Color     = '#0078D4'
        })
        $Global:FlameChartLastPhase = $MatchedPhase.Name
        $Global:FlameChartLastTime  = $LogTime
        Render-FlameChart
    }
}

function Render-PhaseTracker {
    <# Rebuilds the phase tracker visual from current status #>
    $pnlPhaseTracker.Children.Clear()
    $lblPhaseEmpty.Visibility = 'Collapsed'
    $BC = $Global:CachedBC

    foreach ($Phase in $Global:PhaseDefinitions) {
        $Key = $Phase.Name
        $Status = $Global:PhaseStatus[$Key]
        if (-not $Status) { $Status = 'pending' }

        $Row = New-Object System.Windows.Controls.StackPanel
        $Row.Orientation = 'Horizontal'
        $Row.Margin = [System.Windows.Thickness]::new(0, 2, 0, 2)

        $Icon = New-Object System.Windows.Controls.TextBlock
        $Icon.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
        $Icon.FontSize = 11
        $Icon.VerticalAlignment = 'Center'
        $Icon.Margin = [System.Windows.Thickness]::new(0, 0, 8, 0)
        $Icon.Width = 16

        $Label = New-Object System.Windows.Controls.TextBlock
        $Label.Text = $Phase.Name
        $Label.FontSize = 11
        $Label.VerticalAlignment = 'Center'

        switch ($Status) {
            'done' {
                $Icon.Text = [string][char]0xE73E  # checkmark
                $Icon.Foreground = $Window.Resources['ThemeSuccess']
                $Label.Foreground = $Window.Resources['ThemeTextMuted']
            }
            'active' {
                $Icon.Text = [string][char]0xEA3A  # spinner/arrow
                $Icon.Foreground = $Window.Resources['ThemeAccent']
                $Label.Foreground = $Window.Resources['ThemeTextPrimary']
                $Label.FontWeight = 'SemiBold'
            }
            default {
                $Icon.Text = [string][char]0xF13C  # circle outline
                $Icon.Foreground = $Window.Resources['ThemeTextDisabled']
                $Label.Foreground = $Window.Resources['ThemeTextDisabled']
            }
        }

        $Row.Children.Add($Icon) | Out-Null
        $Row.Children.Add($Label) | Out-Null
        $pnlPhaseTracker.Children.Add($Row) | Out-Null
    }
}

# ==============================================================================
# SECTION 7F: LOG EXPORT
# ==============================================================================

function Copy-LogToClipboard {
    <#
    .SYNOPSIS
        Copies the entire log content to the clipboard.
    #>
    try {
        $FullRange = New-Object System.Windows.Documents.TextRange(
            $rtbOutput.Document.ContentStart,
            $rtbOutput.Document.ContentEnd)
        $Text = $FullRange.Text
        if ([string]::IsNullOrWhiteSpace($Text)) {
            $lblStatus.Text = "Nothing to copy — log is empty"
            return
        }
        [System.Windows.Clipboard]::SetText($Text)
        $LineCount = ($Text -split "`n").Count
        $lblStatus.Text = "Copied $LineCount lines to clipboard"
        Write-DebugLog "Log copied to clipboard ($LineCount lines)"
    } catch {
        $lblStatus.Text = "Copy failed: $($_.Exception.Message)"
    }
}

function Export-LogToFile {
    <#
    .SYNOPSIS
        Exports the current RichTextBox content to a plain text file.
    #>
    $Dialog = New-Object Microsoft.Win32.SaveFileDialog
    $Dialog.Title = "Save Packer Log"
    $Dialog.Filter = "Text files (*.txt)|*.txt|Log files (*.log)|*.log|All files (*.*)|*.*"
    $Dialog.DefaultExt = ".txt"
    $Dialog.FileName = "packer-log-$(Get-Date -Format 'yyyyMMdd-HHmmss')"

    if ($Dialog.ShowDialog() -eq $true) {
        try {
            $FullRange = New-Object System.Windows.Documents.TextRange(
                $rtbOutput.Document.ContentStart,
                $rtbOutput.Document.ContentEnd)
            [System.IO.File]::WriteAllText($Dialog.FileName, $FullRange.Text, [System.Text.Encoding]::UTF8)
            $lblStatus.Text = "Log saved to $($Dialog.FileName)"
            Write-DebugLog "Log exported: $($Dialog.FileName)"
        } catch {
            $lblStatus.Text = "Failed to save: $($_.Exception.Message)"
        }
    }
}

# ==============================================================================
# SECTION 7G: SEVERITY BADGES (#1)
# ==============================================================================

function Get-SeverityBadge {
    <#
    .SYNOPSIS
        Returns a severity label and color for a log line, used to prepend a
        colored pill badge (ERR, WARN, INFO, OK, DBG) before the text.
    #>
    param([string]$Color)
    switch ($Color) {
        '#D13438' { return @{ Label = 'ERR';  Bg = '#D13438'; Fg = '#FFFFFF' } }
        '#FFB900' { return @{ Label = 'WARN'; Bg = '#FFB900'; Fg = '#1B1B1B' } }
        '#107C10' { return @{ Label = 'OK';   Bg = '#107C10'; Fg = '#FFFFFF' } }
        '#0078D4' { return @{ Label = 'STEP'; Bg = '#0078D4'; Fg = '#FFFFFF' } }
        '#60CDFF' { return @{ Label = 'PROV'; Bg = '#60CDFF'; Fg = '#1B1B1B' } }
        '#8764B8' { return @{ Label = 'PKG';  Bg = '#8764B8'; Fg = '#FFFFFF' } }
        '#FF8C00' { return @{ Label = 'SCRP'; Bg = '#FF8C00'; Fg = '#1B1B1B' } }
        '#888888' { return @{ Label = 'META'; Bg = '#555555'; Fg = '#CCCCCC' } }
        '#444444' { return $null }  # spinner lines — no badge
        default   { return @{ Label = 'INFO'; Bg = '#333333'; Fg = '#AAAAAA' } }
    }
}

# ==============================================================================
# SECTION 7H: LOG STATS TRACKER (#9)
# ==============================================================================

$Global:LogStats = @{
    TotalLines    = 0
    ErrorCount    = 0
    WarningCount  = 0
    SuccessCount  = 0
    LinesPerSec   = 0.0
    StartTime     = $null
    TopErrors     = [System.Collections.Generic.Dictionary[string,int]]::new()
}

function Reset-LogStats {
    $Global:LogStats.TotalLines  = 0
    $Global:LogStats.ErrorCount  = 0
    $Global:LogStats.WarningCount = 0
    $Global:LogStats.SuccessCount = 0
    $Global:LogStats.LinesPerSec = 0.0
    $Global:LogStats.StartTime   = $null
    $Global:LogStats.TopErrors.Clear()
    Update-StatsDisplay
}

function Update-LogStats {
    param([string]$Color, [string]$CleanText)
    $Global:LogStats.TotalLines++
    if (-not $Global:LogStats.StartTime) { $Global:LogStats.StartTime = Get-Date }
    switch ($Color) {
        '#D13438' {
            $Global:LogStats.ErrorCount++
            # Track top error messages (first 80 chars as key)
            $Key = if ($CleanText.Length -gt 80) { $CleanText.Substring(0, 80) } else { $CleanText }
            if ($Global:LogStats.TopErrors.ContainsKey($Key)) {
                $Global:LogStats.TopErrors[$Key]++
            } else {
                if ($Global:LogStats.TopErrors.Count -lt 10) {
                    $Global:LogStats.TopErrors[$Key] = 1
                }
            }
        }
        '#FFB900' { $Global:LogStats.WarningCount++ }
        '#107C10' { $Global:LogStats.SuccessCount++ }
    }
    # LPS is computed on display update only (see Update-StatsDisplay)
}

function Update-StatsDisplay {
    param([DateTime]$Now = (Get-Date))
    if (-not $lblStatsErrors) { return }
    $S = $Global:LogStats
    # Compute LPS only at display time (not per-line)
    if ($S.StartTime) {
        $ElapsedSec = ($Now - $S.StartTime).TotalSeconds
        if ($ElapsedSec -gt 0) {
            $S.LinesPerSec = [math]::Round($S.TotalLines / $ElapsedSec, 1)
        }
    }
    $lblStatsErrors.Text   = "$($S.ErrorCount)"
    $lblStatsWarnings.Text = "$($S.WarningCount)"
    $lblStatsSuccess.Text  = "$($S.SuccessCount)"
    $lblStatsLines.Text    = "$($S.TotalLines)"
    $lblStatsLps.Text      = "$($S.LinesPerSec)/s"
}

# ==============================================================================
# SECTION 7I: TAIL-FOLLOW INDICATOR (#10)
# ==============================================================================

$Global:IsFollowing = $true

function Update-FollowIndicator {
    param([bool]$Following)
    $Global:IsFollowing = $Following
    if ($lblFollowIndicator) {
        if ($Following) {
            $lblFollowIndicator.Text = [string][char]0x25BC + " Following"   # ▼ standard Unicode
            $lblFollowIndicator.Foreground = $Window.Resources['ThemeTextDim']
        } else {
            $lblFollowIndicator.Text = [string][char]0x25A0 + " Paused"       # ■ standard Unicode
            $lblFollowIndicator.Foreground = $Window.Resources['ThemeWarning']
        }
    }
}

# ==============================================================================
# SECTION 7J: JSON STRUCTURED VIEW (#11)
# ==============================================================================

function Test-JsonLine {
    <# Returns $true if the line appears to be a JSON object #>
    param([string]$Line)
    $Line = $Line.Trim()
    return ($Line.StartsWith('{') -and $Line.EndsWith('}') -and $Line.Length -gt 10)
}

function Format-JsonPretty {
    <# Pretty-prints a JSON string with indentation #>
    param([string]$Json)
    try {
        $Obj = $Json | ConvertFrom-Json -ErrorAction Stop
        return ($Obj | ConvertTo-Json -Depth 10)
    } catch {
        return $null
    }
}

# ==============================================================================
# SECTION 7K: VIRTUAL SCROLL LINE CAP (#15)
# ==============================================================================

$Global:MaxParagraphs = 15000
$Global:MaxLinesPerTick = 200          # batch cap — keeps UI responsive during backfill
$Global:TrimmedCount  = 0
$Global:ParagraphCount = 0             # O(1) counter — avoids linked-list walk in Trim-FlowDocument

function Trim-FlowDocument {
    <#
    .SYNOPSIS
        Removes oldest paragraphs when the document exceeds MaxParagraphs.
        Uses O(1) ParagraphCount instead of walking the linked list.
        Also trims corresponding minimap/heatmap data to avoid phantom dots (E1).
    #>
    if ($Global:ParagraphCount -le $Global:MaxParagraphs) { return }

    $Doc = $rtbOutput.Document
    $RemoveCount = $Global:ParagraphCount - $Global:MaxParagraphs + 1000  # trim 1000 extra to avoid frequent trims
    $Block = $Doc.Blocks.FirstBlock
    $Removed = 0
    for ($i = 0; $i -lt $RemoveCount -and $Block; $i++) {
        $Next = $Block.NextBlock
        $Doc.Blocks.Remove($Block)
        $Block = $Next
        $Removed++
    }
    $Global:ParagraphCount -= $Removed
    $Global:TrimmedCount += $Removed

    # Trim corresponding minimap/heatmap data (E1 — prevent phantom dots)
    if ($Removed -gt 0) {
        if ($Global:MinimapDots.Count -gt $Removed) {
            $Global:MinimapDots.RemoveRange(0, $Removed)
        } else {
            $Global:MinimapDots.Clear()
        }
        $Global:MinimapLastRenderedIndex = 0
        if ($cnvMinimap) { $cnvMinimap.Children.Clear() }

        if ($Global:HeatmapData.Count -gt $Removed) {
            $Global:HeatmapData.RemoveRange(0, $Removed)
        } else {
            $Global:HeatmapData.Clear()
        }
        $Global:HeatmapLastRenderedIndex = 0
        if ($cnvHeatmap) { $cnvHeatmap.Children.Clear() }
    }
}

# ==============================================================================
# SECTION 7L: BUILD HISTORY (#16)
# ==============================================================================

$Global:BuildHistoryPath = Join-Path $Global:Root "build_history.json"
$Global:BuildHistory = [System.Collections.Generic.List[PSObject]]::new()

function Load-BuildHistory {
    if (Test-Path $Global:BuildHistoryPath) {
        try {
            $Json = Get-Content $Global:BuildHistoryPath -Raw | ConvertFrom-Json
            $Global:BuildHistory.Clear()
            foreach ($Item in $Json) {
                $Global:BuildHistory.Add($Item)
            }
        } catch { }
    }
}

function Save-BuildHistory {
    try {
        $Global:BuildHistory | ConvertTo-Json -Depth 5 | Set-Content $Global:BuildHistoryPath -Force
    } catch { }
}

function Add-BuildToHistory {
    param([PSCustomObject]$Build, [string]$FinalStatus)
    # Avoid duplicates by template name + start time
    $Existing = $Global:BuildHistory | Where-Object {
        $_.TemplateName -eq $Build.TemplateName -and $_.StartTime -eq $Build.StartTime
    } | Select-Object -First 1
    # Compute duration from build start
    $DurationMin = 0
    if ($Build.StartTime) {
        $DurationMin = [Math]::Round(((Get-Date).ToUniversalTime() - $Build.StartTime).TotalMinutes, 1)
    }

    # Snapshot phase durations for historical anomaly detection
    $PhaseDurations = $null
    if ($Global:PhaseTimings.Count -gt 0) {
        $PhaseDurations = @{}
        $Now = Get-Date
        foreach ($PT in $Global:PhaseTimings) {
            $End = if ($PT.EndTime) { $PT.EndTime } else { $Now }
            $PhaseDurations[$PT.Name] = ($End - $PT.StartTime).TotalSeconds
        }
    }

    if ($Existing) {
        $Existing.FinalStatus = $FinalStatus
        $Existing.LastViewedAt = (Get-Date).ToString('o')
        $Existing.LinesCaptured = $Global:LogStats.TotalLines
        $Existing.ErrorCount = $Global:LogStats.ErrorCount
        $Existing.DurationMin = $DurationMin
        $Existing.PhaseDurations = $PhaseDurations
    } else {
        $Entry = [PSCustomObject]@{
            TemplateName   = $Build.TemplateName
            ResourceGroup  = $Build.ResourceGroup
            Location       = $Build.Location
            StartTime      = $Build.StartTime
            FinalStatus    = $FinalStatus
            LastViewedAt   = (Get-Date).ToString('o')
            LinesCaptured  = $Global:LogStats.TotalLines
            ErrorCount     = $Global:LogStats.ErrorCount
            DurationMin    = $DurationMin
            PhaseDurations = $PhaseDurations
        }
        $Global:BuildHistory.Insert(0, $Entry)
        # Keep max 50 entries
        while ($Global:BuildHistory.Count -gt 50) {
            $Global:BuildHistory.RemoveAt($Global:BuildHistory.Count - 1)
        }
    }
    Save-BuildHistory
}

function Render-BuildHistory {
    if (-not $pnlBuildHistory) { return }
    $pnlBuildHistory.Children.Clear()
    $BC = $Global:CachedBC

    if ($Global:BuildHistory.Count -eq 0) {
        $Empty = New-Object System.Windows.Controls.TextBlock
        $Empty.Text = "No previous builds"
        $Empty.FontSize = 10
        $Empty.FontStyle = 'Italic'
        $Empty.Foreground = $Window.Resources['ThemeTextDim']
        $pnlBuildHistory.Children.Add($Empty) | Out-Null
        return
    }

    foreach ($H in ($Global:BuildHistory | Select-Object -First 10)) {
        $Row = New-Object System.Windows.Controls.Border
        $Row.CornerRadius = [System.Windows.CornerRadius]::new(4)
        $Row.Padding = [System.Windows.Thickness]::new(8, 5, 8, 5)
        $Row.Margin = [System.Windows.Thickness]::new(0, 0, 0, 3)
        $Row.Background = $Window.Resources['ThemeDeepBg']
        $Row.Cursor = [System.Windows.Input.Cursors]::Hand

        $Inner = New-Object System.Windows.Controls.StackPanel

        $NameTxt = New-Object System.Windows.Controls.TextBlock
        $NameTxt.Text = $H.TemplateName
        $NameTxt.FontSize = 10.5
        $NameTxt.FontWeight = 'SemiBold'
        $NameTxt.Foreground = $Window.Resources['ThemeTextBody']
        $NameTxt.TextTrimming = 'CharacterEllipsis'

        $DetailTxt = New-Object System.Windows.Controls.TextBlock
        $DetailTxt.FontSize = 9
        $StatusClr = switch ($H.FinalStatus) {
            'Completed' { & $Global:RemapLight '#107C10' }
            'Error'     { & $Global:RemapLight '#D13438' }
            'Running'   { & $Global:RemapLight '#0078D4' }
            default     { & $Global:RemapLight '#888888' }
        }
        try { $DetailTxt.Foreground = $BC.ConvertFromString($StatusClr) } catch { $DetailTxt.Foreground = $Window.Resources['ThemeTextDim'] }
        $ViewedAt = try { ([DateTime]$H.LastViewedAt).ToString('MMM dd HH:mm') } catch { 'N/A' }
        $DetailTxt.Text = "$($H.FinalStatus) | $($H.LinesCaptured) lines | $ViewedAt"

        $Inner.Children.Add($NameTxt) | Out-Null
        $Inner.Children.Add($DetailTxt) | Out-Null
        $Row.Child = $Inner
        $pnlBuildHistory.Children.Add($Row) | Out-Null
    }
}

# ==============================================================================
# SECTION 7M: SOUND ALERTS (#18)
# ==============================================================================

$Global:SoundEnabled = $false
$Global:LastSoundTime = [DateTime]::MinValue
$Global:SoundCooldownSec = 5  # Don't play more than once per 5 seconds

function Play-ErrorSound {
    if (-not $Global:SoundEnabled) { return }
    if (((Get-Date) - $Global:LastSoundTime).TotalSeconds -lt $Global:SoundCooldownSec) { return }
    $Global:LastSoundTime = Get-Date
    try {
        [System.Media.SystemSounds]::Exclamation.Play()
    } catch { }
}

function Play-CompletionSound {
    if (-not $Global:SoundEnabled) { return }
    try {
        [System.Media.SystemSounds]::Asterisk.Play()
    } catch { }
}

# ==============================================================================
# SECTION 7N: SMART NOTIFICATIONS (#8)
# ==============================================================================

function Show-ToastNotification {
    <#
    .SYNOPSIS
        Shows a Windows balloon tip notification via a reusable NotifyIcon.
    #>
    param([string]$Title, [string]$Message, [string]$Type = 'Info')
    try {
        # Reuse a single NotifyIcon to avoid handle leaks
        if (-not $Global:NotifyIcon) {
            $Global:NotifyIcon = New-Object System.Windows.Forms.NotifyIcon
            $Global:NotifyIcon.Icon = [System.Drawing.SystemIcons]::Information
        }
        $Global:NotifyIcon.Visible = $true
        $BalloonType = switch ($Type) {
            'Error'   { [System.Windows.Forms.ToolTipIcon]::Error }
            'Warning' { [System.Windows.Forms.ToolTipIcon]::Warning }
            default   { [System.Windows.Forms.ToolTipIcon]::Info }
        }
        $Global:NotifyIcon.ShowBalloonTip(5000, $Title, $Message, $BalloonType)
        # Hide after 6 seconds (don't dispose — reuse)
        if (-not $Global:NotifyHideTimer) {
            $Global:NotifyHideTimer = New-Object System.Windows.Threading.DispatcherTimer
            $Global:NotifyHideTimer.Interval = [TimeSpan]::FromSeconds(6)
            $Global:NotifyHideTimer.Add_Tick({
                $Global:NotifyHideTimer.Stop()
                if ($Global:NotifyIcon) { $Global:NotifyIcon.Visible = $false }
            })
        } else {
            $Global:NotifyHideTimer.Stop()
        }
        $Global:NotifyHideTimer.Start()
    } catch { }
}

$Global:NotifyOnError    = $true
$Global:NotifyOnComplete = $true
$Global:LastNotifyTime   = [DateTime]::MinValue
$Global:NotifyCooldownSec = 15

function Notify-OnError {
    param([string]$Line)
    if (-not $Global:NotifyOnError) { return }
    if (-not $Window.IsActive) {  # Only notify when window is not focused
        if (((Get-Date) - $Global:LastNotifyTime).TotalSeconds -ge $Global:NotifyCooldownSec) {
            $Global:LastNotifyTime = Get-Date
            $Short = if ($Line.Length -gt 120) { $Line.Substring(0, 120) + '...' } else { $Line }
            Show-ToastNotification -Title "Build Error Detected" -Message $Short -Type 'Error'
        }
    }
}

function Notify-OnBuildComplete {
    param([string]$TemplateName)
    if (-not $Global:NotifyOnComplete) { return }
    Show-ToastNotification -Title "Build Completed" -Message "$TemplateName has finished." -Type 'Info'
    Play-CompletionSound
}

# ==============================================================================
# SECTION 7P: MINIMAP SCROLLBAR (#20)
# ==============================================================================

$Global:MinimapDots = [System.Collections.Generic.List[PSObject]]::new()
$Global:MinimapLineTotal = 0
$Global:MinimapLastRenderedIndex = 0    # incremental render marker

function Add-MinimapDot {
    <# Records a colored dot position in the minimap data #>
    param([string]$Color, [int]$LineIndex)
    if ($Color -eq '#D13438' -or $Color -eq '#FFB900' -or $Color -eq '#107C10') {
        $Global:MinimapDots.Add([PSCustomObject]@{ Color = $Color; Line = $LineIndex })
    }
    $Global:MinimapLineTotal = $LineIndex
}

function Render-Minimap {
    <# Incrementally adds new dots to the minimap Canvas. Full rebuild only on resize. #>
    if (-not $cnvMinimap) { return }
    $H = $cnvMinimap.ActualHeight
    if ($H -le 0 -or $Global:MinimapLineTotal -le 0) { return }

    $BC = $Global:CachedBC
    $W  = $cnvMinimap.ActualWidth

    # Full rebuild if Canvas was cleared (reset) or no children yet
    $FullRebuild = ($cnvMinimap.Children.Count -eq 0 -and $Global:MinimapDots.Count -gt 0)
    $StartIdx = if ($FullRebuild) { 0 } else { $Global:MinimapLastRenderedIndex }

    for ($di = $StartIdx; $di -lt $Global:MinimapDots.Count; $di++) {
        $Dot = $Global:MinimapDots[$di]
        $Y = ($Dot.Line / $Global:MinimapLineTotal) * $H
        $Rect = New-Object System.Windows.Shapes.Rectangle
        $Rect.Width = $W
        $Rect.Height = [Math]::Max(2, $H / $Global:MinimapLineTotal)
        $Rect.Fill = $BC.ConvertFromString((& $Global:RemapLight $Dot.Color))
        $Rect.Opacity = 0.7
        [System.Windows.Controls.Canvas]::SetTop($Rect, $Y)
        [System.Windows.Controls.Canvas]::SetLeft($Rect, 0)
        $cnvMinimap.Children.Add($Rect) | Out-Null
    }
    $Global:MinimapLastRenderedIndex = $Global:MinimapDots.Count

    # Remove old viewport indicator (last 2 children if present: viewport + stroke)
    # Re-draw viewport indicator fresh each time
    $ViewTag = 'minimap-viewport'
    for ($vi = $cnvMinimap.Children.Count - 1; $vi -ge 0; $vi--) {
        if ($cnvMinimap.Children[$vi].Tag -eq $ViewTag) {
            $cnvMinimap.Children.RemoveAt($vi)
        }
    }
    if ($rtbOutput.ExtentHeight -gt 0) {
        $ViewTop = ($rtbOutput.VerticalOffset / $rtbOutput.ExtentHeight) * $H
        $ViewH   = ($rtbOutput.ViewportHeight / $rtbOutput.ExtentHeight) * $H
        $ViewH   = [Math]::Max($ViewH, 8)
        $ViewRect = New-Object System.Windows.Shapes.Rectangle
        $ViewRect.Width = $W
        $ViewRect.Height = $ViewH
        $ViewRect.Fill = $Window.Resources['ThemeScrollThumb']
        $ViewRect.Opacity = 0.5
        $ViewRect.Stroke = $Window.Resources['ThemeBorderElevated']
        $ViewRect.StrokeThickness = 0.5
        $ViewRect.Tag = $ViewTag
        [System.Windows.Controls.Canvas]::SetTop($ViewRect, $ViewTop)
        [System.Windows.Controls.Canvas]::SetLeft($ViewRect, 0)
        $cnvMinimap.Children.Add($ViewRect) | Out-Null
    }
}

function Reset-Minimap {
    $Global:MinimapDots.Clear()
    $Global:MinimapLineTotal = 0
    $Global:MinimapLastRenderedIndex = 0
    if ($cnvMinimap) { $cnvMinimap.Children.Clear() }
}

# ==============================================================================
# SECTION 7Q: DIFF AGAINST PREVIOUS BUILD (#12)
# ==============================================================================

$Global:SavedLogsDir = Join-Path $Global:Root "saved_logs"

function Save-CurrentLogForDiff {
    <# Auto-saves the current log text to a local file for future diff comparisons #>
    param([string]$TemplateName)
    if (-not (Test-Path $Global:SavedLogsDir)) {
        New-Item -Path $Global:SavedLogsDir -ItemType Directory -Force | Out-Null
    }
    $SafeName = $TemplateName -replace '[^\w\-\.]', '_'
    $FilePath = Join-Path $Global:SavedLogsDir "$SafeName-latest.log"
    try {
        $FullRange = New-Object System.Windows.Documents.TextRange(
            $rtbOutput.Document.ContentStart, $rtbOutput.Document.ContentEnd)
        [System.IO.File]::WriteAllText($FilePath, $FullRange.Text, [System.Text.Encoding]::UTF8)
        Write-DebugLog "Log auto-saved for diff: $FilePath"
    } catch { }
}

function AutoSave-BuildLog {
    <# Auto-saves the current log with timestamp to saved_logs folder #>
    param([string]$TemplateName, [string]$Status)
    if (-not (Test-Path $Global:SavedLogsDir)) {
        New-Item -Path $Global:SavedLogsDir -ItemType Directory -Force | Out-Null
    }
    $SafeName = $TemplateName -replace '[^\w\-\.]', '_'
    $Ts = Get-Date -Format 'yyyyMMdd-HHmmss'
    $FilePath = Join-Path $Global:SavedLogsDir "${SafeName}_${Status}_${Ts}.log"
    try {
        $FullRange = New-Object System.Windows.Documents.TextRange(
            $rtbOutput.Document.ContentStart, $rtbOutput.Document.ContentEnd)
        $LogText = $FullRange.Text
        if ([string]::IsNullOrWhiteSpace($LogText)) { return }
        [System.IO.File]::WriteAllText($FilePath, $LogText, [System.Text.Encoding]::UTF8)
        $lblStatus.Text = "Log auto-saved: $([System.IO.Path]::GetFileName($FilePath))"
        Write-DebugLog "Log auto-saved ($Status): $FilePath"
    } catch {
        Write-DebugLog "AutoSave-BuildLog failed: $($_.Exception.Message)"
    }
}

function Show-DiffDialog {
    <#
    .SYNOPSIS
        Opens a file dialog to select a previously saved log and shows a
        simple diff summary in the log output.
    #>
    $Dialog = New-Object Microsoft.Win32.OpenFileDialog
    $Dialog.Title = "Select previous build log for comparison"
    $Dialog.Filter = "Log files (*.log;*.txt)|*.log;*.txt|All files (*.*)|*.*"
    if (Test-Path $Global:SavedLogsDir) { $Dialog.InitialDirectory = $Global:SavedLogsDir }

    if ($Dialog.ShowDialog() -ne $true) { return }

    try {
        $OldLines = [System.IO.File]::ReadAllLines($Dialog.FileName) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        $NewRange = New-Object System.Windows.Documents.TextRange(
            $rtbOutput.Document.ContentStart, $rtbOutput.Document.ContentEnd)
        $NewLines = ($NewRange.Text -split "`n") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        $OldSet = [System.Collections.Generic.HashSet[string]]::new($OldLines)
        $NewSet = [System.Collections.Generic.HashSet[string]]::new($NewLines)

        $Added   = $NewLines | Where-Object { -not $OldSet.Contains($_) }
        $Removed = $OldLines | Where-Object { -not $NewSet.Contains($_) }

        $BC = $Global:CachedBC
        # Header
        $HdrPara = New-Object System.Windows.Documents.Paragraph
        $HdrPara.Margin = [System.Windows.Thickness]::new(0, 10, 0, 5)
        $HdrRun = New-Object System.Windows.Documents.Run("[DIFF] Comparing against: $($Dialog.FileName)")
        $HdrRun.Foreground = $BC.ConvertFromString((& $Global:RemapLight '#60CDFF'))
        $HdrRun.FontWeight = 'Bold'
        $HdrPara.Inlines.Add($HdrRun)
        $rtbOutput.Document.Blocks.Add($HdrPara)

        $SummaryPara = New-Object System.Windows.Documents.Paragraph
        $SummaryPara.Margin = [System.Windows.Thickness]::new(0)
        $SumRun = New-Object System.Windows.Documents.Run("+$($Added.Count) new lines | -$($Removed.Count) removed lines")
        $SumRun.Foreground = $Window.Resources['ThemeTextDim']
        $SummaryPara.Inlines.Add($SumRun)
        $rtbOutput.Document.Blocks.Add($SummaryPara)

        $Sep = New-Object System.Windows.Documents.Paragraph
        $Sep.Margin = [System.Windows.Thickness]::new(0)
        $SepRun = New-Object System.Windows.Documents.Run([string]::new([char]0x2500, 60))
        $SepRun.Foreground = $Window.Resources['ThemeBorder']
        $Sep.Inlines.Add($SepRun)
        $rtbOutput.Document.Blocks.Add($Sep)

        # Show first 50 added lines
        $ShowCount = [Math]::Min($Added.Count, 50)
        for ($i = 0; $i -lt $ShowCount; $i++) {
            $P = New-Object System.Windows.Documents.Paragraph
            $P.Margin = [System.Windows.Thickness]::new(0)
            $R = New-Object System.Windows.Documents.Run("+ $($Added[$i])")
            $R.Foreground = $BC.ConvertFromString((& $Global:RemapLight '#107C10'))
            $P.Inlines.Add($R)
            $rtbOutput.Document.Blocks.Add($P)
        }

        # Show first 50 removed lines
        $ShowCount = [Math]::Min($Removed.Count, 50)
        for ($i = 0; $i -lt $ShowCount; $i++) {
            $P = New-Object System.Windows.Documents.Paragraph
            $P.Margin = [System.Windows.Thickness]::new(0)
            $R = New-Object System.Windows.Documents.Run("- $($Removed[$i])")
            $R.Foreground = $BC.ConvertFromString((& $Global:RemapLight '#D13438'))
            $P.Inlines.Add($R)
            $rtbOutput.Document.Blocks.Add($P)
        }

        $rtbOutput.ScrollToEnd()
        $lblStatus.Text = "Diff: +$($Added.Count) / -$($Removed.Count) lines"
    } catch {
        $lblStatus.Text = "Diff failed: $($_.Exception.Message)"
    }
}

# ==============================================================================
# SECTION 7R: DURATION FLAME CHART (#7)
# ==============================================================================

$Global:PhaseTimings = [System.Collections.Generic.List[PSObject]]::new()
$Global:FlameChartLastPhase = $null
$Global:FlameChartLastTime  = $null

function Update-PhaseTiming {
    <# Tracks when phases start/end for duration flame chart #>
    param([string]$Line)

    # This function is now merged with Update-PhaseFromLine to avoid
    # double regex matching the same line against PhaseDefinitions.
    # Update-PhaseFromLine calls this internally when a phase match is found.
    # Kept as a stub for backward compatibility — no-op.
}

function Render-FlameChart {
    <# Renders horizontal duration bars in the flame chart flyout panel #>
    if (-not $pnlFlameChart) { return }
    $pnlFlameChart.Children.Clear()
    $BC = $Global:CachedBC

    # Update badge count on the header
    if ($lblFlameCount) {
        $lblFlameCount.Text = if ($Global:PhaseTimings.Count -gt 0) { "$($Global:PhaseTimings.Count) phases" } else { '' }
    }

    if ($Global:PhaseTimings.Count -eq 0) { return }

    $EarliestStart = ($Global:PhaseTimings | Sort-Object StartTime | Select-Object -First 1).StartTime
    $Now = Get-Date
    $TotalSpan = ($Now - $EarliestStart).TotalSeconds
    if ($TotalSpan -le 0) { return }

    $ChartWidth = 280  # fixed width for flyout

    $PhaseColors = @{
        'Prepare Environment' = (& $Global:RemapLight '#60CDFF'); 'Install WinGet' = (& $Global:RemapLight '#0078D4')
        'Language Packs' = (& $Global:RemapLight '#8764B8'); 'Base Packages' = (& $Global:RemapLight '#FF8C00')
        'KR/Custom Packages' = (& $Global:RemapLight '#FFB900'); 'Traders Packages' = (& $Global:RemapLight '#D13438')
        'Windows Optimization' = (& $Global:RemapLight '#107C10'); 'Remove AppX' = (& $Global:RemapLight '#E74856')
        'Timezone Config' = (& $Global:RemapLight '#888888'); 'SysPrep/Generalize' = (& $Global:RemapLight '#60CDFF')
        'Publish Image' = (& $Global:RemapLight '#107C10')
    }

    foreach ($PT in $Global:PhaseTimings) {
        $End = if ($PT.EndTime) { $PT.EndTime } else { $Now }
        $Duration = ($End - $PT.StartTime).TotalSeconds
        $OffsetPx = (($PT.StartTime - $EarliestStart).TotalSeconds / $TotalSpan) * $ChartWidth
        $WidthPx = ($Duration / $TotalSpan) * $ChartWidth
        $WidthPx = [Math]::Max($WidthPx, 3)

        $Container = New-Object System.Windows.Controls.Grid
        $Container.Margin = [System.Windows.Thickness]::new(0, 1, 0, 1)
        $Container.Height = 16

        $PhaseClr = if ($PhaseColors.ContainsKey($PT.Name)) { $PhaseColors[$PT.Name] } else { & $Global:RemapLight '#0078D4' }

        $Bar = New-Object System.Windows.Shapes.Rectangle
        $Bar.RadiusX = 2; $Bar.RadiusY = 2
        $Bar.Fill = $BC.ConvertFromString($PhaseClr)
        $Bar.Opacity = 0.8
        $Bar.Width = $WidthPx
        $Bar.HorizontalAlignment = 'Left'
        $Bar.Margin = [System.Windows.Thickness]::new($OffsetPx, 0, 0, 0)

        $Label = New-Object System.Windows.Controls.TextBlock
        $DurText = if ($Duration -ge 60) { "{0:0}m{1:00}s" -f [int]($Duration/60), [int]($Duration%60) }
                   else { "{0:0}s" -f [int]$Duration }
        $Label.Text = "$($PT.Name) ($DurText)"
        $Label.FontSize = 8.5
        $Label.Foreground = $Window.Resources['ThemeTextDim']
        $Label.VerticalAlignment = 'Center'
        $Label.Margin = [System.Windows.Thickness]::new([Math]::Max(0, $OffsetPx + $WidthPx + 4), 0, 0, 0)
        $Label.HorizontalAlignment = 'Left'

        $Container.Children.Add($Bar) | Out-Null
        $Container.Children.Add($Label) | Out-Null
        $pnlFlameChart.Children.Add($Container) | Out-Null
    }
}

function Reset-FlameChart {
    $Global:PhaseTimings.Clear()
    $Global:FlameChartLastPhase = $null
    $Global:FlameChartLastTime  = $null
    if ($pnlFlameChart) { $pnlFlameChart.Children.Clear() }
}

# ==============================================================================
# SECTION 7S: BUILD HEALTH SCORE (#R2-1)
# ==============================================================================

$Global:HealthScore = 'N/A'
$Global:HealthGrade = 'N/A'

function Get-HealthGrade {
    param([double]$Score)
    if ($Score -ge 95) { return @{ Grade = 'A+'; Color = (& $Global:RemapLight '#107C10') } }
    if ($Score -ge 90) { return @{ Grade = 'A';  Color = (& $Global:RemapLight '#107C10') } }
    if ($Score -ge 80) { return @{ Grade = 'B';  Color = (& $Global:RemapLight '#3B7D23') } }
    if ($Score -ge 70) { return @{ Grade = 'C';  Color = (& $Global:RemapLight '#FFB900') } }
    if ($Score -ge 60) { return @{ Grade = 'D';  Color = (& $Global:RemapLight '#FF8C00') } }
    return @{ Grade = 'F'; Color = (& $Global:RemapLight '#D13438') }
}

function Update-HealthScore {
    <#
    .SYNOPSIS
        Computes a 0-100 build health score from error/warning/line ratios and historical duration.
    #>
    param([DateTime]$Now = (Get-Date))
    $Stats = $Global:LogStats
    $Total = $Stats.TotalLines
    if ($Total -lt 10) { return }

    # Base: start at 100, deduct for issues
    $Score = 100.0

    # Error penalty: -5 per error (capped at -40)
    $ErrPenalty = [Math]::Min($Stats.ErrorCount * 5, 40)
    $Score -= $ErrPenalty

    # Warning penalty: -1 per warning (capped at -20)
    $WarnPenalty = [Math]::Min($Stats.WarningCount * 1, 20)
    $Score -= $WarnPenalty

    # Error ratio penalty: if >5% errors, extra -10
    if ($Total -gt 0 -and ($Stats.ErrorCount / $Total) -gt 0.05) { $Score -= 10 }

    # Duration anomaly: compare to historical avg
    if ($Global:SyncHash.BuildStartTime) {
        $Elapsed = ($Now.ToUniversalTime() - $Global:SyncHash.BuildStartTime).TotalMinutes
        $HistAvg = Get-HistoricalAvgDuration
        if ($HistAvg -gt 0 -and $Elapsed -gt ($HistAvg * 1.5)) {
            $Score -= 10  # Running much longer than usual
        }
    }

    $Score = [Math]::Max(0, [Math]::Min(100, $Score))
    $Global:HealthScore = [Math]::Round($Score, 0)
    $Global:HealthGrade = Get-HealthGrade $Score

    # Update UI
    if ($lblHealthGrade -and $lblHealthScore) {
        $BC = $Global:CachedBC
        $lblHealthGrade.Text = $Global:HealthGrade.Grade
        try { $lblHealthGrade.Foreground = $BC.ConvertFromString($Global:HealthGrade.Color) } catch { }
        $lblHealthScore.Text = "$($Global:HealthScore)/100"
    }
}

function Get-HistoricalAvgDuration {
    <# Returns average build duration in minutes from build history (cached 30s). #>
    if ($Global:HistAvgCache -and $Global:HistAvgCacheTime -and
        ((Get-Date) - $Global:HistAvgCacheTime).TotalSeconds -lt 30) {
        return $Global:HistAvgCache
    }
    $Completed = $Global:BuildHistory | Where-Object { $_.FinalStatus -eq 'Completed' -and $_.DurationMin -gt 0 }
    if ($Completed.Count -eq 0) { $Global:HistAvgCache = 0; $Global:HistAvgCacheTime = Get-Date; return 0 }
    $Sum = ($Completed | ForEach-Object { $_.DurationMin } | Measure-Object -Sum).Sum
    $Global:HistAvgCache = $Sum / $Completed.Count
    $Global:HistAvgCacheTime = Get-Date
    return $Global:HistAvgCache
}

# ==============================================================================
# SECTION 7T: ANOMALY DETECTION (#R2-3)
# ==============================================================================

# Phase anomaly detection removed — functions were dead code (C1/C2).
# Get-PhaseHistoricalAvg and Test-PhaseAnomaly were defined but never called.
# Anomaly data is still captured in build history via Add-BuildToHistory.$PhaseDurations.

# ==============================================================================
# SECTION 7U: LIVE ETA & PROGRESS (#R2-4)
# ==============================================================================

$Global:ETAPercent = 0
$Global:ETARemaining = ''

function Update-ETA {
    <# Estimates build progress % and remaining time based on historical data #>
    param([DateTime]$Now = (Get-Date))
    if (-not $Global:SyncHash.BuildStartTime) { return }

    $ElapsedMin = ($Now.ToUniversalTime() - $Global:SyncHash.BuildStartTime).TotalMinutes
    $AvgDur = Get-HistoricalAvgDuration

    if ($AvgDur -gt 0) {
        $Pct = [Math]::Min(99, [Math]::Round(($ElapsedMin / $AvgDur) * 100, 0))
        $Global:ETAPercent = $Pct
        $RemainingMin = [Math]::Max(0, $AvgDur - $ElapsedMin)
        if ($RemainingMin -lt 1) {
            $Global:ETARemaining = '<1 min left'
        } else {
            $Global:ETARemaining = "$([Math]::Round($RemainingMin, 0)) min left"
        }
    } else {
        # No history — show elapsed only, no ETA
        $Global:ETAPercent = 0
        $Global:ETARemaining = 'No data for ETA'
    }

    # Update UI
    if ($lblETAPercent) { $lblETAPercent.Text = "$($Global:ETAPercent)%" }
    if ($lblETARemaining) { $lblETARemaining.Text = $Global:ETARemaining }
    if ($prgETA) {
        $prgETA.Value = $Global:ETAPercent
    }
}

# ==============================================================================
# SECTION 7V: PACKAGE INSTALL TRACKER (#R2-5)
# ==============================================================================

$Global:PackageTracker = [System.Collections.Generic.Dictionary[string, PSObject]]::new()

$Global:PkgPatterns = @{
    WinGetStart    = 'Found\s+(.+?)\s+\[([^\]]+)\]\s+Version'
    CustomStart    = '->\s*Installing Package:\s+(\S+)'
    WinGetInstall  = 'Installing\s+(.*)'
    WinGetSuccess  = 'Successfully installed|->\s*SUCCESS\.|Restart your PC to finish installation'
    WinGetFailed   = 'Install failed|Installation abandoned|Package .* not found|->\s*FAILED|Installer hash does not match'
    WinGetSkip     = 'No applicable update found|Already installed|No newer version'
    WinGetDownload = '(\d+)%\s*\|'
    ChocoInstall   = 'choco install\s+(\S+)'
    ChocoSuccess   = 'The install of .+ was successful'
    ChocoFailed    = 'The install of .+ was NOT successful|ERROR'
}

function Update-PackageTracker {
    <# Parses a log line to track package install status #>
    param([string]$Line)

    # Extract log timestamp (from the packer log line), fall back to wall clock
    $LogTime = Get-Date
    $TsMatch = $Global:TimestampRegex.Match($Line)
    if ($TsMatch.Success) {
        $TsRaw = if ($TsMatch.Groups[1].Success) { $TsMatch.Groups[1].Value } else { $TsMatch.Groups[2].Value }
        try { $LogTime = [DateTime]::Parse($TsRaw, [System.Globalization.CultureInfo]::InvariantCulture) } catch { }
    }

    # Custom script package discovery  (-> Installing Package: <id>)
    # This line fires BEFORE WinGet's 'Found' line, so we add the package early
    if ($Line -match $Global:PkgPatterns.CustomStart) {
        $PkgId = $Matches[1].Trim() -replace '\s*\(.*$', ''   # strip any trailing override text
        $Key = $PkgId
        if (-not $Global:PackageTracker.ContainsKey($Key)) {
            $Global:PackageTracker[$Key] = [PSCustomObject]@{
                Name    = $PkgId       # Will be updated by WinGet 'Found' line below
                Id      = $PkgId
                Status  = 'Pending'
                Icon    = [char]0x23F3   # hourglass
                StartAt = $LogTime
                EndAt   = $null
                Pct     = 0
            }
        }
        return
    }

    # WinGet package discovery  (Found <name> [<id>] Version <ver>)
    if ($Line -match $Global:PkgPatterns.WinGetStart) {
        $PkgName = $Matches[1].Trim()
        $PkgId = $Matches[2].Trim()
        $Key = $PkgId
        if ($Global:PackageTracker.ContainsKey($Key)) {
            # Update friendly name from custom-script discovery
            $Global:PackageTracker[$Key].Name = $PkgName
        } else {
            $Global:PackageTracker[$Key] = [PSCustomObject]@{
                Name    = $PkgName
                Id      = $PkgId
                Status  = 'Pending'
                Icon    = [char]0x23F3   # hourglass
                StartAt = $LogTime
                EndAt   = $null
                Pct     = 0
            }
        }
        return
    }

    # WinGet download progress
    if ($Line -match $Global:PkgPatterns.WinGetDownload) {
        $Pct = [int]$Matches[1]
        $LastPkg = $Global:PackageTracker.Values | Where-Object { $_.Status -eq 'Downloading' -or $_.Status -eq 'Pending' } | Select-Object -Last 1
        if ($LastPkg) {
            $LastPkg.Status = 'Downloading'
            $LastPkg.Icon = [char]0x2B07  # down arrow
            $LastPkg.Pct = $Pct
        }
        return
    }

    # WinGet installing  (match 'Starting package install' but NOT '-> Installing Package:')
    if ($Line -match 'Starting package install' -or ($Line -match 'Installing' -and $Line -notmatch '->\s*Installing Package:')) {
        $LastPkg = $Global:PackageTracker.Values | Where-Object { $_.Status -in @('Pending','Downloading') } | Select-Object -Last 1
        if ($LastPkg) {
            $LastPkg.Status = 'Installing'
            $LastPkg.Icon = [char]0x2699  # gear
            $LastPkg.Pct = 0
        }
        return
    }

    # WinGet success
    if ($Line -match $Global:PkgPatterns.WinGetSuccess) {
        $LastPkg = $Global:PackageTracker.Values | Where-Object { $_.Status -in @('Pending','Downloading','Installing') } | Select-Object -Last 1
        if ($LastPkg) {
            $LastPkg.Status = 'Installed'
            $LastPkg.Icon = [char]0x2705  # check
            $LastPkg.EndAt = $LogTime
        }
        return
    }

    # WinGet failed
    if ($Line -match $Global:PkgPatterns.WinGetFailed) {
        $LastPkg = $Global:PackageTracker.Values | Where-Object { $_.Status -in @('Pending','Downloading','Installing') } | Select-Object -Last 1
        if ($LastPkg) {
            $LastPkg.Status = 'Failed'
            $LastPkg.Icon = [char]0x274C  # cross
            $LastPkg.EndAt = $LogTime
        }
        return
    }

    # WinGet skipped
    if ($Line -match $Global:PkgPatterns.WinGetSkip) {
        $LastPkg = $Global:PackageTracker.Values | Where-Object { $_.Status -eq 'Pending' } | Select-Object -Last 1
        if ($LastPkg) {
            $LastPkg.Status = 'Skipped'
            $LastPkg.Icon = [char]0x23ED  # skip
            $LastPkg.EndAt = $LogTime
        }
        return
    }
}

function Render-PackageTracker {
    <# Renders the package tracker panel #>
    if (-not $pnlPackageTracker) { return }
    $pnlPackageTracker.Children.Clear()
    $BC = $Global:CachedBC

    if ($Global:PackageTracker.Count -eq 0) {
        $Empty = New-Object System.Windows.Controls.TextBlock
        $Empty.Text = "No packages detected yet"
        $Empty.FontSize = 10
        $Empty.FontStyle = 'Italic'
        $Empty.Foreground = $Window.Resources['ThemeTextDim']
        $pnlPackageTracker.Children.Add($Empty) | Out-Null
        return
    }

    foreach ($Pkg in $Global:PackageTracker.Values) {
        $Row = New-Object System.Windows.Controls.Grid
        $Row.Margin = [System.Windows.Thickness]::new(0, 0, 0, 2)
        $Col0 = New-Object System.Windows.Controls.ColumnDefinition; $Col0.Width = [System.Windows.GridLength]::new(20)
        $Col1 = New-Object System.Windows.Controls.ColumnDefinition; $Col1.Width = [System.Windows.GridLength]::new(1, 'Star')
        $Col2 = New-Object System.Windows.Controls.ColumnDefinition; $Col2.Width = [System.Windows.GridLength]::Auto
        $Row.ColumnDefinitions.Add($Col0)
        $Row.ColumnDefinitions.Add($Col1)
        $Row.ColumnDefinitions.Add($Col2)

        # Status icon
        $IconTxt = New-Object System.Windows.Controls.TextBlock
        $IconTxt.Text = [string]$Pkg.Icon
        $IconTxt.FontSize = 11
        $IconTxt.VerticalAlignment = 'Center'
        [System.Windows.Controls.Grid]::SetColumn($IconTxt, 0)

        # Package name
        $NameTxt = New-Object System.Windows.Controls.TextBlock
        $NameTxt.Text = $Pkg.Name
        $NameTxt.FontSize = 10
        $NameTxt.TextTrimming = 'CharacterEllipsis'
        $NameTxt.VerticalAlignment = 'Center'
        $NameTxt.Foreground = $Window.Resources['ThemeTextBody']
        [System.Windows.Controls.Grid]::SetColumn($NameTxt, 1)

        # Duration or progress
        $DurTxt = New-Object System.Windows.Controls.TextBlock
        $DurTxt.FontSize = 9
        $DurTxt.FontFamily = [System.Windows.Media.FontFamily]::new("Consolas")
        $DurTxt.VerticalAlignment = 'Center'
        $DurTxt.Foreground = $Window.Resources['ThemeTextDim']
        [System.Windows.Controls.Grid]::SetColumn($DurTxt, 2)

        switch ($Pkg.Status) {
            'Downloading' { $DurTxt.Text = "$($Pkg.Pct)%"; $DurTxt.Foreground = $Window.Resources['ThemeAccent'] }
            'Installing'  { $DurTxt.Text = "..."; $DurTxt.Foreground = $Window.Resources['ThemeWarning'] }
            'Installed'   {
                if ($Pkg.EndAt -and $Pkg.StartAt) {
                    $Dur = ($Pkg.EndAt - $Pkg.StartAt).TotalSeconds
                    $DurTxt.Text = "$([Math]::Round($Dur,0))s"
                } else { $DurTxt.Text = "done" }
                $DurTxt.Foreground = $Window.Resources['ThemeSuccess']
            }
            'Failed'      { $DurTxt.Text = "FAIL"; $DurTxt.Foreground = $Window.Resources['ThemeError'] }
            'Skipped'     { $DurTxt.Text = "skip"; $DurTxt.Foreground = $Window.Resources['ThemeTextDim'] }
            default       { $DurTxt.Text = "" }
        }

        $Row.Children.Add($IconTxt) | Out-Null
        $Row.Children.Add($NameTxt) | Out-Null
        $Row.Children.Add($DurTxt) | Out-Null
        $pnlPackageTracker.Children.Add($Row) | Out-Null
    }
}

function Reset-PackageTracker {
    $Global:PackageTracker.Clear()
    if ($pnlPackageTracker) { $pnlPackageTracker.Children.Clear() }
}

# ==============================================================================
# SECTION 7W: ERROR CLUSTERING (#R2-6)
# ==============================================================================

$Global:ErrorClusters = [System.Collections.Generic.Dictionary[string, PSObject]]::new()

function Add-ErrorCluster {
    <# Groups identical error messages with count tracking #>
    param([string]$ErrorText)
    # Normalize: strip timestamps, line numbers, hex addresses
    $Normalized = $ErrorText -replace '\d{4}-\d{2}-\d{2}T?\d{2}:\d{2}:\d{2}[^\s]*', '' `
                             -replace '0x[0-9A-Fa-f]+', '0x...' `
                             -replace '\b\d{5,}\b', 'N' `
                             -replace '\s+', ' '
    $Normalized = $Normalized.Trim()
    if ($Normalized.Length -lt 5) { return }

    if ($Global:ErrorClusters.ContainsKey($Normalized)) {
        $Global:ErrorClusters[$Normalized].Count++
        $Global:ErrorClusters[$Normalized].LastSeen = (Get-Date).ToString('HH:mm:ss')
        $Global:ErrorClusters[$Normalized].LastLine = $Global:LogStats.TotalLines
    } else {
        $Global:ErrorClusters[$Normalized] = [PSCustomObject]@{
            Message   = $ErrorText.Substring(0, [Math]::Min(120, $ErrorText.Length))
            Count     = 1
            FirstSeen = (Get-Date).ToString('HH:mm:ss')
            LastSeen  = (Get-Date).ToString('HH:mm:ss')
            FirstLine = $Global:LogStats.TotalLines
            LastLine  = $Global:LogStats.TotalLines
        }
    }
}

function Render-ErrorClusters {
    <# Renders grouped errors in the error cluster panel #>
    if (-not $pnlErrorClusters) { return }
    $pnlErrorClusters.Children.Clear()
    $BC = $Global:CachedBC

    if ($Global:ErrorClusters.Count -eq 0) {
        $Empty = New-Object System.Windows.Controls.TextBlock
        $Empty.Text = "No errors detected"
        $Empty.FontSize = 10
        $Empty.FontStyle = 'Italic'
        $Empty.Foreground = $Window.Resources['ThemeSuccess']
        $pnlErrorClusters.Children.Add($Empty) | Out-Null
        return
    }

    $Sorted = $Global:ErrorClusters.Values | Sort-Object Count -Descending | Select-Object -First 8
    foreach ($Cluster in $Sorted) {
        $Card = New-Object System.Windows.Controls.Border
        $Card.Background = $Window.Resources['ThemeDeepBg']
        $Card.CornerRadius = [System.Windows.CornerRadius]::new(4)
        $Card.Padding = [System.Windows.Thickness]::new(6, 4, 6, 4)
        $Card.Margin = [System.Windows.Thickness]::new(0, 0, 0, 3)

        $Inner = New-Object System.Windows.Controls.StackPanel

        # Count badge + message
        $TopRow = New-Object System.Windows.Controls.StackPanel
        $TopRow.Orientation = 'Horizontal'

        $CountBadge = New-Object System.Windows.Controls.TextBlock
        $CountBadge.Text = "x$($Cluster.Count)"
        $CountBadge.FontSize = 9
        $CountBadge.FontWeight = 'Bold'
        $CountBadge.Foreground = $Window.Resources['ThemeError']
        $CountBadge.Margin = [System.Windows.Thickness]::new(0, 0, 6, 0)
        $CountBadge.VerticalAlignment = 'Center'

        $MsgTxt = New-Object System.Windows.Controls.TextBlock
        $MsgTxt.Text = $Cluster.Message
        $MsgTxt.FontSize = 9.5
        $MsgTxt.TextTrimming = 'CharacterEllipsis'
        $MsgTxt.Foreground = $Window.Resources['ThemeTextBody']

        $TopRow.Children.Add($CountBadge) | Out-Null
        $TopRow.Children.Add($MsgTxt) | Out-Null

        # Time range
        $TimeTxt = New-Object System.Windows.Controls.TextBlock
        $TimeTxt.FontSize = 8.5
        $TimeTxt.Foreground = $Window.Resources['ThemeTextDim']
        if ($Cluster.Count -gt 1) {
            $TimeTxt.Text = "Lines $($Cluster.FirstLine)-$($Cluster.LastLine) | $($Cluster.FirstSeen)-$($Cluster.LastSeen)"
        } else {
            $TimeTxt.Text = "Line $($Cluster.FirstLine) | $($Cluster.FirstSeen)"
        }

        $Inner.Children.Add($TopRow) | Out-Null
        $Inner.Children.Add($TimeTxt) | Out-Null
        $Card.Child = $Inner
        $pnlErrorClusters.Children.Add($Card) | Out-Null
    }
}

function Reset-ErrorClusters {
    $Global:ErrorClusters.Clear()
    if ($pnlErrorClusters) { $pnlErrorClusters.Children.Clear() }
}

# ==============================================================================
# SECTION 7X: KNOWN ISSUE AUTO-TAGGER (#R2-8)
# ==============================================================================

$Global:KnownIssues = @(
    @{
        Pattern = 'HTTP 429|Too Many Requests|rate.?limit'
        Title   = 'WinGet Rate Limited (HTTP 429)'
        Fix     = 'CDN/WinGet is throttling. Wait 10-15 min or configure a private WinGet source.'
        Link    = 'https://learn.microsoft.com/windows/package-manager/winget/source'
    }
    @{
        Pattern = '0x800f0954|CBS_E_STORE'
        Title   = 'DISM/CBS Feature Store Error'
        Fix     = 'Enable WSUS fallback or configure Features on Demand source. Check network access to Windows Update.'
        Link    = 'https://learn.microsoft.com/troubleshoot/windows-server/deployment/error-0x800f0954-install-features'
    }
    @{
        Pattern = 'sysprep.*failed|sysprep.*error|Sysprep was not able'
        Title   = 'Sysprep Failure'
        Fix     = 'Check AppX provisioned packages, auto-update settings, and Panther logs. Remove problematic UWP apps before generalize.'
        Link    = 'https://learn.microsoft.com/windows-hardware/manufacture/desktop/sysprep--generalize--a-windows-installation'
    }
    @{
        Pattern = 'AzureADLogin.*failed|AADLoginForWindows.*error'
        Title   = 'Azure AD Login Extension Failure'
        Fix     = 'Ensure VM identity is configured and AADLoginForWindows extension is compatible with OS version.'
        Link    = 'https://learn.microsoft.com/azure/active-directory/devices/howto-vm-sign-in-azure-ad-windows'
    }
    @{
        Pattern = 'Could not resolve host|Name or service not known|NetworkError|getaddrinfo'
        Title   = 'DNS/Network Resolution Failure'
        Fix     = 'Check NSG rules, Private DNS zones, and ensure packer staging VNET has internet access for package downloads.'
        Link    = 'https://learn.microsoft.com/azure/virtual-machines/image-builder-networking'
    }
    @{
        Pattern = 'Access is denied|HTTP 403|Forbidden'
        Title   = 'Access Denied / 403'
        Fix     = 'Verify managed identity has Contributor + Storage Blob Data Reader on staging RG. Check RBAC assignments.'
        Link    = 'https://learn.microsoft.com/azure/virtual-machines/image-builder-permissions-powershell'
    }
    @{
        Pattern = 'Timeout|timed out|Operation timed out|deadline exceeded'
        Title   = 'Operation Timeout'
        Fix     = 'Increase buildTimeoutInMinutes in packer/AIB template. Default is 240 min for large images.'
        Link    = 'https://learn.microsoft.com/azure/virtual-machines/image-builder-json#properties-buildtimeoutinminutes'
    }
    @{
        Pattern = 'No space left on device|disk is full|InsufficientDiskSpace'
        Title   = 'Disk Space Exhausted'
        Fix     = 'Increase OS disk size in packer template or clean temp files before sysprep. Consider 128GB+ for large image builds.'
        Link    = 'https://learn.microsoft.com/azure/virtual-machines/image-builder-json#properties-vmprofile'
    }
    @{
        Pattern = 'Restart-Computer|reboot pending|PendingReboot'
        Title   = 'Pending Reboot Detected'
        Fix     = 'A restart was requested or is pending. Ensure packer restarts are handled with windows-restart provisioner.'
        Link    = 'https://developer.hashicorp.com/packer/integrations/hashicorp/windows/latest/components/provisioner/windows-restart'
    }
    @{
        Pattern = 'SharedImageGallery.*error|Image version.*failed'
        Title   = 'Gallery Image Publish Failure'
        Fix     = 'Verify gallery exists, replica count, and managed identity permissions on Compute Gallery resource.'
        Link    = 'https://learn.microsoft.com/azure/virtual-machines/image-builder-gallery'
    }
)

$Global:TaggedIssues = [System.Collections.Generic.List[PSObject]]::new()

# Pre-compile known issue patterns
foreach ($KI in $Global:KnownIssues) {
    $KI['CompiledRx'] = [regex]::new($KI.Pattern, 'Compiled, IgnoreCase')
}

function Test-KnownIssue {
    <# Checks a line against pre-compiled known issue patterns, returns match or $null #>
    param([string]$Line)
    foreach ($Issue in $Global:KnownIssues) {
        if ($Issue.CompiledRx.IsMatch($Line)) {
            return $Issue
        }
    }
    return $null
}

function Add-KnownIssueTag {
    <# Records a matched known issue occurrence #>
    param([PSCustomObject]$Issue, [int]$LineNum)
    $Existing = $Global:TaggedIssues | Where-Object { $_.Title -eq $Issue.Title }
    if ($Existing) {
        $Existing.Count++
        $Existing.LastLine = $LineNum
    } else {
        $Global:TaggedIssues.Add([PSCustomObject]@{
            Title    = $Issue.Title
            Fix      = $Issue.Fix
            Link     = $Issue.Link
            Count    = 1
            FirstLine = $LineNum
            LastLine  = $LineNum
        })
    }
}

function Render-KnownIssues {
    <# Renders known issue tags in the sidebar panel #>
    if (-not $pnlKnownIssues) { return }
    $pnlKnownIssues.Children.Clear()
    $BC = $Global:CachedBC

    if ($Global:TaggedIssues.Count -eq 0) { return }

    foreach ($Tag in $Global:TaggedIssues) {
        $Card = New-Object System.Windows.Controls.Border
        $Card.Background = if ($Global:IsLightMode) { $BC.ConvertFromString("#FFF0F0") } else { $BC.ConvertFromString("#1A0D0D") }
        $Card.CornerRadius = [System.Windows.CornerRadius]::new(4)
        $Card.Padding = [System.Windows.Thickness]::new(7, 5, 7, 5)
        $Card.Margin = [System.Windows.Thickness]::new(0, 0, 0, 4)
        $Card.BorderBrush = if ($Global:IsLightMode) { $BC.ConvertFromString("#E0C0C0") } else { $BC.ConvertFromString("#3D1F1F") }
        $Card.BorderThickness = [System.Windows.Thickness]::new(1)

        $Inner = New-Object System.Windows.Controls.StackPanel

        # Title with info icon
        $TitleRow = New-Object System.Windows.Controls.StackPanel
        $TitleRow.Orientation = 'Horizontal'

        $InfoIcon = New-Object System.Windows.Controls.TextBlock
        $InfoIcon.Text = [string][char]0x24D8  # circled i
        $InfoIcon.FontSize = 12
        $InfoIcon.Foreground = $Window.Resources['ThemeWarning']
        $InfoIcon.Margin = [System.Windows.Thickness]::new(0, 0, 5, 0)
        $InfoIcon.VerticalAlignment = 'Center'

        $TitleTxt = New-Object System.Windows.Controls.TextBlock
        $TitleTxt.Text = "$($Tag.Title) (x$($Tag.Count))"
        $TitleTxt.FontSize = 10
        $TitleTxt.FontWeight = 'SemiBold'
        $TitleTxt.Foreground = $Window.Resources['ThemeWarning']
        $TitleTxt.TextWrapping = 'Wrap'

        $TitleRow.Children.Add($InfoIcon) | Out-Null
        $TitleRow.Children.Add($TitleTxt) | Out-Null

        # Fix suggestion
        $FixTxt = New-Object System.Windows.Controls.TextBlock
        $FixTxt.Text = $Tag.Fix
        $FixTxt.FontSize = 9
        $FixTxt.TextWrapping = 'Wrap'
        $FixTxt.Foreground = $Window.Resources['ThemeTextMuted']
        $FixTxt.Margin = [System.Windows.Thickness]::new(0, 2, 0, 0)

        $Inner.Children.Add($TitleRow) | Out-Null
        $Inner.Children.Add($FixTxt) | Out-Null
        $Card.Child = $Inner

        # Tooltip with link
        $Card.ToolTip = "MS Learn: $($Tag.Link)"

        $pnlKnownIssues.Children.Add($Card) | Out-Null
    }
}

function Reset-KnownIssues {
    $Global:TaggedIssues.Clear()
    if ($pnlKnownIssues) { $pnlKnownIssues.Children.Clear() }
}

# ==============================================================================
# SECTION 7Y: BUILD STREAK & CONFETTI (#R2-9)
# ==============================================================================

$Global:BuildStreak = 0
$Global:ConfettiActive = $false

function Get-BuildStreak {
    <# Computes current consecutive successful build count from history #>
    $Streak = 0
    foreach ($H in $Global:BuildHistory) {
        if ($H.FinalStatus -eq 'Completed') { $Streak++ }
        else { break }
    }
    return $Streak
}

function Update-StreakDisplay {
    $Global:BuildStreak = Get-BuildStreak
    if ($lblStreak) {
        if ($Global:BuildStreak -gt 0) {
            $lblStreak.Text = [char]::ConvertFromUtf32(0x1F525) + " $($Global:BuildStreak) build streak!"
            if ($Global:BuildStreak -ge 10) {
                $lblStreak.Foreground = $Global:CachedBC.ConvertFromString((& $Global:RemapLight '#FF8C00'))
            } elseif ($Global:BuildStreak -ge 5) {
                $lblStreak.Foreground = $Global:CachedBC.ConvertFromString((& $Global:RemapLight '#FFB900'))
            } else {
                $lblStreak.Foreground = $Window.Resources['ThemeTextMuted']
            }
        } else {
            $lblStreak.Text = ""
        }
    }
}

function Start-ConfettiAnimation {
    <# Spawns confetti particles on the canvas #>
    if (-not $cnvConfetti -or $Global:ConfettiActive) { return }
    # E3: Guard against confetti when window has no width (minimized/zero-size)
    if ($Window.ActualWidth -lt 10) { return }
    $Global:ConfettiActive = $true
    $cnvConfetti.Visibility = 'Visible'
    $cnvConfetti.Children.Clear()
    $BC = $Global:CachedBC

    $Colors = @('#0078D4','#60CDFF','#107C10','#FFB900','#FF8C00','#D13438','#8764B8','#E74856','#00CC6A')
    $Rnd = New-Object System.Random

    for ($i = 0; $i -lt 60; $i++) {
        $Rect = New-Object System.Windows.Shapes.Rectangle
        $Rect.Width = $Rnd.Next(4, 10)
        $Rect.Height = $Rnd.Next(4, 10)
        $Rect.Fill = $BC.ConvertFromString($Colors[$Rnd.Next(0, $Colors.Count)])
        $Rect.Opacity = 0.9
        $StartX = $Rnd.Next(0, [int]$Window.ActualWidth)
        [System.Windows.Controls.Canvas]::SetLeft($Rect, $StartX)
        [System.Windows.Controls.Canvas]::SetTop($Rect, -20)
        $cnvConfetti.Children.Add($Rect) | Out-Null

        # Animate fall
        $FallAnim = New-Object System.Windows.Media.Animation.DoubleAnimation
        $FallAnim.From = -20
        $FallAnim.To = $Window.ActualHeight + 20
        $FallAnim.Duration = [TimeSpan]::FromSeconds($Rnd.NextDouble() * 2 + 1.5)
        $FallAnim.BeginTime = [TimeSpan]::FromMilliseconds($Rnd.Next(0, 800))
        $Rect.BeginAnimation([System.Windows.Controls.Canvas]::TopProperty, $FallAnim)

        # Animate sideways drift
        $DriftAnim = New-Object System.Windows.Media.Animation.DoubleAnimation
        $DriftAnim.From = $StartX
        $DriftAnim.To = $StartX + $Rnd.Next(-100, 100)
        $DriftAnim.Duration = $FallAnim.Duration
        $DriftAnim.BeginTime = $FallAnim.BeginTime
        $Rect.BeginAnimation([System.Windows.Controls.Canvas]::LeftProperty, $DriftAnim)
    }

    # Cleanup timer
    $ConfettiCleanup = New-Object System.Windows.Threading.DispatcherTimer
    $ConfettiCleanup.Interval = [TimeSpan]::FromSeconds(4)
    $ConfettiCleanup.Add_Tick({
        $cnvConfetti.Children.Clear()
        $cnvConfetti.Visibility = 'Collapsed'
        $Global:ConfettiActive = $false
        ($this).Stop()
    })
    $ConfettiCleanup.Start()
}

# ==============================================================================
# SECTION 7Z: ACHIEVEMENT BADGES (#R2-10)
# ==============================================================================

$Global:AchievementsPath = Join-Path $Global:Root "achievements.json"
$Global:Achievements = @{}

$Global:AchievementDefs = @(
    @{ Id = 'first_build';    Name = 'First Build';       Desc = 'Monitored your first build';           Icon = [char]::ConvertFromUtf32(0x1F680) }
    @{ Id = 'zero_errors';    Name = 'Clean Build';       Desc = 'Completed a build with zero errors';   Icon = [char]::ConvertFromUtf32(0x2728) }
    @{ Id = 'speed_demon';    Name = 'Speed Demon';       Desc = 'Build completed in under 30 minutes';  Icon = [char]::ConvertFromUtf32(0x26A1) }
    @{ Id = 'night_owl';      Name = 'Night Owl';         Desc = 'Monitored a build after midnight';     Icon = [char]::ConvertFromUtf32(0x1F989) }
    @{ Id = 'marathon';        Name = 'Marathon Runner';   Desc = 'Survived a 2+ hour build';             Icon = [char]::ConvertFromUtf32(0x1F3C3) }
    @{ Id = 'streak_5';       Name = 'On Fire';           Desc = '5 consecutive successful builds';       Icon = [char]::ConvertFromUtf32(0x1F525) }
    @{ Id = 'streak_10';      Name = 'Unstoppable';       Desc = '10 consecutive successful builds';      Icon = [char]::ConvertFromUtf32(0x1F4AA) }
    @{ Id = 'streak_25';      Name = 'Legendary';         Desc = '25 consecutive successful builds';      Icon = [char]::ConvertFromUtf32(0x1F3C6) }
    @{ Id = 'hundred_builds'; Name = 'Centurion';         Desc = 'Monitored 100 builds total';            Icon = [char]::ConvertFromUtf32(0x1F4AF) }
    @{ Id = 'error_hunter';   Name = 'Error Hunter';      Desc = 'Encountered 50+ errors in one build';   Icon = [char]::ConvertFromUtf32(0x1F50D) }
)

function Load-Achievements {
    if (Test-Path $Global:AchievementsPath) {
        try {
            $Json = Get-Content $Global:AchievementsPath -Raw | ConvertFrom-Json
            $Global:Achievements = @{}
            foreach ($Prop in $Json.PSObject.Properties) {
                $Global:Achievements[$Prop.Name] = $Prop.Value
            }
        } catch { $Global:Achievements = @{} }
    }
}

function Save-Achievements {
    try {
        $Global:Achievements | ConvertTo-Json -Depth 3 | Set-Content $Global:AchievementsPath -Force
    } catch { }
}

function Unlock-Achievement {
    param([string]$Id)
    if ($Global:Achievements[$Id]) { return }  # Already unlocked
    $Def = $Global:AchievementDefs | Where-Object { $_.Id -eq $Id }
    if (-not $Def) { return }

    $Global:Achievements[$Id] = (Get-Date).ToString('o')
    Save-Achievements

    # Toast notification
    Show-ToastNotification "Achievement Unlocked!" "$($Def.Icon) $($Def.Name) — $($Def.Desc)"

    # Update badge shelf
    Render-Achievements
}

function Check-BuildAchievements {
    <# Called on build completion to check and unlock achievements #>
    param([string]$Status)

    # First build
    Unlock-Achievement 'first_build'

    # Clean build
    if ($Status -eq 'Completed' -and $Global:LogStats.ErrorCount -eq 0) {
        Unlock-Achievement 'zero_errors'
    }

    # Speed demon
    if ($Status -eq 'Completed' -and $Global:SyncHash.BuildStartTime) {
        $Dur = ((Get-Date).ToUniversalTime() - $Global:SyncHash.BuildStartTime).TotalMinutes
        if ($Dur -lt 30) { Unlock-Achievement 'speed_demon' }
        if ($Dur -ge 120) { Unlock-Achievement 'marathon' }
    }

    # Night owl
    if ((Get-Date).Hour -lt 5 -or (Get-Date).Hour -ge 23) {
        Unlock-Achievement 'night_owl'
    }

    # Streaks
    $Streak = Get-BuildStreak
    if ($Streak -ge 5) { Unlock-Achievement 'streak_5' }
    if ($Streak -ge 10) { Unlock-Achievement 'streak_10' }
    if ($Streak -ge 25) { Unlock-Achievement 'streak_25' }

    # Centurion
    if ($Global:BuildHistory.Count -ge 100) { Unlock-Achievement 'hundred_builds' }

    # Error hunter
    if ($Global:LogStats.ErrorCount -ge 50) { Unlock-Achievement 'error_hunter' }
}

function Render-Achievements {
    if (-not $pnlAchievements) { return }
    $pnlAchievements.Children.Clear()
    $BC = $Global:CachedBC

    $Unlocked = $Global:AchievementDefs | Where-Object { $Global:Achievements[$_.Id] }
    $Locked   = $Global:AchievementDefs | Where-Object { -not $Global:Achievements[$_.Id] }

    # Unlocked badges
    foreach ($Def in $Unlocked) {
        $Badge = New-Object System.Windows.Controls.Border
        $Badge.Background = if ($Global:IsLightMode) { $BC.ConvertFromString("#E8F5E8") } else { $BC.ConvertFromString("#1A2F1A") }
        $Badge.CornerRadius = [System.Windows.CornerRadius]::new(6)
        $Badge.Padding = [System.Windows.Thickness]::new(6, 3, 6, 3)
        $Badge.Margin = [System.Windows.Thickness]::new(0, 0, 4, 4)
        $Badge.BorderBrush = if ($Global:IsLightMode) { $BC.ConvertFromString("#C0E0C0") } else { $BC.ConvertFromString("#2D5A2D") }
        $Badge.BorderThickness = [System.Windows.Thickness]::new(1)

        $BadgeTxt = New-Object System.Windows.Controls.TextBlock
        $BadgeTxt.Text = "$($Def.Icon) $($Def.Name)"
        $BadgeTxt.FontSize = 9.5
        $BadgeTxt.Foreground = $Window.Resources['ThemeSuccess']
        $Badge.Child = $BadgeTxt
        $BadgeDate = try { ([DateTime]$Global:Achievements[$Def.Id]).ToString('MMM dd') } catch { '' }
        $Badge.ToolTip = "$($Def.Desc) — Unlocked $BadgeDate"

        $pnlAchievements.Children.Add($Badge) | Out-Null
    }

    # Locked badges (dimmed)
    foreach ($Def in $Locked) {
        $Badge = New-Object System.Windows.Controls.Border
        $Badge.Background = $Window.Resources['ThemeDeepBg']
        $Badge.CornerRadius = [System.Windows.CornerRadius]::new(6)
        $Badge.Padding = [System.Windows.Thickness]::new(6, 3, 6, 3)
        $Badge.Margin = [System.Windows.Thickness]::new(0, 0, 4, 4)
        $Badge.Opacity = 0.4

        $BadgeTxt = New-Object System.Windows.Controls.TextBlock
        $BadgeTxt.Text = "🔒 $($Def.Name)"
        $BadgeTxt.FontSize = 9.5
        $BadgeTxt.Foreground = $Window.Resources['ThemeTextDisabled']
        $Badge.Child = $BadgeTxt
        $Badge.ToolTip = $Def.Desc

        $pnlAchievements.Children.Add($Badge) | Out-Null
    }
}

# ==============================================================================
# SECTION 7AA: CUSTOMIZATION SCRIPT PROFILER (#R2-11)
# ==============================================================================

$Global:ScriptProfiler = [System.Collections.Generic.Dictionary[string, PSObject]]::new()
$Global:CurrentProfilerScript = $null
$Global:ProfilerScriptStartTime = $null

# Regex to extract explicit "Time taken: HH:MM:SS.fff" / "Time: HH:MM:SS" from script output
$Global:ProfilerTimeTakenRegex = [regex]::new(
    'Time\s*(?:taken)?:?\s*(\d{1,2}:\d{2}:\d{2}(?:\.\d+)?)',
    'Compiled, IgnoreCase'
)

$Global:ProfilerPatterns = @(
    # Primary: Match CUSTOMIZER PHASE "Executing <script>.ps1" announcements (all pipelines)
    @{ Pattern = 'CUSTOMIZER PHASE\s*:?\s*Executing\s+(\S+?)\.ps1'; Name = 'generic' }
    # Secondary: Script names embedded in script output
    @{ Pattern = '(Install-BasePkgs|Install-KrPkgs|Install-TradersPkgs|Install-LanguagePkgsOnline|InstallLanguagePacks|InstallDbBasePkgs)'; Name = '$1' }
    @{ Pattern = '(Remove-?AppxPackages|RemoveUserApps|Update-?ProvisionedAppxPackages)'; Name = '$1' }
    @{ Pattern = '(WindowsOptimization|DisableAutoUpdates|ResetAutoUpdateSettings)'; Name = '$1' }
    @{ Pattern = '(TimezoneRedirection|AdminSysPrep|UpdateWinGet|Customization)\.ps1'; Name = '$1' }
    @{ Pattern = 'Running command.*?([A-Za-z_-]+\.ps1)'; Name = 'generic' }
    @{ Pattern = 'CUSTOMIZER PHASE.*Running:\s*(.+?)\s*\.\.\.'; Name = 'generic' }
    @{ Pattern = 'Apply-Customizations|Customization\.ps1'; Name = 'Apply-Customizations' }
)

function Update-ScriptProfiler {
    <#
    .SYNOPSIS
        Detects customization script transitions and profiles durations.
        Uses explicit "Time taken" output from scripts for accurate duration
        (Packer blob logs lack timestamps, making wall-clock timing unreliable
        during backfill where all lines are processed in a few seconds).
    #>
    param([string]$Line)

    # ── 1. Check for "Time taken: HH:MM:SS.fff" / "Time: HH:MM:SS" output ──
    # Scripts output their own stopwatch elapsed time.  This is the most accurate
    # duration source — especially during blob backfill where Get-Date is useless.
    $TtMatch = $Global:ProfilerTimeTakenRegex.Match($Line)
    if ($TtMatch.Success -and $Global:CurrentProfilerScript) {
        if ($Global:ScriptProfiler.ContainsKey($Global:CurrentProfilerScript)) {
            try {
                $TS = [TimeSpan]::Parse($TtMatch.Groups[1].Value)
                $Current = $Global:ScriptProfiler[$Global:CurrentProfilerScript]
                $Current.DurationSec = $TS.TotalSeconds
                $Current.ExplicitDuration = $true
                if (-not $Current.EndAt) { $Current.EndAt = $Current.StartAt }
            } catch { <# Unparseable timespan, ignore #> }
        }
    }

    # ── 2. Detect script name from line ──
    $DetectedScript = $null
    foreach ($Pat in $Global:ProfilerPatterns) {
        if ($Line -match $Pat.Pattern) {
            if ($Pat.Name -eq 'generic') {
                $DetectedScript = $Matches[1]
            } elseif ($Pat.Name -eq '$1') {
                $DetectedScript = $Matches[1]
            } else {
                $DetectedScript = $Pat.Name
            }
            break
        }
    }

    if (-not $DetectedScript) { return }

    # Extract log timestamp, fall back to wall clock
    $Now = Get-Date
    $TsMatch = $Global:TimestampRegex.Match($Line)
    if ($TsMatch.Success) {
        $TsRaw = if ($TsMatch.Groups[1].Success) { $TsMatch.Groups[1].Value } else { $TsMatch.Groups[2].Value }
        try { $Now = [DateTime]::Parse($TsRaw, [System.Globalization.CultureInfo]::InvariantCulture) } catch { }
    }

    # Finish previous script (only if no explicit duration was already set)
    if ($Global:CurrentProfilerScript -and $Global:CurrentProfilerScript -ne $DetectedScript) {
        if ($Global:ScriptProfiler.ContainsKey($Global:CurrentProfilerScript)) {
            $Prev = $Global:ScriptProfiler[$Global:CurrentProfilerScript]
            if (-not $Prev.EndAt) {
                $Prev.EndAt = $Now
                if (-not $Prev.ExplicitDuration) {
                    $Prev.DurationSec = ($Now - $Prev.StartAt).TotalSeconds
                }
            }
        }
    }

    # Start new script (or new run of same script)
    if (-not $Global:ScriptProfiler.ContainsKey($DetectedScript)) {
        $Global:ScriptProfiler[$DetectedScript] = [PSCustomObject]@{
            Name             = $DetectedScript
            StartAt          = $Now
            EndAt            = $null
            DurationSec      = 0
            ExplicitDuration = $false
        }
    } elseif ($Global:ScriptProfiler[$DetectedScript].EndAt -or $Global:ScriptProfiler[$DetectedScript].ExplicitDuration) {
        # Script already finished (has EndAt or explicit duration) — this is a SECOND run.
        # Create a new entry with "#N" suffix to track each invocation separately.
        $RunNum = 2
        while ($Global:ScriptProfiler.ContainsKey("$DetectedScript #$RunNum")) { $RunNum++ }
        $SuffixedName = "$DetectedScript #$RunNum"
        $Global:ScriptProfiler[$SuffixedName] = [PSCustomObject]@{
            Name             = $SuffixedName
            StartAt          = $Now
            EndAt            = $null
            DurationSec      = 0
            ExplicitDuration = $false
        }
        $DetectedScript = $SuffixedName
    }

    $Global:CurrentProfilerScript = $DetectedScript
    $Global:ProfilerScriptStartTime = $Now
}

function Render-ScriptProfiler {
    <# Renders the script profiler panel with bars #>
    if (-not $pnlScriptProfiler) { return }
    $pnlScriptProfiler.Children.Clear()
    $BC = $Global:CachedBC

    if ($Global:ScriptProfiler.Count -eq 0) {
        $Empty = New-Object System.Windows.Controls.TextBlock
        $Empty.Text = "No scripts detected yet"
        $Empty.FontSize = 10
        $Empty.FontStyle = 'Italic'
        $Empty.Foreground = $Window.Resources['ThemeTextDim']
        $pnlScriptProfiler.Children.Add($Empty) | Out-Null
        return
    }

    # Close the currently running script for display
    $Scripts = $Global:ScriptProfiler.Values | ForEach-Object {
        $S = $_
        $Dur = if ($S.DurationSec -gt 0) { $S.DurationSec }
              elseif ($S.StartAt) { ((Get-Date) - $S.StartAt).TotalSeconds }
              else { 0 }
        [PSCustomObject]@{ Name = $S.Name; Duration = $Dur; Explicit = [bool]$S.ExplicitDuration }
    } | Sort-Object Duration -Descending

    $TotalDur = ($Scripts | Measure-Object Duration -Sum).Sum
    if ($TotalDur -le 0) { $TotalDur = 1 }
    $MaxDur = ($Scripts | Measure-Object Duration -Maximum).Maximum
    if ($MaxDur -le 0) { $MaxDur = 1 }

    foreach ($S in $Scripts) {
        $Row = New-Object System.Windows.Controls.Grid
        $Row.Margin = [System.Windows.Thickness]::new(0, 0, 0, 3)
        $C0 = New-Object System.Windows.Controls.ColumnDefinition; $C0.Width = [System.Windows.GridLength]::new(1, 'Star')
        $C1 = New-Object System.Windows.Controls.ColumnDefinition; $C1.Width = [System.Windows.GridLength]::Auto
        $Row.ColumnDefinitions.Add($C0)
        $Row.ColumnDefinitions.Add($C1)

        $NameTxt = New-Object System.Windows.Controls.TextBlock
        $NameTxt.Text = $S.Name
        $NameTxt.FontSize = 9.5
        $NameTxt.TextTrimming = 'CharacterEllipsis'
        $NameTxt.Foreground = $Window.Resources['ThemeTextBody']
        [System.Windows.Controls.Grid]::SetColumn($NameTxt, 0)

        $DurStr = if ($S.Duration -ge 60) { "$([Math]::Round($S.Duration/60, 1))m" }
                  elseif ($S.Duration -ge 1)  { "$([Math]::Round($S.Duration, 0))s" }
                  elseif ($S.Duration -gt 0)  { '<1s' }
                  else { '0s' }
        # Prefix "~" for wall-clock estimated durations (no explicit "Time taken" from script)
        if (-not $S.Explicit -and $S.Duration -gt 0) { $DurStr = "~$DurStr" }
        $PctStr = "$([Math]::Round(($S.Duration / $TotalDur) * 100, 0))%"
        $DurTxt = New-Object System.Windows.Controls.TextBlock
        $DurTxt.Text = "$DurStr ($PctStr)"
        $DurTxt.FontSize = 9
        $DurTxt.FontFamily = [System.Windows.Media.FontFamily]::new("Consolas")
        $DurTxt.Foreground = $Window.Resources['ThemeTextDim']
        [System.Windows.Controls.Grid]::SetColumn($DurTxt, 1)

        $Row.Children.Add($NameTxt) | Out-Null
        $Row.Children.Add($DurTxt) | Out-Null
        $pnlScriptProfiler.Children.Add($Row) | Out-Null

        # Progress bar
        $BarBg = New-Object System.Windows.Controls.Border
        $BarBg.Height = 3
        $BarBg.Background = $Window.Resources['ThemeDeepBg']
        $BarBg.CornerRadius = [System.Windows.CornerRadius]::new(1.5)
        $BarBg.Margin = [System.Windows.Thickness]::new(0, 1, 0, 2)

        $BarInner = New-Object System.Windows.Controls.Border
        $BarInner.Height = 3
        $BarInner.CornerRadius = [System.Windows.CornerRadius]::new(1.5)
        $BarInner.HorizontalAlignment = 'Left'
        $BarPct = [Math]::Min(100, ($S.Duration / $MaxDur) * 100)
        $BarInner.Width = [Math]::Max(2, ($BarPct / 100) * 180)  # 180px max

        # Color: longest is red, others are blue
        $BarColor = if ($S -eq $Scripts[0]) { & $Global:RemapLight '#D13438' } else { & $Global:RemapLight '#0078D4' }
        $BarInner.Background = $BC.ConvertFromString($BarColor)

        $BarBg.Child = $BarInner
        $pnlScriptProfiler.Children.Add($BarBg) | Out-Null
    }
}

function Reset-ScriptProfiler {
    $Global:ScriptProfiler.Clear()
    $Global:CurrentProfilerScript = $null
    $Global:ProfilerScriptStartTime = $null
    if ($pnlScriptProfiler) { $pnlScriptProfiler.Children.Clear() }
}

# ==============================================================================
# SECTION 7AB: PERSISTENT USER PREFERENCES (#R2-13)
# ==============================================================================

$Global:PrefsPath = Join-Path $Global:Root "user_prefs.json"

function Load-UserPrefs {
    <# Loads saved preferences and applies to UI #>
    if (-not (Test-Path $Global:PrefsPath)) { return }
    try {
        $P = Get-Content $Global:PrefsPath -Raw | ConvertFrom-Json

        if ($null -ne $P.WordWrap)       { $chkWordWrap.IsChecked    = $P.WordWrap }
        if ($null -ne $P.SoundEnabled)    { $chkSoundAlerts.IsChecked = $P.SoundEnabled; $Global:SoundEnabled = $P.SoundEnabled }
        if ($null -ne $P.AutoRefresh)     { $chkAutoRefresh.IsChecked = $P.AutoRefresh }
        if ($null -ne $P.AutoScanBuilds)  { $chkAutoScanBuilds.IsChecked = $P.AutoScanBuilds }
        if ($null -ne $P.PollInterval)    { $txtPollInterval.Text     = [string]$P.PollInterval }
        if ($null -ne $P.IsLightMode)     { $Global:SuppressThemeHandler = $true; $Global:IsLightMode = $P.IsLightMode; & $ApplyTheme $P.IsLightMode; $btnThemeToggle.Content = if ($P.IsLightMode) { [char]0x263E } else { [char]0x2600 }; if ($chkDarkMode) { $chkDarkMode.IsChecked = -not $P.IsLightMode }; $Global:SuppressThemeHandler = $false }
        if ($null -ne $P.DebugEnabled)    { $chkDebug.IsChecked = $P.DebugEnabled }

        # Analytics panel visibility
        if ($null -ne $P.AnalyticsPanelVisible -and $P.AnalyticsPanelVisible -eq $false) {
            Toggle-RightPanel $false
        }
        # Right panel section toggles
        if ($null -ne $P.RightToggles) {
            foreach ($Key in @('Stats','Profiler','History')) {
                if ($null -ne $P.RightToggles.$Key) { $Global:RightRailToggles[$Key] = [bool]$P.RightToggles.$Key }
            }
            Update-RightRailIndicators
            Apply-RightSections
        }

        # Left panel collapse state
        if ($null -ne $P.LeftPanelOpen -and $P.LeftPanelOpen -eq $false) {
            Toggle-LeftPanel $false
        }
        # Left panel section toggles
        if ($null -ne $P.LeftToggles) {
            foreach ($Key in @('Builds','Info','Settings')) {
                if ($null -ne $P.LeftToggles.$Key) { $Global:LeftRailToggles[$Key] = [bool]$P.LeftToggles.$Key }
            }
            Update-RailIndicators
            Apply-LeftSections
        }

        # Window position/size (E2: validate against virtual screen bounds)
        if ($null -ne $P.WindowState) {
            if ($P.WindowState -ne 'Maximized' -and $null -ne $P.WindowLeft) {
                $VW = [System.Windows.SystemParameters]::VirtualScreenWidth
                $VH = [System.Windows.SystemParameters]::VirtualScreenHeight
                $VL = [System.Windows.SystemParameters]::VirtualScreenLeft
                $VT = [System.Windows.SystemParameters]::VirtualScreenTop
                $PL = $P.WindowLeft; $PT = $P.WindowTop; $PW = $P.WindowWidth; $PH = $P.WindowHeight
                # Ensure at least 100px of the window is visible on any monitor
                if ($PL -gt ($VL + $VW - 100) -or ($PL + $PW) -lt ($VL + 100) -or
                    $PT -gt ($VT + $VH - 100) -or ($PT + $PH) -lt ($VT + 100)) {
                    Write-DebugLog "Saved window position off-screen, centering"
                    $Window.WindowStartupLocation = 'CenterScreen'
                } else {
                    $Window.WindowState = 'Normal'
                    $Window.Left   = $PL
                    $Window.Top    = $PT
                    $Window.Width  = $PW
                    $Window.Height = $PH
                }
            }
        }

        Write-DebugLog "User preferences loaded"
    } catch {
        Write-DebugLog "Failed to load preferences: $($_.Exception.Message)"
    }
}

function Save-UserPrefs {
    <# Persists current UI state to JSON #>
    try {
        $P = [PSCustomObject]@{
            WordWrap        = [bool]$chkWordWrap.IsChecked
            SoundEnabled    = [bool]$chkSoundAlerts.IsChecked
            AutoRefresh     = [bool]$chkAutoRefresh.IsChecked
            AutoScanBuilds  = [bool]$chkAutoScanBuilds.IsChecked
            PollInterval    = $txtPollInterval.Text
            IsLightMode     = $Global:IsLightMode
            DebugEnabled    = [bool]$chkDebug.IsChecked
            AnalyticsPanelVisible = $Global:RightPanelOpen
            LeftPanelOpen   = $Global:LeftPanelOpen
            LeftToggles     = $Global:LeftRailToggles
            RightToggles    = $Global:RightRailToggles
            WindowState     = $Window.WindowState.ToString()
            WindowLeft      = $Window.Left
            WindowTop       = $Window.Top
            WindowWidth     = $Window.Width
            WindowHeight    = $Window.Height
        }
        $P | ConvertTo-Json | Set-Content $Global:PrefsPath -Force
        Write-DebugLog "User preferences saved"
    } catch {
        Write-DebugLog "Failed to save preferences: $($_.Exception.Message)"
    }
}

# ==============================================================================
# SECTION 7AC: ERROR HEATMAP TIMELINE (#R2-14)
# ==============================================================================

$Global:HeatmapData = [System.Collections.Generic.List[PSObject]]::new()
$Global:HeatmapLastRenderedIndex = 0   # incremental render marker

function Add-HeatmapPoint {
    <# Adds a data point to the timeline heatmap #>
    param([string]$Color, [int]$LineIndex)
    $Severity = switch ($Color) {
        '#D13438' { 'error' }
        '#FFB900' { 'warning' }
        '#107C10' { 'success' }
        default   { $null }
    }
    if ($Severity) {
        $Global:HeatmapData.Add([PSCustomObject]@{
            Severity  = $Severity
            LineIndex = $LineIndex
            TimeStamp = Get-Date
        })
    }
}

function Render-Heatmap {
    <# Incrementally adds new dots to the heatmap strip. Viewport indicator refreshed each call. #>
    if (-not $cnvHeatmap) { return }
    if ($Global:HeatmapData.Count -eq 0) { return }

    $W = $cnvHeatmap.ActualWidth
    if ($W -le 0) { $W = 200 }
    $H = $cnvHeatmap.ActualHeight
    if ($H -le 0) { $H = 16 }

    $MaxLine = $Global:LogStats.TotalLines
    if ($MaxLine -le 0) { $MaxLine = 1 }

    $BC = $Global:CachedBC

    # Full rebuild if canvas was cleared (reset)
    $FullRebuild = ($cnvHeatmap.Children.Count -eq 0 -and $Global:HeatmapData.Count -gt 0)
    $StartIdx = if ($FullRebuild) { 0 } else { $Global:HeatmapLastRenderedIndex }

    # Draw only new colored dots
    for ($hi = $StartIdx; $hi -lt $Global:HeatmapData.Count; $hi++) {
        $Point = $Global:HeatmapData[$hi]
        $X = ($Point.LineIndex / $MaxLine) * $W
        $DotColor = switch ($Point.Severity) {
            'error'   { & $Global:RemapLight '#D13438' }
            'warning' { & $Global:RemapLight '#FFB900' }
            'success' { & $Global:RemapLight '#107C10' }
        }
        $Dot = New-Object System.Windows.Shapes.Ellipse
        $Dot.Width = 3
        $Dot.Height = $H
        $Dot.Fill = $BC.ConvertFromString($DotColor)
        $Dot.Opacity = 0.7
        [System.Windows.Controls.Canvas]::SetLeft($Dot, [Math]::Min($X, $W - 3))
        [System.Windows.Controls.Canvas]::SetTop($Dot, 0)
        $cnvHeatmap.Children.Add($Dot) | Out-Null
    }
    $Global:HeatmapLastRenderedIndex = $Global:HeatmapData.Count

    # Remove old viewport indicator and redraw
    $ViewTag = 'heatmap-viewport'
    for ($vi = $cnvHeatmap.Children.Count - 1; $vi -ge 0; $vi--) {
        if ($cnvHeatmap.Children[$vi].Tag -eq $ViewTag) {
            $cnvHeatmap.Children.RemoveAt($vi)
        }
    }
    if ($Global:LogStats.TotalLines -gt 0 -and $rtbOutput.ExtentHeight -gt 0) {
        $ViewStart = ($rtbOutput.VerticalOffset / $rtbOutput.ExtentHeight) * $W
        $ViewWidth = [Math]::Max(8, ($rtbOutput.ViewportHeight / $rtbOutput.ExtentHeight) * $W)
        $ViewRect = New-Object System.Windows.Shapes.Rectangle
        $ViewRect.Width = $ViewWidth
        $ViewRect.Height = $H
        $ViewRect.Fill = $Window.Resources['ThemeScrollThumb']
        $ViewRect.Opacity = 0.5
        $ViewRect.Stroke = $Window.Resources['ThemeAccentLight']
        $ViewRect.StrokeThickness = 0.5
        $ViewRect.Tag = $ViewTag
        [System.Windows.Controls.Canvas]::SetLeft($ViewRect, $ViewStart)
        [System.Windows.Controls.Canvas]::SetTop($ViewRect, 0)
        $cnvHeatmap.Children.Add($ViewRect) | Out-Null
    }
}

function Reset-Heatmap {
    $Global:HeatmapData.Clear()
    $Global:HeatmapLastRenderedIndex = 0
    if ($cnvHeatmap) { $cnvHeatmap.Children.Clear() }
}

# ==============================================================================
# SECTION 7AD: BUILD COST ESTIMATOR (#R2-15)
# ==============================================================================

# AIB default VM SKU pricing (approximate USD/hr for common SKUs)
$Global:VMCostTable = @{
    'Standard_D2s_v3'  = 0.096
    'Standard_D4s_v3'  = 0.192
    'Standard_D8s_v3'  = 0.384
    'Standard_D2ds_v5' = 0.115
    'Standard_D4ds_v5' = 0.230
    'Standard_D2s_v5'  = 0.096
    'Standard_D4s_v5'  = 0.192
    'Standard_E2s_v3'  = 0.126
    'Standard_D2as_v5' = 0.086
    '_default'         = 0.192  # Default to D4s_v3 pricing
}

$Global:DetectedVMSku = $null
$Global:EstimatedCostPerHr = $Global:VMCostTable['_default']

function Detect-VMSku {
    <# Attempts to detect VM SKU from packer log output #>
    param([string]$Line)
    if ($Global:DetectedVMSku) { return }  # Already detected
    if ($Line -match '(Standard_[A-Za-z0-9_]+)') {
        $Sku = $Matches[1]
        if ($Global:VMCostTable.ContainsKey($Sku)) {
            $Global:DetectedVMSku = $Sku
            $Global:EstimatedCostPerHr = $Global:VMCostTable[$Sku]
        } else {
            $Global:DetectedVMSku = $Sku
            # Keep default rate if not in table
        }
        Write-DebugLog "Cost estimator: detected VM SKU = $Sku ($($Global:EstimatedCostPerHr)/hr)"
    }
}

function Update-CostEstimate {
    <# Updates the running cost ticker #>
    param([DateTime]$Now = (Get-Date))
    if (-not $Global:SyncHash.BuildStartTime) { return }
    if (-not $lblCostEstimate) { return }

    $ElapsedHrs = ($Now.ToUniversalTime() - $Global:SyncHash.BuildStartTime).TotalHours
    $Cost = $ElapsedHrs * $Global:EstimatedCostPerHr
    $PerMin = $Global:EstimatedCostPerHr / 60

    $CostStr = '$' + [Math]::Round($Cost, 2).ToString('0.00')
    $RateStr = '$' + [Math]::Round($PerMin, 4).ToString('0.0000') + '/min'

    $lblCostEstimate.Text = "$CostStr"
    if ($lblCostRate) { $lblCostRate.Text = $RateStr }

    $SkuStr = if ($Global:DetectedVMSku) { $Global:DetectedVMSku } else { "Default (D4s v3)" }
    if ($lblCostSku) { $lblCostSku.Text = $SkuStr }
}

# ==============================================================================
# SECTION 7AE: WINGET DOWNLOAD PROGRESS (#R2-17)
# ==============================================================================

# WinGet progress is integrated into the Package Tracker (Section 7V)
# The PkgPatterns.WinGetDownload regex captures percentage from download lines.
# The Render-PackageTracker function shows live % for downloading packages.
# No separate section needed — this enhancement is built into 7V.

# ==============================================================================
# SECTION 7AF: DARK/LIGHT THEME AUTO-DETECT (#R2-20)
# ==============================================================================

function Get-SystemTheme {
    <# Detects Windows system theme (returns $true for light mode) #>
    try {
        $RegPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize'
        $AppsUseLight = Get-ItemPropertyValue $RegPath 'AppsUseLightTheme' -ErrorAction SilentlyContinue
        return ($AppsUseLight -eq 1)
    } catch {
        return $false  # Default to dark
    }
}

function Apply-SystemTheme {
    <# Applies theme based on Windows system setting #>
    $IsLight = Get-SystemTheme
    $Global:IsLightMode = $IsLight
    & $ApplyTheme $IsLight
    if ($btnThemeToggle) {
        $btnThemeToggle.Content = if ($IsLight) { [char]0x263E } else { [char]0x2600 }
    }
    Write-DebugLog "Auto-detected system theme: $(if ($IsLight) { 'Light' } else { 'Dark' })"
}

# ==============================================================================
# SECTION 8: PREREQUISITE CHECK UI
# ==============================================================================

function Show-PrereqStatus {
    <#
    .SYNOPSIS
        Populates the prerequisite error panel with module status and remediation command.
    #>
    param([array]$Results)

    $pnlModuleStatus.Children.Clear()
    $AllGood = $true

    foreach ($Mod in $Results) {
        $Row = New-Object System.Windows.Controls.Grid
        $Col0 = New-Object System.Windows.Controls.ColumnDefinition; $Col0.Width = "Auto"
        $Col1 = New-Object System.Windows.Controls.ColumnDefinition; $Col1.Width = "*"
        $Col2 = New-Object System.Windows.Controls.ColumnDefinition; $Col2.Width = "Auto"
        $Row.ColumnDefinitions.Add($Col0)
        $Row.ColumnDefinitions.Add($Col1)
        $Row.ColumnDefinitions.Add($Col2)
        $Row.Margin = [System.Windows.Thickness]::new(0, 3, 0, 3)

        # Status icon
        $Icon = New-Object System.Windows.Controls.TextBlock
        $Icon.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
        $Icon.FontSize = 14
        $Icon.VerticalAlignment = 'Center'
        $Icon.Margin = [System.Windows.Thickness]::new(0, 0, 10, 0)
        if ($Mod.Available) {
            $Icon.Text = [char]0xE73E   # Checkmark
            $Icon.Foreground = $Window.Resources['ThemeSuccess']
        } else {
            $Icon.Text = [char]0xE711   # X mark
            $Icon.Foreground = $Window.Resources['ThemeError']
            $AllGood = $false
        }
        [System.Windows.Controls.Grid]::SetColumn($Icon, 0)
        $Row.Children.Add($Icon) | Out-Null

        # Module name + purpose
        $NameBlock = New-Object System.Windows.Controls.StackPanel
        $ModName = New-Object System.Windows.Controls.TextBlock
        $ModName.Text = $Mod.Name
        $ModName.FontSize = 13
        $ModName.FontWeight = 'SemiBold'
        $ModName.Foreground = $Window.Resources['ThemeTextPrimary']
        $NameBlock.Children.Add($ModName) | Out-Null
        $ModPurpose = New-Object System.Windows.Controls.TextBlock
        $ModPurpose.Text = $Mod.Purpose
        $ModPurpose.FontSize = 11
        $ModPurpose.Foreground = $Window.Resources['ThemeTextDim']
        $NameBlock.Children.Add($ModPurpose) | Out-Null
        [System.Windows.Controls.Grid]::SetColumn($NameBlock, 1)
        $Row.Children.Add($NameBlock) | Out-Null

        # Version badge
        $VerBadge = New-Object System.Windows.Controls.TextBlock
        $VerBadge.FontSize = 11
        $VerBadge.VerticalAlignment = 'Center'
        if ($Mod.Available) {
            $VerBadge.Text = "v$($Mod.Installed)"
            $VerBadge.Foreground = $Window.Resources['ThemeTextMuted']
        } else {
            $VerBadge.Text = "NOT FOUND"
            $VerBadge.Foreground = $Window.Resources['ThemeError']
            $VerBadge.FontWeight = 'Bold'
        }
        [System.Windows.Controls.Grid]::SetColumn($VerBadge, 2)
        $Row.Children.Add($VerBadge) | Out-Null

        $pnlModuleStatus.Children.Add($Row) | Out-Null
    }

    # Build remediation command
    $Command = Get-RemediationCommand -ModuleResults $Results
    if ($Command) {
        $txtPrereqCommand.Text = $Command
        $pnlPrereqError.Visibility = 'Visible'
        $btnLogin.IsEnabled = $false
        $btnRefreshBuilds.IsEnabled = $false
    } else {
        $pnlPrereqError.Visibility = 'Collapsed'
        $btnLogin.IsEnabled = $true
        $btnRefreshBuilds.IsEnabled = $true
    }
    return $AllGood
}

# ==============================================================================
# SECTION 9: AZURE AUTHENTICATION
# ==============================================================================

# NOTE: Token validation is now inlined inside each Start-BackgroundWork scriptblock
#       (Get-AzResourceGroup | Select -First 1) to avoid blocking the UI thread.

function Set-AuthUIConnected {
    <#
    .SYNOPSIS
        Updates all auth-related UI elements to the "connected" state.
        Accepts pre-fetched subscription list (fetched in background to avoid UI hangs).
    #>
    param(
        [string]$AccountId,
        [string]$SubName,
        [string]$SubId,
        [array]$Subscriptions = @()
    )
    Write-DebugLog "Set-AuthUIConnected: acct=$AccountId sub=$SubName subs=$($Subscriptions.Count)"
    $Global:IsAuthenticated = $true
    $BC = $Global:CachedBC
    $authDot.Fill = $Window.Resources['ThemeSuccess']
    $lblAuthStatus.Text = $AccountId
    $lblAuthStatus.Foreground = $Window.Resources['ThemeTextPrimary']

    # Populate subscription dropdown using ComboBoxItem objects (reliable across runspace boundaries)
    $cmbSubscription.Items.Clear()
    if ($Subscriptions.Count -gt 0) {
        for ($i = 0; $i -lt $Subscriptions.Count; $i++) {
            $Item = New-Object System.Windows.Controls.ComboBoxItem
            $Item.Content = [string]$Subscriptions[$i].Name
            $Item.Tag     = [string]$Subscriptions[$i].Id
            $cmbSubscription.Items.Add($Item) | Out-Null
            if ($Subscriptions[$i].Id -eq $SubId) {
                $cmbSubscription.SelectedIndex = $i
            }
        }
    } else {
        # Fallback: show current sub only
        $Item = New-Object System.Windows.Controls.ComboBoxItem
        $Item.Content = [string]$SubName
        $Item.Tag     = [string]$SubId
        $cmbSubscription.Items.Add($Item) | Out-Null
        $cmbSubscription.SelectedIndex = 0
    }

    $cmbSubscription.Visibility = 'Visible'
    $btnLogin.Visibility = 'Collapsed'
    $btnLogout.Visibility = 'Visible'
    $lblStatus.Text = "Connected as $AccountId"

    # Update ticker connection badge
    if ($lblTickerConnection) {
        $lblTickerConnection.Text = 'ONLINE'
        $lblTickerConnection.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#00C853')
    }
    if ($badgeConnection) {
        $badgeConnection.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#1500C853')
    }
    if ($lblTickerSub -and $SubName) {
        $lblTickerSub.Text = $SubName.ToUpper()
    }
    Write-DebugLog "Set-AuthUIConnected: done"
}

function Set-AuthUIDisconnected {
    <#
    .SYNOPSIS
        Updates all auth-related UI elements to the "disconnected" state.
    #>
    param([string]$Reason = "Not connected")
    Write-DebugLog "Set-AuthUIDisconnected: reason=$Reason"
    $Global:IsAuthenticated = $false
    $BC = $Global:CachedBC
    $authDot.Fill = $Window.Resources['ThemeError']
    $lblAuthStatus.Text = $Reason
    $lblAuthStatus.Foreground = $Window.Resources['ThemeTextMuted']
    $cmbSubscription.Items.Clear()
    $cmbSubscription.Visibility = 'Collapsed'
    $btnLogin.Visibility = 'Visible'
    $btnLogout.Visibility = 'Collapsed'

    # Update ticker connection badge
    if ($lblTickerConnection) {
        $lblTickerConnection.Text = 'OFFLINE'
        $lblTickerConnection.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#D13438')
    }
    if ($badgeConnection) {
        $badgeConnection.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#15D13438')
    }
    if ($lblTickerSub)  { $lblTickerSub.Text = '' }
    if ($lblTickerRegion) { $lblTickerRegion.Text = '' }
    Update-TickerStrip -Builds @()
}

function Connect-ToAzure {
    <#
    .SYNOPSIS
        Attempts interactive Azure login via WAM (Windows Account Manager).
    .DESCRIPTION
        Connect-AzAccount with WAM shows a native Windows dialog that requires the
        UI thread's STA message pump — it cannot run in a background runspace.
        The login itself blocks briefly (user picks an account), then post-login
        work (subscription enumeration + build scan) runs in the background.
    #>
    $lblAuthStatus.Text = "Signing in..."
    $lblStatus.Text = "Authenticating with Entra ID..."
    $Window.Cursor = [System.Windows.Input.Cursors]::Wait
    $btnLogin.IsEnabled = $false
    Write-DebugLog "Connect-ToAzure: >>> ENTER"

    try {
        Write-DebugLog "Connect-ToAzure: calling Get-AzContext..."
        $Context = Get-AzContext -ErrorAction SilentlyContinue
        Write-DebugLog "Connect-ToAzure: Get-AzContext returned (Account=$($Context.Account.Id), Sub=$($Context.Subscription.Name))"

        # If a cached context exists, validate the token with a lightweight call
        $NeedLogin = $true
        if ($Context -and $Context.Account) {
            $lblAuthStatus.Text = "Validating token for $($Context.Account.Id)..."
            Write-DebugLog "Connect-ToAzure: cached context exists for $($Context.Account.Id), validating..."
            try {
                # Lightweight token validation — no network call if token is fresh
                $null = Get-AzSubscription -SubscriptionId $Context.Subscription.Id -ErrorAction Stop -WarningAction SilentlyContinue
                Write-DebugLog "Connect-ToAzure: cached token is valid, skipping WAM re-auth"
                $NeedLogin = $false
            } catch {
                Write-DebugLog "Connect-ToAzure: cached token expired/invalid: $($_.Exception.Message)"
            }
        }

        if ($NeedLogin) {
            # Interactive browser auth — disable WAM, supports MFA/CA
            Write-DebugLog "Connect-ToAzure: disabling WAM, using interactive browser..."
            $lblAuthStatus.Text = "Opening browser for sign-in..."
            $lblStatus.Text = "Browser sign-in (check your default browser)..."
            try { Update-AzConfig -EnableLoginByWam $false -ErrorAction SilentlyContinue | Out-Null } catch { }
            $ConnectParams = @{ WarningAction = 'SilentlyContinue'; ErrorAction = 'Stop' }
            if ($Context -and $Context.Tenant.Id) {
                $ConnectParams['TenantId'] = $Context.Tenant.Id
                Write-DebugLog "Connect-ToAzure: using cached TenantId=$($Context.Tenant.Id)"
            }
            Connect-AzAccount @ConnectParams | Out-Null
            Write-DebugLog "Connect-ToAzure: <<< Connect-AzAccount (browser) returned"
        }

        Write-DebugLog "Connect-ToAzure: calling Get-AzContext post-login..."
        $Context = Get-AzContext
        Write-DebugLog "Connect-ToAzure: post-login context: Account=$($Context.Account.Id) Sub=$($Context.Subscription.Name)"

        if ($Context -and $Context.Account) {
            $lblAuthStatus.Text = $Context.Account.Id
            $lblStatus.Text = "Signed in — loading subscriptions..."
            Write-DebugLog "Connect-ToAzure: auth OK, dispatching sub enum to background"

            # Post-login heavy work (subscription enumeration) → background
            $PostAcctId  = $Context.Account.Id
            $PostSubName = $Context.Subscription.Name
            $PostSubId   = $Context.Subscription.Id

            Write-DebugLog "Connect-ToAzure: calling Start-BackgroundWork for sub enumeration..."
            Start-BackgroundWork -Variables @{
                AcctId  = $PostAcctId
                SubName = $PostSubName
                SubId   = $PostSubId
            } -Work {
                $R = @{ AccountId = $AcctId; SubName = $SubName; SubId = $SubId; Subs = @() }
                try {
                    $RawSubs = @(Get-AzSubscription -WarningAction SilentlyContinue -ErrorAction Stop |
                                Where-Object { $_.State -eq 'Enabled' })
                    $R.Subs = @($RawSubs | ForEach-Object {
                        [PSCustomObject]@{ Name = $_.Name; Id = $_.Id }
                    })
                } catch { }
                return $R
            } -OnComplete {
                param($Results, $Errors)
                Write-DebugLog "Connect-ToAzure OnComplete: >>> ENTER"
                $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ AccountId = ''; SubName = ''; SubId = ''; Subs = @() } }
                Write-DebugLog "Connect-ToAzure OnComplete: $($R.Subs.Count) subscription(s) found"
                Write-DebugLog "Connect-ToAzure OnComplete: calling Set-AuthUIConnected..."
                Set-AuthUIConnected -AccountId $R.AccountId -SubName $R.SubName -SubId $R.SubId -Subscriptions $R.Subs
                Write-DebugLog "Connect-ToAzure OnComplete: calling Get-ActiveBuilds..."
                Get-ActiveBuilds
                Write-DebugLog "Connect-ToAzure OnComplete: restoring cursor"
                $Window.Cursor = [System.Windows.Input.Cursors]::Arrow
                Write-DebugLog "Connect-ToAzure OnComplete: <<< EXIT"
            }
            Write-DebugLog "Connect-ToAzure: Start-BackgroundWork dispatched, returning to event loop"
        } else {
            $lblAuthStatus.Text = "Sign-in cancelled"
            $lblStatus.Text = "Sign-in was cancelled"
            Write-DebugLog "Connect-ToAzure: no context after login (cancelled?)"
            $Window.Cursor = [System.Windows.Input.Cursors]::Arrow
        }
    } catch {
        $BC = $Global:CachedBC
        $authDot.Fill = $Window.Resources['ThemeError']
        $lblAuthStatus.Text = "Sign-in failed"
        $lblAuthStatus.Foreground = $Window.Resources['ThemeError']
        $lblStatus.Text = "Auth error: $($_.Exception.Message)"
        Write-DebugLog "Connect-ToAzure EXCEPTION: $($_.Exception.Message)"
        Write-DebugLog "Connect-ToAzure EXCEPTION type: $($_.Exception.GetType().FullName)"
        $Window.Cursor = [System.Windows.Input.Cursors]::Arrow
    } finally {
        $btnLogin.IsEnabled = $true
        Write-DebugLog "Connect-ToAzure: <<< EXIT (finally)"
    }
}

function Disconnect-FromAzure {
    <#
    .SYNOPSIS
        Logs out of Azure and resets UI state.
    #>
    try {
        Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null
    } catch { }

    $BC = $Global:CachedBC
    $authDot.Fill = $Window.Resources['ThemeError']
    $lblAuthStatus.Text = "Not connected"
    $lblAuthStatus.Foreground = $Window.Resources['ThemeTextMuted']
    $cmbSubscription.Items.Clear()
    $cmbSubscription.Visibility = 'Collapsed'
    $btnLogin.Visibility = 'Visible'
    $btnLogout.Visibility = 'Collapsed'
    $lstBuilds.ItemsSource = $null
    $lblBuildCount.Text = "0 builds found"
    $pnlEmptyState.Visibility = 'Visible'
    $lblEmptyTitle.Text = "No active builds"
    $lblEmptySubtitle.Text = "Sign in to discover running AIB builds"
    $lblStatus.Text = "Disconnected"

    # Stop any active streaming
    $Global:SyncHash.StopFlag = $true
    $Global:SyncHash.IsStreaming = $false
    $Global:SyncHash.SelectedBuild = $null
}

# ==============================================================================
# SECTION 10: BUILD DISCOVERY
# ==============================================================================

function Get-ActiveBuilds {
    <#
    .SYNOPSIS
        Discovers all currently running AIB builds in a background runspace.
    .DESCRIPTION
        Runs all Azure API calls (Get-AzResourceGroup, Get-AzContainerGroup) off the
        UI thread so the window never freezes. Results are dispatched back to the UI
        via the DispatcherTimer polling system.
    #>
    $lblStatus.Text = "Scanning for active builds..."
    $lblBuildCount.Text = "Scanning..."
    $buildScanDot.Opacity = 1
    $btnRefreshBuilds.IsEnabled = $false
    Write-DebugLog "Get-ActiveBuilds: >>> ENTER, dispatching scan to background"

    Start-BackgroundWork -Work {
        $R = @{ Builds = @(); StagingRGCount = 0; Error = ''; AuthError = $false }
        try {
            # Validate token with a subscription-level call
            try {
                Get-AzResourceGroup -ErrorAction Stop | Select-Object -First 1 | Out-Null
            } catch {
                $R.Error = $_.Exception.Message
                if ($R.Error -match 'credentials|expired|Connect-AzAccount|unauthorized|authentication') {
                    $R.AuthError = $true
                }
                return $R
            }

            # Find AIB staging resource groups
            $StagingRGs = Get-AzResourceGroup -ErrorAction Stop | Where-Object {
                ($_.Tags -and $_.Tags['createdBy'] -eq 'AzureVMImageBuilder') -or
                ($_.ResourceGroupName -match '^IT_.*imagebuilder.*_[0-9a-f]{8}-')
            }
            $R.StagingRGCount = if ($StagingRGs) { @($StagingRGs).Count } else { 0 }

            if (-not $StagingRGs -or $R.StagingRGCount -eq 0) { return $R }

            $Builds = [System.Collections.Generic.List[PSObject]]::new()

            foreach ($RG in $StagingRGs) {
                try {
                    $ContainerGroups = Get-AzContainerGroup -ResourceGroupName $RG.ResourceGroupName -ErrorAction SilentlyContinue

                    foreach ($CG in $ContainerGroups) {
                        if ($CG.Name -notlike 'vmimagebuilder-build-container-*') { continue }

                        $TemplateName = if ($RG.Tags -and $RG.Tags['imageTemplateName']) {
                            $RG.Tags['imageTemplateName']
                        } else {
                            $Parts = $RG.ResourceGroupName -split '_'
                            if ($Parts.Count -ge 3) { $Parts[2..($Parts.Count-2)] -join '_' } else { $CG.Name }
                        }

                        $StartTime = $null
                        $Elapsed = "N/A"
                        if ($CG.Containers -and $CG.Containers.Count -gt 0) {
                            try {
                                $ContainerDetail = $CG.Containers[0]
                                if ($ContainerDetail.InstanceViewCurrentState) {
                                    $StartTime = $ContainerDetail.InstanceViewCurrentState.StartTime
                                }
                            } catch { }
                        }
                        if (-not $StartTime -and $RG.Tags -and $RG.Tags['imageTemplateName']) {
                            if ($TemplateName -match 't_(\d{13})') {
                                try {
                                    $UnixMs = [long]$Matches[1]
                                    $StartTime = [DateTimeOffset]::FromUnixTimeMilliseconds($UnixMs).UtcDateTime
                                } catch { }
                            }
                        }
                        if ($StartTime) {
                            $Span = (Get-Date).ToUniversalTime() - $StartTime
                            if ($Span.TotalHours -ge 1) {
                                $Elapsed = "{0}h {1}m" -f [int]$Span.TotalHours, $Span.Minutes
                            } else {
                                $Elapsed = "{0}m" -f [int]$Span.TotalMinutes
                            }
                        }

                        $State = "Unknown"
                        try {
                            $ProvState = $CG.ProvisioningState
                            $ContainerState = if ($CG.Containers -and $CG.Containers[0].InstanceViewCurrentState) {
                                $CG.Containers[0].InstanceViewCurrentState.State
                            } else { $null }

                            if ($ContainerState) {
                                $State = $ContainerState
                            } elseif ($ProvState -eq 'Succeeded') {
                                # ProvisioningState "Succeeded" means the container was
                                # created successfully — NOT that the build finished.
                                # If the container group still exists, the build is running.
                                $State = "Running"
                            } elseif ($ProvState) {
                                $State = $ProvState
                            }
                        } catch { }

                        $StateColor = switch ($State) {
                            'Running'    { '#0078D4' }  # Microsoft Blue
                            'Waiting'    { '#8764B8' }  # Purple
                            'Terminated' { '#D13438' }  # Red
                            'Failed'     { '#D13438' }
                            default      { '#797775' }  # Gray
                        }

                        $Build = [PSCustomObject]@{
                            TemplateName   = $TemplateName
                            ResourceGroup  = $RG.ResourceGroupName
                            ContainerGroup = $CG.Name
                            ContainerName  = if ($CG.Containers) { $CG.Containers[0].Name } else { "packerizercontainer" }
                            Location       = $RG.Location
                            State          = $State
                            StateColor     = $StateColor
                            Elapsed        = $Elapsed
                            StartTime      = $StartTime
                        }
                        $Builds.Add($Build)
                    }
                } catch { continue }
            }
            $R.Builds = @($Builds)
        } catch {
            $R.Error = $_.Exception.Message
        }
        return $R
    } -OnComplete {
        param($Results, $Errors)
        Write-DebugLog "Get-ActiveBuilds OnComplete: >>> ENTER"
        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ Builds = @(); Error = 'No result'; AuthError = $false } }
        Write-DebugLog "Get-ActiveBuilds OnComplete: builds=$($R.Builds.Count) stagingRGs=$($R.StagingRGCount) error='$($R.Error)' authErr=$($R.AuthError)"

        if ($R.AuthError) {
            Set-AuthUIDisconnected -Reason "Credentials expired"
            $lblBuildCount.Text = "Auth expired"
            $pnlEmptyState.Visibility = 'Visible'
            $lblEmptyTitle.Text = "Session expired"
            $lblEmptySubtitle.Text = "Click Sign In to re-authenticate"
            $lblStatus.Text = "Token expired — please sign in again"
            Write-DebugLog "Get-ActiveBuilds: auth error — $($R.Error)"
        } elseif ($R.Error) {
            $lblBuildCount.Text = "Error scanning"
            $lblStatus.Text = "Error: $($R.Error)"
            $pnlEmptyState.Visibility = 'Visible'
            $lblEmptyTitle.Text = "Scan failed"
            $lblEmptySubtitle.Text = $R.Error
            Write-DebugLog "Get-ActiveBuilds ERROR: $($R.Error)"
        } elseif ($R.Builds.Count -gt 0) {
            $lstBuilds.ItemsSource = $R.Builds
            $lblBuildCount.Text = "$($R.Builds.Count) build$(if($R.Builds.Count -ne 1){'s'}) found"
            $pnlEmptyState.Visibility = 'Collapsed'
            $lblStatus.Text = "Found $($R.Builds.Count) active build$(if($R.Builds.Count -ne 1){'s'})"
            Write-DebugLog "Builds discovered: $($R.Builds.Count)"
        } else {
            $lstBuilds.ItemsSource = $null
            $lblBuildCount.Text = "0 builds found"
            $pnlEmptyState.Visibility = 'Visible'

            # Reset all sidebar panels when no builds are selected
            $Global:ParagraphCount = 0
            Reset-PhaseTracker
            Reset-LogStats
            Reset-FlameChart
            Reset-Minimap
            Reset-PackageTracker
            Reset-ErrorClusters
            Reset-KnownIssues
            Reset-ScriptProfiler
            Reset-Heatmap
            $Global:DetectedVMSku = $null
            $Global:ETAPercent = 0
            if ($lblHealthGrade)  { $lblHealthGrade.Text = '--' }
            if ($lblHealthScore)  { $lblHealthScore.Text = '' }
            if ($lblETAPercent)   { $lblETAPercent.Text = '' }
            if ($lblETARemaining) { $lblETARemaining.Text = '' }
            if ($prgETA)          { $prgETA.Value = 0 }
            if ($lblCostEstimate) { $lblCostEstimate.Text = '$0.00' }
            if ($lblCostRate)     { $lblCostRate.Text = '' }
            if ($lblCostSku)      { $lblCostSku.Text = '' }

            if ($R.StagingRGCount -gt 0) {
                $lblEmptyTitle.Text = "No running containers"
                $lblEmptySubtitle.Text = "Found $($R.StagingRGCount) staging RG(s) but no active build containers"
            } else {
                $lblEmptyTitle.Text = "No active builds"
                $lblEmptySubtitle.Text = "No AIB staging resource groups found in current subscription"
            }
            $lblStatus.Text = "Scan complete — no active builds"
        }
        # Update ticker strip with build data
        Update-TickerStrip -Builds @($R.Builds)
        $buildScanDot.Opacity = 0
        $btnRefreshBuilds.IsEnabled = $true
        Write-DebugLog "Get-ActiveBuilds OnComplete: <<< EXIT"
    }
}

function Update-TickerStrip {
    <#
    .SYNOPSIS
        Updates the Build Status Ticker Strip counts and badge visibility.
    .DESCRIPTION
        Takes a list of build objects and counts Running/Failed/Succeeded states,
        then updates the ticker strip badges accordingly.
    #>
    param([array]$Builds = @())

    $TotalCount = if ($Builds) { $Builds.Count } else { 0 }
    $Running    = @($Builds | Where-Object { $_.State -eq 'Running' -or $_.State -eq 'Waiting' }).Count
    $Failed     = @($Builds | Where-Object { $_.State -eq 'Failed' -or $_.State -eq 'Terminated' }).Count
    $Succeeded  = @($Builds | Where-Object { $_.State -eq 'Succeeded' }).Count

    if ($lblTickerBuilds) { $lblTickerBuilds.Text = $TotalCount.ToString() }

    # Running badge
    if ($badgeRunning) {
        if ($Running -gt 0) {
            $badgeRunning.Visibility = 'Visible'
            $lblTickerRunning.Text = "$Running Running"
        } else {
            $badgeRunning.Visibility = 'Collapsed'
        }
    }

    # Failed badge
    if ($badgeFailed) {
        if ($Failed -gt 0) {
            $badgeFailed.Visibility = 'Visible'
            $lblTickerFailed.Text = "$Failed Failed"
        } else {
            $badgeFailed.Visibility = 'Collapsed'
        }
    }

    # Succeeded badge
    if ($badgeSucceeded) {
        if ($Succeeded -gt 0) {
            $badgeSucceeded.Visibility = 'Visible'
            $lblTickerSucceeded.Text = "$Succeeded OK"
        } else {
            $badgeSucceeded.Visibility = 'Collapsed'
        }
    }

    # Populate region from first build if available
    if ($lblTickerRegion -and $Builds.Count -gt 0 -and $Builds[0].Location) {
        $lblTickerRegion.Text = $Builds[0].Location.ToUpper()
    }
}

# ==============================================================================
# SECTION 11: LOG STREAMING (Worker Runspace)
# ==============================================================================

function Start-LogStreaming {
    <#
    .SYNOPSIS
        Starts a background runspace that polls the container log and enqueues new lines.
    .DESCRIPTION
        The worker runspace calls Get-AzContainerInstanceLog at the configured poll interval.
        It tracks the last-seen log length to send only new (delta) lines to the UI queue.
    #>
    param([PSCustomObject]$Build)

    # Stop any existing stream
    Stop-LogStreaming
    Write-DebugLog "Start-LogStreaming: $($Build.TemplateName) in $($Build.ResourceGroup)"

    $Global:SyncHash.StopFlag = $false
    $Global:SyncHash.IsStreaming = $true
    $Global:SyncHash.SelectedBuild = $Build
    $Global:SyncHash.LastLogLength = 0
    $Global:SyncHash.TotalLines = 0
    $Global:SyncHash.BuildStartTime = $Build.StartTime

    # Parse poll interval
    $PollSec = 5
    try {
        $Val = [int]$txtPollInterval.Text
        if ($Val -ge 1 -and $Val -le 300) { $PollSec = $Val }
    } catch { }
    $Global:SyncHash.PollIntervalSec = $PollSec

    # Create Initial State Session with required modules
    $ISS = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $ISS.ExecutionPolicy = [Microsoft.PowerShell.ExecutionPolicy]::Bypass

    # Create runspace
    $Runspace = [RunspaceFactory]::CreateRunspace($ISS)
    $Runspace.ApartmentState = 'STA'
    $Runspace.ThreadOptions = 'ReuseThread'
    $Runspace.Open()

    $PowerShell = [PowerShell]::Create()
    $PowerShell.Runspace = $Runspace

    # Bridge variables
    $PowerShell.Runspace.SessionStateProxy.SetVariable('SyncHash', $Global:SyncHash)
    $PowerShell.Runspace.SessionStateProxy.SetVariable('BuildInfo', $Build)

    $PowerShell.AddScript({
        # Worker thread — polls Azure Container Instance logs
        $ErrorActionPreference = 'Continue'

        # Force InvariantCulture for consistent formatting
        [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::InvariantCulture

        $RG = $BuildInfo.ResourceGroup
        $CGName = $BuildInfo.ContainerGroup
        $CName = $BuildInfo.ContainerName
        $PollMs = $SyncHash.PollIntervalSec * 1000

        # Initial status message
        $SyncHash.LogQueue.Enqueue(@{
            Text  = "Connecting to container: $CGName"
            Color = "#0078D4"
            Bold  = $true
        })
        $SyncHash.LogQueue.Enqueue(@{
            Text  = "Resource Group: $RG | Container: $CName | Poll: $($SyncHash.PollIntervalSec)s"
            Color = "#888888"
            Bold  = $false
        })
        $SyncHash.LogQueue.Enqueue(@{
            Text  = [string]::new([char]0x2500, 80)
            Color = "#333333"
            Bold  = $false
        })

        # === BACKFILL: Try to fetch the full Packer log from blob storage ===
        # AIB writes the complete customization.log to a storage account in the
        # staging resource group.  ACI stdout is capped at ~4 MB, so long builds
        # lose early lines.  We backfill from the blob first, then switch to
        # ACI polling for live tail.
        $BlobLineCount = 0
        try {
            $SA = Get-AzStorageAccount -ResourceGroupName $RG -ErrorAction SilentlyContinue |
                  Select-Object -First 1
            if ($SA) {
                $Ctx = $SA.Context
                # AIB places logs under   packerlogs/<run-guid>/customization.log
                $Blobs = Get-AzStorageBlob -Container 'packerlogs' -Context $Ctx -ErrorAction SilentlyContinue |
                         Where-Object { $_.Name -like '*/customization.log' } |
                         Sort-Object -Property LastModified -Descending

                $LogBlob = $Blobs | Select-Object -First 1
                if ($LogBlob) {
                    $SyncHash.LogQueue.Enqueue(@{
                        Text  = "[BACKFILL] Fetching full log from blob: $($LogBlob.Name) ($([math]::Round($LogBlob.Length / 1KB, 1)) KB)"
                        Color = "#8764B8"
                        Bold  = $false
                    })

                    # Download blob content to memory
                    $TempFile = [System.IO.Path]::GetTempFileName()
                    try {
                        Get-AzStorageBlobContent -Blob $LogBlob.Name -Container 'packerlogs' `
                            -Context $Ctx -Destination $TempFile -Force -ErrorAction Stop | Out-Null

                        $BlobLines = [System.IO.File]::ReadAllLines($TempFile)
                        $BlobLineCount = $BlobLines.Count

                        foreach ($BLine in $BlobLines) {
                            $Trimmed = $BLine.TrimEnd("`r")
                            if ([string]::IsNullOrWhiteSpace($Trimmed)) { continue }
                            $SyncHash.LogQueue.Enqueue(@{
                                Text  = $Trimmed
                                Color = "__AUTO__"
                                Bold  = $false
                            })
                        }

                        $SyncHash.LogQueue.Enqueue(@{
                            Text  = "[BACKFILL] Loaded $BlobLineCount lines from blob storage"
                            Color = "#107C10"
                            Bold  = $true
                        })
                        $SyncHash.LogQueue.Enqueue(@{
                            Text  = [string]::new([char]0x2500, 80)
                            Color = "#333333"
                            Bold  = $false
                        })
                    } finally {
                        if (Test-Path $TempFile) { Remove-Item $TempFile -Force -ErrorAction SilentlyContinue }
                    }
                }
            }
        } catch {
            $SyncHash.LogQueue.Enqueue(@{
                Text  = "[BACKFILL] Could not fetch blob log: $($_.Exception.Message)"
                Color = "#FFB900"
                Bold  = $false
            })
        }

        # === LIVE TAIL: Poll ACI logs for new lines ===
        $PreviousLineCount = 0
        $ConsecutiveErrors = 0
        $MaxConsecutiveErrors = 10
        $SkippedFirstPoll = $false   # when blob backfill ran, skip first ACI batch

        while (-not $SyncHash.StopFlag) {
            try {
                # Fetch the full log
                $Log = Get-AzContainerInstanceLog `
                    -ResourceGroupName $RG `
                    -ContainerGroupName $CGName `
                    -Name $CName `
                    -ErrorAction Stop

                $ConsecutiveErrors = 0   # Reset on success

                if ($Log) {
                    $Lines = $Log -split "`n"
                    $TotalLines = $Lines.Count

                    # If blob backfill loaded data, skip the first ACI batch
                    # (those lines are already displayed from the blob).
                    if ($BlobLineCount -gt 0 -and -not $SkippedFirstPoll) {
                        $SkippedFirstPoll = $true
                        $PreviousLineCount = $TotalLines
                        $SyncHash.TotalLines = $BlobLineCount + $TotalLines
                        $SyncHash.LogQueue.Enqueue(@{
                            Text  = "[LIVE] ACI buffer has $TotalLines lines (already in blob). Streaming new lines only."
                            Color = "#888888"
                            Bold  = $false
                        })
                    }

                    # Send only NEW lines (delta)
                    if ($TotalLines -gt $PreviousLineCount) {
                        $NewLines = $Lines[$PreviousLineCount..($TotalLines - 1)]
                        foreach ($Line in $NewLines) {
                            $Trimmed = $Line.TrimEnd("`r")
                            if ([string]::IsNullOrWhiteSpace($Trimmed)) { continue }

                            $SyncHash.LogQueue.Enqueue(@{
                                Text  = $Trimmed
                                Color = "__AUTO__"    # Signal the UI to apply color rules
                                Bold  = $false
                            })
                        }
                        $PreviousLineCount = $TotalLines
                        $SyncHash.TotalLines = $BlobLineCount + $TotalLines
                    }
                }
            } catch {
                $ConsecutiveErrors++
                $ErrMsg = $_.Exception.Message

                # Check if container/RG was deleted (build completed/cleanup)
                if ($ErrMsg -match 'ResourceGroupNotFound|ResourceNotFound|ContainerGroupNotFound|not found|does not exist') {
                    $SyncHash.LogQueue.Enqueue(@{
                        Text  = "`n[BUILD COMPLETED] Container has been removed — build finished or was cleaned up."
                        Color = "#FFB900"
                        Bold  = $true
                    })
                    $SyncHash.LogQueue.Enqueue(@{
                        Text  = "Log streaming stopped. Total lines captured: $PreviousLineCount"
                        Color = "#888888"
                        Bold  = $false
                    })
                    break
                }

                if ($ConsecutiveErrors -ge $MaxConsecutiveErrors) {
                    $SyncHash.LogQueue.Enqueue(@{
                        Text  = "[ERROR] Too many consecutive errors ($MaxConsecutiveErrors). Stopping log stream."
                        Color = "#D13438"
                        Bold  = $true
                    })
                    $SyncHash.LogQueue.Enqueue(@{
                        Text  = "Last error: $ErrMsg"
                        Color = "#D13438"
                        Bold  = $false
                    })
                    break
                }

                # Transient error — log it and retry
                $SyncHash.LogQueue.Enqueue(@{
                    Text  = "[WARN] Poll error ($ConsecutiveErrors/$MaxConsecutiveErrors): $ErrMsg"
                    Color = "#FFB900"
                    Bold  = $false
                })
            }

            # Sleep for poll interval
            Start-Sleep -Milliseconds $PollMs
        }

        $SyncHash.IsStreaming = $false
    }) | Out-Null

    # Store runspace references for cleanup
    $Global:WorkerPS = $PowerShell
    $Global:WorkerHandle = $PowerShell.BeginInvoke()

    $lblLogTemplateName.Text = $Build.TemplateName
    $lblStatus.Text = "Streaming log: $($Build.TemplateName)"
}

function Stop-LogStreaming {
    <#
    .SYNOPSIS
        Stops the active log streaming worker.
    #>
    $Global:SyncHash.StopFlag = $true
    if ($Global:WorkerPS) {
        try {
            if ($Global:WorkerHandle -and -not $Global:WorkerHandle.IsCompleted) {
                $Global:WorkerPS.Stop()
            }
            # Dispose runspace first to release the thread, then the PowerShell instance
            try { $Global:WorkerPS.Runspace.Dispose() } catch { }
            $Global:WorkerPS.Dispose()
        } catch { }
        $Global:WorkerPS = $null
        $Global:WorkerHandle = $null
    }
    $Global:SyncHash.IsStreaming = $false
}

# ==============================================================================
# SECTION 12: DISPATCHER TIMER (Queue → UI Bridge)
# ==============================================================================

$Global:CachedBC = [System.Windows.Media.BrushConverter]::new()

$Timer = New-Object System.Windows.Threading.DispatcherTimer
$Timer.Interval = [TimeSpan]::FromMilliseconds(50)
$Timer.Add_Tick({
    $RTB = $rtbOutput
    if (-not $RTB) { return }
    $TickNow = Get-Date   # cache once — avoids 12+ Get-Date syscalls per tick

    # ── Background job completion polling ───────────────────────────────────
    for ($bi = $Global:BgJobs.Count - 1; $bi -ge 0; $bi--) {
        $Job = $Global:BgJobs[$bi]
        $ElapsedSec = ($TickNow - $Job.StartedAt).TotalSeconds
        if ($Job.AsyncResult.IsCompleted) {
            Write-DebugLog "BgJob[$bi]: COMPLETED after $([math]::Round($ElapsedSec,1))s — invoking callback"
            try {
                $BgResult   = $Job.PS.EndInvoke($Job.AsyncResult)
                $BgErrors   = @($Job.PS.Streams.Error)
                Write-DebugLog "BgJob[$bi]: EndInvoke OK, results=$($BgResult.Count) errors=$($BgErrors.Count)"
                if ($BgErrors.Count -gt 0) {
                    foreach ($e in $BgErrors) {
                        Write-DebugLog "BgJob[$bi]: ERROR STREAM: $($e.ToString())"
                    }
                }
                Write-DebugLog "BgJob[$bi]: calling OnComplete..."
                & $Job.OnComplete $BgResult $BgErrors
                Write-DebugLog "BgJob[$bi]: OnComplete finished"
            } catch {
                Write-DebugLog "BgJob[$bi]: callback EXCEPTION: $($_.Exception.Message)"
            }
            try { $Job.PS.Dispose() } catch { }
            try { $Job.Runspace.Dispose() } catch { }
            $Global:BgJobs.RemoveAt($bi)
            Write-DebugLog "BgJob[$bi]: disposed, remaining jobs=$($Global:BgJobs.Count)"
        } else {
            # Heartbeat every 5s — use floor to avoid duplicates across ticks
            $Bucket = [math]::Floor($ElapsedSec / 5)
            if ($Bucket -ge 1 -and -not $Job.LastBucket) { $Job.LastBucket = 0 }
            if ($Bucket -ge 1 -and $Bucket -ne $Job.LastBucket) {
                $Job.LastBucket = $Bucket
                Write-DebugLog "BgJob[$bi]: still running ($([math]::Round($ElapsedSec,0))s)..."
            }
        }
    }

    # Smart scroll latch (Bagel Commander pattern) — moved into queue processing block

    # Process log queue (batched to stay responsive during backfill)
    $ContentAdded = $false
    $BC = $Global:CachedBC
    $RemapLight = $Global:RemapLight
    $LineIndex = $Global:LogStats.TotalLines
    $BatchCount = 0
    $QueueRemaining = $Global:SyncHash.LogQueue.Count

    while ($Global:SyncHash.LogQueue.Count -gt 0 -and $BatchCount -lt $Global:MaxLinesPerTick) {
        $Item = $Global:SyncHash.LogQueue.Dequeue()
        $BatchCount++
        if (-not $Item -or -not $Item.Text) { continue }
      try {
        # Strip ANSI escape codes
        $CleanText = Strip-AnsiCodes $Item.Text

        # Apply color rules
        $Color = $Item.Color
        $Bold = $Item.Bold
        if ($Color -eq "__AUTO__") {
            $Rules = Get-PackerLineColor -Line $CleanText
            $Color = $Rules.Color
            $Bold = $Rules.Bold
        }

        # Update stats (#9)
        Update-LogStats -Color $Color -CleanText $CleanText
        $LineIndex++

        # Minimap dot (#20)
        Add-MinimapDot -Color $Color -LineIndex $LineIndex

        # Phase tracker update
        Update-PhaseFromLine $CleanText

        # Duration flame chart update (#7)
        Update-PhaseTiming $CleanText

        # Package tracker (#R2-5)
        Update-PackageTracker $CleanText

        # Script profiler (#R2-11)
        Update-ScriptProfiler $CleanText

        # Error clustering (#R2-6)
        if ($Color -eq "#D13438") {
            Add-ErrorCluster $CleanText
        }

        # Known issue auto-tagger (#R2-8)
        $KnownMatch = Test-KnownIssue $CleanText
        if ($KnownMatch) {
            Add-KnownIssueTag $KnownMatch $LineIndex
        }

        # Heatmap data point (#R2-14)
        Add-HeatmapPoint -Color $Color -LineIndex $LineIndex

        # Cost estimator: detect VM SKU (#R2-15)
        Detect-VMSku $CleanText

        # Smart notifications + sound (#8, #18)
        if ($Color -eq "#D13438") {
            Notify-OnError $CleanText
            Play-ErrorSound
        }
        if ($CleanText -match 'BUILD COMPLETED|Container has been removed|Build .* finished') {
            $TN = if ($Global:SyncHash.SelectedBuild) { $Global:SyncHash.SelectedBuild.TemplateName } else { "Build" }
            Notify-OnBuildComplete $TN
            if ($Global:SyncHash.SelectedBuild) {
                Add-BuildToHistory $Global:SyncHash.SelectedBuild 'Completed'
                Save-CurrentLogForDiff $TN
                AutoSave-BuildLog $TN 'completed'
                # Check achievements (#R2-10)
                Check-BuildAchievements 'Completed'
                Update-StreakDisplay
                # Confetti on milestone streaks (#R2-9)
                $CurStreak = Get-BuildStreak
                if ($CurStreak -in @(5, 10, 25, 50, 100)) {
                    Start-ConfettiAnimation
                }
            }
        }
        # Failed build detection
        if ($CleanText -match 'Packer build failed|Build .* failed|Error building template|Build not successful') {
            $TN2 = if ($Global:SyncHash.SelectedBuild) { $Global:SyncHash.SelectedBuild.TemplateName } else { 'Build' }
            if ($Global:SyncHash.SelectedBuild) {
                Add-BuildToHistory $Global:SyncHash.SelectedBuild 'Error'
                AutoSave-BuildLog $TN2 'failed'
                Check-BuildAchievements 'Error'
                Update-StreakDisplay
            }
        }

        # Apply light mode remap
        $CanonicalColor = $Color  # preserve un-remapped color for badge matching
        $Color = & $RemapLight $Color

        # Create paragraph + run
        $Para = New-Object System.Windows.Documents.Paragraph
        $Para.Margin = [System.Windows.Thickness]::new(0)
        $Para.LineHeight = 1

        # Severity badge (#1) — use un-remapped color for canonical matching
        $Badge = Get-SeverityBadge $CanonicalColor
        if ($Badge) {
            $BadgeRun = New-Object System.Windows.Documents.Run(" $($Badge.Label) ")
            $BadgeRun.FontSize = 9
            $BadgeRun.FontWeight = 'Bold'
            $BadgeRun.FontFamily = [System.Windows.Media.FontFamily]::new("Consolas")
            try {
                $BadgeRun.Background = $BC.ConvertFromString($Badge.Bg)
                $BadgeRun.Foreground = $BC.ConvertFromString($Badge.Fg)
            } catch { }
            $Para.Inlines.Add($BadgeRun)
            $SpacerRun = New-Object System.Windows.Documents.Run(" ")
            $SpacerRun.FontSize = 9
            $Para.Inlines.Add($SpacerRun)
        }

        # Timestamp delta prefix
        $Delta = Get-RelativeTimestamp $CleanText
        if ($Delta) {
            $DeltaRun = New-Object System.Windows.Documents.Run("$Delta ")
            $DeltaRun.Foreground = $Window.Resources['ThemeTextDisabled']
            $DeltaRun.FontSize = 10
            $Para.Inlines.Add($DeltaRun)
        }

        # JSON structured view (#11) — detect and pretty-print
        $DisplayText = $CleanText
        if (Test-JsonLine $CleanText) {
            $Pretty = Format-JsonPretty $CleanText
            if ($Pretty) { $DisplayText = $Pretty }
        }

        $Run = New-Object System.Windows.Documents.Run($DisplayText)
        try {
            $Run.Foreground = $BC.ConvertFromString($Color)
        } catch {
            $Run.Foreground = $Window.Resources['ThemeTextBody']
        }
        if ($Bold) { $Run.FontWeight = 'Bold' }

        $Para.Inlines.Add($Run)

        $RTB.Document.Blocks.Add($Para)
        $Global:ParagraphCount++
        $ContentAdded = $true
      } catch {
        Write-DebugLog "Tick: line processing error: $($_.Exception.Message)"
      }
    }

    # Backfill progress — show how many lines remain in queue
    $QueueLeft = $Global:SyncHash.LogQueue.Count
    if ($QueueLeft -gt 0 -and $ContentAdded) {
        $lblStatus.Text = "Loading backfill... $QueueLeft lines remaining"
    } elseif ($QueueLeft -eq 0 -and $QueueRemaining -gt $Global:MaxLinesPerTick) {
        $lblStatus.Text = "Backfill complete"
    }

    # Virtual scroll line cap (#15)
    if ($ContentAdded) { Trim-FlowDocument }

    # Tail-follow indicator (#10)
    $IsLatched = ($RTB.VerticalOffset + $RTB.ViewportHeight) -ge ($RTB.ExtentHeight - 10.0)
    if ($Global:IsFollowing -ne $IsLatched) {
        Update-FollowIndicator $IsLatched
    }

    # Auto-scroll if latched
    if ($ContentAdded -and $IsLatched) {
        $RTB.ScrollToEnd()
    }

    # Periodic updates (every ~500ms worth of ticks)
    if ($ContentAdded) {
        Update-StatsDisplay -Now $TickNow
        Update-HealthScore -Now $TickNow
        Update-CostEstimate -Now $TickNow
    }

    # Throttled panel renders (every 3s to avoid excessive redraws)
    # Also force a final render when the queue drains (last batch complete)
    if (-not $Global:LastPanelRender) { $Global:LastPanelRender = [DateTime]::MinValue }
    $QueueJustDrained = $ContentAdded -and $Global:SyncHash.LogQueue.Count -eq 0 -and $QueueRemaining -gt 0
    if (($ContentAdded -and ($TickNow - $Global:LastPanelRender).TotalSeconds -ge 3) -or $QueueJustDrained) {
        $Global:LastPanelRender = $TickNow
        Render-PackageTracker
        Render-ErrorClusters
        Render-KnownIssues
        Render-ScriptProfiler
        Render-Heatmap
    }

    # Live ETA update (every 5s)
    if (-not $Global:LastETAUpdate) { $Global:LastETAUpdate = [DateTime]::MinValue }
    if ($Global:SyncHash.IsStreaming -and ($TickNow - $Global:LastETAUpdate).TotalSeconds -ge 5) {
        $Global:LastETAUpdate = $TickNow
        Update-ETA -Now $TickNow
    }

    # Minimap refresh (throttled — every 2s)
    if (-not $Global:LastMinimapRender) { $Global:LastMinimapRender = [DateTime]::MinValue }
    if ($ContentAdded -and ($TickNow - $Global:LastMinimapRender).TotalSeconds -ge 2) {
        $Global:LastMinimapRender = $TickNow
        Render-Minimap
    }

    # Auto-scan builds (every 60s if enabled) — E4: use cached auth state instead of Get-AzContext
    if ($chkAutoScanBuilds.IsChecked -and $Global:SyncHash.IsStreaming -eq $false) {
        if (-not $Global:LastBuildScan) { $Global:LastBuildScan = $TickNow }
        if (($TickNow - $Global:LastBuildScan).TotalSeconds -ge 60) {
            $Global:LastBuildScan = $TickNow
            if ($Global:IsAuthenticated) {
                Write-DebugLog "Auto-scan: triggering build refresh"
                Get-ActiveBuilds
            }
        }
    }

    # Update line count
    if ($Global:SyncHash.TotalLines -gt 0) {
        $lblLineCount.Text = "Lines: $($Global:SyncHash.TotalLines)"
    }

    # Update elapsed time
    if ($Global:SyncHash.BuildStartTime -and $Global:SyncHash.IsStreaming) {
        $Span = $TickNow.ToUniversalTime() - $Global:SyncHash.BuildStartTime
        $lblElapsed.Text = "Elapsed: {0:hh\:mm\:ss}" -f $Span
    }

    # Status dot pulse (streaming indicator)
    if ($Global:SyncHash.IsStreaming) {
        $Tick = [int]($TickNow.Millisecond / 500)
        $statusDot.Opacity = if ($Tick -eq 0) { 1.0 } else { 0.4 }
    } else {
        $statusDot.Opacity = 1.0
    }
})

# ==============================================================================
# SECTION 13: EVENT HANDLERS
# ==============================================================================

# Title bar: minimize, maximize, close, drag, theme toggle
$btnMinimize.Add_Click({ $Window.WindowState = 'Minimized' })
$btnMaximize.Add_Click({
    if ($Window.WindowState -eq 'Maximized') {
        $Window.WindowState = 'Normal'
    } else {
        $Window.WindowState = 'Maximized'
    }
})
$btnClose.Add_Click({
    Stop-LogStreaming
    $Timer.Stop()
    $Window.Close()
})

# Theme toggle
$btnThemeToggle.Add_Click({
    $Global:IsLightMode = -not $Global:IsLightMode
    & $ApplyTheme $Global:IsLightMode
    $btnThemeToggle.Content = if ($Global:IsLightMode) { [char]0x263E } else { [char]0x2600 }  # moon / sun
    if ($chkDarkMode) { $chkDarkMode.IsChecked = -not $Global:IsLightMode }
})

# Phase durations flyout
if ($btnFlameChartToggle) {
    $btnFlameChartToggle.Add_MouseLeftButtonUp({
        Render-FlameChart
        $popFlameChart.IsOpen = -not $popFlameChart.IsOpen
    })
}

# Collapsible settings panel
if ($pnlSettingsHeader) {
    $pnlSettingsHeader.Add_MouseLeftButtonUp({
        if ($pnlSettingsBody.Visibility -eq 'Collapsed') {
            $pnlSettingsBody.Visibility = 'Visible'
            $lblSettingsChevron.Text = [char]0xE70D   # chevron down
        } else {
            $pnlSettingsBody.Visibility = 'Collapsed'
            $lblSettingsChevron.Text = [char]0xE76C   # chevron right
        }
    })
}

# Dark mode checkbox syncs with theme
if ($chkDarkMode) {
    $chkDarkMode.IsChecked = -not $Global:IsLightMode
    $chkDarkMode.Add_Click({
        if ($Global:SuppressThemeHandler) { return }
        $Global:IsLightMode = -not $chkDarkMode.IsChecked
        & $ApplyTheme $Global:IsLightMode
        $btnThemeToggle.Content = if ($Global:IsLightMode) { [char]0x263E } else { [char]0x2600 }
    })
}

# ==============================================================================
# RIGHT SIDEBAR: HAMBURGER MENU + ICON RAIL (mirrors left rail pattern)
# ==============================================================================

$Global:RightPanelOpen = $true
$Global:RightRailToggles = @{ Stats = $true; Profiler = $true; History = $true }

function Toggle-RightPanel {
    <# Expands or collapses the right analytics detail panel #>
    param([bool]$Open)
    $Global:RightPanelOpen = $Open
    if ($Open) {
        if ($pnlRightSidebar) { $pnlRightSidebar.Visibility = 'Visible' }
        if ($splitterRight)   { $splitterRight.Visibility = 'Visible' }
        if ($colRightPanel)   { $colRightPanel.Width = [System.Windows.GridLength]::new(320); $colRightPanel.MinWidth = 200 }
        if ($colRightSplitter){ $colRightSplitter.Width = [System.Windows.GridLength]::Auto }
    } else {
        if ($pnlRightSidebar) { $pnlRightSidebar.Visibility = 'Collapsed' }
        if ($splitterRight)   { $splitterRight.Visibility = 'Collapsed' }
        if ($colRightPanel)   { $colRightPanel.Width = [System.Windows.GridLength]::new(0); $colRightPanel.MinWidth = 0 }
        if ($colRightSplitter){ $colRightSplitter.Width = [System.Windows.GridLength]::new(0) }
    }
}

function Update-RightRailIndicators {
    <# Highlights toggled-on right rail icons with accent color + indicator bar #>
    $Pairs = @(
        @{ Section = 'Stats';    Indicator = $railRStatsIndicator;    Icon = $railRStatsIcon },
        @{ Section = 'Profiler'; Indicator = $railRProfilerIndicator; Icon = $railRProfilerIcon },
        @{ Section = 'History';  Indicator = $railRHistoryIndicator;  Icon = $railRHistoryIcon }
    )
    foreach ($P in $Pairs) {
        $IsOn = $Global:RightRailToggles[$P.Section]
        $P.Indicator.Visibility = if ($IsOn) { 'Visible' } else { 'Collapsed' }
        $P.Icon.Foreground = if ($IsOn) { $Window.Resources['ThemeAccentLight'] } else { $Window.Resources['ThemeTextMuted'] }
    }
}

function Apply-RightSections {
    <# Shows/hides right panel content groups based on toggle state; auto-collapses if all off #>
    if ($grpRightStats)    { $grpRightStats.Visibility    = if ($Global:RightRailToggles.Stats)    { 'Visible' } else { 'Collapsed' } }
    if ($grpRightProfiler) { $grpRightProfiler.Visibility = if ($Global:RightRailToggles.Profiler) { 'Visible' } else { 'Collapsed' } }
    if ($grpRightHistory)  { $grpRightHistory.Visibility  = if ($Global:RightRailToggles.History)  { 'Visible' } else { 'Collapsed' } }
    # Auto-collapse if nothing is on
    $AnyOn = $Global:RightRailToggles.Values -contains $true
    if (-not $AnyOn -and $Global:RightPanelOpen) { Toggle-RightPanel $false }
}

# Right hamburger: toggle panel open/close (re-enables all if all off)
if ($btnRightHamburger) {
    $btnRightHamburger.Add_Click({
        if ($Global:RightPanelOpen) {
            Toggle-RightPanel $false
        } else {
            # If all sections are off, re-enable all
            if (-not ($Global:RightRailToggles.Values -contains $true)) {
                $Global:RightRailToggles.Stats = $true
                $Global:RightRailToggles.Profiler = $true
                $Global:RightRailToggles.History = $true
                Update-RightRailIndicators
                Apply-RightSections
            }
            Toggle-RightPanel $true
        }
    })
}

# Right rail icon clicks: toggle individual sections
foreach ($RailBtn in @($railRStats, $railRProfiler, $railRHistory)) {
    if ($RailBtn) {
        $RailBtn.Add_Click({
            $Section = $this.Tag
            $Global:RightRailToggles[$Section] = -not $Global:RightRailToggles[$Section]
            Update-RightRailIndicators
            Apply-RightSections
            # If we just toggled something ON and panel is closed, open it
            if ($Global:RightRailToggles[$Section] -and -not $Global:RightPanelOpen) {
                Toggle-RightPanel $true
            }
        })
    }
}

# Initialize: all sections visible on startup
Apply-RightSections

# ==============================================================================
# LEFT SIDEBAR: HAMBURGER MENU + ICON RAIL (Bagel Commander pattern)
# ==============================================================================

$Global:LeftPanelOpen = $true
$Global:LeftRailToggles = @{ Builds = $true; Info = $true; Settings = $true }

function Toggle-LeftPanel {
    <# Expands or collapses the left detail panel #>
    param([bool]$Open)
    $Global:LeftPanelOpen = $Open
    if ($Open) {
        if ($pnlLeftSidebar) { $pnlLeftSidebar.Visibility = 'Visible' }
        if ($splitterLeft)   { $splitterLeft.Visibility = 'Visible' }
        if ($colLeftPanel)   { $colLeftPanel.Width = [System.Windows.GridLength]::new(280); $colLeftPanel.MinWidth = 200 }
        if ($colLeftSplitter){ $colLeftSplitter.Width = [System.Windows.GridLength]::Auto }
    } else {
        if ($pnlLeftSidebar) { $pnlLeftSidebar.Visibility = 'Collapsed' }
        if ($splitterLeft)   { $splitterLeft.Visibility = 'Collapsed' }
        if ($colLeftPanel)   { $colLeftPanel.Width = [System.Windows.GridLength]::new(0); $colLeftPanel.MinWidth = 0 }
        if ($colLeftSplitter){ $colLeftSplitter.Width = [System.Windows.GridLength]::new(0) }
    }
}

function Update-RailIndicators {
    <# Highlights toggled-on left rail icons with accent color + indicator bar #>
    $Pairs = @(
        @{ Section = 'Builds';   Indicator = $railBuildsIndicator;   Icon = $railBuildsIcon },
        @{ Section = 'Info';     Indicator = $railInfoIndicator;     Icon = $railInfoIcon },
        @{ Section = 'Settings'; Indicator = $railSettingsIndicator; Icon = $railSettingsIcon }
    )
    foreach ($P in $Pairs) {
        $IsOn = $Global:LeftRailToggles[$P.Section]
        $P.Indicator.Visibility = if ($IsOn) { 'Visible' } else { 'Collapsed' }
        $P.Icon.Foreground = if ($IsOn) { $Window.Resources['ThemeAccentLight'] } else { $Window.Resources['ThemeTextMuted'] }
    }
}

function Apply-LeftSections {
    <# Shows/hides left panel content sections based on toggle state; auto-collapses if all off #>
    $BuildsOn   = $Global:LeftRailToggles.Builds
    $InfoOn     = $Global:LeftRailToggles.Info
    $SettingsOn = $Global:LeftRailToggles.Settings

    # Builds section: header (row 0) + list (row 1)
    if ($grpLeftBuildsHeader) { $grpLeftBuildsHeader.Visibility = if ($BuildsOn) { 'Visible' } else { 'Collapsed' } }
    if ($lstBuilds)           { $lstBuilds.Visibility           = if ($BuildsOn) { 'Visible' } else { 'Collapsed' } }
    # Only show empty state when Builds section is ON *and* there are no builds in the list
    $hasBuilds = $lstBuilds -and $lstBuilds.ItemsSource -and @($lstBuilds.ItemsSource).Count -gt 0
    if ($pnlEmptyState)       { $pnlEmptyState.Visibility       = if ($BuildsOn -and -not $hasBuilds) { 'Visible' } else { 'Collapsed' } }
    # Adjust row heights
    if ($rowLeftHeader) { $rowLeftHeader.Height = if ($BuildsOn) { [System.Windows.GridLength]::Auto } else { [System.Windows.GridLength]::new(0) } }
    if ($rowLeftBuilds) {
        $rowLeftBuilds.Height    = if ($BuildsOn) { [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) } else { [System.Windows.GridLength]::new(0) }
        $rowLeftBuilds.MinHeight = if ($BuildsOn) { 120 } else { 0 }
    }

    # Info section in row 2
    if ($grpLeftInfo) { $grpLeftInfo.Visibility = if ($InfoOn) { 'Visible' } else { 'Collapsed' } }

    # Settings section in row 2
    if ($grpLeftSettings) { $grpLeftSettings.Visibility = if ($SettingsOn) { 'Visible' } else { 'Collapsed' } }

    # Bottom border (row 2): visible if either Info or Settings is on
    $BottomOn = $InfoOn -or $SettingsOn
    if ($grpLeftBottomBorder) { $grpLeftBottomBorder.Visibility = if ($BottomOn) { 'Visible' } else { 'Collapsed' } }
    if ($rowLeftBottom) {
        $rowLeftBottom.Height = if ($BottomOn) { [System.Windows.GridLength]::new(2, [System.Windows.GridUnitType]::Star) } else { [System.Windows.GridLength]::new(0) }
    }

    # Auto-collapse if nothing is on
    $AnyOn = $Global:LeftRailToggles.Values -contains $true
    if (-not $AnyOn -and $Global:LeftPanelOpen) { Toggle-LeftPanel $false }
}

# Hamburger button: toggle panel open/close (re-enables all if all off)
if ($btnHamburger) {
    $btnHamburger.Add_Click({
        if ($Global:LeftPanelOpen) {
            Toggle-LeftPanel $false
        } else {
            # If all sections are off, re-enable all
            if (-not ($Global:LeftRailToggles.Values -contains $true)) {
                $Global:LeftRailToggles.Builds = $true
                $Global:LeftRailToggles.Info = $true
                $Global:LeftRailToggles.Settings = $true
                Update-RailIndicators
                Apply-LeftSections
            }
            Toggle-LeftPanel $true
        }
    })
}

# Rail icon clicks: toggle individual sections
foreach ($RailBtn in @($railBuilds, $railInfo, $railSettings)) {
    if ($RailBtn) {
        $RailBtn.Add_Click({
            $Section = $this.Tag
            $Global:LeftRailToggles[$Section] = -not $Global:LeftRailToggles[$Section]
            Update-RailIndicators
            Apply-LeftSections
            # If we just toggled something ON and panel is closed, open it
            if ($Global:LeftRailToggles[$Section] -and -not $Global:LeftPanelOpen) {
                Toggle-LeftPanel $true
            }
        })
    }
}

# Initialize left sections
Apply-LeftSections

# Auth
$btnLogin.Add_Click({ Connect-ToAzure })
$btnLogout.Add_Click({ Disconnect-FromAzure })

# Subscription selector: switch context and re-scan (background)
$cmbSubscription.Add_SelectionChanged({
    $SelectedSub = $cmbSubscription.SelectedItem
    if (-not $SelectedSub) { return }

    # ComboBoxItem: Name is .Content, Id is .Tag
    $SubName = $SelectedSub.Content
    $SubId   = $SelectedSub.Tag

    # Avoid re-triggering when we programmatically set the index
    $CurrentContext = Get-AzContext -ErrorAction SilentlyContinue
    if ($CurrentContext -and $CurrentContext.Subscription.Id -eq $SubId) { return }

    Write-DebugLog "Switching subscription to: $SubName ($SubId)"
    $lblStatus.Text = "Switching to $SubName..."
    $Window.Cursor = [System.Windows.Input.Cursors]::Wait

    Start-BackgroundWork -Variables @{ TargetSubId = $SubId; TargetSubName = $SubName } -Work {
        $R = @{ Success = $false; Error = '' }
        try {
            Set-AzContext -SubscriptionId $TargetSubId -ErrorAction Stop | Out-Null
            $R.Success = $true
        } catch {
            $R.Error = $_.Exception.Message
        }
        return $R
    } -OnComplete {
        param($Results, $Errors)
        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ Success = $false; Error = 'No result' } }
        if ($R.Success) {
            $lblStatus.Text = "Switched subscription — scanning builds..."
            Write-DebugLog "Subscription switch OK — rescanning"
            $Window.Cursor = [System.Windows.Input.Cursors]::Arrow
            Get-ActiveBuilds
        } else {
            $lblStatus.Text = "Failed to switch subscription: $($R.Error)"
            Write-DebugLog "Subscription switch FAILED: $($R.Error)"
            $Window.Cursor = [System.Windows.Input.Cursors]::Arrow
        }
    }
})

# Build list: refresh
$btnRefreshBuilds.Add_Click({
    $Context = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $Context -or -not $Context.Account) {
        $lblStatus.Text = "Sign in first to discover builds"
        return
    }
    Get-ActiveBuilds
})

# Ticker strip: Scan Builds button removed (was redundant with btnRefreshBuilds)

# Build list: selection changed → start log streaming
$lstBuilds.Add_SelectionChanged({
    $Selected = $lstBuilds.SelectedItem
    if (-not $Selected) { return }

    # Clear existing log & reset phase tracker
    $rtbOutput.Document.Blocks.Clear()
    $Global:ParagraphCount = 0
    Reset-PhaseTracker
    Reset-LogStats
    Reset-FlameChart
    Reset-Minimap
    Reset-PackageTracker
    Reset-ErrorClusters
    Reset-KnownIssues
    Reset-ScriptProfiler
    Reset-Heatmap
    $Global:DetectedVMSku = $null
    $Global:EstimatedCostPerHr = $Global:VMCostTable['_default']
    $Global:ETAPercent = 0
    $Global:ETARemaining = ''
    $Global:HealthScore = 'N/A'
    $Global:HealthGrade = 'N/A'
    if ($lblHealthGrade)  { $lblHealthGrade.Text = 'N/A' }
    if ($lblHealthScore)  { $lblHealthScore.Text = 'N/A' }
    if ($lblETAPercent)   { $lblETAPercent.Text = '0%' }
    if ($lblETARemaining) { $lblETARemaining.Text = 'No data for ETA' }
    if ($prgETA)          { $prgETA.Value = 0 }
    if ($lblCostEstimate) { $lblCostEstimate.Text = '$0.00' }
    if ($lblCostRate)     { $lblCostRate.Text = '' }
    if ($lblCostSku)      { $lblCostSku.Text = '' }

    # Start streaming
    Start-LogStreaming -Build $Selected
})

# Log toolbar: clear
$btnClearLog.Add_Click({
    $rtbOutput.Document.Blocks.Clear()
    $Global:ParagraphCount = 0
    $Global:SyncHash.LastLogLength = 0
    $lblLineCount.Text = "Lines: 0"
    Reset-PhaseTracker
    Reset-LogStats
    Reset-FlameChart
    Reset-Minimap
    Reset-PackageTracker
    Reset-ErrorClusters
    Reset-KnownIssues
    Reset-ScriptProfiler
    Reset-Heatmap
    $Global:DetectedVMSku = $null
    $Global:EstimatedCostPerHr = $Global:VMCostTable['_default']
    $Global:ETAPercent = 0
    $Global:ETARemaining = ''
    $Global:HealthScore = 'N/A'
    $Global:HealthGrade = 'N/A'
    if ($lblHealthGrade)  { $lblHealthGrade.Text = 'N/A' }
    if ($lblHealthScore)  { $lblHealthScore.Text = 'N/A' }
    if ($lblETAPercent)   { $lblETAPercent.Text = '0%' }
    if ($lblETARemaining) { $lblETARemaining.Text = 'No data for ETA' }
    if ($prgETA)          { $prgETA.Value = 0 }
    if ($lblCostEstimate) { $lblCostEstimate.Text = '$0.00' }
    if ($lblCostRate)     { $lblCostRate.Text = '' }
    if ($lblCostSku)      { $lblCostSku.Text = '' }
})

# Log toolbar: save + copy
$btnSaveLog.Add_Click({ Export-LogToFile })
$btnCopyLog.Add_Click({ Copy-LogToClipboard })

# Search: highlight on text change (debounced via 300ms delay)
$Global:SearchTimer = New-Object System.Windows.Threading.DispatcherTimer
$Global:SearchTimer.Interval = [TimeSpan]::FromMilliseconds(300)
$Global:SearchTimer.Add_Tick({
    $Global:SearchTimer.Stop()
    Search-LogHighlight $txtSearch.Text
})
$txtSearch.Add_TextChanged({
    $Global:SearchTimer.Stop()
    $Global:SearchTimer.Start()
})

# Search: prev/next buttons
$btnSearchPrev.Add_Click({ Search-Prev })
$btnSearchNext.Add_Click({ Search-Next })

# Keyboard shortcuts: F3 = next, Shift+F3 = prev, Ctrl+F = focus search, Ctrl+S = save
$Window.Add_KeyDown({
    param($s, $e)
    if ($e.Key -eq 'F3') {
        if ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Shift) {
            Search-Prev
        } else {
            Search-Next
        }
        $e.Handled = $true
    }
    elseif ($e.Key -eq 'F' -and ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control)) {
        $txtSearch.Focus()
        $txtSearch.SelectAll()
        $e.Handled = $true
    }
    elseif ($e.Key -eq 'S' -and ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control)) {
        Export-LogToFile
        $e.Handled = $true
    }
    elseif ($e.Key -eq 'C' -and ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control) -and
           ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Shift)) {
        Copy-LogToClipboard
        $e.Handled = $true
    }
})

# Word wrap toggle
$chkWordWrap.Add_Checked({
    $rtbOutput.Document.PageWidth = [double]::NaN
    $rtbOutput.HorizontalScrollBarVisibility = 'Disabled'
})
$chkWordWrap.Add_Unchecked({
    $rtbOutput.Document.PageWidth = 5000
    $rtbOutput.HorizontalScrollBarVisibility = 'Auto'
})

# Copy prereq command
$btnCopyCommand.Add_Click({
    if ($txtPrereqCommand.Text) {
        [System.Windows.Clipboard]::SetText($txtPrereqCommand.Text)
        $lblStatus.Text = "Command copied to clipboard"
    }
})

# Retry prereq check
$btnRetryPrereq.Add_Click({
    $Results = Test-Prerequisites
    $AllGood = Show-PrereqStatus -Results $Results
    if ($AllGood) {
        $lblStatus.Text = "All modules available — ready to connect"
    }
})

# Poll interval validation
$txtPollInterval.Add_LostFocus({
    try {
        $Val = [int]$txtPollInterval.Text
        if ($Val -lt 1) { $txtPollInterval.Text = "1" }
        if ($Val -gt 300) { $txtPollInterval.Text = "300" }
    } catch {
        $txtPollInterval.Text = "5"
    }
})

# Debug toggle
$chkDebug.Add_Checked({
    $pnlDebugOverlay.Visibility = 'Visible'
    Write-DebugLog "Debug overlay enabled"
})
$chkDebug.Add_Unchecked({
    $pnlDebugOverlay.Visibility = 'Collapsed'
})
$btnDebugClear.Add_Click({
    $Global:DebugSB.Clear()
    $Global:DebugLineCount = 0
    $txtDebugLog.Text   = ''
    $lblDebugCount.Text = ''
})
$btnDebugClose.Add_Click({
    $pnlDebugOverlay.Visibility = 'Collapsed'
    $chkDebug.IsChecked = $false
})

# Sound alerts toggle (#18)
if ($chkSoundAlerts) {
    $chkSoundAlerts.Add_Checked({  $Global:SoundEnabled = $true  })
    $chkSoundAlerts.Add_Unchecked({ $Global:SoundEnabled = $false })
}



# Diff button (#12)
if ($btnDiffLog) {
    $btnDiffLog.Add_Click({ Show-DiffDialog })
}

# Follow indicator click to re-latch (#10)
if ($lblFollowIndicator) {
    $lblFollowIndicator.Add_MouseLeftButtonDown({
        $rtbOutput.ScrollToEnd()
        Update-FollowIndicator $true
    })
}

# Minimap click to jump (#20)
if ($cnvMinimap) {
    $cnvMinimap.Add_MouseLeftButtonDown({
        param($s, $e)
        $ClickY = $e.GetPosition($cnvMinimap).Y
        $Ratio = $ClickY / $cnvMinimap.ActualHeight
        $TargetOffset = $Ratio * $rtbOutput.ExtentHeight
        $rtbOutput.ScrollToVerticalOffset($TargetOffset)
    })
}

# Heatmap click to jump (#R2-14)
if ($cnvHeatmap) {
    $cnvHeatmap.Add_MouseLeftButtonDown({
        param($s, $e)
        $ClickX = $e.GetPosition($cnvHeatmap).X
        $Ratio = $ClickX / $cnvHeatmap.ActualWidth
        $TargetOffset = $Ratio * $rtbOutput.ExtentHeight
        $rtbOutput.ScrollToVerticalOffset($TargetOffset)
    })
}

# Build history - Load on startup
Load-BuildHistory
Load-Achievements

# Window closing cleanup — save preferences
$Window.Add_Closing({
    Save-UserPrefs
    Stop-LogStreaming
    $Timer.Stop()
})

# ==============================================================================
# SECTION 14: STARTUP SEQUENCE
# ==============================================================================

# 1. Check prerequisites (no auto-install — surface error with remediation steps)
Write-DebugLog "Startup: checking prerequisites..."
$PrereqResults = Test-Prerequisites
$AllModulesPresent = Show-PrereqStatus -Results $PrereqResults
Write-DebugLog "Prerequisites: AllPresent=$AllModulesPresent"

if ($AllModulesPresent) {
    $lblStatus.Text = "Ready — sign in to start"

    # 2. Check if already authenticated — run in background so the window
    #    appears instantly and never shows "Not Responding".
    Write-DebugLog "Startup: calling Get-AzContext..."
    $ExistingContext = Get-AzContext -ErrorAction SilentlyContinue
    Write-DebugLog "Startup: Get-AzContext returned (Account=$($ExistingContext.Account.Id))"
    if ($ExistingContext -and $ExistingContext.Account) {
        Write-DebugLog "Startup: cached context found for $($ExistingContext.Account.Id)"
        $lblAuthStatus.Text = $ExistingContext.Account.Id
        $lblStatus.Text = "Validating cached token for $($ExistingContext.Account.Id)..."

        $StartupAcct   = $ExistingContext.Account.Id
        $StartupSubName = $ExistingContext.Subscription.Name
        $StartupSubId   = $ExistingContext.Subscription.Id

        Write-DebugLog "Startup: dispatching token validation to background..."
        Start-BackgroundWork -Variables @{
            AcctId  = $StartupAcct
            SubName = $StartupSubName
            SubId   = $StartupSubId
        } -Work {
            $R = @{ TokenValid = $false; AccountId = $AcctId; SubName = $SubName; SubId = $SubId; Subs = @(); Error = '' }
            try {
                Get-AzResourceGroup -ErrorAction Stop | Select-Object -First 1 | Out-Null
                $R.TokenValid = $true
                $RawSubs = @(Get-AzSubscription -ErrorAction SilentlyContinue |
                            Where-Object { $_.State -eq 'Enabled' })
                $R.Subs = @($RawSubs | ForEach-Object {
                    [PSCustomObject]@{ Name = $_.Name; Id = $_.Id }
                })
            } catch {
                $R.Error = $_.Exception.Message
            }
            return $R
        } -OnComplete {
            param($Results, $Errors)
            Write-DebugLog "Startup OnComplete: >>> ENTER"
            $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ TokenValid = $false; Error = 'No result' } }
            Write-DebugLog "Startup OnComplete: tokenValid=$($R.TokenValid) subs=$($R.Subs.Count) error='$($R.Error)'"
            if ($R.TokenValid) {
                Write-DebugLog "Startup OnComplete: calling Set-AuthUIConnected..."
                Set-AuthUIConnected -AccountId $R.AccountId -SubName $R.SubName -SubId $R.SubId -Subscriptions $R.Subs
                $lblStatus.Text = "Logged in as $($R.AccountId) — scanning builds..."
                Write-DebugLog "Startup OnComplete: calling Get-ActiveBuilds..."
                Get-ActiveBuilds
            } else {
                Set-AuthUIDisconnected -Reason "Token expired for $($R.AccountId)"
                $lblStatus.Text = "Cached token expired — click Sign In to re-authenticate"
                Write-DebugLog "Startup OnComplete: token expired — $($R.Error)"
            }
            Write-DebugLog "Startup OnComplete: <<< EXIT"
        }
    }
} else {
    $lblStatus.Text = "Missing required modules — see error panel"
}

# 3. Start the dispatcher timer (must start BEFORE ShowDialog so background job callbacks work)
$Timer.Start()

# 4. Render build history, achievements, streak, and load preferences
Render-BuildHistory
Render-Achievements
Update-StreakDisplay

# 5. Load user preferences (applies saved theme, toggles, window position)
Load-UserPrefs

# 5b. Apply word wrap state programmatically (fixes B5 — XAML initial state mismatch)
if ($chkWordWrap.IsChecked) {
    $rtbOutput.Document.PageWidth = [double]::NaN
    $rtbOutput.HorizontalScrollBarVisibility = 'Disabled'
} else {
    $rtbOutput.Document.PageWidth = 5000
    $rtbOutput.HorizontalScrollBarVisibility = 'Auto'
}

# 6. Default to dark mode (user can toggle via Settings > Dark mode)
if (-not (Test-Path $Global:PrefsPath)) {
    $Global:IsLightMode = $false
    & $ApplyTheme $false
    if ($btnThemeToggle) { $btnThemeToggle.Content = [char]0x2600 }
    if ($chkDarkMode)    { $chkDarkMode.IsChecked = $true }
}

# 7. Bring window to front on launch, then release Topmost so it doesn't
#    permanently cover other apps.
$Window.Topmost = $true
$Window.Add_ContentRendered({
    $Window.Topmost = $false
    $Window.Activate()
})
$Window.ShowDialog() | Out-Null

# 8. Cleanup on exit
Stop-LogStreaming
$Timer.Stop()
# Dispose NotifyIcon
if ($Global:NotifyIcon) { try { $Global:NotifyIcon.Visible = $false; $Global:NotifyIcon.Dispose() } catch { } }
# Dispose any remaining background jobs
for ($i = $Global:BgJobs.Count - 1; $i -ge 0; $i--) {
    try { $Global:BgJobs[$i].PS.Dispose() } catch { }
    try { $Global:BgJobs[$i].Runspace.Dispose() } catch { }
}
