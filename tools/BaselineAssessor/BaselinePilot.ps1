#Requires -Version 5.1
<#
.SYNOPSIS
    BaselinePilot - Windows 11 Client Security Baseline Assessment Tool.
.DESCRIPTION
    PowerShell/WPF tool for assessing Windows 11 client machines against the
    Microsoft Security Baselines (Intune MDM 24H2 + GPO SCT 25H2) plus
    operational health checks. Supports headless collection import, automated
    check evaluation, weighted scoring, and rich HTML report export.
.NOTES
    Author : Anton Romanyuk
    Version: 0.1.0
    Date   : 2026-03-31
#>

# ===============================================================================
# SECTION 1: PRE-LOAD & INITIALIZATION
# ===============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'

$env:PSModulePath = ($env:PSModulePath -split ';' |
    Where-Object { $_ -notlike '*OneDrive*' }) -join ';'

$Global:Root = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Global:Root)) { $Global:Root = $PWD.Path }

$Global:AppVersion       = "0.1.0-alpha"
$Global:AppTitle         = "BaselinePilot v$($Global:AppVersion)"
$Global:PrefsPath        = Join-Path $Global:Root "user_prefs.json"
$Global:AssessmentDir    = Join-Path $Global:Root "assessments"
$Global:ReportDir        = Join-Path $Global:Root "reports"
$Global:DebugLogFile     = Join-Path $env:TEMP "BaselinePilot_debug.log"

$Global:AutoSaveFile      = Join-Path $Global:Root "_autosave.json"
$Global:AutoSaveEnabled   = $false
$Global:AutoSaveInterval  = 60
$Global:OpenAfterExport   = $true
$Global:VerboseLogging    = $false
$Global:MaxBackups        = 10

foreach ($dir in @($Global:AssessmentDir, $Global:ReportDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
        Write-Host "[INIT] Created directory: $dir" -ForegroundColor DarkGray
    }
}

$Script:LOG_MAX_LINES       = 500
$Script:TIMER_INTERVAL_MS   = 50
$Script:TOAST_DURATION_MS   = 4000

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName System.Windows.Forms

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
# ===============================================================================

$Global:SyncHash = [Hashtable]::Synchronized(@{
    StatusQueue  = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())
    LogQueue     = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())
    StopFlag     = $false
})

$Global:BgJobs          = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
$Global:TimerProcessing = $false

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
# ===============================================================================

$XamlPath = Join-Path $Global:Root "BaselinePilot_UI.xaml"
if (-not (Test-Path $XamlPath)) {
    [System.Windows.MessageBox]::Show(
        "CRITICAL: 'BaselinePilot_UI.xaml' not found in:`n$Global:Root",
        "BaselinePilot", 'OK', 'Error') | Out-Null
    exit 1
}

$XamlContent = Get-Content $XamlPath -Raw -Encoding UTF8
Write-Host "[XAML] Loaded XAML: $([math]::Round($XamlContent.Length / 1024, 1)) KB" -ForegroundColor DarkGray
$XamlContent = $XamlContent -replace 'Title="BaselinePilot"', "Title=`"$Global:AppTitle`""

try {
    $Window = [Windows.Markup.XamlReader]::Parse($XamlContent)
    Write-Host "[XAML] Window parsed successfully" -ForegroundColor DarkGray
} catch {
    [System.Windows.MessageBox]::Show(
        "XAML Parse Error:`n$($_.Exception.Message)",
        "BaselinePilot", 'OK', 'Error') | Out-Null
    exit 1
}

# ===============================================================================
# SECTION 5: ELEMENT REFERENCES
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
$Global:ActiveFilePath = $null

# Icon rail
$railDashboard     = $Window.FindName("railDashboard")
$railBaseline      = $Window.FindName("railBaseline")
$railFindings      = $Window.FindName("railFindings")
$railReport        = $Window.FindName("railReport")
$railSettings      = $Window.FindName("railSettings")
$railDashboardIndicator  = $Window.FindName("railDashboardIndicator")
$railBaselineIndicator   = $Window.FindName("railBaselineIndicator")
$railFindingsIndicator   = $Window.FindName("railFindingsIndicator")
$railReportIndicator     = $Window.FindName("railReportIndicator")
$railSettingsIndicator   = $Window.FindName("railSettingsIndicator")

# Sidebar panels
$pnlSidebar              = $Window.FindName("pnlSidebar")
$pnlSidebarDashboard     = $Window.FindName("pnlSidebarDashboard")
$pnlSidebarBaseline      = $Window.FindName("pnlSidebarBaseline")
$pnlSidebarFindings      = $Window.FindName("pnlSidebarFindings")
$pnlSidebarReport        = $Window.FindName("pnlSidebarReport")
$pnlSidebarSettings      = $Window.FindName("pnlSidebarSettings")

# Sidebar — Dashboard
$txtCustomerName   = $Window.FindName("txtCustomerName")
$txtAssessorName   = $Window.FindName("txtAssessorName")
$txtAssessmentDate = $Window.FindName("txtAssessmentDate")
$bdrJoinType       = $Window.FindName("bdrJoinType")
$lblJoinTypeIcon   = $Window.FindName("lblJoinTypeIcon")
$lblJoinType       = $Window.FindName("lblJoinType")
$lblStatTotalChecks = $Window.FindName("lblStatTotalChecks")
$lblStatPass       = $Window.FindName("lblStatPass")
$lblStatFail       = $Window.FindName("lblStatFail")
$lblStatWarning    = $Window.FindName("lblStatWarning")
$lblStatNA         = $Window.FindName("lblStatNA")
$btnImportCollection = $Window.FindName("btnImportCollection")
$btnSaveAssessment = $Window.FindName("btnSaveAssessment")
$btnNewAssessment  = $Window.FindName("btnNewAssessment")
$lstSavedAssessments = $Window.FindName("lstSavedAssessments")
$lblSavedCount     = $Window.FindName("lblSavedCount")
$lblSavedEmpty     = $Window.FindName("lblSavedEmpty")

# Sidebar — Baseline
$pnlCategoryList   = $Window.FindName("pnlCategoryList")
$lblProgressChecked = $Window.FindName("lblProgressChecked")
$barProgress       = $Window.FindName("barProgress")
$chkShowDeviationsOnly = $Window.FindName("chkShowDeviationsOnly")
$cmbBaselinePriority   = $Window.FindName("cmbBaselinePriority")

# Sidebar — Findings filter
$cmbFilterSeverity = $Window.FindName("cmbFilterSeverity")
$cmbFilterStatus   = $Window.FindName("cmbFilterStatus")
$cmbFilterCategory = $Window.FindName("cmbFilterCategory")
$cmbFilterPriority = $Window.FindName("cmbFilterPriority")
$cmbFilterOrigin   = $Window.FindName("cmbFilterOrigin")
$cmbFilterEffort   = $Window.FindName("cmbFilterEffort")
$txtFindingsSearch = $Window.FindName("txtFindingsSearch")

# Sidebar — Report
$btnExportHtml     = $Window.FindName("btnExportHtml")
$btnExportCsv      = $Window.FindName("btnExportCsv")
$btnExportJson     = $Window.FindName("btnExportJson")
$cmbScoreView      = $Window.FindName("cmbScoreView")
$lblOverallScore   = $Window.FindName("lblOverallScore")
$lblOverallLabel   = $Window.FindName("lblOverallLabel")
$btnCopySummary    = $Window.FindName("btnCopySummary")

# Settings
$chkDarkMode         = $Window.FindName("chkDarkMode")
$txtDefaultAssessor  = $Window.FindName("txtDefaultAssessor")
$cmbBaselineVersion  = $Window.FindName("cmbBaselineVersion")
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
$pnlTabBaseline    = $Window.FindName("pnlTabBaseline")
$pnlTabFindings    = $Window.FindName("pnlTabFindings")
$pnlTabReport      = $Window.FindName("pnlTabReport")
$pnlTabSettings    = $Window.FindName("pnlTabSettings")

# Dashboard content
$bdrScoreHero       = $Window.FindName("bdrScoreHero")
$lblHeroScore       = $Window.FindName("lblHeroScore")
$lblHeroRiskScore   = $Window.FindName("lblHeroRiskScore")
$lblHeroHostname    = $Window.FindName("lblHeroHostname")
$lblHeroOS          = $Window.FindName("lblHeroOS")
$lblHeroBuild       = $Window.FindName("lblHeroBuild")
$lblHeroJoinType    = $Window.FindName("lblHeroJoinType")
$lblHeroCollected   = $Window.FindName("lblHeroCollected")
$pnlHeroDonut      = $Window.FindName("pnlHeroDonut")
$pnlHeroLegend     = $Window.FindName("pnlHeroLegend")
$pnlScoreCards      = $Window.FindName("pnlScoreCards")
$pnlCategoryBars    = $Window.FindName("pnlCategoryBars")
$pnlMaturityAlignment = $Window.FindName("pnlMaturityAlignment")
$pnlTopFindings     = $Window.FindName("pnlTopFindings")

# Baseline content
$pnlBaselineChecks = $Window.FindName("pnlBaselineChecks")

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
# ===============================================================================

$Global:FullLogSB        = [System.Text.StringBuilder]::new()
$Global:FullLogLineCount = 0
$Global:FullLogMaxLines  = 1000
$Global:DebugLineCount   = 0
$Global:DebugMaxLines    = $Script:LOG_MAX_LINES
$Global:DebugOverlayEnabled = $false

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
# ===============================================================================

$ApplyTheme = {
    param([bool]$IsLight)
    $Palette = if ($IsLight) { $Global:ThemeLight } else { $Global:ThemeDark }

    foreach ($Key in $Palette.Keys) {
        try {
            $NewColor = [System.Windows.Media.ColorConverter]::ConvertFromString($Palette[$Key])
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
# ===============================================================================

function Show-Toast {
    param(
        [string]$Message,
        [ValidateSet('Success','Error','Warning','Info')][string]$Type = 'Info',
        [int]$DurationMs = $Script:TOAST_DURATION_MS
    )
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
# ===============================================================================

function Show-ThemedDialog {
    param(
        [string]$Title    = 'BaselinePilot',
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

    # Header
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

    $Sep = [System.Windows.Controls.Border]::new()
    $Sep.Height = 1
    $Sep.Margin = [System.Windows.Thickness]::new(0,0,0,16)
    $Sep.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeBorder')
    [void]$MainStack.Children.Add($Sep)

    if ($Message) {
        $MsgBlock = [System.Windows.Controls.TextBlock]::new()
        $MsgBlock.Text = $Message
        $MsgBlock.FontSize   = 12
        $MsgBlock.TextWrapping = 'Wrap'
        $MsgBlock.Margin     = [System.Windows.Thickness]::new(0,0,0,24)
        $MsgBlock.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextSecondary')
        [void]$MainStack.Children.Add($MsgBlock)
    }

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

    $OuterBorder.Add_MouseLeftButtonDown({
        try { $Dlg.DragMove() } catch { }
    }.GetNewClosure())

    $Dlg.ShowDialog() | Out-Null
    $DlgResult = if ($Dlg.Tag) { $Dlg.Tag } else { 'Cancel' }
    return $DlgResult
}

# ===============================================================================
# SECTION 9: TAB SWITCHING
# ===============================================================================

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

$Global:ActivityLogOpen       = $true
$Global:ActivityLogSavedHeight = 160

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
    if ($chkShowActivityLog -and $chkShowActivityLog.IsChecked -ne $Show) {
        $chkShowActivityLog.IsChecked = $Show
    }
}

function Switch-Tab {
    param([string]$TabName)
    $Global:ActiveTabName = $TabName

    # Hide all content panels
    $pnlTabDashboard.Visibility  = 'Collapsed'
    $pnlTabBaseline.Visibility   = 'Collapsed'
    $pnlTabFindings.Visibility   = 'Collapsed'
    $pnlTabReport.Visibility     = 'Collapsed'
    $pnlTabSettings.Visibility   = 'Collapsed'

    # Hide all sidebar panels
    $pnlSidebarDashboard.Visibility  = 'Collapsed'
    $pnlSidebarBaseline.Visibility   = 'Collapsed'
    $pnlSidebarFindings.Visibility   = 'Collapsed'
    $pnlSidebarReport.Visibility     = 'Collapsed'
    $pnlSidebarSettings.Visibility   = 'Collapsed'

    # Reset rail indicators
    foreach ($ind in @($railDashboardIndicator, $railBaselineIndicator, $railFindingsIndicator,
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
        'Baseline' {
            Invoke-TabFade $pnlTabBaseline
            $pnlSidebarBaseline.Visibility = 'Visible'
            if ($railBaselineIndicator) { $railBaselineIndicator.Background = $AccentBrush }
            # Defer heavy rendering so tab shows instantly
            Invoke-DeferredRender { Render-BaselineChecks }
        }
        'Findings' {
            Invoke-TabFade $pnlTabFindings
            $pnlSidebarFindings.Visibility = 'Visible'
            if ($railFindingsIndicator) { $railFindingsIndicator.Background = $AccentBrush }
            Invoke-DeferredRender { Render-Findings }
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

# Deferred rendering: schedules work on the Dispatcher at Background priority
# so the tab shows immediately, then content populates without freezing UI
function Invoke-DeferredRender {
    param([ScriptBlock]$Work)
    $Window.Dispatcher.BeginInvoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [Action]$Work
    )
}

# ===============================================================================
# SECTION 10: ASSESSMENT DATA MODEL
# ===============================================================================

$Global:Assessment = [PSCustomObject]@{
    CustomerName    = ''
    AssessorName    = ''
    Date            = (Get-Date -Format 'yyyy-MM-dd')
    Notes           = ''
    CollectionData  = $null   # Primary collection JSON (latest import)
    JoinType        = ''      # 'EntraID', 'ADDS', 'Hybrid', or ''
    Checks          = [System.Collections.ArrayList]::new()
    ManualOverrides = @{}
    # Multi-machine support
    Machines        = [System.Collections.ArrayList]::new()  # All imported machines
    MachineCount    = 0
}
$Global:CatScoreRefs = @{}

# Load check definitions from external JSON
$Global:ChecksJsonPath = Join-Path $Global:Root 'checks.json'
if (-not (Test-Path $Global:ChecksJsonPath)) {
    [System.Windows.MessageBox]::Show(
        "CRITICAL: 'checks.json' not found in:`n$Global:Root",
        'BaselinePilot', 'OK', 'Error') | Out-Null
    exit 1
}

try {
    $ChecksFile = Get-Content $Global:ChecksJsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $Global:CheckDefinitions = $ChecksFile.checks
    Write-Host "[INIT] Loaded $($Global:CheckDefinitions.Count) check definitions from checks.json" -ForegroundColor DarkGray
} catch {
    [System.Windows.MessageBox]::Show(
        "Failed to parse checks.json:`n$($_.Exception.Message)",
        'BaselinePilot', 'OK', 'Error') | Out-Null
    exit 1
}

# Load CSP metadata for enriched remediation guidance
$Global:CspMetadata = @{}
$Global:CspDbAge = 0
$CspPath = Join-Path $Global:Root 'csp_metadata.json'
if (Test-Path $CspPath) {
    try {
        $CspRaw = Get-Content $CspPath -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($prop in $CspRaw.PSObject.Properties) {
            $Global:CspMetadata[$prop.Name] = $prop.Value
        }
        Write-Host "[INIT] Loaded $($Global:CspMetadata.Count) CSP settings from csp_metadata.json" -ForegroundColor DarkGray

        # Staleness check: warn if JSON older than 90 days
        $Global:CspDbAge = ((Get-Date) - (Get-Item $CspPath).LastWriteTime).Days
        if ($Global:CspDbAge -gt 90) {
            Write-Host "[INIT] WARNING: CSP metadata is $($Global:CspDbAge) days old. Run Build-CspDatabase.ps1 to refresh." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "[INIT] Failed to load csp_metadata.json: $($_.Exception.Message)" -ForegroundColor Yellow
    }
} else {
    Write-Host "[INIT] csp_metadata.json not found — run Build-CspDatabase.ps1 to generate, or copy from GPODocumenter" -ForegroundColor Yellow
}

# Lookup CSP enrichment data for a check (returns hashtable with ADMX, GP path, registry, allowed values)
function Get-CspEnrichment {
    param([string]$CspKey)
    if (-not $CspKey -or $Global:CspMetadata.Count -eq 0) { return $null }
    # Try exact match first
    $entry = $Global:CspMetadata[$CspKey]
    if ($entry) { return $entry }
    # Try with Area/Setting format from collectionKey
    # e.g., "defender.RealTimeProtectionEnabled" -> try "Defender/AllowRealtimeMonitoring"
    return $null
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
        Weight         = [int]$Def.weight
        Priority       = $Def.priority       # Baseline or Operational
        Origin         = $Def.origin          # INTUNE, SCT, OPS
        Effort         = $Def.effort
        Impact         = $Def.impact
        Remediation    = $Def.remediation
        Reference      = $Def.reference
        CollectionKeys = $Def.collectionKeys
        BaselineValue  = $Def.baselineValue
        CspPath        = $Def.cspPath         # CSP key for enrichment lookup
        Rationale      = $Def.rationale       # Why this setting matters (from baseline blog) (e.g. "Defender/AllowRealtimeMonitoring")
        applicableTo   = $Def.applicableTo
        threshold      = $Def.threshold
        ActualValue    = $null
        Status         = 'Not Assessed'
        AutoStatus     = $null            # Stores auto-evaluated status; set during eval, used for Undo
        Excluded       = $false
        Details        = ''
        Notes          = ''
        Source         = $Def.type
        AffectedMachines = [System.Collections.ArrayList]::new()  # Multi-machine: tracks which hosts have this finding
    })
}

# ===============================================================================
# SECTION 11: SCORING ENGINE
# ===============================================================================

function Get-Categories {
    return @($Global:Assessment.Checks | Select-Object -ExpandProperty Category -Unique)
}

function Get-CategoryScore {
    param([string]$Category)
    $Checks = @($Global:Assessment.Checks | Where-Object {
        $_.Category -eq $Category -and $_.Status -ne 'N/A' -and -not $_.Excluded
    })
    if ($Checks.Count -eq 0) { return -1 }
    $AnyAssessed = @($Checks | Where-Object { $_.Status -ne 'Not Assessed' }).Count
    if ($AnyAssessed -eq 0) { return -1 }
    $WeightedScore = 0; $TotalWeight = 0
    foreach ($C in $Checks) {
        $W = [math]::Max(1, [int]$C.Weight)
        $Points = switch ($C.Status) {
            'Pass'         { 100 }
            'Warning'      { 50 }
            'Fail'         { 0 }
            'Not Assessed' { 0 }
            default        { 0 }
        }
        $WeightedScore += $Points * $W
        $TotalWeight += $W
    }
    if ($TotalWeight -eq 0) { return -1 }
    return [math]::Round($WeightedScore / $TotalWeight, 0)
}

function Get-OverallScore {
    $AllChecks = @($Global:Assessment.Checks | Where-Object {
        $_.Status -ne 'N/A' -and -not $_.Excluded
    })
    if ($AllChecks.Count -eq 0) { return -1 }
    $AnyAssessed = @($AllChecks | Where-Object { $_.Status -ne 'Not Assessed' }).Count
    if ($AnyAssessed -eq 0) { return -1 }
    $WeightedSum = 0; $TotalWeight = 0
    foreach ($C in $AllChecks) {
        $W = [math]::Max(1, [int]$C.Weight)
        $Points = switch ($C.Status) {
            'Pass'         { 100 }
            'Warning'      { 50 }
            'Fail'         { 0 }
            'Deferred'     { 0 }   # Deferred = not yet fixed, counts as Fail
            'Not Assessed' { 0 }
            default        { 0 }
        }
        $WeightedSum += $Points * $W
        $TotalWeight += $W
    }
    if ($TotalWeight -eq 0) { return -1 }
    return [math]::Round($WeightedSum / $TotalWeight, 0)
}

function Get-BaselineCompliancePercent {
    $Checks = @($Global:Assessment.Checks | Where-Object {
        $_.Status -ne 'N/A' -and $_.Status -ne 'Not Assessed' -and -not $_.Excluded
    })
    if ($Checks.Count -eq 0) { return -1 }
    $Passed = @($Checks | Where-Object { $_.Status -eq 'Pass' }).Count
    return [math]::Round(($Passed / $Checks.Count) * 100, 0)
}

# Maturity dimensions mapped to 9 BaselinePilot categories
$Global:MaturityDimensions = [ordered]@{
    Security           = @{ Categories = @('Security Configuration','User Account Control');  Label = 'Security & UAC' }
    EndpointProtection = @{ Categories = @('Defender & Endpoint Security');                    Label = 'Endpoint Protection' }
    Authentication     = @{ Categories = @('Authentication & Credentials');                    Label = 'Authentication' }
    NetworkSecurity    = @{ Categories = @('Network Security');                                Label = 'Network Security' }
    Monitoring         = @{ Categories = @('Monitoring & Audit','Operations & Health','Performance & Stability'); Label = 'Monitoring & Operations' }
    DataProtection     = @{ Categories = @('Data Protection');                                 Label = 'Data Protection' }
}

function Get-DimensionScore {
    param([string]$DimensionKey)
    $Dim = $Global:MaturityDimensions[$DimensionKey]
    if (-not $Dim) { return -1 }
    $Checks = @($Global:Assessment.Checks | Where-Object {
        $_.Category -in $Dim.Categories -and $_.Status -ne 'N/A' -and -not $_.Excluded
    })
    if ($Checks.Count -eq 0) { return -1 }
    $AnyAssessed = @($Checks | Where-Object { $_.Status -ne 'Not Assessed' }).Count
    if ($AnyAssessed -eq 0) { return -1 }
    $WeightedSum = 0; $TotalWeight = 0
    foreach ($C in $Checks) {
        $W = [math]::Max(1, [int]$C.Weight)
        $Points = switch ($C.Status) { 'Pass' { 100 } 'Warning' { 50 } default { 0 } }
        $WeightedSum += $Points * $W
        $TotalWeight += $W
    }
    if ($TotalWeight -eq 0) { return -1 }
    return [math]::Round($WeightedSum / $TotalWeight, 0)
}

function Get-MaturityLevel {
    param([int]$Score)
    if ($Score -lt 0) { return 'N/A' }
    if ($Score -ge 90) { return 'Optimized' }
    if ($Score -ge 75) { return 'Managed' }
    if ($Score -ge 55) { return 'Defined' }
    if ($Score -ge 35) { return 'Developing' }
    return 'Initial'
}

# ===============================================================================
# SECTION 12: UI RENDERING
# ===============================================================================

function Update-QuickStats {
    $All      = $Global:Assessment.Checks
    $Pass     = @($All | Where-Object { $_.Status -eq 'Pass' }).Count
    $Fail     = @($All | Where-Object { $_.Status -eq 'Fail' }).Count
    $Warn     = @($All | Where-Object { $_.Status -eq 'Warning' }).Count
    $NA       = @($All | Where-Object { $_.Status -eq 'N/A' }).Count
    $lblStatTotalChecks.Text = "$($All.Count)"
    $lblStatPass.Text    = "$Pass"
    $lblStatFail.Text    = "$Fail"
    $lblStatWarning.Text = "$Warn"
    $lblStatNA.Text      = "$NA"
}

function Update-Progress {
    $Total    = $Global:Assessment.Checks.Count
    $Assessed = @($Global:Assessment.Checks | Where-Object { $_.Status -ne 'Not Assessed' }).Count
    $lblProgressChecked.Text = "$Assessed / $Total checks evaluated"
    $barProgress.Maximum = $Total
    $barProgress.Value   = $Assessed
}

function Update-Dashboard {
    Update-QuickStats

    $Compliance = Get-BaselineCompliancePercent
    $RiskScore  = Get-OverallScore
    $lblHeroScore.Text     = if ($Compliance -ge 0) { "$Compliance%" } else { [char]0x2014 + '%' }
    $lblHeroRiskScore.Text = if ($RiskScore -ge 0)  { "$RiskScore"   } else { [string][char]0x2014 }

    if ($Compliance -ge 80) { $lblHeroScore.Foreground = $Window.Resources['ThemeSuccess'] }
    elseif ($Compliance -ge 50) { $lblHeroScore.Foreground = $Window.Resources['ThemeWarning'] }
    elseif ($Compliance -ge 0) { $lblHeroScore.Foreground = $Window.Resources['ThemeError'] }
    else { $lblHeroScore.Foreground = $Window.Resources['ThemeAccent'] }

    if ($Global:Assessment.CollectionData) {
        $SI = $Global:Assessment.CollectionData.systemInfo
        if ($SI) {
            $lblHeroHostname.Text  = if ($SI.ComputerName) { $SI.ComputerName } else { [string][char]0x2014 }
            $lblHeroOS.Text        = if ($SI.OSCaption) { $SI.OSCaption } else { [string][char]0x2014 }
            $lblHeroBuild.Text     = if ($SI.OSBuild) { $SI.OSBuild } else { [string][char]0x2014 }
            $lblHeroCollected.Text = if ($Global:Assessment.CollectionData._metadata.collectedAt) {
                $Global:Assessment.CollectionData._metadata.collectedAt
            } else { [string][char]0x2014 }
        }
        $lblHeroJoinType.Text = if ($Global:Assessment.JoinType) { $Global:Assessment.JoinType } else { [string][char]0x2014 }
    }

    $lblOverallScore.Text = if ($RiskScore -ge 0) { "$RiskScore%" } else { [char]0x2014 + '%' }

    # ═══════════════════════════════════════════════════════
    # DONUT CHART — rendered into hero card
    # ═══════════════════════════════════════════════════════
    $pnlHeroDonut.Children.Clear()
    $pnlHeroLegend.Children.Clear()

    $All      = $Global:Assessment.Checks
    $PassCnt  = @($All | Where-Object { $_.Status -eq 'Pass' }).Count
    $FailCnt  = @($All | Where-Object { $_.Status -eq 'Fail' }).Count
    $WarnCnt  = @($All | Where-Object { $_.Status -eq 'Warning' }).Count
    $NACnt    = @($All | Where-Object { $_.Status -eq 'N/A' }).Count
    $NotCnt   = @($All | Where-Object { $_.Status -eq 'Not Assessed' }).Count
    $AccRiskCnt = @($All | Where-Object { $_.Status -eq 'Accepted Risk' }).Count
    $DeferCnt   = @($All | Where-Object { $_.Status -eq 'Deferred' }).Count
    $Total    = $All.Count

    $DonutSize = 120; $DonutThickness = 18; $Radius = ($DonutSize - $DonutThickness) / 2
    $CX = $DonutSize / 2; $CY = $DonutSize / 2
    $Canvas = New-Object System.Windows.Controls.Canvas
    $Canvas.Width = $DonutSize; $Canvas.Height = $DonutSize

    $Segments = @()
    if ($PassCnt -gt 0)    { $Segments += @{ Count=$PassCnt;    Color=$(if($Global:IsLightMode){'#15803D'}else{'#00C853'}) } }
    if ($FailCnt -gt 0)    { $Segments += @{ Count=$FailCnt;    Color=$(if($Global:IsLightMode){'#DC2626'}else{'#FF5000'}) } }
    if ($WarnCnt -gt 0)    { $Segments += @{ Count=$WarnCnt;    Color=$(if($Global:IsLightMode){'#C2410C'}else{'#F59E0B'}) } }
    if ($AccRiskCnt -gt 0) { $Segments += @{ Count=$AccRiskCnt; Color=$(if($Global:IsLightMode){'#B45309'}else{'#D97706'}) } }
    if ($DeferCnt -gt 0)   { $Segments += @{ Count=$DeferCnt;   Color=$(if($Global:IsLightMode){'#0078D4'}else{'#3B9FE3'}) } }
    if ($NACnt -gt 0)      { $Segments += @{ Count=$NACnt;      Color=$(if($Global:IsLightMode){'#888888'}else{'#555555'}) } }
    if ($NotCnt -gt 0)     { $Segments += @{ Count=$NotCnt;     Color=$(if($Global:IsLightMode){'#D0D0D0'}else{'#2A2A2A'}) } }

    if ($Segments.Count -eq 0 -or $Total -eq 0) {
        $EmptyEllipse = New-Object System.Windows.Shapes.Ellipse
        $EmptyEllipse.Width = $DonutSize; $EmptyEllipse.Height = $DonutSize
        $EmptyEllipse.StrokeThickness = $DonutThickness
        $EmptyEllipse.Stroke = $Global:CachedBC.ConvertFromString($(if($Global:IsLightMode){'#E0E0E0'}else{'#2A2A2A'}))
        $EmptyEllipse.Fill = [System.Windows.Media.Brushes]::Transparent
        [void]$Canvas.Children.Add($EmptyEllipse)
    } else {
        $StartAngle = -90
        foreach ($Seg in $Segments) {
            $SweepAngle = ($Seg.Count / $Total) * 360
            if ($SweepAngle -lt 0.5) { continue }
            $Path = New-Object System.Windows.Shapes.Path
            $Path.StrokeThickness = $DonutThickness
            $Path.Stroke = $Global:CachedBC.ConvertFromString($Seg.Color)
            $Path.Fill = [System.Windows.Media.Brushes]::Transparent
            $Path.StrokeStartLineCap = 'Flat'; $Path.StrokeEndLineCap = 'Flat'
            $StartRad = $StartAngle * [Math]::PI / 180
            $EndRad   = ($StartAngle + $SweepAngle) * [Math]::PI / 180
            $X1 = $CX + $Radius * [Math]::Cos($StartRad); $Y1 = $CY + $Radius * [Math]::Sin($StartRad)
            $X2 = $CX + $Radius * [Math]::Cos($EndRad);   $Y2 = $CY + $Radius * [Math]::Sin($EndRad)
            $Fig = New-Object System.Windows.Media.PathFigure
            $Fig.StartPoint = [System.Windows.Point]::new($X1, $Y1); $Fig.IsClosed = $false
            $Arc = New-Object System.Windows.Media.ArcSegment
            $Arc.Point = [System.Windows.Point]::new($X2, $Y2)
            $Arc.Size = [System.Windows.Size]::new($Radius, $Radius)
            $Arc.IsLargeArc = ($SweepAngle -gt 180); $Arc.SweepDirection = 'Clockwise'
            [void]$Fig.Segments.Add($Arc)
            $Geo = New-Object System.Windows.Media.PathGeometry; [void]$Geo.Figures.Add($Fig)
            $Path.Data = $Geo; [void]$Canvas.Children.Add($Path)
            $StartAngle += $SweepAngle
        }
    }

    # Center label
    $CenterLabel = New-Object System.Windows.Controls.TextBlock
    $CenterLabel.Text = if ($Compliance -ge 0) { "$Compliance%" } else { [string][char]0x2014 }
    $CenterLabel.FontSize = 18; $CenterLabel.FontWeight = [System.Windows.FontWeights]::Bold
    $CenterLabel.HorizontalAlignment = 'Center'; $CenterLabel.VerticalAlignment = 'Center'
    if ($Compliance -ge 80) { $CenterLabel.Foreground = $Window.Resources['ThemeSuccess'] }
    elseif ($Compliance -ge 50) { $CenterLabel.Foreground = $Window.Resources['ThemeWarning'] }
    elseif ($Compliance -ge 0) { $CenterLabel.Foreground = $Window.Resources['ThemeError'] }
    else { $CenterLabel.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeAccent') }
    $CenterGrid = New-Object System.Windows.Controls.Grid
    $CenterGrid.Width = $DonutSize; $CenterGrid.Height = $DonutSize
    [void]$CenterGrid.Children.Add($Canvas); [void]$CenterGrid.Children.Add($CenterLabel)
    [void]$pnlHeroDonut.Children.Add($CenterGrid)

    # Legend items
    foreach ($LI in @(
        @{ Label='Pass'; Count=$PassCnt; Color=$(if($Global:IsLightMode){'#15803D'}else{'#00C853'}) }
        @{ Label='Fail'; Count=$FailCnt; Color=$(if($Global:IsLightMode){'#DC2626'}else{'#FF5000'}) }
        @{ Label='Warning'; Count=$WarnCnt; Color=$(if($Global:IsLightMode){'#C2410C'}else{'#F59E0B'}) }
        @{ Label='Accepted Risk'; Count=$AccRiskCnt; Color=$(if($Global:IsLightMode){'#B45309'}else{'#D97706'}) }
        @{ Label='Deferred'; Count=$DeferCnt; Color=$(if($Global:IsLightMode){'#0078D4'}else{'#3B9FE3'}) }
        @{ Label='N/A'; Count=$NACnt; Color=$(if($Global:IsLightMode){'#888888'}else{'#555555'}) }
        @{ Label='Not Assessed'; Count=$NotCnt; Color=$(if($Global:IsLightMode){'#D0D0D0'}else{'#2A2A2A'}) }
    )) {
        $LRow = New-Object System.Windows.Controls.StackPanel; $LRow.Orientation = 'Horizontal'
        $LRow.Margin = [System.Windows.Thickness]::new(0,1,0,1)
        $Dot = New-Object System.Windows.Shapes.Ellipse
        $Dot.Width = 8; $Dot.Height = 8; $Dot.Fill = $Global:CachedBC.ConvertFromString($LI.Color)
        $Dot.VerticalAlignment = 'Center'; $Dot.Margin = [System.Windows.Thickness]::new(0,0,6,0)
        $LblTB = New-Object System.Windows.Controls.TextBlock
        $LblTB.Text = "$($LI.Label): $($LI.Count)"; $LblTB.FontSize = 11
        $LblTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
        [void]$LRow.Children.Add($Dot); [void]$LRow.Children.Add($LblTB)
        [void]$pnlHeroLegend.Children.Add($LRow)
    }

    # ═══════════════════════════════════════════════════════
    # CATEGORY SCORE CARDS
    # ═══════════════════════════════════════════════════════
    $pnlScoreCards.Children.Clear()
    foreach ($Cat in (Get-Categories)) {
        $Score = Get-CategoryScore $Cat
        $ShortCat = ($Cat -split ' ')[0]
        if ($ShortCat.Length -gt 12) { $ShortCat = $ShortCat.Substring(0,10) + '..' }
        $Card = New-Object System.Windows.Controls.Border
        $Card.CornerRadius = [System.Windows.CornerRadius]::new(8)
        $Card.Padding = [System.Windows.Thickness]::new(12,8,12,8)
        $Card.Margin  = [System.Windows.Thickness]::new(0,0,8,8)
        $Card.MinWidth = 100
        $Card.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeCardBg')
        $Card.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderCard')
        $Card.BorderThickness = [System.Windows.Thickness]::new(1)
        $SP = New-Object System.Windows.Controls.StackPanel
        $CatLabel = New-Object System.Windows.Controls.TextBlock
        $CatLabel.Text = $ShortCat; $CatLabel.FontSize = 10; $CatLabel.ToolTip = $Cat
        $CatLabel.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
        $ScoreLabel = New-Object System.Windows.Controls.TextBlock
        $ScoreLabel.Text = if ($Score -ge 0) { "$Score%" } else { [string][char]0x2014 }
        $ScoreLabel.FontSize = 20; $ScoreLabel.FontWeight = [System.Windows.FontWeights]::Bold
        if ($Score -ge 80) { $ScoreLabel.Foreground = $Window.Resources['ThemeSuccess'] }
        elseif ($Score -ge 50) { $ScoreLabel.Foreground = $Window.Resources['ThemeWarning'] }
        elseif ($Score -ge 0) { $ScoreLabel.Foreground = $Window.Resources['ThemeError'] }
        else { $ScoreLabel.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted') }
        [void]$SP.Children.Add($CatLabel); [void]$SP.Children.Add($ScoreLabel)
        $Card.Child = $SP; [void]$pnlScoreCards.Children.Add($Card)
    }

    # ═══════════════════════════════════════════════════════
    # CATEGORY BREAKDOWN — Full-width segmented bars
    # ═══════════════════════════════════════════════════════
    $pnlCategoryBars.Children.Clear()
    foreach ($Cat in (Get-Categories)) {
        $CatChecks = @($Global:Assessment.Checks | Where-Object { $_.Category -eq $Cat })
        $Pass = @($CatChecks | Where-Object { $_.Status -eq 'Pass' }).Count
        $Fail = @($CatChecks | Where-Object { $_.Status -eq 'Fail' }).Count
        $Warn = @($CatChecks | Where-Object { $_.Status -eq 'Warning' }).Count
        $Ttl  = $CatChecks.Count

        $RowBorder = New-Object System.Windows.Controls.Border
        $RowBorder.CornerRadius = [System.Windows.CornerRadius]::new(6)
        $RowBorder.Padding = [System.Windows.Thickness]::new(12,8,12,8)
        $RowBorder.Margin  = [System.Windows.Thickness]::new(0,0,0,4)
        $RowBorder.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeCardBg')

        $RowSP = New-Object System.Windows.Controls.StackPanel

        $LabelRow = New-Object System.Windows.Controls.DockPanel
        $CatTB = New-Object System.Windows.Controls.TextBlock
        $CatTB.Text = $Cat; $CatTB.FontSize = 12; $CatTB.FontWeight = [System.Windows.FontWeights]::SemiBold
        $CatTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
        $CountTB = New-Object System.Windows.Controls.TextBlock
        $CountTB.Text = "$Pass/$Ttl pass"; $CountTB.FontSize = 11
        $CountTB.HorizontalAlignment = 'Right'
        $CountTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
        [System.Windows.Controls.DockPanel]::SetDock($CountTB, 'Right')
        [void]$LabelRow.Children.Add($CountTB); [void]$LabelRow.Children.Add($CatTB)
        [void]$RowSP.Children.Add($LabelRow)

        # Full-width segmented bar using Grid with star columns
        $BarGrid = New-Object System.Windows.Controls.Grid
        $BarGrid.Height = 8; $BarGrid.Margin = [System.Windows.Thickness]::new(0,6,0,0)

        # Background track (full width)
        $Track = New-Object System.Windows.Controls.Border
        $Track.CornerRadius = [System.Windows.CornerRadius]::new(4)
        $Track.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeDeepBg')
        [void]$BarGrid.Children.Add($Track)

        if ($Ttl -gt 0) {
            # Use a single-row Grid with proportional columns for true full-width
            $SegGrid = New-Object System.Windows.Controls.Grid
            $SegGrid.ClipToBounds = $true

            $SegDefs = @(
                @{ Cnt=$Pass; Color=$(if($Global:IsLightMode){'#15803D'}else{'#00C853'}) }
                @{ Cnt=$Warn; Color=$(if($Global:IsLightMode){'#C2410C'}else{'#F59E0B'}) }
                @{ Cnt=$Fail; Color=$(if($Global:IsLightMode){'#DC2626'}else{'#FF5000'}) }
            )
            $colIdx = 0
            foreach ($SD in $SegDefs) {
                if ($SD.Cnt -le 0) { continue }
                $col = New-Object System.Windows.Controls.ColumnDefinition
                $col.Width = [System.Windows.GridLength]::new($SD.Cnt, 'Star')
                [void]$SegGrid.ColumnDefinitions.Add($col)
                $SegBdr = New-Object System.Windows.Controls.Border
                $SegBdr.Height = 8; $SegBdr.Background = $Global:CachedBC.ConvertFromString($SD.Color)
                [System.Windows.Controls.Grid]::SetColumn($SegBdr, $colIdx)
                [void]$SegGrid.Children.Add($SegBdr)
                $colIdx++
            }
            # Remaining (NA + Not Assessed) as empty space
            $Remaining = $Ttl - $Pass - $Warn - $Fail
            if ($Remaining -gt 0) {
                $col = New-Object System.Windows.Controls.ColumnDefinition
                $col.Width = [System.Windows.GridLength]::new($Remaining, 'Star')
                [void]$SegGrid.ColumnDefinitions.Add($col)
            }

            $ClipBorder = New-Object System.Windows.Controls.Border
            $ClipBorder.CornerRadius = [System.Windows.CornerRadius]::new(4)
            $ClipBorder.ClipToBounds = $true
            $ClipBorder.Child = $SegGrid
            [void]$BarGrid.Children.Add($ClipBorder)
        }

        [void]$RowSP.Children.Add($BarGrid)
        $RowBorder.Child = $RowSP
        [void]$pnlCategoryBars.Children.Add($RowBorder)
    }

    # ═══════════════════════════════════════════════════════
    # MATURITY RADAR (compact, beside category cards)
    # ═══════════════════════════════════════════════════════
    $pnlMaturityAlignment.Children.Clear()

    $DimKeys   = @($Global:MaturityDimensions.Keys)
    $DimCount  = $DimKeys.Count
    $RadarSize = 200
    $RadarCX   = $RadarSize / 2; $RadarCY = $RadarSize / 2
    $RadarR    = ($RadarSize / 2) - 28

    $RadarCard = New-Object System.Windows.Controls.Border
    $RadarCard.CornerRadius = [System.Windows.CornerRadius]::new(8)
    $RadarCard.Padding = [System.Windows.Thickness]::new(12,12,12,12)
    $RadarCard.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeCardBg')
    $RadarCard.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderCard')
    $RadarCard.BorderThickness = [System.Windows.Thickness]::new(1)

    $RadarCanvas = New-Object System.Windows.Controls.Canvas
    $RadarCanvas.Width = $RadarSize; $RadarCanvas.Height = $RadarSize

    # Zone rings
    for ($z = 4; $z -ge 0; $z--) {
        $ZoneR = $RadarR * (($z + 1) / 5)
        $ZonePoly = New-Object System.Windows.Shapes.Polygon
        $ZonePoly.Stroke = $Global:CachedBC.ConvertFromString($(if($Global:IsLightMode){'#15000000'}else{'#15FFFFFF'}))
        $ZonePoly.StrokeThickness = 0.5
        $ZonePoly.Fill = $Global:CachedBC.ConvertFromString($(if($Global:IsLightMode){'#04000000'}else{'#04FFFFFF'}))
        $ZonePoints = New-Object System.Windows.Media.PointCollection
        for ($d = 0; $d -lt $DimCount; $d++) {
            $Angle = (([math]::PI * 2 * $d) / $DimCount) - ([math]::PI / 2)
            [void]$ZonePoints.Add([System.Windows.Point]::new($RadarCX + $ZoneR * [Math]::Cos($Angle), $RadarCY + $ZoneR * [Math]::Sin($Angle)))
        }
        $ZonePoly.Points = $ZonePoints; [void]$RadarCanvas.Children.Add($ZonePoly)
    }

    # Axes
    for ($d = 0; $d -lt $DimCount; $d++) {
        $Angle = (([math]::PI * 2 * $d) / $DimCount) - ([math]::PI / 2)
        $AxisLine = New-Object System.Windows.Shapes.Line
        $AxisLine.X1 = $RadarCX; $AxisLine.Y1 = $RadarCY
        $AxisLine.X2 = $RadarCX + $RadarR * [Math]::Cos($Angle); $AxisLine.Y2 = $RadarCY + $RadarR * [Math]::Sin($Angle)
        $AxisLine.Stroke = $Global:CachedBC.ConvertFromString($(if($Global:IsLightMode){'#20000000'}else{'#20FFFFFF'}))
        $AxisLine.StrokeThickness = 0.5; [void]$RadarCanvas.Children.Add($AxisLine)
    }

    # Data polygon
    $DataPoly = New-Object System.Windows.Shapes.Polygon
    $DataPoly.Stroke = $Global:CachedBC.ConvertFromString($(if($Global:IsLightMode){'#0078D4'}else{'#60CDFF'}))
    $DataPoly.StrokeThickness = 2
    $DataPoly.Fill = $Global:CachedBC.ConvertFromString($(if($Global:IsLightMode){'#200078D4'}else{'#2060CDFF'}))
    $DataPoints = New-Object System.Windows.Media.PointCollection
    $DimScores = @()
    for ($d = 0; $d -lt $DimCount; $d++) {
        $DScore = Get-DimensionScore $DimKeys[$d]; if ($DScore -lt 0) { $DScore = 0 }
        $DimScores += $DScore
        $Angle = (([math]::PI * 2 * $d) / $DimCount) - ([math]::PI / 2)
        $DR = $RadarR * ($DScore / 100)
        [void]$DataPoints.Add([System.Windows.Point]::new($RadarCX + $DR * [Math]::Cos($Angle), $RadarCY + $DR * [Math]::Sin($Angle)))
    }
    $DataPoly.Points = $DataPoints; [void]$RadarCanvas.Children.Add($DataPoly)

    # Dots + labels
    for ($d = 0; $d -lt $DimCount; $d++) {
        $DScore = $DimScores[$d]
        $Angle = (([math]::PI * 2 * $d) / $DimCount) - ([math]::PI / 2)
        $DR = $RadarR * ($DScore / 100)
        $Dot = New-Object System.Windows.Shapes.Ellipse
        $Dot.Width = 6; $Dot.Height = 6
        $Dot.Fill = $Global:CachedBC.ConvertFromString($(if($Global:IsLightMode){'#0078D4'}else{'#60CDFF'}))
        [System.Windows.Controls.Canvas]::SetLeft($Dot, $RadarCX + $DR * [Math]::Cos($Angle) - 3)
        [System.Windows.Controls.Canvas]::SetTop($Dot, $RadarCY + $DR * [Math]::Sin($Angle) - 3)
        [void]$RadarCanvas.Children.Add($Dot)

        # Axis label
        $Dim = $Global:MaturityDimensions[$DimKeys[$d]]
        $LblR = $RadarR + 16
        $AxisLbl = New-Object System.Windows.Controls.TextBlock
        $AxisLbl.Text = ($Dim.Label -split ' ')[0]; $AxisLbl.FontSize = 8
        $AxisLbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
        $AxisLbl.Measure([System.Windows.Size]::new([double]::PositiveInfinity, [double]::PositiveInfinity))
        $LblW = [math]::Max($AxisLbl.DesiredSize.Width, 25)
        [System.Windows.Controls.Canvas]::SetLeft($AxisLbl, $RadarCX + $LblR * [Math]::Cos($Angle) - $LblW/2)
        [System.Windows.Controls.Canvas]::SetTop($AxisLbl, $RadarCY + $LblR * [Math]::Sin($Angle) - 6)
        [void]$RadarCanvas.Children.Add($AxisLbl)
    }

    # Use horizontal Grid: radar left, legend right
    $RadarGrid = New-Object System.Windows.Controls.Grid
    $RCol1 = New-Object System.Windows.Controls.ColumnDefinition; $RCol1.Width = [System.Windows.GridLength]::new($RadarSize)
    $RCol2 = New-Object System.Windows.Controls.ColumnDefinition; $RCol2.Width = [System.Windows.GridLength]::new(1, 'Star')
    [void]$RadarGrid.ColumnDefinitions.Add($RCol1)
    [void]$RadarGrid.ColumnDefinitions.Add($RCol2)

    [System.Windows.Controls.Grid]::SetColumn($RadarCanvas, 0)
    [void]$RadarGrid.Children.Add($RadarCanvas)

    # Maturity legend (right of radar)
    $LegendSP = New-Object System.Windows.Controls.StackPanel
    $LegendSP.VerticalAlignment = 'Center'
    $LegendSP.Margin = [System.Windows.Thickness]::new(16,0,4,0)
    $LegendSP.MinWidth = 200
    for ($d = 0; $d -lt $DimCount; $d++) {
        $Dim = $Global:MaturityDimensions[$DimKeys[$d]]
        $DScore = $DimScores[$d]; $Level = Get-MaturityLevel $DScore
        $MRow = New-Object System.Windows.Controls.DockPanel; $MRow.Margin = [System.Windows.Thickness]::new(0,2,0,2)
        $MLbl = New-Object System.Windows.Controls.TextBlock; $MLbl.Text = $Dim.Label; $MLbl.FontSize = 11
        $MLbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
        $MVal = New-Object System.Windows.Controls.TextBlock; $MVal.Text = "$DScore%"; $MVal.FontSize = 11
        $MVal.HorizontalAlignment = 'Right'; $MVal.Margin = [System.Windows.Thickness]::new(12,0,0,0)
        [System.Windows.Controls.DockPanel]::SetDock($MVal, 'Right')
        if ($DScore -ge 75) { $MVal.Foreground = $Window.Resources['ThemeSuccess'] }
        elseif ($DScore -ge 35) { $MVal.Foreground = $Window.Resources['ThemeWarning'] }
        else { $MVal.Foreground = $Window.Resources['ThemeError'] }
        [void]$MRow.Children.Add($MVal); [void]$MRow.Children.Add($MLbl)
        [void]$LegendSP.Children.Add($MRow)
    }
    [System.Windows.Controls.Grid]::SetColumn($LegendSP, 1)
    [void]$RadarGrid.Children.Add($LegendSP)

    $RadarCard.Child = $RadarGrid
    [void]$pnlMaturityAlignment.Children.Add($RadarCard)

    # ═══════════════════════════════════════════════════════
    # TOP FINDINGS PREVIEW
    # ═══════════════════════════════════════════════════════
    $pnlTopFindings.Children.Clear()
    $SevOrder = @{ Critical=0; High=1; Medium=2; Low=3 }
    $TopFindings = @($Global:Assessment.Checks | Where-Object { $_.Status -eq 'Fail' } |
        Sort-Object { $SevOrder[$_.Severity] }, { -[int]$_.Weight } | Select-Object -First 5)

    if ($TopFindings.Count -eq 0) {
        $EmptyTB = New-Object System.Windows.Controls.TextBlock
        $EmptyTB.Text = "No failed checks. Import a collection JSON to begin assessment."
        $EmptyTB.FontSize = 12; $EmptyTB.FontStyle = 'Italic'
        $EmptyTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
        [void]$pnlTopFindings.Children.Add($EmptyTB)
    }

    foreach ($F in $TopFindings) {
        $FRow = New-Object System.Windows.Controls.Border
        $FRow.CornerRadius = [System.Windows.CornerRadius]::new(6)
        $FRow.Padding = [System.Windows.Thickness]::new(12,6,12,6)
        $FRow.Margin  = [System.Windows.Thickness]::new(0,0,0,3)
        $FRow.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeCardBg')
        $FDP = New-Object System.Windows.Controls.DockPanel
        $SDot = New-Object System.Windows.Shapes.Ellipse
        $SDot.Width = 8; $SDot.Height = 8; $SDot.VerticalAlignment = 'Center'
        $SDot.Margin = [System.Windows.Thickness]::new(0,0,8,0)
        switch ($F.Severity) {
            'Critical' { $SDot.Fill = $Window.Resources['ThemeError'] }
            'High'     { $SDot.Fill = $Window.Resources['ThemeWarning'] }
            default    { $SDot.Fill = $Window.Resources['ThemeAccent'] }
        }
        $FName = New-Object System.Windows.Controls.TextBlock
        $FName.Text = "$($F.Id) $($F.Name)"; $FName.FontSize = 12
        $FName.TextTrimming = 'CharacterEllipsis'; $FName.VerticalAlignment = 'Center'
        $FName.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
        $FSev = New-Object System.Windows.Controls.TextBlock
        $FSev.Text = $F.Severity; $FSev.FontSize = 10; $FSev.VerticalAlignment = 'Center'
        $FSev.Margin = [System.Windows.Thickness]::new(8,0,0,0)
        [System.Windows.Controls.DockPanel]::SetDock($FSev, 'Right')
        switch ($F.Severity) {
            'Critical' { $FSev.Foreground = $Window.Resources['ThemeError'] }
            'High'     { $FSev.Foreground = $Window.Resources['ThemeWarning'] }
            default    { $FSev.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted') }
        }
        [void]$FDP.Children.Add($FSev); [void]$FDP.Children.Add($SDot); [void]$FDP.Children.Add($FName)
        $FRow.Child = $FDP; [void]$pnlTopFindings.Children.Add($FRow)
    }

    Update-Progress
}


function Render-BaselineChecks {
    # Keep header elements (first 2 children = title + subtitle)
    while ($pnlBaselineChecks.Children.Count -gt 2) {
        $pnlBaselineChecks.Children.RemoveAt($pnlBaselineChecks.Children.Count - 1)
    }

    # Show loading indicator
    $LoadingTB = New-Object System.Windows.Controls.TextBlock
    $LoadingTB.Text = "Loading checks..."
    $LoadingTB.FontSize = 13; $LoadingTB.FontStyle = 'Italic'
    $LoadingTB.Margin = [System.Windows.Thickness]::new(0,8,0,8)
    $LoadingTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
    [void]$pnlBaselineChecks.Children.Add($LoadingTB)

    $PriorityFilter = $null
    if ($cmbBaselinePriority -and $cmbBaselinePriority.SelectedItem) {
        $Sel = $cmbBaselinePriority.SelectedItem.Content
        if ($Sel -ne 'All') { $PriorityFilter = $Sel }
    }
    $DeviationsOnly = $chkShowDeviationsOnly -and $chkShowDeviationsOnly.IsChecked

    # Build category work queue
    $CatQueue = [System.Collections.Queue]::new()
    foreach ($Cat in (Get-Categories)) {
        $CatChecks = @($Global:Assessment.Checks | Where-Object { $_.Category -eq $Cat })
        if ($PriorityFilter) { $CatChecks = @($CatChecks | Where-Object { $_.Priority -eq $PriorityFilter }) }
        if ($DeviationsOnly) { $CatChecks = @($CatChecks | Where-Object { $_.Status -in @('Fail','Warning','Deferred','Accepted Risk') }) }
        if ($CatChecks.Count -eq 0) { continue }
        $CatQueue.Enqueue(@{ Cat = $Cat; Checks = $CatChecks })
    }

    # Remove loading indicator once first batch starts
    $LoadingRef = $LoadingTB
    $PanelRef = $pnlBaselineChecks
    $QueueRef = $CatQueue

    # Render categories one at a time via timer for responsive UI
    $RenderTimer = New-Object System.Windows.Threading.DispatcherTimer
    $RenderTimer.Interval = [TimeSpan]::FromMilliseconds(15)
    $TimerRef = $RenderTimer
    $FirstBatch = $true

    $RenderTimer.Add_Tick({
        if ($QueueRef.Count -eq 0) {
            $TimerRef.Stop()
            Update-Progress
            return
        }

        # Remove loading indicator on first batch
        if ($FirstBatch) {
            $PanelRef.Children.Remove($LoadingRef)
            $Script:FirstBatch = $false
        }

        $Item = $QueueRef.Dequeue()
        Render-CategoryBlock -Panel $PanelRef -Category $Item.Cat -Checks $Item.Checks
    }.GetNewClosure())

    $RenderTimer.Start()
}

# Renders a single category block (header + check cards) into the parent panel
function Render-CategoryBlock {
    param($Panel, [string]$Category, $Checks)

    $PassCnt = @($Checks | Where-Object { $_.Status -eq 'Pass' }).Count
    $FailCnt = @($Checks | Where-Object { $_.Status -eq 'Fail' }).Count
    $WarnCnt = @($Checks | Where-Object { $_.Status -eq 'Warning' }).Count
    $AccRCnt = @($Checks | Where-Object { $_.Status -eq 'Accepted Risk' }).Count
    $DefCnt  = @($Checks | Where-Object { $_.Status -eq 'Deferred' }).Count

    # Start collapsed if >15 checks to further speed rendering
    $StartCollapsed = ($Checks.Count -gt 15)

        # ── Collapsible category header ──
        $CatBorder = New-Object System.Windows.Controls.Border
        $CatBorder.CornerRadius = [System.Windows.CornerRadius]::new(8)
        $CatBorder.Padding = [System.Windows.Thickness]::new(12,8,12,8)
        $CatBorder.Margin  = [System.Windows.Thickness]::new(0,12,0,4)
        $CatBorder.Cursor = [System.Windows.Input.Cursors]::Hand
        $CatBorder.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeSurfaceBg')
        $CatBorder.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderCard')
        $CatBorder.BorderThickness = [System.Windows.Thickness]::new(1)

        $CatDP = New-Object System.Windows.Controls.DockPanel

        $Chevron = New-Object System.Windows.Controls.TextBlock
        $Chevron.Text = [char]0xE70D
        $Chevron.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
        $Chevron.FontSize = 10; $Chevron.VerticalAlignment = 'Center'
        $Chevron.Margin = [System.Windows.Thickness]::new(0,0,8,0)
        $Chevron.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')

        $CatTB = New-Object System.Windows.Controls.TextBlock
        $CatTB.Text = "$Category ($($Checks.Count))"
        $CatTB.FontSize = 14; $CatTB.FontWeight = [System.Windows.FontWeights]::SemiBold
        $CatTB.VerticalAlignment = 'Center'
        $CatTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextPrimary')

        # Status counts
        $CountsSP = New-Object System.Windows.Controls.StackPanel
        $CountsSP.Orientation = 'Horizontal'; $CountsSP.VerticalAlignment = 'Center'
        [System.Windows.Controls.DockPanel]::SetDock($CountsSP, 'Right')

        foreach ($sc in @(
            @{ Cnt=$PassCnt; Lbl='pass'; Res='ThemeSuccess' }
            @{ Cnt=$FailCnt; Lbl='fail'; Res='ThemeError' }
            @{ Cnt=$WarnCnt; Lbl='warn'; Res='ThemeWarning' }
            @{ Cnt=$AccRCnt; Lbl='accepted'; Res='ThemeWarning' }
            @{ Cnt=$DefCnt;  Lbl='deferred'; Res='ThemeAccent' }
        )) {
            if ($sc.Cnt -gt 0) {
                $cntTB = New-Object System.Windows.Controls.TextBlock
                $cntTB.Text = "$($sc.Cnt)"; $cntTB.FontSize = 11; $cntTB.FontWeight = [System.Windows.FontWeights]::SemiBold
                $cntTB.Foreground = $Window.Resources[$sc.Res]; $cntTB.Margin = [System.Windows.Thickness]::new(0,0,3,0)
                $lblTB = New-Object System.Windows.Controls.TextBlock
                $lblTB.Text = $sc.Lbl; $lblTB.FontSize = 11; $lblTB.Margin = [System.Windows.Thickness]::new(0,0,10,0)
                $lblTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
                [void]$CountsSP.Children.Add($cntTB); [void]$CountsSP.Children.Add($lblTB)
            }
        }

        [void]$CatDP.Children.Add($CountsSP)
        [void]$CatDP.Children.Add($Chevron)
        [void]$CatDP.Children.Add($CatTB)
        $CatBorder.Child = $CatDP

        $ChecksPanel = New-Object System.Windows.Controls.StackPanel
        $ChecksPanel.Visibility = 'Visible'

        $ChevronRef = $Chevron; $PanelRef = $ChecksPanel
        $CatBorder.Add_MouseLeftButtonDown({
            if ($PanelRef.Visibility -eq 'Visible') {
                $PanelRef.Visibility = 'Collapsed'
                $ChevronRef.Text = [char]0xE76C
            } else {
                $PanelRef.Visibility = 'Visible'
                $ChevronRef.Text = [char]0xE70D
            }
        }.GetNewClosure())

        [void]$Panel.Children.Add($CatBorder)

        foreach ($Chk in $Checks) {
            $Card = New-Object System.Windows.Controls.Border
            $Card.CornerRadius = [System.Windows.CornerRadius]::new(8)
            $Card.Padding = [System.Windows.Thickness]::new(16,10,16,10)
            $Card.Margin  = [System.Windows.Thickness]::new(0,0,0,4)
            $Card.BorderThickness = [System.Windows.Thickness]::new(1)

            # Left accent stripe based on status
            $AccentColor = switch ($Chk.Status) {
                'Pass'          { if ($Global:IsLightMode) { '#15803D' } else { '#00C853' } }
                'Fail'          { if ($Global:IsLightMode) { '#DC2626' } else { '#FF5000' } }
                'Warning'       { if ($Global:IsLightMode) { '#C2410C' } else { '#F59E0B' } }
                'Accepted Risk' { if ($Global:IsLightMode) { '#B45309' } else { '#D97706' } }
                'Deferred'      { if ($Global:IsLightMode) { '#0078D4' } else { '#3B9FE3' } }
                default         { if ($Global:IsLightMode) { '#D0D0D0' } else { '#333333' } }
            }

            # Subtle status-tinted card background & border
            $CardBg = switch ($Chk.Status) {
                'Pass'          { if ($Global:IsLightMode) { '#0A15803D' } else { '#0A00C853' } }
                'Fail'          { if ($Global:IsLightMode) { '#0ADC2626' } else { '#0AFF5000' } }
                'Warning'       { if ($Global:IsLightMode) { '#08C2410C' } else { '#08F59E0B' } }
                'Accepted Risk' { if ($Global:IsLightMode) { '#08B45309' } else { '#08D97706' } }
                'Deferred'      { if ($Global:IsLightMode) { '#080078D4' } else { '#083B9FE3' } }
                default         { $null }
            }
            $CardBorder = switch ($Chk.Status) {
                'Pass'          { if ($Global:IsLightMode) { '#2015803D' } else { '#1800C853' } }
                'Fail'          { if ($Global:IsLightMode) { '#20DC2626' } else { '#18FF5000' } }
                'Warning'       { if ($Global:IsLightMode) { '#18C2410C' } else { '#14F59E0B' } }
                'Accepted Risk' { if ($Global:IsLightMode) { '#18B45309' } else { '#14D97706' } }
                'Deferred'      { if ($Global:IsLightMode) { '#180078D4' } else { '#143B9FE3' } }
                default         { $null }
            }
            if ($CardBg) {
                $Card.Background  = $Global:CachedBC.ConvertFromString($CardBg)
                $Card.BorderBrush = $Global:CachedBC.ConvertFromString($CardBorder)
            } else {
                $Card.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeCardBg')
                $Card.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderCard')
            }

            $OuterGrid = New-Object System.Windows.Controls.Grid
            $StripeCol = New-Object System.Windows.Controls.ColumnDefinition; $StripeCol.Width = [System.Windows.GridLength]::new(4)
            $ContentCol = New-Object System.Windows.Controls.ColumnDefinition; $ContentCol.Width = [System.Windows.GridLength]::new(1, 'Star')
            [void]$OuterGrid.ColumnDefinitions.Add($StripeCol)
            [void]$OuterGrid.ColumnDefinitions.Add($ContentCol)

            $Stripe = New-Object System.Windows.Controls.Border
            $Stripe.CornerRadius = [System.Windows.CornerRadius]::new(2)
            $Stripe.Background = $Global:CachedBC.ConvertFromString($AccentColor)
            $Stripe.Margin = [System.Windows.Thickness]::new(0,2,8,2)
            [System.Windows.Controls.Grid]::SetColumn($Stripe, 0)
            [void]$OuterGrid.Children.Add($Stripe)

            $SP = New-Object System.Windows.Controls.StackPanel
            [System.Windows.Controls.Grid]::SetColumn($SP, 1)

            # ── Row 1: Status icon + Name + Severity + Status dropdown ──
            $Row1 = New-Object System.Windows.Controls.DockPanel

            # Status badge (read-only — overrides are done via action buttons)
            $StatusBadge = New-Object System.Windows.Controls.Border
            $StatusBadge.CornerRadius = [System.Windows.CornerRadius]::new(4)
            $StatusBadge.Padding = [System.Windows.Thickness]::new(6,2,6,2)
            $StatusBadge.Margin = [System.Windows.Thickness]::new(8,0,0,0)
            $StatusBadge.BorderThickness = [System.Windows.Thickness]::new(1)
            $StatusBadgeTB = New-Object System.Windows.Controls.TextBlock
            $StatusBadgeTB.FontSize = 10; $StatusBadgeTB.FontWeight = [System.Windows.FontWeights]::SemiBold
            $StatusBadgeTB.VerticalAlignment = 'Center'
            switch ($Chk.Status) {
                'Pass'          { $StatusBadgeTB.Text = 'PASS'; $StatusBadge.Background = $Global:CachedBC.ConvertFromString('#2015803D')
                                  $StatusBadgeTB.Foreground = $Window.Resources['ThemeSuccess']; $StatusBadge.BorderBrush = $Global:CachedBC.ConvertFromString('#4015803D') }
                'Fail'          { $StatusBadgeTB.Text = 'FAIL'; $StatusBadge.Background = $Global:CachedBC.ConvertFromString('#20DC2626')
                                  $StatusBadgeTB.Foreground = $Window.Resources['ThemeError']; $StatusBadge.BorderBrush = $Global:CachedBC.ConvertFromString('#40DC2626') }
                'Warning'       { $StatusBadgeTB.Text = 'WARN'; $StatusBadge.Background = $Global:CachedBC.ConvertFromString('#20F59E0B')
                                  $StatusBadgeTB.Foreground = $Window.Resources['ThemeWarning']; $StatusBadge.BorderBrush = $Global:CachedBC.ConvertFromString('#40F59E0B') }
                'N/A'           { $StatusBadgeTB.Text = 'N/A'; $StatusBadge.Background = $Global:CachedBC.ConvertFromString('#20888888')
                                  $StatusBadgeTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
                                  $StatusBadge.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderSubtle') }
                'Accepted Risk' { $StatusBadgeTB.Text = 'RISK ACCEPTED'; $StatusBadge.Background = $Global:CachedBC.ConvertFromString('#20F59E0B')
                                  $StatusBadgeTB.Foreground = $Window.Resources['ThemeWarning']; $StatusBadge.BorderBrush = $Global:CachedBC.ConvertFromString('#40F59E0B') }
                'Deferred'      { $StatusBadgeTB.Text = 'DEFERRED'; $StatusBadge.Background = $Global:CachedBC.ConvertFromString('#200078D4')
                                  $StatusBadgeTB.Foreground = $Window.Resources['ThemeAccent']; $StatusBadge.BorderBrush = $Global:CachedBC.ConvertFromString('#400078D4') }
                default         { $StatusBadgeTB.Text = 'NOT ASSESSED'; $StatusBadge.Background = $Global:CachedBC.ConvertFromString('#20888888')
                                  $StatusBadgeTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
                                  $StatusBadge.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderSubtle') }
            }
            $StatusBadge.Child = $StatusBadgeTB
            [System.Windows.Controls.DockPanel]::SetDock($StatusBadge, 'Right')

            # Severity badge
            $SevBadge = New-Object System.Windows.Controls.Border
            $SevBadge.CornerRadius = [System.Windows.CornerRadius]::new(4)
            $SevBadge.Padding = [System.Windows.Thickness]::new(6,2,6,2)
            $SevBadge.Margin = [System.Windows.Thickness]::new(8,0,0,0)
            [System.Windows.Controls.DockPanel]::SetDock($SevBadge, 'Right')
            $SevTB = New-Object System.Windows.Controls.TextBlock
            $SevTB.Text = $Chk.Severity; $SevTB.FontSize = 10; $SevTB.FontWeight = [System.Windows.FontWeights]::SemiBold
            switch ($Chk.Severity) {
                'Critical' { $SevBadge.Background = $Global:CachedBC.ConvertFromString('#30FF5000'); $SevTB.Foreground = $Window.Resources['ThemeError'] }
                'High'     { $SevBadge.Background = $Global:CachedBC.ConvertFromString('#30F59E0B'); $SevTB.Foreground = $Window.Resources['ThemeWarning'] }
                'Medium'   { $SevBadge.Background = $Global:CachedBC.ConvertFromString('#300078D4'); $SevTB.Foreground = $Window.Resources['ThemeAccent'] }
                default    { $SevBadge.Background = $Global:CachedBC.ConvertFromString('#30888888'); $SevTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted') }
            }
            $SevBadge.Child = $SevTB

            # Status icon
            $StatusIcon = New-Object System.Windows.Controls.TextBlock
            $StatusIcon.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
            $StatusIcon.FontSize = 14; $StatusIcon.VerticalAlignment = 'Center'
            $StatusIcon.Margin = [System.Windows.Thickness]::new(0,0,8,0)
            switch ($Chk.Status) {
                'Pass'          { $StatusIcon.Text = [char]0xE73E; $StatusIcon.Foreground = $Window.Resources['ThemeSuccess'] }
                'Fail'          { $StatusIcon.Text = [char]0xEA39; $StatusIcon.Foreground = $Window.Resources['ThemeError'] }
                'Warning'       { $StatusIcon.Text = [char]0xE7BA; $StatusIcon.Foreground = $Window.Resources['ThemeWarning'] }
                'N/A'           { $StatusIcon.Text = [char]0xE711; $StatusIcon.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted') }
                'Accepted Risk' { $StatusIcon.Text = [char]0xE8FB; $StatusIcon.Foreground = $Window.Resources['ThemeWarning'] }
                'Deferred'      { $StatusIcon.Text = [char]0xE823; $StatusIcon.Foreground = $Window.Resources['ThemeAccent'] }
                default         { $StatusIcon.Text = [char]0xE9CE; $StatusIcon.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted') }
            }

            $NameTB = New-Object System.Windows.Controls.TextBlock
            $NameTB.Text = "$($Chk.Id)  $($Chk.Name)"
            $NameTB.FontSize = 13; $NameTB.FontWeight = [System.Windows.FontWeights]::SemiBold
            $NameTB.TextTrimming = 'CharacterEllipsis'; $NameTB.VerticalAlignment = 'Center'
            $NameTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')

            [void]$Row1.Children.Add($StatusBadge)
            # Manual review badge for Not Assessed checks
            if ($Chk.Status -eq 'Not Assessed') {
                $ManualBadge = New-Object System.Windows.Controls.Border
                $ManualBadge.CornerRadius = [System.Windows.CornerRadius]::new(4)
                $ManualBadge.Padding = [System.Windows.Thickness]::new(5,1,5,1)
                $ManualBadge.Margin = [System.Windows.Thickness]::new(0,0,4,0)
                $ManualBadge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeSurfaceBg')
                $ManualBadge.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderSubtle')
                $ManualBadge.BorderThickness = [System.Windows.Thickness]::new(1)
                $ManualTB = New-Object System.Windows.Controls.TextBlock
                $ManualTB.Text = 'MANUAL'; $ManualTB.FontSize = 8; $ManualTB.FontWeight = [System.Windows.FontWeights]::Bold
                $ManualTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextFaintest')
                $ManualBadge.Child = $ManualTB
                [System.Windows.Controls.DockPanel]::SetDock($ManualBadge, 'Right')
                [void]$Row1.Children.Add($ManualBadge)
            }
            [void]$Row1.Children.Add($SevBadge)

            # Origin badge(s) — show recommendation source (SCT, INTUNE, OPS)
            if ($Chk.Origin) {
                foreach ($o in ($Chk.Origin -split ',')) {
                    $oBadge = New-Object System.Windows.Controls.Border
                    $oBadge.CornerRadius = [System.Windows.CornerRadius]::new(4)
                    $oBadge.Padding = [System.Windows.Thickness]::new(5,1,5,1)
                    $oBadge.Margin = [System.Windows.Thickness]::new(4,0,0,0)
                    $oBadge.BorderThickness = [System.Windows.Thickness]::new(1)
                    $oTB = New-Object System.Windows.Controls.TextBlock
                    $oTB.Text = $o.Trim().ToUpper(); $oTB.FontSize = 9; $oTB.FontWeight = [System.Windows.FontWeights]::SemiBold
                    switch ($o.Trim().ToUpper()) {
                        'SCT'    { $oBadge.Background = $Global:CachedBC.ConvertFromString('#200078D4'); $oTB.Foreground = $Window.Resources['ThemeAccent']
                                   $oBadge.BorderBrush = $Global:CachedBC.ConvertFromString('#400078D4') }
                        'INTUNE' { $oBadge.Background = $Global:CachedBC.ConvertFromString('#20008272'); $oTB.Foreground = $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#008272' } else { '#00C9A7' }))
                                   $oBadge.BorderBrush = $Global:CachedBC.ConvertFromString('#40008272') }
                        'OPS'    { $oBadge.Background = $Global:CachedBC.ConvertFromString('#20888888')
                                   $oTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
                                   $oBadge.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderSubtle') }
                        default  { $oBadge.Background = $Global:CachedBC.ConvertFromString('#20888888')
                                   $oTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
                                   $oBadge.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderSubtle') }
                    }
                    $oBadge.Child = $oTB
                    [System.Windows.Controls.DockPanel]::SetDock($oBadge, 'Right')
                    [void]$Row1.Children.Add($oBadge)
                }
            }

            [void]$Row1.Children.Add($StatusIcon)
            [void]$Row1.Children.Add($NameTB)
            [void]$SP.Children.Add($Row1)

            # Card element refs for action-button visual updates
            $ChkRef = $Chk; $IconRef = $StatusIcon; $StripeRef = $Stripe; $CardRef = $Card
            $BadgeRef = $StatusBadge; $BadgeTBRef = $StatusBadgeTB

            # ── Row 2: Expected vs Actual ──
            if ($null -ne $Chk.BaselineValue -or $null -ne $Chk.ActualValue) {
                $Row2 = New-Object System.Windows.Controls.WrapPanel
                $Row2.Margin = [System.Windows.Thickness]::new(22,4,0,0)
                if ($null -ne $Chk.BaselineValue) {
                    $ExpTB = New-Object System.Windows.Controls.TextBlock
                    $ExpTB.Text = "Expected: $($Chk.BaselineValue)"; $ExpTB.FontSize = 11
                    $ExpTB.Margin = [System.Windows.Thickness]::new(0,0,16,0)
                    $ExpTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
                    [void]$Row2.Children.Add($ExpTB)
                }
                if ($null -ne $Chk.ActualValue) {
                    $ActTB = New-Object System.Windows.Controls.TextBlock
                    $ActTB.Text = "Actual: $($Chk.ActualValue)"; $ActTB.FontSize = 11
                    if ($Chk.Status -eq 'Fail') { $ActTB.Foreground = $Window.Resources['ThemeError'] }
                    elseif ($Chk.Status -eq 'Pass') { $ActTB.Foreground = $Window.Resources['ThemeSuccess'] }
                    else { $ActTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextSecondary') }
                    [void]$Row2.Children.Add($ActTB)
                }
                [void]$SP.Children.Add($Row2)
            }

            # ── Row 3: Description ──
            $DescTB = New-Object System.Windows.Controls.TextBlock
            $DescTB.Text = $Chk.Description; $DescTB.FontSize = 11; $DescTB.TextWrapping = 'Wrap'
            $DescTB.Margin = [System.Windows.Thickness]::new(22,4,0,0)
            $DescTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextSecondary')
            [void]$SP.Children.Add($DescTB)

            # ── Row 4: Action buttons ── show for Fail/Warning/Not Assessed/Accepted Risk/Deferred
            if ($Chk.Status -in @('Fail','Warning','Not Assessed','Accepted Risk','Deferred')) {
                $ActionRow = New-Object System.Windows.Controls.WrapPanel
                $ActionRow.Margin = [System.Windows.Thickness]::new(22,8,0,4)

                # Decision label
                $DecLbl = New-Object System.Windows.Controls.TextBlock
                $DecLbl.Text = "Action:"; $DecLbl.FontSize = 11; $DecLbl.VerticalAlignment = 'Center'
                $DecLbl.Margin = [System.Windows.Thickness]::new(0,0,8,0)
                $DecLbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
                [void]$ActionRow.Children.Add($DecLbl)

                # Show Undo button if already overridden (Accepted Risk / Deferred / N/A set by user)
                $IsOverridden = ($Chk.Status -in @('Accepted Risk','Deferred'))
                if ($IsOverridden -and $Chk.AutoStatus) {
                    $UndoBtn = New-Object System.Windows.Controls.Border
                    $UndoBtn.CornerRadius = [System.Windows.CornerRadius]::new(4)
                    $UndoBtn.Padding = [System.Windows.Thickness]::new(8,3,8,3)
                    $UndoBtn.Margin = [System.Windows.Thickness]::new(0,0,8,0)
                    $UndoBtn.Cursor = [System.Windows.Input.Cursors]::Hand
                    $UndoBtn.BorderThickness = [System.Windows.Thickness]::new(1)
                    $UndoBtn.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeSurfaceBg')
                    $UndoBtn.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderSubtle')
                    $UndoTB = New-Object System.Windows.Controls.TextBlock
                    $UndoTB.Text = [char]0xE7A7 + " Undo"; $UndoTB.FontSize = 10
                    $UndoTB.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets, Segoe UI")
                    $UndoTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeAccent')
                    $UndoBtn.Child = $UndoTB
                    $UndoChkRef = $ChkRef
                    $UndoBtn.Add_MouseLeftButtonDown({
                        $UndoChkRef.Status   = $UndoChkRef.AutoStatus
                        $UndoChkRef.Details   = ''
                        $UndoChkRef.Excluded = $false
                        Set-Dirty; Update-Dashboard; Render-BaselineChecks
                    }.GetNewClosure())
                    [void]$ActionRow.Children.Add($UndoBtn)
                }

                # Decision buttons: Remediate | Accept Risk | N/A | Defer
                $Decisions = @(
                    @{ Text='Remediate';   Color='ThemeSuccess';   Tag='Remediate';  StatusVal=$null }
                    @{ Text='Accept Risk'; Color='ThemeWarning';   Tag='AcceptRisk'; StatusVal='Accepted Risk' }
                    @{ Text='N/A';         Color='ThemeTextMuted'; Tag='NA';         StatusVal='N/A' }
                    @{ Text='Defer';       Color='ThemeAccent';    Tag='Defer';      StatusVal='Deferred' }
                )
                $AllDecBtns = @()
                foreach ($Dec in $Decisions) {
                    $DecBtn = New-Object System.Windows.Controls.Border
                    $DecBtn.CornerRadius = [System.Windows.CornerRadius]::new(4)
                    $DecBtn.Padding = [System.Windows.Thickness]::new(8,3,8,3)
                    $DecBtn.Margin = [System.Windows.Thickness]::new(0,0,4,0)
                    $DecBtn.Cursor = [System.Windows.Input.Cursors]::Hand
                    $DecBtn.BorderThickness = [System.Windows.Thickness]::new(1)
                    $DecBtn.Tag = $Dec.Tag

                    $DecTB = New-Object System.Windows.Controls.TextBlock
                    $DecTB.Text = $Dec.Text; $DecTB.FontSize = 10; $DecTB.FontWeight = [System.Windows.FontWeights]::SemiBold
                    $DecTB.VerticalAlignment = 'Center'
                    $DecBtn.Child = $DecTB

                    # Highlight if this decision is already selected
                    $CurrentDec = $Chk.Details
                    if ($CurrentDec -eq $Dec.Tag) {
                        $DecBtn.Background = $Window.Resources[$Dec.Color]
                        $DecTB.Foreground = [System.Windows.Media.Brushes]::White
                        $DecBtn.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, $Dec.Color)
                    } else {
                        $DecBtn.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeSurfaceBg')
                        $DecBtn.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderSubtle')
                        $DecTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $Dec.Color)
                    }

                    $AllDecBtns += @{ Btn = $DecBtn; TB = $DecTB; Tag = $Dec.Tag; ColorKey = $Dec.Color; StatusVal = $Dec.StatusVal }
                    [void]$ActionRow.Children.Add($DecBtn)
                }

                # Click handlers — update status, card visuals, and button styles
                $DecChkRef = $ChkRef; $DecIconRef = $IconRef; $DecStripeRef = $StripeRef
                $DecCardRef = $CardRef; $DecBadgeRef = $BadgeRef; $DecBadgeTBRef = $BadgeTBRef
                $BtnsRef = $AllDecBtns
                foreach ($BtnInfo in $AllDecBtns) {
                    $ClickTag  = $BtnInfo.Tag
                    $ClickStatusVal = $BtnInfo.StatusVal
                    $BtnInfo.Btn.Add_MouseLeftButtonDown({
                        $DecChkRef.Details = $ClickTag
                        # Status-changing actions (AcceptRisk, NA, Defer) update the check status
                        if ($ClickStatusVal) {
                            $DecChkRef.Status   = $ClickStatusVal
                            $DecChkRef.Excluded = ($ClickTag -eq 'AcceptRisk')
                            # Update status badge text + colors
                            switch ($ClickStatusVal) {
                                'Accepted Risk' {
                                    $DecBadgeTBRef.Text = 'RISK ACCEPTED'
                                    $DecBadgeRef.Background = $Global:CachedBC.ConvertFromString('#20F59E0B')
                                    $DecBadgeTBRef.Foreground = $Window.Resources['ThemeWarning']
                                    $DecBadgeRef.BorderBrush = $Global:CachedBC.ConvertFromString('#40F59E0B')
                                }
                                'N/A' {
                                    $DecBadgeTBRef.Text = 'N/A'
                                    $DecBadgeRef.Background = $Global:CachedBC.ConvertFromString('#20888888')
                                    $DecBadgeTBRef.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
                                    $DecBadgeRef.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderSubtle')
                                }
                                'Deferred' {
                                    $DecBadgeTBRef.Text = 'DEFERRED'
                                    $DecBadgeRef.Background = $Global:CachedBC.ConvertFromString('#200078D4')
                                    $DecBadgeTBRef.Foreground = $Window.Resources['ThemeAccent']
                                    $DecBadgeRef.BorderBrush = $Global:CachedBC.ConvertFromString('#400078D4')
                                }
                            }
                            # Update icon
                            switch ($ClickStatusVal) {
                                'Accepted Risk' { $DecIconRef.Text = [char]0xE8FB; $DecIconRef.Foreground = $Window.Resources['ThemeWarning'] }
                                'N/A'           { $DecIconRef.Text = [char]0xE711; $DecIconRef.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted') }
                                'Deferred'      { $DecIconRef.Text = [char]0xE823; $DecIconRef.Foreground = $Window.Resources['ThemeAccent'] }
                            }
                            # Update stripe
                            $sc = switch ($ClickStatusVal) {
                                'Accepted Risk' { if ($Global:IsLightMode) { '#C2410C' } else { '#F59E0B' } }
                                'N/A'           { if ($Global:IsLightMode) { '#888888' } else { '#555555' } }
                                'Deferred'      { if ($Global:IsLightMode) { '#0078D4' } else { '#0078D4' } }
                                default         { if ($Global:IsLightMode) { '#D0D0D0' } else { '#333333' } }
                            }
                            $DecStripeRef.Background = $Global:CachedBC.ConvertFromString($sc)
                            # Update card tint
                            $bg = switch ($ClickStatusVal) {
                                'Accepted Risk' { if ($Global:IsLightMode) { '#08C2410C' } else { '#08F59E0B' } }
                                'N/A'           { $null }
                                'Deferred'      { if ($Global:IsLightMode) { '#080078D4' } else { '#080078D4' } }
                                default         { $null }
                            }
                            if ($bg) {
                                $DecCardRef.Background  = $Global:CachedBC.ConvertFromString($bg)
                                $DecCardRef.BorderBrush = $Global:CachedBC.ConvertFromString($bg.Replace('#08','#18'))
                            } else {
                                $DecCardRef.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeCardBg')
                                $DecCardRef.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderCard')
                            }
                        } else {
                            # Remediate: keep status as-is, just mark decision
                            $DecChkRef.Excluded = $false
                        }
                        Set-Dirty
                        # Update all button styles in-place
                        foreach ($b in $BtnsRef) {
                            if ($b.Tag -eq $ClickTag) {
                                $b.Btn.Background = $Window.Resources[$b.ColorKey]
                                $b.TB.Foreground = [System.Windows.Media.Brushes]::White
                                $b.Btn.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, $b.ColorKey)
                            } else {
                                $b.Btn.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeSurfaceBg')
                                $b.Btn.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderSubtle')
                                $b.TB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $b.ColorKey)
                            }
                        }
                        Update-Dashboard
                    }.GetNewClosure())
                }

                # Effort badge
                if ($Chk.Effort) {
                    $EffBadge = New-Object System.Windows.Controls.Border
                    $EffBadge.CornerRadius = [System.Windows.CornerRadius]::new(4)
                    $EffBadge.Padding = [System.Windows.Thickness]::new(6,3,6,3)
                    $EffBadge.Margin = [System.Windows.Thickness]::new(12,0,0,0)
                    $EffBadge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeDeepBg')
                    $EffTB = New-Object System.Windows.Controls.TextBlock
                    $EffTB.Text = $Chk.Effort; $EffTB.FontSize = 10
                    $EffTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
                    $EffBadge.Child = $EffTB
                    [void]$ActionRow.Children.Add($EffBadge)
                }

                [void]$SP.Children.Add($ActionRow)
            }

            # ── Row 5: Remediation + CSP enrichment (shown for ALL checks) ──
            if ($Chk.Remediation -or $Chk.CspPath -or $Chk.Rationale) {
                    $RemBorder = New-Object System.Windows.Controls.Border
                    $RemBorder.Margin = [System.Windows.Thickness]::new(22,4,0,0)
                    $RemBorder.Padding = [System.Windows.Thickness]::new(10,6,10,6)
                    $RemBorder.CornerRadius = [System.Windows.CornerRadius]::new(4)
                    $RemBorder.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeDeepBg')
                    $RemBorder.BorderThickness = [System.Windows.Thickness]::new(1,0,0,0)
                    $RemBorder.BorderBrush = $Window.Resources['ThemeAccent']

                    $RemSP = New-Object System.Windows.Controls.StackPanel

                    # Rationale: "Why this matters" section
                    if ($Chk.Rationale) {
                        $RatLbl = New-Object System.Windows.Controls.TextBlock
                        $RatLbl.Text = "Why this matters"; $RatLbl.FontSize = 10; $RatLbl.FontWeight = [System.Windows.FontWeights]::Bold
                        $RatLbl.Foreground = $Window.Resources['ThemeWarning']
                        $RatLbl.Margin = [System.Windows.Thickness]::new(0,0,0,3)
                        $RatTB = New-Object System.Windows.Controls.TextBlock
                        $RatTB.Text = $Chk.Rationale; $RatTB.FontSize = 11; $RatTB.TextWrapping = 'Wrap'
                        $RatTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextSecondary')
                        $RatTB.Margin = [System.Windows.Thickness]::new(0,0,0,6)
                        [void]$RemSP.Children.Add($RatLbl)
                        [void]$RemSP.Children.Add($RatTB)
                    }

                    if ($Chk.Remediation) {
                        $RemLbl = New-Object System.Windows.Controls.TextBlock
                        $RemLbl.Text = "Remediation"; $RemLbl.FontSize = 10; $RemLbl.FontWeight = [System.Windows.FontWeights]::Bold
                        $RemLbl.Foreground = $Window.Resources['ThemeAccentLight']
                        $RemLbl.Margin = [System.Windows.Thickness]::new(0,0,0,4)
                        $RemTB = New-Object System.Windows.Controls.TextBlock
                        $RemTB.Text = $Chk.Remediation; $RemTB.FontSize = 11; $RemTB.TextWrapping = 'Wrap'
                        $RemTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
                        [void]$RemSP.Children.Add($RemLbl)
                        [void]$RemSP.Children.Add($RemTB)
                    }

                    # CSP enrichment from csp_metadata.json
                    $CspData = if ($Chk.CspPath) { Get-CspEnrichment $Chk.CspPath } else { $null }
                    if ($CspData) {
                        $CspSep = New-Object System.Windows.Controls.Border
                        $CspSep.Height = 1; $CspSep.Margin = [System.Windows.Thickness]::new(0,6,0,6)
                        $CspSep.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeBorder')
                        [void]$RemSP.Children.Add($CspSep)

                        $CspLbl = New-Object System.Windows.Controls.TextBlock
                        $CspLbl.Text = "Policy Details (CSP)"; $CspLbl.FontSize = 10; $CspLbl.FontWeight = [System.Windows.FontWeights]::Bold
                        $CspLbl.Foreground = $Window.Resources['ThemeGreenAccent']
                        $CspLbl.Margin = [System.Windows.Thickness]::new(0,0,0,4)
                        [void]$RemSP.Children.Add($CspLbl)

                        # CSP Path
                        $CspPathTB = New-Object System.Windows.Controls.TextBlock
                        $CspPathTB.Text = "CSP: ./Device/Vendor/MSFT/Policy/Config/$($Chk.CspPath)"
                        $CspPathTB.FontSize = 10; $CspPathTB.FontFamily = [System.Windows.Media.FontFamily]::new("Cascadia Mono, Consolas")
                        $CspPathTB.TextWrapping = 'Wrap'
                        $CspPathTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextSecondary')
                        [void]$RemSP.Children.Add($CspPathTB)

                        # GP Mapping
                        if ($CspData.GPMapping) {
                            $GpM = $CspData.GPMapping
                            if ($GpM.Path) {
                                $GpTB = New-Object System.Windows.Controls.TextBlock
                                $GpTB.Text = "GPO: $($GpM.Location) > $($GpM.Path)"
                                $GpTB.FontSize = 10; $GpTB.TextWrapping = 'Wrap'; $GpTB.Margin = [System.Windows.Thickness]::new(0,2,0,0)
                                $GpTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextSecondary')
                                [void]$RemSP.Children.Add($GpTB)
                            }
                            if ($GpM.'ADMX File Name') {
                                $AdmxTB = New-Object System.Windows.Controls.TextBlock
                                $AdmxTB.Text = "ADMX: $($GpM.'ADMX File Name')"; $AdmxTB.FontSize = 10
                                $AdmxTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
                                [void]$RemSP.Children.Add($AdmxTB)
                            }
                            if ($GpM.'Registry Key Name' -and $GpM.'Registry Value Name') {
                                $RegTB = New-Object System.Windows.Controls.TextBlock
                                $RegTB.Text = "Registry: HKLM\$($GpM.'Registry Key Name')\$($GpM.'Registry Value Name')"
                                $RegTB.FontSize = 10; $RegTB.FontFamily = [System.Windows.Media.FontFamily]::new("Cascadia Mono, Consolas")
                                $RegTB.TextWrapping = 'Wrap'; $RegTB.Margin = [System.Windows.Thickness]::new(0,2,0,0)
                                $RegTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
                                [void]$RemSP.Children.Add($RegTB)
                            }
                        }

                        # Allowed values
                        if ($CspData.AllowedValues) {
                            $AvLbl = New-Object System.Windows.Controls.TextBlock
                            $AvLbl.Text = "Allowed Values:"; $AvLbl.FontSize = 10; $AvLbl.Margin = [System.Windows.Thickness]::new(0,4,0,2)
                            $AvLbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
                            [void]$RemSP.Children.Add($AvLbl)
                            foreach ($avProp in $CspData.AllowedValues.PSObject.Properties) {
                                $AvRow = New-Object System.Windows.Controls.TextBlock
                                $AvRow.Text = "  $($avProp.Name) = $($avProp.Value)"
                                $AvRow.FontSize = 10; $AvRow.FontFamily = [System.Windows.Media.FontFamily]::new("Cascadia Mono, Consolas")
                                $AvRow.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
                                [void]$RemSP.Children.Add($AvRow)
                            }
                        }
                    }

                    $RemBorder.Child = $RemSP
                    [void]$SP.Children.Add($RemBorder)
                }

            # ── Row 6: Notes field ──
            $NotesRow = New-Object System.Windows.Controls.DockPanel
            $NotesRow.Margin = [System.Windows.Thickness]::new(22,6,0,0)
            $NotesLbl = New-Object System.Windows.Controls.TextBlock
            $NotesLbl.Text = "Notes:"; $NotesLbl.FontSize = 10; $NotesLbl.VerticalAlignment = 'Top'
            $NotesLbl.Margin = [System.Windows.Thickness]::new(0,4,8,0)
            $NotesLbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
            $NotesTxt = New-Object System.Windows.Controls.TextBox
            $NotesTxt.Text = $Chk.Notes; $NotesTxt.FontSize = 11; $NotesTxt.Height = 28
            $NotesTxt.Padding = [System.Windows.Thickness]::new(6,4,6,4)
            $NotesTxt.TextWrapping = 'Wrap'
            $NotesTxt.SetResourceReference([System.Windows.Controls.TextBox]::BackgroundProperty, 'ThemeInputBg')
            $NotesTxt.SetResourceReference([System.Windows.Controls.TextBox]::ForegroundProperty, 'ThemeTextBody')
            $NotesTxt.SetResourceReference([System.Windows.Controls.TextBox]::BorderBrushProperty, 'ThemeBorderSubtle')
            $NotesTxt.BorderThickness = [System.Windows.Thickness]::new(1)

            $NotesChkRef = $Chk
            $NotesTxt.Add_LostFocus({
                if ($NotesChkRef.Notes -ne $NotesTxt.Text) {
                    $NotesChkRef.Notes = $NotesTxt.Text
                    Set-Dirty
                }
            }.GetNewClosure())

            [void]$NotesRow.Children.Add($NotesLbl)
            [void]$NotesRow.Children.Add($NotesTxt)
            [void]$SP.Children.Add($NotesRow)

            [void]$OuterGrid.Children.Add($SP)
            $Card.Child = $OuterGrid
            [void]$ChecksPanel.Children.Add($Card)
        }

        [void]$Panel.Children.Add($ChecksPanel)
}


function Render-Findings {
    while ($pnlFindingsList.Children.Count -gt 2) {
        $pnlFindingsList.Children.RemoveAt($pnlFindingsList.Children.Count - 1)
    }

    # Start from all assessed checks; Status filter controls what's shown
    $StatusFilter = 'All'
    if ($cmbFilterStatus -and $cmbFilterStatus.SelectedItem) {
        $StatusFilter = $cmbFilterStatus.SelectedItem.Content
    }

    if ($StatusFilter -eq 'All') {
        # Default: show actionable findings (Fail, Warning, Deferred, Accepted Risk)
        $Findings = @($Global:Assessment.Checks | Where-Object { $_.Status -in @('Fail','Warning','Deferred','Accepted Risk') })
    } else {
        # Specific status selected — show all checks matching that status
        $Findings = @($Global:Assessment.Checks | Where-Object { $_.Status -eq $StatusFilter })
    }

    # Apply filters
    if ($cmbFilterSeverity -and $cmbFilterSeverity.SelectedItem) {
        $Sel = $cmbFilterSeverity.SelectedItem.Content
        if ($Sel -ne 'All') { $Findings = @($Findings | Where-Object { $_.Severity -eq $Sel }) }
    }
    if ($cmbFilterCategory -and $cmbFilterCategory.SelectedItem) {
        $Sel = $cmbFilterCategory.SelectedItem.Content
        if ($Sel -ne 'All') { $Findings = @($Findings | Where-Object { $_.Category -eq $Sel }) }
    }
    if ($cmbFilterPriority -and $cmbFilterPriority.SelectedItem) {
        $Sel = $cmbFilterPriority.SelectedItem.Content
        if ($Sel -ne 'All') { $Findings = @($Findings | Where-Object { $_.Priority -eq $Sel }) }
    }
    if ($cmbFilterOrigin -and $cmbFilterOrigin.SelectedItem) {
        $Sel = $cmbFilterOrigin.SelectedItem.Content
        if ($Sel -ne 'All') { $Findings = @($Findings | Where-Object { $_.Origin -match $Sel }) }
    }
    if ($cmbFilterEffort -and $cmbFilterEffort.SelectedItem) {
        $Sel = $cmbFilterEffort.SelectedItem.Content
        if ($Sel -ne 'All') { $Findings = @($Findings | Where-Object { $_.Effort -eq $Sel }) }
    }
    if ($txtFindingsSearch -and $txtFindingsSearch.Text.Trim()) {
        $Term = $txtFindingsSearch.Text.Trim()
        $Findings = @($Findings | Where-Object {
            $_.Name -match [regex]::Escape($Term) -or $_.Description -match [regex]::Escape($Term) -or $_.Id -match [regex]::Escape($Term)
        })
    }

    # Sort by severity weight desc
    $SevOrder = @{ Critical=0; High=1; Medium=2; Low=3 }
    $Findings = @($Findings | Sort-Object { $SevOrder[$_.Severity] })

    $lblFindingsCount.Text = "$($Findings.Count) findings"

    foreach ($F in $Findings) {
        $Card = New-Object System.Windows.Controls.Border
        $Card.CornerRadius = [System.Windows.CornerRadius]::new(8)
        $Card.Padding = [System.Windows.Thickness]::new(16,12,16,12)
        $Card.Margin  = [System.Windows.Thickness]::new(0,0,0,8)
        $Card.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeCardBg')
        $Card.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderCard')
        $Card.BorderThickness = [System.Windows.Thickness]::new(1)

        $SP = New-Object System.Windows.Controls.StackPanel

        # Header row
        $HRow = New-Object System.Windows.Controls.DockPanel
        $StatusIcon = New-Object System.Windows.Controls.TextBlock
        $StatusIcon.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
        $StatusIcon.FontSize = 14; $StatusIcon.VerticalAlignment = 'Center'
        $StatusIcon.Margin = [System.Windows.Thickness]::new(0,0,8,0)
        if ($F.Status -eq 'Fail') { $StatusIcon.Text = [char]0xEA39; $StatusIcon.Foreground = $Window.Resources['ThemeError'] }
        elseif ($F.Status -eq 'Accepted Risk') { $StatusIcon.Text = [char]0xE8FB; $StatusIcon.Foreground = $Window.Resources['ThemeWarning'] }
        elseif ($F.Status -eq 'Deferred') { $StatusIcon.Text = [char]0xE823; $StatusIcon.Foreground = $Window.Resources['ThemeAccent'] }
        else { $StatusIcon.Text = [char]0xE7BA; $StatusIcon.Foreground = $Window.Resources['ThemeWarning'] }

        # Severity + effort badges on right
        $BadgePanel = New-Object System.Windows.Controls.StackPanel
        $BadgePanel.Orientation = 'Horizontal'
        [System.Windows.Controls.DockPanel]::SetDock($BadgePanel, 'Right')

        $EffortBadge = New-Object System.Windows.Controls.Border
        $EffortBadge.CornerRadius = [System.Windows.CornerRadius]::new(4)
        $EffortBadge.Padding = [System.Windows.Thickness]::new(6,2,6,2)
        $EffortBadge.Margin = [System.Windows.Thickness]::new(4,0,0,0)
        $EffortBadge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeSurfaceBg')
        $EffortTB = New-Object System.Windows.Controls.TextBlock
        $EffortTB.Text = $F.Effort; $EffortTB.FontSize = 10
        $EffortTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
        $EffortBadge.Child = $EffortTB
        [void]$BadgePanel.Children.Add($EffortBadge)

        $NameTB = New-Object System.Windows.Controls.TextBlock
        $NameTB.Text = "$($F.Id)  $($F.Name)"
        $NameTB.FontSize = 13; $NameTB.FontWeight = [System.Windows.FontWeights]::SemiBold
        $NameTB.VerticalAlignment = 'Center'; $NameTB.TextTrimming = 'CharacterEllipsis'
        $NameTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')

        [void]$HRow.Children.Add($BadgePanel)
        [void]$HRow.Children.Add($StatusIcon)
        [void]$HRow.Children.Add($NameTB)
        [void]$SP.Children.Add($HRow)

        # Impact
        if ($F.Impact) {
            $ImpTB = New-Object System.Windows.Controls.TextBlock
            $ImpTB.Text = $F.Impact; $ImpTB.FontSize = 11; $ImpTB.TextWrapping = 'Wrap'
            $ImpTB.Margin = [System.Windows.Thickness]::new(22,6,0,0)
            $ImpTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextSecondary')
            [void]$SP.Children.Add($ImpTB)
        }

        # Remediation
        if ($F.Remediation) {
            $RemLabel = New-Object System.Windows.Controls.TextBlock
            $RemLabel.Text = "Remediation:"; $RemLabel.FontSize = 11; $RemLabel.FontWeight = [System.Windows.FontWeights]::SemiBold
            $RemLabel.Margin = [System.Windows.Thickness]::new(22,6,0,2)
            $RemLabel.Foreground = $Window.Resources['ThemeAccentLight']
            [void]$SP.Children.Add($RemLabel)

            $RemTB = New-Object System.Windows.Controls.TextBlock
            $RemTB.Text = $F.Remediation; $RemTB.FontSize = 11; $RemTB.TextWrapping = 'Wrap'
            $RemTB.Margin = [System.Windows.Thickness]::new(22,0,0,0)
            $RemTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
            [void]$SP.Children.Add($RemTB)
        }

        $Card.Child = $SP
        [void]$pnlFindingsList.Children.Add($Card)
    }
}

function Get-ExecutiveSummary {
    # Generates a plain-text executive summary from evaluated assessment data.
    # Used for both the Report tab display and clipboard copy.
    $Checks   = $Global:Assessment.Checks
    $Data     = $Global:Assessment.CollectionData
    $Score    = Get-OverallScore
    $Compliance = Get-BaselineCompliancePercent
    $Maturity = Get-MaturityLevel $Score

    $Total     = $Checks.Count
    $Assessed  = @($Checks | Where-Object { $_.Status -ne 'Not Assessed' }).Count
    $Pass      = @($Checks | Where-Object { $_.Status -eq 'Pass' }).Count
    $Fail      = @($Checks | Where-Object { $_.Status -eq 'Fail' }).Count
    $Warn      = @($Checks | Where-Object { $_.Status -eq 'Warning' }).Count
    $NA        = @($Checks | Where-Object { $_.Status -eq 'N/A' }).Count
    $Accepted  = @($Checks | Where-Object { $_.Status -eq 'Accepted Risk' }).Count
    $Deferred  = @($Checks | Where-Object { $_.Status -eq 'Deferred' }).Count

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine('BaselinePilot Executive Summary')
    [void]$sb.AppendLine('================================')
    [void]$sb.AppendLine("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
    [void]$sb.AppendLine('')

    # ── Device Info ──
    [void]$sb.AppendLine('DEVICE INFORMATION')
    [void]$sb.AppendLine('------------------')
    if ($Data -and $Data.systemInfo) {
        $si = $Data.systemInfo
        [void]$sb.AppendLine("  Hostname:     $($si.hostname)")
        [void]$sb.AppendLine("  OS:           $($si.OSCaption)")
        [void]$sb.AppendLine("  Build:        $($si.displayVersion) (Build $($si.osBuild))")
        [void]$sb.AppendLine("  CPU:          $($si.cpuName.Trim())")
        [void]$sb.AppendLine("  RAM:          $($si.ramGB) GB")
        [void]$sb.AppendLine("  TPM:          $(if ($si.tpmVersion) { $si.tpmVersion } else { 'N/A' })")
        [void]$sb.AppendLine("  Secure Boot:  $(if ($si.secureBootEnabled) { 'Enabled' } else { 'Disabled' })")
        [void]$sb.AppendLine("  Supported:    $(if ($si.isSupported) { 'Yes' } else { 'No' })")
    }
    [void]$sb.AppendLine("  Join Type:    $($Global:Assessment.JoinType)")
    [void]$sb.AppendLine("  Customer:     $($Global:Assessment.CustomerName)")
    [void]$sb.AppendLine("  Assessor:     $($Global:Assessment.AssessorName)")
    if ($Data -and $Data._metadata -and $Data._metadata.timestamp) {
        [void]$sb.AppendLine("  Collected:    $($Data._metadata.timestamp)")
    }
    [void]$sb.AppendLine('')

    # ── Score Summary ──
    [void]$sb.AppendLine('SCORE SUMMARY')
    [void]$sb.AppendLine('-------------')
    [void]$sb.AppendLine("  Compliance:       $(if ($Compliance -ge 0) { "$Compliance%" } else { 'N/A' })")
    [void]$sb.AppendLine("  Risk Score:       $(if ($Score -ge 0) { "$Score/100" } else { 'N/A' })")
    [void]$sb.AppendLine("  Maturity Level:   $Maturity")
    [void]$sb.AppendLine("  Total Checks:     $Total ($Assessed evaluated)")
    [void]$sb.AppendLine("  Pass: $Pass  |  Fail: $Fail  |  Warning: $Warn  |  N/A: $NA")
    if ($Accepted -gt 0 -or $Deferred -gt 0) {
        [void]$sb.AppendLine("  Accepted Risk: $Accepted  |  Deferred: $Deferred")
    }
    [void]$sb.AppendLine('')

    # ── Category Scores ──
    [void]$sb.AppendLine('CATEGORY SCORES')
    [void]$sb.AppendLine('---------------')
    foreach ($Cat in (Get-Categories)) {
        $CatScore = Get-CategoryScore $Cat
        $CatChecks = @($Checks | Where-Object { $_.Category -eq $Cat })
        $CatFail   = @($CatChecks | Where-Object { $_.Status -eq 'Fail' -or $_.Status -eq 'Deferred' }).Count
        $CatPass   = @($CatChecks | Where-Object { $_.Status -eq 'Pass' }).Count
        $ScoreStr  = if ($CatScore -ge 0) { "$CatScore%" } else { 'N/A' }
        [void]$sb.AppendLine("  $($Cat.PadRight(32)) $($ScoreStr.PadLeft(5))   ($CatPass pass, $CatFail fail)")
    }
    [void]$sb.AppendLine('')

    # ── Key Passes (strong areas) ──
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
        ($_.Status -eq 'Fail' -or $_.Status -eq 'Deferred') -and
        $_.Severity -in @('Critical','High')
    } | Sort-Object @{E={if($_.Severity -eq 'Critical'){0}else{1}}}, Id)
    if ($critFails.Count -gt 0) {
        [void]$sb.AppendLine("CRITICAL & HIGH FAILURES ($($critFails.Count))")
        [void]$sb.AppendLine('---------------------------------------------')
        foreach ($f in $critFails) {
            $statusTag = if ($f.Status -eq 'Deferred') { ' [DEFERRED]' } else { '' }
            [void]$sb.AppendLine("  [$($f.Severity.ToUpper().PadRight(8))] $($f.Id)  $($f.Name)$statusTag")
            if ($f.ActualValue) {
                [void]$sb.AppendLine("             Current: $($f.ActualValue)   Expected: $($f.BaselineValue)")
            }
        }
        [void]$sb.AppendLine('')
    }

    # ── Medium/Low Failures (count only) ──
    $medLowFails = @($Checks | Where-Object {
        ($_.Status -eq 'Fail' -or $_.Status -eq 'Deferred') -and
        $_.Severity -in @('Medium','Low')
    })
    if ($medLowFails.Count -gt 0) {
        $medCount = @($medLowFails | Where-Object { $_.Severity -eq 'Medium' }).Count
        $lowCount = @($medLowFails | Where-Object { $_.Severity -eq 'Low' }).Count
        [void]$sb.AppendLine("OTHER FAILURES: $medCount Medium, $lowCount Low")
        [void]$sb.AppendLine('')
    }

    # ── Quick Win Remediation Priority ──
    $quickWins = @($Checks | Where-Object {
        $_.Status -eq 'Fail' -and $_.Effort -eq 'Quick Win'
    } | Sort-Object @{E={switch($_.Severity){'Critical'{0}'High'{1}'Medium'{2}default{3}}}}, Id | Select-Object -First 10)
    if ($quickWins.Count -gt 0) {
        [void]$sb.AppendLine("QUICK WIN REMEDIATION (top $($quickWins.Count))")
        [void]$sb.AppendLine('----------------------------------------------')
        $qi = 0
        foreach ($qw in $quickWins) {
            $qi++
            [void]$sb.AppendLine("  $qi. [$($qw.Severity)] $($qw.Name) ($($qw.Id))")
            if ($qw.Remediation) {
                $remText = if ($qw.Remediation.Length -gt 100) { $qw.Remediation.Substring(0,100) + '...' } else { $qw.Remediation }
                [void]$sb.AppendLine("     $remText")
            }
        }
        [void]$sb.AppendLine('')
    }

    # ── Accepted Risk / Deferred items ──
    $overrides = @($Checks | Where-Object { $_.Status -eq 'Accepted Risk' -or $_.Status -eq 'Deferred' })
    if ($overrides.Count -gt 0) {
        [void]$sb.AppendLine("GOVERNANCE OVERRIDES ($($overrides.Count))")
        [void]$sb.AppendLine('--------------------------------------')
        foreach ($ov in $overrides) {
            [void]$sb.AppendLine("  [$($ov.Status.PadRight(13))] $($ov.Id)  $($ov.Name)")
        }
        [void]$sb.AppendLine('')
    }

    # ── Scoring Methodology ──
    [void]$sb.AppendLine('SCORING METHODOLOGY')
    [void]$sb.AppendLine('-------------------')
    [void]$sb.AppendLine('  Compliance = Pass / (Total - N/A - Not Assessed - Accepted Risk)')
    [void]$sb.AppendLine('  Risk Score = weighted: Critical(5x), High(4x), Medium(3x), Low(2x)')
    [void]$sb.AppendLine('  Pass=100pts, Warning=50pts, Fail/Deferred=0pts')
    [void]$sb.AppendLine('  Accepted Risk items excluded from both scores.')

    return $sb.ToString()
}

function Get-ExecutiveSummaryRtf {
    # Generates a richly formatted RTF document from evaluated assessment data.
    # Pastes beautifully into Word, Outlook, Teams, OneNote.
    $Checks     = $Global:Assessment.Checks
    $Data       = $Global:Assessment.CollectionData
    $Score      = Get-OverallScore
    $Compliance = Get-BaselineCompliancePercent
    $Maturity   = Get-MaturityLevel $Score

    $Total    = $Checks.Count
    $Assessed = @($Checks | Where-Object { $_.Status -ne 'Not Assessed' }).Count
    $Pass     = @($Checks | Where-Object { $_.Status -eq 'Pass' }).Count
    $Fail     = @($Checks | Where-Object { $_.Status -eq 'Fail' }).Count
    $Warn     = @($Checks | Where-Object { $_.Status -eq 'Warning' }).Count
    $NA       = @($Checks | Where-Object { $_.Status -eq 'N/A' }).Count
    $Accepted = @($Checks | Where-Object { $_.Status -eq 'Accepted Risk' }).Count
    $Deferred = @($Checks | Where-Object { $_.Status -eq 'Deferred' }).Count

    # RTF helper — escape special chars
    $esc = { param($t) if (-not $t) { return '' }; "$t".Replace('\','\\').Replace('{','\{').Replace('}','\}') }

    # Color table indices (1-based after auto-entry 0):
    # 0=black(auto) 1=accent(#0078D4) 2=green(#107C10) 3=red(#D13438) 4=orange(#CA5010)
    # 5=gray(#8A8886) 6=darkBg(#F3F2F1) 7=white 8=blue(#0063B1) 9=amber(#D48C00)
    $rtf = [System.Text.StringBuilder]::new()
    [void]$rtf.Append('{\rtf1\ansi\deff0')
    [void]$rtf.Append('{\fonttbl{\f0\fswiss\fcharset0 Segoe UI;}{\f1\fmodern\fcharset0 Cascadia Mono;}}')
    [void]$rtf.Append('{\colortbl;')
    [void]$rtf.Append('\red0\green120\blue212;')   # 1 accent blue
    [void]$rtf.Append('\red16\green124\blue16;')    # 2 green
    [void]$rtf.Append('\red209\green52\blue56;')    # 3 red
    [void]$rtf.Append('\red202\green80\blue16;')    # 4 orange
    [void]$rtf.Append('\red138\green136\blue134;')  # 5 gray
    [void]$rtf.Append('\red243\green242\blue241;')  # 6 light bg
    [void]$rtf.Append('\red255\green255\blue255;')  # 7 white
    [void]$rtf.Append('\red0\green99\blue177;')     # 8 dark blue
    [void]$rtf.Append('\red212\green140\blue0;')    # 9 amber
    [void]$rtf.Append('\red50\green50\blue50;')     # 10 dark text
    [void]$rtf.Append('}')
    # Default font, size 20 half-points = 10pt
    [void]$rtf.Append('\f0\fs20\cf10 ')

    # ── Title ──
    [void]$rtf.Append('\pard\sb0\sa120\qc{\f0\fs36\b\cf1 BaselinePilot Executive Summary}\par')
    [void]$rtf.Append("{\f0\fs18\cf5 Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm')  \u8226  Microsoft Security Baseline Assessment}\par")
    # Thin accent line
    [void]$rtf.Append('\pard\sb40\sa120{\f0\fs4\cf1\brdrb\brdrs\brdrw10\brsp40 \par}')

    # ── Score Hero (2 big numbers side by side) ──
    [void]$rtf.Append('\pard\sb0\sa60\qc')
    $compStr = if ($Compliance -ge 0) { "$Compliance%" } else { 'N/A' }
    $riskStr = if ($Score -ge 0) { "$Score/100" } else { 'N/A' }
    # Compliance big number
    [void]$rtf.Append("{{\f0\fs56\b\cf1 $compStr}  }")
    [void]$rtf.Append('{\f0\fs20\cf5 Compliance}')
    [void]$rtf.Append('{\f0\fs28  \u8195  }')  # em space separator
    # Risk Score big number
    $riskColor = if ($Score -ge 75) { '\cf2' } elseif ($Score -ge 50) { '\cf4' } elseif ($Score -ge 0) { '\cf3' } else { '\cf5' }
    [void]$rtf.Append("{{\f0\fs56\b$riskColor $riskStr}  }")
    [void]$rtf.Append('{\f0\fs20\cf5 Risk Score}')
    [void]$rtf.Append('{\f0\fs28  \u8195  }')
    [void]$rtf.Append("{{\f0\fs28\b\cf8 $Maturity}  }")
    [void]$rtf.Append('{\f0\fs20\cf5 Maturity}\par')

    # Status counts bar
    [void]$rtf.Append('\pard\sb0\sa80\qc{\f0\fs18 ')
    [void]$rtf.Append("{\cf2\b $Pass} Pass   ")
    [void]$rtf.Append("{\cf3\b $Fail} Fail   ")
    [void]$rtf.Append("{\cf4\b $Warn} Warning   ")
    [void]$rtf.Append("{\cf5 $NA N/A}")
    if ($Accepted -gt 0 -or $Deferred -gt 0) {
        [void]$rtf.Append("   {\cf9 $Accepted Accepted Risk}   {\cf1 $Deferred Deferred}")
    }
    [void]$rtf.Append("   \u8226  $Total total, $Assessed evaluated")
    [void]$rtf.Append('}\par')

    # ═══════════════ Section helper ═══════════════
    $SectionHeader = {
        param($title, $icon)
        # Accent colored header with top spacing
        [void]$rtf.Append("\pard\sb200\sa80\keepn{\f0\fs24\b\cf1 $icon  $(& $esc $title)}\par")
        # Thin line under header
        [void]$rtf.Append('\pard\sb0\sa60{\f0\fs2\cf5\brdrb\brdrs\brdrw5\brsp20 \par}')
    }

    # ═══════════════ DEVICE INFORMATION ═══════════════
    & $SectionHeader 'Device Information' '\u9889'
    # Table: 2 columns, key-value pairs
    [void]$rtf.Append('\pard\sb0\sa0{\trowd\trgaph80')
    [void]$rtf.Append('\clcbpat6\cellx2600\cellx8500')
    $deviceRows = [System.Collections.ArrayList]::new()
    if ($Data -and $Data.systemInfo) {
        $si = $Data.systemInfo
        [void]$deviceRows.Add(@('Hostname',     $si.hostname))
        [void]$deviceRows.Add(@('OS',           $si.OSCaption))
        [void]$deviceRows.Add(@('Build',        "$($si.displayVersion) (Build $($si.osBuild))"))
        [void]$deviceRows.Add(@('CPU',          "$($si.cpuName)".Trim()))
        [void]$deviceRows.Add(@('RAM',          "$($si.ramGB) GB"))
        [void]$deviceRows.Add(@('TPM',          $(if ($si.tpmVersion) { $si.tpmVersion } else { 'N/A' })))
        [void]$deviceRows.Add(@('Secure Boot',  $(if ($si.secureBootEnabled) { 'Enabled' } else { 'Disabled' })))
    }
    [void]$deviceRows.Add(@('Join Type',    $Global:Assessment.JoinType))
    [void]$deviceRows.Add(@('Customer',     $Global:Assessment.CustomerName))
    [void]$deviceRows.Add(@('Assessor',     $Global:Assessment.AssessorName))
    if ($Data -and $Data._metadata -and $Data._metadata.timestamp) {
        [void]$deviceRows.Add(@('Collected', "$($Data._metadata.timestamp)"))
    }
    foreach ($row in $deviceRows) {
        [void]$rtf.Append('\trowd\trgaph80\clcbpat6\cellx2600\cellx8500')
        [void]$rtf.Append("\pard\intbl\sb20\sa20{\f0\fs18\b\cf8  $(& $esc $row[0])}\cell")
        [void]$rtf.Append("\pard\intbl\sb20\sa20{\f0\fs18\cf10  $(& $esc $row[1])}\cell")
        [void]$rtf.Append('\row')
    }
    [void]$rtf.Append('}')

    # ═══════════════ CATEGORY SCORES ═══════════════
    & $SectionHeader 'Category Scores' '\u9733'
    # Table header
    [void]$rtf.Append('\pard\sb0\sa0{\trowd\trgaph80')
    [void]$rtf.Append('\clcbpat1\cellx4200\clcbpat1\cellx5400\clcbpat1\cellx6400\clcbpat1\cellx7400\clcbpat1\cellx8500')
    [void]$rtf.Append('\pard\intbl\sb30\sa30{\f0\fs17\b\cf7  Category}\cell')
    [void]$rtf.Append('\pard\intbl\sb30\sa30\qc{\f0\fs17\b\cf7 Score}\cell')
    [void]$rtf.Append('\pard\intbl\sb30\sa30\qc{\f0\fs17\b\cf7 Pass}\cell')
    [void]$rtf.Append('\pard\intbl\sb30\sa30\qc{\f0\fs17\b\cf7 Fail}\cell')
    [void]$rtf.Append('\pard\intbl\sb30\sa30\qc{\f0\fs17\b\cf7 Total}\cell')
    [void]$rtf.Append('\row')
    $catIdx = 0
    foreach ($Cat in (Get-Categories)) {
        $CatScore  = Get-CategoryScore $Cat
        $CatChecks = @($Checks | Where-Object { $_.Category -eq $Cat })
        $CatTotal  = $CatChecks.Count
        $CatFail   = @($CatChecks | Where-Object { $_.Status -eq 'Fail' -or $_.Status -eq 'Deferred' }).Count
        $CatPass   = @($CatChecks | Where-Object { $_.Status -eq 'Pass' }).Count
        $ScoreStr  = if ($CatScore -ge 0) { "$CatScore%" } else { 'N/A' }
        $scoreClr  = if ($CatScore -ge 80) { '\cf2' } elseif ($CatScore -ge 50) { '\cf4' } elseif ($CatScore -ge 0) { '\cf3' } else { '\cf5' }
        $bgPat     = if ($catIdx % 2 -eq 0) { '\clcbpat6' } else { '' }
        [void]$rtf.Append("\trowd\trgaph80${bgPat}\cellx4200${bgPat}\cellx5400${bgPat}\cellx6400${bgPat}\cellx7400${bgPat}\cellx8500")
        [void]$rtf.Append("\pard\intbl\sb20\sa20{\f0\fs18\cf10  $(& $esc $Cat)}\cell")
        [void]$rtf.Append("\pard\intbl\sb20\sa20\qc{\f0\fs18\b$scoreClr $ScoreStr}\cell")
        [void]$rtf.Append("\pard\intbl\sb20\sa20\qc{\f0\fs18\cf2 $CatPass}\cell")
        [void]$rtf.Append("\pard\intbl\sb20\sa20\qc{\f0\fs18\cf3 $CatFail}\cell")
        [void]$rtf.Append("\pard\intbl\sb20\sa20\qc{\f0\fs18\cf5 $CatTotal}\cell")
        [void]$rtf.Append('\row')
        $catIdx++
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
        ($_.Status -eq 'Fail' -or $_.Status -eq 'Deferred') -and
        $_.Severity -in @('Critical','High')
    } | Sort-Object @{E={if($_.Severity -eq 'Critical'){0}else{1}}}, Id)
    if ($critFails.Count -gt 0) {
        & $SectionHeader "Critical & High Failures ($($critFails.Count))" '\u9888'
        # Table
        [void]$rtf.Append('\pard\sb0\sa0{\trowd\trgaph80')
        [void]$rtf.Append('\clcbpat1\cellx1200\clcbpat1\cellx2200\clcbpat1\cellx5600\clcbpat1\cellx7000\clcbpat1\cellx8500')
        [void]$rtf.Append('\pard\intbl\sb30\sa30{\f0\fs17\b\cf7  Sev}\cell')
        [void]$rtf.Append('\pard\intbl\sb30\sa30{\f0\fs17\b\cf7  ID}\cell')
        [void]$rtf.Append('\pard\intbl\sb30\sa30{\f0\fs17\b\cf7  Check Name}\cell')
        [void]$rtf.Append('\pard\intbl\sb30\sa30{\f0\fs17\b\cf7  Current}\cell')
        [void]$rtf.Append('\pard\intbl\sb30\sa30{\f0\fs17\b\cf7  Expected}\cell')
        [void]$rtf.Append('\row')
        $fi = 0
        foreach ($f in $critFails) {
            $sevClr  = if ($f.Severity -eq 'Critical') { '\cf3' } else { '\cf4' }
            $bgPat   = if ($fi % 2 -eq 0) { '\clcbpat6' } else { '' }
            $curVal  = if ($f.ActualValue) { & $esc "$($f.ActualValue)" } else { '\u8212' }
            $expVal  = if ($f.BaselineValue) { & $esc "$($f.BaselineValue)" } else { '\u8212' }
            $nameStr = & $esc $f.Name
            if ($f.Status -eq 'Deferred') { $nameStr = "$nameStr {\f0\fs15\i\cf1 [DEFERRED]}" }
            [void]$rtf.Append("\trowd\trgaph80${bgPat}\cellx1200${bgPat}\cellx2200${bgPat}\cellx5600${bgPat}\cellx7000${bgPat}\cellx8500")
            [void]$rtf.Append("\pard\intbl\sb20\sa20{\f0\fs17\b$sevClr  $($f.Severity)}\cell")
            [void]$rtf.Append("\pard\intbl\sb20\sa20{\f1\fs16\cf8  $($f.Id)}\cell")
            [void]$rtf.Append("\pard\intbl\sb20\sa20{\f0\fs17\cf10  $nameStr}\cell")
            [void]$rtf.Append("\pard\intbl\sb20\sa20{\f1\fs16\cf3  $curVal}\cell")
            [void]$rtf.Append("\pard\intbl\sb20\sa20{\f1\fs16\cf2  $expVal}\cell")
            [void]$rtf.Append('\row')
            $fi++
        }
        [void]$rtf.Append('}')
    }

    # ═══════════════ OTHER FAILURES ═══════════════
    $medLowFails = @($Checks | Where-Object {
        ($_.Status -eq 'Fail' -or $_.Status -eq 'Deferred') -and
        $_.Severity -in @('Medium','Low')
    })
    if ($medLowFails.Count -gt 0) {
        $medCount = @($medLowFails | Where-Object { $_.Severity -eq 'Medium' }).Count
        $lowCount = @($medLowFails | Where-Object { $_.Severity -eq 'Low' }).Count
        [void]$rtf.Append("\pard\sb120\sa60{\f0\fs18\cf5 Additional failures: {\b\cf4 $medCount} Medium, {\b\cf5 $lowCount} Low}\par")
    }

    # ═══════════════ QUICK WIN REMEDIATION ═══════════════
    $quickWins = @($Checks | Where-Object {
        $_.Status -eq 'Fail' -and $_.Effort -eq 'Quick Win'
    } | Sort-Object @{E={switch($_.Severity){'Critical'{0}'High'{1}'Medium'{2}default{3}}}}, Id | Select-Object -First 10)
    if ($quickWins.Count -gt 0) {
        & $SectionHeader "Quick Win Remediation (top $($quickWins.Count))" '\u9889'
        $qi = 0
        foreach ($qw in $quickWins) {
            $qi++
            $sevClr = switch ($qw.Severity) { 'Critical' { '\cf3' } 'High' { '\cf4' } default { '\cf5' } }
            [void]$rtf.Append("\pard\sb40\sa0\li200{\f0\fs18\b\cf10 ${qi}. }{\f0\fs17$sevClr [$($qw.Severity)]} {\f0\fs18\cf10 $(& $esc $qw.Name)} {\f1\fs15\cf5 ($($qw.Id))}\par")
            if ($qw.Remediation) {
                $remText = if ($qw.Remediation.Length -gt 120) { $qw.Remediation.Substring(0,120) + '...' } else { $qw.Remediation }
                [void]$rtf.Append("\pard\sb0\sa20\li400{\f0\fs16\i\cf5 $(& $esc $remText)}\par")
            }
        }
    }

    # ═══════════════ GOVERNANCE OVERRIDES ═══════════════
    $overrides = @($Checks | Where-Object { $_.Status -eq 'Accepted Risk' -or $_.Status -eq 'Deferred' })
    if ($overrides.Count -gt 0) {
        & $SectionHeader "Governance Overrides ($($overrides.Count))" '\u9881'
        foreach ($ov in $overrides) {
            $stClr = if ($ov.Status -eq 'Accepted Risk') { '\cf9' } else { '\cf1' }
            [void]$rtf.Append("\pard\sb20\sa0\li200{\f0\fs17$stClr [$($ov.Status)]} {\f1\fs15\cf8 $($ov.Id)} {\f0\fs17\cf10 $(& $esc $ov.Name)}\par")
        }
    }

    # ═══════════════ SCORING METHODOLOGY ═══════════════
    [void]$rtf.Append('\pard\sb200\sa60{\f0\fs2\cf5\brdrb\brdrs\brdrw5\brsp20 \par}')
    [void]$rtf.Append('\pard\sb40\sa20{\f0\fs16\i\cf5 Scoring: Compliance = Pass / (Total \u8722 N/A \u8722 Not Assessed \u8722 Accepted Risk). ')
    [void]$rtf.Append('Risk Score = weighted average: Critical(5\u215), High(4\u215), Medium(3\u215), Low(2\u215). ')
    [void]$rtf.Append('Pass=100pts, Warning=50pts, Fail/Deferred=0pts. Accepted Risk excluded from both scores.}\par')

    [void]$rtf.Append('}')
    return $rtf.ToString()
}

function Update-ReportPreview {
    $paraReportSummary.Inlines.Clear()
    $Global:ExecutiveSummaryText = Get-ExecutiveSummary

    # Render with section-colored headers
    $HeaderBrush  = $Window.Resources['ThemeAccent']
    $BodyBrush    = $Window.Resources['ThemeTextBody']
    $MutedBrush   = $Window.Resources['ThemeTextMuted']

    foreach ($line in ($Global:ExecutiveSummaryText -split "`n")) {
        if ($line -match '^={3,}' -or $line -match '^-{3,}') {
            # Separator lines in muted
            $run = New-Object System.Windows.Documents.Run("$line`n")
            $run.Foreground = $MutedBrush
            $paraReportSummary.Inlines.Add($run)
        } elseif ($line -match '^[A-Z][A-Z &/()]+\s*$' -or $line -match '^[A-Z][A-Z &/(]+\(') {
            # ALL-CAPS section headers in accent color
            $run = New-Object System.Windows.Documents.Run("$line`n")
            $run.Foreground = $HeaderBrush
            $run.FontWeight = 'Bold'
            $paraReportSummary.Inlines.Add($run)
        } elseif ($line -match '^BaselinePilot') {
            # Title in accent + large
            $run = New-Object System.Windows.Documents.Run("$line`n")
            $run.Foreground = $HeaderBrush
            $run.FontWeight = 'Bold'
            $run.FontSize = 16
            $paraReportSummary.Inlines.Add($run)
        } elseif ($line -match '\[CRITICAL') {
            $run = New-Object System.Windows.Documents.Run("$line`n")
            $run.Foreground = [System.Windows.Media.Brushes]::IndianRed
            $paraReportSummary.Inlines.Add($run)
        } elseif ($line -match '\[HIGH') {
            $run = New-Object System.Windows.Documents.Run("$line`n")
            $run.Foreground = [System.Windows.Media.Brushes]::Orange
            $paraReportSummary.Inlines.Add($run)
        } elseif ($line -match '^\s+\+') {
            # Pass items (+ prefix) in green
            $run = New-Object System.Windows.Documents.Run("$line`n")
            $run.Foreground = [System.Windows.Media.Brushes]::MediumSeaGreen
            $paraReportSummary.Inlines.Add($run)
        } else {
            $run = New-Object System.Windows.Documents.Run("$line`n")
            $run.Foreground = $BodyBrush
            $paraReportSummary.Inlines.Add($run)
        }
    }
}

# ===============================================================================
# SECTION 13: COLLECTION IMPORT & AUTO-EVALUATION
# ===============================================================================

function Import-CollectionJson {
    param([string]$Path)

    try {
        $Json = Get-Content $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        Show-Toast "Failed to parse collection JSON: $($_.Exception.Message)" -Type 'Error'
        return
    }

    if (-not $Json._metadata -or -not $Json.systemInfo) {
        Show-Toast "Not a valid BaselinePilot collection file" -Type 'Error'
        return
    }

    $Global:Assessment.CollectionData = $Json

    # Track this machine in the multi-machine list
    $machineName = if ($Json.systemInfo.hostname) { $Json.systemInfo.hostname }
                   elseif ($Json.systemInfo.ComputerName) { $Json.systemInfo.ComputerName }
                   else { 'Unknown' }
    $machineEntry = @{
        Hostname      = $machineName
        ImportedAt    = (Get-Date).ToString('o')
        FilePath      = $Path
        OSBuild       = $Json.systemInfo.osBuild
        OSVersion     = if ($Json.systemInfo.versionName) { $Json.systemInfo.versionName } else { $Json.systemInfo.displayVersion }
        JoinType      = $null  # Set below
        IsSupported   = $Json.systemInfo.isSupported
        CollectedAt   = $Json._metadata.timestamp
    }
    # Check if machine already imported (by hostname) — replace if so
    $existingIdx = -1
    for ($mi = 0; $mi -lt $Global:Assessment.Machines.Count; $mi++) {
        if ($Global:Assessment.Machines[$mi].Hostname -eq $machineName) { $existingIdx = $mi; break }
    }
    if ($existingIdx -ge 0) {
        $Global:Assessment.Machines[$existingIdx] = $machineEntry
        Write-DebugLog "Updated existing machine: $machineName" -Level 'INFO'
    } else {
        [void]$Global:Assessment.Machines.Add($machineEntry)
        Write-DebugLog "Added new machine: $machineName (total: $($Global:Assessment.Machines.Count))" -Level 'INFO'
    }
    $Global:Assessment.MachineCount = $Global:Assessment.Machines.Count

    # Detect join type (collection stores booleans: azureAdJoined=true/false)
    $JT = $Json.joinType
    if ($JT) {
        if ($JT.azureAdJoined -and -not $JT.domainJoined) {
            $Global:Assessment.JoinType = 'Entra ID (Intune)'
        } elseif ($JT.domainJoined -and -not $JT.azureAdJoined) {
            $Global:Assessment.JoinType = 'AD DS (GPO)'
        } elseif ($JT.domainJoined -and $JT.azureAdJoined) {
            $Global:Assessment.JoinType = 'Hybrid Joined'
        } else {
            $Global:Assessment.JoinType = 'Workgroup'
        }
    }
    $machineEntry.JoinType = $Global:Assessment.JoinType

    # Update join type badge
    $lblJoinType.Text = $Global:Assessment.JoinType
    if ($Global:Assessment.JoinType -match 'Entra') {
        $lblJoinTypeIcon.Text = [char]0xE753
    } elseif ($Global:Assessment.JoinType -match 'AD DS') {
        $lblJoinTypeIcon.Text = [char]0xE968
    } else {
        $lblJoinTypeIcon.Text = [char]0xE8AF
    }

    # Auto-evaluate checks against collection data
    # Determine applicability based on join type
    $IsEntraOnly = $Global:Assessment.JoinType -match 'Entra'
    $IsADDS      = $Global:Assessment.JoinType -match 'AD DS'
    $IsHybrid    = $Global:Assessment.JoinType -match 'Hybrid'

    $EvalCount = 0
    $SkipCount = 0
    $TotalAuto = @($Global:Assessment.Checks | Where-Object { $_.Type -eq 'Auto' -and $_.CollectionKeys }).Count
    $Step = 0

    foreach ($Chk in $Global:Assessment.Checks) {
        if ($Chk.Type -ne 'Auto') { continue }
        if (-not $Chk.CollectionKeys) { continue }
        $Step++

        # Join type applicability filtering:
        # - Entra ID only: skip GPO-only checks (no domain controller, no GPOs applied)
        # - AD DS only: skip Intune-only checks (no MDM enrollment)
        # - Hybrid: evaluate everything
        if ($Chk.applicableTo -and $Chk.applicableTo.Count -gt 0) {
            $applicable = $false
            if ($IsEntraOnly) {
                $applicable = 'intune' -in $Chk.applicableTo
            } elseif ($IsADDS) {
                $applicable = 'gpo' -in $Chk.applicableTo
            } else {
                $applicable = $true  # Hybrid or Workgroup — check everything
            }
            if (-not $applicable) {
                $Chk.Status = 'N/A'
                $Chk.Details = "Not applicable — $($Global:Assessment.JoinType) device (requires $(($Chk.applicableTo | Where-Object { $_ -ne 'intune' -and $_ -ne 'gpo' }) -join ', ')$(if ('gpo' -in $Chk.applicableTo -and 'intune' -notin $Chk.applicableTo) { 'GPO/domain join' } elseif ('intune' -in $Chk.applicableTo -and 'gpo' -notin $Chk.applicableTo) { 'Intune/MDM' }))"
                $SkipCount++
                continue
            }
        }

        # Update progress bar every 10 checks
        if ($Step % 10 -eq 0 -and $barProgress) {
            $barProgress.Maximum = $TotalAuto
            $barProgress.Value = $Step
            [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
                [System.Windows.Threading.DispatcherPriority]::Background,
                [Action]{ }
            )
        }

        foreach ($Key in $Chk.CollectionKeys) {
            $ActualValue = Resolve-CollectionKey -Json $Json -Key $Key
            if ($null -ne $ActualValue) {
                $Chk.ActualValue = $ActualValue
                # Evaluate against baseline value if defined
                if ($null -ne $Chk.BaselineValue) {
                    $Expected = "$($Chk.BaselineValue)"
                    $Actual   = "$ActualValue"
                    if ($Actual -eq $Expected) {
                        $Chk.Status = 'Pass'
                    } else {
                        $Chk.Status = 'Fail'
                    }
                } else {
                    # No baseline value — check for numeric thresholds first
                    $thresholdHandled = $false

                    # Pre-process: if value is an array of boot/logon events and threshold expects ms, extract durations
                    if ($Chk.threshold -and ($null -ne $Chk.threshold.maxBootMs -or $null -ne $Chk.threshold.maxLogonMs) -and
                        ($ActualValue -is [System.Object[]] -or $ActualValue -is [System.Collections.ArrayList])) {
                        $boots = @($ActualValue)
                        $durations = @($boots | ForEach-Object { $_.BootDurationMs } | Where-Object { $null -ne $_ })
                        $bootCount = $boots.Count
                        if ($durations.Count -gt 0) {
                            $maxMs = ($durations | Measure-Object -Maximum).Maximum
                            $avgMs = [math]::Round(($durations | Measure-Object -Average).Average, 0)
                            $limitMs = if ($Chk.threshold.maxBootMs) { [double]$Chk.threshold.maxBootMs } else { [double]$Chk.threshold.maxLogonMs }
                            $limitLabel = if ($Chk.threshold.maxBootMs) { 'Boot' } else { 'Logon' }
                            $exceeded = @($durations | Where-Object { $_ -gt $limitMs }).Count
                            $Chk.ActualValue = "$([math]::Round($maxMs/1000,1))s max, $([math]::Round($avgMs/1000,1))s avg ($bootCount boots)"
                            if ($exceeded -gt 0) {
                                $Chk.Status = 'Fail'
                                $Chk.Details = "$exceeded of $($durations.Count) $($limitLabel.ToLower()) times exceed $([int]($limitMs/1000))s limit"
                            } else {
                                $Chk.Status = 'Pass'
                            }
                        } else {
                            # No duration data available — report boot count only
                            $dates = @($boots | ForEach-Object { $_.TimeCreated } | Where-Object { $_ })
                            $dateRange = if ($dates.Count -ge 2) { "$("$($dates[-1])".Substring(0,10)) to $("$($dates[0])".Substring(0,10))" } elseif ($dates.Count -eq 1) { "$($dates[0])".Substring(0,10) } else { '' }
                            $Chk.ActualValue = "$bootCount boot events (no duration data)"
                            $Chk.Status = 'Warning'
                            $Chk.Details = "Boot events found but BootDuration not recorded by OS$(if ($dateRange) { " ($dateRange)" })"
                        }
                        $thresholdHandled = $true
                    }

                    if (-not $thresholdHandled -and $Chk.threshold -and ($ActualValue -is [int] -or $ActualValue -is [long] -or $ActualValue -is [double])) {
                        $numVal = [double]$ActualValue
                        if ($null -ne $Chk.threshold.maxDays) {
                            # "days since X" — fail if over threshold
                            $limit = [double]$Chk.threshold.maxDays
                            $Chk.ActualValue = "$([int]$numVal) days"
                            if ($numVal -gt $limit) { $Chk.Status = 'Fail'; $Chk.Details = "Exceeds ${limit}-day limit" }
                            elseif ($numVal -gt ($limit * 0.8)) { $Chk.Status = 'Warning'; $Chk.Details = "Approaching ${limit}-day limit" }
                            else { $Chk.Status = 'Pass' }
                            $thresholdHandled = $true
                        } elseif ($null -ne $Chk.threshold.maxAge) {
                            # "age in days" — fail if stale
                            $limit = [double]$Chk.threshold.maxAge
                            $Chk.ActualValue = "$([int]$numVal) days old"
                            if ($numVal -gt $limit) { $Chk.Status = 'Fail'; $Chk.Details = "Stale — exceeds ${limit}-day max age" }
                            else { $Chk.Status = 'Pass' }
                            $thresholdHandled = $true
                        } elseif ($null -ne $Chk.threshold.minGB) {
                            # "minimum free space" — fail if below
                            $limit = [double]$Chk.threshold.minGB
                            $Chk.ActualValue = "$([math]::Round($numVal,1)) GB free"
                            if ($numVal -lt $limit) { $Chk.Status = 'Fail'; $Chk.Details = "Below ${limit} GB minimum" }
                            elseif ($numVal -lt ($limit * 2)) { $Chk.Status = 'Warning'; $Chk.Details = "Low — approaching ${limit} GB minimum" }
                            else { $Chk.Status = 'Pass' }
                            $thresholdHandled = $true
                        } elseif ($null -ne $Chk.threshold.maxBootMs) {
                            # "boot time in ms" — fail if slow
                            $limit = [double]$Chk.threshold.maxBootMs
                            $Chk.ActualValue = "$([math]::Round($numVal/1000,1))s"
                            if ($numVal -gt $limit) { $Chk.Status = 'Fail'; $Chk.Details = "Boot time exceeds $([int]($limit/1000))s limit" }
                            else { $Chk.Status = 'Pass' }
                            $thresholdHandled = $true
                        } elseif ($null -ne $Chk.threshold.maxLogonMs) {
                            # "logon time in ms" — fail if slow
                            $limit = [double]$Chk.threshold.maxLogonMs
                            $Chk.ActualValue = "$([math]::Round($numVal/1000,1))s"
                            if ($numVal -gt $limit) { $Chk.Status = 'Fail'; $Chk.Details = "Logon time exceeds $([int]($limit/1000))s limit" }
                            else { $Chk.Status = 'Pass' }
                            $thresholdHandled = $true
                        }
                    }
                    if (-not $thresholdHandled) {
                    # No baseline value — use heuristic evaluation
                    if ($ActualValue -is [bool]) {
                        $Chk.Status = if ($ActualValue) { 'Pass' } else { 'Fail' }
                    } elseif ($ActualValue -is [int] -or $ActualValue -is [long]) {
                        $Chk.Status = 'Warning'
                        $Chk.Details = "Value=$ActualValue (no baseline to compare)"
                    } elseif ($ActualValue -is [string] -and $ActualValue -eq '') {
                        $Chk.Status = 'Warning'
                        $Chk.Details = 'Not configured (empty)'
                    } elseif ($ActualValue -is [string] -and $ActualValue -match 'No Auditing') {
                        $Chk.Status = 'Fail'
                    } elseif ($ActualValue -is [string] -and $ActualValue -match 'Success') {
                        $Chk.Status = 'Pass'
                    } elseif ($ActualValue -is [System.Object[]] -or $ActualValue -is [System.Collections.ArrayList]) {
                        # Array-based checks (drivers.unsigned, scheduledTasks.failedTasks, event arrays, etc.)
                        $filteredArr = @($ActualValue)

                        # Filter by eventIds if specified in check definition
                        if ($Chk.eventIds -and $Chk.eventIds.Count -gt 0) {
                            $eids = @($Chk.eventIds)
                            $filteredArr = @($filteredArr | Where-Object { $_.id -in $eids })
                        }

                        # Filter by filterField/filterValues if specified (e.g., LOLBin process names)
                        if ($Chk.filterField -and $Chk.filterValues -and $Chk.filterValues.Count -gt 0) {
                            $fField  = $Chk.filterField
                            $fValues = @($Chk.filterValues)
                            $filteredArr = @($filteredArr | Where-Object {
                                $val = if ($fField -like 'props.*') { $_.props.($fField -replace '^props\.','') } else { $_.$fField }
                                if ($val) {
                                    $leaf = [System.IO.Path]::GetFileName($val)
                                    $leaf -in $fValues
                                }
                            })
                        }

                        $arrCount = $filteredArr.Count

                        # Check for count-based threshold (evalMode: count)
                        if ($Chk.threshold -and $Chk.threshold.evalMode -eq 'count' -and $null -ne $Chk.threshold.count) {
                            $limit = [int]$Chk.threshold.count
                            $Chk.ActualValue = "$arrCount events"
                            if ($arrCount -gt $limit) {
                                $Chk.Status = 'Fail'
                                $Chk.Details = "$arrCount events exceed threshold of $limit in $($Chk.threshold.windowDays) days"
                            } elseif ($arrCount -gt 0) {
                                $Chk.Status = 'Warning'
                                $Chk.Details = "$arrCount events (threshold: $limit)"
                            } else {
                                $Chk.Status = 'Pass'
                                $Chk.Details = 'No events found'
                            }
                        } else {
                            $Chk.ActualValue = "$arrCount items"
                            if ($arrCount -eq 0) {
                                $Chk.Status = 'Pass'
                                $Chk.Details = 'None found'
                            } else {
                                $Chk.Status = 'Fail'
                                # Build summary from first few items
                                $preview = @($ActualValue | Select-Object -First 3 | ForEach-Object {
                                    if ($_.Name) { $_.Name } elseif ($_.DeviceName) { $_.DeviceName } else { "$_" }
                                }) -join ', '
                                $Chk.Details = "$arrCount found: $preview$(if ($arrCount -gt 3) { '...' })"
                            }
                        }
                    } elseif ($ActualValue -is [PSCustomObject] -or $ActualValue -is [hashtable]) {
                        # Event log capacity check (evalMode: logCapacity)
                        if ($Chk.threshold -and $Chk.threshold.evalMode -eq 'logCapacity' -and $null -ne $ActualValue.FileSize -and $null -ne $ActualValue.MaxSizeKB) {
                            $fileSizeKB = [double]$ActualValue.FileSize
                            $maxSizeKB  = [double]$ActualValue.MaxSizeKB
                            $usagePct   = if ($maxSizeKB -gt 0) { [math]::Round(($fileSizeKB / $maxSizeKB) * 100, 1) } else { 0 }
                            $maxPct     = if ($Chk.threshold.maxUsagePct) { [double]$Chk.threshold.maxUsagePct } else { 90 }
                            $recordCount = if ($ActualValue.RecordCount) { [int]$ActualValue.RecordCount } else { 0 }
                            $Chk.ActualValue = "$usagePct% full ($recordCount records, $([math]::Round($fileSizeKB/1024))/$([math]::Round($maxSizeKB/1024)) MB)"
                            if ($usagePct -ge $maxPct) {
                                $Chk.Status = 'Warning'
                                $Chk.Details = "Log is ${usagePct}% full — risk of event loss. Overflow: $($ActualValue.OverflowAction)"
                            } else {
                                $Chk.Status = 'Pass'
                                $Chk.Details = "Overflow: $($ActualValue.OverflowAction)"
                            }
                        }
                        # Event data with threshold evaluation
                        elseif ($null -ne $ActualValue.count -and $Chk.threshold) {
                            $evtCount  = [int]$ActualValue.count
                            $threshold = [int]$Chk.threshold.count
                            $Chk.ActualValue = "$evtCount events"
                            if ($ActualValue.topUsers) {
                                $topList = ($ActualValue.topUsers | ForEach-Object { "$($_.user)($($_.count))" }) -join ', '
                                $Chk.Details = "Top: $topList"
                            }
                            if ($ActualValue.firstEvent) {
                                $span = "$($ActualValue.firstEvent.Substring(0,10)) – $($ActualValue.lastEvent.Substring(0,10))"
                                $Chk.Details = if ($Chk.Details) { "$($Chk.Details) | $span" } else { $span }
                            }
                            if ($evtCount -gt $threshold) {
                                $Chk.Status = 'Fail'
                            } elseif ($evtCount -gt 0) {
                                $Chk.Status = 'Warning'
                            } else {
                                $Chk.Status = 'Pass'
                            }
                        }
                        # Event data without threshold — report count, flag for review
                        elseif ($null -ne $ActualValue.count) {
                            $evtCount = [int]$ActualValue.count
                            $Chk.ActualValue = "$evtCount events"
                            if ($evtCount -gt 0) {
                                $Chk.Status = 'Warning'
                                $Chk.Details = "$evtCount events collected — no threshold defined"
                            } else {
                                $Chk.Status = 'Pass'
                            }
                        }
                        # Service or other complex object
                        else {
                            $svcStatus = $ActualValue.Status
                            $svcStart  = $ActualValue.StartType
                            $enabled   = $ActualValue.Enabled
                            if ($svcStatus) {
                                $Chk.ActualValue = "$svcStatus ($svcStart)"
                            } elseif (-not $svcStatus -and -not $enabled) {
                                # Whole-section object without Status/Enabled — summarize key properties
                                $props = @($ActualValue.PSObject.Properties | Where-Object { $_.Value -ne $null -and $_.Value -isnot [System.Object[]] -and $_.Value -isnot [PSCustomObject] } | Select-Object -First 6)
                                if ($props.Count -gt 0) {
                                    $summary = ($props | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join '; '
                                    $Chk.ActualValue = $summary
                                } else {
                                    $propCount = @($ActualValue.PSObject.Properties).Count
                                    $Chk.ActualValue = "$propCount properties collected"
                                }
                            }
                            if ($svcStatus -eq 'Running' -or $enabled -eq $true) { $Chk.Status = 'Pass' }
                            elseif ($svcStatus -eq 'Stopped' -or $enabled -eq $false) { $Chk.Status = 'Warning' }
                            else { $Chk.Status = 'Warning'; $Chk.Details = 'Collected but needs manual review' }
                        }
                    } else {
                        $Chk.Status = 'Warning'
                        $Chk.Details = 'Collected but needs manual review'
                    }
                    } # end if (-not $thresholdHandled)
                }
                $EvalCount++
                break  # Use first matching key
            }
        }

        # MDM precedence: if registry value is null but device is MDM-enrolled,
        # check if the corresponding CSP area is managed by Intune.
        # If so, the setting may be configured via MDM (different registry path).
        if ($Chk.Status -eq 'Not Assessed' -and $Json.mdmEnrollment.mdmEnrolled -and $Json.mdmEnrollment.managedAreas) {
            $cspArea = $null
            # Derive CSP area from cspPath (e.g., "Defender/AllowRealtimeMonitoring" -> "Defender")
            if ($Chk.CspPath) {
                $cspArea = ($Chk.CspPath -split '/')[0]
            }
            # Or from collectionKey to CSP area mapping
            if (-not $cspArea -and $Chk.CollectionKeys) {
                $keyHints = @{
                    'defender'   = 'Defender'
                    'firewall'   = 'WindowsFirewall'
                    'bitlocker'  = 'BitLocker'
                    'winrmConfig' = 'RemoteManagement'
                }
                foreach ($Key in $Chk.CollectionKeys) {
                    $section = ($Key -split '\.')[0]
                    if ($keyHints.ContainsKey($section)) { $cspArea = $keyHints[$section]; break }
                }
            }

            if ($cspArea -and $Json.mdmEnrollment.managedAreas.$cspArea) {
                $mdmArea = $Json.mdmEnrollment.managedAreas.$cspArea
                $Chk.Status = 'Warning'
                $Chk.ActualValue = 'Managed via Intune'
                $Chk.Details = "Setting managed by MDM ($($mdmArea.settingCount) settings in $cspArea area). GPO registry path is null — verify correct value in Intune portal."
                $EvalCount++
            }
        }

        # "Not configured" = Fail for baseline checks:
        # If the check is still Not Assessed, all collectionKeys resolved to null.
        # For checks with a baselineValue, null means "not configured" which means
        # the insecure default is in effect — this is a finding (Fail).
        # For checks without a baselineValue, null means we can't assess — leave as Not Assessed.
        if ($Chk.Status -eq 'Not Assessed' -and $null -ne $Chk.BaselineValue) {
            $Chk.Status = 'Fail'
            $Chk.ActualValue = 'Not configured'
            $Chk.Details = "Registry/policy value not found — setting uses insecure Windows default. Expected: $($Chk.BaselineValue)"
            $EvalCount++
        }
    }

    # Reset progress bar
    if ($barProgress) { $barProgress.Value = 0 }

    # Snapshot auto-evaluated status so users can Undo manual overrides
    foreach ($Chk in $Global:Assessment.Checks) {
        $Chk.AutoStatus = $Chk.Status
    }

    # Track affected machines for each finding (multi-machine support)
    foreach ($Chk in $Global:Assessment.Checks) {
        if ($Chk.Status -in @('Fail','Warning') -and $machineName -notin $Chk.AffectedMachines) {
            [void]$Chk.AffectedMachines.Add($machineName)
        }
    }

    $machineNote = if ($Global:Assessment.MachineCount -gt 1) { " ($($Global:Assessment.MachineCount) machines total)" } else { " (import more machines for aggregate view)" }
    Write-DebugLog "Collection imported: $machineName — $EvalCount checks evaluated, $SkipCount skipped$machineNote" -Level 'SUCCESS'
    Show-Toast "${machineName}: $EvalCount evaluated ($SkipCount N/A)$machineNote" -Type 'Success'

    Set-Dirty
    Update-Dashboard
    Render-BaselineChecks
    Render-Findings
    Unlock-Achievement 'first_import'
    Check-AssessmentAchievements
}

function Resolve-CollectionKey {
    param($Json, [string]$Key)
    # Split only on the FIRST dot to get the top-level section
    $dotIdx = $Key.IndexOf('.')
    if ($dotIdx -lt 0) {
        # No dot — direct top-level property
        try { return $Json.$Key } catch { return $null }
    }

    $section = $Key.Substring(0, $dotIdx)
    $remainder = $Key.Substring($dotIdx + 1)

    $sectionObj = $null
    try { $sectionObj = $Json.$section } catch { return $null }
    if ($null -eq $sectionObj) { return $null }

    # For registryBaselines, securityPolicy, auditPolicy — the remainder is a
    # literal property name (may contain spaces, backslashes, dots).
    # Use direct PSObject property lookup instead of recursive dot traversal.
    if ($section -in @('registryBaselines', 'securityPolicy', 'auditPolicy', 'tlsConfig')) {
        try {
            $prop = $sectionObj.PSObject.Properties[$remainder]
            if ($prop -and $null -ne $prop.Value) { return $prop.Value }
        } catch { }

        # Fallback for registryBaselines: try alternate path (Policies\ <-> Microsoft\)
        # CSP/MDM writes to SOFTWARE\Microsoft\ while GPO writes to SOFTWARE\Policies\Microsoft\
        # On hybrid devices, both can have values — MDMWinsOverGP determines precedence.
        if ($section -eq 'registryBaselines') {
            $altRemainder = $null
            $isGpoPath = $false
            if ($remainder -match '^HKLM\\SOFTWARE\\Policies\\Microsoft\\(.+)$') {
                $altRemainder = "HKLM\SOFTWARE\Microsoft\$($Matches[1])"
                $isGpoPath = $true
            } elseif ($remainder -match '^HKLM\\SOFTWARE\\Microsoft\\(.+)$' -and $remainder -notmatch 'CurrentVersion|PolicyManager') {
                $altRemainder = "HKLM\SOFTWARE\Policies\Microsoft\$($Matches[1])"
                $isGpoPath = $false
            }
            if ($altRemainder) {
                $altValue = $null
                try {
                    $altProp = $sectionObj.PSObject.Properties[$altRemainder]
                    if ($altProp -and $null -ne $altProp.Value) { $altValue = $altProp.Value }
                } catch { }

                # If primary had a value too, we have a conflict — respect MDMWinsOverGP
                $primaryValue = try { $sectionObj.PSObject.Properties[$remainder].Value } catch { $null }
                if ($null -ne $primaryValue -and $null -ne $altValue) {
                    # Both paths have values — pick based on precedence
                    $mdmWins = $Json.mdmEnrollment.mdmWinsOverGP -eq 1
                    if ($isGpoPath -and $mdmWins) {
                        return $altValue   # Check asked for GPO path but MDM wins — use CSP value
                    } elseif (-not $isGpoPath -and -not $mdmWins) {
                        return $altValue   # Check asked for CSP path but GPO wins — use GPO value
                    }
                    return $primaryValue   # Default: use the path the check asked for
                }
                # Only alternate has a value
                if ($null -ne $altValue) { return $altValue }
            }
            # Return primary value even if it's $null (explicit null is different from missing)
            try {
                $prop2 = $sectionObj.PSObject.Properties[$remainder]
                if ($prop2) { return $prop2.Value }
            } catch { }
        }
        return $null
    }

    # For other sections (defender, firewall, services, etc.) — navigate dot-separated
    $Parts = $remainder -split '\.'
    $Current = $sectionObj
    foreach ($Part in $Parts) {
        if ($null -eq $Current) { return $null }
        try { $Current = $Current.$Part } catch { return $null }
    }
    return $Current
}

# ===============================================================================
# SECTION 14: SAVE / LOAD ASSESSMENT
# ===============================================================================

function Sync-AssessmentFromUI {
    $Global:Assessment.CustomerName = $txtCustomerName.Text
    $Global:Assessment.AssessorName = $txtAssessorName.Text
    $Global:Assessment.Date         = $txtAssessmentDate.Text
    $Global:Assessment.Notes        = if ($txtNotes) { $txtNotes.Text } else { '' }
}

function Set-Dirty {
    if (-not $Global:IsDirty) {
        $Global:IsDirty = $true
        if ($dotDirty)    { $dotDirty.Visibility = 'Visible' }
        if ($lblDirtyText) { $lblDirtyText.Visibility = 'Visible' }
    }
}

function Clear-Dirty {
    $Global:IsDirty = $false
    if ($dotDirty)    { $dotDirty.Visibility = 'Collapsed' }
    if ($lblDirtyText) { $lblDirtyText.Visibility = 'Collapsed' }
}

function AutoSave-Assessment {
    if (-not $Global:AutoSaveEnabled) { return }
    try {
        Sync-AssessmentFromUI
        $JsonStr = $Global:Assessment | ConvertTo-Json -Depth 10

        [System.IO.File]::WriteAllText($Global:AutoSaveFile, $JsonStr, [System.Text.Encoding]::UTF8)

        $BackupDir = Join-Path $Global:Root '_backups'
        if (-not (Test-Path $BackupDir)) {
            New-Item -Path $BackupDir -ItemType Directory -Force | Out-Null
        }
        $BackupPath = Join-Path $BackupDir "backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
        [System.IO.File]::WriteAllText($BackupPath, $JsonStr, [System.Text.Encoding]::UTF8)

        $OldBackups = @(Get-ChildItem $BackupDir -Filter 'backup_*.json' -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -Skip $Global:MaxBackups)
        foreach ($Old in $OldBackups) {
            try { [System.IO.File]::Delete($Old.FullName) } catch { }
        }
    } catch { }
}

function Save-Assessment {
    Sync-AssessmentFromUI

    if (-not $Global:Assessment.CustomerName.Trim()) {
        Show-ThemedDialog `
            -Title 'Customer Name Required' `
            -Message 'Please enter a customer name before saving.' `
            -Icon ([char]0xE77B) `
            -IconColor 'ThemeWarning' `
            -Buttons @( @{ Text = 'OK'; IsAccent = $true; Result = 'OK' } )
        Switch-Tab 'Dashboard'
        $txtCustomerName.Focus()
        return
    }

    $NeedNewPath = (-not $Global:ActiveFilePath) -or (-not (Test-Path $Global:ActiveFilePath))
    if (-not $NeedNewPath -and $Global:Assessment.CustomerName.Trim() -and
        (Split-Path $Global:ActiveFilePath -Leaf) -match '^untitled_') {
        $NeedNewPath = $true
    }

    if ($NeedNewPath) {
        $CustSlug = ($Global:Assessment.CustomerName -replace '[^\w\-]','_').Trim('_')
        $Path = Join-Path $Global:AssessmentDir "${CustSlug}_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
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

function Test-AssessmentDirty {
    $Assessed = @($Global:Assessment.Checks | Where-Object { $_.Status -ne 'Not Assessed' }).Count
    $HasNotes = @($Global:Assessment.Checks | Where-Object { $_.Notes }).Count
    return ($Assessed -gt 0 -or $HasNotes -gt 0)
}

function Confirm-OverwriteAssessment {
    param([string]$Action = 'import')
    if (-not (Test-AssessmentDirty)) { return $true }

    $Assessed = @($Global:Assessment.Checks | Where-Object { $_.Status -ne 'Not Assessed' }).Count
    $Result = Show-ThemedDialog `
        -Title 'Unsaved Work' `
        -Message "You have $Assessed check(s) assessed.`n`nSave before ${Action}?" `
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

function Reset-Assessment {
    if (-not (Confirm-OverwriteAssessment -Action 'starting a new assessment')) { return }

    $Global:Assessment.CustomerName = ''
    $Global:Assessment.AssessorName = if ($txtDefaultAssessor) { $txtDefaultAssessor.Text } else { '' }
    $Global:Assessment.Date         = (Get-Date -Format 'yyyy-MM-dd')
    $Global:Assessment.Notes        = ''
    $Global:Assessment.CollectionData = $null
    $Global:Assessment.JoinType     = ''
    $Global:Assessment.ManualOverrides = @{}
    $Global:Assessment.Machines.Clear()
    $Global:Assessment.MachineCount = 0
    $Global:ActiveFilePath = $null

    $Global:Assessment.Checks.Clear()
    foreach ($Def in $Global:CheckDefinitions) {
        [void]$Global:Assessment.Checks.Add([PSCustomObject]@{
            Id=$Def.id;Category=$Def.category;Name=$Def.name;Description=$Def.description
            Severity=$Def.severity;Type=$Def.type;Weight=[int]$Def.weight;Priority=$Def.priority
            Origin=$Def.origin;Effort=$Def.effort;Impact=$Def.impact;Remediation=$Def.remediation
            Reference=$Def.reference;CollectionKeys=$Def.collectionKeys;BaselineValue=$Def.baselineValue
            CspPath=$Def.cspPath;Rationale=$Def.rationale;applicableTo=$Def.applicableTo;threshold=$Def.threshold
            eventIds=$Def.eventIds;filterField=$Def.filterField;filterValues=$Def.filterValues
            ActualValue=$null;Status='Not Assessed';AutoStatus=$null;Excluded=$false;Details='';Notes='';Source=$Def.type
            AffectedMachines=[System.Collections.ArrayList]::new()
        })
    }

    $txtCustomerName.Text   = ''
    $txtAssessorName.Text   = $Global:Assessment.AssessorName
    $txtAssessmentDate.Text = $Global:Assessment.Date
    $lblJoinType.Text       = 'Not Detected'
    $lblHeroHostname.Text   = '—'
    $lblHeroOS.Text         = '—'
    $lblHeroBuild.Text      = '—'
    $lblHeroJoinType.Text   = '—'
    $lblHeroCollected.Text  = '—'

    Clear-Dirty
    Update-Dashboard
    Render-BaselineChecks
    Render-Findings
    Switch-Tab 'Dashboard'
    Write-DebugLog "New assessment started" -Level 'INFO'
    Show-Toast "New assessment created" -Type 'Info'
}

function Load-Assessment {
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'JSON Files (*.json)|*.json'
    $dlg.InitialDirectory = $Global:AssessmentDir
    $dlg.Title = 'Load Assessment or Collection JSON'

    if ($dlg.ShowDialog() -eq $true) {
        if (-not (Confirm-OverwriteAssessment -Action 'loading a file')) { return }
        try {
            $Json = Get-Content $dlg.FileName -Raw -Encoding UTF8 | ConvertFrom-Json

            if ($Json._metadata -and $Json.systemInfo) {
                # Collection JSON — create fresh assessment, then import
                Reset-Assessment
                Import-CollectionJson -Path $dlg.FileName
            } elseif ($Json.Checks) {
                # Saved assessment — restore
                $Global:Assessment = $Json
                $Global:Assessment.Checks = [System.Collections.ArrayList]@($Json.Checks)
                # Merge new checks from checks.json
                $ExIds = @($Json.Checks | ForEach-Object { $_.Id })
                foreach ($Def in $Global:CheckDefinitions) {
                    if ($Def.id -notin $ExIds) {
                        [void]$Global:Assessment.Checks.Add([PSCustomObject]@{
                            Id=$Def.id;Category=$Def.category;Name=$Def.name;Description=$Def.description
                            Severity=$Def.severity;Type=$Def.type;Weight=[int]$Def.weight;Priority=$Def.priority
                            Origin=$Def.origin;Effort=$Def.effort;Impact=$Def.impact;Remediation=$Def.remediation
                            Reference=$Def.reference;CollectionKeys=$Def.collectionKeys;BaselineValue=$Def.baselineValue
                            CspPath=$Def.cspPath;Rationale=$Def.rationale
                            applicableTo=$Def.applicableTo;threshold=$Def.threshold
                            eventIds=$Def.eventIds;filterField=$Def.filterField;filterValues=$Def.filterValues
                            ActualValue=$null;Status='Not Assessed';AutoStatus=$null;Excluded=$false;Details='';Notes='';Source=$Def.type
                            AffectedMachines=[System.Collections.ArrayList]::new()
                        })
                    }
                }
                $txtCustomerName.Text   = $Json.CustomerName
                $txtAssessorName.Text   = $Json.AssessorName
                $txtAssessmentDate.Text = $Json.Date
                if ($txtNotes) { $txtNotes.Text = $Json.Notes }

                Update-Dashboard
                Update-Progress
                Render-BaselineChecks
                Render-Findings
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

function Refresh-AssessmentList {
    $lstSavedAssessments.Items.Clear()

    if (-not (Test-Path $Global:AssessmentDir)) {
        $lblSavedEmpty.Visibility = 'Visible'
        $lblSavedCount.Text = ''
        return
    }

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
        $DisplayName  = $F.BaseName
        $Subtitle     = $F.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
        $ScorePercent = -1

        try {
            $Peek = Get-Content $F.FullName -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json
            if ($Peek._metadata -and $Peek.systemInfo) { continue }  # Skip collection files
            if ($Peek.Checks) {
                $DisplayName = if ($Peek.CustomerName) { $Peek.CustomerName } else { 'Untitled Assessment' }
                $Assessed = @($Peek.Checks | Where-Object { $_.Status -ne 'Not Assessed' }).Count
                if ($Assessed -gt 0) {
                    $ScoringChecks = @($Peek.Checks | Where-Object { $_.Status -ne 'N/A' })
                    $WSum = 0; $TW = 0
                    foreach ($SC in $ScoringChecks) {
                        $W = [math]::Max(1, [int]$SC.Weight)
                        $Pts = switch ($SC.Status) { 'Pass' { 100 } 'Warning' { 50 } default { 0 } }
                        $WSum += $Pts * $W; $TW += $W
                    }
                    if ($TW -gt 0) { $ScorePercent = [math]::Round($WSum / $TW, 0) }
                }
            }
        } catch { }

        $Item = New-Object System.Windows.Controls.ListViewItem
        $Item.Tag = $F.FullName
        $Item.Cursor = [System.Windows.Input.Cursors]::Hand

        $SP = New-Object System.Windows.Controls.StackPanel
        $Row1 = New-Object System.Windows.Controls.DockPanel

        $NameTB = New-Object System.Windows.Controls.TextBlock
        $NameTB.Text = $DisplayName; $NameTB.FontSize = 12; $NameTB.FontWeight = [System.Windows.FontWeights]::SemiBold
        $NameTB.TextTrimming = 'CharacterEllipsis'; $NameTB.MaxWidth = 160
        $NameTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')

        $SubTB = New-Object System.Windows.Controls.TextBlock
        $SubTB.Text = $Subtitle; $SubTB.FontSize = 11
        $SubTB.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextFaintest')

        if ($ScorePercent -ge 0) {
            $ScoreTB = New-Object System.Windows.Controls.TextBlock
            $ScoreTB.Text = "$ScorePercent%"; $ScoreTB.FontSize = 12; $ScoreTB.FontWeight = [System.Windows.FontWeights]::Bold
            $ScoreTB.VerticalAlignment = 'Center'
            $ScoreTB.Foreground = if ($ScorePercent -ge 80) { $Window.Resources['ThemeSuccess'] }
                                  elseif ($ScorePercent -ge 50) { $Window.Resources['ThemeWarning'] }
                                  else { $Window.Resources['ThemeError'] }
            [System.Windows.Controls.DockPanel]::SetDock($ScoreTB, 'Right')
            [void]$Row1.Children.Add($ScoreTB)
        }

        $NameSP = New-Object System.Windows.Controls.StackPanel
        [void]$NameSP.Children.Add($NameTB)
        [void]$NameSP.Children.Add($SubTB)
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

function Load-UserPrefs {
    if (-not (Test-Path $Global:PrefsPath)) { return }
    try {
        $P = Get-Content $Global:PrefsPath -Raw -Encoding UTF8 | ConvertFrom-Json

        if ($null -ne $P.IsLightMode -and $P.IsLightMode -ne $Global:IsLightMode) {
            $Global:SuppressThemeHandler = $true
            $Global:IsLightMode = $P.IsLightMode
            & $ApplyTheme $P.IsLightMode
            $chkDarkMode.IsChecked = -not $P.IsLightMode
            $btnThemeToggle.Content = if ($P.IsLightMode) { [char]0x263D } else { [char]0x2600 }
            $Global:SuppressThemeHandler = $false
        }
        if ($P.DefaultAssessor) { $txtDefaultAssessor.Text = $P.DefaultAssessor }
        if ($P.ExportPath)      { $Global:ReportDir = $P.ExportPath; $txtExportPath.Text = $P.ExportPath }
        if ($null -ne $P.AutoSaveEnabled)  { $Global:AutoSaveEnabled = $P.AutoSaveEnabled; $chkAutoSave.IsChecked = $P.AutoSaveEnabled }
        if ($null -ne $P.AutoSaveInterval) { $Global:AutoSaveInterval = $P.AutoSaveInterval }
        if ($null -ne $P.OpenAfterExport)  { $Global:OpenAfterExport = $P.OpenAfterExport; $chkOpenAfterExport.IsChecked = $P.OpenAfterExport }
        if ($null -ne $P.VerboseLogging)   { $Global:VerboseLogging = $P.VerboseLogging; $chkVerboseLog.IsChecked = $P.VerboseLogging }
        if ($null -ne $P.MaxBackups)       { $Global:MaxBackups = $P.MaxBackups }
        if ($P.Achievements) {
            $P.Achievements.PSObject.Properties | ForEach-Object {
                $Global:Achievements[$_.Name] = $_.Value
            }
        }
        if ($null -ne $P.ActivityLogOpen) { Toggle-ActivityLog $P.ActivityLogOpen }
        if ($P.AchievementsCollapsed) {
            $Window.FindName('pnlAchievements').Visibility = 'Collapsed'
            $Window.FindName('lblAchievementChevron').Text = [char]0xE76C
        }

        # Restore window position
        if ($P.WindowState -eq 'Normal' -and $P.WindowWidth -gt 100 -and $P.WindowHeight -gt 100) {
            $Window.WindowState = 'Normal'
            $Window.Left   = $P.WindowLeft
            $Window.Top    = $P.WindowTop
            $Window.Width  = $P.WindowWidth
            $Window.Height = $P.WindowHeight
            # Verify on-screen
            $VW = [System.Windows.SystemParameters]::VirtualScreenWidth
            $VH = [System.Windows.SystemParameters]::VirtualScreenHeight
            if ($Window.Left -lt -100 -or $Window.Left -gt $VW -or $Window.Top -lt -100 -or $Window.Top -gt $VH) {
                $Window.WindowStartupLocation = 'CenterScreen'
                Write-DebugLog "Saved window position off-screen, centering" -Level 'DEBUG'
            }
        }

        Write-DebugLog "User preferences loaded" -Level 'DEBUG'
    } catch {
        Write-DebugLog "Failed to load preferences: $($_.Exception.Message)" -Level 'WARN'
    }
}

# ===============================================================================
# SECTION 15B: ACHIEVEMENTS
# ===============================================================================

$Global:AchievementDefs = @(
    @{ Id='first_import';       Name='Data Collector';   Icon='📥'; Desc='Import your first collection file' }
    @{ Id='first_save';         Name='First Steps';      Icon='🎯'; Desc='Save your first assessment' }
    @{ Id='five_saves';         Name='Repeat Auditor';   Icon='📋'; Desc='Save 5 assessments' }
    @{ Id='ten_saves';          Name='Assessment Pro';   Icon='🏆'; Desc='Save 10 assessments' }
    @{ Id='full_sweep';         Name='Full Sweep';       Icon='🧹'; Desc='Evaluate all checks in a single run' }
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
    @{ Id='zero_critical';      Name='Zero Critical';    Icon='🛡'; Desc='No critical-severity failures' }
    @{ Id='fifty_pass';         Name='Half Way';         Icon='⭐'; Desc='Pass 50 checks in a single assessment' }
    @{ Id='hundred_pass';       Name='Century Club';     Icon='💯'; Desc='Pass 100 checks' }
    @{ Id='baseline_hero';      Name='Baseline Hero';    Icon='🏅'; Desc='Achieve 90%+ baseline compliance' }
    @{ Id='speed_demon';        Name='Speed Demon';      Icon='⚡'; Desc='Complete assessment in under 5 minutes' }
)

$Global:Achievements = @{}

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
        $Badge.Width = 26; $Badge.Height = 26
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
        $Lbl = New-Object System.Windows.Controls.TextBlock
        $Lbl.Text = if ($IsUnlocked) { $Def.Icon } else { '?' }
        $Lbl.FontSize = 12
        $Lbl.HorizontalAlignment = 'Center'
        $Lbl.VerticalAlignment = 'Center'
        if (-not $IsUnlocked) {
            $Lbl.Foreground = $Global:CachedBC.ConvertFromString($(if ($Global:IsLightMode) { '#AAAAAA' } else { '#555555' }))
        }
        $Badge.Child = $Lbl
        [void]$pnl.Children.Add($Badge)
    }
    if ($lbl) { $lbl.Text = "$Unlocked/$($Global:AchievementDefs.Count)" }
}

function Check-AssessmentAchievements {
    $Checks   = $Global:Assessment.Checks
    $Assessed = @($Checks | Where-Object { $_.Status -ne 'Not Assessed' -and $_.Status -ne 'N/A' })
    $Passed   = @($Assessed | Where-Object { $_.Status -eq 'Pass' })
    $Failed   = @($Assessed | Where-Object { $_.Status -eq 'Fail' })

    if ($Assessed.Count -gt 0) { Unlock-Achievement 'first_save' }

    $SavedFiles = @(Get-ChildItem -Path $Global:AssessmentDir -Filter '*.json' -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notmatch '_autosave' })
    if ($SavedFiles.Count -ge 5)  { Unlock-Achievement 'five_saves' }
    if ($SavedFiles.Count -ge 10) { Unlock-Achievement 'ten_saves' }

    $NotAssessed = @($Checks | Where-Object { $_.Status -eq 'Not Assessed' -and -not $_.Excluded })
    if ($NotAssessed.Count -eq 0 -and $Checks.Count -gt 0) { Unlock-Achievement 'full_sweep' }

    $Categories = @($Checks | ForEach-Object { $_.Category } | Sort-Object -Unique)
    foreach ($Cat in $Categories) {
        $CatChecks = @($Assessed | Where-Object { $_.Category -eq $Cat })
        if ($CatChecks.Count -ge 3 -and @($CatChecks | Where-Object { $_.Status -ne 'Pass' }).Count -eq 0) {
            Unlock-Achievement 'perfect_category'
            break
        }
    }

    $OverallScore = Get-OverallScore
    if ($OverallScore -ge 75) { Unlock-Achievement 'maturity_managed' }
    if ($OverallScore -ge 90) { Unlock-Achievement 'maturity_optimized' }

    $Compliance = Get-BaselineCompliancePercent
    if ($Compliance -ge 90) { Unlock-Achievement 'baseline_hero' }

    $CritFail = @($Failed | Where-Object { $_.Severity -eq 'Critical' })
    if ($Assessed.Count -ge 10 -and $CritFail.Count -eq 0) { Unlock-Achievement 'zero_critical' }

    if ($Passed.Count -ge 50)  { Unlock-Achievement 'fifty_pass' }
    if ($Passed.Count -ge 100) { Unlock-Achievement 'hundred_pass' }

    $WithNotes = @($Checks | Where-Object { $_.Notes }).Count
    if ($WithNotes -ge 5) { Unlock-Achievement 'note_taker' }

    $Hour = (Get-Date).Hour
    if ($Hour -ge 0 -and $Hour -lt 5)  { Unlock-Achievement 'night_owl' }
    if ($Hour -ge 5 -and $Hour -lt 7)  { Unlock-Achievement 'early_bird' }
    if ((Get-Date).DayOfWeek -in @('Saturday','Sunday')) { Unlock-Achievement 'weekend_warrior' }
}

# ===============================================================================
# SECTION 16: EXPORT FUNCTIONS
# ===============================================================================

function Export-CsvReport {
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'CSV Files (*.csv)|*.csv'
    $dlg.FileName = "BaselinePilot_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $dlg.InitialDirectory = $Global:ReportDir

    if ($dlg.ShowDialog() -eq $true) {
        try {
            $Global:Assessment.Checks | Select-Object Id, Category, Name, Description, Status, Severity, Weight, Priority, Origin, Effort, Excluded, Details, Notes |
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

function Export-JsonAssessment {
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'JSON Files (*.json)|*.json'
    $dlg.FileName = "BaselinePilot_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
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

function Export-HtmlReport {
    Sync-AssessmentFromUI
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'HTML Files (*.html)|*.html'
    $dlg.FileName = "BaselinePilot_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    $dlg.InitialDirectory = $Global:ReportDir

    if ($dlg.ShowDialog() -ne $true) { return }

    try {
        $enc = { param($s) [System.Net.WebUtility]::HtmlEncode("$s") }

        $Score      = Get-OverallScore
        $Compliance = Get-BaselineCompliancePercent
        $All    = $Global:Assessment.Checks
        $Pass   = @($All | Where-Object { $_.Status -eq 'Pass' }).Count
        $Fail   = @($All | Where-Object { $_.Status -eq 'Fail' }).Count
        $Warn   = @($All | Where-Object { $_.Status -eq 'Warning' }).Count
        $NA     = @($All | Where-Object { $_.Status -eq 'N/A' }).Count
        $NotA   = @($All | Where-Object { $_.Status -eq 'Not Assessed' }).Count
        $AccRisk = @($All | Where-Object { $_.Status -eq 'Accepted Risk' }).Count
        $Defer  = @($All | Where-Object { $_.Status -eq 'Deferred' }).Count
        $Total  = $All.Count
        $Assessed = $Total - $NotA

        $ScoreClass = if ($Compliance -ge 80) { 'green' } elseif ($Compliance -ge 50) { 'orange' } else { 'red' }
        $RiskClass  = if ($Score -ge 80) { 'green' } elseif ($Score -ge 50) { 'orange' } else { 'red' }

        # Progress bar percentages
        $PctPass = if ($Assessed -gt 0) { [math]::Round(($Pass / $Assessed) * 100, 1) } else { 0 }
        $PctFail = if ($Assessed -gt 0) { [math]::Round(($Fail / $Assessed) * 100, 1) } else { 0 }
        $PctWarn = if ($Assessed -gt 0) { [math]::Round(($Warn / $Assessed) * 100, 1) } else { 0 }

        # Effort counts
        $EffortQuick = @($All | Where-Object { $_.Status -eq 'Fail' -and $_.Effort -eq 'Quick Win' }).Count
        $EffortSome  = @($All | Where-Object { $_.Status -eq 'Fail' -and $_.Effort -eq 'Some Effort' }).Count
        $EffortMajor = @($All | Where-Object { $_.Status -eq 'Fail' -and $_.Effort -eq 'Major Effort' }).Count

        # Category sections
        $CatSections = ''
        foreach ($Cat in (Get-Categories)) {
            $CatChecks = @($All | Where-Object { $_.Category -eq $Cat })
            $CatPass = @($CatChecks | Where-Object { $_.Status -eq 'Pass' }).Count
            $CatFail = @($CatChecks | Where-Object { $_.Status -eq 'Fail' }).Count
            $CatWarn = @($CatChecks | Where-Object { $_.Status -eq 'Warning' }).Count
            $CatScore = Get-CategoryScore $Cat
            $CatColor = if ($CatScore -ge 80) { 'var(--green)' } elseif ($CatScore -ge 50) { 'var(--orange)' } else { 'var(--red)' }
            $CatScoreText = if ($CatScore -ge 0) { "$CatScore%" } else { [string][char]0x2014 }
            $CatScoreClass = if ($CatScore -ge 80) { 'green' } elseif ($CatScore -ge 50) { 'orange' } else { 'red' }
            $CatPct = if ($CatChecks.Count -gt 0) { [math]::Round(($CatPass / $CatChecks.Count) * 100) } else { 0 }

            # Check rows
            $CheckRows = ''
            foreach ($Chk in ($CatChecks | Sort-Object { @{Critical=0;High=1;Medium=2;Low=3}[$_.Severity] })) {
                $StatusBadge = switch ($Chk.Status) {
                    'Pass'          { "<span class='status-badge pass'>Pass</span>" }
                    'Fail'          { "<span class='status-badge fail'>Fail</span>" }
                    'Warning'       { "<span class='status-badge warning'>Warn</span>" }
                    'N/A'           { "<span class='status-badge na'>N/A</span>" }
                    'Accepted Risk' { "<span class='status-badge warning' style='opacity:0.8'>Risk Accepted</span>" }
                    'Deferred'      { "<span class='status-badge' style='background:var(--accent);color:#fff'>Deferred</span>" }
                    default         { "<span class='status-badge notassessed'>$([char]0x2014)</span>" }
                }
                $SevClass = switch ($Chk.Severity) { 'Critical' { 'sev-critical' } 'High' { 'sev-high' } 'Medium' { 'sev-medium' } default { 'sev-low' } }
                $ExpVal = if ($null -ne $Chk.BaselineValue) { & $enc "$($Chk.BaselineValue)" } else { [string][char]0x2014 }
                $ActVal = if ($null -ne $Chk.ActualValue) { & $enc "$($Chk.ActualValue)" } else { [string][char]0x2014 }

                # Weight dots
                $Wt = [math]::Min(5, [math]::Max(1, [int]$Chk.Weight))
                $WtDots = ("$([char]0x25CF)" * $Wt) + ("$([char]0x25CB)" * (5 - $Wt))

                # Origin badges
                $OriginBadges = ''
                if ($Chk.Origin) {
                    foreach ($o in ($Chk.Origin -split ',')) {
                        $o = $o.Trim()
                        $oClass = switch ($o) { 'INTUNE' { 'waf' } 'SCT' { 'caf' } 'OPS' { 'lza' } default { 'avd' } }
                        $OriginBadges += "<span class='origin-badge $oClass'>$o</span>"
                    }
                }

                # Detail column: impact + remediation for failures
                $DetailHtml = ''
                if ($Chk.Rationale -and $Chk.Status -in @('Fail','Warning')) {
                    $DetailHtml += "<div style='margin-bottom:4px;padding:4px 8px;background:var(--orange-dim);border-radius:var(--radius-xs);font-size:10px;color:var(--text-secondary)'><strong style='color:var(--orange-text)'>Why:</strong> $(& $enc $Chk.Rationale)</div>"
                }
                if ($Chk.Status -in @('Fail','Warning')) {
                    if ($Chk.Impact) {
                        $DetailHtml += "<div class='impact-box'><strong>Impact:</strong> $(& $enc $Chk.Impact)</div>"
                    }
                    if ($Chk.Remediation) {
                        $DetailHtml += "<div class='remediation-box'><strong>Fix:</strong> $(& $enc $Chk.Remediation)</div>"
                    }
                }
                if ($Chk.Notes) {
                    $DetailHtml += "<div style='margin-top:4px;font-size:10px;color:var(--text-dim);font-style:italic'>Note: $(& $enc $Chk.Notes)</div>"
                }
                if ($Chk.Details -and $Chk.Details -in @('Remediate','AcceptRisk','Defer')) {
                    $DecColor = switch ($Chk.Details) { 'Remediate' { 'var(--green-text)' } 'AcceptRisk' { 'var(--orange-text)' } default { 'var(--text-dim)' } }
                    $DetailHtml += "<div style='margin-top:4px;font-size:10px'><span style='color:$DecColor;font-weight:600'>Decision: $($Chk.Details)</span></div>"
                }

                $RefLink = if ($Chk.Reference) { "<a href='$(& $enc $Chk.Reference)' target='_blank' style='color:var(--accent-text);text-decoration:none;font-size:10px'>Docs</a>" } else { '' }

                $RowClass = if ($Chk.Status -eq 'Not Assessed' -or $Chk.Excluded) { " class='excluded-row'" } else { '' }

                $CheckRows += @"
<tr$RowClass data-search="$(& $enc "$($Chk.Id) $($Chk.Name) $($Chk.Description)")">
  <td style='font-weight:600;color:var(--text-dim);font-size:11px'>$($Chk.Id)</td>
  <td><strong>$(& $enc $Chk.Name)</strong><br><small style='color:var(--text-dim)'>$(& $enc $Chk.Description)</small></td>
  <td>$ExpVal</td>
  <td>$ActVal</td>
  <td>$StatusBadge</td>
  <td><span class='sev $SevClass'>$($Chk.Severity)</span></td>
  <td><span class='wt'>$WtDots</span></td>
  <td>$OriginBadges</td>
  <td>$DetailHtml</td>
  <td>$RefLink</td>
</tr>
"@
            }

            $CatSections += @"
<div class='section-header' onclick="this.className=this.className.indexOf('open')>=0?'section-header':'section-header open';var c=this.nextElementSibling;c.className=c.className.indexOf('open')>=0?'section-content':'section-content open'">
  <span class='chevron'>&#9658;</span>
  <span class='section-dot' style='background:$CatColor'></span>
  $(& $enc $Cat)
  <span class='section-score' style='color:$CatColor'>$CatScoreText</span>
  <span class='section-count'>$CatPass pass / $CatFail fail / $($CatChecks.Count) total</span>
</div>
<div class='section-content'>
  <div class='cat-progress'><div class='cat-progress-fill' style='width:${CatPct}%;background:$CatColor'></div></div>
  <table class='check-table'>
    <thead><tr><th>ID</th><th>Check</th><th>Expected</th><th>Actual</th><th>Status</th><th>Severity</th><th>Wt</th><th>Origin</th><th>Details</th><th>Ref</th></tr></thead>
    <tbody>$CheckRows</tbody>
  </table>
</div>
"@
        }

        # Maturity dimension cards
        $DimCards = ''
        foreach ($DimKey in $Global:MaturityDimensions.Keys) {
            $Dim = $Global:MaturityDimensions[$DimKey]
            $DScore = Get-DimensionScore $DimKey
            if ($DScore -lt 0) { $DScore = 0 }
            $DLevel = Get-MaturityLevel $DScore
            $DColor = if ($DScore -ge 75) { 'var(--green)' } elseif ($DScore -ge 35) { 'var(--orange)' } else { 'var(--red)' }
            $DClass = if ($DScore -ge 75) { 'green' } elseif ($DScore -ge 35) { 'orange' } else { 'red' }
            $DimCards += @"
<div class='dim-card'>
  <div class='dim-header'><span class='dim-label'>$(& $enc $Dim.Label)</span><span class='dim-score $DClass'>$DScore%</span></div>
  <div class='dim-bar'><div class='dim-bar-fill' style='width:${DScore}%;background:$DColor'></div></div>
  <div class='dim-meta'><span>$DLevel</span></div>
</div>
"@
        }

        # System info
        $SysHost = if ($Global:Assessment.CollectionData.systemInfo.hostname) { & $enc $Global:Assessment.CollectionData.systemInfo.hostname } else { [string][char]0x2014 }
        $SysOS   = if ($Global:Assessment.CollectionData.systemInfo.OSCaption) { & $enc $Global:Assessment.CollectionData.systemInfo.OSCaption } else { [string][char]0x2014 }
        $SysBuild = if ($Global:Assessment.CollectionData.systemInfo.OSBuild) { & $enc $Global:Assessment.CollectionData.systemInfo.OSBuild } else { [string][char]0x2014 }

        $Html = @"
<!DOCTYPE html>
<html lang="en" data-theme="dark">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>BaselinePilot Report$(if($Global:Assessment.CustomerName){" - $(& $enc $Global:Assessment.CustomerName)"})</title>
<style>
:root{--font:'Inter','Segoe UI Variable Display','Segoe UI',sans-serif;--bg:#0a0a0c;--bg2:#111113;--surface:#141416;--card:#1a1a1e;--card-hover:#222226;--card-border:#2a2a2e;--border:rgba(255,255,255,0.06);--border-strong:rgba(255,255,255,0.12);--text:#e4e4e7;--text-bright:#fff;--text-secondary:#a1a1aa;--text-dim:#71717a;--text-faint:#52525b;--accent:#60cdff;--accent-dim:rgba(96,205,255,0.1);--accent-text:#60cdff;--green:#00c853;--green-dim:rgba(0,200,83,0.12);--green-text:#4ade80;--red:#ff5000;--red-dim:rgba(255,80,0,0.12);--red-text:#fb923c;--orange:#f59e0b;--orange-dim:rgba(245,158,11,0.12);--orange-text:#fbbf24;--radius:12px;--radius-sm:8px;--radius-xs:6px;--shadow:0 1px 3px rgba(0,0,0,0.3)}
[data-theme="light"]{--bg:#f8f9fa;--bg2:#fff;--surface:#f4f4f5;--card:#fff;--card-hover:#f0f0f2;--card-border:#e4e4e7;--border:rgba(0,0,0,0.08);--border-strong:rgba(0,0,0,0.15);--text:#18181b;--text-bright:#09090b;--text-secondary:#52525b;--text-dim:#71717a;--text-faint:#a1a1aa;--accent:#0078d4;--accent-dim:rgba(0,120,212,0.08);--accent-text:#0066b8;--green:#16a34a;--green-dim:rgba(22,163,74,0.08);--green-text:#15803d;--red:#dc2626;--red-dim:rgba(220,38,38,0.08);--red-text:#b91c1c;--orange:#ea580c;--orange-dim:rgba(234,88,12,0.08);--orange-text:#c2410c;--shadow:0 1px 3px rgba(0,0,0,0.06)}
*,*::before,*::after{margin:0;padding:0;box-sizing:border-box}html{scroll-behavior:smooth}
body{font-family:var(--font);background:var(--bg);color:var(--text);line-height:1.5}
.toolbar{position:sticky;top:0;z-index:100;background:var(--bg2);border-bottom:1px solid var(--border);padding:10px 32px;display:flex;align-items:center;gap:12px;backdrop-filter:blur(12px)}
.toolbar h1{font-size:15px;font-weight:600;color:var(--text-bright);white-space:nowrap;display:flex;align-items:center;gap:8px}
.toolbar .spacer{flex:1}
.toolbar input{background:var(--card);border:1px solid var(--card-border);border-radius:var(--radius-sm);padding:5px 12px;color:var(--text);font-size:11px;font-family:var(--font);width:200px;outline:none}
.toolbar input:focus{border-color:var(--accent)}
.btn{background:var(--card);border:1px solid var(--card-border);border-radius:var(--radius-sm);padding:5px 12px;color:var(--text-secondary);font-size:11px;cursor:pointer;font-family:var(--font);transition:all .15s}
.btn:hover{background:var(--card-hover);color:var(--text)}
.container{max-width:1400px;margin:0 auto;padding:24px 40px 60px}
.meta-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:6px;margin-bottom:24px;background:var(--card);border:1px solid var(--card-border);border-radius:var(--radius);padding:14px 20px}
.meta-item{font-size:11px;color:var(--text-dim)}.meta-item span{font-weight:600;color:var(--text)}
.score-hero{text-align:center;padding:28px 0 24px;margin-bottom:24px;border-bottom:1px solid var(--border)}
.score-hero .score-row{display:flex;justify-content:center;gap:48px;align-items:flex-end}
.score-hero .score-block .val{font-size:56px;font-weight:800;letter-spacing:-2px;line-height:1}
.score-hero .score-block .lbl{font-size:11px;color:var(--text-dim);margin-top:2px;text-transform:uppercase;letter-spacing:1px}
.green{color:var(--green)}.orange{color:var(--orange)}.red{color:var(--red)}.accent{color:var(--accent)}
.summary{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:28px}
.stat-card{background:var(--card);border:1px solid var(--card-border);border-radius:var(--radius);padding:16px 20px;position:relative;overflow:hidden;transition:border-color .2s}
.stat-card:hover{border-color:var(--border-strong)}
.stat-card .accent-bar{position:absolute;top:0;left:0;right:0;height:3px;border-radius:var(--radius) var(--radius) 0 0}
.stat-card .num{font-size:28px;font-weight:800;letter-spacing:-1px;line-height:1}.stat-card .label{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.5px;color:var(--text-dim);margin-top:4px}
.stat-card.pass .num{color:var(--green)}.stat-card.pass .accent-bar{background:var(--green)}
.stat-card.warn .num{color:var(--orange)}.stat-card.warn .accent-bar{background:var(--orange)}
.stat-card.fail .num{color:var(--red)}.stat-card.fail .accent-bar{background:var(--red)}
.stat-card.total .num{color:var(--accent)}.stat-card.total .accent-bar{background:var(--accent)}
.stat-card.na .num{color:var(--text-dim)}.stat-card.na .accent-bar{background:var(--text-faint)}
.progress-bar{display:flex;height:8px;border-radius:4px;overflow:hidden;margin-bottom:28px;background:var(--card);border:1px solid var(--card-border)}
.seg{transition:width .6s ease}.seg-pass{background:var(--green)}.seg-warn{background:var(--orange)}.seg-fail{background:var(--red)}
.effort-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:28px}
.effort-card{background:var(--card);border:1px solid var(--card-border);border-radius:var(--radius);padding:16px;text-align:center}
.effort-card .effort-icon{font-size:22px;margin-bottom:2px}.effort-card .effort-count{font-size:24px;font-weight:800;letter-spacing:-1px}
.effort-card .effort-label{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.5px;color:var(--text-dim);margin-top:2px}
.effort-card.quick .effort-count{color:var(--green-text)}.effort-card.some .effort-count{color:var(--orange-text)}.effort-card.major .effort-count{color:var(--red-text)}
.maturity-dims{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:28px}
.dim-card{background:var(--card);border:1px solid var(--card-border);border-radius:var(--radius-sm);padding:14px 18px;transition:border-color .2s}
.dim-card:hover{border-color:var(--border-strong)}
.dim-card .dim-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:6px}
.dim-card .dim-label{font-size:12px;font-weight:600;color:var(--text)}.dim-card .dim-score{font-size:18px;font-weight:800;letter-spacing:-1px}
.dim-card .dim-bar{height:4px;border-radius:2px;background:var(--surface);overflow:hidden}.dim-card .dim-bar-fill{height:100%;border-radius:2px;transition:width .4s ease}
.dim-card .dim-meta{margin-top:4px;font-size:10px;color:var(--text-faint)}
.section-header{display:flex;align-items:center;gap:10px;padding:12px 18px;cursor:pointer;background:var(--card);border:1px solid var(--card-border);border-radius:var(--radius);margin-bottom:2px;font-size:13px;font-weight:600;color:var(--text);user-select:none;transition:border-color .2s}
.section-header:hover{border-color:var(--border-strong)}
.section-header .chevron{font-size:10px;color:var(--text-dim);transition:transform .2s}.section-header.open .chevron{transform:rotate(90deg)}
.section-header .section-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0}
.section-header .section-score{font-size:12px;font-weight:500;margin-left:auto}.section-header .section-count{font-size:10px;color:var(--text-faint);background:var(--surface);border:1px solid var(--card-border);border-radius:12px;padding:1px 8px}
.section-content{display:none;margin-bottom:12px}.section-content.open{display:block}
.cat-progress{height:4px;border-radius:2px;background:var(--surface);margin:6px 18px 0;overflow:hidden}.cat-progress-fill{height:100%;border-radius:2px}
.check-table{width:100%;border-collapse:separate;border-spacing:0;table-layout:fixed}
.check-table thead th{padding:7px 12px;font-size:9px;font-weight:600;text-transform:uppercase;letter-spacing:.8px;color:var(--text-faint);border-bottom:1px solid var(--border);text-align:left;background:var(--surface);position:sticky;top:44px}
.check-table thead th:nth-child(1){width:65px}.check-table thead th:nth-child(2){width:auto}.check-table thead th:nth-child(3){width:80px}.check-table thead th:nth-child(4){width:80px}
.check-table thead th:nth-child(5){width:55px}.check-table thead th:nth-child(6){width:60px}.check-table thead th:nth-child(7){width:40px}.check-table thead th:nth-child(8){width:60px}.check-table thead th:nth-child(9){width:auto}.check-table thead th:nth-child(10){width:35px}
.check-table tbody tr{transition:background .1s}.check-table tbody tr:hover{background:var(--card-hover)}.check-table td{padding:8px 12px;font-size:11px;border-bottom:1px solid var(--border);vertical-align:top;word-wrap:break-word}
.status-badge{display:inline-flex;align-items:center;gap:4px;padding:2px 8px;border-radius:20px;font-size:9px;font-weight:700;letter-spacing:.3px;text-transform:uppercase}
.status-badge.pass{background:var(--green-dim);color:var(--green-text)}.status-badge.warning{background:var(--orange-dim);color:var(--orange-text)}.status-badge.fail{background:var(--red-dim);color:var(--red-text)}.status-badge.na,.status-badge.notassessed{background:var(--surface);color:var(--text-faint)}
.sev{font-size:10px;font-weight:600}.sev-critical{color:var(--red-text)}.sev-high{color:var(--orange-text)}.sev-medium{color:var(--text-secondary)}.sev-low{color:var(--text-faint)}
.wt{font-size:8px;letter-spacing:1px;color:var(--text-dim)}
.origin-badge{display:inline-block;padding:1px 5px;border-radius:3px;font-size:8px;font-weight:600;margin-right:2px}
.origin-badge.waf{background:var(--accent-dim);color:var(--accent-text)}.origin-badge.caf{background:var(--green-dim);color:var(--green-text)}.origin-badge.lza{background:var(--orange-dim);color:var(--orange-text)}.origin-badge.avd{background:var(--surface);color:var(--text-dim)}
.excluded-row{opacity:.4}
.impact-box{margin-top:4px;padding:4px 8px;background:var(--accent-dim);border-left:3px solid var(--orange-text);border-radius:var(--radius-xs);font-size:10px;color:var(--text-secondary)}
.impact-box strong{color:var(--orange-text)}
.remediation-box{margin-top:3px;padding:4px 8px;background:var(--card);border-left:3px solid var(--green-text);border-radius:var(--radius-xs);font-size:10px;color:var(--text-dim)}
.remediation-box strong{color:var(--green-text)}
.footer{margin-top:36px;padding-top:16px;border-top:1px solid var(--border);display:flex;justify-content:space-between;color:var(--text-faint);font-size:10px}
.footer-dot{width:6px;height:6px;border-radius:50%;background:var(--accent);display:inline-block;margin-right:6px}
@media(max-width:900px){.summary{grid-template-columns:repeat(2,1fr)}.container{padding:16px}.maturity-dims{grid-template-columns:repeat(2,1fr)}.effort-grid{grid-template-columns:repeat(2,1fr)}}
@media print{.toolbar{position:static}.btn,.toolbar input{display:none}body{background:#fff;color:#000}.stat-card,.section-header,.check-table th{background:#f5f5f5;border-color:#ddd}.section-content{display:block!important}}
</style>
</head>
<body>
<div class='toolbar'>
  <h1><span class='footer-dot'></span>BaselinePilot Report</h1>
  <div class='spacer'></div>
  <input type='text' id='searchBox' placeholder='Search checks...' oninput="var v=this.value.toLowerCase();var rows=document.querySelectorAll('.check-table tbody tr');for(var i=0;i<rows.length;i++){rows[i].style.display=rows[i].getAttribute('data-search').toLowerCase().indexOf(v)>=0?'':'none'}">
  <button class='btn' onclick="var t=document.documentElement.getAttribute('data-theme');document.documentElement.setAttribute('data-theme',t==='dark'?'light':'dark')">&#9728; Theme</button>
  <button class='btn' onclick="var s=document.getElementsByClassName('section-content');for(var i=0;i<s.length;i++){s[i].className='section-content open'};var h=document.getElementsByClassName('section-header');for(var j=0;j<h.length;j++){h[j].className='section-header open'}">Expand All</button>
  <button class='btn' onclick="var s=document.getElementsByClassName('section-content open');while(s.length>0){s[0].className='section-content'};var h=document.getElementsByClassName('section-header open');while(h.length>0){h[0].className='section-header'}">Collapse</button>
  <button class='btn' onclick='window.print()'>&#128424; Print</button>
</div>
<div class='container'>
<div class='meta-grid'>
  <div class='meta-item'>Customer <span>$(& $enc $Global:Assessment.CustomerName)</span></div>
  <div class='meta-item'>Assessor <span>$(& $enc $Global:Assessment.AssessorName)</span></div>
  <div class='meta-item'>Date <span>$($Global:Assessment.Date)</span></div>
  <div class='meta-item'>Join Type <span>$(& $enc $Global:Assessment.JoinType)</span></div>
  <div class='meta-item'>Host <span>$SysHost</span></div>
  <div class='meta-item'>OS Build <span>$SysBuild</span></div>
  <div class='meta-item'>Tool <span>BaselinePilot v$Global:AppVersion</span></div>
  <div class='meta-item'>Checks <span>$Total ($Assessed evaluated)</span></div>
</div>
<div class='score-hero'>
  <div class='score-row'>
    <div class='score-block'><div class='val $ScoreClass'>$(if ($Compliance -ge 0) { "$Compliance%" } else { [string][char]0x2014 })</div><div class='lbl'>Baseline Compliance</div></div>
    <div class='score-block'><div class='val $RiskClass'>$(if ($Score -ge 0) { "$Score" } else { [string][char]0x2014 })</div><div class='lbl'>Weighted Risk Score</div></div>
  </div>
</div>
<div class='summary'>
  <div class='stat-card pass'><div class='accent-bar'></div><div class='num'>$Pass</div><div class='label'>Pass</div></div>
  <div class='stat-card fail'><div class='accent-bar'></div><div class='num'>$Fail</div><div class='label'>Fail</div></div>
  <div class='stat-card warn'><div class='accent-bar'></div><div class='num'>$Warn</div><div class='label'>Warning</div></div>
  $(if ($AccRisk -gt 0) { "<div class='stat-card warn' style='opacity:0.8'><div class='accent-bar'></div><div class='num'>$AccRisk</div><div class='label'>Accepted Risk</div></div>" })
  $(if ($Defer -gt 0) { "<div class='stat-card' style='border-left:3px solid var(--accent)'><div class='num' style='color:var(--accent)'>$Defer</div><div class='label'>Deferred</div></div>" })
  <div class='stat-card na'><div class='accent-bar'></div><div class='num'>$NotA</div><div class='label'>Not Assessed</div></div>
  <div class='stat-card total'><div class='accent-bar'></div><div class='num'>$Total</div><div class='label'>Total Checks</div></div>
</div>
<div class='progress-bar'><div class='seg seg-pass' style='width:${PctPass}%'></div><div class='seg seg-warn' style='width:${PctWarn}%'></div><div class='seg seg-fail' style='width:${PctFail}%'></div></div>
$(if ($Fail -gt 0) { @"
<div style='margin-bottom:28px'>
<div style='font-size:15px;font-weight:700;color:var(--text-bright);margin-bottom:4px'>Remediation Priority</div>
<div style='font-size:11px;color:var(--text-dim);margin-bottom:12px'>Failed checks grouped by required effort</div>
<div class='effort-grid'>
  <div class='effort-card quick'><div class='effort-icon'>&#9889;</div><div class='effort-count'>$EffortQuick</div><div class='effort-label'>Quick Wins</div></div>
  <div class='effort-card some'><div class='effort-icon'>&#128736;</div><div class='effort-count'>$EffortSome</div><div class='effort-label'>Some Effort</div></div>
  <div class='effort-card major'><div class='effort-icon'>&#127959;</div><div class='effort-count'>$EffortMajor</div><div class='effort-label'>Major Effort</div></div>
</div>
</div>
"@ })
<div style='font-size:15px;font-weight:700;color:var(--text-bright);margin-bottom:4px'>Maturity Dimensions</div>
<div style='font-size:11px;color:var(--text-dim);margin-bottom:12px'>Composite scores across 6 security dimensions</div>
<div class='maturity-dims'>$DimCards</div>
$CatSections
<div class='footer'>
  <span><span class='footer-dot'></span>BaselinePilot v$Global:AppVersion</span>
  <span>Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm') | $Total checks | $(& $enc $Global:Assessment.JoinType)</span>
</div>
</div>
</body></html>
"@

        [System.IO.File]::WriteAllText($dlg.FileName, $Html, [System.Text.Encoding]::UTF8)
        Write-DebugLog "HTML report exported: $($dlg.FileName)" -Level 'SUCCESS'
        Unlock-Achievement 'export_html'
        Show-Toast "HTML report exported" -Type 'Success'
        if ($Global:OpenAfterExport) { Start-Process $dlg.FileName }
    } catch {
        Write-DebugLog "HTML export failed: $($_.Exception.Message)" -Level 'ERROR'
        Show-Toast "HTML export failed: $($_.Exception.Message)" -Type 'Error'
    }
}

# SECTION 17: EVENT HANDLERS
# ===============================================================================

# Title bar buttons
$btnMinimize.Add_Click({ $Window.WindowState = 'Minimized' })
$btnMaximize.Add_Click({
    $Window.WindowState = if ($Window.WindowState -eq 'Maximized') { 'Normal' } else { 'Maximized' }
})
$btnClose.Add_Click({ $Window.Close() })

# Theme toggle
$btnThemeToggle.Add_Click({
    if ($Global:SuppressThemeHandler) { return }
    $NewLight = -not $Global:IsLightMode
    $Global:SuppressThemeHandler = $true
    & $ApplyTheme $NewLight
    $btnThemeToggle.Content = if ($NewLight) { [char]0x263D } else { [char]0x2600 }
    $chkDarkMode.IsChecked = -not $NewLight
    $Global:SuppressThemeHandler = $false
    Unlock-Achievement 'theme_toggle'
    Update-AchievementBadges
    Save-UserPrefs
})

$chkDarkMode.Add_Checked({
    if ($Global:SuppressThemeHandler) { return }
    $Global:SuppressThemeHandler = $true
    & $ApplyTheme $false
    $btnThemeToggle.Content = [char]0x2600
    $Global:SuppressThemeHandler = $false
    Update-AchievementBadges
    Save-UserPrefs
})
$chkDarkMode.Add_Unchecked({
    if ($Global:SuppressThemeHandler) { return }
    $Global:SuppressThemeHandler = $true
    & $ApplyTheme $true
    $btnThemeToggle.Content = [char]0x263D
    $Global:SuppressThemeHandler = $false
    Update-AchievementBadges
    Save-UserPrefs
})

# Settings controls
$chkAutoSave.Add_Checked({  $Global:AutoSaveEnabled = $true })
$chkAutoSave.Add_Unchecked({ $Global:AutoSaveEnabled = $false })
$chkOpenAfterExport.Add_Checked({  $Global:OpenAfterExport = $true })
$chkOpenAfterExport.Add_Unchecked({ $Global:OpenAfterExport = $false })
$chkVerboseLog.Add_Checked({  $Global:VerboseLogging = $true })
$chkVerboseLog.Add_Unchecked({ $Global:VerboseLogging = $false })
$chkShowActivityLog.Add_Checked({  Toggle-ActivityLog $true })
$chkShowActivityLog.Add_Unchecked({ Toggle-ActivityLog $false })

$cmbAutoSaveInterval.Add_SelectionChanged({
    $Sel = $cmbAutoSaveInterval.SelectedItem
    if ($Sel) {
        $Val = [int]($Sel.Content -replace 's','')
        $Global:AutoSaveInterval = $Val
        if ($Global:AutoSaveTimer) {
            $Global:AutoSaveTimer.Interval = [TimeSpan]::FromSeconds($Val)
        }
    }
})

$cmbMaxBackups.Add_SelectionChanged({
    $Sel = $cmbMaxBackups.SelectedItem
    if ($Sel) { $Global:MaxBackups = [int]$Sel.Content }
})

$btnBrowseExportPath.Add_Click({
    $FBD = New-Object System.Windows.Forms.FolderBrowserDialog
    $FBD.SelectedPath = $Global:ReportDir
    if ($FBD.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $Global:ReportDir = $FBD.SelectedPath
        $txtExportPath.Text = $FBD.SelectedPath
        Save-UserPrefs
    }
})

$btnResetPrefs.Add_Click({
    $Result = Show-ThemedDialog `
        -Title 'Reset Settings' `
        -Message "This will reset all settings to defaults.`nYour assessment data will not be affected." `
        -Icon ([char]0xE7BA) `
        -IconColor 'ThemeWarning' `
        -Buttons @(
            @{ Text = 'Reset'; Result = 'Yes'; IsAccent = $true }
            @{ Text = 'Cancel'; Result = 'No' }
        )
    if ($Result -ne 'Yes') { return }

    $Global:SuppressThemeHandler = $true
    $Global:IsLightMode = $false; & $ApplyTheme $false
    $btnThemeToggle.Content = [char]0x2600
    $chkDarkMode.IsChecked = $true
    $Global:SuppressThemeHandler = $false

    $txtDefaultAssessor.Text = ''
    $txtNotes.Text = ''
    $Global:AutoSaveEnabled = $true; $chkAutoSave.IsChecked = $true
    $Global:AutoSaveInterval = 60; $cmbAutoSaveInterval.SelectedIndex = 1
    if ($Global:AutoSaveTimer) { $Global:AutoSaveTimer.Interval = [TimeSpan]::FromSeconds(60) }
    $Global:ReportDir = Join-Path $Global:Root 'reports'; $txtExportPath.Text = $Global:ReportDir
    $Global:OpenAfterExport = $true; $chkOpenAfterExport.IsChecked = $true
    $Global:VerboseLogging = $false; $chkVerboseLog.IsChecked = $false
    $Global:MaxBackups = 10; $cmbMaxBackups.SelectedIndex = 1
    Toggle-ActivityLog $true; $chkShowActivityLog.IsChecked = $true

    Update-Dashboard; Render-BaselineChecks; Render-Findings; Update-ReportPreview; Refresh-AssessmentList
    Write-DebugLog "All settings reset to defaults" -Level 'INFO'
    Show-Toast 'Settings restored to defaults' -Type 'Info'
})

# Achievement header collapse/expand
$pnlAchievementHeader = $Window.FindName("pnlAchievementHeader")
$pnlAchievementHeader.Add_MouseLeftButtonDown({
    $pnl = $Window.FindName('pnlAchievements')
    $chev = $Window.FindName('lblAchievementChevron')
    if ($pnl.Visibility -eq 'Visible') {
        $pnl.Visibility = 'Collapsed'
        $chev.Text = [char]0xE76C
    } else {
        $pnl.Visibility = 'Visible'
        $chev.Text = [char]0xE76B
    }
})

# Icon rail navigation
$railDashboard.Add_Click({  Switch-Tab 'Dashboard' })
$railBaseline.Add_Click({   Switch-Tab 'Baseline' })
$railFindings.Add_Click({   Switch-Tab 'Findings' })
$railReport.Add_Click({     Switch-Tab 'Report' })
$railSettings.Add_Click({   Switch-Tab 'Settings' })

# Sidebar buttons
$btnSaveAssessment.Add_Click({ Save-Assessment; Refresh-AssessmentList })
$btnImportCollection.Add_Click({
    if (-not (Confirm-OverwriteAssessment -Action 'importing a collection')) { return }
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'JSON Files (*.json)|*.json'
    $dlg.Title = 'Import Collection JSON'
    if ($dlg.ShowDialog() -eq $true) {
        Import-CollectionJson -Path $dlg.FileName
        Refresh-AssessmentList
    }
})
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
            if ($Json.Checks) {
                $Global:Assessment = $Json
                $Global:Assessment.Checks = [System.Collections.ArrayList]@($Json.Checks)
                $ExIds = @($Json.Checks | ForEach-Object { $_.Id })
                foreach ($Def in $Global:CheckDefinitions) {
                    if ($Def.id -notin $ExIds) {
                        [void]$Global:Assessment.Checks.Add([PSCustomObject]@{
                            Id=$Def.id;Category=$Def.category;Name=$Def.name;Description=$Def.description
                            Severity=$Def.severity;Type=$Def.type;Weight=[int]$Def.weight;Priority=$Def.priority
                            Origin=$Def.origin;Effort=$Def.effort;Impact=$Def.impact;Remediation=$Def.remediation
                            Reference=$Def.reference;CollectionKeys=$Def.collectionKeys;BaselineValue=$Def.baselineValue
                            CspPath=$Def.cspPath;Rationale=$Def.rationale;applicableTo=$Def.applicableTo;threshold=$Def.threshold
                            eventIds=$Def.eventIds;filterField=$Def.filterField;filterValues=$Def.filterValues
                            ActualValue=$null;Status='Not Assessed';AutoStatus=$null;Excluded=$false;Details='';Notes='';Source=$Def.type
                            AffectedMachines=[System.Collections.ArrayList]::new()
                        })
                    }
                }
                $txtCustomerName.Text   = $Json.CustomerName
                $txtAssessorName.Text   = $Json.AssessorName
                $txtAssessmentDate.Text = $Json.Date
                if ($txtNotes) { $txtNotes.Text = $Json.Notes }
                Update-Dashboard; Update-Progress; Render-BaselineChecks; Render-Findings
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
$btnCopySummary.Add_Click({
    if ($Global:ExecutiveSummaryText) {
        $dataObj = New-Object System.Windows.DataObject
        $dataObj.SetData([System.Windows.DataFormats]::Rtf, (Get-ExecutiveSummaryRtf))
        $dataObj.SetData([System.Windows.DataFormats]::UnicodeText, $Global:ExecutiveSummaryText)
        [System.Windows.Clipboard]::SetDataObject($dataObj, $true)
        Show-Toast 'Executive summary copied to clipboard (rich text)' -Type 'Success'
    } else {
        Show-Toast 'No summary available — import a collection first' -Type 'Warning'
    }
})

# Filter changes trigger re-render
$cmbFilterSeverity.Add_SelectionChanged({ Invoke-DeferredRender { Render-Findings } })
$cmbFilterStatus.Add_SelectionChanged({ Invoke-DeferredRender { Render-Findings } })
$cmbFilterCategory.Add_SelectionChanged({ Invoke-DeferredRender { Render-Findings } })
$cmbFilterPriority.Add_SelectionChanged({ Invoke-DeferredRender { Render-Findings } })
$cmbFilterOrigin.Add_SelectionChanged({ Invoke-DeferredRender { Render-Findings } })
$cmbFilterEffort.Add_SelectionChanged({ Invoke-DeferredRender { Render-Findings } })
$txtFindingsSearch.Add_TextChanged({ Invoke-DeferredRender { Render-Findings } })

# Baseline sidebar filter changes
$chkShowDeviationsOnly.Add_Checked({ Invoke-DeferredRender { Render-BaselineChecks } })
$chkShowDeviationsOnly.Add_Unchecked({ Invoke-DeferredRender { Render-BaselineChecks } })
$cmbBaselinePriority.Add_SelectionChanged({ Invoke-DeferredRender { Render-BaselineChecks } })

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

# Dirty tracking on text fields
$txtCustomerName.Add_TextChanged({ Set-Dirty })
$txtAssessorName.Add_TextChanged({ Set-Dirty })

# ===============================================================================
# SECTION 18: BACKGROUND JOB POLLING TIMER
# ===============================================================================

$Timer = New-Object System.Windows.Threading.DispatcherTimer
$Timer.Interval = [TimeSpan]::FromMilliseconds($Script:TIMER_INTERVAL_MS)
$Timer.Add_Tick({
    if ($Global:TimerProcessing) { return }
    $Global:TimerProcessing = $true
    try {
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

# Auto-save timer
$Global:AutoSaveTimer = New-Object System.Windows.Threading.DispatcherTimer
$Global:AutoSaveTimer.Interval = [TimeSpan]::FromSeconds($Global:AutoSaveInterval)
$Global:AutoSaveTimer.Add_Tick({ AutoSave-Assessment })
$Global:AutoSaveTimer.Start()

# ===============================================================================
# SECTION 19: INITIALIZATION & STARTUP
# ===============================================================================

Write-DebugLog "$Global:AppTitle starting..." -Level 'INFO'
Write-DebugLog "[INIT] Root: $Global:Root" -Level 'DEBUG'

$txtAssessmentDate.Text = (Get-Date -Format 'yyyy-MM-dd')
$lblSplashVersion.Text = "v$Global:AppVersion"
$txtExportPath.Text = $Global:ReportDir

# Splash step 1
$dotStep1.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, 'ThemeSuccess')
$lblStep1.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
$lblSplashStatus.Text = "Loading assemblies..."

# Load preferences
Load-UserPrefs

# Render badges after prefs (theme may have changed)
Update-AchievementBadges

# Splash step 2
$dotStep2.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, 'ThemeSuccess')
$lblStep2.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
$lblSplashStatus.Text = "Loading check definitions..."

Write-DebugLog "[INIT] $($Global:CheckDefinitions.Count) check definitions loaded" -Level 'INFO'

# Splash step 3
$dotStep3.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, 'ThemeSuccess')
$lblStep3.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
$lblSplashStatus.Text = "Ready!"

# Pre-fill assessor name from prefs
if ($txtDefaultAssessor.Text -and -not $txtAssessorName.Text) {
    $txtAssessorName.Text = $txtDefaultAssessor.Text
}

# Populate initial views
Update-Dashboard
Render-BaselineChecks
Refresh-AssessmentList

# Dismiss splash after short delay
$SplashTimer = New-Object System.Windows.Threading.DispatcherTimer
$SplashTimer.Interval = [TimeSpan]::FromMilliseconds(800)
$SplashTimerRef = $SplashTimer
$SplashTimer.Add_Tick({
    $SplashTimerRef.Stop()
    $pnlSplash.Visibility = 'Collapsed'
    Write-DebugLog "[INIT] Splash dismissed, ready" -Level 'DEBUG'
})
$SplashTimer.Start()

Write-DebugLog "[INIT] Startup complete — $($Global:Assessment.Checks.Count) checks, $(Get-Categories | Measure-Object | Select-Object -ExpandProperty Count) categories" -Level 'SUCCESS'

# CSP staleness warning (after splash so toast is visible)
if ($Global:CspDbAge -gt 90) {
    Write-DebugLog "[INIT] CSP metadata is $($Global:CspDbAge) days old — remediation data may be outdated" -Level 'WARN'
    Show-Toast "CSP metadata is $($Global:CspDbAge) days old. Run Build-CspDatabase.ps1 to refresh." -Type 'Warning' -DurationMs 8000
} elseif ($Global:CspMetadata.Count -eq 0) {
    Write-DebugLog "[INIT] No CSP metadata loaded — remediation enrichment unavailable" -Level 'WARN'
}

# ===============================================================================
# SECTION 20: WINDOW EVENTS & SHOWDIALOG
# ===============================================================================

$Window.Add_Closing({
    param($sender, $e)
    if ($Global:IsDirty -or (Test-AssessmentDirty)) {
        $Result = Show-ThemedDialog `
            -Title 'Unsaved Changes' `
            -Message 'You have unsaved assessment data. Save before closing?' `
            -Icon ([char]0xE7BA) `
            -IconColor 'ThemeWarning' `
            -Buttons @(
                @{ Text = 'Save & Close'; Result = 'Save'; IsAccent = $true }
                @{ Text = 'Discard';      Result = 'Discard' }
                @{ Text = 'Cancel';       Result = 'Cancel' }
            )
        switch ($Result) {
            'Save'    { Save-Assessment }
            'Cancel'  { $e.Cancel = $true; return }
        }
    }

    # Cleanup
    Save-UserPrefs
    $Timer.Stop()
    $Global:AutoSaveTimer.Stop()
    foreach ($Job in @($Global:BgJobs)) {
        try { $Job.PS.Stop(); $Job.PS.Dispose(); $Job.RS.Dispose() } catch { }
    }
    Write-DebugLog "BaselinePilot closed" -Level 'INFO'
})

$Window.Add_StateChanged({
    if ($Window.WindowState -eq 'Maximized') {
        $btnMaximize.Content = [char]0xE923
    } else {
        $btnMaximize.Content = [char]0xE922
    }
})

# Bring to foreground
$Window.Add_ContentRendered({
    try {
        $helper = [System.Windows.Interop.WindowInteropHelper]::new($Window)
        [ForegroundHelper]::SetForegroundWindow($helper.Handle) | Out-Null
    } catch { }
})

$Window.ShowDialog() | Out-Null