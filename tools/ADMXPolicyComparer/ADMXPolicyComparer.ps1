<#
.SYNOPSIS
    ADMX Policy Comparer - WPF GUI for comparing Group Policy ADML/ADMX files.
.DESCRIPTION
    Modern dark/light-themed WPF application that compares two versions of Microsoft
    Group Policy ADMX/ADML templates and presents colour-coded results with filtering,
    search, side-by-side word-level diffs, registry metadata, and export to HTML/CSV.
    Includes achievement/gamification system, comparison history with result persistence,
    daily streak tracker, and a rich debug console with minimap and heatmap.

    The tool parses both ADML (localised strings) and ADMX (policy definitions) to
    extract registry keys, value names, value types, scope, enabled/disabled values,
    and enumeration choices for each policy setting.

    Layout: scripts/ADMXPolicyComparer/
      - ADMXPolicyComparer.ps1     Main script (this file)
      - ADMXPolicyComparer_UI.xaml WPF layout definition

    Major sections (in source order):
      Pre-flight             Assembly loading, DPI awareness, globals
      Theme engine           Dark/light palette application
      Debug console          Write-DebugLog, minimap, heatmap, log panel
      Progress & status      Shimmer bar, status dot
      Preferences            Load/Save user settings (JSON)
      History & results      Comparison history, result persistence (JSON)
      Streak & achievements  Daily streak, 24 achievement definitions
      Log formatting         Colour-coded comparison log (RichTextBox)
      Animations             Toast, confetti, count-up, breathe, fade, transitions
      Navigation             Tab switching, theme transition, folder picker, modal dialog
      Version scanning       Enumerate ADMX template versions
      ADML/ADMX parsing      String extraction, category/registry/element metadata
      Comparison engine      File-by-file diff with trivial-change filter
      Filtering & search     CollectionView filter, policy grouping, file sidebar
      Dashboard              Stat cards, recent comparisons, export history
      Export                 HTML report (with sidebar, search, theme) and CSV
      Results UI             Empty state, status bar, breakdown chart
      Detail pane            Word-level diff (LCS), registry info bar, metadata display
      Window lifecycle       Load, close, keyboard shortcuts, micro-interactions
.NOTES
    Requires: Windows 10+, PowerShell 5.1+, .NET Framework 4.7.2+
    Author:   Anton Romanyuk
    Version:  0.1.0-alpha
#>

# -- Pre-flight ----------------------------------------------------------------
param([switch]$Debug)
$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Web

# DPI awareness
try {
    if (-not ([System.Management.Automation.PSTypeName]'DpiHelper').Type) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class DpiHelper {
    [DllImport("shcore.dll")] public static extern int SetProcessDpiAwareness(int value);
}
"@
    }
    [void][DpiHelper]::SetProcessDpiAwareness(2)
} catch {}

# -- Globals -------------------------------------------------------------------
$Script:AppVersion   = '0.1.0-alpha'
$Script:AppDir       = Split-Path -Parent $MyInvocation.MyCommand.Definition
$Script:PrefsPath    = Join-Path $Script:AppDir 'user_prefs.json'
$Script:AchievePath  = Join-Path $Script:AppDir 'achievements.json'
$Script:HistoryPath  = Join-Path $Script:AppDir 'comparison_history.json'
$Script:ResultsDir   = Join-Path $Script:AppDir 'results'
$Script:StreakPath   = Join-Path $Script:AppDir 'streak.json'
$Script:LogPath      = Join-Path $Script:AppDir 'debug.log'

$Script:AllResults   = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
$Script:LoadedHistoryId = $null
$Script:LoadedOlderVersion = $null
$Script:LoadedNewerVersion = $null
$Script:CurrentFilter = 'All'
$Script:CurrentSearch = ''
$Script:AnimationsDisabled = $false
$Script:IsComparing  = $false
$Script:BottomPanelVisible = $false
$Script:BottomPanelExpanded = $false
$Script:LogLineCount = 0
$Script:LogBuffer    = [System.Collections.Generic.List[PSCustomObject]]::new()
$Script:MaxLogLines  = 500
$Script:FullLogSB    = [System.Text.StringBuilder]::new(8192)
$Script:FullLogLines = 0
$Script:DebugOverlayEnabled = $false
$Script:MinimapLastIdx  = 0
$Script:HeatmapLastIdx  = 0

# -- File-based debug log (survives crashes, rotates at 2 MB) -----------------
try {
    if ((Test-Path $Script:LogPath) -and (Get-Item $Script:LogPath).Length -gt 2MB) {
        $prevLog = $Script:LogPath + '.prev'
        if (Test-Path $prevLog) { Remove-Item $prevLog -Force -ErrorAction SilentlyContinue }
        Rename-Item $Script:LogPath $prevLog -Force -ErrorAction SilentlyContinue
    }
    $header = "`n=== ADMX Policy Comparer v$($Script:AppVersion) - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
    [System.IO.File]::AppendAllText($Script:LogPath, $header + "`r`n")
} catch { <# best-effort #> }

# Theme palettes
$Script:ThemeDark = @{
    ThemeAppBg='#FF111113'; ThemePanelBg='#FF18181B'; ThemeCardBg='#FF1E1E1E'
    ThemeInputBg='#FF141414'; ThemeDeepBg='#FF0D0D0D'
    ThemeAccent='#FF0078D4'; ThemeAccentHover='#FF1A8AD4'; ThemeAccentLight='#FF60CDFF'
    ThemeGreenAccent='#FF00C853'
    ThemeTextPrimary='#FFFFFFFF'; ThemeTextBody='#FFE0E0E0'; ThemeTextSecondary='#FFA1A1AA'
    ThemeTextMuted='#FF8B8B93'; ThemeTextDim='#FF7A7A84'
    ThemeHoverBg='#FF27272B'; ThemeSelectedBg='#FF2A2A2A'; ThemePressedBg='#FF1A1A1A'
    ThemeError='#FFFF5000'; ThemeWarning='#FFF59E0B'; ThemeSuccess='#FF00C853'
    ThemeBorder='#19FFFFFF'; ThemeBorderCard='#19FFFFFF'; ThemeBorderElevated='#FF333333'
    ThemeBorderSubtle='#19FFFFFF'
    ThemeScrollTrack='#08FFFFFF'; ThemeScrollThumb='#40FFFFFF'; ThemeOutputBg='#FF0F0F11'
    StatusAdded='#FF00C853'; StatusRemoved='#FFFF5000'; StatusModified='#FFF59E0B'
    StatusFileAdded='#FF00E676'; StatusFileRemoved='#FFFF6E40'
}
$Script:ThemeLight = @{
    ThemeAppBg='#FFF5F5F5'; ThemePanelBg='#FFFFFFFF'; ThemeCardBg='#FFFFFFFF'
    ThemeInputBg='#FFF0F0F0'; ThemeDeepBg='#FFE8E8E8'
    ThemeAccent='#FF0078D4'; ThemeAccentHover='#FF106EBE'; ThemeAccentLight='#FF005A9E'
    ThemeGreenAccent='#FF00C853'
    ThemeTextPrimary='#FF111111'; ThemeTextBody='#FF333333'; ThemeTextSecondary='#FF6B6B73'
    ThemeTextMuted='#FF636370'; ThemeTextDim='#FF707078'
    ThemeHoverBg='#FFEBEBEB'; ThemeSelectedBg='#FFE0E0E0'; ThemePressedBg='#FFD5D5D5'
    ThemeError='#FFD32F2F'; ThemeWarning='#FFF57F17'; ThemeSuccess='#FF2E7D32'
    ThemeBorder='#19000000'; ThemeBorderCard='#19000000'; ThemeBorderElevated='#FFCCCCCC'
    ThemeBorderSubtle='#19000000'
    ThemeScrollTrack='#08000000'; ThemeScrollThumb='#40000000'; ThemeOutputBg='#FFF0F0F2'
    StatusAdded='#FF1B7D3A'; StatusRemoved='#FFC62828'; StatusModified='#FFB8860B'
    StatusFileAdded='#FF2E7D32'; StatusFileRemoved='#FFC41C00'
}

# -- Load XAML -----------------------------------------------------------------
$xamlPath = Join-Path $Script:AppDir 'ADMXPolicyComparer_UI.xaml'
$xamlContent = [System.IO.File]::ReadAllText($xamlPath)
$xml = [xml]$xamlContent
$reader = [System.Xml.XmlNodeReader]::new($xml)
$Window = [System.Windows.Markup.XamlReader]::Load($reader)
$reader.Close()

# Named element references
$ui = @{}
@(
    'WindowBorder','TitleBar','TitleText','VersionBadge','BtnMinimize','BtnMaximize','BtnClose','MaximizeIcon'
    'NavDashboard','NavCompare','NavResults','NavSettings','NavToggleSidebar','NavToggleConsole'
    'LeftSidebarPanel','SidebarDashboard','SidebarCompare','SidebarResults','SidebarSettings'
    'DashSidebarSummary','DashAchievementCount','DashAchievementWrap'
    'BasePathText','BtnBrowseBasePath','BtnScanVersions','VersionCountText','AvailableVersionsList'
    'FileFilterAllBtn','FileFilterList','ResultsSidebarSummary'
    'PanelDashboard','PanelCompare','PanelResults','PanelSettings'
    'StatComparisons','StatChanges','StatFiles','StatStreak'
    'RecentComparisonsList','NoRecentText','BtnQuickCompare','BtnQuickExport','BtnQuickSettings','ExportHistoryCombo'
    'OlderVersionCombo','NewerVersionCombo','BtnRunComparison','ComparisonStatusText'
    'ComparisonProgressBar','ComparisonLogText'
    'FilterAll','FilterAdded','FilterRemoved','FilterModified','FilterFileAdded','FilterFileRemoved'
    'SearchBox','SearchPlaceholder','ResultsDataGrid','ResultsSubtitle','ResultsCountText'
    'ResultsEmptyGuide','BtnGetStarted','BtnResultsGoCompare'
    'BtnExportHtml','BtnExportCsv'
    'ThemeToggle','AnimationsToggle','DebugOverlayToggle','BasePathSetting','ExportPathSetting'
    'BtnBrowseBasePathSetting','BtnBrowseExportPath','BtnResetAll'
    'AchievementSummaryText','AchievementProgressBar','AchievementsPanel'
    'StatusText','statusDot','StatusRight'
    'ToastBorder','ToastTitle','ToastMessage'
    'ConfettiCanvas'
    # Title bar buttons
    'BtnHelp','BtnThemeToggle'
    # Debug console elements
    'pnlBottomPanel','rtbActivityLog','docActivityLog','paraLog','logScroller'
    'cnvMinimap','cnvHeatmap','btnToggleBottomSize','btnClearLog','btnHideBottom'
    'ConsoleLineCount','BottomSplitter','BottomPanelRow'
    # Shimmer progress
    'pnlGlobalProgress','prgGlobal','brdGlobalShimmer','shimmerTranslate','lblGlobalProgress'
    'StatCard1','StatCard2','StatCard3','StatCard4'
    'SearchIcon','StatusResults','StatusStreak'
    'BtnSearchClear','CompareArrow','ArrowBounce'
    'LinkWin10_22H2','LinkWin11_23H2','LinkWin11_24H2','LinkWin11_25H2','LinkWin11_26H1'
    'SidebarLinkWin10_22H2','SidebarLinkWin11_23H2','SidebarLinkWin11_24H2','SidebarLinkWin11_25H2','SidebarLinkWin11_26H1'
    'ToastIcon','ToastAccentBar'
    'pnlDetailPane','lblDetailIcon','lblDetailTitle','lblDetailFile'
    'rtbDetailOld','docDetailOld','paraDetailOld','rtbDetailNew','docDetailNew','paraDetailNew'
    'btnCopyOldValue','btnCopyNewValue','btnCopyStringId'
    'pnlBreakdownChart','pnlChartBar','pnlChartLegend'
    'pnlRegistryBar','lblRegistryPath','lblRegistryValue','badgeRegistryScope','lblRegistryScope','lblRegistryDetails'
    'BtnGroupByPolicy'
) | ForEach-Object { $ui[$_] = $Window.FindName($_) }

$ui.ResultsDataGrid.ItemsSource = $Script:AllResults
$ui.VersionBadge.Text  = "v$($Script:AppVersion)"
$ui.StatusRight.Text   = "ADMX Policy Comparer v$($Script:AppVersion)"
# -- Theme engine --------------------------------------------------------------
    <#
    .SYNOPSIS
        Applies a colour palette to all WPF dynamic resource brushes.
    .PARAMETER palette
        Hashtable mapping resource key names to hex colour strings (e.g. #FF0078D4).
    #>
function Set-Theme([hashtable]$palette) {
    foreach ($kv in $palette.GetEnumerator()) {
        $newBrush = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.ColorConverter]::ConvertFromString($kv.Value))
        $Window.Resources[$kv.Key] = $newBrush
    }
}

# -- Debug Console / Write-DebugLog -------------------------------------------
$Script:LogLevelColorsDark = @{
    INFO    = '#60CDFF'
    SUCCESS = '#00C853'
    WARN    = '#F59E0B'
    ERROR   = '#FF5000'
    DEBUG   = '#A78BFA'
    STEP    = '#F472B6'
    SYSTEM  = '#71717A'
}
$Script:LogLevelColorsLight = @{
    INFO    = '#0078D4'
    SUCCESS = '#008A2E'
    WARN    = '#B86E00'
    ERROR   = '#CC0000'
    DEBUG   = '#8888AA'
    STEP    = '#D63384'
    SYSTEM  = '#636370'
}
$Script:LogLevelColors = $Script:LogLevelColorsDark

    <#
    .SYNOPSIS
        Writes a timestamped, level-coded log entry to console, RichTextBox, and disk.
    .DESCRIPTION
        Central logging function. Entries are written to the PowerShell console, appended to
        a rotating disk log file, stored in a ring buffer, and displayed in the WPF activity
        log panel with colour-coded formatting. DEBUG-level messages are suppressed unless
        the debug overlay is enabled.
    #>
function Write-DebugLog {
    param(
        [string]$Message,
        [ValidateSet('INFO','SUCCESS','WARN','ERROR','DEBUG','STEP','SYSTEM')]
        [string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format 'HH:mm:ss.fff'
    $line = "[$timestamp] [$Level] $Message"

    # Always write to PS console
    Write-Host $line -ForegroundColor DarkGray

    # Append to disk log (survives crashes)
    try { [System.IO.File]::AppendAllText($Script:LogPath, $line + "`r`n") } catch {}

    # Full log buffer (keeps everything for replay)
    $Script:FullLogLines++
    if ($Script:FullLogLines -gt ($Script:MaxLogLines * 2)) {
        $text = $Script:FullLogSB.ToString()
        $nl = $text.IndexOf("`n")
        if ($nl -ge 0) {
            $Script:FullLogSB.Clear()
            $Script:FullLogSB.Append($text.Substring($nl + 1)) | Out-Null
            $Script:FullLogLines--
        }
    }
    $Script:FullLogSB.AppendLine($line) | Out-Null

    # Skip DEBUG-level from visible log unless debug overlay is enabled
    if ($Level -eq 'DEBUG' -and -not $Script:DebugOverlayEnabled) { return }

    $color = $Script:LogLevelColors[$Level]

    # Buffer for minimap/heatmap
    $entry = [PSCustomObject]@{ Timestamp = $timestamp; Level = $Level; Message = $Message; Color = $color }
    $Script:LogBuffer.Add($entry)
    $Script:LogLineCount++

    # Ring buffer: trim old lines
    if ($Script:LogBuffer.Count -gt $Script:MaxLogLines) {
        $Script:LogBuffer.RemoveAt(0)
        if ($ui.paraLog -and $ui.paraLog.Inlines.Count -ge 2) {
            try { $ui.paraLog.Inlines.Remove($ui.paraLog.Inlines.FirstInline) } catch {}
            try { $ui.paraLog.Inlines.Remove($ui.paraLog.Inlines.FirstInline) } catch {}
        }
    }

    # Add to RichTextBox
    if ($ui.paraLog) {
        $converter = [System.Windows.Media.BrushConverter]::new()
        $tsColor = if ($Script:Prefs.IsLightMode) { '#8B8B93' } else { '#52525B' }
        if ($ui.paraLog.Inlines.Count -gt 0) {
            $ui.paraLog.Inlines.Add([System.Windows.Documents.LineBreak]::new())
        }
        # Timestamp run
        $tsRun = [System.Windows.Documents.Run]::new("[$timestamp] ")
        $tsRun.Foreground = $converter.ConvertFromString($tsColor)
        $tsRun.FontSize = 10.5
        $ui.paraLog.Inlines.Add($tsRun)

        # Level badge + message
        $msgRun = [System.Windows.Documents.Run]::new("[$Level] $Message")
        $msgRun.Foreground = $converter.ConvertFromString($color)
        $msgRun.FontSize = 11
        $ui.paraLog.Inlines.Add($msgRun)

        # Collect minimap/heatmap data for non-INFO levels
        if ($Level -ne 'INFO' -and $Level -ne 'SYSTEM') {
            $Script:MinimapDots.Add([PSCustomObject]@{ Color = $color; Line = $Script:LogLineCount })
            $Script:HeatmapDots.Add([PSCustomObject]@{ Color = $color; Line = $Script:LogLineCount })
        }

        # Auto-scroll
        $ui.logScroller.ScrollToEnd()
    }

    # Update line count
    if ($ui.ConsoleLineCount) {
        $ui.ConsoleLineCount.Text = "$($Script:LogLineCount) lines"
    }

    # Update minimap + heatmap periodically
    if ($Script:LogLineCount % 5 -eq 0) {
        Update-Minimap
        Update-Heatmap
    }
}

# Incremental minimap/heatmap dot lists
$Script:MinimapDots  = [System.Collections.Generic.List[PSCustomObject]]::new()
$Script:HeatmapDots  = [System.Collections.Generic.List[PSCustomObject]]::new()

    <#
    .SYNOPSIS
        Incrementally renders coloured dots on the vertical minimap canvas for new log entries.
    #>
function Update-Minimap {
    $cnv = $ui.cnvMinimap
    if (-not $cnv -or $cnv.ActualHeight -le 0) { return }
    $H = $cnv.ActualHeight
    $Total = [Math]::Max($Script:LogLineCount, 1)
    $converter = [System.Windows.Media.BrushConverter]::new()
    # Render new dots since last call
    for ($i = $Script:MinimapLastIdx; $i -lt $Script:MinimapDots.Count; $i++) {
        $dot = $Script:MinimapDots[$i]
        $Y = ($dot.Line / $Total) * $H
        $rect = [System.Windows.Shapes.Rectangle]::new()
        $rect.Width = 10; $rect.Height = 2
        $rect.Fill = $converter.ConvertFromString($dot.Color)
        [System.Windows.Controls.Canvas]::SetLeft($rect, 2)
        [System.Windows.Controls.Canvas]::SetTop($rect, $Y)
        $cnv.Children.Add($rect) | Out-Null
    }
    $Script:MinimapLastIdx = $Script:MinimapDots.Count
    # Viewport indicator
    $tag = 'minimapViewport'
    $old = $cnv.Children | Where-Object { $_.Tag -eq $tag }
    if ($old) { $cnv.Children.Remove($old) }
    if ($ui.logScroller.ExtentHeight -gt 0) {
        $vpTop = ($ui.logScroller.VerticalOffset / $ui.logScroller.ExtentHeight) * $H
        $vpH   = [Math]::Max(($ui.logScroller.ViewportHeight / $ui.logScroller.ExtentHeight) * $H, 4)
        $vp = [System.Windows.Shapes.Rectangle]::new()
        $vp.Width = 14; $vp.Height = $vpH
        $vp.Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#30FFFFFF')
        $vp.Tag = $tag
        [System.Windows.Controls.Canvas]::SetLeft($vp, 0)
        [System.Windows.Controls.Canvas]::SetTop($vp, $vpTop)
        $cnv.Children.Add($vp) | Out-Null
    }
}

    <#
    .SYNOPSIS
        Incrementally renders coloured dots on the horizontal heatmap canvas for new log entries.
    #>
function Update-Heatmap {
    $cnv = $ui.cnvHeatmap
    if (-not $cnv -or $cnv.ActualWidth -le 0) { return }
    $W = $cnv.ActualWidth
    $Total = [Math]::Max($Script:LogLineCount, 1)
    $converter = [System.Windows.Media.BrushConverter]::new()
    # Render new dots since last call
    for ($i = $Script:HeatmapLastIdx; $i -lt $Script:HeatmapDots.Count; $i++) {
        $dot = $Script:HeatmapDots[$i]
        $X = ($dot.Line / $Total) * $W
        $rect = [System.Windows.Shapes.Rectangle]::new()
        $rect.Width = 2; $rect.Height = 10
        $rect.Fill = $converter.ConvertFromString($dot.Color)
        [System.Windows.Controls.Canvas]::SetLeft($rect, $X)
        [System.Windows.Controls.Canvas]::SetTop($rect, 0)
        $cnv.Children.Add($rect) | Out-Null
    }
    $Script:HeatmapLastIdx = $Script:HeatmapDots.Count
    # Viewport indicator
    $tag = 'heatmapViewport'
    $old = $cnv.Children | Where-Object { $_.Tag -eq $tag }
    if ($old) { $cnv.Children.Remove($old) }
    if ($ui.logScroller.ExtentHeight -gt 0) {
        $vpLeft = ($ui.logScroller.VerticalOffset / $ui.logScroller.ExtentHeight) * $W
        $vpW   = [Math]::Max(($ui.logScroller.ViewportHeight / $ui.logScroller.ExtentHeight) * $W, 4)
        $vp = [System.Windows.Shapes.Rectangle]::new()
        $vp.Width = $vpW; $vp.Height = 14
        $vp.Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#30FFFFFF')
        $vp.Tag = $tag
        [System.Windows.Controls.Canvas]::SetLeft($vp, $vpLeft)
        [System.Windows.Controls.Canvas]::SetTop($vp, 0)
        $cnv.Children.Add($vp) | Out-Null
    }
}

    <#
    .SYNOPSIS
        Rebuilds the entire RichTextBox activity log from the in-memory buffer.
    .DESCRIPTION
        Used after theme switches or debug overlay toggles to re-render all log lines
        with updated colour palettes. Resets minimap and heatmap canvases.
    #>
function Rebuild-LogPanel {
    if (-not $ui.paraLog) { return }
    $ui.paraLog.Inlines.Clear()
    $Script:LogLineCount = 0
    $Script:LogBuffer.Clear()
    # Reset minimap + heatmap
    $Script:MinimapDots.Clear(); $Script:MinimapLastIdx = 0
    $Script:HeatmapDots.Clear(); $Script:HeatmapLastIdx = 0
    if ($ui.cnvMinimap) { $ui.cnvMinimap.Children.Clear() }
    if ($ui.cnvHeatmap) { $ui.cnvHeatmap.Children.Clear() }
    $allLines = $Script:FullLogSB.ToString() -split "`n" | Where-Object { $_.Trim() }
    if (-not $Script:DebugOverlayEnabled) { $allLines = $allLines | Where-Object { $_ -notmatch '\[DEBUG\]' } }
    $converter = [System.Windows.Media.BrushConverter]::new()
    $isLight = $Script:Prefs.IsLightMode
    $palette = if ($isLight) { $Script:LogLevelColorsLight } else { $Script:LogLevelColorsDark }
    $tsColor = if ($isLight) { '#8B8B93' } else { '#52525B' }
    foreach ($L in $allLines) {
        if ($ui.paraLog.Inlines.Count -gt 0) {
            $ui.paraLog.Inlines.Add([System.Windows.Documents.LineBreak]::new())
        }
        # Determine level color by parsing the line
        $lvl = 'INFO'
        if ($L -match '\[(ERROR|WARN|SUCCESS|DEBUG|STEP|SYSTEM|INFO)\]') { $lvl = $Matches[1] }
        $color = $palette[$lvl]
        if (-not $color) { $color = $palette['INFO'] }
        # Timestamp portion (first bracketed section)
        if ($L -match '^(\[[^\]]+\]\s)(.*)$') {
            $tsRun = [System.Windows.Documents.Run]::new($Matches[1])
            $tsRun.Foreground = $converter.ConvertFromString($tsColor)
            $tsRun.FontSize = 10.5
            $ui.paraLog.Inlines.Add($tsRun)
            $msgRun = [System.Windows.Documents.Run]::new($Matches[2].TrimEnd())
            $msgRun.Foreground = $converter.ConvertFromString($color)
            $msgRun.FontSize = 11
            $ui.paraLog.Inlines.Add($msgRun)
        } else {
            $run = [System.Windows.Documents.Run]::new($L.TrimEnd())
            $run.Foreground = $converter.ConvertFromString($color)
            $run.FontSize = 11
            $ui.paraLog.Inlines.Add($run)
        }
        $Script:LogLineCount++
        # Rebuild entry
        $entry = [PSCustomObject]@{ Timestamp = ''; Level = $lvl; Message = $L; Color = $color }
        $Script:LogBuffer.Add($entry)
        if ($lvl -ne 'INFO' -and $lvl -ne 'SYSTEM') {
            $Script:MinimapDots.Add([PSCustomObject]@{ Color = $color; Line = $Script:LogLineCount })
            $Script:HeatmapDots.Add([PSCustomObject]@{ Color = $color; Line = $Script:LogLineCount })
        }
    }
    if ($ui.ConsoleLineCount) { $ui.ConsoleLineCount.Text = "$($Script:LogLineCount) lines" }
    $ui.logScroller.ScrollToEnd()
    Update-Minimap
    Update-Heatmap
}

# -- Console panel toggle -----------------------------------------------------
    <#
    .SYNOPSIS
        Toggles the bottom debug console panel open or closed.
    #>
function Toggle-ConsolePanel {
    if ($Script:BottomPanelVisible) {
        Hide-ConsolePanel
    } else {
        Show-ConsolePanel
    }
}

    <#
    .SYNOPSIS
        Opens the bottom debug console panel with an optional fade-in animation.
    #>
function Show-ConsolePanel {
    $Script:BottomPanelVisible = $true
    $ui.BottomPanelRow.Height = [System.Windows.GridLength]::new(200)
    $ui.BottomSplitter.Visibility = 'Visible'
    # #35: Fade-in console panel
    if (-not $Script:AnimationsDisabled -and $ui.pnlBottomPanel) {
        $ui.pnlBottomPanel.Opacity = 0
        $consoleFade = [System.Windows.Media.Animation.DoubleAnimation]::new(0, 1,
            [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(200)))
        $consoleFade.EasingFunction = [System.Windows.Media.Animation.CubicEase]::new()
        $ui.pnlBottomPanel.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $consoleFade)
    }
    Write-DebugLog 'Debug console opened' -Level SYSTEM
}

    <#
    .SYNOPSIS
        Collapses the bottom debug console panel.
    #>
function Hide-ConsolePanel {
    $Script:BottomPanelVisible = $false
    $ui.BottomPanelRow.Height = [System.Windows.GridLength]::new(0)
    $ui.BottomSplitter.Visibility = 'Collapsed'
}

    <#
    .SYNOPSIS
        Toggles the console panel between compact (200px) and expanded (450px) height.
    #>
function Toggle-ConsolePanelSize {
    if ($Script:BottomPanelExpanded) {
        $ui.BottomPanelRow.Height = [System.Windows.GridLength]::new(200)
        $Script:BottomPanelExpanded = $false
    } else {
        $ui.BottomPanelRow.Height = [System.Windows.GridLength]::new(450)
        $Script:BottomPanelExpanded = $true
    }
}

# Wire console buttons
$ui.NavToggleConsole.Add_Click({ Toggle-ConsolePanel }.GetNewClosure())

# Wire sidebar toggle
$Script:SidebarCollapsed = $false
$Script:SidebarSavedWidth = 270
if ($ui.NavToggleSidebar) {
    $ui.NavToggleSidebar.Add_Click({
        if ($Script:SidebarCollapsed) {
            $ui.SidebarColumn.Width = [System.Windows.GridLength]::new($Script:SidebarSavedWidth)
            $ui.LeftSidebarPanel.Visibility = 'Visible'
            $Script:SidebarCollapsed = $false
            $ui.NavToggleSidebar.Content = [char]0xE89F
        } else {
            $Script:SidebarSavedWidth = $ui.SidebarColumn.ActualWidth
            if ($Script:SidebarSavedWidth -lt 100) { $Script:SidebarSavedWidth = 270 }
            $ui.SidebarColumn.Width = [System.Windows.GridLength]::new(0)
            $ui.LeftSidebarPanel.Visibility = 'Collapsed'
            $Script:SidebarCollapsed = $true
            $ui.NavToggleSidebar.Content = [char]0xE8A0
        }
    })
}
$ui.btnToggleBottomSize.Add_Click({ Toggle-ConsolePanelSize }.GetNewClosure())
$ui.btnClearLog.Add_Click({
    $ui.paraLog.Inlines.Clear()
    $Script:LogBuffer.Clear()
    $Script:LogLineCount = 0
    $Script:FullLogSB.Clear()
    $Script:FullLogLines = 0
    $ui.ConsoleLineCount.Text = '0 lines'
    $Script:MinimapDots.Clear(); $Script:MinimapLastIdx = 0
    $Script:HeatmapDots.Clear(); $Script:HeatmapLastIdx = 0
    if ($ui.cnvMinimap)  { $ui.cnvMinimap.Children.Clear() }
    if ($ui.cnvHeatmap)  { $ui.cnvHeatmap.Children.Clear() }
    Write-DebugLog 'Console cleared' -Level SYSTEM
}.GetNewClosure())
$ui.btnHideBottom.Add_Click({ Hide-ConsolePanel }.GetNewClosure())

# Click-to-jump on minimap
if ($ui.cnvMinimap) {
    $ui.cnvMinimap.Add_MouseLeftButtonDown({
        param($s, $e)
        $pos = $e.GetPosition($s)
        $ratio = $pos.Y / [Math]::Max($s.ActualHeight, 1)
        $target = $ratio * $ui.logScroller.ExtentHeight
        $ui.logScroller.ScrollToVerticalOffset($target)
        Update-Minimap
    }.GetNewClosure())
}
# Click-to-jump on heatmap
if ($ui.cnvHeatmap) {
    $ui.cnvHeatmap.Add_MouseLeftButtonDown({
        param($s, $e)
        $pos = $e.GetPosition($s)
        $ratio = $pos.X / [Math]::Max($s.ActualWidth, 1)
        $target = $ratio * $ui.logScroller.ExtentHeight
        $ui.logScroller.ScrollToVerticalOffset($target)
        Update-Heatmap
    }.GetNewClosure())
}
# Update viewport indicators on scroll
if ($ui.logScroller) {
    $ui.logScroller.Add_ScrollChanged({
        Update-Minimap
        Update-Heatmap
    }.GetNewClosure())
}

# -- Shimmer progress ---------------------------------------------------------
    <#
    .SYNOPSIS
        Shows the global shimmer progress bar with the given status text.
    #>
function Show-GlobalProgress([string]$text) {
    if ($ui.pnlGlobalProgress) {
        $ui.pnlGlobalProgress.Visibility = 'Visible'
        if ($ui.lblGlobalProgress) { $ui.lblGlobalProgress.Text = $text }
    }
}
    <#
    .SYNOPSIS
        Hides the global shimmer progress bar.
    #>
function Hide-GlobalProgress {
    if ($ui.pnlGlobalProgress) {
        $ui.pnlGlobalProgress.Visibility = 'Collapsed'
    }
}
    <#
    .SYNOPSIS
        Updates the global progress bar width and optional status text.
    .PARAMETER pct
        Progress fraction (0.0 to 1.0).
    #>
function Update-GlobalProgress([double]$pct, [string]$text) {
    if ($ui.prgGlobal) {
        $parentW = $ui.prgGlobal.Parent.ActualWidth
        if ($parentW -gt 0) { $ui.prgGlobal.Width = $pct * $parentW }
    }
    if ($ui.lblGlobalProgress -and $text) { $ui.lblGlobalProgress.Text = $text }
}

# -- Status bar helpers --------------------------------------------------------
    <#
    .SYNOPSIS
        Updates the status bar text and optional coloured indicator dot.
    #>
function Set-Status([string]$text, [string]$dotColor) {
    $ui.StatusText.Text = $text
    if ($dotColor -and $ui.statusDot) {
        $ui.statusDot.Fill = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.ColorConverter]::ConvertFromString($dotColor))
    }
}
# -- Preferences ---------------------------------------------------------------
$Script:Prefs = @{
    IsLightMode = $false; DisableAnimations = $false
    BasePath = 'C:\Program Files (x86)\Microsoft Group Policy'
    ExportPath = [System.IO.Path]::Combine([Environment]::GetFolderPath('Desktop'), 'ADMX_Reports')
    WindowState = 'Normal'; WindowLeft = -1; WindowTop = -1
    WindowWidth = 1440; WindowHeight = 920
}

    <#
    .SYNOPSIS
        Loads user preferences from user_prefs.json and applies them to the UI.
    .DESCRIPTION
        Restores persisted settings including theme, animation toggle, base path,
        export path, and window geometry. Falls back to defaults on error.
    #>
function Load-Preferences {
    if (Test-Path $Script:PrefsPath) {
        try {
            $saved = Get-Content $Script:PrefsPath -Raw | ConvertFrom-Json
            foreach ($prop in $saved.PSObject.Properties) {
                $Script:Prefs[$prop.Name] = $prop.Value
            }
        } catch { Write-DebugLog "Failed loading prefs: $_" -Level WARN }
    }
    # Apply to UI
    if ($ui.ThemeToggle)       { $ui.ThemeToggle.IsChecked = -not $Script:Prefs.IsLightMode }
    if ($ui.AnimationsToggle)  { $ui.AnimationsToggle.IsChecked = -not $Script:Prefs.DisableAnimations }
    if ($ui.BasePathSetting)   { $ui.BasePathSetting.Text = $Script:Prefs.BasePath }
    if ($ui.BasePathText)      { $ui.BasePathText.Text = $Script:Prefs.BasePath }
    if ($ui.ExportPathSetting) { $ui.ExportPathSetting.Text = $Script:Prefs.ExportPath }
    $Script:AnimationsDisabled = $Script:Prefs.DisableAnimations
    if ($Script:Prefs.IsLightMode) {
        Set-Theme $Script:ThemeLight
        $Script:LogLevelColors = $Script:LogLevelColorsLight
    } else {
        Set-Theme $Script:ThemeDark
        $Script:LogLevelColors = $Script:LogLevelColorsDark
    }
    Write-DebugLog 'Preferences loaded' -Level SYSTEM
}

    <#
    .SYNOPSIS
        Persists current UI state and window geometry to user_prefs.json.
    #>
function Save-Preferences {
    $Script:Prefs.IsLightMode       = -not [bool]$ui.ThemeToggle.IsChecked
    $Script:Prefs.DisableAnimations = -not [bool]$ui.AnimationsToggle.IsChecked
    if ($ui.BasePathSetting)   { $Script:Prefs.BasePath = $ui.BasePathSetting.Text }
    if ($ui.ExportPathSetting) { $Script:Prefs.ExportPath = $ui.ExportPathSetting.Text }
    if ($Window.WindowState -eq 'Normal') {
        $Script:Prefs.WindowLeft   = [int]$Window.Left
        $Script:Prefs.WindowTop    = [int]$Window.Top
        $Script:Prefs.WindowWidth  = [int]$Window.Width
        $Script:Prefs.WindowHeight = [int]$Window.Height
    } elseif ($Window.WindowState -eq 'Maximized') {
        $rb = $Window.RestoreBounds
        if ($rb -and $rb.Width -gt 0) {
            $Script:Prefs.WindowLeft   = [int]$rb.Left
            $Script:Prefs.WindowTop    = [int]$rb.Top
            $Script:Prefs.WindowWidth  = [int]$rb.Width
            $Script:Prefs.WindowHeight = [int]$rb.Height
        }
    }
    $Script:Prefs.WindowState = $Window.WindowState.ToString()
    $Script:Prefs | ConvertTo-Json -Depth 4 | Set-Content $Script:PrefsPath -Encoding UTF8
    Write-DebugLog 'Preferences saved' -Level SYSTEM
}

# -- Comparison history --------------------------------------------------------
$Script:History = @()
    <#
    .SYNOPSIS
        Loads comparison history entries from comparison_history.json.
    #>
function Load-History {
    if (Test-Path $Script:HistoryPath) {
        try { $Script:History = @(Get-Content $Script:HistoryPath -Raw | ConvertFrom-Json) } catch { $Script:History = @() }
    }
}
    <#
    .SYNOPSIS
        Saves comparison history to disk and prunes orphaned result files.
    #>
function Save-History {
    $Script:History | ConvertTo-Json -Depth 4 | Set-Content $Script:HistoryPath -Encoding UTF8
    # Prune orphaned result files
    if (Test-Path $Script:ResultsDir) {
        $validIds = @($Script:History | ForEach-Object { $_.Id } | Where-Object { $_ })
        Get-ChildItem $Script:ResultsDir -Filter '*.json' | Where-Object {
            $_.BaseName -notin $validIds
        } | Remove-Item -Force -ErrorAction SilentlyContinue
    }
}
    <#
    .SYNOPSIS
        Creates a new history entry, saves results, and trims history to 50 entries.
    #>
function Add-HistoryEntry($older, $newer, $changes, $files) {
    $id = (Get-Date).ToString('yyyyMMdd_HHmmss') + '_' + [guid]::NewGuid().ToString('N').Substring(0,8)
    $entry = [PSCustomObject]@{
        Id = $id
        OlderVersion = $older; NewerVersion = $newer
        ChangeCount = $changes; FileCount = $files
        Date = (Get-Date).ToString('yyyy-MM-dd HH:mm')
    }
    $Script:History = @($entry) + @($Script:History | Select-Object -First 49)
    Save-History
    Save-ComparisonResults $id
    $Script:LoadedHistoryId = $id
    Write-DebugLog "History entry added: $older -> $newer ($changes changes) [ID: $id]" -Level DEBUG
}

    <#
    .SYNOPSIS
        Serialises the current AllResults collection to a JSON file keyed by history ID.
    #>
function Save-ComparisonResults([string]$id) {
    if (-not (Test-Path $Script:ResultsDir)) { New-Item -Path $Script:ResultsDir -ItemType Directory -Force | Out-Null }
    $data = @($Script:AllResults | ForEach-Object {
        @{ ChangeType=$_.ChangeType; FileName=$_.FileName; StringId=$_.StringId
           OldValue=$_.OldValue; NewValue=$_.NewValue
           Category=$_.Category; StringType=$_.StringType; PolicyGroup=$_.PolicyGroup; ValueType=$_.ValueType
           RegistryKey=$_.RegistryKey; ValueName=$_.ValueName; Scope=$_.Scope }
    })
    $data | ConvertTo-Json -Depth 4 -Compress | Set-Content (Join-Path $Script:ResultsDir "$id.json") -Encoding UTF8
}

    <#
    .SYNOPSIS
        Loads a previously saved comparison result set by history ID.
    .OUTPUTS
        [bool] True if results were loaded successfully.
    #>
function Load-ComparisonResults([string]$id) {
    $path = Join-Path $Script:ResultsDir "$id.json"
    if (-not (Test-Path $path)) {
        Write-DebugLog "Results file not found for ID $id" -Level WARN
        return $false
    }
    try {
        $data = Get-Content $path -Raw | ConvertFrom-Json
        $Script:AllResults.Clear()
        foreach ($item in $data) {
            $Script:AllResults.Add([PSCustomObject]@{
                ChangeType = $item.ChangeType; FileName = $item.FileName
                StringId = $item.StringId; OldValue = $item.OldValue; NewValue = $item.NewValue
                Category = $item.Category; StringType = $item.StringType; PolicyGroup = $item.PolicyGroup; ValueType = $item.ValueType
                RegistryKey = $item.RegistryKey; ValueName = $item.ValueName; Scope = $item.Scope
            })
        }
        $Script:LoadedHistoryId = $id
        Write-DebugLog "Loaded $($Script:AllResults.Count) results for ID $id" -Level SUCCESS
        return $true
    } catch {
        Write-DebugLog "Failed to load results for ID $id - $_" -Level ERROR
        return $false
    }
}

# -- Streak tracker ------------------------------------------------------------
$Script:Streak = @{ LastDate = ''; Count = 0; TotalComparisons = 0 }
    <#
    .SYNOPSIS
        Loads the daily usage streak data from streak.json.
    #>
function Load-Streak {
    if (Test-Path $Script:StreakPath) {
        try {
            $s = Get-Content $Script:StreakPath -Raw | ConvertFrom-Json
            $Script:Streak.LastDate = $s.LastDate
            $Script:Streak.Count    = $s.Count
            $Script:Streak.TotalComparisons = $s.TotalComparisons
        } catch {}
    }
}
<# .SYNOPSIS  Persists streak data to streak.json. #>
function Save-Streak { $Script:Streak | ConvertTo-Json | Set-Content $Script:StreakPath -Encoding UTF8 }
    <#
    .SYNOPSIS
        Increments the daily usage streak counter and updates the total comparison count.
    .DESCRIPTION
        Tracks consecutive-day usage. If the last comparison was yesterday the streak
        increments; if it was today it holds; otherwise it resets to 1.
    #>
function Update-Streak {
    $today = (Get-Date).ToString('yyyy-MM-dd')
    $yesterday = (Get-Date).AddDays(-1).ToString('yyyy-MM-dd')
    if ($Script:Streak.LastDate -eq $today) {
        # same day
    } elseif ($Script:Streak.LastDate -eq $yesterday) {
        $Script:Streak.Count++
    } else {
        $Script:Streak.Count = 1
    }
    $Script:Streak.LastDate = $today
    $Script:Streak.TotalComparisons++
    Save-Streak
    Write-DebugLog "Streak updated: $($Script:Streak.Count) days, $($Script:Streak.TotalComparisons) total" -Level DEBUG
}

# -- Achievement system --------------------------------------------------------
$Script:AchievementDefs = @(
    @{ Id='first_steps';    Name='First Steps';      Emoji='&#x1F680;'; Desc='Complete your first comparison' }
    @{ Id='deep_diver';     Name='Deep Diver';       Emoji='&#x1F30A;'; Desc='Complete 5 comparisons' }
    @{ Id='policy_guru';    Name='Policy Guru';      Emoji='&#x1F9D9;'; Desc='Complete 10 comparisons' }
    @{ Id='veteran_admin';  Name='Veteran Admin';    Emoji='&#x1F396;'; Desc='Complete 25 comparisons' }
    @{ Id='centurion';      Name='Centurion';        Emoji='&#x1F3C6;'; Desc='Complete 100 comparisons' }
    @{ Id='eagle_eye';      Name='Eagle Eye';        Emoji='&#x1F985;'; Desc='Find 100+ changes in one comparison' }
    @{ Id='needle_finder';  Name='Needle Finder';    Emoji='&#x1FAA1;'; Desc='Find exactly 1 change in a comparison' }
    @{ Id='clean_slate';    Name='Clean Slate';      Emoji='&#x2728;';  Desc='Compare with zero differences' }
    @{ Id='speed_demon';    Name='Speed Demon';      Emoji='&#x26A1;';  Desc='Complete comparison in under 5 seconds' }
    @{ Id='night_owl';      Name='Night Owl';        Emoji='&#x1F989;'; Desc='Compare between midnight and 5 AM' }
    @{ Id='early_bird';     Name='Early Bird';       Emoji='&#x1F426;'; Desc='Compare between 5 AM and 7 AM' }
    @{ Id='weekend_warrior';Name='Weekend Warrior';  Emoji='&#x1F3D6;'; Desc='Compare on a weekend' }
    @{ Id='export_master';  Name='Export Master';    Emoji='&#x1F4E4;'; Desc='Export your first report' }
    @{ Id='batch_reporter'; Name='Batch Reporter';   Emoji='&#x1F4DA;'; Desc='Export 5 reports' }
    @{ Id='filter_pro';     Name='Filter Pro';       Emoji='&#x1F50D;'; Desc='Use every filter type' }
    @{ Id='theme_switcher'; Name='Theme Switcher';   Emoji='&#x1F3A8;'; Desc='Toggle between dark and light mode' }
    @{ Id='explorer';       Name='Explorer';         Emoji='&#x1F5C2;'; Desc='Browse to a custom policy folder' }
    @{ Id='bookworm';       Name='Bookworm';         Emoji='&#x1F4D6;'; Desc='Find 50+ modified strings in one run' }
    @{ Id='spring_cleaning';Name='Spring Cleaning';  Emoji='&#x1F9F9;'; Desc='Find 50+ removed strings in one run' }
    @{ Id='fresh_paint';    Name='Fresh Paint';      Emoji='&#x1F58C;'; Desc='Find 50+ added strings in one run' }
    @{ Id='streak_3';       Name='On a Roll';        Emoji='&#x1F525;'; Desc='3-day comparison streak' }
    @{ Id='streak_7';       Name='On Fire';          Emoji='&#x1F525;'; Desc='7-day comparison streak' }
    @{ Id='streak_30';      Name='Unstoppable';      Emoji='&#x1F4AA;'; Desc='30-day comparison streak' }
    @{ Id='completionist';  Name='Completionist';    Emoji='&#x1F451;'; Desc='Unlock 20 achievements' }
)

$Script:Unlocked = @{}
    <#
    .SYNOPSIS
        Loads unlocked achievement data from achievements.json.
    #>
function Load-Achievements {
    if (Test-Path $Script:AchievePath) {
        try {
            $data = Get-Content $Script:AchievePath -Raw | ConvertFrom-Json
            foreach ($p in $data.PSObject.Properties) { $Script:Unlocked[$p.Name] = $p.Value }
        } catch {}
    }
    Write-DebugLog "Loaded $($Script:Unlocked.Count) achievements" -Level SYSTEM
}
    <#
    .SYNOPSIS
        Persists unlocked achievement data to achievements.json.
    #>
function Save-Achievements {
    $Script:Unlocked | ConvertTo-Json -Depth 2 | Set-Content $Script:AchievePath -Encoding UTF8
}

    <#
    .SYNOPSIS
        Unlocks an achievement by ID, triggers toast notification and confetti.
    .DESCRIPTION
        Marks the achievement as unlocked with a timestamp, saves to disk, re-renders
        badge grids, shows a success toast with confetti, and applies a scale-bounce
        animation to the newly unlocked badge. Triggers the Completionist achievement
        if 20+ badges are unlocked.
    #>
function Unlock-Achievement([string]$id) {
    if ($Script:Unlocked.ContainsKey($id)) { return }
    $def = $Script:AchievementDefs | Where-Object { $_.Id -eq $id }
    if (-not $def) { return }
    $Script:Unlocked[$id] = (Get-Date).ToString('o')
    Save-Achievements
    Render-AchievementBadges
    $emojiText = [System.Net.WebUtility]::HtmlDecode($def.Emoji)
    Set-ToastType 'success'
    Show-Toast "$emojiText $($def.Name)" $def.Desc
    Write-DebugLog "Achievement unlocked: $($def.Name) - $($def.Desc)" -Level SUCCESS
    Start-Confetti

    # #20: Scale bounce on newly unlocked badge
    if ($ui.AchievementsPanel -and -not $Script:AnimationsDisabled) {
        foreach ($child in $ui.AchievementsPanel.Children) {
            if ($child -is [System.Windows.Controls.Border] -and $child.ToolTip -and
                $child.ToolTip.ToString().Contains($def.Name)) {
                $child.RenderTransformOrigin = [System.Windows.Point]::new(0.5, 0.5)
                $child.RenderTransform = [System.Windows.Media.ScaleTransform]::new(1, 1)
                $scaleUp = [System.Windows.Media.Animation.DoubleAnimation]::new()
                $scaleUp.From = 0.3; $scaleUp.To = 1.2
                $scaleUp.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(250))
                $scaleUp.AutoReverse = $true
                $scaleUp.EasingFunction = [System.Windows.Media.Animation.CubicEase]::new()
                $scaleUp.EasingFunction.EasingMode = [System.Windows.Media.Animation.EasingMode]::EaseOut
                $child.RenderTransform.BeginAnimation([System.Windows.Media.ScaleTransform]::ScaleXProperty, $scaleUp)
                $child.RenderTransform.BeginAnimation([System.Windows.Media.ScaleTransform]::ScaleYProperty, $scaleUp)
                $child.Effect = [System.Windows.Media.Effects.DropShadowEffect]@{
                    Color = [System.Windows.Media.ColorConverter]::ConvertFromString('#0078D4')
                    BlurRadius = 12; ShadowDepth = 0; Opacity = 0.5
                }
                $glowRef = $child
                $clearGlow = [System.Windows.Threading.DispatcherTimer]::new()
                $clearGlow.Interval = [TimeSpan]::FromMilliseconds(800)
                $clearGlow.Add_Tick({
                    $this.Stop()
                    $glowRef.Effect = $null
                }.GetNewClosure())
                $clearGlow.Start()
                break
            }
        }
    }
    if ($Script:Unlocked.Count -ge 20 -and -not $Script:Unlocked.ContainsKey('completionist')) {
        Unlock-Achievement 'completionist'
    }
}

$Script:ExportCount = 0
$Script:FiltersUsed = [System.Collections.Generic.HashSet[string]]::new()

    <#
    .SYNOPSIS
        Evaluates post-comparison metrics and unlocks any matching achievements.
    .DESCRIPTION
        Called after each comparison run. Checks total comparisons, change thresholds,
        time-of-day, day-of-week, streak length, and per-type counts against all
        achievement definitions.
    #>
function Check-ComparisonAchievements($changeCount, $addedCount, $removedCount, $modifiedCount, $elapsed) {
    $total = $Script:Streak.TotalComparisons
    if ($total -ge 1)   { Unlock-Achievement 'first_steps' }
    if ($total -ge 5)   { Unlock-Achievement 'deep_diver' }
    if ($total -ge 10)  { Unlock-Achievement 'policy_guru' }
    if ($total -ge 25)  { Unlock-Achievement 'veteran_admin' }
    if ($total -ge 100) { Unlock-Achievement 'centurion' }
    if ($changeCount -ge 100) { Unlock-Achievement 'eagle_eye' }
    if ($changeCount -eq 1)   { Unlock-Achievement 'needle_finder' }
    if ($changeCount -eq 0)   { Unlock-Achievement 'clean_slate' }
    if ($elapsed.TotalSeconds -lt 5) { Unlock-Achievement 'speed_demon' }
    $hour = (Get-Date).Hour
    if ($hour -ge 0 -and $hour -lt 5) { Unlock-Achievement 'night_owl' }
    if ($hour -ge 5 -and $hour -lt 7) { Unlock-Achievement 'early_bird' }
    if ((Get-Date).DayOfWeek -in 'Saturday','Sunday') { Unlock-Achievement 'weekend_warrior' }
    if ($modifiedCount -ge 50) { Unlock-Achievement 'bookworm' }
    if ($removedCount -ge 50)  { Unlock-Achievement 'spring_cleaning' }
    if ($addedCount -ge 50)    { Unlock-Achievement 'fresh_paint' }
    if ($Script:Streak.Count -ge 3)  { Unlock-Achievement 'streak_3' }
    if ($Script:Streak.Count -ge 7)  { Unlock-Achievement 'streak_7' }
    if ($Script:Streak.Count -ge 30) { Unlock-Achievement 'streak_30' }
}

    <#
    .SYNOPSIS
        Renders achievement badge grids in the Settings panel and Dashboard sidebar.
    .DESCRIPTION
        Creates emoji badge borders (unlocked=coloured, locked=muted) with hover scale
        animations for both the full Settings panel and the compact Dashboard sidebar.
        Updates the achievement count label and animated progress bar.
    #>
function Render-AchievementBadges {
    # Settings panel achievements
    $wrap = $ui.AchievementsPanel
    if ($wrap) {
        $wrap.Children.Clear()
        foreach ($def in $Script:AchievementDefs) {
            $isUnlocked = $Script:Unlocked.ContainsKey($def.Id)
            $badge = [System.Windows.Controls.Border]::new()
            $badge.Width  = 36; $badge.Height = 36
            $badge.CornerRadius = [System.Windows.CornerRadius]::new(6)
            $badge.Margin = [System.Windows.Thickness]::new(0,0,4,4)
            $badge.Cursor = [System.Windows.Input.Cursors]::Hand
            $badge.RenderTransformOrigin = [System.Windows.Point]::new(0.5, 0.5)
            $badge.RenderTransform = [System.Windows.Media.ScaleTransform]::new(1, 1)
            $badgeRef = $badge
            $badge.Add_MouseEnter({
                if (-not $Script:AnimationsDisabled) {
                    $scale = [System.Windows.Media.Animation.DoubleAnimation]::new(1, 1.15,
                        [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(100)))
                    $scale.EasingFunction = [System.Windows.Media.Animation.CubicEase]::new()
                    $this.RenderTransform.BeginAnimation([System.Windows.Media.ScaleTransform]::ScaleXProperty, $scale)
                    $this.RenderTransform.BeginAnimation([System.Windows.Media.ScaleTransform]::ScaleYProperty, $scale)
                }
            }.GetNewClosure())
            $badge.Add_MouseLeave({
                if (-not $Script:AnimationsDisabled) {
                    $scale = [System.Windows.Media.Animation.DoubleAnimation]::new(1.15, 1,
                        [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(100)))
                    $this.RenderTransform.BeginAnimation([System.Windows.Media.ScaleTransform]::ScaleXProperty, $scale)
                    $this.RenderTransform.BeginAnimation([System.Windows.Media.ScaleTransform]::ScaleYProperty, $scale)
                }
            }.GetNewClosure())
            if ($isUnlocked) {
                $badge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeSelectedBg')
                $badge.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeAccent')
                $badge.BorderThickness = [System.Windows.Thickness]::new(1)
                $tb = [System.Windows.Controls.TextBlock]::new()
                $tb.Text = [System.Net.WebUtility]::HtmlDecode($def.Emoji)
                $tb.FontSize = 16
                $tb.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextPrimary')
                $tb.HorizontalAlignment = 'Center'
                $tb.VerticalAlignment   = 'Center'
                $badge.Child = $tb
                $badge.ToolTip = "$($def.Name) - $($def.Desc)"
            } else {
                $badge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeDeepBg')
                $badge.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderSubtle')
                $badge.BorderThickness = [System.Windows.Thickness]::new(1)
                $tb = [System.Windows.Controls.TextBlock]::new()
                $tb.Text = '?'
                $tb.FontSize = 14; $tb.FontWeight = 'Bold'
                $tb.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
                $tb.HorizontalAlignment = 'Center'
                $tb.VerticalAlignment   = 'Center'
                $badge.Child = $tb
                $badge.ToolTip = '???'
            }
            $wrap.Children.Add($badge)
        }
    }

    # Dashboard sidebar badges
    $dashWrap = $ui.DashAchievementWrap
    if ($dashWrap) {
        $dashWrap.Children.Clear()
        foreach ($def in $Script:AchievementDefs) {
            $isUnlocked = $Script:Unlocked.ContainsKey($def.Id)
            $badge = [System.Windows.Controls.Border]::new()
            $badge.Width  = 28; $badge.Height = 28
            $badge.CornerRadius = [System.Windows.CornerRadius]::new(4)
            $badge.Margin = [System.Windows.Thickness]::new(0,0,2,2)
            if ($isUnlocked) {
                $badge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeSelectedBg')
                $badge.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeAccent')
                $badge.BorderThickness = [System.Windows.Thickness]::new(1)
                $tb = [System.Windows.Controls.TextBlock]::new()
                $tb.Text = [System.Net.WebUtility]::HtmlDecode($def.Emoji)
                $tb.FontSize = 12
                $tb.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextPrimary')
                $tb.HorizontalAlignment = 'Center'; $tb.VerticalAlignment = 'Center'
                $badge.Child = $tb
                $badge.ToolTip = $def.Name
            } else {
                $badge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeDeepBg')
                $badge.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorderSubtle')
                $badge.BorderThickness = [System.Windows.Thickness]::new(1)
                $tb = [System.Windows.Controls.TextBlock]::new()
                $tb.Text = '?'; $tb.FontSize = 10; $tb.FontWeight = 'Bold'
                $tb.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
                $tb.HorizontalAlignment = 'Center'; $tb.VerticalAlignment = 'Center'
                $badge.Child = $tb
            }
            $dashWrap.Children.Add($badge)
        }
    }

    $countText = "$($Script:Unlocked.Count) / $($Script:AchievementDefs.Count)"
    if ($ui.DashAchievementCount)  { $ui.DashAchievementCount.Text = $countText }
    if ($ui.AchievementSummaryText) { $ui.AchievementSummaryText.Text = $countText }

    # Progress bar
    if ($ui.AchievementProgressBar -and $ui.AchievementProgressBar.Parent) {
        $parentW = $ui.AchievementProgressBar.Parent.ActualWidth
        if ($parentW -gt 0 -and $Script:AchievementDefs.Count -gt 0) {
            $pct = $Script:Unlocked.Count / $Script:AchievementDefs.Count
            Set-AnimatedWidth $ui.AchievementProgressBar ($pct * $parentW)
        }
    }
}
# -- Toast notification --------------------------------------------------------
$Script:ToastTimer = $null
# #8: RichTextBox helper — replaces plain .Text property usage
    <#
    .SYNOPSIS
        Replaces the comparison log RichTextBox content with colour-coded formatted text.
    .DESCRIPTION
        Parses each line for keywords (ERROR, File Added, File Removed, Added, Removed,
        Modified, summary) and applies themed colour formatting with prefix icons.
    #>
function Set-LogText([string]$text) {
    if (-not $ui.ComparisonLogText) { return }
    $doc = $ui.ComparisonLogText.Document
    $doc.Blocks.Clear()
    $converter = [System.Windows.Media.BrushConverter]::new()
    foreach ($line in ($text -split "`n")) {
        $para = [System.Windows.Documents.Paragraph]::new()
        $para.Margin = [System.Windows.Thickness]::new(0,1,0,1)
        if ($line -match '^\s*ERROR') {
            $icon = [System.Windows.Documents.Run]::new([char]0x26A0 + ' ')
            $icon.Foreground = $Window.Resources['ThemeError']
            $icon.FontSize = 13
            $para.Inlines.Add($icon)
            $run = [System.Windows.Documents.Run]::new($line)
            $run.Foreground = $Window.Resources['ThemeError']
            $para.Inlines.Add($run)
        } elseif ($line -match '^\s*(File Added|FILE ADDED)') {
            $para.Margin = [System.Windows.Thickness]::new(0,6,0,2)
            $icon = [System.Windows.Documents.Run]::new([char]0x2295 + ' ')
            $icon.Foreground = $converter.ConvertFromString('#00E676')
            $icon.FontWeight = 'Bold'; $icon.FontSize = 14
            $para.Inlines.Add($icon)
            $run = [System.Windows.Documents.Run]::new($line)
            $run.Foreground = $converter.ConvertFromString('#00E676')
            $run.FontWeight = 'Bold'
            $para.Inlines.Add($run)
        } elseif ($line -match '^\s*(File Removed|FILE REMOVED)') {
            $para.Margin = [System.Windows.Thickness]::new(0,6,0,2)
            $icon = [System.Windows.Documents.Run]::new([char]0x2296 + ' ')
            $icon.Foreground = $converter.ConvertFromString('#FF6E40')
            $icon.FontWeight = 'Bold'; $icon.FontSize = 14
            $para.Inlines.Add($icon)
            $run = [System.Windows.Documents.Run]::new($line)
            $run.Foreground = $converter.ConvertFromString('#FF6E40')
            $run.FontWeight = 'Bold'
            $para.Inlines.Add($run)
        } elseif ($line -match '^\s*(Added|New|ADDED)') {
            $icon = [System.Windows.Documents.Run]::new('+ ')
            $icon.Foreground = $Window.Resources['ThemeSuccess']
            $icon.FontWeight = 'Bold'
            $para.Inlines.Add($icon)
            $run = [System.Windows.Documents.Run]::new($line)
            $run.Foreground = $Window.Resources['ThemeSuccess']
            $para.Inlines.Add($run)
        } elseif ($line -match '^\s*(Removed|REMOVED)') {
            $icon = [System.Windows.Documents.Run]::new('- ')
            $icon.Foreground = $Window.Resources['ThemeError']
            $icon.FontWeight = 'Bold'
            $para.Inlines.Add($icon)
            $run = [System.Windows.Documents.Run]::new($line)
            $run.Foreground = $Window.Resources['ThemeError']
            $para.Inlines.Add($run)
        } elseif ($line -match '^\s*(Modified|Changed|MODIFIED)') {
            $icon = [System.Windows.Documents.Run]::new('~ ')
            $icon.Foreground = $Window.Resources['ThemeWarning']
            $icon.FontWeight = 'Bold'
            $para.Inlines.Add($icon)
            $run = [System.Windows.Documents.Run]::new($line)
            $run.Foreground = $Window.Resources['ThemeWarning']
            $para.Inlines.Add($run)
        } elseif ($line -match '^\s*---' -or $line -match '^\s*===') {
            $run = [System.Windows.Documents.Run]::new($line)
            $run.Foreground = $Window.Resources['ThemeBorderElevated']
            $para.Inlines.Add($run)
        } elseif ($line -match '^\s*Processing|^\s*Scanning|^\s*Comparing') {
            $run = [System.Windows.Documents.Run]::new($line)
            $run.Foreground = $converter.ConvertFromString('#60CDFF')
            $para.Inlines.Add($run)
        } elseif ($line -match '^\s*Found\s+\d+|^\s*Total:') {
            $run = [System.Windows.Documents.Run]::new($line)
            $run.Foreground = $converter.ConvertFromString('#A78BFA')
            $run.FontWeight = 'SemiBold'
            $para.Inlines.Add($run)
        } else {
            $run = [System.Windows.Documents.Run]::new($line)
            $para.Inlines.Add($run)
        }
        $doc.Blocks.Add($para)
    }
    $ui.ComparisonLogText.ScrollToEnd()
}
    <#
    .SYNOPSIS
        Appends colour-coded formatted text to the comparison log RichTextBox.
    #>
function Append-LogText([string]$text) {
    if (-not $ui.ComparisonLogText) { return }
    $doc = $ui.ComparisonLogText.Document
    $converter = [System.Windows.Media.BrushConverter]::new()
    foreach ($line in ($text -split "`n")) {
        $para = [System.Windows.Documents.Paragraph]::new()
        $para.Margin = [System.Windows.Thickness]::new(0,1,0,1)
        if ($line -match '^\s*ERROR') {
            $icon = [System.Windows.Documents.Run]::new([char]0x26A0 + ' ')
            $icon.Foreground = $Window.Resources['ThemeError']
            $icon.FontSize = 13
            $para.Inlines.Add($icon)
            $run = [System.Windows.Documents.Run]::new($line)
            $run.Foreground = $Window.Resources['ThemeError']
            $para.Inlines.Add($run)
        } elseif ($line -match '^\s*(File Added|FILE ADDED)') {
            $para.Margin = [System.Windows.Thickness]::new(0,6,0,2)
            $icon = [System.Windows.Documents.Run]::new([char]0x2295 + ' ')
            $icon.Foreground = $converter.ConvertFromString('#00E676')
            $icon.FontWeight = 'Bold'; $icon.FontSize = 14
            $para.Inlines.Add($icon)
            $run = [System.Windows.Documents.Run]::new($line)
            $run.Foreground = $converter.ConvertFromString('#00E676')
            $run.FontWeight = 'Bold'
            $para.Inlines.Add($run)
        } elseif ($line -match '^\s*(File Removed|FILE REMOVED)') {
            $para.Margin = [System.Windows.Thickness]::new(0,6,0,2)
            $icon = [System.Windows.Documents.Run]::new([char]0x2296 + ' ')
            $icon.Foreground = $converter.ConvertFromString('#FF6E40')
            $icon.FontWeight = 'Bold'; $icon.FontSize = 14
            $para.Inlines.Add($icon)
            $run = [System.Windows.Documents.Run]::new($line)
            $run.Foreground = $converter.ConvertFromString('#FF6E40')
            $run.FontWeight = 'Bold'
            $para.Inlines.Add($run)
        } elseif ($line -match '^\s*(Added|New|ADDED)') {
            $icon = [System.Windows.Documents.Run]::new('+ ')
            $icon.Foreground = $Window.Resources['ThemeSuccess']
            $icon.FontWeight = 'Bold'
            $para.Inlines.Add($icon)
            $run = [System.Windows.Documents.Run]::new($line)
            $run.Foreground = $Window.Resources['ThemeSuccess']
            $para.Inlines.Add($run)
        } elseif ($line -match '^\s*(Removed|REMOVED)') {
            $icon = [System.Windows.Documents.Run]::new('- ')
            $icon.Foreground = $Window.Resources['ThemeError']
            $icon.FontWeight = 'Bold'
            $para.Inlines.Add($icon)
            $run = [System.Windows.Documents.Run]::new($line)
            $run.Foreground = $Window.Resources['ThemeError']
            $para.Inlines.Add($run)
        } elseif ($line -match '^\s*(Modified|Changed|MODIFIED)') {
            $icon = [System.Windows.Documents.Run]::new('~ ')
            $icon.Foreground = $Window.Resources['ThemeWarning']
            $icon.FontWeight = 'Bold'
            $para.Inlines.Add($icon)
            $run = [System.Windows.Documents.Run]::new($line)
            $run.Foreground = $Window.Resources['ThemeWarning']
            $para.Inlines.Add($run)
        } elseif ($line -match '^\s*---' -or $line -match '^\s*===') {
            $run = [System.Windows.Documents.Run]::new($line)
            $run.Foreground = $Window.Resources['ThemeBorderElevated']
            $para.Inlines.Add($run)
        } elseif ($line -match '^\s*Processing|^\s*Scanning|^\s*Comparing') {
            $run = [System.Windows.Documents.Run]::new($line)
            $run.Foreground = $converter.ConvertFromString('#60CDFF')
            $para.Inlines.Add($run)
        } elseif ($line -match '^\s*Found\s+\d+|^\s*Total:') {
            $run = [System.Windows.Documents.Run]::new($line)
            $run.Foreground = $converter.ConvertFromString('#A78BFA')
            $run.FontWeight = 'SemiBold'
            $para.Inlines.Add($run)
        } else {
            $run = [System.Windows.Documents.Run]::new($line)
            $para.Inlines.Add($run)
        }
        $doc.Blocks.Add($para)
    }
    $ui.ComparisonLogText.ScrollToEnd()
}
# #39: Version badge subtle opacity pulse on data refresh
    <#
    .SYNOPSIS
        Triggers a quick opacity pulse on the version badge in the title bar.
    #>
function Invoke-VersionBadgeBlink {
    if ($Script:AnimationsDisabled -or -not $ui.VersionBadge) { return }
    $pulse = [System.Windows.Media.Animation.DoubleAnimation]::new()
    $pulse.From = 1.0; $pulse.To = 0.3
    $pulse.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(200))
    $pulse.AutoReverse = $true
    $ui.VersionBadge.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $pulse)
}

# #40: Button loading spinner helpers
    <#
    .SYNOPSIS
        Puts a button into a loading state with a spinning icon and disabled interaction.
    #>
function Set-ButtonLoading([System.Windows.Controls.Button]$btn) {
    $btn.IsEnabled = $false
    if ($btn.Content -is [System.Windows.Controls.StackPanel]) {
        $sp = $btn.Content
        $iconTb = $sp.Children | Where-Object { $_ -is [System.Windows.Controls.TextBlock] } | Select-Object -First 1
        if ($iconTb) {
            $Script:OrigButtonIcon = $iconTb.Text
            $iconTb.Text = [char]0xF16A  # Sync/refresh icon
            $iconTb.RenderTransformOrigin = [System.Windows.Point]::new(0.5, 0.5)
            $iconTb.RenderTransform = [System.Windows.Media.RotateTransform]::new(0)
            $rot = [System.Windows.Media.Animation.DoubleAnimation]::new()
            $rot.From = 0; $rot.To = 360
            $rot.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(1000))
            $rot.RepeatBehavior = [System.Windows.Media.Animation.RepeatBehavior]::Forever
            $iconTb.RenderTransform.BeginAnimation([System.Windows.Media.RotateTransform]::AngleProperty, $rot)
        }
    }
}
    <#
    .SYNOPSIS
        Restores a button from loading state back to its original icon and enabled state.
    #>
function Clear-ButtonLoading([System.Windows.Controls.Button]$btn) {
    $btn.IsEnabled = $true
    if ($btn.Content -is [System.Windows.Controls.StackPanel]) {
        $sp = $btn.Content
        $iconTb = $sp.Children | Where-Object { $_ -is [System.Windows.Controls.TextBlock] } | Select-Object -First 1
        if ($iconTb -and $Script:OrigButtonIcon) {
            $iconTb.RenderTransform.BeginAnimation([System.Windows.Media.RotateTransform]::AngleProperty, $null)
            $iconTb.Text = $Script:OrigButtonIcon
        }
    }
}
    <#
    .SYNOPSIS
        Displays an animated toast notification with auto-dismiss and click-to-dismiss.
    .DESCRIPTION
        Shows a slide-in notification at the top-right with a type-specific icon and colour.
        Auto-dismisses after 4 seconds with a slide-out animation. Supports click-to-dismiss.
    .PARAMETER type
        Toast style: info, success, warning, or error.
    #>
function Show-Toast([string]$title, [string]$message, [string]$type = 'info') {
    Set-ToastType $type
    if (-not $ui.ToastBorder) { return }
    $ui.ToastTitle.Text   = $title
    $ui.ToastMessage.Text = $message
    $ui.ToastBorder.Visibility = 'Visible'

    # #33: Toast dismiss on click
    $ui.ToastBorder.IsHitTestVisible = $true

    $tt = $ui.ToastBorder.RenderTransform
    if (-not $Script:AnimationsDisabled -and $tt -is [System.Windows.Media.TranslateTransform]) {
        $tt.X = 400
        $anim = [System.Windows.Media.Animation.DoubleAnimation]::new(400, 0,
            [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(350)))
        $anim.EasingFunction = [System.Windows.Media.Animation.CubicEase]::new()
        $anim.EasingFunction.EasingMode = 'EaseOut'
        $tt.BeginAnimation([System.Windows.Media.TranslateTransform]::XProperty, $anim)
    } elseif ($tt -is [System.Windows.Media.TranslateTransform]) {
        $tt.X = 0
    }

    if ($Script:ToastTimer) { $Script:ToastTimer.Stop() }
    if ($Script:ToastHideTimer) { $Script:ToastHideTimer.Stop() }

    # Both timers created at SAME scope level (flat) — no nested .GetNewClosure()
    $ToastRef = $ui.ToastBorder

    # Pre-create hide timer (runs after slide-out animation finishes)
    $Script:ToastHideTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $Script:ToastHideTimer.Interval = [TimeSpan]::FromMilliseconds(350)
    $Script:ToastHideTimer.Add_Tick({
        $this.Stop()
        $ToastRef.Visibility = 'Collapsed'
    }.GetNewClosure())

    # Main dismiss timer (starts the slide-out, then starts the hide timer)
    $HideTimerRef = $Script:ToastHideTimer
    $Script:ToastTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $Script:ToastTimer.Interval = [TimeSpan]::FromSeconds(4)
    $Script:ToastTimer.Add_Tick({
        $this.Stop()
        $ttInner = $ToastRef.RenderTransform
        if (-not $Script:AnimationsDisabled -and $ttInner -is [System.Windows.Media.TranslateTransform]) {
            $anim = [System.Windows.Media.Animation.DoubleAnimation]::new(0, 400,
                [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(300)))
            $anim.EasingFunction = [System.Windows.Media.Animation.CubicEase]::new()
            $anim.EasingFunction.EasingMode = 'EaseIn'
            $ttInner.BeginAnimation([System.Windows.Media.TranslateTransform]::XProperty, $anim)
            $HideTimerRef.Start()
        } else {
            $ToastRef.Visibility = 'Collapsed'
        }
    }.GetNewClosure())
    $Script:ToastTimer.Start()

    # #33: Toast click-to-dismiss
    if (-not $Script:ToastClickWired) {
        $Script:ToastClickWired = $true
        $ToastBorderRef = $ui.ToastBorder
        $ui.ToastBorder.Add_MouseLeftButtonDown({
            if ($Script:ToastTimer) { $Script:ToastTimer.Stop() }
            $ttInner = $ToastBorderRef.RenderTransform
            if (-not $Script:AnimationsDisabled -and $ttInner -is [System.Windows.Media.TranslateTransform]) {
                $anim = [System.Windows.Media.Animation.DoubleAnimation]::new(0, 400,
                    [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(200)))
                $anim.EasingFunction = [System.Windows.Media.Animation.CubicEase]::new()
                $anim.EasingFunction.EasingMode = 'EaseIn'
                $ttInner.BeginAnimation([System.Windows.Media.TranslateTransform]::XProperty, $anim)
                if ($Script:ToastHideTimer) { $Script:ToastHideTimer.Start() }
            } else {
                $ToastBorderRef.Visibility = 'Collapsed'
            }
        }.GetNewClosure())
    }
    Write-DebugLog "Toast: $title - $message" -Level DEBUG
}

# -- Confetti ------------------------------------------------------------------
    <#
    .SYNOPSIS
        Launches a confetti particle animation across the full window canvas.
    .DESCRIPTION
        Creates 55 randomly coloured, sized, and rotated rectangles that fall with
        gravity-eased animation and horizontal drift. Auto-cleans up after 5 seconds
        with a fade-out transition.
    #>
function Start-Confetti {
    if ($Script:AnimationsDisabled) { return }
    if (-not $ui.ConfettiCanvas) { return }
    $canvas = $ui.ConfettiCanvas
    $canvas.Children.Clear()
    $canvas.Visibility = 'Visible'
    $colors = @('#FF4444','#FFD700','#00C853','#60CDFF','#0078D4','#B388FF','#FF6D00','#E040FB')
    $rng = [System.Random]::new()
    $w = $Window.ActualWidth; $h = $Window.ActualHeight
    if ($w -lt 100) { $w = 1440 }; if ($h -lt 100) { $h = 920 }

    for ($i = 0; $i -lt 55; $i++) {
        $rect = [System.Windows.Shapes.Rectangle]::new()
        $sz = $rng.Next(4,11)
        $rect.Width  = $sz
        $rect.Height = $sz * (0.5 + $rng.NextDouble())
        $rect.Fill   = [System.Windows.Media.BrushConverter]::new().ConvertFromString($colors[$rng.Next($colors.Count)])
        $rect.Opacity = 0.9
        $rect.RadiusX = 1; $rect.RadiusY = 1
        $rect.RenderTransform = [System.Windows.Media.RotateTransform]::new(0, $sz/2, $sz/2)
        $startX = $rng.NextDouble() * $w
        [System.Windows.Controls.Canvas]::SetLeft($rect, $startX)
        [System.Windows.Controls.Canvas]::SetTop($rect, -20)
        $canvas.Children.Add($rect)

        $dur = 2000 + $rng.Next(2500)
        $ts  = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds($dur))
        $ease = [System.Windows.Media.Animation.CubicEase]::new()
        $ease.EasingMode = 'EaseIn'

        $animY = [System.Windows.Media.Animation.DoubleAnimation]::new(-20, $h + 20, $ts)
        $animY.EasingFunction = $ease
        $rect.BeginAnimation([System.Windows.Controls.Canvas]::TopProperty, $animY)

        $drift = $rng.NextDouble() * 240 - 120
        $animX = [System.Windows.Media.Animation.DoubleAnimation]::new($startX, $startX + $drift, $ts)
        $rect.BeginAnimation([System.Windows.Controls.Canvas]::LeftProperty, $animX)

        $animR = [System.Windows.Media.Animation.DoubleAnimation]::new(0, $rng.Next(-720,720), $ts)
        $rect.RenderTransform.BeginAnimation([System.Windows.Media.RotateTransform]::AngleProperty, $animR)
    }

    $CanvasRef = $canvas
    $cleanupTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $cleanupTimer.Interval = [TimeSpan]::FromSeconds(5)
    $cleanupTimer.Add_Tick({
        $this.Stop()
        $CanvasRef.Children.Clear()
        # #34: Fade confetti canvas out instead of hard collapse
    if (-not $Script:AnimationsDisabled) {
        $fadeOut = [System.Windows.Media.Animation.DoubleAnimation]::new(1, 0,
            [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(400)))
        $fadeOut.EasingFunction = [System.Windows.Media.Animation.CubicEase]::new()
        $cRef = $CanvasRef
        $fadeOut.Add_Completed({ $cRef.Visibility = 'Collapsed'; $cRef.Opacity = 1 }.GetNewClosure())
        $CanvasRef.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $fadeOut)
    } else {
        $CanvasRef.Visibility = 'Collapsed'
    }
    }.GetNewClosure())
    $cleanupTimer.Start()
}

# -- Tab navigation ------------------------------------------------------------
$Script:Tabs = @('Dashboard','Compare','Results','Settings')
$Script:ActiveTab = 'Dashboard'

    <#
    .SYNOPSIS
        Applies a quick opacity fade-in animation when showing a tab panel.
    #>
function Invoke-TabFade([System.Windows.UIElement]$Panel) {
    $Panel.Visibility = 'Visible'
    if ($Script:AnimationsDisabled) { $Panel.Opacity = 1; return }
    $Panel.Opacity = 0
    $Fade = [System.Windows.Media.Animation.DoubleAnimation]::new()
    $Fade.From     = 0
    $Fade.To       = 1
    $Fade.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(150))
    $Fade.EasingFunction = [System.Windows.Media.Animation.CubicEase]::new()
    $Fade.EasingFunction.EasingMode = 'EaseOut'
    $Panel.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $Fade)
}

# -- Animated Progress Helper (Robinhood-style smooth fill) --------------------
    <#
    .SYNOPSIS
        Smoothly animates a FrameworkElement width to the target value.
    #>
function Set-AnimatedWidth([System.Windows.FrameworkElement]$element, [double]$targetWidth) {
    if ($Script:AnimationsDisabled -or $targetWidth -le 0) {
        $element.Width = $targetWidth
        return
    }
    $anim = [System.Windows.Media.Animation.DoubleAnimation]::new()
    $anim.To = $targetWidth
    $anim.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(250))
    $anim.EasingFunction = [System.Windows.Media.Animation.QuadraticEase]::new()
    $anim.EasingFunction.EasingMode = 'EaseOut'
    $element.BeginAnimation([System.Windows.FrameworkElement]::WidthProperty, $anim)
}
# #24: Stat card count-up animation (fake timer-based number increment)
$Script:CountUpTimers = @{}
    <#
    .SYNOPSIS
        Animates a TextBlock numeric value from 0 up to the target using eased increments.
    .DESCRIPTION
        Uses a DispatcherTimer to increment the displayed number over ~375ms with
        quadratic easing. Previous animations on the same TextBlock are stopped first.
    #>
function Invoke-CountUp([System.Windows.Controls.TextBlock]$textBlock, [int]$target) {
    # Stop any previous timer animating this TextBlock
    $key = $textBlock.Name
    if ($key -and $Script:CountUpTimers[$key]) {
        $Script:CountUpTimers[$key].Stop()
        $Script:CountUpTimers.Remove($key)
    }
    if ($Script:AnimationsDisabled -or $target -le 0) { $textBlock.Text = "$target"; return }
    $steps = [Math]::Min(15, $target)
    $step = 0
    $tbRef = $textBlock; $tgtRef = $target; $stepsRef = $steps
    $timer = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [TimeSpan]::FromMilliseconds(25)
    $timer.Add_Tick({
        $step++
        $progress = [Math]::Min(1.0, $step / $stepsRef)
        $eased = 1 - (1 - $progress) * (1 - $progress)
        $val = [Math]::Round($tgtRef * $eased)
        $tbRef.Text = "$val"
        if ($step -ge $stepsRef) {
            $this.Stop()
            $tbRef.Text = "$tgtRef"
            if ($key -and $Script:CountUpTimers) { $Script:CountUpTimers.Remove($key) }
        }
    }.GetNewClosure())
    if ($key) { $Script:CountUpTimers[$key] = $timer }
    $timer.Start()
}

# #28: Stat card subtle breathe animation (very subtle opacity pulse on idle)
    <#
    .SYNOPSIS
        Applies a subtle continuous opacity pulse to the four dashboard stat cards.
    #>
function Start-CardBreathe {
    if ($Script:AnimationsDisabled) { return }
    foreach ($name in @('StatCard1','StatCard2','StatCard3','StatCard4')) {
        $card = $ui[$name]
        if ($card) {
            $breathe = [System.Windows.Media.Animation.DoubleAnimation]::new()
            $breathe.From = 0.92; $breathe.To = 1.0
            $breathe.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(3000))
            $breathe.AutoReverse = $true
            $breathe.RepeatBehavior = [System.Windows.Media.Animation.RepeatBehavior]::Forever
            $breathe.EasingFunction = [System.Windows.Media.Animation.SineEase]::new()
            $breathe.EasingFunction.EasingMode = 'EaseInOut'
            $card.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $breathe)
        }
    }
}

# #32: Update filter pill text with counts
    <#
    .SYNOPSIS
        Updates the filter pill button labels with current per-type result counts.
    #>
function Update-FilterPillCounts {
    if (-not $Script:AllResults -or $Script:AllResults.Count -eq 0) { return }
    $added = ($Script:AllResults | Where-Object { $_.ChangeType -eq 'Added' }).Count
    $removed = ($Script:AllResults | Where-Object { $_.ChangeType -eq 'Removed' }).Count
    $modified = ($Script:AllResults | Where-Object { $_.ChangeType -eq 'Modified' }).Count
    $fAdded = ($Script:AllResults | Where-Object { $_.ChangeType -eq 'File Added' }).Count
    $fRemoved = ($Script:AllResults | Where-Object { $_.ChangeType -eq 'File Removed' }).Count
    # Update FilterAll content
    if ($ui.FilterAll) {
        $ui.FilterAll.Content = "All ($($Script:AllResults.Count))"
    }
    # Update each pill — find the TextBlock child and update
    foreach ($pair in @(
        @($ui.FilterAdded, "Added ($added)"),
        @($ui.FilterRemoved, "Removed ($removed)"),
        @($ui.FilterModified, "Modified ($modified)"),
        @($ui.FilterFileAdded, "File Added ($fAdded)"),
        @($ui.FilterFileRemoved, "File Removed ($fRemoved)")
    )) {
        $pill = $pair[0]; $label = $pair[1]
        if ($pill -and $pill.Content -is [System.Windows.Controls.StackPanel]) {
            $sp = $pill.Content
            $tb = $sp.Children | Where-Object { $_ -is [System.Windows.Controls.TextBlock] } | Select-Object -Last 1
            if ($tb) { $tb.Text = $label }
        }
    }
}


# -- Status Dot Pulse Animation ------------------------------------------------
    <#
    .SYNOPSIS
        Starts a repeating opacity pulse animation on the status bar indicator dot.
    #>
function Start-StatusPulse {
    if ($Script:AnimationsDisabled) { return }
    $anim = [System.Windows.Media.Animation.DoubleAnimation]::new()
    $anim.From = 0.4; $anim.To = 1.0
    $anim.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(800))
    $anim.AutoReverse = $true
    $anim.RepeatBehavior = [System.Windows.Media.Animation.RepeatBehavior]::Forever
    $ui.statusDot.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $anim)
}
    <#
    .SYNOPSIS
        Stops the status bar dot pulse animation and resets opacity to 1.0.
    #>
function Stop-StatusPulse {
    $ui.statusDot.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $null)
    $ui.statusDot.Opacity = 1.0
}
    <#
    .SYNOPSIS
        Navigates to the specified tab, showing/hiding panels and sidebars with fade transitions.
    .PARAMETER tab
        Tab name: Dashboard, Compare, Results, or Settings.
    #>
function Switch-Tab([string]$tab) {
    $Script:ActiveTab = $tab
    foreach ($t in $Script:Tabs) {
        $panel   = $ui["Panel$t"]
        $sidebar = $ui["Sidebar$t"]
        $nav     = $ui["Nav$t"]
        if ($t -eq $tab) {
            if ($panel)   { Invoke-TabFade $panel }
            if ($sidebar) { Invoke-TabFade $sidebar }
        } else {
            # #21: Fade-out before collapse
            if ($panel -and $panel.Visibility -eq 'Visible') {
                if (-not $Script:AnimationsDisabled) {
                    $panel.Opacity = 0; $panel.Visibility = 'Collapsed'
                } else {
                    $panel.Visibility = 'Collapsed'
                }
            }
            if ($sidebar -and $sidebar.Visibility -eq 'Visible') {
                if (-not $Script:AnimationsDisabled) {
                    $sidebar.Opacity = 0; $sidebar.Visibility = 'Collapsed'
                } else {
                    $sidebar.Visibility = 'Collapsed'
                }
            }
        }
        if ($nav) { $nav.Tag = if ($t -eq $tab) { 'Active' } else { $null } }
    }
    Set-Status $tab '#00C853'
    Write-DebugLog "Navigated to $tab" -Level STEP
}

$ui.NavDashboard.Add_Click({ Switch-Tab 'Dashboard' })
$ui.NavCompare.Add_Click({ Switch-Tab 'Compare' })
$ui.NavResults.Add_Click({ Switch-Tab 'Results' })
$ui.NavSettings.Add_Click({ Switch-Tab 'Settings' })

# -- Window chrome events ------------------------------------------------------
$ui.TitleBar.Add_MouseLeftButtonDown({
    if ($_.ClickCount -eq 2) {
        if ($Window.WindowState -eq 'Maximized') { $Window.WindowState = 'Normal' }
        else { $Window.WindowState = 'Maximized' }
    } else { $Window.DragMove() }
})

$ui.BtnMinimize.Add_Click({ $Window.WindowState = 'Minimized' })
$ui.BtnMaximize.Add_Click({
    if ($Window.WindowState -eq 'Maximized') { $Window.WindowState = 'Normal' }
    else { $Window.WindowState = 'Maximized' }
})
$ui.BtnClose.Add_Click({ $Window.Close() })

$Window.Add_StateChanged({
    if ($Window.WindowState -eq 'Maximized') {
        $ui.MaximizeIcon.Text = [char]0xE923
        $ui.WindowBorder.CornerRadius = [System.Windows.CornerRadius]::new(0)
        $ui.WindowBorder.BorderThickness = [System.Windows.Thickness]::new(0)
        $ui.WindowBorder.Margin = [System.Windows.Thickness]::new(7)
    } else {
        $ui.MaximizeIcon.Text = [char]0xE922
        $ui.WindowBorder.CornerRadius = [System.Windows.CornerRadius]::new(8)
        $ui.WindowBorder.BorderThickness = [System.Windows.Thickness]::new(1)
        $ui.WindowBorder.Margin = [System.Windows.Thickness]::new(0)
    }
})

# -- Settings events -----------------------------------------------------------
# Title bar theme toggle button
if ($ui.BtnThemeToggle) {
    $ui.BtnThemeToggle.Add_Click({
        # Toggle the settings checkbox (which fires Checked/Unchecked events)
        $ui.ThemeToggle.IsChecked = -not $ui.ThemeToggle.IsChecked
    })
}
# Title bar help button
if ($ui.BtnHelp) {
    $ui.BtnHelp.Add_Click({
        $helpText = @"
ADMX Policy Comparer - Keyboard Shortcuts

Ctrl+`  Toggle debug console
Ctrl+B  Toggle sidebar
Ctrl+E  Export HTML report
Ctrl+F  Focus search box
F5      Run comparison

1-4     Switch tabs (when not in text field)
"@
        Show-Toast 'Keyboard Shortcuts' $helpText
        Write-DebugLog 'Help shown' -Level INFO
    })
}

    <#
    .SYNOPSIS
        Wraps a theme change in a brief opacity fade-out/fade-in transition.
    #>
function Invoke-ThemeTransition([scriptblock]$applyTheme) {
    if ($Script:AnimationsDisabled) { & $applyTheme; return }
    $fadeOut = [System.Windows.Media.Animation.DoubleAnimation]::new(1, 0.85,
        [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(100)))
    $fadeOut.EasingFunction = [System.Windows.Media.Animation.QuadraticEase]::new()
    $WindowRef = $Window; $applyRef = $applyTheme
    $fadeOut.Add_Completed({
        & $applyRef
        $fadeIn = [System.Windows.Media.Animation.DoubleAnimation]::new(0.85, 1,
            [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(150)))
        $fadeIn.EasingFunction = [System.Windows.Media.Animation.QuadraticEase]::new()
        $ui.WindowBorder.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $fadeIn)
    }.GetNewClosure())
    $ui.WindowBorder.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $fadeOut)
}

$ui.ThemeToggle.Add_Unchecked({
    Invoke-ThemeTransition {
        Set-Theme $Script:ThemeLight
        $Script:LogLevelColors = $Script:LogLevelColorsLight
        Rebuild-LogPanel
        Render-AchievementBadges
        Unlock-Achievement 'theme_switcher'
        if ($ui.BtnThemeToggle) { $ui.BtnThemeToggle.Content = [char]0xE708 }
    }
    Write-DebugLog 'Switched to light theme' -Level INFO
})
$ui.ThemeToggle.Add_Checked({
    Invoke-ThemeTransition {
        Set-Theme $Script:ThemeDark
        $Script:LogLevelColors = $Script:LogLevelColorsDark
        Rebuild-LogPanel
        Render-AchievementBadges
        if ($ui.BtnThemeToggle) { $ui.BtnThemeToggle.Content = [char]0xE706 }
    }
    Write-DebugLog 'Switched to dark theme' -Level INFO
})
$ui.AnimationsToggle.Add_Unchecked({ $Script:AnimationsDisabled = $true; Write-DebugLog 'Animations disabled' -Level INFO })
$ui.AnimationsToggle.Add_Checked({  $Script:AnimationsDisabled = $false; Write-DebugLog 'Animations enabled' -Level INFO })
if ($ui.DebugOverlayToggle) {
    $ui.DebugOverlayToggle.Add_Checked({
        $Script:DebugOverlayEnabled = $true
        Rebuild-LogPanel
        if (-not $Script:BottomPanelVisible) { Show-ConsolePanel }
        Write-DebugLog 'Debug overlay enabled - verbose logging active' -Level INFO
    }.GetNewClosure())
    $ui.DebugOverlayToggle.Add_Unchecked({
        $Script:DebugOverlayEnabled = $false
        Rebuild-LogPanel
        Write-DebugLog 'Debug overlay disabled' -Level INFO
    }.GetNewClosure())
}

# #2: Modern folder picker — tries Windows API Code Pack first, then WinForms fallback
    <#
    .SYNOPSIS
        Opens a modern folder picker dialog.
    .DESCRIPTION
        Tries the Windows API Code Pack Vista+ dialog first, falls back to WinForms
        FolderBrowserDialog. Returns the selected path or $null if cancelled.
    .OUTPUTS
        [string] Selected folder path, or $null.
    #>
function Browse-Folder([string]$current) {
    # Try Windows API Code Pack modern Vista+ dialog
    try {
        $asm = [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.WindowsAPICodePack.Shell')
        if ($asm) {
            $dlg = New-Object Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog
            $dlg.IsFolderPicker = $true
            $dlg.Title = 'Select Folder'
            if ($current -and (Test-Path $current)) { $dlg.InitialDirectory = $current }
            if ($dlg.ShowDialog() -eq 'Ok') { return $dlg.FileName }
            return $null
        }
    } catch {}

    # Fallback: WinForms FolderBrowserDialog with UseDescriptionForTitle (modern on .NET 4.8+)
    try { Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue } catch {}
    $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dlg.SelectedPath = $current
    $dlg.Description  = 'Select folder'
    $dlg.UseDescriptionForTitle = $true
    if ($dlg.ShowDialog() -eq 'OK') { return $dlg.SelectedPath }
    return $null
}

$ui.BtnBrowseBasePath.Add_Click({
    $path = Browse-Folder $ui.BasePathText.Text
    if ($path) {
        $ui.BasePathText.Text    = $path
        if ($ui.BasePathSetting) { $ui.BasePathSetting.Text = $path }
        Write-DebugLog "Base path changed: $path" -Level INFO
        Unlock-Achievement 'explorer'
    }
})
if ($ui.BtnBrowseBasePathSetting) {
    $ui.BtnBrowseBasePathSetting.Add_Click({
        $path = Browse-Folder $ui.BasePathSetting.Text
        if ($path) {
            $ui.BasePathSetting.Text = $path
            $ui.BasePathText.Text    = $path
            Write-DebugLog "Base path changed (settings): $path" -Level INFO
            Unlock-Achievement 'explorer'
        }
    })
}
if ($ui.BtnBrowseExportPath) {
    $ui.BtnBrowseExportPath.Add_Click({
        $path = Browse-Folder $ui.ExportPathSetting.Text
        if ($path) {
            $ui.ExportPathSetting.Text = $path
            Write-DebugLog "Export path changed: $path" -Level INFO
        }
    })
}

# -- Custom Modal Dialog (replaces Win95 MessageBox) --------------------------
    <#
    .SYNOPSIS
        Displays a themed modal confirmation dialog with custom button labels.
    .DESCRIPTION
        Creates a fullscreen overlay with a centred card containing icon, title, message,
        and two action buttons (cancel + danger-styled confirm). Blocks the UI dispatcher
        until a choice is made.
    .OUTPUTS
        [bool] $true if confirmed, $false if cancelled.
    #>
function Show-CustomDialog([string]$title, [string]$message, [string]$confirmText = 'Yes', [string]$cancelText = 'No') {
    $overlay = [System.Windows.Controls.Border]::new()
    $overlay.Background = [System.Windows.Media.SolidColorBrush]::new(
        [System.Windows.Media.Color]::FromArgb(160, 0, 0, 0))
    $overlay.SetValue([System.Windows.Controls.Panel]::ZIndexProperty, 1000)

    $card = [System.Windows.Controls.Border]::new()
    $card.Background      = $Window.Resources['ThemeCardBg']
    $card.CornerRadius     = [System.Windows.CornerRadius]::new(16)
    $card.Padding          = [System.Windows.Thickness]::new(32, 28, 32, 24)
    $card.MinWidth         = 380; $card.MaxWidth = 480
    $card.HorizontalAlignment = 'Center'; $card.VerticalAlignment = 'Center'
    $card.BorderBrush      = $Window.Resources['ThemeBorderElevated']
    $card.BorderThickness  = [System.Windows.Thickness]::new(1)
    $card.Effect = [System.Windows.Media.Effects.DropShadowEffect]@{
        Color = [System.Windows.Media.Colors]::Black; BlurRadius = 28
        ShadowDepth = 0; Opacity = 0.5 }

    $stack = [System.Windows.Controls.StackPanel]::new()

    # Icon
    $icon = [System.Windows.Controls.TextBlock]::new()
    $icon.Text = [char]0xE7BA; $icon.FontFamily = 'Segoe MDL2 Assets'
    $icon.FontSize = 32; $icon.Foreground = $Window.Resources['ThemeWarning']
    $icon.Margin = [System.Windows.Thickness]::new(0,0,0,16)
    $stack.Children.Add($icon)

    # Title
    $tb = [System.Windows.Controls.TextBlock]::new()
    $tb.Text = $title; $tb.FontSize = 16; $tb.FontWeight = 'SemiBold'
    $tb.Foreground = $Window.Resources['ThemeTextPrimary']
    $tb.Margin = [System.Windows.Thickness]::new(0,0,0,8)
    $stack.Children.Add($tb)

    # Message
    $msg = [System.Windows.Controls.TextBlock]::new()
    $msg.Text = $message; $msg.FontSize = 13; $msg.TextWrapping = 'Wrap'
    $msg.Foreground = $Window.Resources['ThemeTextSecondary']
    $msg.Margin = [System.Windows.Thickness]::new(0,0,0,24)
    $stack.Children.Add($msg)

    # Button row
    $btnRow = [System.Windows.Controls.StackPanel]::new()
    $btnRow.Orientation = 'Horizontal'; $btnRow.HorizontalAlignment = 'Right'

    $Script:DialogResult = $false

    # Cancel button
    $btnCancel = [System.Windows.Controls.Button]::new()
    $btnCancel.Content = $cancelText; $btnCancel.FontSize = 13
    $btnCancel.Padding = [System.Windows.Thickness]::new(20,9,20,9)
    $btnCancel.Margin  = [System.Windows.Thickness]::new(0,0,10,0)
    $btnCancel.Cursor  = [System.Windows.Input.Cursors]::Hand
    $btnCancel.Style   = $Window.Resources['GhostButton']
    $cancelOverlay = $overlay
    $btnCancel.Add_Click({
        $Script:DialogResult = $false
        $cancelOverlay.Visibility = 'Collapsed'
    }.GetNewClosure())
    $btnRow.Children.Add($btnCancel)

    # Confirm button (danger style)
    $btnOk = [System.Windows.Controls.Button]::new()
    $btnOk.Content = $confirmText; $btnOk.FontSize = 13
    $btnOk.Foreground = [System.Windows.Media.Brushes]::White
    $btnOk.Padding = [System.Windows.Thickness]::new(20,9,20,9)
    $btnOk.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnOk.Style  = $Window.Resources['DangerGhostButton']
    $okBorder = [System.Windows.Controls.Border]::new()
    $okBorder.Background = [System.Windows.Media.SolidColorBrush]::new(
        [System.Windows.Media.ColorConverter]::ConvertFromString('#FF5000'))
    $okBorder.CornerRadius = [System.Windows.CornerRadius]::new(6)
    $okBorder.Padding = [System.Windows.Thickness]::new(20,9,20,9)
    $okBorder.Cursor = [System.Windows.Input.Cursors]::Hand
    $okTb = [System.Windows.Controls.TextBlock]::new()
    $okTb.Text = $confirmText; $okTb.FontSize = 13; $okTb.FontWeight = 'SemiBold'
    $okTb.Foreground = [System.Windows.Media.Brushes]::White
    $okBorder.Child = $okTb
    $okOverlay = $overlay
    $okBorder.Add_MouseLeftButtonDown({
        $Script:DialogResult = $true
        $okOverlay.Visibility = 'Collapsed'
    }.GetNewClosure())
    $okBorder.Add_MouseEnter({ $this.Background = [System.Windows.Media.SolidColorBrush]::new(
        [System.Windows.Media.ColorConverter]::ConvertFromString('#CC4000')) }.GetNewClosure())
    $okBorder.Add_MouseLeave({ $this.Background = [System.Windows.Media.SolidColorBrush]::new(
        [System.Windows.Media.ColorConverter]::ConvertFromString('#FF5000')) }.GetNewClosure())
    $btnRow.Children.Add($okBorder)

    $stack.Children.Add($btnRow)
    $card.Child = $stack
    $overlay.Child = $card

    # Add to window root
    $rootGrid = $ui.WindowBorder.Child
    $rootGrid.Children.Add($overlay)

    # Fade in
    if (-not $Script:AnimationsDisabled) {
        $overlay.Opacity = 0
        $fadeIn = [System.Windows.Media.Animation.DoubleAnimation]::new(0, 1,
            [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(150)))
        $overlay.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $fadeIn)
    }

    # Block with dispatcher until overlay is collapsed
    $frame = [System.Windows.Threading.DispatcherFrame]::new()
    $overlayRef = $overlay; $frameRef = $frame
    $overlay.Add_IsVisibleChanged({
        if ($overlayRef.Visibility -eq 'Collapsed') {
            $frameRef.Continue = $false
        }
    }.GetNewClosure())
    [System.Windows.Threading.Dispatcher]::PushFrame($frame)

    $rootGrid.Children.Remove($overlay)
    return $Script:DialogResult
}
$ui.BtnResetAll.Add_Click({
    $confirmed = Show-CustomDialog -title 'Reset All Data' `
        -message 'This will permanently erase all achievements, settings, and comparison history. This action cannot be undone.' `
        -confirmText 'Reset Everything' -cancelText 'Cancel'
    if ($confirmed) {
        if (Test-Path $Script:PrefsPath) { Remove-Item $Script:PrefsPath -Force }
        if (Test-Path $Script:AchievePath) { Remove-Item $Script:AchievePath -Force }
        if (Test-Path $Script:HistoryPath) { Remove-Item $Script:HistoryPath -Force }
        if (Test-Path $Script:ResultsDir) { Remove-Item $Script:ResultsDir -Recurse -Force }
        if (Test-Path $Script:StreakPath) { Remove-Item $Script:StreakPath -Force }
        $Script:Unlocked = @{}
        $Script:History = @()
        $Script:Streak = @{ LastDate = ''; Count = 0; TotalComparisons = 0 }
        $Script:Prefs = @{
            IsLightMode=$false; DisableAnimations=$false
            BasePath='C:\Program Files (x86)\Microsoft Group Policy'
            ExportPath=[System.IO.Path]::Combine([Environment]::GetFolderPath('Desktop'), 'ADMX_Reports')
            WindowState='Normal'; WindowLeft=-1; WindowTop=-1; WindowWidth=1440; WindowHeight=920
        }
        Load-Preferences
        Render-AchievementBadges
        Update-Dashboard
        Write-DebugLog 'All data reset to defaults' -Level WARN
        Show-Toast 'Data Reset' 'All settings, achievements, and history cleared.' 'warning'
    }
})
# -- Version scanning ----------------------------------------------------------
$Script:AvailableVersions = @()
    <#
    .SYNOPSIS
        Scans the base policy folder for available ADMX/ADML template versions.
    .DESCRIPTION
        Enumerates subdirectories containing PolicyDefinitions\en-US, sorts them by
        Windows version code (e.g. 25H2), and populates the older/newer combo boxes.
    #>
function Scan-Versions {
    $basePath = $ui.BasePathText.Text
    $Script:AvailableVersions = @()
    $ui.OlderVersionCombo.Items.Clear()
    $ui.NewerVersionCombo.Items.Clear()
    if ($ui.AvailableVersionsList) { $ui.AvailableVersionsList.Items.Clear() }

    Write-DebugLog "Scanning versions in: $basePath" -Level INFO

    if (-not (Test-Path $basePath)) {
        if ($ui.VersionCountText) { $ui.VersionCountText.Text = 'Base path not found' }
        Set-LogText "Error: '$basePath' does not exist."
        Write-DebugLog "Base path not found: $basePath" -Level ERROR
        Set-Status 'Path not found' '#FF5000'
        return
    }
    # Version-aware sort: extract build codes like (1709), (20H2), (25H2) and sort chronologically
    $versionOrder = @{
        '1507'=1; '1511'=2; '1607'=3; '1703'=4; '1709'=5; '1803'=6; '1809'=7; '1903'=8; '1909'=9;
        '2004'=10; '20H2'=11; '21H1'=12; '21H2'=13; '22H2'=14; '23H2'=15; '24H2'=16; '25H2'=17; '26H2'=18
    }
    $releases = Get-ChildItem -Path $basePath -Directory -ErrorAction SilentlyContinue |
        Where-Object { Test-Path (Join-Path $_.FullName 'PolicyDefinitions\en-US') } |
        Sort-Object {
            $name = $_.Name
            # Extract version code from parentheses, e.g. "Windows 11 Sep 2025 Update (25H2)" -> "25H2"
            $code = if ($name -match '\((\d{4}|[\d]{2}H\d)\)') { $matches[1] } else { '' }
            # Extract major Windows version: 10 or 11
            $major = if ($name -match 'Windows\s+(\d+)') { [int]$matches[1] } else { 0 }
            # Sort key: major version * 100 + known order, or parse numeric code
            $order = if ($versionOrder.ContainsKey($code)) { $versionOrder[$code] }
                     elseif ($code -match '^\d{4}$') { [int]$code }
                     elseif ($code -match '^(\d{2})H(\d)$') { [int]$matches[1] * 10 + [int]$matches[2] + 200 }
                     else { 999 }
            '{0:D4}-{1:D6}' -f $major, $order
        }

    $Script:AvailableVersions = $releases
    if ($ui.VersionCountText) { $ui.VersionCountText.Text = "$($releases.Count) versions found" }
    Write-DebugLog "Found $($releases.Count) versions with PolicyDefinitions\en-US" -Level SUCCESS
    Invoke-VersionBadgeBlink

    foreach ($r in $releases) {
        $item = [System.Windows.Controls.ComboBoxItem]::new()
        $item.Content = $r.Name
        $item.Tag     = $r.FullName
        $ui.OlderVersionCombo.Items.Add($item)

        $item2 = [System.Windows.Controls.ComboBoxItem]::new()
        $item2.Content = $r.Name
        $item2.Tag     = $r.FullName
        $ui.NewerVersionCombo.Items.Add($item2)

        if ($ui.AvailableVersionsList) {
            $lbItem = [System.Windows.Controls.TextBlock]::new()
            $lbItem.Text = $r.Name; $lbItem.FontSize = 12
            $lbItem.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
            $ui.AvailableVersionsList.Items.Add($lbItem)
        }
        Write-DebugLog "  Version: $($r.Name)" -Level DEBUG
    }
    if ($releases.Count -ge 2) {
        $ui.OlderVersionCombo.SelectedIndex = $releases.Count - 2
        $ui.NewerVersionCombo.SelectedIndex = $releases.Count - 1
    }
    Set-LogText "Found $($releases.Count) versions in '$basePath'."
    Set-Status "Ready - $($releases.Count) versions" '#00C853'
}
$ui.BtnScanVersions.Add_Click({ Scan-Versions })

# -- ADML comparison engine ----------------------------------------------------
    <#
    .SYNOPSIS
        Normalises line endings and trims whitespace for consistent string comparison.
    #>
function Normalize-Text([string]$text) {
    if ([string]::IsNullOrEmpty($text)) { return '' }
    return $text.Replace("`r`n","`n").Replace("`r","`n").Trim()
}

    <#
    .SYNOPSIS
        Parses an ADML file and returns a hashtable of string ID to localised text.
    .OUTPUTS
        [hashtable] Keys are string IDs, values are normalised string content.
    #>
function Get-AdmlStrings([string]$filePath) {
    $table = @{}
    if (-not (Test-Path $filePath)) { return $table }
    try {
        $xml = [xml]([System.IO.File]::ReadAllText($filePath, [System.Text.Encoding]::UTF8))
        $strings = $xml.policyDefinitionResources.resources.stringTable.string
        if ($strings) {
            foreach ($s in $strings) {
                if ($s.id -and $s.id.Trim()) {
                    $table[$s.id.Trim()] = Normalize-Text $s.'#text'
                }
            }
        }
    } catch {
        Write-DebugLog "Error parsing $filePath : $_" -Level ERROR
    }
    return $table
}

    <#
    .SYNOPSIS
        Extracts rich policy metadata from the corresponding ADMX file.
    .DESCRIPTION
        Parses the ADMX XML for each policy element and extracts: parent category,
        registry key/value name, policy class (Machine/User/Both), enabled/disabled
        values, and the full elements tree (enum choices, decimal ranges, text/boolean/
        list/multiText types). Results are keyed by the ADML string IDs referenced in
        displayName and explainText attributes.
    .OUTPUTS
        [hashtable] Keys are ADML string IDs, values are metadata hashtables with
        Policy, Category, RegistryKey, ValueName, Scope, ValueType, EnabledValue,
        DisabledValue, and Elements properties.
    #>
function Get-AdmxCategories([string]$admlDir, [string]$fileName) {
    $categories = @{}
    # ADMX files typically sit one level above the en-US folder
    $admxFile = Join-Path (Split-Path $admlDir) ($fileName -replace '\.adml$', '.admx')
    if (-not (Test-Path $admxFile)) {
        $admxFile = Join-Path $admlDir ($fileName -replace '\.adml$', '.admx')
    }
    if (-not (Test-Path $admxFile)) { return $categories }
    try {
        $xml = [xml]([System.IO.File]::ReadAllText($admxFile, [System.Text.Encoding]::UTF8))
        foreach ($node in $xml.GetElementsByTagName('policy')) {
            $pName = $node.GetAttribute('name')
            $dispRef = $node.GetAttribute('displayName')
            $expRef = $node.GetAttribute('explainText')
            $regKey = $node.GetAttribute('key')
            $valName = $node.GetAttribute('valueName')
            $pClass = $node.GetAttribute('class')
            $parentCat = ''
            $catNode = $node.SelectSingleNode('parentCategory')
            if ($catNode) { $parentCat = $catNode.GetAttribute('ref') }

            # Extract enabled/disabled values
            $enabledVal = ''; $disabledVal = ''
            $enNode = $node.SelectSingleNode('enabledValue/*')
            if ($enNode) { $enabledVal = "$($enNode.LocalName): $($enNode.GetAttribute('value'))" }
            $disNode = $node.SelectSingleNode('disabledValue/*')
            if ($disNode) { $disabledVal = "$($disNode.LocalName): $($disNode.GetAttribute('value'))" }

            # Determine value type from valueName and elements
            $valType = ''
            $elements = @()
            $elemNodes = $node.SelectSingleNode('elements')
            if ($elemNodes) {
                foreach ($el in $elemNodes.ChildNodes) {
                    if ($el.NodeType -ne 'Element') { continue }
                    $elType = $el.LocalName   # enum, decimal, text, boolean, list, multiText
                    $elId = $el.GetAttribute('id')
                    $elDisp = $el.GetAttribute('displayName')
                    $elKey = $el.GetAttribute('key')
                    $elValName = $el.GetAttribute('valueName')

                    $elemInfo = @{ Type=$elType; Id=$elId; DisplayName=$elDisp; Key=$elKey; ValueName=$elValName }

                    switch ($elType) {
                        'enum' {
                            $choices = @()
                            foreach ($item in $el.SelectNodes('item')) {
                                $iDisp = $item.GetAttribute('displayName')
                                $iVal = $item.SelectSingleNode('value/*')
                                $iValStr = if ($iVal) { "$($iVal.LocalName):$($iVal.GetAttribute('value'))" } else { '' }
                                $choices += "$iDisp=$iValStr"
                            }
                            $elemInfo.Choices = $choices
                            $elemInfo.ValueType = 'enum'
                        }
                        'decimal' {
                            $elemInfo.ValueType = 'DWORD'
                            $elemInfo.MinValue = $el.GetAttribute('minValue')
                            $elemInfo.MaxValue = $el.GetAttribute('maxValue')
                        }
                        'text'      { $elemInfo.ValueType = 'REG_SZ' }
                        'multiText' { $elemInfo.ValueType = 'REG_MULTI_SZ' }
                        'boolean'   { $elemInfo.ValueType = 'DWORD (boolean)' }
                        'list'      { $elemInfo.ValueType = 'REG_SZ (list)'; $elemInfo.Additive = $el.GetAttribute('additive') }
                    }
                    $elements += $elemInfo
                }
            }
            # Infer value type from enabled/disabled values if no elements
            if (-not $valType -and $enabledVal -match '^decimal:') { $valType = 'DWORD' }
            elseif (-not $valType -and $enabledVal -match '^string:') { $valType = 'REG_SZ' }
            elseif ($elements.Count -eq 1) { $valType = $elements[0].ValueType }
            elseif ($elements.Count -gt 1) { $valType = 'composite' }
            elseif ($valName -and -not $valType) { $valType = 'DWORD' }

            $info = @{
                Policy=$pName; Category=$parentCat
                RegistryKey=$regKey; ValueName=$valName; Scope=$pClass
                ValueType=$valType; EnabledValue=$enabledVal; DisabledValue=$disabledVal
                Elements=$elements
            }
            if ($dispRef -match '\$\(string\.(.+?)\)') {
                $stringId = $Matches[1]
                $categories[$stringId] = $info
            }
            if ($expRef -match '\$\(string\.(.+?)\)') {
                $stringId = $Matches[1]
                if (-not $categories.ContainsKey($stringId)) {
                    $categories[$stringId] = $info + @{ IsExplain=$true }
                }
            }
            # Map enum/list/text item displayName strings to this policy
            foreach ($elem in $node.SelectNodes('.//item|.//enum|.//list|.//text|.//boolean/*')) {
                $elemDisp = $elem.GetAttribute('displayName')
                if ($elemDisp -match '\$\(string\.(.+?)\)') {
                    $sid = $Matches[1]
                    if (-not $categories.ContainsKey($sid)) {
                        $categories[$sid] = @{
                            Policy=$pName; Category=$parentCat
                            RegistryKey=$regKey; ValueName=''; Scope=$pClass
                        }
                    }
                }
            }
        }
    } catch {
        Write-DebugLog "Error parsing ADMX categories for $fileName : $_" -Level DEBUG
    }
    return $categories
}

    <#
    .SYNOPSIS
        Executes the full ADML/ADMX comparison between the selected older and newer versions.
    .DESCRIPTION
        Compares two PolicyDefinitions\en-US directories file by file. Detects added,
        removed, and modified ADML strings, plus entirely added/removed files. Enriches
        results with ADMX-derived metadata (category, registry path, value type, policy
        group). Filters out trivial whitespace-only changes. Updates the streak, history,
        achievements, dashboard, and navigates to the Results tab on completion.
    #>
function Run-Comparison {
    if ($Script:IsComparing) { return }
    $olderItem = $ui.OlderVersionCombo.SelectedItem
    $newerItem = $ui.NewerVersionCombo.SelectedItem
    if (-not $olderItem -or -not $newerItem) {
        Show-Toast 'Selection Required' 'Please select both older and newer versions.' 'warning'
        return
    }
    if ($olderItem.Content -eq $newerItem.Content) {
        Show-Toast 'Same Version' 'Please select two different versions to compare.' 'warning'
        return
    }

    $Script:IsComparing = $true
    Start-StatusPulse
    $ui.BtnRunComparison.IsEnabled = $false
    $ui.ComparisonStatusText.Text  = 'Comparing...'

    # #27: Apply shimmer gradient to progress bar
    $shimmerBrush = [System.Windows.Media.LinearGradientBrush]::new()
    $shimmerBrush.StartPoint = [System.Windows.Point]::new(0, 0)
    $shimmerBrush.EndPoint = [System.Windows.Point]::new(1, 0)
    $shimmerBrush.GradientStops.Add([System.Windows.Media.GradientStop]::new(
        [System.Windows.Media.ColorConverter]::ConvertFromString('#0078D4'), 0.0))
    $shimmerBrush.GradientStops.Add([System.Windows.Media.GradientStop]::new(
        [System.Windows.Media.ColorConverter]::ConvertFromString('#60CDFF'), 0.5))
    $shimmerBrush.GradientStops.Add([System.Windows.Media.GradientStop]::new(
        [System.Windows.Media.ColorConverter]::ConvertFromString('#0078D4'), 1.0))
    $ui.ComparisonProgressBar.Background = $shimmerBrush
    Set-Status 'Comparing...' '#F59E0B'
    Show-GlobalProgress 'Running comparison...'
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    $sourcePath    = Join-Path $olderItem.Tag 'PolicyDefinitions\en-US'
    $referencePath = Join-Path $newerItem.Tag 'PolicyDefinitions\en-US'
    $olderName     = $olderItem.Content
    $newerName     = $newerItem.Content

    Write-DebugLog "Starting comparison: $olderName -> $newerName" -Level STEP
    Write-DebugLog "Source: $sourcePath" -Level DEBUG
    Write-DebugLog "Reference: $referencePath" -Level DEBUG

    $Script:AllResults.Clear()
    $Script:PolicyMetadata = @{}
    $Script:LoadedOlderVersion = $null
    $Script:LoadedNewerVersion = $null
    Set-LogText "Comparing '$olderName' with '$newerName'...`n"

    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [Action]{ }, [System.Windows.Threading.DispatcherPriority]::Render)

    $addedCount = 0; $removedCount = 0; $modifiedCount = 0
    $fileAddedCount = 0; $fileRemovedCount = 0
    $Script:TrivialSkipCount = 0
    $totalFiles = 0

    try {
        $refFiles = Get-ChildItem -Path $referencePath -Filter '*.adml' -ErrorAction Stop
        $srcFiles = Get-ChildItem -Path $sourcePath    -Filter '*.adml' -ErrorAction Stop
        Write-DebugLog "Reference files: $($refFiles.Count), Source files: $($srcFiles.Count)" -Level INFO
        $fileDiff = Compare-Object -ReferenceObject $refFiles -DifferenceObject $srcFiles -Property Name -IncludeEqual

        $allFiles = $fileDiff | Where-Object { $_.Name }
        $totalIdx = 0; $totalCount = @($allFiles).Count

        # --- Trivial change detection helper ---
        # Returns $true if the difference between two strings is only whitespace or punctuation changes
        # with a Levenshtein distance of $Threshold or less (default 2).
        function Test-TrivialChange {
            param([string]$Old, [string]$New, [int]$Threshold = 2)
            # Quick length check — if length difference exceeds threshold, it's not trivial
            if ([Math]::Abs($Old.Length - $New.Length) -gt $Threshold) { return $false }
            # Strip all whitespace and punctuation, then compare — if alphanumeric content is identical, it's trivial
            $stripOld = $Old -replace '[^\p{L}\p{N}]',''
            $stripNew = $New -replace '[^\p{L}\p{N}]',''
            if ($stripOld -eq $stripNew) { return $true }
            return $false
        }

        # --- Policy group derivation helper ---
        # Returns the parent policy name for a string ID, using ADMX mapping if available
        function Get-PolicyGroup {
            param([string]$StringId, [hashtable]$CatInfo)
            if ($CatInfo -and $CatInfo.Policy) { return $CatInfo.Policy }
            # Derive from StringId by stripping known suffixes
            $id = $StringId -replace '_(Help|Explain|ExplainText|DisplayName|Supported|SupportedOn)$',''
            $id = $id -replace '_\d+$',''
            return $id
        }

        # files added (only in newer)
        $addedFiles = $fileDiff | Where-Object { $_.SideIndicator -eq '<=' }
        foreach ($af in $addedFiles) {
            $totalIdx++; $totalFiles++; $fileAddedCount++
            Update-ComparisonProgress $totalIdx $totalCount "Added file: $($af.Name)"
            Write-DebugLog "File added: $($af.Name)" -Level SUCCESS
            $fp = Join-Path $referencePath $af.Name
            $strings = Get-AdmlStrings $fp
            $afCats = Get-AdmxCategories $referencePath $af.Name
            $Script:AllResults.Add([PSCustomObject]@{
                ChangeType='File Added'; FileName=$af.Name; StringId='-'
                OldValue=''; NewValue="Entire file added ($($strings.Count) strings)"
                Category=''; StringType='File'; PolicyGroup=''
                RegistryKey=''; ValueName=''; Scope=''
            })
            foreach ($key in $strings.Keys | Sort-Object) {
                $addedCount++
                $catInfo = if ($afCats.ContainsKey($key)) { $afCats[$key] }
                           else { @{ Category=''; Policy=''; RegistryKey=''; ValueName=''; Scope=''; ValueType=''; EnabledValue=''; DisabledValue=''; Elements=@() } }
                $cat = if ($catInfo.Category) { $catInfo.Category } else { '' }
                $regKey = if ($catInfo.RegistryKey) { $catInfo.RegistryKey } else { '' }
                $valName = if ($catInfo.ValueName) { $catInfo.ValueName } else { '' }
                $scope = if ($catInfo.Scope) { $catInfo.Scope } else { '' }
                $vType = if ($catInfo.ValueType) { $catInfo.ValueType } else { '' }
                $pGroup = Get-PolicyGroup -StringId $key -CatInfo $catInfo
                if ($catInfo.Policy -and -not $Script:PolicyMetadata.ContainsKey($catInfo.Policy)) {
                    $Script:PolicyMetadata[$catInfo.Policy] = $catInfo
                }
                $sType = if ($key -match '_Help$|_Explain$' -or $catInfo.IsExplain) { 'Explain' }
                         elseif ($key -match '_Supported$') { 'SupportedOn' }
                         else { 'Display' }
                $Script:AllResults.Add([PSCustomObject]@{
                    ChangeType='Added'; FileName=$af.Name; StringId=$key
                    OldValue=''; NewValue=$strings[$key]
                    Category=$cat; StringType=$sType; PolicyGroup=$pGroup
                    RegistryKey=$regKey; ValueName=$valName; Scope=$scope; ValueType=$vType
                })
            }
        }

        # files removed (only in older)
        $removedFiles = $fileDiff | Where-Object { $_.SideIndicator -eq '=>' -and $_.Name }
        foreach ($rf in $removedFiles) {
            $totalIdx++; $totalFiles++; $fileRemovedCount++
            Update-ComparisonProgress $totalIdx $totalCount "Removed file: $($rf.Name)"
            Write-DebugLog "File removed: $($rf.Name)" -Level WARN
            $fp = Join-Path $sourcePath $rf.Name
            $strings = Get-AdmlStrings $fp
            $rfCats = Get-AdmxCategories $sourcePath $rf.Name
            $Script:AllResults.Add([PSCustomObject]@{
                ChangeType='File Removed'; FileName=$rf.Name; StringId='-'
                OldValue="Entire file removed ($($strings.Count) strings)"; NewValue=''
                Category=''; StringType='File'; PolicyGroup=''
                RegistryKey=''; ValueName=''; Scope=''
            })
            foreach ($key in $strings.Keys | Sort-Object) {
                $removedCount++
                $catInfo = if ($rfCats.ContainsKey($key)) { $rfCats[$key] }
                           else { @{ Category=''; Policy=''; RegistryKey=''; ValueName=''; Scope=''; ValueType=''; EnabledValue=''; DisabledValue=''; Elements=@() } }
                $cat = if ($catInfo.Category) { $catInfo.Category } else { '' }
                $regKey = if ($catInfo.RegistryKey) { $catInfo.RegistryKey } else { '' }
                $valName = if ($catInfo.ValueName) { $catInfo.ValueName } else { '' }
                $scope = if ($catInfo.Scope) { $catInfo.Scope } else { '' }
                $vType = if ($catInfo.ValueType) { $catInfo.ValueType } else { '' }
                $pGroup = Get-PolicyGroup -StringId $key -CatInfo $catInfo
                if ($catInfo.Policy -and -not $Script:PolicyMetadata.ContainsKey($catInfo.Policy)) {
                    $Script:PolicyMetadata[$catInfo.Policy] = $catInfo
                }
                $sType = if ($key -match '_Help$|_Explain$' -or $catInfo.IsExplain) { 'Explain' }
                         elseif ($key -match '_Supported$') { 'SupportedOn' }
                         else { 'Display' }
                $Script:AllResults.Add([PSCustomObject]@{
                    ChangeType='Removed'; FileName=$rf.Name; StringId=$key
                    OldValue=$strings[$key]; NewValue=''
                    Category=$cat; StringType=$sType; PolicyGroup=$pGroup
                    RegistryKey=$regKey; ValueName=$valName; Scope=$scope; ValueType=$vType
                })
            }
        }

        # common files
        $commonFiles = $fileDiff | Where-Object { $_.SideIndicator -eq '==' }
        foreach ($cf in $commonFiles) {
            $totalIdx++; $totalFiles++
            Update-ComparisonProgress $totalIdx $totalCount $cf.Name
            Write-DebugLog "Comparing file: $($cf.Name)" -Level DEBUG
            try {
                $srcStrings = Get-AdmlStrings (Join-Path $sourcePath $cf.Name)
                $refStrings = Get-AdmlStrings (Join-Path $referencePath $cf.Name)

                $srcCats = Get-AdmxCategories $sourcePath $cf.Name
                $refCats = Get-AdmxCategories $referencePath $cf.Name
                $allKeys = [System.Collections.Generic.HashSet[string]]::new()
                foreach ($k in $srcStrings.Keys) { [void]$allKeys.Add($k) }
                foreach ($k in $refStrings.Keys) { [void]$allKeys.Add($k) }

                $fileChanges = 0
                foreach ($key in ($allKeys | Sort-Object)) {
                    $inSrc = $srcStrings.ContainsKey($key)
                    $inRef = $refStrings.ContainsKey($key)
                    $catInfo = if ($refCats.ContainsKey($key)) { $refCats[$key] }
                                 elseif ($srcCats.ContainsKey($key)) { $srcCats[$key] }
                                 else { @{ Category=''; Policy=''; RegistryKey=''; ValueName=''; Scope=''; ValueType=''; EnabledValue=''; DisabledValue=''; Elements=@() } }
                    $cat = if ($catInfo.Category) { $catInfo.Category } else { '' }
                    $regKey = if ($catInfo.RegistryKey) { $catInfo.RegistryKey } else { '' }
                    $valName = if ($catInfo.ValueName) { $catInfo.ValueName } else { '' }
                    $scope = if ($catInfo.Scope) { $catInfo.Scope } else { '' }
                    $vType = if ($catInfo.ValueType) { $catInfo.ValueType } else { '' }
                    $pGroup = Get-PolicyGroup -StringId $key -CatInfo $catInfo
                    # Store rich metadata for detail pane lookup
                    if ($catInfo.Policy -and -not $Script:PolicyMetadata.ContainsKey($catInfo.Policy)) {
                        $Script:PolicyMetadata[$catInfo.Policy] = $catInfo
                    }
                    $sType = if ($key -match '_Help$|_Explain$' -or $catInfo.IsExplain) { 'Explain' }
                             elseif ($key -match '_Supported$') { 'SupportedOn' }
                             else { 'Display' }
                    if ($inRef -and -not $inSrc) {
                        $addedCount++; $fileChanges++
                        $Script:AllResults.Add([PSCustomObject]@{
                            ChangeType='Added'; FileName=$cf.Name; StringId=$key
                            OldValue=''; NewValue=$refStrings[$key]
                            Category=$cat; StringType=$sType; PolicyGroup=$pGroup
                            RegistryKey=$regKey; ValueName=$valName; Scope=$scope; ValueType=$vType
                        })
                    } elseif ($inSrc -and -not $inRef) {
                        $removedCount++; $fileChanges++
                        $Script:AllResults.Add([PSCustomObject]@{
                            ChangeType='Removed'; FileName=$cf.Name; StringId=$key
                            OldValue=$srcStrings[$key]; NewValue=''
                            Category=$cat; StringType=$sType; PolicyGroup=$pGroup
                            RegistryKey=$regKey; ValueName=$valName; Scope=$scope; ValueType=$vType
                        })
                    } elseif ($srcStrings[$key] -ne $refStrings[$key]) {
                        if (Test-TrivialChange $srcStrings[$key] $refStrings[$key]) {
                            $Script:TrivialSkipCount++
                            continue
                        }
                        $modifiedCount++; $fileChanges++
                        $Script:AllResults.Add([PSCustomObject]@{
                            ChangeType='Modified'; FileName=$cf.Name; StringId=$key
                            OldValue=$srcStrings[$key]; NewValue=$refStrings[$key]
                            Category=$cat; StringType=$sType; PolicyGroup=$pGroup
                            RegistryKey=$regKey; ValueName=$valName; Scope=$scope; ValueType=$vType
                        })
                    }
                }
                if ($fileChanges -gt 0) {
                    Write-DebugLog "  $($cf.Name): $fileChanges changes" -Level INFO
                }
            } catch {
                Write-DebugLog "Error comparing $($cf.Name): $_" -Level ERROR
                $Script:AllResults.Add([PSCustomObject]@{
                    ChangeType='Modified'; FileName=$cf.Name; StringId='ERROR'
                    OldValue=''; NewValue=$_.Exception.Message
                    Category=''; StringType='Error'; PolicyGroup=''
                    RegistryKey=''; ValueName=''; Scope=''
                })
            }
        }
    } catch {
        Append-LogText "`nERROR: $($_.Exception.Message)"
        Write-DebugLog "Comparison failed: $($_.Exception.Message)" -Level ERROR
    }

    $sw.Stop()
    $totalChanges = $addedCount + $removedCount + $modifiedCount + $fileAddedCount + $fileRemovedCount

    # Update tracking
    Update-Streak
    Add-HistoryEntry $olderName $newerName $totalChanges $totalFiles
    Check-ComparisonAchievements $totalChanges $addedCount $removedCount $modifiedCount $sw.Elapsed

    # Update UI
    Set-AnimatedWidth $ui.ComparisonProgressBar $ui.ComparisonProgressBar.Parent.ActualWidth
    $ui.ComparisonStatusText.Text = "Done in $($sw.Elapsed.TotalSeconds.ToString('F1'))s"
    Hide-GlobalProgress
    $elapsed = $sw.Elapsed.TotalSeconds.ToString('F1')
    $trivialNote = if ($Script:TrivialSkipCount -gt 0) { " | Trivial changes skipped: $($Script:TrivialSkipCount)" } else { '' }
    Append-LogText @"
`nComparison complete in ${elapsed}s.
Files: $totalFiles | Added: $addedCount | Removed: $removedCount | Modified: $modifiedCount
File Added: $fileAddedCount | File Removed: $fileRemovedCount | Total Changes: $totalChanges$trivialNote
"@

    Write-DebugLog "Comparison complete in ${elapsed}s - $totalChanges total changes ($addedCount added, $removedCount removed, $modifiedCount modified, $($Script:TrivialSkipCount) trivial skipped)" -Level SUCCESS
    Set-Status "Done - $totalChanges changes" '#00C853'

    # Update results panel
    $ui.ResultsSubtitle.Text = "$olderName -> $newerName  |  $totalChanges changes across $totalFiles files"
    Apply-Filters
    Populate-FileFilterList
    Update-Dashboard

    $Script:IsComparing = $false
    Stop-StatusPulse
    $ui.BtnRunComparison.IsEnabled = $true

    if ($totalChanges -gt 0) { Switch-Tab 'Results' }
}

    <#
    .SYNOPSIS
        Updates the comparison progress bar, global shimmer, and status text mid-comparison.
    #>
function Update-ComparisonProgress($current, $total, $msg) {
    if ($total -gt 0) {
        $pct = [Math]::Min(($current / $total), 1.0)
        $parentWidth = $ui.ComparisonProgressBar.Parent.ActualWidth
        if ($parentWidth -gt 0) {
            Set-AnimatedWidth $ui.ComparisonProgressBar ($pct * $parentWidth)
        }
        Update-GlobalProgress $pct "[$current/$total] $msg"
    }
    $ui.ComparisonStatusText.Text = "[$current/$total] $msg"
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [Action]{ }, [System.Windows.Threading.DispatcherPriority]::Render)
}

$ui.BtnRunComparison.Add_Click({ Run-Comparison })
$ui.BtnQuickCompare.Add_Click({ Switch-Tab 'Compare' })
$ui.BtnQuickSettings.Add_Click({ Switch-Tab 'Settings' })
# -- Filtering -----------------------------------------------------------------
    <#
    .SYNOPSIS
        Applies the active change-type filter, search text, and policy grouping to the results view.
    .DESCRIPTION
        Configures a CollectionView filter predicate that matches on ChangeType (from filter pills)
        and searches across FileName, StringId, OldValue, NewValue, Category, RegistryKey, and
        ValueName. Optionally groups results by PolicyGroup. Updates the results count summary,
        filter pill counts, breakdown chart, and status bar.
    #>
function Apply-Filters {
    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($Script:AllResults)
    if (-not $view) { return }
    $filterType = $Script:CurrentFilter
    $searchText = $Script:CurrentSearch.ToLower()

    $view.Filter = [Predicate[object]]{
        param($item)
        if ($filterType -ne 'All') {
            if ($item.ChangeType -ne $filterType) { return $false }
        }
        if ($searchText) {
            $match = ($item.FileName -and $item.FileName.ToLower().Contains($searchText)) -or
                     ($item.StringId -and $item.StringId.ToLower().Contains($searchText)) -or
                     ($item.OldValue -and $item.OldValue.ToLower().Contains($searchText)) -or
                     ($item.NewValue -and $item.NewValue.ToLower().Contains($searchText)) -or
                     ($item.Category -and $item.Category.ToLower().Contains($searchText)) -or
                     ($item.RegistryKey -and $item.RegistryKey.ToLower().Contains($searchText)) -or
                     ($item.ValueName -and $item.ValueName.ToLower().Contains($searchText))
            if (-not $match) { return $false }
        }
        return $true
    }

    # Grouping by policy
    $view.GroupDescriptions.Clear()
    if ($Script:GroupByPolicy -and $ui.BtnGroupByPolicy.IsChecked) {
        $view.GroupDescriptions.Add(
            [System.Windows.Data.PropertyGroupDescription]::new('PolicyGroup'))
    }

    $visibleCount = 0
    foreach ($item in $view) { $visibleCount++ }
    # Rich results summary
    $added   = @($Script:AllResults | Where-Object { $_.ChangeType -eq 'Added' }).Count
    $removed = @($Script:AllResults | Where-Object { $_.ChangeType -eq 'Removed' }).Count
    $modified= @($Script:AllResults | Where-Object { $_.ChangeType -eq 'Modified' }).Count
    $fAdded  = @($Script:AllResults | Where-Object { $_.ChangeType -eq 'File Added' }).Count
    $fRemoved= @($Script:AllResults | Where-Object { $_.ChangeType -eq 'File Removed' }).Count
    $summary = "$visibleCount of $($Script:AllResults.Count) results"
    $parts = @()
    if ($added -gt 0)    { $parts += "+$added" }
    if ($removed -gt 0)  { $parts += "-$removed" }
    if ($modified -gt 0) { $parts += "~$modified" }
    if ($fAdded -gt 0)   { $parts += "$fAdded files added" }
    if ($fRemoved -gt 0) { $parts += "$fRemoved files removed" }
    if ($parts.Count -gt 0) { $summary += "  |  $($parts -join '  ·  ')" }
    $ui.ResultsCountText.Text = $summary
    if (Get-Command Update-StatusBar -ErrorAction SilentlyContinue) { Update-StatusBar }
    Update-FilterPillCounts
    Update-BreakdownChart

    # #23: DataGrid fade-in animation on results load
    if (-not $Script:AnimationsDisabled -and $ui.ResultsDataGrid -and -not $Script:DataGridFadeApplied) {
        $Script:DataGridFadeApplied = $true
        $ui.ResultsDataGrid.Opacity = 0
        $gridFade = [System.Windows.Media.Animation.DoubleAnimation]::new(0, 1,
            [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(300)))
        $gridFade.EasingFunction = [System.Windows.Media.Animation.CubicEase]::new()
        $gridFade.EasingFunction.EasingMode = 'EaseOut'
        $ui.ResultsDataGrid.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $gridFade)
    }
    Update-ResultsEmptyState
}

# Filter pill click handlers
$ui.FilterAll.Add_Click({
    $Script:CurrentFilter = 'All'
    $ui.FilterAll.IsChecked = $true
    $ui.FilterAdded.IsChecked = $false; $ui.FilterRemoved.IsChecked = $false
    $ui.FilterModified.IsChecked = $false; $ui.FilterFileAdded.IsChecked = $false
    $ui.FilterFileRemoved.IsChecked = $false
    Apply-Filters
    $Script:FiltersUsed.Add('All') | Out-Null
    if ($Script:FiltersUsed.Count -ge 6) { Unlock-Achievement 'filter_pro' }
    Write-DebugLog 'Filter: All' -Level DEBUG
})
$ui.FilterAdded.Add_Click({
    $Script:CurrentFilter = 'Added'
    $ui.FilterAll.IsChecked = $false; $ui.FilterAdded.IsChecked = $true
    $ui.FilterRemoved.IsChecked = $false; $ui.FilterModified.IsChecked = $false
    $ui.FilterFileAdded.IsChecked = $false; $ui.FilterFileRemoved.IsChecked = $false
    Apply-Filters
    $Script:FiltersUsed.Add('Added') | Out-Null
    if ($Script:FiltersUsed.Count -ge 6) { Unlock-Achievement 'filter_pro' }
    Write-DebugLog 'Filter: Added' -Level DEBUG
})
$ui.FilterRemoved.Add_Click({
    $Script:CurrentFilter = 'Removed'
    $ui.FilterAll.IsChecked = $false; $ui.FilterAdded.IsChecked = $false
    $ui.FilterRemoved.IsChecked = $true; $ui.FilterModified.IsChecked = $false
    $ui.FilterFileAdded.IsChecked = $false; $ui.FilterFileRemoved.IsChecked = $false
    Apply-Filters
    $Script:FiltersUsed.Add('Removed') | Out-Null
    if ($Script:FiltersUsed.Count -ge 6) { Unlock-Achievement 'filter_pro' }
    Write-DebugLog 'Filter: Removed' -Level DEBUG
})
$ui.FilterModified.Add_Click({
    $Script:CurrentFilter = 'Modified'
    $ui.FilterAll.IsChecked = $false; $ui.FilterAdded.IsChecked = $false
    $ui.FilterRemoved.IsChecked = $false; $ui.FilterModified.IsChecked = $true
    $ui.FilterFileAdded.IsChecked = $false; $ui.FilterFileRemoved.IsChecked = $false
    Apply-Filters
    $Script:FiltersUsed.Add('Modified') | Out-Null
    if ($Script:FiltersUsed.Count -ge 6) { Unlock-Achievement 'filter_pro' }
    Write-DebugLog 'Filter: Modified' -Level DEBUG
})
$ui.FilterFileAdded.Add_Click({
    $Script:CurrentFilter = 'File Added'
    $ui.FilterAll.IsChecked = $false; $ui.FilterAdded.IsChecked = $false
    $ui.FilterRemoved.IsChecked = $false; $ui.FilterModified.IsChecked = $false
    $ui.FilterFileAdded.IsChecked = $true; $ui.FilterFileRemoved.IsChecked = $false
    Apply-Filters
    $Script:FiltersUsed.Add('File Added') | Out-Null
    if ($Script:FiltersUsed.Count -ge 6) { Unlock-Achievement 'filter_pro' }
    Write-DebugLog 'Filter: File Added' -Level DEBUG
})
$ui.FilterFileRemoved.Add_Click({
    $Script:CurrentFilter = 'File Removed'
    $ui.FilterAll.IsChecked = $false; $ui.FilterAdded.IsChecked = $false
    $ui.FilterRemoved.IsChecked = $false; $ui.FilterModified.IsChecked = $false
    $ui.FilterFileAdded.IsChecked = $false; $ui.FilterFileRemoved.IsChecked = $true
    Apply-Filters
    $Script:FiltersUsed.Add('File Removed') | Out-Null
    if ($Script:FiltersUsed.Count -ge 6) { Unlock-Achievement 'filter_pro' }
    Write-DebugLog 'Filter: File Removed' -Level DEBUG
})

# Group by Policy toggle
$Script:GroupByPolicy = $false
if ($ui.BtnGroupByPolicy) {
    $ui.BtnGroupByPolicy.Add_Click({
        $Script:GroupByPolicy = $ui.BtnGroupByPolicy.IsChecked
        Apply-Filters
        Write-DebugLog "Group by policy: $($Script:GroupByPolicy)" -Level DEBUG
    })
}

# Search
$ui.SearchBox.Add_TextChanged({
    $Script:CurrentSearch = $ui.SearchBox.Text
    $ui.SearchPlaceholder.Visibility = if ($ui.SearchBox.Text) { 'Collapsed' } else { 'Visible' }
    if ($ui.BtnSearchClear) {
        $ui.BtnSearchClear.Visibility = if ($ui.SearchBox.Text) { 'Visible' } else { 'Collapsed' }
    }
    Apply-Filters
})

# Search clear button (#16)
if ($ui.BtnSearchClear) {
    $ui.BtnSearchClear.Add_Click({
        $ui.SearchBox.Text = ''
        $ui.SearchBox.Focus()
    })
}

# -- File filter list (sidebar) ------------------------------------------------
    <#
    .SYNOPSIS
        Populates the sidebar file filter list with per-file change counts.
    .DESCRIPTION
        Counts results per ADML file, renders each as a labelled list item with a badge count,
        and builds the Results sidebar summary showing added/removed/modified totals.
    #>
function Populate-FileFilterList {
    if (-not $ui.FileFilterList) { return }
    $ui.FileFilterList.Items.Clear()
    $fileCounts = @{}
    foreach ($item in $Script:AllResults) {
        if ($item.FileName) {
            if (-not $fileCounts.ContainsKey($item.FileName)) { $fileCounts[$item.FileName] = 0 }
            $fileCounts[$item.FileName]++
        }
    }
    foreach ($fname in ($fileCounts.Keys | Sort-Object)) {
        $sp = [System.Windows.Controls.StackPanel]::new()
        $sp.Orientation = 'Horizontal'
        $tb = [System.Windows.Controls.TextBlock]::new()
        $tb.Text = $fname; $tb.FontSize = 11.5; $tb.VerticalAlignment = 'Center'
        $tb.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
        $tb.TextTrimming = 'CharacterEllipsis'; $tb.MaxWidth = 170
        $countBd = [System.Windows.Controls.Border]::new()
        $countBd.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeHoverBg')
        $countBd.CornerRadius = [System.Windows.CornerRadius]::new(8)
        $countBd.Padding = [System.Windows.Thickness]::new(6,1,6,1)
        $countBd.Margin  = [System.Windows.Thickness]::new(8,0,0,0)
        $countTb = [System.Windows.Controls.TextBlock]::new()
        $countTb.Text = $fileCounts[$fname].ToString()
        $countTb.FontSize = 10
        $countTb.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextSecondary')
        $countBd.Child = $countTb
        $sp.Children.Add($tb); $sp.Children.Add($countBd)
        $ui.FileFilterList.Items.Add($sp)
    }

    # Summary in sidebar
    if ($ui.ResultsSidebarSummary) {
        $ui.ResultsSidebarSummary.Children.Clear()
        $added   = @($Script:AllResults | Where-Object { $_.ChangeType -eq 'Added' }).Count
        $removed = @($Script:AllResults | Where-Object { $_.ChangeType -eq 'Removed' }).Count
        $modified= @($Script:AllResults | Where-Object { $_.ChangeType -eq 'Modified' }).Count
        foreach ($pair in @(@('Added',$added,'#00C853'),@('Removed',$removed,'#FF5000'),@('Modified',$modified,'#F59E0B'))) {
            $sp = [System.Windows.Controls.StackPanel]::new()
            $sp.Orientation = 'Horizontal'; $sp.Margin = [System.Windows.Thickness]::new(0,2,0,2)
            $dot = [System.Windows.Shapes.Ellipse]::new()
            $dot.Width = 8; $dot.Height = 8; $dot.Margin = [System.Windows.Thickness]::new(0,0,8,0)
            $dot.Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString($pair[2])
            $dot.VerticalAlignment = 'Center'
            $lbl = [System.Windows.Controls.TextBlock]::new()
            $lbl.Text = "$($pair[0]): $($pair[1])"; $lbl.FontSize = 12
            $lbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextSecondary')
            $sp.Children.Add($dot); $sp.Children.Add($lbl)
            $ui.ResultsSidebarSummary.Children.Add($sp)
        }
    }
    Write-DebugLog "File filter list populated: $($fileCounts.Count) files" -Level DEBUG
}

if ($ui.FileFilterList) {
    $ui.FileFilterList.Add_SelectionChanged({
        if ($ui.FileFilterList.SelectedItem) {
            $sp = $ui.FileFilterList.SelectedItem
            if ($sp -is [System.Windows.Controls.StackPanel] -and $sp.Children.Count -gt 0) {
                $fileName = $sp.Children[0].Text
                $ui.SearchBox.Text = $fileName
            }
        }
    })
}
if ($ui.FileFilterAllBtn) {
    $ui.FileFilterAllBtn.Add_MouseLeftButtonDown({
        $ui.SearchBox.Text = ''
        if ($ui.FileFilterList) { $ui.FileFilterList.SelectedIndex = -1 }
    })
}
# -- Dashboard updates ---------------------------------------------------------
    <#
    .SYNOPSIS
        Refreshes all Dashboard tab elements: stat cards, recent comparisons, and achievements.
    .DESCRIPTION
        Reloads history and streak data, animates stat card counters, populates the recent
        comparisons list with click-to-load entries, fills the export history combo box,
        and starts the card breathe animation.
    #>
function Update-Dashboard {
    Load-History
    Load-Streak
    if ($ui.StatComparisons) { Invoke-CountUp $ui.StatComparisons $Script:Streak.TotalComparisons }
    if ($ui.StatChanges) {
        $chgVal = 0; foreach ($h in $Script:History) { $v = $h.ChangeCount; if ($v -is [array]) { $v = $v[0] }; $chgVal += [int]$v }
        Invoke-CountUp $ui.StatChanges $chgVal
    }
    if ($ui.StatFiles) {
        $fileVal = 0; foreach ($h in $Script:History) { $v = $h.FileCount; if ($v -is [array]) { $v = $v[0] }; $fileVal += [int]$v }
        Invoke-CountUp $ui.StatFiles $fileVal
    }
    if ($ui.StatStreak)      { Invoke-CountUp $ui.StatStreak $Script:Streak.Count }

    # Sidebar summary
    if ($ui.DashSidebarSummary) {
        if ($Script:History.Count -gt 0) {
            $last = $Script:History[0]
            $ui.DashSidebarSummary.Text = "Last: $($last.OlderVersion) -> $($last.NewerVersion)`n$($last.ChangeCount) changes on $($last.Date)"
        } else {
            $ui.DashSidebarSummary.Text = 'No comparisons yet'
        }
    }

    # Recent comparisons list
    if ($ui.RecentComparisonsList) {
        $ui.RecentComparisonsList.Items.Clear()
        if ($Script:History.Count -gt 0) {
            if ($ui.NoRecentText) { $ui.NoRecentText.Visibility = 'Collapsed' }
            foreach ($h in ($Script:History | Select-Object -First 10)) {
                $sp = [System.Windows.Controls.Grid]::new()
                $sp.Margin = [System.Windows.Thickness]::new(0,2,0,2)
                $col0 = [System.Windows.Controls.ColumnDefinition]::new()
                $col0.Width = [System.Windows.GridLength]::new(0, 'Auto')
                $col1 = [System.Windows.Controls.ColumnDefinition]::new()
                $col1.Width = [System.Windows.GridLength]::new(1, 'Star')
                $col2 = [System.Windows.Controls.ColumnDefinition]::new(); $col2.Width = 'Auto'
                $col3 = [System.Windows.Controls.ColumnDefinition]::new(); $col3.Width = 'Auto'
                $sp.ColumnDefinitions.Add($col0)
                $sp.ColumnDefinitions.Add($col1); $sp.ColumnDefinitions.Add($col2); $sp.ColumnDefinitions.Add($col3)

                # Icon badge (WinGet MM ConfigListItem pattern)
                $iconBorder = [System.Windows.Controls.Border]::new()
                $iconBorder.Width = 28; $iconBorder.Height = 28
                $iconBorder.CornerRadius = [System.Windows.CornerRadius]::new(6)

                # #36: Hover shadow elevation on recent list cards
                $gridRef = $sp
                $sp.Add_MouseEnter({
                    $gridRef.Effect = [System.Windows.Media.Effects.DropShadowEffect]@{
                        Color = 'Black'; BlurRadius = 8; ShadowDepth = 2; Opacity = 0.25; Direction = 270
                    }
                }.GetNewClosure())
                $sp.Add_MouseLeave({
                    $gridRef.Effect = $null
                }.GetNewClosure())
                $iconBorder.Margin = [System.Windows.Thickness]::new(0,0,10,0)
                $iconBorder.Background = [System.Windows.Media.SolidColorBrush]::new(
                    [System.Windows.Media.ColorConverter]::ConvertFromString('#200078D4'))
                $iconTb = [System.Windows.Controls.TextBlock]::new()
                $iconTb.Text = [char]0xE8AB
                $iconTb.FontFamily = [System.Windows.Media.FontFamily]::new('Segoe MDL2 Assets')
                $iconTb.FontSize = 13
                $iconTb.Foreground = $Window.Resources['ThemeAccent']
                $iconTb.HorizontalAlignment = 'Center'; $iconTb.VerticalAlignment = 'Center'
                $iconBorder.Child = $iconTb
                [System.Windows.Controls.Grid]::SetColumn($iconBorder, 0)
                $sp.Children.Add($iconBorder)

                # Version text stack
                $verStack = [System.Windows.Controls.StackPanel]::new()
                $verStack.VerticalAlignment = 'Center'
                $verTb = [System.Windows.Controls.TextBlock]::new()
                $olderRun = [System.Windows.Documents.Run]::new($h.OlderVersion)
                $olderRun.Foreground = $Window.Resources['ThemeTextSecondary']
                $verTb.Inlines.Add($olderRun)
                $arrow = [System.Windows.Documents.Run]::new('  ➜  ')
                $arrow.Foreground = $Window.Resources['ThemeTextDim']
                $verTb.Inlines.Add($arrow)
                $newer = [System.Windows.Documents.Run]::new($h.NewerVersion)
                $newer.Foreground = $Window.Resources['ThemeAccentLight']
                $newer.FontWeight = 'SemiBold'
                $verTb.Inlines.Add($newer)
                $verTb.FontSize = 12
                $verStack.Children.Add($verTb)
                $subTb = [System.Windows.Controls.TextBlock]::new()
                $subTb.Text = "$($h.ChangeCount) changes"
                $subTb.FontSize = 10; $subTb.Foreground = $Window.Resources['ThemeTextDim']
                $verStack.Children.Add($subTb)
                [System.Windows.Controls.Grid]::SetColumn($verStack, 1)
                $sp.Children.Add($verStack)

                $dateTb = [System.Windows.Controls.TextBlock]::new()
                $dateTb.Text = $h.Date; $dateTb.FontSize = 10
                $dateTb.Foreground = $Window.Resources['ThemeTextDim']
                $dateTb.VerticalAlignment = 'Center'
                [System.Windows.Controls.Grid]::SetColumn($dateTb, 3)
                $sp.Children.Add($dateTb)

                # Click-to-load: clicking a history item loads its stored results
                $sp.Cursor = [System.Windows.Input.Cursors]::Hand
                $entryRef = $h
                $sp.Add_MouseLeftButtonDown({
                    $eid = $entryRef.Id
                    if (-not $eid) {
                        Show-Toast 'No Data' 'This comparison was saved before result storage was added.' 'warning'
                        return
                    }
                    $loaded = Load-ComparisonResults $eid
                    if ($loaded) {
                        $Script:LoadedOlderVersion = $entryRef.OlderVersion
                        $Script:LoadedNewerVersion = $entryRef.NewerVersion
                        $ui.ResultsSubtitle.Text = "$($entryRef.OlderVersion) -> $($entryRef.NewerVersion)  |  $($entryRef.ChangeCount) changes (loaded from history)"
                        Apply-Filters
                        Switch-Tab 'Results'
                        Show-Toast 'Loaded' "Restored $($entryRef.ChangeCount) results from $($entryRef.Date)" 'success'
                    } else {
                        Show-Toast 'Not Found' 'Result data for this comparison is no longer available.' 'warning'
                    }
                }.GetNewClosure())

                $ui.RecentComparisonsList.Items.Add($sp)
            }
        } else {
            if ($ui.NoRecentText) { $ui.NoRecentText.Visibility = 'Visible' }
        }
    }

    # Populate export history combo
    if ($ui.ExportHistoryCombo) {
        $ui.ExportHistoryCombo.Items.Clear()
        foreach ($h in ($Script:History | Select-Object -First 20)) {
            $item = [System.Windows.Controls.ComboBoxItem]::new()
            $item.Content = "$($h.OlderVersion) → $($h.NewerVersion)  ($($h.ChangeCount) changes, $($h.Date))"
            $item.Tag = $h
            $ui.ExportHistoryCombo.Items.Add($item)
        }
        if ($ui.ExportHistoryCombo.Items.Count -gt 0) {
            $ui.ExportHistoryCombo.SelectedIndex = 0
        }
    }

    Start-CardBreathe
    Write-DebugLog 'Dashboard updated' -Level DEBUG
}

# -- Export --------------------------------------------------------------------
    <#
    .SYNOPSIS
        Exports the current (optionally filtered) comparison results as a styled HTML report.
    .DESCRIPTION
        Generates a self-contained HTML file with dark/light theme toggle, collapsible sidebar,
        filter pills, search, and per-file expandable sections. Each row shows change type,
        string ID, category, registry path, and old/new values with inline diff highlighting.
        Respects the currently active filter and search. Unlocks export achievements.
    #>

function Get-HtmlDiffSpans([string]$oldText, [string]$newText) {
    $diff = Get-WordDiffRuns $oldText $newText
    $oldHtml = [System.Text.StringBuilder]::new()
    foreach ($run in $diff.OldRuns) {
        $enc = [System.Web.HttpUtility]::HtmlEncode($run.Text)
        if ($run.Type -eq 'deleted') { [void]$oldHtml.Append("<span class='diff-del'>$enc</span>") }
        else { [void]$oldHtml.Append($enc) }
    }
    $newHtml = [System.Text.StringBuilder]::new()
    foreach ($run in $diff.NewRuns) {
        $enc = [System.Web.HttpUtility]::HtmlEncode($run.Text)
        if ($run.Type -eq 'added') { [void]$newHtml.Append("<span class='diff-add'>$enc</span>") }
        else { [void]$newHtml.Append($enc) }
    }
    return @{ Old=$oldHtml.ToString(); New=$newHtml.ToString() }
}

function Export-Html {
    if ($Script:AllResults.Count -eq 0) {
        Show-Toast 'No Data' 'Run a comparison first before exporting.' 'warning'
        return
    }
    # Use filtered view if active filters exist
    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($Script:AllResults)
    $visibleItems = @($view | ForEach-Object { $_ })
    $exportingFiltered = ($visibleItems.Count -lt $Script:AllResults.Count)
    if ($exportingFiltered) {
        Write-DebugLog "Exporting filtered subset: $($visibleItems.Count) of $($Script:AllResults.Count)" -Level INFO
    }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Filter = 'HTML files (*.html)|*.html'
    $dlg.FileName = "ADMX_Comparison_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    if ($ui.ExportPathSetting -and $ui.ExportPathSetting.Text -and (Test-Path $ui.ExportPathSetting.Text)) {
        $dlg.InitialDirectory = $ui.ExportPathSetting.Text
    }
    if (-not $dlg.ShowDialog()) { return }

    Write-DebugLog "Exporting HTML to $($dlg.FileName)" -Level STEP

    # Use loaded history versions if available, otherwise use combo selection
    $older = if ($Script:LoadedOlderVersion) { $Script:LoadedOlderVersion }
             elseif ($ui.OlderVersionCombo.SelectedItem) { $ui.OlderVersionCombo.SelectedItem.Content }
             else { '?' }
    $newer = if ($Script:LoadedNewerVersion) { $Script:LoadedNewerVersion }
             elseif ($ui.NewerVersionCombo.SelectedItem) { $ui.NewerVersionCombo.SelectedItem.Content }
             else { '?' }
    $olderEnc = [System.Web.HttpUtility]::HtmlEncode($older)
    $newerEnc = [System.Web.HttpUtility]::HtmlEncode($newer)
    $genDate  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    $exportData = if ($exportingFiltered) { $visibleItems } else { $Script:AllResults }
    $added    = @($exportData | Where-Object { $_.ChangeType -eq 'Added' }).Count
    $removed  = @($exportData | Where-Object { $_.ChangeType -eq 'Removed' }).Count
    $modified = @($exportData | Where-Object { $_.ChangeType -eq 'Modified' }).Count
    $fAdded   = @($exportData | Where-Object { $_.ChangeType -eq 'File Added' }).Count
    $fRemoved = @($exportData | Where-Object { $_.ChangeType -eq 'File Removed' }).Count
    $total    = @($exportData).Count

    $grouped = $exportData | Group-Object FileName

    # Build table of contents and file sections
    $tocHtml = [System.Text.StringBuilder]::new()
    $bodyHtml = [System.Text.StringBuilder]::new()
    $fileIdx = 0
    foreach ($g in ($grouped | Sort-Object Name)) {
        $fileIdx++
        $fname = [System.Web.HttpUtility]::HtmlEncode($g.Name)
        $fAdds = @($g.Group | Where-Object { $_.ChangeType -eq 'Added' }).Count
        $fRems = @($g.Group | Where-Object { $_.ChangeType -eq 'Removed' }).Count
        $fMods = @($g.Group | Where-Object { $_.ChangeType -eq 'Modified' }).Count
        $fFA   = @($g.Group | Where-Object { $_.ChangeType -eq 'File Added' }).Count
        $fFR   = @($g.Group | Where-Object { $_.ChangeType -eq 'File Removed' }).Count

        # TOC entry
        $tocBadges = ''
        if ($fAdds -gt 0) { $tocBadges += "<span class='toc-badge toc-added'>+$fAdds</span>" }
        if ($fRems -gt 0) { $tocBadges += "<span class='toc-badge toc-removed'>-$fRems</span>" }
        if ($fMods -gt 0) { $tocBadges += "<span class='toc-badge toc-modified'>~$fMods</span>" }
        if ($fFA -gt 0)   { $tocBadges += "<span class='toc-badge toc-file-added'>new</span>" }
        if ($fFR -gt 0)   { $tocBadges += "<span class='toc-badge toc-file-removed'>del</span>" }
        [void]$tocHtml.Append("<a href='#file-$fileIdx' class='toc-item'><span class='toc-name'>$fname</span>$tocBadges</a>")

        # File section — collapsible
        [void]$bodyHtml.Append("<details class='file-section' id='file-$fileIdx' open>")
        [void]$bodyHtml.Append("<summary class='file-header'><span class='file-name'>$fname</span>")
        [void]$bodyHtml.Append("<span class='file-badges'>")
        if ($fAdds -gt 0) { [void]$bodyHtml.Append("<span class='badge badge-added'>+$fAdds</span>") }
        if ($fRems -gt 0) { [void]$bodyHtml.Append("<span class='badge badge-removed'>-$fRems</span>") }
        if ($fMods -gt 0) { [void]$bodyHtml.Append("<span class='badge badge-modified'>~$fMods</span>") }
        if ($fFA -gt 0)   { [void]$bodyHtml.Append("<span class='badge badge-file-added'>File Added</span>") }
        if ($fFR -gt 0)   { [void]$bodyHtml.Append("<span class='badge badge-file-removed'>File Removed</span>") }
        [void]$bodyHtml.Append("<span class='badge badge-total'>$($g.Count)</span>")
        [void]$bodyHtml.Append("</span></summary>")
        [void]$bodyHtml.Append("<table><thead><tr><th class='col-status'>Status</th><th class='col-cat'>Category</th><th class='col-sid'>String ID</th><th class='col-reg'>Registry Path</th><th class='col-old'>Old Value</th><th class='col-new'>New Value</th></tr></thead><tbody>")
        # Sub-group by PolicyGroup within each file
        $policyGroups = $g.Group | Group-Object PolicyGroup
        foreach ($pg in ($policyGroups | Sort-Object Name)) {
            if ($pg.Name) {
                $pgEnc = [System.Web.HttpUtility]::HtmlEncode($pg.Name)
                [void]$bodyHtml.Append("<tr class='policy-group-row'><td colspan='6'><span class='pg-icon'>&#xE8CB;</span> $pgEnc <span class='pg-count'>$($pg.Count)</span></td></tr>")
            }
        foreach ($row in $pg.Group) {
            $cssClass = switch ($row.ChangeType) {
                'Added'        { 'added' }
                'Removed'      { 'removed' }
                'Modified'     { 'modified' }
                'File Added'   { 'file-added' }
                'File Removed' { 'file-removed' }
                default        { '' }
            }
            $statusIcon = switch ($row.ChangeType) {
                'Added'        { '&#x2795;' }
                'Removed'      { '&#x2796;' }
                'Modified'     { '&#x270F;' }
                'File Added'   { '&#x1F4C4;' }
                'File Removed' { '&#x1F5D1;' }
                default        { '' }
            }
            [void]$bodyHtml.Append("<tr class='data-row $cssClass' data-type='$($row.ChangeType)'>")
            [void]$bodyHtml.Append("<td class='status-cell'>$statusIcon $([System.Web.HttpUtility]::HtmlEncode($row.ChangeType))</td>")
            $catEnc = [System.Web.HttpUtility]::HtmlEncode($row.Category)
            [void]$bodyHtml.Append("<td title='$catEnc'>$catEnc</td>")
            [void]$bodyHtml.Append("<td><code title='$([System.Web.HttpUtility]::HtmlEncode($row.StringId))'>$([System.Web.HttpUtility]::HtmlEncode($row.StringId))</code></td>")
            $regPath = if ($row.RegistryKey) {
                $vn = if ($row.ValueName) { "\$($row.ValueName)" } else { '' }
                $full = [System.Web.HttpUtility]::HtmlEncode("$($row.RegistryKey)$vn")
                $vtBadge = if ($row.ValueType) { " <span class='badge badge-total'>$([System.Web.HttpUtility]::HtmlEncode($row.ValueType))</span>" } else { '' }
                $scBadge = if ($row.Scope -and $row.Scope -ne 'Both') { " <span class='badge badge-modified'>$([System.Web.HttpUtility]::HtmlEncode($row.Scope))</span>" } else { '' }
                "<code title='$full'>$full</code>$vtBadge$scBadge"
            } else { '' }
            [void]$bodyHtml.Append("<td>$regPath</td>")
            $oldEnc = [System.Web.HttpUtility]::HtmlEncode($row.OldValue)
            $newEnc = [System.Web.HttpUtility]::HtmlEncode($row.NewValue)
            if ($row.ChangeType -eq 'Modified' -and $row.OldValue -and $row.NewValue) {
                $diffHtml = Get-HtmlDiffSpans $row.OldValue $row.NewValue
                [void]$bodyHtml.Append("<td title='$oldEnc'><pre>$($diffHtml.Old)</pre></td>")
                [void]$bodyHtml.Append("<td title='$newEnc'><pre>$($diffHtml.New)</pre></td>")
            } else {
                [void]$bodyHtml.Append("<td title='$oldEnc'><pre>$oldEnc</pre></td>")
                [void]$bodyHtml.Append("<td title='$newEnc'><pre>$newEnc</pre></td>")
            }
            [void]$bodyHtml.Append("</tr>")
        }
        } # end policy group loop
        [void]$bodyHtml.Append("</tbody></table></details>")
    }

    $html = @"
<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ADMX Comparison — $olderEnc vs $newerEnc</title>
<style>
:root {
  --bg: #111113; --bg-card: #1A1A1D; --bg-surface: #18181B; --bg-hover: #27272B;
  --text: #E4E4E7; --text-secondary: #A1A1AA; --text-dim: #71717A;
  --border: #27272B; --border-card: #333;
  --accent: #60CDFF; --accent-light: #90DEFF;
  --green: #00C853; --green-bg: #0D2818; --green-light: #00E676; --green-bg-light: #0A3D1A;
  --red: #FF5000; --red-bg: #2D0F0F; --red-light: #FF6E40; --red-bg-light: #3D1010;
  --yellow: #F59E0B; --yellow-bg: #2D2200;
  --purple: #B388FF;
}
.light {
  --bg: #FAFAFA; --bg-card: #FFFFFF; --bg-surface: #F4F4F5; --bg-hover: #E4E4E7;
  --text: #18181B; --text-secondary: #52525B; --text-dim: #A1A1AA;
  --border: #E4E4E7; --border-card: #D4D4D8;
  --accent: #0078D4; --accent-light: #005A9E;
  --green: #16A34A; --green-bg: #DCFCE7; --green-light: #15803D; --green-bg-light: #BBF7D0;
  --red: #DC2626; --red-bg: #FEE2E2; --red-light: #B91C1C; --red-bg-light: #FECACA;
  --yellow: #D97706; --yellow-bg: #FEF3C7;
  --purple: #7C3AED;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Segoe UI Variable Display','Segoe UI',system-ui,sans-serif; background: var(--bg); color: var(--text); font-size: 14px; line-height: 1.5; }

/* Header */
.header { position: sticky; top: 0; z-index: 100; background: var(--bg); border-bottom: 1px solid var(--border); padding: 16px 28px; backdrop-filter: blur(12px); }
.header-inner { max-width: 1400px; margin: 0 auto; display: flex; align-items: center; justify-content: space-between; gap: 16px; flex-wrap: wrap; }
.header h1 { font-size: 20px; font-weight: 700; color: var(--text); display: flex; align-items: center; gap: 10px; }
.header h1 .icon { font-size: 24px; }
.header-meta { font-size: 12px; color: var(--text-dim); }
.header-actions { display: flex; gap: 8px; align-items: center; }
.header-actions button { background: var(--bg-card); border: 1px solid var(--border-card); color: var(--text-secondary); padding: 6px 14px; border-radius: 8px; cursor: pointer; font-size: 12px; font-family: inherit; transition: all .15s; }
.header-actions button:hover { background: var(--bg-hover); color: var(--text); }
.header-actions button.active { background: var(--accent); color: #fff; border-color: var(--accent); }

/* Layout */
.container { max-width: 1400px; margin: 0 auto; padding: 20px 28px; display: grid; grid-template-columns: 240px 1fr; gap: 20px; transition: grid-template-columns .25s ease; }
.container.sidebar-collapsed { grid-template-columns: 0 1fr; gap: 0; }
.container.sidebar-collapsed .sidebar { width: 0; overflow: hidden; opacity: 0; padding: 0; pointer-events: none; }
@media (max-width: 900px) { .container { grid-template-columns: 1fr; } .sidebar { display: none; } }
.sidebar { transition: opacity .2s ease, width .25s ease; }

/* Summary Cards */
.summary-bar { display: grid; grid-template-columns: repeat(auto-fit, minmax(130px, 1fr)); gap: 12px; margin-bottom: 20px; grid-column: 1 / -1; }
.stat-card { background: var(--bg-card); border: 1px solid var(--border-card); border-radius: 12px; padding: 16px 18px; text-align: center; position: relative; overflow: hidden; }
.stat-card::before { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px; border-radius: 12px 12px 0 0; }
.stat-card.sc-added::before { background: linear-gradient(90deg, var(--green), var(--green-light)); }
.stat-card.sc-removed::before { background: linear-gradient(90deg, var(--red), var(--red-light)); }
.stat-card.sc-modified::before { background: linear-gradient(90deg, var(--yellow), #FBBF24); }
.stat-card.sc-fa::before { background: linear-gradient(90deg, #16A34A, #4ADE80); }
.stat-card.sc-fr::before { background: linear-gradient(90deg, #B91C1C, #F87171); }
.stat-card.sc-total::before { background: linear-gradient(90deg, var(--accent), var(--accent-light)); }
.stat-card .num { font-size: 28px; font-weight: 700; }
.stat-card .label { font-size: 11px; color: var(--text-dim); margin-top: 2px; }
.stat-card.sc-added .num { color: var(--green); } .stat-card.sc-removed .num { color: var(--red); }
.stat-card.sc-modified .num { color: var(--yellow); } .stat-card.sc-total .num { color: var(--accent); }
.stat-card.sc-fa .num { color: var(--green); } .stat-card.sc-fr .num { color: var(--red); }

/* Version info */
.version-bar { grid-column: 1 / -1; background: var(--bg-card); border: 1px solid var(--border-card); border-radius: 12px; padding: 16px 20px; display: flex; align-items: center; gap: 20px; flex-wrap: wrap; margin-bottom: 8px; }
.version-bar strong { color: var(--text-dim); font-size: 11px; text-transform: uppercase; letter-spacing: .5px; }
.version-bar .ver { color: var(--text); font-weight: 600; }
.version-bar .arrow { color: var(--accent); font-size: 18px; }
.version-bar .date { margin-left: auto; color: var(--text-dim); font-size: 12px; }

/* Sidebar / TOC */
.sidebar { position: sticky; top: 80px; max-height: calc(100vh - 100px); overflow-y: auto; }
.toc { background: var(--bg-card); border: 1px solid var(--border-card); border-radius: 12px; padding: 14px; }
.toc h3 { font-size: 12px; color: var(--text-dim); text-transform: uppercase; letter-spacing: .5px; margin-bottom: 10px; padding: 0 6px; }
.toc-item { display: flex; align-items: center; gap: 6px; padding: 6px 8px; border-radius: 6px; text-decoration: none; color: var(--text-secondary); font-size: 12px; transition: all .15s; flex-wrap: wrap; }
.toc-item:hover { background: var(--bg-hover); color: var(--text); }
.toc-name { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.toc-badge { font-size: 9px; padding: 1px 5px; border-radius: 8px; font-weight: 600; }
.toc-added { background: var(--green-bg); color: var(--green); }
.toc-removed { background: var(--red-bg); color: var(--red); }
.toc-modified { background: var(--yellow-bg); color: var(--yellow); }
.toc-file-added { background: var(--green-bg-light); color: var(--green-light); }
.toc-file-removed { background: var(--red-bg-light); color: var(--red-light); }

/* Search */
.search-container { grid-column: 1 / -1; margin-bottom: 4px; }
.search-box { width: 100%; padding: 10px 16px 10px 40px; background: var(--bg-card); border: 1px solid var(--border-card); border-radius: 10px; color: var(--text); font-size: 14px; font-family: inherit; outline: none; transition: border-color .15s; }
.search-box:focus { border-color: var(--accent); }
.search-wrap { position: relative; }
.search-wrap::before { content: '\1F50D'; position: absolute; left: 14px; top: 50%; transform: translateY(-50%); font-size: 14px; opacity: .5; }
.search-count { position: absolute; right: 14px; top: 50%; transform: translateY(-50%); font-size: 11px; color: var(--text-dim); }
.no-match { display: none; }

/* Filter pills */
.filter-bar { grid-column: 1 / -1; display: flex; gap: 6px; flex-wrap: wrap; margin-bottom: 8px; }
.filter-pill { padding: 5px 12px; border-radius: 20px; border: 1px solid var(--border-card); background: var(--bg-card); color: var(--text-secondary); font-size: 12px; cursor: pointer; transition: all .15s; font-family: inherit; }
.filter-pill:hover { background: var(--bg-hover); }
.filter-pill.active { background: var(--accent); color: #fff; border-color: var(--accent); }

/* File sections */
.file-section { background: var(--bg-card); border: 1px solid var(--border-card); border-radius: 12px; margin-bottom: 16px; overflow: hidden; }
.file-header { padding: 14px 18px; cursor: pointer; display: flex; align-items: center; gap: 10px; font-size: 14px; font-weight: 600; color: var(--accent); list-style: none; user-select: none; }
.file-header::-webkit-details-marker { display: none; }
.file-header::before { content: '\25BC'; font-size: 10px; color: var(--text-dim); transition: transform .2s; }
details:not([open]) .file-header::before { transform: rotate(-90deg); }
.file-name { flex: 1; }
.file-badges { display: flex; gap: 4px; flex-wrap: wrap; }
table { border-collapse: collapse; width: 100%; table-layout: fixed; }
thead { position: sticky; top: 60px; z-index: 10; }
th { background: var(--bg-surface); color: var(--text-dim); font-size: 11px; font-weight: 600; padding: 8px 12px; text-align: left; border-bottom: 1px solid var(--border); text-transform: uppercase; letter-spacing: .3px; overflow: hidden; text-overflow: ellipsis; }
th.col-status { width: 90px; } th.col-cat { width: 12%; } th.col-sid { width: 18%; } th.col-reg { width: 22%; } th.col-old { width: 24%; } th.col-new { width: 24%; }
td { padding: 8px 12px; border-bottom: 1px solid var(--border); vertical-align: top; overflow-wrap: break-word; word-break: break-word; }
td.status-cell { white-space: nowrap; font-size: 12px; }
td code, td pre { max-width: 100%; display: block; word-break: break-all; }
tr.data-row:hover { background: var(--bg-hover); }
tr.data-row:hover td { position: relative; }
tr.data-row:hover td code, tr.data-row:hover td pre { overflow: visible; white-space: pre-wrap; text-overflow: clip; }
.added    { border-left: 3px solid var(--green); } .added td:first-child { color: var(--green); font-weight: 600; }
.removed  { border-left: 3px solid var(--red); }   .removed td:first-child { color: var(--red); font-weight: 600; }
.modified { border-left: 3px solid var(--yellow); } .modified td:first-child { color: var(--yellow); font-weight: 600; }
.file-added   { border-left: 3px solid var(--green-light); } .file-added td { color: var(--green-light); font-weight: 600; }
.file-removed { border-left: 3px solid var(--red-light); }   .file-removed td { color: var(--red-light); font-weight: 600; }
code { font-family: 'Cascadia Code','Fira Code',Consolas,monospace; font-size: 12px; color: var(--accent); background: var(--bg-surface); padding: 2px 6px; border-radius: 4px; white-space: normal; word-break: break-all; }
pre { white-space: pre-wrap; word-break: break-word; margin: 0; font-family: 'Cascadia Code','Fira Code',Consolas,monospace; font-size: 12px; line-height: 1.6; }

/* Inline diff highlights */
.diff-del { background: var(--red-bg); color: var(--red); font-weight: 600; border-radius: 3px; padding: 0 2px; }
.diff-add { background: var(--green-bg); color: var(--green); font-weight: 600; border-radius: 3px; padding: 0 2px; }

/* Badges */
.badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 10px; font-weight: 600; }
.badge-added { background: var(--green-bg); color: var(--green); }
.badge-removed { background: var(--red-bg); color: var(--red); }
.badge-modified { background: var(--yellow-bg); color: var(--yellow); }
.badge-file-added { background: var(--green-bg-light); color: var(--green-light); }
.badge-file-removed { background: var(--red-bg-light); color: var(--red-light); }
.badge-total { background: rgba(96,205,255,.12); color: var(--accent); }

/* Policy group sub-headers */
.policy-group-row td { background: var(--bg-surface); font-weight: 600; font-size: 12px; color: var(--accent); padding: 6px 12px; border-bottom: 2px solid var(--border); letter-spacing: .2px; }
.policy-group-row td .pg-icon { font-family: 'Segoe UI Emoji','Segoe MDL2 Assets'; margin-right: 6px; opacity: .7; }
.policy-group-row td .pg-count { margin-left: 8px; font-size: 10px; font-weight: 400; color: var(--text-dim); background: var(--bg-card); padding: 1px 7px; border-radius: 8px; }

/* Custom Scrollbar */
::-webkit-scrollbar { width: 10px; height: 10px; }
::-webkit-scrollbar-track { background: var(--bg-surface); border-radius: 6px; }
::-webkit-scrollbar-thumb { background: var(--text-dim); border-radius: 6px; border: 2px solid var(--bg-surface); }
::-webkit-scrollbar-thumb:hover { background: var(--text-secondary); }
::-webkit-scrollbar-corner { background: var(--bg); }
* { scrollbar-width: thin; scrollbar-color: var(--text-dim) var(--bg-surface); }
.sidebar::-webkit-scrollbar { width: 6px; }
.sidebar::-webkit-scrollbar-thumb { border: 1px solid var(--bg-surface); }

/* Scroll to top */
.scroll-top { position: fixed; bottom: 24px; right: 24px; width: 40px; height: 40px; border-radius: 50%; background: var(--accent); color: #fff; border: none; cursor: pointer; font-size: 18px; display: none; z-index: 200; box-shadow: 0 4px 12px rgba(0,0,0,.3); transition: opacity .2s; }
.scroll-top:hover { opacity: .85; }

/* Print */
@media print {
  .header { position: static; } .sidebar, .header-actions, .scroll-top, .search-container, .filter-bar { display: none !important; }
  .container { display: block; } body { background: #fff; color: #000; font-size: 11px; padding: 12px; }
  .file-section, .stat-card, .version-bar, .toc { border-color: #ccc; background: #fff; }
  .file-header { color: #0078D4; } th { background: #f0f0f0; color: #333; }
  td, th { border-color: #ddd; } details { break-inside: avoid; }
  .added { background: #e8f5e9; } .removed { background: #ffebee; } .modified { background: #fff8e1; }
  .stat-card .num { color: #333 !important; } pre, code { font-size: 10px; }
}
</style></head><body>

<!-- Sticky Header -->
<div class="header">
  <div class="header-inner">
    <h1><span class="icon">&#x1F4CB;</span> ADMX Policy Comparison</h1>
    <div class="header-meta">$olderEnc &#x27A1; $newerEnc &nbsp;|&nbsp; $genDate</div>
    <div class="header-actions">
      <button onclick="toggleSidebar()" id="sidebarBtn" title="Toggle file list sidebar">&#x2630; Sidebar</button>
      <button onclick="toggleTheme()" id="themeBtn" title="Toggle light/dark theme">&#x1F319; Theme</button>
      <button onclick="expandAll()" title="Expand all sections">&#x25BC; Expand All</button>
      <button onclick="collapseAll()" title="Collapse all sections">&#x25B6; Collapse All</button>
      <button onclick="window.print()" title="Print report">&#x1F5A8; Print</button>
    </div>
  </div>
</div>

<div class="container">

<!-- Version Bar -->
<div class="version-bar">
  <div><strong>Source (Older)</strong><div class="ver">$olderEnc</div></div>
  <div class="arrow">&#x27A1;</div>
  <div><strong>Reference (Newer)</strong><div class="ver">$newerEnc</div></div>
  <div class="date">Generated $genDate</div>
</div>

<!-- Summary Cards -->
<div class="summary-bar">
  <div class="stat-card sc-added"><div class="num">$added</div><div class="label">Added</div></div>
  <div class="stat-card sc-removed"><div class="num">$removed</div><div class="label">Removed</div></div>
  <div class="stat-card sc-modified"><div class="num">$modified</div><div class="label">Modified</div></div>
  <div class="stat-card sc-fa"><div class="num">$fAdded</div><div class="label">Files Added</div></div>
  <div class="stat-card sc-fr"><div class="num">$fRemoved</div><div class="label">Files Removed</div></div>
  <div class="stat-card sc-total"><div class="num">$total</div><div class="label">Total Changes</div></div>
</div>

<!-- Search -->
<div class="search-container">
  <div class="search-wrap">
    <input type="text" class="search-box" id="searchInput" placeholder="Search policies, files, values..." oninput="filterRows()">
    <span class="search-count" id="searchCount"></span>
  </div>
</div>

<!-- Filter Pills -->
<div class="filter-bar">
  <button class="filter-pill active" data-filter="all" onclick="setFilter('all',this)">All ($total)</button>
  <button class="filter-pill" data-filter="Added" onclick="setFilter('Added',this)">&#x2795; Added ($added)</button>
  <button class="filter-pill" data-filter="Removed" onclick="setFilter('Removed',this)">&#x2796; Removed ($removed)</button>
  <button class="filter-pill" data-filter="Modified" onclick="setFilter('Modified',this)">&#x270F; Modified ($modified)</button>
  <button class="filter-pill" data-filter="File Added" onclick="setFilter('File Added',this)">&#x1F4C4; File Added ($fAdded)</button>
  <button class="filter-pill" data-filter="File Removed" onclick="setFilter('File Removed',this)">&#x1F5D1; File Removed ($fRemoved)</button>
</div>

<!-- Sidebar TOC -->
<div class="sidebar"><div class="toc"><h3>Files ($($grouped.Count))</h3>
$($tocHtml.ToString())
</div></div>

<!-- Main Content -->
<div class="main-content">
$($bodyHtml.ToString())
</div>

</div>

<!-- Scroll to top -->
<button class="scroll-top" id="scrollTop" onclick="window.scrollTo({top:0,behavior:'smooth'})">&#x2191;</button>

<script>
// Sidebar toggle
function toggleSidebar() {
  var c = document.querySelector('.container');
  c.classList.toggle('sidebar-collapsed');
  var btn = document.getElementById('sidebarBtn');
  btn.textContent = c.classList.contains('sidebar-collapsed') ? '\u2630 Show Sidebar' : '\u2630 Sidebar';
}

// Theme toggle
function toggleTheme() {
  document.body.classList.toggle('light');
  var btn = document.getElementById('themeBtn');
  btn.textContent = document.body.classList.contains('light') ? '\u2600\uFE0F Theme' : '\uD83C\uDF19 Theme';
}

// Expand/Collapse all
function expandAll()   { document.querySelectorAll('.file-section').forEach(function(d){ d.open = true; }); }
function collapseAll() { document.querySelectorAll('.file-section').forEach(function(d){ d.open = false; }); }

// Filter by type
var currentFilter = 'all';
function setFilter(type, btn) {
  currentFilter = type;
  document.querySelectorAll('.filter-pill').forEach(function(p){ p.classList.remove('active'); });
  btn.classList.add('active');
  filterRows();
}

// Search + filter
function filterRows() {
  var query = document.getElementById('searchInput').value.toLowerCase();
  var rows = document.querySelectorAll('.data-row');
  var shown = 0;
  rows.forEach(function(row) {
    var matchFilter = (currentFilter === 'all' || row.getAttribute('data-type') === currentFilter);
    var matchSearch = (!query || row.textContent.toLowerCase().indexOf(query) !== -1);
    row.style.display = (matchFilter && matchSearch) ? '' : 'none';
    if (matchFilter && matchSearch) shown++;
  });
  var countEl = document.getElementById('searchCount');
  if (query || currentFilter !== 'all') {
    countEl.textContent = shown + ' of $total';
  } else {
    countEl.textContent = '';
  }
  // Hide empty file sections + policy group headers with no visible rows
  document.querySelectorAll('.file-section').forEach(function(sec) {
    var hasVisible = false;
    sec.querySelectorAll('.data-row').forEach(function(r) { if (r.style.display !== 'none') hasVisible = true; });
    sec.style.display = hasVisible ? '' : 'none';
    // Toggle policy-group-row visibility based on sibling data-rows
    sec.querySelectorAll('.policy-group-row').forEach(function(pgr) {
      var sib = pgr.nextElementSibling, anyVis = false;
      while (sib && !sib.classList.contains('policy-group-row')) {
        if (sib.classList.contains('data-row') && sib.style.display !== 'none') anyVis = true;
        sib = sib.nextElementSibling;
      }
      pgr.style.display = anyVis ? '' : 'none';
    });
  });
}

// Scroll to top button
window.addEventListener('scroll', function() {
  document.getElementById('scrollTop').style.display = window.scrollY > 300 ? 'block' : 'none';
});

// Keyboard shortcut: Ctrl+F focuses search
document.addEventListener('keydown', function(e) {
  if ((e.ctrlKey || e.metaKey) && e.key === 'f') {
    e.preventDefault();
    document.getElementById('searchInput').focus();
  }
});
</script>
</body></html>
"@
    [System.IO.File]::WriteAllText($dlg.FileName, $html, [System.Text.Encoding]::UTF8)

    $Script:ExportCount++
    if ($Script:ExportCount -ge 1) { Unlock-Achievement 'export_master' }
    if ($Script:ExportCount -ge 5) { Unlock-Achievement 'batch_reporter' }

    Write-DebugLog "HTML export complete: $($dlg.FileName)" -Level SUCCESS
    Show-Toast 'Export Complete' "HTML report saved." 'success'
}

    <#
    .SYNOPSIS
        Exports the current (optionally filtered) comparison results as a CSV file.
    .DESCRIPTION
        Opens a SaveFileDialog and writes all visible results with columns: ChangeType,
        FileName, Category, StringId, RegistryKey, ValueName, Scope, OldValue, NewValue.
        Properly escapes embedded double quotes. Unlocks export achievements.
    #>
function Export-Csv {
    if ($Script:AllResults.Count -eq 0) {
        Show-Toast 'No Data' 'Run a comparison first before exporting.' 'warning'
        return
    }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Filter = 'CSV files (*.csv)|*.csv'
    $dlg.FileName = "ADMX_Comparison_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    if ($ui.ExportPathSetting -and $ui.ExportPathSetting.Text -and (Test-Path $ui.ExportPathSetting.Text)) {
        $dlg.InitialDirectory = $ui.ExportPathSetting.Text
    }
    if (-not $dlg.ShowDialog()) { return }

    Write-DebugLog "Exporting CSV to $($dlg.FileName)" -Level STEP

    # Use filtered view if active
    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($Script:AllResults)
    $csvData = @($view | ForEach-Object { $_ })
    if ($csvData.Count -lt $Script:AllResults.Count) {
        Write-DebugLog "Exporting filtered CSV: $($csvData.Count) of $($Script:AllResults.Count)" -Level INFO
    }
    $lines = @('"ChangeType","FileName","Category","StringId","RegistryKey","ValueName","Scope","OldValue","NewValue"')
    foreach ($r in $csvData) {
        $lines += '"{0}","{1}","{2}","{3}","{4}","{5}","{6}","{7}","{8}"' -f `
            ($r.ChangeType -replace '"','""'),
            ($r.FileName -replace '"','""'),
            ($r.Category -replace '"','""'),
            ($r.StringId -replace '"','""'),
            ($r.RegistryKey -replace '"','""'),
            ($r.ValueName -replace '"','""'),
            ($r.Scope -replace '"','""'),
            ($r.OldValue -replace '"','""'),
            ($r.NewValue -replace '"','""')
    }
    [System.IO.File]::WriteAllLines($dlg.FileName, $lines, [System.Text.Encoding]::UTF8)

    $Script:ExportCount++
    if ($Script:ExportCount -ge 1) { Unlock-Achievement 'export_master' }
    if ($Script:ExportCount -ge 5) { Unlock-Achievement 'batch_reporter' }

    Write-DebugLog "CSV export complete: $($dlg.FileName)" -Level SUCCESS
    Show-Toast 'Export Complete' "CSV saved." 'success'
}

$ui.BtnExportHtml.Add_Click({ Export-Html })
$ui.BtnExportCsv.Add_Click({ Export-Csv })

# -- Getting Started navigation buttons ----------------------------------------
if ($ui.BtnGetStarted) {
    $ui.BtnGetStarted.Add_Click({ Switch-Tab 'Compare' })
}
if ($ui.BtnResultsGoCompare) {
    $ui.BtnResultsGoCompare.Add_Click({ Switch-Tab 'Compare' })
}

# -- ADMX Template download links ---------------------------------------------------
$Script:DownloadUrls = @{
    'Win10_22H2' = 'https://www.microsoft.com/en-us/download/details.aspx?id=104677'
    'Win11_23H2' = 'https://www.microsoft.com/en-us/download/details.aspx?id=105667'
    'Win11_24H2' = 'https://www.microsoft.com/en-us/download/details.aspx?id=106255'
    'Win11_25H2' = 'https://www.microsoft.com/en-us/download/details.aspx?id=106621'
    'Win11_26H1' = 'https://www.microsoft.com/en-us/download/details.aspx?id=107188'
}

foreach ($linkKey in $Script:DownloadUrls.Keys) {
    $url = $Script:DownloadUrls[$linkKey]
    foreach ($prefix in @('Link','SidebarLink')) {
        $el = $ui["${prefix}${linkKey}"]
        if ($el) {
            $urlRef = $url
            $el.Add_MouseLeftButtonDown({
                [System.Diagnostics.Process]::Start([System.Diagnostics.ProcessStartInfo]@{ FileName = $urlRef; UseShellExecute = $true })
                Write-DebugLog "Opened download: $urlRef" -Level INFO
            }.GetNewClosure())
            # Underline on hover
            $elRef = $el
            $el.Add_MouseEnter({
                $elRef.Opacity = 0.7
            }.GetNewClosure())
            $el.Add_MouseLeave({
                $elRef.Opacity = 1.0
            }.GetNewClosure())
        }
    }
}

# -- Results empty state visibility helper -------------------------------------
    <#
    .SYNOPSIS
        Shows or hides the no-results empty-state guide based on whether results exist.
    #>
function Update-ResultsEmptyState {
    if ($ui.ResultsEmptyGuide) {
        if ($Script:AllResults.Count -eq 0) {
            $ui.ResultsEmptyGuide.Visibility = 'Visible'
        } else {
            $ui.ResultsEmptyGuide.Visibility = 'Collapsed'
        }
    }
}
$ui.BtnQuickExport.Add_Click({
    # Load the selected history entry before exporting
    if ($ui.ExportHistoryCombo -and $ui.ExportHistoryCombo.SelectedItem) {
        $sel = $ui.ExportHistoryCombo.SelectedItem.Tag
        if ($sel -and $sel.Id) {
            $loaded = Load-ComparisonResults $sel.Id
            if ($loaded) {
                $Script:LoadedOlderVersion = $sel.OlderVersion
                $Script:LoadedNewerVersion = $sel.NewerVersion
                Apply-Filters
            } else {
                Show-Toast 'Not Found' 'Result data for this comparison is no longer available.' 'warning'
                return
            }
        }
    }
    Export-Html
})

# -- Keyboard shortcuts --------------------------------------------------------
$Window.Add_KeyDown({
    if ($_.Key -eq 'F1') { if ($ui.BtnHelp) { $ui.BtnHelp.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent)) } }
    elseif ($_.Key -eq 'F5') { Switch-Tab 'Compare'; Run-Comparison }
    elseif ($_.Key -eq 'E' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') { Export-Html }
    elseif ($_.Key -eq 'F' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') {
        Switch-Tab 'Results'; $ui.SearchBox.Focus()
    }
    elseif ($_.Key -eq 'OemTilde' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') {
        Toggle-ConsolePanel
    }
    elseif ($_.Key -eq 'B' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') {
        if ($ui.NavToggleSidebar) { $ui.NavToggleSidebar.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent)) }
    }
    elseif ($_.Key -eq 'D1' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'None') {
        if (-not ($ui.SearchBox.IsFocused -or $ui.BasePathSetting.IsFocused -or $ui.ExportPathSetting.IsFocused)) { Switch-Tab 'Dashboard' }
    }
    elseif ($_.Key -eq 'D2' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'None') {
        if (-not ($ui.SearchBox.IsFocused -or $ui.BasePathSetting.IsFocused -or $ui.ExportPathSetting.IsFocused)) { Switch-Tab 'Compare' }
    }
    elseif ($_.Key -eq 'D3' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'None') {
        if (-not ($ui.SearchBox.IsFocused -or $ui.BasePathSetting.IsFocused -or $ui.ExportPathSetting.IsFocused)) { Switch-Tab 'Results' }
    }
    elseif ($_.Key -eq 'D4' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'None') {
        if (-not ($ui.SearchBox.IsFocused -or $ui.BasePathSetting.IsFocused -or $ui.ExportPathSetting.IsFocused)) { Switch-Tab 'Settings' }
    }
})

# -- UI Micro-interactions (WinUI 3 / Robinhood patterns) ---------------------

# IMPROVEMENT #3/#17: Stat Card Hover — lift effect with border glow
foreach ($cardName in @('StatCard1','StatCard2','StatCard3','StatCard4')) {
    $card = $ui[$cardName]
    if ($card) {
        $card.Add_MouseEnter({
            $this.BorderBrush = $Window.Resources['ThemeBorderElevated']
            $this.Effect = [System.Windows.Media.Effects.DropShadowEffect]@{
                Color = [System.Windows.Media.Colors]::Black; BlurRadius = 12
                ShadowDepth = 2; Opacity = 0.25
            }
        }.GetNewClosure())
        $card.Add_MouseLeave({
            $this.BorderBrush = $Window.Resources['ThemeBorderCard']
            $this.Effect = $null
        }.GetNewClosure())
    }
}

# IMPROVEMENT #9: Search Box Icon Color — accent on focus
if ($ui.SearchBox -and $ui.SearchIcon) {
    $ui.SearchBox.Add_GotFocus({
        $ui.SearchIcon.Foreground = $Window.Resources['ThemeAccent']
    }.GetNewClosure())
    $ui.SearchBox.Add_LostFocus({
        $ui.SearchIcon.Foreground = $Window.Resources['ThemeTextDim']
    }.GetNewClosure())
}

# IMPROVEMENT #11: Settings cards don't need PS1 — hover handled via #3 pattern
# (Settings cards are static borders without names — keeping them clean)

# IMPROVEMENT #10: Status Bar Enhancement — update results count on filter
    <#
    .SYNOPSIS
        Updates the status bar with the current visible result count and streak display.
    #>
function Update-StatusBar {
    if ($ui.StatusResults) {
        $count = if ($ui.ResultsDataGrid.ItemsSource) {
            @($ui.ResultsDataGrid.ItemsSource).Count
        } else { 0 }
        $ui.StatusResults.Text = "$count results"
    }
    if ($ui.StatusStreak) {
        $streakCount = [int]$Script:Streak.Count
        $ui.StatusStreak.Text = if ($streakCount -gt 0) { "$streakCount day streak" } else { '' }
    }
}

# -- Context Menu (right-click on DataGrid rows) -------------------------------
if ($ui.ResultsDataGrid) {
    $ctxMenu = [System.Windows.Controls.ContextMenu]::new()
    $ctxMenu.SetResourceReference([System.Windows.Controls.Control]::BackgroundProperty, 'ThemeCardBg')
    $ctxMenu.SetResourceReference([System.Windows.Controls.Control]::BorderBrushProperty, 'ThemeBorderElevated')
    $ctxMenu.BorderThickness = [System.Windows.Thickness]::new(1)
    $ctxMenu.Padding          = [System.Windows.Thickness]::new(2)

    $menuItems = @(
        @{ Header='Copy String ID'; Icon=[char]0xE8C8; Action={
            $sel = $ui.ResultsDataGrid.SelectedItem
            if ($sel -and $sel.StringId) { [System.Windows.Clipboard]::SetText($sel.StringId); Show-Toast 'Copied' $sel.StringId 'success' }
        }},
        @{ Header='Copy Old Value'; Icon=[char]0xE8C8; Action={
            $sel = $ui.ResultsDataGrid.SelectedItem
            if ($sel -and $sel.OldValue) { [System.Windows.Clipboard]::SetText($sel.OldValue); Show-Toast 'Copied' 'Old value copied' 'success' }
        }},
        @{ Header='Copy New Value'; Icon=[char]0xE8C8; Action={
            $sel = $ui.ResultsDataGrid.SelectedItem
            if ($sel -and $sel.NewValue) { [System.Windows.Clipboard]::SetText($sel.NewValue); Show-Toast 'Copied' 'New value copied' 'success' }
        }},
        @{ Header='Copy Row as Text'; Icon=[char]0xE8C8; Action={
            $sel = $ui.ResultsDataGrid.SelectedItem
            if ($sel) {
                $txt = "[$($sel.ChangeType)] $($sel.FileName) | $($sel.StringId)`nOld: $($sel.OldValue)`nNew: $($sel.NewValue)"
                if ($sel.RegistryKey) { $txt += "`nRegistry: $($sel.RegistryKey)\$($sel.ValueName)" }
                [System.Windows.Clipboard]::SetText($txt); Show-Toast 'Copied' 'Full row copied' 'success'
            }
        }},
        @{ Header='Filter to this file'; Icon=[char]0xE71C; Action={
            $sel = $ui.ResultsDataGrid.SelectedItem
            if ($sel -and $sel.FileName) { $ui.SearchBox.Text = $sel.FileName }
        }}
    )
    foreach ($mi in $menuItems) {
        $menuItem = [System.Windows.Controls.MenuItem]::new()
        $menuItem.Header = $mi.Header
        $menuItem.SetResourceReference([System.Windows.Controls.Control]::ForegroundProperty, 'ThemeTextBody')
        $iconTb = [System.Windows.Controls.TextBlock]::new()
        $iconTb.Text = [string]$mi.Icon
        $iconTb.FontFamily = [System.Windows.Media.FontFamily]::new('Segoe MDL2 Assets')
        $iconTb.FontSize = 12
        $iconTb.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
        $menuItem.Icon = $iconTb
        $act = $mi.Action
        $menuItem.Add_Click($act)
        [void]$ctxMenu.Items.Add($menuItem)
    }
    $ui.ResultsDataGrid.ContextMenu = $ctxMenu
    Write-DebugLog 'Context menu wired to DataGrid' -Level DEBUG
}

# -- Detail Pane: word-level diff engine ---------------------------------------
    <#
    .SYNOPSIS
        Computes word-level diff runs between two strings using LCS.
    .DESCRIPTION
        Splits both texts on whitespace boundaries, builds an LCS (Longest Common
        Subsequence) table using a jagged array, backtracks to produce a minimal edit
        script, and returns parallel run arrays for old and new text with each run
        tagged as plain, deleted, or added. Falls back to plain runs for texts
        exceeding 500 words to avoid performance issues.
    .OUTPUTS
        [hashtable] With keys OldRuns and NewRuns, each an array of @{ Text; Type }.
    #>
function Get-WordDiffRuns([string]$oldText, [string]$newText) {
    $oldWords = if ($oldText) { $oldText -split '(\s+)' } else { @() }
    $newWords = if ($newText) { $newText -split '(\s+)' } else { @() }
    $m = $oldWords.Count; $n = $newWords.Count
    # Skip LCS for very long texts — too slow in PS
    if ($m -gt 500 -or $n -gt 500) {
        return @{
            OldRuns = @(@{ Text=$oldText; Type='plain' })
            NewRuns = @(@{ Text=$newText; Type='plain' })
        }
    }
    # Build LCS table — use jagged array to avoid PS [int[,]] indexing pitfalls
    $dp = [int[][]]::new($m + 1)
    for ($x = 0; $x -le $m; $x++) { $dp[$x] = [int[]]::new($n + 1) }
    for ($i = 1; $i -le $m; $i++) {
        for ($j = 1; $j -le $n; $j++) {
            if ($oldWords[$i-1] -ceq $newWords[$j-1]) {
                $dp[$i][$j] = $dp[$i-1][$j-1] + 1
            } else {
                $dp[$i][$j] = [Math]::Max($dp[$i-1][$j], $dp[$i][$j-1])
            }
        }
    }
    # Backtrack
    $ops = [System.Collections.Generic.List[string]]::new()
    $i = $m; $j = $n
    while ($i -gt 0 -or $j -gt 0) {
        if ($i -gt 0 -and $j -gt 0 -and $oldWords[$i-1] -ceq $newWords[$j-1]) {
            $ops.Add("S:$($i-1):$($j-1)"); $i--; $j--
        } elseif ($j -gt 0 -and ($i -eq 0 -or $dp[$i][$j-1] -ge $dp[$i-1][$j])) {
            $ops.Add("A:$($j-1)"); $j--
        } else {
            $ops.Add("D:$($i-1)"); $i--
        }
    }
    $ops.Reverse()
    # Build runs
    $oldRuns = [System.Collections.Generic.List[hashtable]]::new()
    $newRuns = [System.Collections.Generic.List[hashtable]]::new()
    $curOldSame = ''; $curOldDel = ''; $curNewSame = ''; $curNewAdd = ''
    foreach ($op in $ops) {
        $parts = $op -split ':', 3
        switch ($parts[0]) {
            'S' {
                if ($curOldDel) { $oldRuns.Add(@{ Text=$curOldDel; Type='deleted' }); $curOldDel = '' }
                if ($curNewAdd) { $newRuns.Add(@{ Text=$curNewAdd; Type='added' });   $curNewAdd = '' }
                $curOldSame += $oldWords[[int]$parts[1]]
                $curNewSame += $newWords[[int]$parts[2]]
            }
            'D' {
                if ($curOldSame) { $oldRuns.Add(@{ Text=$curOldSame; Type='plain' }); $curOldSame = '' }
                if ($curNewSame) { $newRuns.Add(@{ Text=$curNewSame; Type='plain' }); $curNewSame = '' }
                $curOldDel += $oldWords[[int]$parts[1]]
            }
            'A' {
                if ($curOldSame) { $oldRuns.Add(@{ Text=$curOldSame; Type='plain' }); $curOldSame = '' }
                if ($curNewSame) { $newRuns.Add(@{ Text=$curNewSame; Type='plain' }); $curNewSame = '' }
                $curNewAdd += $newWords[[int]$parts[1]]
            }
        }
    }
    if ($curOldSame) { $oldRuns.Add(@{ Text=$curOldSame; Type='plain' }) }
    if ($curOldDel)  { $oldRuns.Add(@{ Text=$curOldDel;  Type='deleted' }) }
    if ($curNewSame) { $newRuns.Add(@{ Text=$curNewSame; Type='plain' }) }
    if ($curNewAdd)  { $newRuns.Add(@{ Text=$curNewAdd;  Type='added' }) }
    return @{ OldRuns=$oldRuns; NewRuns=$newRuns }
}

# -- Detail Pane: selection handler --------------------------------------------
    <#
    .SYNOPSIS
        Populates the detail pane with the selected DataGrid row diff and metadata.
    .DESCRIPTION
        Updates the detail header (icon, title, file name), registry info bar (key path,
        value name, scope badge, value type badge, policy element details from PolicyMetadata),
        and renders word-level diff runs in the side-by-side old/new RichTextBox panes
        with colour-coded inline highlighting.
    #>
function Update-DetailPane {
    $sel = $ui.ResultsDataGrid.SelectedItem
    if (-not $sel -or -not $ui.paraDetailOld) { return }
    # Header
    $typeIcons = @{ 'Added'=[char]0xE710; 'Removed'=[char]0xE738; 'Modified'=[char]0xE70F
                    'File Added'=[char]0xE8E5; 'File Removed'=[char]0xE74D }
    $ui.lblDetailIcon.Text  = if ($typeIcons[$sel.ChangeType]) { $typeIcons[$sel.ChangeType] } else { [char]0xE946 }
    $ui.lblDetailTitle.Text = "$($sel.ChangeType): $($sel.StringId)"
    $catLabel  = if ($sel.Category)  { "  |  Category: $($sel.Category)" } else { '' }
    $typeLabel = if ($sel.StringType -and $sel.StringType -ne 'Display') { "  |  $($sel.StringType)" } else { '' }
    $ui.lblDetailFile.Text  = "$($sel.FileName)$catLabel$typeLabel"
    # Registry info bar
    if ($sel.RegistryKey -and $ui.pnlRegistryBar) {
        $ui.pnlRegistryBar.Visibility = 'Visible'
        $ui.lblRegistryPath.Text  = $sel.RegistryKey
        $vn = if ($sel.ValueName) { "  \  $($sel.ValueName)" } else { '' }
        $vtBadge = if ($sel.ValueType) { "  [$($sel.ValueType)]" } else { '' }
        $ui.lblRegistryValue.Text = "$vn$vtBadge"
        if ($sel.Scope -and $sel.Scope -ne 'Both' -and $ui.badgeRegistryScope) {
            $ui.badgeRegistryScope.Visibility = 'Visible'
            $ui.lblRegistryScope.Text = $sel.Scope.ToUpper()
        } else {
            if ($ui.badgeRegistryScope) { $ui.badgeRegistryScope.Visibility = 'Collapsed' }
        }
        # Show policy element details (enum values, ranges) if available
        if ($ui.lblRegistryDetails) {
            $policyName = $sel.PolicyGroup
            $meta = if ($policyName -and $Script:PolicyMetadata.ContainsKey($policyName)) { $Script:PolicyMetadata[$policyName] } else { $null }
            if ($meta -and ($meta.Elements.Count -gt 0 -or $meta.EnabledValue -or $meta.DisabledValue)) {
                $parts = @()
                if ($meta.EnabledValue) { $parts += "Enabled=$($meta.EnabledValue)" }
                if ($meta.DisabledValue) { $parts += "Disabled=$($meta.DisabledValue)" }
                foreach ($el in $meta.Elements) {
                    $elDesc = "$($el.Type)"
                    if ($el.ValueName) { $elDesc += " '$($el.ValueName)'" }
                    elseif ($el.Id) { $elDesc += " '$($el.Id)'" }
                    if ($el.ValueType) { $elDesc += " ($($el.ValueType))" }
                    if ($el.MinValue -or $el.MaxValue) { $elDesc += " [$($el.MinValue)..$($el.MaxValue)]" }
                    if ($el.Choices -and $el.Choices.Count -gt 0) {
                        $choiceStr = ($el.Choices | ForEach-Object { $_ -replace '\$\(string\.(.+?)\)','$1' }) -join ' | '
                        $elDesc += ": $choiceStr"
                    }
                    $parts += $elDesc
                }
                $ui.lblRegistryDetails.Text = $parts -join '  ·  '
                $ui.lblRegistryDetails.Visibility = 'Visible'
            } else {
                $ui.lblRegistryDetails.Text = ''
                $ui.lblRegistryDetails.Visibility = 'Collapsed'
            }
        }
    } else {
        if ($ui.pnlRegistryBar) { $ui.pnlRegistryBar.Visibility = 'Collapsed' }
    }
    # Clear
    $ui.paraDetailOld.Inlines.Clear()
    $ui.paraDetailNew.Inlines.Clear()
    $brConv = [System.Windows.Media.BrushConverter]::new()
    $delBg  = $brConv.ConvertFromString('#40FF5000')
    $addBg  = $brConv.ConvertFromString('#4000C853')
    $delFg  = $brConv.ConvertFromString('#FFFF6E40')
    $addFg  = $brConv.ConvertFromString('#FF69F0AE')
    $plainFg = $Window.Resources['ThemeTextBody']
    $dimFg   = $Window.Resources['ThemeTextDim']
    if ($sel.ChangeType -eq 'Modified') {
        $diff = Get-WordDiffRuns $sel.OldValue $sel.NewValue
        foreach ($run in $diff.OldRuns) {
            $r = [System.Windows.Documents.Run]::new($run.Text)
            if ($run.Type -eq 'deleted') { $r.Background = $delBg; $r.Foreground = $delFg; $r.FontWeight = 'SemiBold' }
            else { $r.Foreground = $plainFg }
            $r.FontSize = 12.5
            $ui.paraDetailOld.Inlines.Add($r)
        }
        foreach ($run in $diff.NewRuns) {
            $r = [System.Windows.Documents.Run]::new($run.Text)
            if ($run.Type -eq 'added') { $r.Background = $addBg; $r.Foreground = $addFg; $r.FontWeight = 'SemiBold' }
            else { $r.Foreground = $plainFg }
            $r.FontSize = 12.5
            $ui.paraDetailNew.Inlines.Add($r)
        }
    } else {
        foreach ($pair in @(@{Val=$sel.OldValue; Para=$ui.paraDetailOld}, @{Val=$sel.NewValue; Para=$ui.paraDetailNew})) {
            if ($pair.Val) {
                $r = [System.Windows.Documents.Run]::new($pair.Val); $r.Foreground = $plainFg; $r.FontSize = 12.5
            } else {
                $r = [System.Windows.Documents.Run]::new('(empty)'); $r.Foreground = $dimFg; $r.FontSize = 12; $r.FontStyle = 'Italic'
            }
            $pair.Para.Inlines.Add($r)
        }
    }
    Write-DebugLog "Detail pane: $($sel.StringId)" -Level DEBUG
}

if ($ui.ResultsDataGrid) {
    $ui.ResultsDataGrid.Add_SelectionChanged({ Update-DetailPane })
}

# -- Detail Pane Copy Buttons -------------------------------------------------
foreach ($btn in @(
    @{ Name='btnCopyOldValue'; Prop='OldValue'; Msg='Old value copied' }
    @{ Name='btnCopyNewValue'; Prop='NewValue'; Msg='New value copied' }
    @{ Name='btnCopyStringId'; Prop='StringId'; Msg='String ID copied' }
)) {
    if ($ui[$btn.Name]) {
        $propName = $btn.Prop; $msg = $btn.Msg
        $ui[$btn.Name].Add_Click({
            $sel = $ui.ResultsDataGrid.SelectedItem
            $val = $sel.$propName
            if ($sel -and $val) { [System.Windows.Clipboard]::SetText([string]$val); Show-Toast 'Copied' $msg 'success' }
        }.GetNewClosure())
    }
}

# -- Dashboard Breakdown Chart ------------------------------------------------
    <#
    .SYNOPSIS
        Renders the horizontal stacked bar chart and legend in the Results panel.
    .DESCRIPTION
        Counts results by change type, creates proportionally sized coloured bars,
        and builds a dot-legend with type labels and counts.
    #>
function Update-BreakdownChart {
    if (-not $ui.pnlChartBar -or -not $ui.pnlChartLegend) { return }
    $ui.pnlChartBar.Children.Clear()
    $ui.pnlChartLegend.Children.Clear()
    if ($Script:AllResults.Count -eq 0) {
        if ($ui.pnlBreakdownChart) { $ui.pnlBreakdownChart.Visibility = 'Collapsed' }
        return
    }
    if ($ui.pnlBreakdownChart) { $ui.pnlBreakdownChart.Visibility = 'Visible' }
    $brConv = [System.Windows.Media.BrushConverter]::new()
    $counts = @{
        'Added'        = @($Script:AllResults | Where-Object { $_.ChangeType -eq 'Added' }).Count
        'Removed'      = @($Script:AllResults | Where-Object { $_.ChangeType -eq 'Removed' }).Count
        'Modified'     = @($Script:AllResults | Where-Object { $_.ChangeType -eq 'Modified' }).Count
        'File Added'   = @($Script:AllResults | Where-Object { $_.ChangeType -eq 'File Added' }).Count
        'File Removed' = @($Script:AllResults | Where-Object { $_.ChangeType -eq 'File Removed' }).Count
    }
    $colors = @{ 'Added'='#00C853'; 'Removed'='#FF5000'; 'Modified'='#F59E0B'; 'File Added'='#00E676'; 'File Removed'='#FF6E40' }
    $total = $Script:AllResults.Count
    # Build a single-row Grid with proportional * columns so bars always fill the full width
    $barGrid = [System.Windows.Controls.Grid]::new()
    $barGrid.Height = 28
    $colIdx = 0
    foreach ($type in @('Added','Removed','Modified','File Added','File Removed')) {
        $c = $counts[$type]; if ($c -eq 0) { continue }
        $pct = $c / $total
        $col = [System.Windows.Controls.ColumnDefinition]::new()
        $col.Width = [System.Windows.GridLength]::new($pct, [System.Windows.GridUnitType]::Star)
        $barGrid.ColumnDefinitions.Add($col)
        $bar = [System.Windows.Controls.Border]::new()
        $bar.Background = $brConv.ConvertFromString($colors[$type])
        $bar.ToolTip = "$type`: $c ($([Math]::Round($pct * 100, 1))%)"
        [System.Windows.Controls.Grid]::SetColumn($bar, $colIdx)
        [void]$barGrid.Children.Add($bar)
        $colIdx++
        # Legend
        $sp = [System.Windows.Controls.StackPanel]::new(); $sp.Orientation = 'Horizontal'
        $sp.Margin = [System.Windows.Thickness]::new(0,0,16,0)
        $dot = [System.Windows.Shapes.Ellipse]::new(); $dot.Width = 8; $dot.Height = 8
        $dot.Margin = [System.Windows.Thickness]::new(0,0,6,0)
        $dot.Fill = $brConv.ConvertFromString($colors[$type]); $dot.VerticalAlignment = 'Center'
        $lbl = [System.Windows.Controls.TextBlock]::new(); $lbl.Text = "$type ($c)"; $lbl.FontSize = 11
        $lbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextSecondary')
        $lbl.VerticalAlignment = 'Center'
        [void]$sp.Children.Add($dot); [void]$sp.Children.Add($lbl)
        [void]$ui.pnlChartLegend.Children.Add($sp)
    }
    [void]$ui.pnlChartBar.Children.Add($barGrid)
    Write-DebugLog "Breakdown chart updated ($total total)" -Level DEBUG
}


# -- Window lifecycle ----------------------------------------------------------
$Window.Add_Loaded({
    Write-DebugLog 'Application loaded' -Level SYSTEM
    Write-DebugLog "Version: $($Script:AppVersion)" -Level SYSTEM
    Write-DebugLog "AppDir: $($Script:AppDir)" -Level SYSTEM

    # Restore window position (clamped to screen bounds)
    if ($Script:Prefs.WindowLeft -ge 0) {
        $sw = [System.Windows.SystemParameters]::VirtualScreenWidth
        $sh = [System.Windows.SystemParameters]::VirtualScreenHeight
        $sl = [System.Windows.SystemParameters]::VirtualScreenLeft
        $st = [System.Windows.SystemParameters]::VirtualScreenTop
        $w  = [Math]::Min([int]$Script:Prefs.WindowWidth,  [int]$sw)
        $h  = [Math]::Min([int]$Script:Prefs.WindowHeight, [int]$sh)
        $l  = [Math]::Max([int]$sl, [Math]::Min([int]$Script:Prefs.WindowLeft, [int]($sl + $sw - $w)))
        $t  = [Math]::Max([int]$st, [Math]::Min([int]$Script:Prefs.WindowTop,  [int]($st + $sh - $h)))
        $Window.Left   = $l
        $Window.Top    = $t
        $Window.Width  = $w
        $Window.Height = $h
    }
    if ($Script:Prefs.WindowState -eq 'Maximized') { $Window.WindowState = 'Maximized' }

    # Auto-scan
    Scan-Versions
    Update-Dashboard
    Render-AchievementBadges
    Update-ResultsEmptyState

    # Auto-load most recent comparison results
    if ($Script:History.Count -gt 0 -and $Script:History[0].Id) {
        $loaded = Load-ComparisonResults $Script:History[0].Id
        if ($loaded) {
            $Script:LoadedOlderVersion = $Script:History[0].OlderVersion
            $Script:LoadedNewerVersion = $Script:History[0].NewerVersion
            $ui.ResultsSubtitle.Text = "$($Script:History[0].OlderVersion) -> $($Script:History[0].NewerVersion)  |  $($Script:History[0].ChangeCount) changes"
            Apply-Filters
            Write-DebugLog "Auto-loaded last comparison: $($Script:AllResults.Count) results" -Level INFO
        }
    }

    Write-DebugLog "Loaded $($Script:Unlocked.Count) unlocked achievements" -Level INFO
    Write-DebugLog "Loaded $($Script:History.Count) history entries" -Level INFO
    Write-DebugLog "Streak: $($Script:Streak.Count) days" -Level INFO
    Write-DebugLog 'Ready. Press Ctrl+` to toggle debug console.' -Level SYSTEM
})

$Window.Add_Closing({
    Write-DebugLog 'Application closing - saving preferences' -Level SYSTEM
    Save-Preferences
})

# -- Window Loaded Animation (#19) + Additional Micro-interactions -------------

# #19: Window entry animation — opacity + slide-up
$Window.Add_ContentRendered({
    if ($Script:AnimationsDisabled) { return }
    $ui.WindowBorder.Opacity = 0
    $ui.WindowBorder.RenderTransform = [System.Windows.Media.TranslateTransform]::new(0, 10)

    $fadeIn = [System.Windows.Media.Animation.DoubleAnimation]::new(0, 1,
        [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(250)))
    $fadeIn.EasingFunction = [System.Windows.Media.Animation.CubicEase]::new()
    $fadeIn.EasingFunction.EasingMode = 'EaseOut'
    $ui.WindowBorder.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $fadeIn)

    $slideUp = [System.Windows.Media.Animation.DoubleAnimation]::new(10, 0,
        [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(250)))
    $slideUp.EasingFunction = [System.Windows.Media.Animation.CubicEase]::new()
    $slideUp.EasingFunction.EasingMode = 'EaseOut'
    $ui.WindowBorder.RenderTransform.BeginAnimation(
        [System.Windows.Media.TranslateTransform]::YProperty, $slideUp)

    # #38: Status bar slide-in with delay
    $statusDotEl = $Window.FindName('statusDot')
    if ($statusDotEl -and $statusDotEl.Parent -and $statusDotEl.Parent.Parent) {
        $statusRow = $statusDotEl.Parent.Parent
        if ($statusRow -is [System.Windows.Controls.Border]) {
            $statusRow.RenderTransform = [System.Windows.Media.TranslateTransform]::new(0, 8)
            $statusRow.Opacity = 0
            $statusSlide = [System.Windows.Media.Animation.DoubleAnimation]::new(8, 0,
                [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(250)))
            $statusSlide.BeginTime = [TimeSpan]::FromMilliseconds(200)
            $statusSlide.EasingFunction = [System.Windows.Media.Animation.CubicEase]::new()
            $statusRow.RenderTransform.BeginAnimation([System.Windows.Media.TranslateTransform]::YProperty, $statusSlide)
            $statusFade = [System.Windows.Media.Animation.DoubleAnimation]::new(0, 1,
                [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(250)))
            $statusFade.BeginTime = [TimeSpan]::FromMilliseconds(200)
            $statusRow.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $statusFade)
        }
    }
}.GetNewClosure())

# #18: Toast icon type mapping
    <#
    .SYNOPSIS
        Sets the toast notification icon and accent colour based on message type.
    .PARAMETER type
        One of: success, warning, error, or info (default).
    #>
function Set-ToastType([string]$type) {
    if (-not $ui.ToastIcon -or -not $ui.ToastAccentBar) { return }
    switch ($type) {
        'success' {
            $ui.ToastIcon.Text = [char]0xE73E
            $ui.ToastIcon.Foreground = $Window.Resources['ThemeSuccess']
            $ui.ToastAccentBar.Background = $Window.Resources['ThemeSuccess']
        }
        'warning' {
            $ui.ToastIcon.Text = [char]0xE7BA
            $ui.ToastIcon.Foreground = $Window.Resources['ThemeWarning']
            $ui.ToastAccentBar.Background = $Window.Resources['ThemeWarning']
        }
        'error' {
            $ui.ToastIcon.Text = [char]0xEA39
            $ui.ToastIcon.Foreground = $Window.Resources['ThemeError']
            $ui.ToastAccentBar.Background = $Window.Resources['ThemeError']
        }
        default {
            $ui.ToastIcon.Text = [char]0xE946
            $ui.ToastIcon.Foreground = $Window.Resources['ThemeAccent']
            $ui.ToastAccentBar.Background = $Window.Resources['ThemeAccent']
        }
    }
}

# #6: Scrollbar auto-hide on MinimalScrollViewer
$scrollViewerStyle = $Window.Resources['MinimalScrollViewer']
# (auto-hide is handled at the scrollbar style level — see Opacity=0.5 in XAML)
# We enhance it by making the MinimalScrollViewer's scrollbar fade on mouse events

# #12: Title bar chrome — ensure buttons are hit-test visible
foreach ($btnName in @('BtnHelp','BtnThemeToggle','BtnMinimize','BtnMaximize','BtnClose')) {
    $btn = $ui[$btnName]
    if ($btn) {
        [System.Windows.Shell.WindowChrome]::SetIsHitTestVisibleInChrome($btn, $true)
    }
}

# -- Keyboard shortcuts (enhanced) ---------------------------------------------
$Window.Add_KeyDown({
    param($s, $e)
    $focused = [System.Windows.Input.Keyboard]::FocusedElement
    $inTextInput = $focused -is [System.Windows.Controls.TextBox] -or
                   $focused -is [System.Windows.Controls.ComboBox] -or
                   $focused -is [System.Windows.Controls.RichTextBox]
    if ($e.Key -eq 'C' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control' -and -not $inTextInput) {
        $sel = $ui.ResultsDataGrid.SelectedItem
        if ($sel) {
            $txt = "[$($sel.ChangeType)] $($sel.FileName) | $($sel.StringId)`nOld: $($sel.OldValue)`nNew: $($sel.NewValue)"
            if ($sel.RegistryKey) { $txt += "`nRegistry: $($sel.RegistryKey)\$($sel.ValueName)" }
            [System.Windows.Clipboard]::SetText($txt)
            Show-Toast 'Copied' 'Row copied to clipboard' 'success'
            $e.Handled = $true
        }
    }
})

# -- Initialize & launch -------------------------------------------------------
Load-Preferences
Load-Achievements
Load-History
Load-Streak

$Window.ShowDialog() | Out-Null
