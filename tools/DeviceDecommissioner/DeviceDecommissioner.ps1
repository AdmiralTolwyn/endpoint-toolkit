# ============================================================================
# Device Decommissioner
# ----------------------------------------------------------------------------
# A WPF GUI tool to remove a single device (workstation) from:
#   - Active Directory   (Remove-ADComputer)
#   - Entra ID           (Remove-MgDevice)
#   - Microsoft Intune   (Remove-MgDeviceManagementManagedDevice)
#   - SCCM/MEMCM         (Remove-CMDevice)
#
# Credentials for AD and SCCM are stored DPAPI-encrypted (per-user) in
# user_creds.dat next to the script. Modern endpoints (Entra ID + Intune)
# use interactive Connect-MgGraph sign-in.
#
# Design language and helpers (theme, log panel, toast, modal, background
# runspace pattern) mirror AIBPipelineCreator 1:1.
# ============================================================================

[CmdletBinding()]
param()

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Per-monitor DPI awareness (crisp at 125%/150%/175%/200% scaling)
try {
    Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;
public static class DpiHelper {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SetProcessDpiAwarenessContext(int value);
}
'@ -ErrorAction SilentlyContinue
    # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
    [DpiHelper]::SetProcessDpiAwarenessContext(-4) | Out-Null
} catch { <# Win10 1703+ only; older OS falls back to system DPI #> }

# ============================================================================
# GLOBALS & CONSTANTS
# ============================================================================
$Global:Root        = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$Global:AppVersion  = '0.1.0-alpha'
$Global:AppTitle    = 'Device Decommissioner'
$Global:SettingsFile = Join-Path $Global:Root 'user_settings.json'
$Global:RecentFile   = Join-Path $Global:Root 'recent_devices.json'
$Global:CredsFile    = Join-Path $Global:Root 'user_creds.dat'
$Global:AuditFile    = Join-Path $Global:Root 'decommission-history.json'
$Global:AchievementsFile = Join-Path $Global:Root 'achievements.json'
$Global:DebugLogFile = Join-Path $env:TEMP 'DeviceDecommissioner_debug.log'

$Script:TIMER_INTERVAL_MS = 50
$Script:TOAST_DURATION_MS = 3500
$Script:LOG_MAX_LINES     = 500
$Script:CONFETTI_COUNT    = 60
$Script:CONFETTI_CLEANUP_MS = 5000

$Global:IsLightMode = $false
$Global:CachedBC    = New-Object System.Windows.Media.BrushConverter

# Synchronized job tracking (background runspaces)
$Global:BgJobs = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())

# Discovery state for current lookup
# All supported systems — used for discovery cards, iteration, and result tracking.
$Script:AllSystems = @('AD','Entra','Intune','Autopilot','SCCM')

$Global:LookupState = @{
    DeviceQuery = ''
    Results = @{
        AD        = $null   # @{ Found=$bool; DisplayName=''; Detail=''; Error='' }
        Entra     = $null
        Intune    = $null
        Autopilot = $null
        SCCM      = $null
    }
}

$Global:ModalCallback = $null   # ScriptBlock invoked on confirm

# ============================================================================
# THEME PALETTES
# ============================================================================
$Global:ThemeDark = @{
    ThemeAppBg='#111111';    ThemePanelBg='#1A1A1A';   ThemeCardBg='#222222';   ThemeCardAltBg='#252528'
    ThemeInputBg='#141414';  ThemeDeepBg='#0D0D0D';    ThemeOutputBg='#080808'; ThemeSurfaceBg='#2A2A2A'
    ThemeHoverBg='#333333';  ThemeSelectedBg='#2E2E32';ThemePressedBg='#181818'
    ThemeAccent='#0078D4';   ThemeAccentHover='#1A8AD4';ThemeAccentLight='#60CDFF';ThemeAccentDim='#004578'
    ThemeGreenAccent='#00C853';ThemeAccentText='#FFFFFF'
    ThemeTextPrimary='#FFFFFF';ThemeTextBody='#E0E0E0';ThemeTextSecondary='#CCCCCC';ThemeTextMuted='#AAAAAA';ThemeTextDim='#888888';ThemeTextDisabled='#666666';ThemeTextFaintest='#4A4A4A'
   
    ThemeBorder='#2A2A2A';   ThemeBorderCard='#333333';ThemeBorderElevated='#3E3E3E';ThemeBorderHover='#505050';ThemeBorderMedium='#1AFFFFFF'
    ThemeScrollThumb='#444444';ThemeScrollTrack='#22FFFFFF'
    ThemeSuccess='#00C853';  ThemeWarning='#FFB900';   ThemeError='#D13438'
    ThemeSidebarBg='#0D0D0D';ThemeSidebarBorder='#222222'
    DotColorBrush='#22FFFFFF';GlowColorBrush='#10005A9E'
}
$Global:ThemeLight = @{
    ThemeAppBg='#F5F5F5';    ThemePanelBg='#FAFAFA';   ThemeCardBg='#FFFFFF';   ThemeCardAltBg='#F8F8F8'
    ThemeInputBg='#FFFFFF';  ThemeDeepBg='#EEEEEE';    ThemeOutputBg='#F0F0F0'; ThemeSurfaceBg='#F2F2F5'
    ThemeHoverBg='#E8E8EC';  ThemeSelectedBg='#E0E0E4';ThemePressedBg='#D8D8DC'
    ThemeAccent='#0078D4';   ThemeAccentHover='#106EBE';ThemeAccentLight='#0063B1';ThemeAccentDim='#CCE4F6'
    ThemeGreenAccent='#107C10';ThemeAccentText='#FFFFFF'
    ThemeTextPrimary='#1A1A1A';ThemeTextBody='#333333';ThemeTextSecondary='#555555';ThemeTextMuted='#6B6B6B';ThemeTextDim='#808080';ThemeTextDisabled='#999999';ThemeTextFaintest='#AAAAAA'
    ThemeBorder='#E0E0E0';   ThemeBorderCard='#D0D0D0';ThemeBorderElevated='#CCCCCC';ThemeBorderHover='#AAAAAA';ThemeBorderMedium='#1A000000'
    ThemeScrollThumb='#888888';ThemeScrollTrack='#18000000'
    ThemeSuccess='#107C10';  ThemeWarning='#D48300';   ThemeError='#D13438'
    ThemeSidebarBg='#EEEEEE';ThemeSidebarBorder='#D0D0D0'
    DotColorBrush='#1A000000';GlowColorBrush='#080078D4'
}

# ============================================================================
# LOG FILE BOOTSTRAP
# ============================================================================
try {
    if ((Test-Path $Global:DebugLogFile) -and (Get-Item $Global:DebugLogFile).Length -gt 2MB) {
        $prev = $Global:DebugLogFile + '.prev'
        if (Test-Path $prev) { Remove-Item $prev -Force -ErrorAction SilentlyContinue }
        Rename-Item $Global:DebugLogFile $prev -Force -ErrorAction SilentlyContinue
    }
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    [System.IO.File]::AppendAllText($Global:DebugLogFile, "`n=== Device Decommissioner v$($Global:AppVersion) - $ts ===`r`n")
} catch { }

# ============================================================================
# BACKGROUND WORK (runspace + 50ms polling)
# ============================================================================
function Start-BackgroundWork {
    <#
    .SYNOPSIS
        Runs a scriptblock in a fresh STA runspace and queues it for the
        50ms poll timer to drain results back onto the UI thread.
    .PARAMETER Work
        The scriptblock to run on the background thread. Receives variables
        listed in -Variables. Should NOT touch any WPF/UI element directly.
    .PARAMETER OnComplete
        Scriptblock invoked on the UI thread once the runspace finishes.
        Receives ($Results, $Errors). Use this to update UI state.
    .PARAMETER Variables
        Hashtable of name -> value pairs to inject into the runspace's
        session state. The work block accesses them by name (no prefix).
    .NOTES
        Multiple background jobs can run in parallel — each gets its own
        runspace. Status is tracked in $Global:BgJobs. UI updates inside
        the work block will silently no-op (different thread + scope).
    #>
    param(
        [Parameter(Mandatory)][ScriptBlock]$Work,
        [Parameter(Mandatory)][ScriptBlock]$OnComplete,
        [hashtable]$Variables = @{}
    )
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $iss.ExecutionPolicy = [Microsoft.PowerShell.ExecutionPolicy]::Bypass
    $rs  = [RunspaceFactory]::CreateRunspace($iss)
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions  = 'ReuseThread'
    $rs.Open()
    $ps = [PowerShell]::Create()
    $ps.Runspace = $rs
    foreach ($k in $Variables.Keys) {
        $ps.Runspace.SessionStateProxy.SetVariable($k, $Variables[$k])
    }
    $null = $ps.AddScript($Work)
    $async = $ps.BeginInvoke()
    [void]$Global:BgJobs.Add(@{
        PS=$ps; Runspace=$rs; AsyncResult=$async; OnComplete=$OnComplete
        StartedAt=(Get-Date); Label=($Variables.Keys -join ',')
    })
    Write-DebugLog "BgWork: queued runspace #$($Global:BgJobs.Count) (vars=$($Variables.Keys -join ','))" -Level 'DEBUG'
}

# ============================================================================
# XAML LOAD
# ============================================================================
$xamlPath = Join-Path $Global:Root 'DeviceDecommissioner_UI.xaml'
if (-not (Test-Path $xamlPath)) {
    [System.Windows.MessageBox]::Show("DeviceDecommissioner_UI.xaml not found in:`n$Global:Root", 'Device Decommissioner', 'OK', 'Error') | Out-Null
    exit 1
}
$xamlContent = [IO.File]::ReadAllText($xamlPath)
try {
    $Window = [Windows.Markup.XamlReader]::Parse($xamlContent)
} catch {
    $msg = $_.Exception.Message
    if ($_.Exception.InnerException) { $msg += "`n`nInner: " + $_.Exception.InnerException.Message }
    [System.Windows.MessageBox]::Show("XAML parse error:`n$msg", 'Device Decommissioner', 'OK', 'Error') | Out-Null
    exit 1
}

# ============================================================================
# ELEMENT EXTRACTION
# ============================================================================
$names = @(
    'lblTitle','lblTitleVersion','lblVersion','lblStatus','lblStatusDetail','statusDot'
    'btnHelp','btnSettings','btnThemeToggle','btnMinimize','btnMaximize','btnClose'
    # Toolbar (Entra auth)
    'authDot','lblAuthStatus','bdrTenantInfo','lblTenantInfo','btnEntraSignIn','btnEntraSignOut'
    # Sign-in callout (shown above device card when no MgContext)
    'pnlEntraCallout','btnEntraSignInCallout'
    'txtDeviceId','btnLookup','btnCancelLookup','btnClearLookup','btnDecommission','btnRefresh','lblActionHint','btnEditCredsAD','btnEditCredsSCCM','chkDryRun','lblDecommissionIcon','lblDecommissionLabel'
    # Discovery cards
    'chkTargetAD','pillAD','dotAD','lblADStatus','lblADDetail','btnCopyAD','prgAD'
    'chkTargetEntra','pillEntra','dotEntra','lblEntraStatus','lblEntraDetail','btnCopyEntra','prgEntra'
    'chkTargetIntune','pillIntune','dotIntune','lblIntuneStatus','lblIntuneDetail','btnCopyIntune','prgIntune'
    'chkTargetSCCM','pillSCCM','dotSCCM','lblSCCMStatus','lblSCCMDetail','btnCopySCCM','prgSCCM'
    'chkTargetAutopilot','pillAutopilot','dotAutopilot','lblAutopilotStatus','lblAutopilotDetail','btnCopyAutopilot','prgAutopilot'
    # Cred indicators
    'dotCredAD','lblCredAD','dotCredSCCM','lblCredSCCM'
    # Modal
    'pnlModalOverlay','lblModalIcon','lblModalTitle','lblModalMessage','txtModalInput','pwdModalInput','btnModalCancel','btnModalConfirm'
    'pnlModalDetails','pnlModalDetailsScroll','pnlModalWarnings','lblModalInputPrompt','pnlModalCommands'
    # Settings
    # Settings (in-app tabbed view)
    'svMainContent','pnlSettingsView','btnSettingsBack'
    'setTabGeneral','setTabAD','setTabEntra','setTabIntune','setTabSCCM','setTabAppearance'
    'setPaneGeneral','setPaneAD','setPaneEntra','setPaneIntune','setPaneSCCM','setPaneAppearance'
    'lblSettingsFilePath','lblRecentFilePath','lblAuditFilePath'
    'txtSetADServer','txtSetADSearchBase','chkSetADEnabled'
    'txtSetEntraTenant','chkSetEntraEnabled','chkSetIntuneEnabled','chkSetAutopilotEnabled'
    'txtSetSCCMServer','txtSetSCCMSite','chkSetSCCMEnabled'
    'rbThemeDark','rbThemeLight','lblLogFilePath','btnSettingsCancel','btnSettingsSave'
    'txtSetRecentDays','btnExportAudit'
    # Log + progress + toasts
    'rtbLog','logParagraph','btnClearLog','btnCopyLog','lblDeviceIdHint'
    'pnlGlobalProgress','lblGlobalProgress','pnlToastHost','rowProgress'
    'pnlPrereqs','lblPrereqsDetail','btnInstallGraph','btnPrereqHelp'
    # Phase 4: rail + sidebar
    'colSidebar','pnlSidebar','btnRailHamburger','btnRailDevice','btnRailSettings','btnRailHelp','btnRailBulk','btnRailHistory'
    # Bulk lookup view
    'pnlBulkView','btnBulkBack','txtBulkInput','btnBulkRun','btnBulkCopy','dgBulkResults','pnlBulkEmpty','lblBulkInputHint'
    # History view
    'pnlHistoryView','btnHistoryBack','btnHistoryRefresh','txtHistoryFilter','chkHistoryHideDryRun','btnHistoryLookup','dgHistory','lblHistoryDetail','pnlHistoryEmpty','lblHistoryFilterHint'
    # Achievements view + canvas
    'btnRailAchievements','pnlAchievementsView','btnAchievementsBack','pnlAchievementsGrid','lblAchievementsCount','cnvConfetti'
    # Output panel chrome
    'rowSplitter','rowLog','btnToggleLog','btnRestoreLog'
    'lstRecentDevices','lblRecentEmpty','btnClearRecent'
)
foreach ($n in $names) { Set-Variable -Name $n -Value $Window.FindName($n) -Scope Script }
$paraLog = $logParagraph

# Pre-build a per-system card element cache so Set-CardStatus doesn't pay the
# Get-Variable cost on every UI update (called several times per lookup result).
# Each entry holds the live WPF references for one of the 5 discovery cards.
$Script:CardElements = @{}
foreach ($sys in $Script:AllSystems) {
    $Script:CardElements[$sys] = @{
        Dot     = Get-Variable -Name "dot$sys"           -ValueOnly -Scope Script -ErrorAction SilentlyContinue
        Label   = Get-Variable -Name "lbl${sys}Status"   -ValueOnly -Scope Script -ErrorAction SilentlyContinue
        Detail  = Get-Variable -Name "lbl${sys}Detail"   -ValueOnly -Scope Script -ErrorAction SilentlyContinue
        Copy    = Get-Variable -Name "btnCopy$sys"       -ValueOnly -Scope Script -ErrorAction SilentlyContinue
        Progress= Get-Variable -Name "prg$sys"           -ValueOnly -Scope Script -ErrorAction SilentlyContinue
    }
}

$lblTitle.Text         = $Global:AppTitle
$lblTitleVersion.Text  = "v$Global:AppVersion"
$lblVersion.Text       = "v$Global:AppVersion"
$lblLogFilePath.Text       = $Global:DebugLogFile
$lblSettingsFilePath.Text  = $Global:SettingsFile
$lblRecentFilePath.Text    = $Global:RecentFile
if ($lblAuditFilePath) { $lblAuditFilePath.Text = $Global:AuditFile }
# Window icon (also set in XAML; runtime fallback handles relative-path edge cases)
try {
    $iconPath = Join-Path $Global:Root 'DeviceDecommissioner.ico'
    if (Test-Path $iconPath) {
        $Window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]$iconPath)
    }
} catch {}

# ============================================================================
# THEME APPLY
# ============================================================================
$ApplyTheme = {
    param([bool]$IsLight)
    $Palette = if ($IsLight) { $Global:ThemeLight } else { $Global:ThemeDark }
    foreach ($Key in $Palette.Keys) {
        $NewColor = [System.Windows.Media.ColorConverter]::ConvertFromString($Palette[$Key])
        $NewBrush = [System.Windows.Media.SolidColorBrush]::new($NewColor)
        $Window.Resources[$Key] = $NewBrush
    }
    $Global:IsLightMode = $IsLight
    if ($btnThemeToggle) {
        $btnThemeToggle.Content = if ($IsLight) { [string][char]0xE708 } else { [string][char]0xE706 }
    }
}

# ============================================================================
# LOGGING
# ============================================================================
$Global:DebugLineCount = 0
$Global:DebugMaxLines  = $Script:LOG_MAX_LINES

function Write-DebugLog {
    param([string]$Message, [string]$Level = 'INFO')
    $ts   = (Get-Date).ToString('HH:mm:ss.fff')
    $line = "[$ts] [$Level] $Message"
    Write-Host $line -ForegroundColor DarkGray
    try { [IO.File]::AppendAllText($Global:DebugLogFile, $line + "`r`n") } catch { }
    if ($Level -eq 'DEBUG' -and -not $Global:DebugOverlayEnabled) { return }
    if (-not $paraLog) { return }
    $color = if ($Global:IsLightMode) {
        switch ($Level) { 'ERROR'{'#CC0000'} 'WARN'{'#B86E00'} 'SUCCESS'{'#008A2E'} 'DEBUG'{'#888888'} default {'#444444'} }
    } else {
        switch ($Level) { 'ERROR'{'#FF4040'} 'WARN'{'#FF9100'} 'SUCCESS'{'#16C60C'} 'DEBUG'{'#666666'} default {'#888888'} }
    }
    if ($paraLog.Inlines.Count -gt 0) {
        $paraLog.Inlines.Add((New-Object System.Windows.Documents.LineBreak))
    }
    $run = New-Object System.Windows.Documents.Run($line)
    $run.Foreground = $Global:CachedBC.ConvertFromString($color)
    $paraLog.Inlines.Add($run)
    $Global:DebugLineCount++
    if ($Global:DebugLineCount -gt $Global:DebugMaxLines -and $paraLog.Inlines.Count -ge 2) {
        $paraLog.Inlines.Remove($paraLog.Inlines.FirstInline)
        $paraLog.Inlines.Remove($paraLog.Inlines.FirstInline)
        $Global:DebugLineCount--
    }
    $rtbLog.ScrollToEnd()
}

# ============================================================================
# GLOBAL PROGRESS
# ============================================================================
function Show-GlobalProgress {
    param([string]$Text='')
    $pnlGlobalProgress.Visibility = 'Visible'
    $lblGlobalProgress.Text = $Text
    if ($rowProgress) { $rowProgress.Height = New-Object System.Windows.GridLength 22 }
}
function Hide-GlobalProgress {
    $pnlGlobalProgress.Visibility = 'Collapsed'
    $lblGlobalProgress.Text = ''
    if ($rowProgress) { $rowProgress.Height = New-Object System.Windows.GridLength 0 }
}

# Restore Look up button + hide Cancel after a lookup ends (success, timeout, cancel).
function Set-LookupButtonsIdle {
    if ($btnLookup) { $btnLookup.IsEnabled = $true; $btnLookup.Visibility = 'Visible' }
    if ($btnCancelLookup) { $btnCancelLookup.Visibility = 'Collapsed' }
}
function Set-Status { param([string]$Text,[string]$Detail='',[ValidateSet('Idle','Busy','Success','Warning','Error')][string]$State='Idle')
    $lblStatus.Text = $Text
    $lblStatusDetail.Text = $Detail
    $color = switch ($State) {
        'Busy'    { '#FFB900' }
        'Success' { '#00C853' }
        'Warning' { '#FFB900' }
        'Error'   { '#D13438' }
        default   { '#888888' }
    }
    $statusDot.Fill = $Global:CachedBC.ConvertFromString($color)
}

# ============================================================================
# TOAST
# ============================================================================
function Show-Toast {
    <#
    .SYNOPSIS
        Slides a colored notification banner into the top-right toast host.
        Auto-dismisses after $DurationMs (default 3500) with a 200ms fade.
    .PARAMETER Type
        Success/Error/Warning/Info — controls border color and icon glyph.
    #>
    param([string]$Message,[ValidateSet('Success','Error','Warning','Info')][string]$Type='Info',[int]$DurationMs=$Script:TOAST_DURATION_MS)
    $colors = if ($Global:IsLightMode) {
        @{ Success=@{Bg='#E6F4EA';Border='#107C10';Icon=[char]0xE73E}
           Error  =@{Bg='#FDECEC';Border='#D13438';Icon=[char]0xEA39}
           Warning=@{Bg='#FFF4CE';Border='#D48300';Icon=[char]0xE7BA}
           Info   =@{Bg='#E6F0FA';Border='#0078D4';Icon=[char]0xE946} }
    } else {
        @{ Success=@{Bg='#0A3D1A';Border='#00C853';Icon=[char]0xE73E}
           Error  =@{Bg='#3D0A0A';Border='#D13438';Icon=[char]0xEA39}
           Warning=@{Bg='#3D2D0A';Border='#FFB900';Icon=[char]0xE7BA}
           Info   =@{Bg='#0A1E3D';Border='#0078D4';Icon=[char]0xE946} }
    }
    $c = $colors[$Type]
    $toast = New-Object System.Windows.Controls.Border
    $toast.Background      = $Global:CachedBC.ConvertFromString($c.Bg)
    $toast.BorderBrush     = $Global:CachedBC.ConvertFromString($c.Border)
    $toast.BorderThickness = [System.Windows.Thickness]::new(1)
    $toast.CornerRadius    = [System.Windows.CornerRadius]::new(12)
    $toast.Padding         = [System.Windows.Thickness]::new(12,8,12,8)
    $toast.Margin          = [System.Windows.Thickness]::new(0,0,0,6)
    $toast.Opacity         = 0
    $sp = New-Object System.Windows.Controls.StackPanel
    $sp.Orientation = 'Horizontal'
    $icon = New-Object System.Windows.Controls.TextBlock
    $icon.Text=$c.Icon; $icon.FontFamily=[System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
    $icon.FontSize=14; $icon.Foreground=$Global:CachedBC.ConvertFromString($c.Border); $icon.VerticalAlignment='Center'
    $icon.Margin=[System.Windows.Thickness]::new(0,0,8,0)
    $msg = New-Object System.Windows.Controls.TextBlock
    $msg.Text=$Message; $msg.FontSize=11; $msg.Foreground=$Window.Resources['ThemeTextPrimary']
    $msg.TextWrapping='Wrap'; $msg.MaxWidth=300; $msg.VerticalAlignment='Center'
    [void]$sp.Children.Add($icon); [void]$sp.Children.Add($msg)
    $toast.Child = $sp
    $pnlToastHost.Children.Insert(0, $toast)

    $fadeIn = New-Object System.Windows.Media.Animation.DoubleAnimation
    $fadeIn.From=0; $fadeIn.To=1; $fadeIn.Duration=[System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(200))
    $toast.BeginAnimation([System.Windows.UIElement]::OpacityProperty,$fadeIn)

    $dismiss = New-Object System.Windows.Threading.DispatcherTimer
    $dismiss.Interval = [TimeSpan]::FromMilliseconds($DurationMs)
    $remove  = New-Object System.Windows.Threading.DispatcherTimer
    $remove.Interval  = [TimeSpan]::FromMilliseconds(250)
    $tRef=$dismiss; $rRef=$remove; $toRef=$toast; $hRef=$pnlToastHost
    $remove.Add_Tick({ $rRef.Stop(); if($hRef -and $toRef){ $hRef.Children.Remove($toRef) } }.GetNewClosure())
    $dismiss.Add_Tick({
        try {
            $tRef.Stop()
            $fadeOut = New-Object System.Windows.Media.Animation.DoubleAnimation
            $fadeOut.From=1; $fadeOut.To=0
            $fadeOut.Duration=[System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(200))
            $toRef.BeginAnimation([System.Windows.UIElement]::OpacityProperty,$fadeOut)
            $rRef.Start()
        } catch {}
    }.GetNewClosure())
    $dismiss.Start()
    Write-DebugLog "Toast [$Type]: $Message" -Level 'DEBUG'
}

# ============================================================================
# MODAL
# ============================================================================
# Internal helper: reset all optional panels to a clean state. Each Show-*
# function calls this then enables only what it needs.
function Reset-ModalPanels {
    if ($pnlModalDetails)       { $pnlModalDetails.ItemsSource  = $null }
    if ($pnlModalWarnings)      { $pnlModalWarnings.ItemsSource = $null }
    if ($pnlModalCommands)      { $pnlModalCommands.ItemsSource = $null }
    if ($pnlModalDetailsScroll) { $pnlModalDetailsScroll.Visibility = 'Collapsed' }
    if ($lblModalInputPrompt)   { $lblModalInputPrompt.Visibility   = 'Collapsed' }
    $txtModalInput.Visibility = 'Collapsed'
    $pwdModalInput.Visibility = 'Collapsed'
}

# Internal helper: wire copy buttons inside the pnlModalCommands ItemsControl
# (because XAML code-behind isn't available in XamlReader.Parse'd UIs).
function Wire-ModalCommandCopyButtons {
    if (-not $pnlModalCommands) { return }
    # Defer until after the ItemsControl has rendered its containers.
    $pnlModalCommands.Dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Loaded, [action]{
        $buttons = @()
        $stack = New-Object System.Collections.Stack
        $stack.Push($pnlModalCommands)
        while ($stack.Count -gt 0) {
            $node = $stack.Pop()
            $count = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($node)
            for ($i = 0; $i -lt $count; $i++) {
                $child = [System.Windows.Media.VisualTreeHelper]::GetChild($node, $i)
                if ($child -is [System.Windows.Controls.Button] -and $child.Tag) { $buttons += $child }
                else { $stack.Push($child) }
            }
        }
        foreach ($b in $buttons) {
            $b.Add_Click({
                $cmd = [string]$this.Tag
                if ($cmd) {
                    [System.Windows.Clipboard]::SetText($cmd)
                    Show-Toast 'Command copied to clipboard' -Type Info -DurationMs 1500
                }
            })
        }
    })
}

function Show-ModalConfirm {
    param(
        [string]$Title,
        [string]$Message,
        [string]$ConfirmLabel = 'OK',
        [string]$Icon = ([string][char]0xE7BA),
        [switch]$HideCancel,
        [array]$DetailItems = @(),     # @(@{ System='AD'; DisplayName='X'; Fields=@(@{Label;Value}) }, ...)
        [string[]]$Warnings = @(),
        [array]$Commands = @(),        # @(@{ Label; Description; Command }, ...)
        [scriptblock]$OnConfirm
    )
    Reset-ModalPanels
    $lblModalIcon.Text     = $Icon
    $lblModalTitle.Text    = $Title
    $lblModalMessage.Text  = $Message
    $lblModalMessage.Visibility = if ($Message) { 'Visible' } else { 'Collapsed' }
    if ($DetailItems -and $DetailItems.Count -gt 0) {
        $pnlModalDetails.ItemsSource = $DetailItems
        $pnlModalDetailsScroll.Visibility = 'Visible'
    }
    if ($Warnings -and $Warnings.Count -gt 0) {
        $pnlModalWarnings.ItemsSource = $Warnings
        $pnlModalDetailsScroll.Visibility = 'Visible'
    }
    if ($Commands -and $Commands.Count -gt 0) {
        $pnlModalCommands.ItemsSource = $Commands
        $pnlModalDetailsScroll.Visibility = 'Visible'
        Wire-ModalCommandCopyButtons
    }
    $btnModalConfirm.Content   = $ConfirmLabel
    $btnModalCancel.Content    = 'Cancel'
    $btnModalCancel.Visibility = if ($HideCancel) { 'Collapsed' } else { 'Visible' }
    $Global:ModalCallback  = $OnConfirm
    $pnlModalOverlay.Visibility = 'Visible'
    [void]$btnModalConfirm.Focus()
    Write-DebugLog "[Modal] show confirm: '$Title' (details=$($DetailItems.Count) warnings=$($Warnings.Count) commands=$($Commands.Count))" -Level 'DEBUG'
}
function Show-ModalInput {
    param(
        [string]$Title,
        [string]$Message = '',
        [string]$DefaultText = '',
        [string]$ConfirmLabel = 'Save',
        [string]$Icon = ([string][char]0xE70F),
        [string]$InputPrompt = '',
        [switch]$AsPassword,
        [array]$DetailItems = @(),
        [string[]]$Warnings = @(),
        [scriptblock]$OnConfirm
    )
    Reset-ModalPanels
    $lblModalIcon.Text   = $Icon
    $lblModalTitle.Text  = $Title
    $lblModalMessage.Text = $Message
    $lblModalMessage.Visibility = if ($Message) { 'Visible' } else { 'Collapsed' }
    if ($DetailItems -and $DetailItems.Count -gt 0) {
        $pnlModalDetails.ItemsSource = $DetailItems
        $pnlModalDetailsScroll.Visibility = 'Visible'
    }
    if ($Warnings -and $Warnings.Count -gt 0) {
        $pnlModalWarnings.ItemsSource = $Warnings
        $pnlModalDetailsScroll.Visibility = 'Visible'
    }
    if ($InputPrompt -and $lblModalInputPrompt) {
        $lblModalInputPrompt.Text = $InputPrompt
        $lblModalInputPrompt.Visibility = 'Visible'
    }
    if ($AsPassword) {
        $pwdModalInput.Visibility = 'Visible'
        $pwdModalInput.Clear()
        [void]$pwdModalInput.Focus()
    } else {
        $txtModalInput.Visibility = 'Visible'
        $txtModalInput.Text = $DefaultText
        [void]$txtModalInput.Focus()
        $txtModalInput.SelectAll()
    }
    $btnModalConfirm.Content   = $ConfirmLabel
    $btnModalCancel.Content    = 'Cancel'
    $btnModalCancel.Visibility = 'Visible'
    $Global:ModalCallback = $OnConfirm
    $pnlModalOverlay.Visibility = 'Visible'
    Write-DebugLog "[Modal] show input: '$Title' (asPassword=$AsPassword details=$($DetailItems.Count))" -Level 'DEBUG'
}
function Hide-Modal {
    Write-DebugLog "[Modal] hide (callback was set: $([bool]$Global:ModalCallback))" -Level 'DEBUG'
    $pnlModalOverlay.Visibility = 'Collapsed'
    Reset-ModalPanels
    $Global:ModalCallback = $null
}

# ============================================================================
# SETTINGS PERSISTENCE
# ============================================================================
$Global:Settings = $null

function Get-DefaultSettings {
    @{
        Theme = 'Dark'
        AD     = @{ Server=''; SearchBase=''; Enabled=$true }
        Entra  = @{ TenantId=''; Enabled=$true }
        Intune = @{ Enabled=$true }
        Autopilot = @{ Enabled=$false }
        SCCM   = @{ Server=''; SiteCode=''; Enabled=$false }
        SidebarVisible  = $false
        LogPanelVisible = $false
        RecentActivityDays = 7
    }
}

# Recent devices live in their own file so settings can stay clean / version-controllable.
$Global:RecentDevices = @()
function Load-RecentDevices {
    $Global:RecentDevices = @()
    if (Test-Path $Global:RecentFile) {
        try {
            $raw = Get-Content $Global:RecentFile -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($raw) { $Global:RecentDevices = @($raw | Where-Object { $_ } | ForEach-Object { [string]$_ }) }
        } catch {
            Write-DebugLog "Recent devices load failed: $($_.Exception.Message)" -Level 'WARN'
        }
    }
}
function Save-RecentDevices {
    try {
        $payload = @($Global:RecentDevices)
        if ($payload.Count -eq 0) {
            if (Test-Path $Global:RecentFile) { Remove-Item $Global:RecentFile -Force -ErrorAction SilentlyContinue }
            return
        }
        # Force JSON array (single-element edge case)
        ($payload | ConvertTo-Json -Depth 2) | Set-Content -Path $Global:RecentFile -Encoding UTF8
    } catch {
        Write-DebugLog "Recent devices save failed: $($_.Exception.Message)" -Level 'WARN'
    }
}

function Load-Settings {
    if (Test-Path $Global:SettingsFile) {
        try {
            $raw = Get-Content $Global:SettingsFile -Raw -Encoding UTF8 | ConvertFrom-Json
            # Convert PSCustomObject -> nested hashtable
            $Global:Settings = Get-DefaultSettings
            foreach ($section in $Script:AllSystems) {
                if ($raw.$section) {
                    foreach ($prop in $raw.$section.PSObject.Properties) {
                        $Global:Settings[$section][$prop.Name] = $prop.Value
                    }
                }
            }
            if ($raw.Theme) { $Global:Settings.Theme = $raw.Theme }
            if ($null -ne $raw.SidebarVisible)  { $Global:Settings.SidebarVisible  = [bool]$raw.SidebarVisible }
            if ($null -ne $raw.LogPanelVisible) { $Global:Settings.LogPanelVisible = [bool]$raw.LogPanelVisible }
            if ($null -ne $raw.RecentActivityDays) { $Global:Settings.RecentActivityDays = [int]$raw.RecentActivityDays }
            # Migrate legacy: RecentDevices used to live in user_settings.json
            if ($raw.RecentDevices -and -not (Test-Path $Global:RecentFile)) {
                $Global:RecentDevices = @($raw.RecentDevices | Where-Object { $_ } | ForEach-Object { [string]$_ })
                Save-RecentDevices
                Write-DebugLog 'Migrated RecentDevices from settings -> recent_devices.json' -Level 'DEBUG'
            }
            Write-DebugLog "Settings loaded from $Global:SettingsFile"
        } catch {
            Write-DebugLog "Settings load failed: $($_.Exception.Message)" -Level 'WARN'
            $Global:Settings = Get-DefaultSettings
        }
    } else {
        $Global:Settings = Get-DefaultSettings
        Write-DebugLog "Settings: using defaults (no file yet)" -Level 'DEBUG'
    }
}

function Save-Settings {
    try {
        $Global:Settings | ConvertTo-Json -Depth 5 | Set-Content -Path $Global:SettingsFile -Encoding UTF8
        Write-DebugLog "Settings saved to $Global:SettingsFile" -Level 'SUCCESS'
    } catch {
        Write-DebugLog "Settings save failed: $($_.Exception.Message)" -Level 'ERROR'
    }
}

# ============================================================================
# CREDENTIAL STORE (DPAPI per-user)
# ============================================================================
# Schema:
# {
#   "AD":   { "User": "CONTOSO\\svc_ad",   "Pass": "<DPAPI-encrypted>" },
#   "SCCM": { "User": "CONTOSO\\svc_sccm", "Pass": "<DPAPI-encrypted>" }
# }
# Pass is ConvertFrom-SecureString output (DPAPI per-user, Windows-only).
$Global:Creds = @{ AD=$null; SCCM=$null }

function Load-Creds {
    if (-not (Test-Path $Global:CredsFile)) { return }
    try {
        $raw = Get-Content $Global:CredsFile -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($key in @('AD','SCCM')) {
            if ($raw.$key -and $raw.$key.User -and $raw.$key.Pass) {
                $sec = ConvertTo-SecureString -String $raw.$key.Pass -ErrorAction Stop
                $Global:Creds[$key] = New-Object System.Management.Automation.PSCredential ($raw.$key.User, $sec)
            }
        }
        Write-DebugLog "Stored credentials loaded (AD=$([bool]$Global:Creds.AD), SCCM=$([bool]$Global:Creds.SCCM))" -Level 'DEBUG'
    } catch {
        Write-DebugLog "Failed to decrypt stored credentials: $($_.Exception.Message)" -Level 'WARN'
        $Global:Creds = @{ AD=$null; SCCM=$null }
    }
}

function Save-Creds {
    try {
        $obj = @{}
        foreach ($key in @('AD','SCCM')) {
            $c = $Global:Creds[$key]
            if ($c) {
                $obj[$key] = @{
                    User = $c.UserName
                    Pass = ConvertFrom-SecureString -SecureString $c.Password
                }
            }
        }
        $obj | ConvertTo-Json -Depth 4 | Set-Content -Path $Global:CredsFile -Encoding UTF8
        # File can only be decrypted by current Windows user (DPAPI).
        Write-DebugLog "Stored credentials saved (encrypted with DPAPI for current user)" -Level 'SUCCESS'
    } catch {
        Write-DebugLog "Failed to save credentials: $($_.Exception.Message)" -Level 'ERROR'
    }
}

function Update-CredIndicators {
    if ($Global:Creds.AD) {
        $dotCredAD.Fill   = $Window.Resources['ThemeSuccess']
        $lblCredAD.Text   = $Global:Creds.AD.UserName
    } else {
        $dotCredAD.Fill   = $Window.Resources['ThemeWarning']
        $lblCredAD.Text   = 'Not configured'
    }
    if ($Global:Creds.SCCM) {
        $dotCredSCCM.Fill = $Window.Resources['ThemeSuccess']
        $lblCredSCCM.Text = $Global:Creds.SCCM.UserName
    } else {
        $dotCredSCCM.Fill = $Window.Resources['ThemeWarning']
        $lblCredSCCM.Text = 'Not configured'
    }
}

# Two-step modal: prompt for username, then password, then save.
function Edit-StoredCreds {
    param([ValidateSet('AD','SCCM')][string]$Target)
    Write-DebugLog "[Creds] Edit-StoredCreds ENTER target=$Target  existing=$([bool]$Global:Creds[$Target])" -Level 'DEBUG'
    $existing = $Global:Creds[$Target]
    $defUser = if ($existing) { $existing.UserName } else { '' }
    Show-ModalInput -Title "$Target credentials - username" -Message "Enter username for $Target operations (e.g. CONTOSO\svc_account)." `
                    -DefaultText $defUser -ConfirmLabel 'Next' -OnConfirm {
        $u = $txtModalInput.Text.Trim()
        Hide-Modal
        if (-not $u) { return }
        Show-ModalInput -Title "$Target credentials - password" -Message "Enter password for $u." `
                        -ConfirmLabel 'Save' -AsPassword -OnConfirm {
            $sec = $pwdModalInput.SecurePassword
            Hide-Modal
            if (-not $sec -or $sec.Length -eq 0) {
                Show-Toast "Password empty, $Target credentials not changed" -Type Warning
                return
            }
            $Global:Creds[$Target] = New-Object System.Management.Automation.PSCredential ($u, $sec)
            Save-Creds
            Update-CredIndicators
            Show-Toast "$Target credentials saved" -Type Success
        }.GetNewClosure()
    }.GetNewClosure()
}

# ============================================================================
# UI HELPERS — discovery card status
# ============================================================================
function Set-CardStatus {
    <#
    .SYNOPSIS
        Updates one discovery card's dot color, status label, detail text, copy
        button visibility, and progress bar based on the lookup state.
    .NOTES
        Element references are pulled from $Script:CardElements (cached at
        startup) instead of repeated Get-Variable calls, which makes this hot
        path roughly 10x faster.
    #>
    param(
        [ValidateSet('AD','Entra','Intune','Autopilot','SCCM')][string]$System,
        [ValidateSet('Idle','Busy','Found','NotFound','Removed','WouldRemove','Skipped','Error','ModuleMissing','Warning','Cancelled')][string]$State,
        [string]$Detail = ''
    )
    $card = $Script:CardElements[$System]
    if (-not $card) { Write-DebugLog "Set-CardStatus: no card cache for '$System'" -Level 'WARN'; return }
    $dot = $card.Dot; $lbl = $card.Label; $det = $card.Detail
    if (-not $dot -or -not $lbl -or -not $det) {
        Write-DebugLog "Set-CardStatus: missing UI element(s) for '$System' (XAML out of sync?)" -Level 'WARN'
        return
    }
    switch ($State) {
        'Idle'           { $dot.Fill = $Window.Resources['ThemeTextDim'];     $lbl.Text = 'Idle' }
        'Busy'           { $dot.Fill = $Window.Resources['ThemeWarning'];     $lbl.Text = 'Searching...' }
        'Found'          { $dot.Fill = $Window.Resources['ThemeSuccess'];     $lbl.Text = 'Found' }
        'NotFound'       { $dot.Fill = $Window.Resources['ThemeTextDim'];     $lbl.Text = 'Not found' }
        'Removed'        { $dot.Fill = $Window.Resources['ThemeAccent'];      $lbl.Text = 'Removed' }
        'WouldRemove'    { $dot.Fill = $Window.Resources['ThemeAccentLight']; $lbl.Text = 'Would remove' }
        'Skipped'        { $dot.Fill = $Window.Resources['ThemeTextDim'];     $lbl.Text = 'Skipped' }
        'Error'          { $dot.Fill = $Window.Resources['ThemeError'];       $lbl.Text = 'Error' }
        'ModuleMissing'  { $dot.Fill = $Window.Resources['ThemeWarning'];     $lbl.Text = 'Module missing' }
        'Warning'        { $dot.Fill = $Window.Resources['ThemeWarning'];     $lbl.Text = 'Warning' }
        'Cancelled'      { $dot.Fill = $Window.Resources['ThemeTextDim'];     $lbl.Text = 'Cancelled' }
    }
    $det.Text = $Detail
    if ($card.Copy) {
        $card.Copy.Visibility = if ($State -in @('Found','Removed','WouldRemove','Error','NotFound','ModuleMissing','Warning') -and $Detail) { 'Visible' } else { 'Collapsed' }
    }
    if ($card.Progress) {
        $card.Progress.Visibility = if ($State -eq 'Busy') { 'Visible' } else { 'Collapsed' }
    }
}

function Reset-AllCards {
    <#
    .SYNOPSIS
        Resets all five discovery cards to Idle state and clears the lookup
        results hashtable. Disables the Decommission button and resets the
        action hint label.
    #>
    foreach ($s in $Script:AllSystems) {
        Set-CardStatus -System $s -State Idle
        $card = $Script:CardElements[$s]
        if ($card -and $card.Copy) { $card.Copy.Visibility = 'Collapsed' }
    }
    foreach ($k in $Script:AllSystems) { $Global:LookupState.Results[$k] = $null }
    $btnDecommission.IsEnabled = $false
    $lblActionHint.Text = 'Look up a device to enable decommissioning.'
}

function Update-DecommissionButton {
    $any = $false
    foreach ($s in $Script:AllSystems) {
        $r = $Global:LookupState.Results[$s]
        $chk = Get-Variable -Name "chkTarget$s" -ValueOnly -Scope Script
        if ($chk.IsChecked -and $r -and $r.Found) { $any = $true; break }
    }
    $btnDecommission.IsEnabled = $any
    $isDry = [bool]$chkDryRun.IsChecked
    if ($lblDecommissionLabel) {
        $lblDecommissionLabel.Text = if ($isDry) { 'Dry-run selected' } else { 'Decommission selected' }
    }
    if ($lblDecommissionIcon) {
        $lblDecommissionIcon.Text = if ($isDry) { [string][char]0xE7BA } else { [string][char]0xE74D }
    }
    $lblActionHint.Text = if (-not $any) { 'No matching device targets selected.' }
                          elseif ($isDry) { "Will validate '$($Global:LookupState.DeviceQuery)' end-to-end without removing anything." }
                          else            { "Ready to remove '$($Global:LookupState.DeviceQuery)' from selected systems." }
}

# ============================================================================
# DEVICE LOOKUP
# ============================================================================
function Start-DeviceLookup {
    <#
    .SYNOPSIS
        Kicks off parallel lookups across all 5 enabled systems for one
        device hostname / wildcard / GUID. Runs through Start-BackgroundWork.
    .DESCRIPTION
        Generates a fresh $myGen counter so any in-flight callbacks from a
        previous lookup are dropped (see Save-LookupResult). Honours per-
        system checkboxes (skipped systems are marked as Skipped instantly).
        A 30-second timeout timer fires if any runspace hangs.
    .NOTES
        AD lookup runs in its own runspace; Entra+Intune+Autopilot share
        one (single Graph context); SCCM has its own. The poll timer drains
        results into Save-LookupResult on the UI thread.
    #>
    param([string]$Query)
    Write-DebugLog "[Lookup] Starting lookup for '$Query'" -Level 'INFO'
    Write-DebugLog "[Lookup] >>> ENTER  query='$Query'  inFlight=$Global:LookupInFlight  gen=$Global:LookupGen" -Level 'DEBUG'
    if (-not $Query) { Write-DebugLog '[Lookup] aborted: empty query' -Level 'DEBUG'; Show-Toast 'Enter a device hostname or ObjectId' -Type Warning; return }
    if ($Global:LookupInFlight) { Write-DebugLog '[Lookup] aborted: another lookup already in flight' -Level 'DEBUG'; Show-Toast 'Lookup already in progress' -Type Warning; return }
    Add-RecentDevice -Name $Query
    Unlock-Achievement 'first_lookup'
    if ($Query -match '[\*\?]') { Unlock-Achievement 'wildcard_user' }
    $Global:LookupInFlight = $true
    $Global:LookupGen = [int]$Global:LookupGen + 1
    $myGen = $Global:LookupGen
    $btnLookup.IsEnabled = $false
    $btnLookup.Visibility       = 'Collapsed'
    if ($btnCancelLookup) { $btnCancelLookup.Visibility = 'Visible'; $btnCancelLookup.IsEnabled = $true }
    $Global:LookupState.DeviceQuery = $Query
    Reset-AllCards
    Set-Status -Text 'Looking up device' -Detail $Query -State 'Busy'
    Show-GlobalProgress -Text "Searching for '$Query' across configured directories..."

    # Honour per-system checkboxes — a system the user unchecked is not searched.
    $runAD        = [bool]$chkTargetAD.IsChecked
    $runEntra     = [bool]$chkTargetEntra.IsChecked
    $runIntune    = [bool]$chkTargetIntune.IsChecked
    $runAutopilot = if ($chkTargetAutopilot) { [bool]$chkTargetAutopilot.IsChecked } else { $false }
    $runSCCM      = [bool]$chkTargetSCCM.IsChecked
    Write-DebugLog "[Lookup] gen=$myGen  targets: AD=$runAD Entra=$runEntra Intune=$runIntune Autopilot=$runAutopilot SCCM=$runSCCM" -Level 'INFO'
    if (-not ($runAD -or $runEntra -or $runIntune -or $runAutopilot -or $runSCCM)) {
        Write-DebugLog '[Lookup] aborted: no systems checked' -Level 'DEBUG'
        Hide-GlobalProgress
        Set-Status -Text 'No targets selected' -State Warning
        Show-Toast 'No systems selected to search' -Type Warning
        $Global:LookupInFlight = $false; Set-LookupButtonsIdle
        return
    }

    # Safety timeout — if any background job hangs we never want LookupInFlight stuck.
    if ($Global:LookupTimeoutTimer) { try { $Global:LookupTimeoutTimer.Stop() } catch {} }
    $Global:LookupTimeoutTimer = New-Object System.Windows.Threading.DispatcherTimer
    $Global:LookupTimeoutTimer.Interval = [TimeSpan]::FromSeconds(30)
    $Global:LookupTimeoutTimer.Add_Tick({
        try { $Global:LookupTimeoutTimer.Stop() } catch {}
        if (-not $Global:LookupInFlight -or $Global:LookupGen -ne $myGen) { return }
        # Log which systems never returned — this is the key diagnostic for hangs.
        # NOTE: hardcode the system list here; $Script:AllSystems is not visible inside
        # DispatcherTimer Tick handlers (they run in a dynamic module scope).
        $pending = @()
        foreach ($s in @('AD','Entra','Intune','Autopilot','SCCM')) {
            if (-not $Global:LookupState.Results[$s]) {
                $pending += $s
                Set-CardStatus -System $s -State Error -Detail 'Timed out (30s). Background runspace did not complete.'
            }
        }
        Write-DebugLog "Lookup timed out after 30s. Still pending: $($pending -join ', '). BgJobs count=$($Global:BgJobs.Count)" -Level 'WARN'
        Hide-GlobalProgress
        Set-Status -Text 'Lookup timed out' -State Warning
        $Global:LookupInFlight = $false
        Set-LookupButtonsIdle
        Show-Toast "Lookup timed out — $($pending -join ', ') did not respond" -Type Warning
    }.GetNewClosure())
    $Global:LookupTimeoutTimer.Start()

    # Mark skipped systems immediately so the completion counter still hits 4.
    $skipResult = @{ Found=$false; DisplayName=''; Detail='Skipped (unchecked).'; Error=''; Raw=$null; Skipped=$true }
    if (-not $runAD)        { Save-LookupResult -System 'AD'        -Result $skipResult -Gen $myGen }
    if (-not $runEntra)      { Save-LookupResult -System 'Entra'     -Result $skipResult -Gen $myGen }
    if (-not $runIntune)     { Save-LookupResult -System 'Intune'    -Result $skipResult -Gen $myGen }
    if (-not $runAutopilot)  { Save-LookupResult -System 'Autopilot' -Result $skipResult -Gen $myGen }
    if (-not $runSCCM)       { Save-LookupResult -System 'SCCM'      -Result $skipResult -Gen $myGen }

    foreach ($s in $Script:AllSystems) {
        $shouldRun = switch ($s) { 'AD'{$runAD} 'Entra'{$runEntra} 'Intune'{$runIntune} 'Autopilot'{$runAutopilot} 'SCCM'{$runSCCM} }
        if ($shouldRun) { Set-CardStatus -System $s -State Busy }
    }

    $settings = $Global:Settings
    $credAD   = $Global:Creds.AD
    $credSCCM = $Global:Creds.SCCM

    # ---- AD ----
    if ($runAD) {
    Write-DebugLog '[Lookup] launching AD background runspace' -Level 'INFO'
    Start-BackgroundWork -Variables @{ Query=$Query; Cred=$credAD; Server=$settings.AD.Server; SearchBase=$settings.AD.SearchBase } -Work {
        $r = @{ Found=$false; DisplayName=''; Detail=''; Error=''; Raw=$null }
        try {
            if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) { $r.Error='ModuleMissing'; return $r }
            Import-Module ActiveDirectory -ErrorAction Stop
            $params = @{ ErrorAction='Stop' }
            if ($Server)     { $params.Server     = $Server }
            if ($SearchBase) { $params.SearchBase = $SearchBase }
            if ($Cred)       { $params.Credential = $Cred }
            # Escape embedded single quotes so the AD filter can't be broken/injected.
            $q = $Query.Replace("'", "''")
            if ($q -match '[\*\?]') {
                $params.Filter = "Name -like '$q' -or sAMAccountName -like '$q'"
            } else {
                $params.Filter = "Name -eq '$q' -or sAMAccountName -eq '$q' -or sAMAccountName -eq '$($q)$'"
            }
            # First query: core properties only (fast, no LAPS which can hang if schema is missing).
            $obj = Get-ADComputer @params -Properties DistinguishedName,LastLogonDate,Enabled,OperatingSystem,Description | Select-Object -First 1
            if ($obj) {
                $r.Found       = $true
                $r.DisplayName = $obj.Name
                $r.Detail      = $obj.DistinguishedName
                $r.Raw         = @{
                    DistinguishedName = $obj.DistinguishedName
                    SamAccountName    = $obj.SamAccountName
                    LastLogonDate     = if ($obj.LastLogonDate) { $obj.LastLogonDate.ToString('o') } else { $null }
                    Enabled           = [bool]$obj.Enabled
                    OperatingSystem   = $obj.OperatingSystem
                    Description       = $obj.Description
                    HasLAPSPassword   = $false
                    BitLockerKeyCount = 0
                }
                # LAPS check (separate query — can fail without blocking the main result)
                try {
                    $lapsParams = @{ Identity=$obj.DistinguishedName; Properties=@('ms-Mcs-AdmPwd','msLAPS-Password'); ErrorAction='Stop' }
                    if ($Server) { $lapsParams.Server = $Server }
                    if ($Cred)   { $lapsParams.Credential = $Cred }
                    $lapsObj = Get-ADComputer @lapsParams
                    if ($lapsObj.'msLAPS-Password' -or $lapsObj.'ms-Mcs-AdmPwd') { $r.Raw.HasLAPSPassword = $true }
                } catch { <# LAPS attrs may not exist in schema — that's fine #> }
                # BitLocker recovery keys (separate query)
                try {
                    $blkParams = @{ Filter="objectClass -eq 'msFVE-RecoveryInformation'"; SearchBase=$obj.DistinguishedName; ErrorAction='SilentlyContinue' }
                    if ($Server) { $blkParams.Server = $Server }
                    if ($Cred)   { $blkParams.Credential = $Cred }
                    $r.Raw.BitLockerKeyCount = @(Get-ADObject @blkParams).Count
                } catch { <# No recovery keys or no access — that's fine #> }
            } else { $r.Detail = 'No computer object matched.' }
        } catch { $r.Error = $_.Exception.Message }
        return $r
    } -OnComplete {
        param($Results,$Errors)
        $r = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ Found=$false; Error='Background job returned no result' } }
        Save-LookupResult -System 'AD' -Result $r -Gen $myGen
    }.GetNewClosure()
    }

    # ---- Entra ID + Intune (Microsoft Graph) ----
    if ($runEntra -or $runIntune -or $runAutopilot) {
    Write-DebugLog "[Lookup] launching Entra+Intune+Autopilot background runspace (entra=$runEntra intune=$runIntune autopilot=$runAutopilot)" -Level 'INFO'
    Start-BackgroundWork -Variables @{ Query=$Query; TenantId=$settings.Entra.TenantId; QueryEntra=$runEntra; QueryIntune=$runIntune; QueryAutopilot=$runAutopilot } -Work {
        $out = @{ Entra=@{Found=$false;DisplayName='';Detail='';Error='';Raw=$null}
                  Intune=@{Found=$false;DisplayName='';Detail='';Error='';Raw=$null}
                  Autopilot=@{Found=$false;DisplayName='';Detail='';Error='';Raw=$null} }
        try {
            if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Identity.DirectoryManagement)) {
                $out.Entra.Error='ModuleMissing'
            }
            if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.DeviceManagement)) {
                $out.Intune.Error='ModuleMissing'
            }
            if ($out.Entra.Error -eq 'ModuleMissing' -and $out.Intune.Error -eq 'ModuleMissing') { return $out }

            # Check for existing Graph context — do NOT attempt interactive auth from a background thread.
            # The user must sign in via the toolbar button first; that runs on the UI thread.
            $ctx = $null
            try { $ctx = Get-MgContext -ErrorAction SilentlyContinue } catch {}
            if (-not $ctx) {
                $errMsg = 'Not signed in to Microsoft Graph. Click the Sign in to Entra button first.'
                if ($QueryEntra  -and $out.Entra.Error  -ne 'ModuleMissing') { $out.Entra.Error  = $errMsg }
                if ($QueryIntune -and $out.Intune.Error -ne 'ModuleMissing') { $out.Intune.Error = $errMsg }
                return $out
            }

            # Entra device
            if ($QueryEntra -and $out.Entra.Error -ne 'ModuleMissing') {
                try {
                    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
                    $devs = @()
                    # Graph $filter only supports startsWith for prefix; map '*' wildcard.
                    if ($Query -match '\*$' -and ($Query -notmatch '^\*')) {
                        $prefix = ($Query.TrimEnd('*')).Replace("'","''")
                        try { $devs = @(Get-MgDevice -Filter "startswith(displayName,'$prefix')" -All -ErrorAction Stop) }
                        catch { $out.Entra.Detail = "Graph startswith() failed: $($_.Exception.Message)" }
                    } elseif ($Query -match '[\*\?]') {
                        # Other wildcards: pull a bounded set then filter client-side.
                        $like = $Query
                        try { $all = @(Get-MgDevice -All -ErrorAction Stop -Top 200) }
                        catch { $all = @(); $out.Entra.Detail = "Graph paged Get-MgDevice failed: $($_.Exception.Message)" }
                        $devs = @($all | Where-Object { $_.DisplayName -like $like })
                    } else {
                        # Try by displayName first, then by deviceId GUID
                        try { $devs = @(Get-MgDevice -Filter "displayName eq '$Query'" -All -ErrorAction Stop) }
                        catch { $out.Entra.Detail = "Graph displayName lookup failed: $($_.Exception.Message)" }
                        if ($devs.Count -eq 0 -and $Query -match '^[0-9a-fA-F\-]{36}$') {
                            try { $devs = @(Get-MgDevice -Filter "deviceId eq '$Query'" -All -ErrorAction Stop) }
                            catch { $out.Entra.Detail = "Graph deviceId lookup failed: $($_.Exception.Message)" }
                        }
                    }
                    if ($devs.Count -gt 0) {
                        $d = $devs[0]
                        $out.Entra.Found       = $true
                        $out.Entra.DisplayName = $d.DisplayName
                        $out.Entra.Detail      = "ObjectId=$($d.Id)  DeviceId=$($d.DeviceId)  OS=$($d.OperatingSystem)"
                        $out.Entra.Raw         = @{
                            Id            = $d.Id
                            DeviceId      = $d.DeviceId
                            DisplayName   = $d.DisplayName
                            OperatingSystem        = $d.OperatingSystem
                            OperatingSystemVersion = $d.OperatingSystemVersion
                            ApproximateLastSignInDateTime = if ($d.ApproximateLastSignInDateTime) { $d.ApproximateLastSignInDateTime.ToString('o') } else { $null }
                            AccountEnabled = [bool]$d.AccountEnabled
                        }
                        # Check BitLocker recovery keys escrowed in Entra
                        try {
                            $blkKeys = @(Get-MgDeviceRegisteredOwner -DeviceId $d.Id -ErrorAction SilentlyContinue)
                        } catch { $blkKeys = @() }
                        try {
                            $blkCount = @(Get-MgInformationProtectionBitlockerRecoveryKey -Filter "deviceId eq '$($d.DeviceId)'" -ErrorAction SilentlyContinue).Count
                            $out.Entra.Raw.BitLockerKeyCount = $blkCount
                        } catch { $out.Entra.Raw.BitLockerKeyCount = 0 }
                    } else {
                        $out.Entra.Detail = 'No Entra device matched.'
                    }
                } catch { $out.Entra.Error = $_.Exception.Message }
            }

            # Intune managed device
            if ($QueryIntune -and $out.Intune.Error -ne 'ModuleMissing') {
                try {
                    Import-Module Microsoft.Graph.DeviceManagement -ErrorAction Stop
                    $mds = @()
                    if ($Query -match '\*$' -and ($Query -notmatch '^\*')) {
                        $prefix = ($Query.TrimEnd('*')).Replace("'","''")
                        try { $mds = @(Get-MgDeviceManagementManagedDevice -Filter "startswith(deviceName,'$prefix')" -All -ErrorAction Stop) }
                        catch { $out.Intune.Detail = "Graph startswith() failed: $($_.Exception.Message)" }
                    } elseif ($Query -match '[\*\?]') {
                        try { $all = @(Get-MgDeviceManagementManagedDevice -All -ErrorAction Stop -Top 200) }
                        catch { $all = @(); $out.Intune.Detail = "Graph paged Get-MgDeviceManagementManagedDevice failed: $($_.Exception.Message)" }
                        $mds = @($all | Where-Object { $_.DeviceName -like $Query })
                    } else {
                        try { $mds = @(Get-MgDeviceManagementManagedDevice -Filter "deviceName eq '$Query'" -All -ErrorAction Stop) }
                        catch { $out.Intune.Detail = "Graph deviceName lookup failed: $($_.Exception.Message)" }
                    }
                    if ($mds.Count -gt 0) {
                        $m = $mds[0]
                        $out.Intune.Found       = $true
                        $out.Intune.DisplayName = $m.DeviceName
                        $out.Intune.Detail      = "Id=$($m.Id)  AzureAdDeviceId=$($m.AzureAdDeviceId)  LastSync=$($m.LastSyncDateTime)"
                        $out.Intune.Raw         = @{
                            Id                 = $m.Id
                            DeviceName         = $m.DeviceName
                            AzureAdDeviceId    = $m.AzureAdDeviceId
                            LastSyncDateTime   = if ($m.LastSyncDateTime) { $m.LastSyncDateTime.ToString('o') } else { $null }
                            OperatingSystem    = $m.OperatingSystem
                            UserPrincipalName  = $m.UserPrincipalName
                            ManagedDeviceOwnerType = [string]$m.ManagedDeviceOwnerType
                        }
                    } else {
                        $out.Intune.Detail = 'No Intune managed device matched.'
                    }
                } catch { $out.Intune.Error = $_.Exception.Message }
            }

            # Autopilot device identity
            if ($QueryAutopilot) {
                try {
                    Import-Module Microsoft.Graph.DeviceManagement.Enrollment -ErrorAction Stop
                    # Match by device name (contains search) — Autopilot API doesn't have a displayName filter,
                    # so we pull all and filter client-side (bounded to 200).
                    $apDevs = @()
                    try { $apDevs = @(Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -All -Top 200 -ErrorAction Stop) } catch {}
                    $match = $apDevs | Where-Object {
                        ($_.DisplayName -eq $Query) -or
                        ($_.DisplayName -like $Query) -or
                        ($_.SerialNumber -eq $Query)
                    } | Select-Object -First 1
                    if ($match) {
                        $out.Autopilot.Found       = $true
                        $out.Autopilot.DisplayName = $match.DisplayName
                        $out.Autopilot.Detail      = "Id=$($match.Id)  Serial=$($match.SerialNumber)  Model=$($match.Model)"
                        $out.Autopilot.Raw         = @{
                            Id           = $match.Id
                            DisplayName  = $match.DisplayName
                            SerialNumber = $match.SerialNumber
                            Model        = $match.Model
                            GroupTag     = $match.GroupTag
                        }
                    } else {
                        $out.Autopilot.Detail = 'No Autopilot device identity matched.'
                    }
                } catch { $out.Autopilot.Error = $_.Exception.Message }
            }
        } catch {
            if ($QueryEntra)    { $out.Entra.Error    = $_.Exception.Message }
            if ($QueryIntune)   { $out.Intune.Error   = $_.Exception.Message }
            if ($QueryAutopilot){ $out.Autopilot.Error= $_.Exception.Message }
        }
        return $out
    } -OnComplete {
        param($Results,$Errors)
        $r = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { $null }
        if (-not $r) {
            $errMsg = if ($Errors -and $Errors.Count -gt 0) { ($Errors[0] | Out-String).Trim() } else { 'Background job returned no result' }
            if ($runEntra)    { Save-LookupResult -System 'Entra'     -Result @{ Found=$false; Error=$errMsg; Detail=''; Raw=$null } -Gen $myGen }
            if ($runIntune)   { Save-LookupResult -System 'Intune'    -Result @{ Found=$false; Error=$errMsg; Detail=''; Raw=$null } -Gen $myGen }
            if ($runAutopilot){ Save-LookupResult -System 'Autopilot' -Result @{ Found=$false; Error=$errMsg; Detail=''; Raw=$null } -Gen $myGen }
            return
        }
        if ($runEntra)    { Save-LookupResult -System 'Entra'     -Result $r.Entra     -Gen $myGen }
        if ($runIntune)   { Save-LookupResult -System 'Intune'    -Result $r.Intune    -Gen $myGen }
        if ($runAutopilot){ Save-LookupResult -System 'Autopilot' -Result $r.Autopilot -Gen $myGen }
    }.GetNewClosure()
    }

    # ---- SCCM ----
    if ($runSCCM) {
    Write-DebugLog '[Lookup] launching SCCM background runspace' -Level 'INFO'
    Start-BackgroundWork -Variables @{ Query=$Query; Server=$settings.SCCM.Server; SiteCode=$settings.SCCM.SiteCode; Cred=$credSCCM } -Work {
        $r = @{ Found=$false; DisplayName=''; Detail=''; Error=''; Raw=$null }
        try {
            if (-not $Server -or -not $SiteCode) { $r.Error='Not configured'; return $r }
            $cmModule = Join-Path $env:SMS_ADMIN_UI_PATH '..\ConfigurationManager.psd1'
            if (-not (Test-Path $cmModule)) {
                $cmModule = "${env:ProgramFiles(x86)}\Microsoft Endpoint Manager\AdminConsole\bin\ConfigurationManager.psd1"
            }
            if (-not (Test-Path $cmModule)) { $r.Error='ModuleMissing'; return $r }
            Import-Module $cmModule -ErrorAction Stop
            $driveName = "$SiteCode" + ':'
            if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
                $np = @{ Name=$SiteCode; PSProvider='CMSite'; Root=$Server; ErrorAction='Stop' }
                if ($Cred) { $np.Credential = $Cred }
                New-PSDrive @np | Out-Null
            }
            Push-Location $driveName
            try {
                $obj = Get-CMDevice -Name $Query -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($obj) {
                    $r.Found       = $true
                    $r.DisplayName = $obj.Name
                    $r.Detail      = "ResourceID=$($obj.ResourceID)  DomainOrWorkgroup=$($obj.DomainOrWorkgroup)  Client=$($obj.IsClient)"
                    $r.Raw         = @{ ResourceID=$obj.ResourceID; Name=$obj.Name }
                } else { $r.Detail = 'No SCCM device matched.' }
            } finally { Pop-Location }
        } catch { $r.Error = $_.Exception.Message }
        return $r
    } -OnComplete {
        param($Results,$Errors)
        $r = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ Found=$false; Error='Background job returned no result' } }
        Save-LookupResult -System 'SCCM' -Result $r -Gen $myGen
    }.GetNewClosure()
    }
}

function Save-LookupResult {
    <#
    .SYNOPSIS
        UI-thread callback that processes one system's lookup result, updates
        the corresponding card, and fires the lookup-complete logic when all
        five systems have reported back.
    .PARAMETER Gen
        Generation counter from when the lookup started. If it doesn't match
        $Global:LookupGen the result is dropped (stale callback from a
        cancelled or replaced lookup).
    #>
    param([string]$System,[hashtable]$Result,[int]$Gen = -1)
    # Stale callback from a cancelled / replaced lookup — ignore.
    if ($Gen -ge 0 -and $Gen -ne $Global:LookupGen) {
        Write-DebugLog "[Lookup] ${System}: stale callback dropped (gen=$Gen, current=$Global:LookupGen)" -Level 'DEBUG'
        return
    }
    Write-DebugLog "[Lookup] ${System}: result received  found=$([bool]$Result.Found)  skipped=$([bool]$Result.Skipped)  error='$($Result.Error)'" -Level 'INFO'
    $Global:LookupState.Results[$System] = $Result
    if (-not $Result) {
        Set-CardStatus -System $System -State Error -Detail 'No result returned from background job.'
    } elseif ($Result.Skipped) {
        Set-CardStatus -System $System -State Skipped -Detail $Result.Detail
    } elseif ($Result.Error -eq 'ModuleMissing') {
        Set-CardStatus -System $System -State ModuleMissing -Detail 'Required PowerShell module is not installed on this machine.'
        Write-DebugLog "${System} lookup: required module not installed" -Level 'WARN'
    } elseif ($Result.Error) {
        Set-CardStatus -System $System -State Error -Detail $Result.Error
        Write-DebugLog "${System} lookup error: $($Result.Error)" -Level 'ERROR'
    } elseif ($Result.Found) {
        Set-CardStatus -System $System -State Found -Detail "$($Result.DisplayName) - $($Result.Detail)"
        Write-DebugLog "${System}: found '$($Result.DisplayName)'" -Level 'SUCCESS'
    } else {
        Set-CardStatus -System $System -State NotFound -Detail $Result.Detail
        Write-DebugLog "${System}: no match for '$($Global:LookupState.DeviceQuery)'"
    }

    # All four systems reported back?
    $done = ($Global:LookupState.Results.GetEnumerator() | Where-Object { $_.Value -ne $null }).Count
    if ($done -ge $Script:AllSystems.Count) {
        if ($Global:LookupTimeoutTimer) { try { $Global:LookupTimeoutTimer.Stop() } catch {} }
        Hide-GlobalProgress
        $found=0; $notfound=0; $errored=0; $skipped=0
        foreach ($v in $Global:LookupState.Results.Values) {
            if ($v.Skipped)        { $skipped++ }
            elseif ($v.Error)      { $errored++ }
            elseif ($v.Found)      { $found++ }
            else                   { $notfound++ }
        }
        $parts = @()
        if ($found)    { $parts += "$found found" }
        if ($notfound) { $parts += "$notfound not found" }
        if ($errored)  { $parts += "$errored error" }
        if ($skipped)  { $parts += "$skipped skipped" }
        $summary = if ($parts.Count) { ($parts -join ', ') } else { 'no results' }
        $state = if ($errored -gt 0 -and $found -eq 0) { 'Error' } elseif ($errored -gt 0) { 'Warning' } elseif ($found -eq 0) { 'Warning' } else { 'Success' }
        Set-Status -Text 'Lookup complete' -Detail "$($Global:LookupState.DeviceQuery) - $summary" -State $state
        Write-DebugLog "Lookup complete: $summary" -Level $(if ($state -eq 'Error') { 'ERROR' } elseif ($state -eq 'Warning') { 'WARN' } else { 'SUCCESS' })
        $Global:LookupInFlight = $false
        Set-LookupButtonsIdle
        Update-DecommissionButton

        # #9 — Better edge states: helpful guidance when nothing matched
        if ($found -eq 0 -and $errored -eq 0) {
            $activeCount = 4 - $skipped
            if ($activeCount -gt 0) {
                Show-Toast "No matches across $activeCount system(s). Check the hostname spelling, tenant, or domain." -Type Warning -DurationMs 5000
            }
        } elseif ($found -eq 0 -and $errored -gt 0) {
            Show-Toast "No matches found and $errored system(s) returned errors. Check connectivity and module install." -Type Error -DurationMs 5000
        }
    }
}

# ============================================================================
# AUDIT TRAIL (#1) — append per-decommission record to decommission-history.json
# ============================================================================
function Save-AuditEntry {
    <#
    .SYNOPSIS
        Appends one decommission record (real or dry-run) to
        decommission-history.json. Invalidates the in-memory audit cache so
        subsequent reads see the new entry.
    #>
    param(
        [string]$DeviceQuery,
        [string[]]$Targets,
        [bool]$DryRun,
        [hashtable]$StepResults   # @{ AD=@{Success;Message}; Entra=…; Intune=…; SCCM=… }
    )
    $entry = [ordered]@{
        Timestamp = (Get-Date).ToString('o')
        Operator  = "$env:USERDOMAIN\$env:USERNAME"
        Machine   = $env:COMPUTERNAME
        Device    = $DeviceQuery
        DryRun    = $DryRun
        Targets   = $Targets
        Results   = @{}
    }
    foreach ($t in $Targets) {
        $sr = if ($StepResults -and $StepResults[$t]) { $StepResults[$t] } else { @{ Success=$false; Message='No result' } }
        $entry.Results[$t] = [ordered]@{ Success=[bool]$sr.Success; Message=$sr.Message }
    }
    try {
        $history = @()
        if (Test-Path $Global:AuditFile) {
            $raw = Get-Content $Global:AuditFile -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction SilentlyContinue
            if ($raw) { $history = @($raw) }
        }
        $history += $entry
        ($history | ConvertTo-Json -Depth 5) | Set-Content -Path $Global:AuditFile -Encoding UTF8
        # Invalidate in-memory cache so the next Read-AuditHistory picks up the new entry.
        $Global:AuditHistoryCache = $null
        Write-DebugLog "Audit: wrote entry to $Global:AuditFile (total $($history.Count) entries)" -Level 'DEBUG'
    } catch {
        Write-DebugLog "Audit: failed to write $Global:AuditFile — $($_.Exception.Message)" -Level 'WARN'
    }
}

# In-memory audit history cache. Populated by Read-AuditHistory and invalidated
# by Save-AuditEntry. Avoids parsing the same JSON file three times during one
# decommission flow (Save -> Check-Achievements -> Refresh-HistoryView).
$Global:AuditHistoryCache = $null

function Read-AuditHistory {
    <#
    .SYNOPSIS
        Returns the parsed audit history, loading from disk only if the cache is
        empty. Returns @() on missing/corrupted file (and logs the failure).
    #>
    if ($null -ne $Global:AuditHistoryCache) { return $Global:AuditHistoryCache }
    $rows = @()
    if (Test-Path $Global:AuditFile) {
        try {
            $raw = Get-Content $Global:AuditFile -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($raw) { $rows = @($raw) }
        } catch {
            Write-DebugLog "Read-AuditHistory: failed to parse $Global:AuditFile — $($_.Exception.Message)" -Level 'WARN'
        }
    }
    $Global:AuditHistoryCache = $rows
    return $rows
}

# ============================================================================
# ACHIEVEMENTS — 30 unlockables, persisted to achievements.json
# ============================================================================
$Global:Achievements = @{}
$Global:AchievementDefs = @(
    # First-time milestones
    @{ Id='first_lookup';      Icon=([char]::ConvertFromUtf32(0x1F50D)); Name='First Lookup';        Desc='Looked up your first device.' }
    @{ Id='first_dryrun';      Icon=([char]::ConvertFromUtf32(0x1F9EA)); Name='Cautious Operator';   Desc='Ran your first dry-run.' }
    @{ Id='first_decomm';      Icon=([char]::ConvertFromUtf32(0x1F389)); Name='First Cut';            Desc='Decommissioned your first device.' }
    @{ Id='first_signin';      Icon=([char]::ConvertFromUtf32(0x1F510)); Name='Signed In';            Desc='Signed in to Entra successfully.' }
    @{ Id='first_inventory';   Icon=([char]::ConvertFromUtf32(0x1F4CA)); Name='Inventory Taker';      Desc='Ran your first inventory check.' }
    # Volume tiers
    @{ Id='five_decomms';      Icon=([char]::ConvertFromUtf32(0x1F525)); Name='Warming Up';           Desc='Decommissioned 5 devices.' }
    @{ Id='ten_decomms';       Icon=([char]::ConvertFromUtf32(0x1F3C6)); Name='Decom Pro';            Desc='Decommissioned 10 devices.' }
    @{ Id='twentyfive_decomms';Icon=([char]::ConvertFromUtf32(0x1F48E)); Name='Diamond Disposer';      Desc='Decommissioned 25 devices.' }
    @{ Id='fifty_decomms';     Icon=([char]::ConvertFromUtf32(0x2B50));  Name='Half Century';         Desc='Decommissioned 50 devices.' }
    @{ Id='hundred_decomms';   Icon=([char]::ConvertFromUtf32(0x1F4AF)); Name='Centurion';            Desc='Decommissioned 100 devices.' }
    # Coverage breadth
    @{ Id='all_systems';       Icon=([char]::ConvertFromUtf32(0x1F308)); Name='Full Spectrum';        Desc='Hit AD + Entra + Intune + Autopilot + SCCM in one decommission.' }
    @{ Id='ad_specialist';     Icon=([char]::ConvertFromUtf32(0x1F4C2)); Name='AD Specialist';        Desc='Removed 10 AD computer objects.' }
    @{ Id='cloud_native';      Icon=([char]::ConvertFromUtf32(0x2601));  Name='Cloud Native';         Desc='Removed 10 Entra devices.' }
    @{ Id='intune_tamer';      Icon=([char]::ConvertFromUtf32(0x1F4F1)); Name='Intune Tamer';         Desc='Removed 10 Intune managed devices.' }
    @{ Id='autopilot_ace';     Icon=([char]::ConvertFromUtf32(0x2708));  Name='Autopilot Ace';        Desc='Removed 5 Autopilot device identities.' }
    @{ Id='sccm_cleaner';      Icon=([char]::ConvertFromUtf32(0x1F9F9)); Name='SCCM Cleaner';         Desc='Removed 10 SCCM resources.' }
    # Time-based
    @{ Id='night_owl';         Icon=([char]::ConvertFromUtf32(0x1F989)); Name='Night Owl';            Desc='Decommissioned between midnight and 5am.' }
    @{ Id='early_bird';        Icon=([char]::ConvertFromUtf32(0x1F305)); Name='Early Bird';           Desc='Decommissioned between 5am and 7am.' }
    @{ Id='weekend_warrior';   Icon=([char]::ConvertFromUtf32(0x1F6E1)); Name='Weekend Warrior';      Desc='Decommissioned on a Saturday or Sunday.' }
    @{ Id='speed_demon';       Icon=([char]::ConvertFromUtf32(0x26A1));  Name='Speed Demon';          Desc='Decommissioned in under 10 seconds (wall clock).' }
    # Safety / hygiene
    @{ Id='dryrun_devotee';    Icon=([char]::ConvertFromUtf32(0x1F9EA)); Name='Dry-run Devotee';      Desc='Ran 10 dry-runs.' }
    @{ Id='heeded_warning';    Icon=([char]::ConvertFromUtf32(0x1F6E1)); Name='Heeded the Warning';   Desc='Cancelled a decommission with active safety warnings.' }
    @{ Id='laps_aware';        Icon=([char]::ConvertFromUtf32(0x1F511)); Name='LAPS-Aware';           Desc='Saw a LAPS warning before pressing Decommission.' }
    @{ Id='bitlocker_aware';   Icon=([char]::ConvertFromUtf32(0x1F512)); Name='BitLocker-Aware';      Desc='Saw a BitLocker key warning before pressing Decommission.' }
    @{ Id='no_typos';          Icon=([char]::ConvertFromUtf32(0x270F));  Name='Pen & Paper';          Desc='Typed the device name correctly first try, 10 times.' }
    # Tooling / power-user
    @{ Id='theme_toggle';      Icon=([char]::ConvertFromUtf32(0x1F3A8)); Name='Chameleon';            Desc='Toggled the theme.' }
    @{ Id='wildcard_user';     Icon=([char]::ConvertFromUtf32(0x2733));  Name='Wildcard';             Desc='Looked up a device using a * or ? pattern.' }
    @{ Id='cancelled_lookup';  Icon=([char]::ConvertFromUtf32(0x274C));  Name='Quick Reflexes';       Desc='Cancelled a running lookup.' }
    @{ Id='exported_audit';    Icon=([char]::ConvertFromUtf32(0x1F4C4)); Name='Reporter';             Desc='Exported the audit trail to CSV.' }
    @{ Id='all_unlocked';      Icon=([char]::ConvertFromUtf32(0x1F451)); Name='Completionist';        Desc='Unlocked every other achievement.' }
)

function Load-Achievements {
    $Global:Achievements = @{}
    if (Test-Path $Global:AchievementsFile) {
        try {
            $raw = Get-Content $Global:AchievementsFile -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($prop in $raw.PSObject.Properties) {
                $Global:Achievements[$prop.Name] = $prop.Value
            }
        } catch {
            Write-DebugLog "Achievements load failed: $($_.Exception.Message)" -Level 'WARN'
        }
    }
}
function Save-Achievements {
    try {
        $Global:Achievements | ConvertTo-Json | Set-Content -Path $Global:AchievementsFile -Encoding UTF8
    } catch {
        Write-DebugLog "Achievements save failed: $($_.Exception.Message)" -Level 'WARN'
    }
}
function Unlock-Achievement {
    <#
    .SYNOPSIS
        Marks an achievement as unlocked, persists to disk, fires the toast +
        confetti, and re-renders the achievements grid. Idempotent — calling
        again on an already-unlocked ID is a no-op.
    .NOTES
        After every unlock (other than 'all_unlocked'), checks whether the
        operator has earned the Completionist badge.
    #>
    param([string]$Id)
    if ($Global:Achievements.ContainsKey($Id)) { return }
    $def = $Global:AchievementDefs | Where-Object { $_.Id -eq $Id } | Select-Object -First 1
    if (-not $def) { return }
    $Global:Achievements[$Id] = (Get-Date).ToString('o')
    Save-Achievements
    Show-Toast "$($def.Icon) Achievement Unlocked: $($def.Name) — $($def.Desc)" -Type Success -DurationMs 5000
    Start-ConfettiAnimation
    Render-Achievements
    Write-DebugLog "Achievement unlocked: $($def.Name) ($Id)" -Level 'SUCCESS'
    # Completionist check (avoid recursion: only check when something else unlocks)
    if ($Id -ne 'all_unlocked') {
        $expected = ($Global:AchievementDefs | Where-Object { $_.Id -ne 'all_unlocked' }).Count
        $unlocked = ($Global:AchievementDefs | Where-Object { $_.Id -ne 'all_unlocked' -and $Global:Achievements.ContainsKey($_.Id) }).Count
        if ($unlocked -ge $expected) { Unlock-Achievement 'all_unlocked' }
    }
}
function Render-Achievements {
    if (-not $pnlAchievementsGrid) { return }
    $pnlAchievementsGrid.Children.Clear()
    $unlockedCount = 0
    foreach ($def in $Global:AchievementDefs) {
        $isUnlocked = $Global:Achievements.ContainsKey($def.Id)
        if ($isUnlocked) { $unlockedCount++ }
        $border = New-Object System.Windows.Controls.Border
        $border.Width  = 130
        $border.Height = 110
        $border.CornerRadius = [System.Windows.CornerRadius]::new(8)
        $border.Margin = [System.Windows.Thickness]::new(0,0,8,8)
        $border.Padding = [System.Windows.Thickness]::new(10,10,10,10)
        $border.BorderThickness = [System.Windows.Thickness]::new(1)
        if ($isUnlocked) {
            $border.Background = $Window.Resources['ThemeCardBg']
            $border.BorderBrush = $Window.Resources['ThemeAccentLight']
            $border.ToolTip = "$($def.Name) — $($def.Desc)`nUnlocked: $($Global:Achievements[$def.Id])"
        } else {
            $border.Background = $Window.Resources['ThemeDeepBg']
            $border.BorderBrush = $Window.Resources['ThemeBorder']
            $border.Opacity = 0.55
            $border.ToolTip = "Locked — $($def.Desc)"
        }
        $sp = New-Object System.Windows.Controls.StackPanel
        $sp.HorizontalAlignment = 'Stretch'
        $iconTb = New-Object System.Windows.Controls.TextBlock
        $iconTb.Text = if ($isUnlocked) { $def.Icon } else { '?' }
        $iconTb.FontSize = 28
        $iconTb.HorizontalAlignment = 'Center'
        if ($isUnlocked) { $iconTb.FontFamily = New-Object System.Windows.Media.FontFamily('Segoe UI Emoji') }
        else { $iconTb.Foreground = $Window.Resources['ThemeTextDisabled'] }
        $nameTb = New-Object System.Windows.Controls.TextBlock
        $nameTb.Text = $def.Name
        $nameTb.FontSize = 11
        $nameTb.FontWeight = 'SemiBold'
        $nameTb.HorizontalAlignment = 'Center'
        $nameTb.TextAlignment = 'Center'
        $nameTb.TextWrapping = 'Wrap'
        $nameTb.Margin = [System.Windows.Thickness]::new(0,6,0,0)
        $nameTb.Foreground = if ($isUnlocked) { $Window.Resources['ThemeTextPrimary'] } else { $Window.Resources['ThemeTextDim'] }
        [void]$sp.Children.Add($iconTb)
        [void]$sp.Children.Add($nameTb)
        $border.Child = $sp
        [void]$pnlAchievementsGrid.Children.Add($border)
    }
    if ($lblAchievementsCount) {
        $lblAchievementsCount.Text = "$unlockedCount/$($Global:AchievementDefs.Count)"
    }
}

# ============================================================================
# CONFETTI ANIMATION (celebratory burst on unlock)
# ============================================================================
function Start-ConfettiAnimation {
    <#
    .SYNOPSIS
        Drops 60 falling/spinning rectangles across the window for ~3-5s.
        Each particle gets independent fall, drift, and spin animations.
    .NOTES
        Click-through canvas (IsHitTestVisible=False) at z-index 9999.
        Auto-cleanup DispatcherTimer fires after $Script:CONFETTI_CLEANUP_MS.
        Re-entrant calls are dropped while the canvas is already visible.
    #>
    if (-not $cnvConfetti) { return }
    if ($cnvConfetti.Visibility -eq 'Visible') { return }
    $W = $Window.ActualWidth
    $H = $Window.ActualHeight
    if ($W -le 0 -or $H -le 0) { return }
    $cnvConfetti.Children.Clear()
    $cnvConfetti.Visibility = 'Visible'
    $colors = @('#FF4444','#FFD700','#00C853','#60CDFF','#0078D4','#B388FF','#FF6D00','#E040FB')
    $rand = [System.Random]::new()
    for ($i = 0; $i -lt $Script:CONFETTI_COUNT; $i++) {
        $size = $rand.Next(4, 10)
        $rect = New-Object System.Windows.Shapes.Rectangle
        $rect.Width  = $size
        $rect.Height = $size * ($rand.NextDouble() * 1.5 + 0.5)
        $rect.Fill   = $Global:CachedBC.ConvertFromString($colors[$rand.Next($colors.Count)])
        $rect.RadiusX = if ($rand.Next(3) -eq 0) { $size / 2 } else { 1 }
        $rect.RadiusY = $rect.RadiusX
        $rect.Opacity = 0.9
        $rect.RenderTransform = [System.Windows.Media.RotateTransform]::new($rand.Next(360))
        $X0 = $rand.NextDouble() * $W
        [System.Windows.Controls.Canvas]::SetLeft($rect, $X0)
        [System.Windows.Controls.Canvas]::SetTop($rect, -20)
        [void]$cnvConfetti.Children.Add($rect)
        $fall = New-Object System.Windows.Media.Animation.DoubleAnimation
        $fall.From = $rand.Next(-40, -10)
        $fall.To   = $H + 20
        $fall.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds($rand.Next(2000, 4500)))
        $ease = New-Object System.Windows.Media.Animation.CubicEase
        $ease.EasingMode = 'EaseIn'
        $fall.EasingFunction = $ease
        $drift = New-Object System.Windows.Media.Animation.DoubleAnimation
        $drift.From = $X0
        $drift.To   = $X0 + $rand.Next(-120, 120)
        $drift.Duration = $fall.Duration
        $spin = New-Object System.Windows.Media.Animation.DoubleAnimation
        $spin.From = 0
        $spin.To   = $rand.Next(-720, 720)
        $spin.Duration = $fall.Duration
        $rect.BeginAnimation([System.Windows.Controls.Canvas]::TopProperty, $fall)
        $rect.BeginAnimation([System.Windows.Controls.Canvas]::LeftProperty, $drift)
        $rect.RenderTransform.BeginAnimation([System.Windows.Media.RotateTransform]::AngleProperty, $spin)
    }
    $cleanTimer = New-Object System.Windows.Threading.DispatcherTimer
    $cleanTimer.Interval = [TimeSpan]::FromMilliseconds($Script:CONFETTI_CLEANUP_MS)
    $cleanTimer.Tag = $cnvConfetti
    $cleanTimer.Add_Tick({
        try {
            $this.Tag.Children.Clear()
            $this.Tag.Visibility = 'Collapsed'
            $this.Stop()
        } catch {
            Write-DebugLog "[ConfettiTimer] cleanup error: $($_.Exception.Message)" -Level 'ERROR'
        }
    })
    $cleanTimer.Start()
}

# ============================================================================
# ACHIEVEMENT TRIGGERS — called from key code paths to evaluate unlock conditions
# ============================================================================
function Check-DecommissionAchievements {
    param(
        [bool]$DryRun,
        [string[]]$Targets,
        [hashtable]$Steps,
        [string]$DeviceName
    )
    # Read the persisted audit file (cached) for total counts.
    $history = Read-AuditHistory
    $real    = @($history | Where-Object { -not $_.DryRun })
    $dryruns = @($history | Where-Object { $_.DryRun })

    # First-time milestones
    if ($DryRun) { Unlock-Achievement 'first_dryrun' }
    else         { Unlock-Achievement 'first_decomm' }

    # Volume tiers (real decommissions only)
    if ($real.Count -ge 5)   { Unlock-Achievement 'five_decomms' }
    if ($real.Count -ge 10)  { Unlock-Achievement 'ten_decomms' }
    if ($real.Count -ge 25)  { Unlock-Achievement 'twentyfive_decomms' }
    if ($real.Count -ge 50)  { Unlock-Achievement 'fifty_decomms' }
    if ($real.Count -ge 100) { Unlock-Achievement 'hundred_decomms' }
    if ($dryruns.Count -ge 10) { Unlock-Achievement 'dryrun_devotee' }

    # Coverage breadth — count real successful per-system removals across history
    if (-not $DryRun -and $Steps) {
        $perSystemCount = @{ AD=0; Entra=0; Intune=0; Autopilot=0; SCCM=0 }
        foreach ($e in $real) {
            if ($e.Results) {
                foreach ($prop in $e.Results.PSObject.Properties) {
                    if ($prop.Value.Success) { $perSystemCount[$prop.Name]++ }
                }
            }
        }
        if ($perSystemCount.AD        -ge 10) { Unlock-Achievement 'ad_specialist' }
        if ($perSystemCount.Entra     -ge 10) { Unlock-Achievement 'cloud_native' }
        if ($perSystemCount.Intune    -ge 10) { Unlock-Achievement 'intune_tamer' }
        if ($perSystemCount.Autopilot -ge 5)  { Unlock-Achievement 'autopilot_ace' }
        if ($perSystemCount.SCCM      -ge 10) { Unlock-Achievement 'sccm_cleaner' }

        # Full spectrum — all 5 in this single run, all successful
        $five = @('AD','Entra','Intune','Autopilot','SCCM')
        $hitAll = $true
        foreach ($s in $five) {
            if (-not $Steps[$s] -or -not $Steps[$s].Success) { $hitAll = $false; break }
        }
        if ($hitAll) { Unlock-Achievement 'all_systems' }
    }

    # Time-based (only for real runs to avoid easy unlocks via dry-runs)
    if (-not $DryRun) {
        $now = Get-Date
        if ($now.Hour -ge 0 -and $now.Hour -lt 5) { Unlock-Achievement 'night_owl' }
        if ($now.Hour -ge 5 -and $now.Hour -lt 7) { Unlock-Achievement 'early_bird' }
        if ($now.DayOfWeek -in @('Saturday','Sunday')) { Unlock-Achievement 'weekend_warrior' }
        if ($Global:DecommStartTime) {
            $elapsed = ($now - $Global:DecommStartTime).TotalSeconds
            if ($elapsed -lt 10) { Unlock-Achievement 'speed_demon' }
        }
    }
}

# ============================================================================
# PRE-FLIGHT SUMMARY (#2) — rich confirm dialog with device details
# ============================================================================
function Build-PreflightSummary {
    param([string]$DeviceQuery, [string[]]$Targets)
    $items   = New-Object System.Collections.ArrayList
    $results = $Global:LookupState.Results

    # Helper: format ISO date as "yyyy-MM-dd HH:mm (X days ago)" for readability.
    $fmtDate = {
        param([string]$iso)
        if (-not $iso) { return '' }
        try {
            $dt = [datetime]::Parse($iso)
            $ago = (Get-Date) - $dt
            $rel = if ($ago.TotalDays -ge 1) { "$([math]::Round($ago.TotalDays, 0)) day(s) ago" }
                   elseif ($ago.TotalHours -ge 1) { "$([math]::Round($ago.TotalHours, 0)) hour(s) ago" }
                   else { "$([math]::Round($ago.TotalMinutes, 0)) min ago" }
            return ('{0} ({1})' -f $dt.ToString('yyyy-MM-dd HH:mm'), $rel)
        } catch { return $iso }
    }

    foreach ($sys in $Targets) {
        $r = $results[$sys]
        if (-not $r -or -not $r.Found) { continue }
        $raw = $r.Raw
        $fields = New-Object System.Collections.ArrayList
        switch ($sys) {
            'AD' {
                if ($raw.OperatingSystem)   { [void]$fields.Add(@{ Label='OS';          Value=[string]$raw.OperatingSystem }) }
                if ($null -ne $raw.Enabled) { [void]$fields.Add(@{ Label='Enabled';     Value=[string]$raw.Enabled }) }
                if ($raw.LastLogonDate)     { [void]$fields.Add(@{ Label='Last logon';  Value=(& $fmtDate $raw.LastLogonDate) }) }
                if ($raw.Description)       { [void]$fields.Add(@{ Label='Description'; Value=[string]$raw.Description }) }
                if ($raw.DistinguishedName) { [void]$fields.Add(@{ Label='DN';          Value=[string]$raw.DistinguishedName }) }
                $label = 'AD'
            }
            'Entra' {
                if ($raw.OperatingSystem) {
                    $os = if ($raw.OperatingSystemVersion) { "$($raw.OperatingSystem) $($raw.OperatingSystemVersion)" } else { [string]$raw.OperatingSystem }
                    [void]$fields.Add(@{ Label='OS'; Value=$os })
                }
                if ($null -ne $raw.AccountEnabled)        { [void]$fields.Add(@{ Label='Enabled';      Value=[string]$raw.AccountEnabled }) }
                if ($raw.ApproximateLastSignInDateTime)   { [void]$fields.Add(@{ Label='Last sign-in'; Value=(& $fmtDate $raw.ApproximateLastSignInDateTime) }) }
                if ($raw.Id)                              { [void]$fields.Add(@{ Label='ObjectId';     Value=[string]$raw.Id }) }
                $label = 'Entra'
            }
            'Intune' {
                if ($raw.OperatingSystem)   { [void]$fields.Add(@{ Label='OS';        Value=[string]$raw.OperatingSystem }) }
                if ($raw.UserPrincipalName) { [void]$fields.Add(@{ Label='User';      Value=[string]$raw.UserPrincipalName }) }
                if ($raw.LastSyncDateTime)  { [void]$fields.Add(@{ Label='Last sync'; Value=(& $fmtDate $raw.LastSyncDateTime) }) }
                $label = 'Intune'
            }
            'Autopilot' {
                if ($raw.SerialNumber) { [void]$fields.Add(@{ Label='Serial';    Value=[string]$raw.SerialNumber }) }
                if ($raw.Model)        { [void]$fields.Add(@{ Label='Model';     Value=[string]$raw.Model }) }
                if ($raw.GroupTag)     { [void]$fields.Add(@{ Label='Group tag'; Value=[string]$raw.GroupTag }) }
                $label = 'Autopilot'
            }
            'SCCM' {
                if ($raw.ResourceID) { [void]$fields.Add(@{ Label='ResourceID'; Value=[string]$raw.ResourceID }) }
                $label = 'SCCM'
            }
        }
        [void]$items.Add([pscustomobject]@{
            System      = $label
            DisplayName = [string]$r.DisplayName
            Fields      = @($fields | ForEach-Object { [pscustomobject]$_ })
        })
    }
    return @($items)
}

# ============================================================================
# SAFETY WARNINGS (#3 BitLocker, #4 LAPS, #5 Recently Active)
# ============================================================================
function Get-SafetyWarnings {
    param([string[]]$Targets)
    $warnings = @()
    $results  = $Global:LookupState.Results
    $recentDays = if ($Global:Settings.RecentActivityDays) { [int]$Global:Settings.RecentActivityDays } else { 7 }
    $now = Get-Date

    foreach ($sys in $Targets) {
        $r = $results[$sys]
        if (-not $r -or -not $r.Found -or -not $r.Raw) { continue }
        $raw = $r.Raw

        # ---- BitLocker recovery key warning (#3) ----
        if ($sys -eq 'AD' -and $raw.BitLockerKeyCount -and [int]$raw.BitLockerKeyCount -gt 0) {
            $warnings += [char]0x26A0 + " AD: $($raw.BitLockerKeyCount) BitLocker recovery key(s) escrowed under this computer object will be LOST."
        }
        if ($sys -eq 'Entra' -and $raw.BitLockerKeyCount -and [int]$raw.BitLockerKeyCount -gt 0) {
            $warnings += [char]0x26A0 + " Entra: $($raw.BitLockerKeyCount) BitLocker recovery key(s) escrowed in Entra ID will be LOST."
        }

        # ---- LAPS password warning (#4) ----
        if ($sys -eq 'AD' -and $raw.HasLAPSPassword) {
            $warnings += [char]0x26A0 + " AD: LAPS password is currently stored — it will be lost when the object is deleted."
        }

        # ---- Recently-active guard (#5) ----
        if ($sys -eq 'AD' -and $raw.LastLogonDate) {
            try {
                $lastLogon = [datetime]::Parse($raw.LastLogonDate)
                $ago = ($now - $lastLogon).TotalDays
                if ($ago -lt $recentDays) {
                    $warnings += [char]0x26A0 + " AD: last logon was $([math]::Round($ago, 1)) day(s) ago — device may still be in use."
                }
            } catch {}
        }
        if ($sys -eq 'Entra' -and $raw.ApproximateLastSignInDateTime) {
            try {
                $lastSign = [datetime]::Parse($raw.ApproximateLastSignInDateTime)
                $ago = ($now - $lastSign).TotalDays
                if ($ago -lt $recentDays) {
                    $warnings += [char]0x26A0 + " Entra: last sign-in was $([math]::Round($ago, 1)) day(s) ago — device may still be in use."
                }
            } catch {}
        }
        if ($sys -eq 'Intune' -and $raw.LastSyncDateTime) {
            try {
                $lastSync = [datetime]::Parse($raw.LastSyncDateTime)
                $ago = ($now - $lastSync).TotalDays
                if ($ago -lt $recentDays) {
                    $warnings += [char]0x26A0 + " Intune: last sync was $([math]::Round($ago, 1)) day(s) ago — device may still be in use."
                }
            } catch {}
        }
    }
    return $warnings
}

# ============================================================================
# DECOMMISSION
# ============================================================================
function Start-Decommission {
    Write-DebugLog '[Decommission] >>> ENTER button click' -Level 'DEBUG'
    $q = $Global:LookupState.DeviceQuery
    if (-not $q) { Write-DebugLog '[Decommission] aborted: no DeviceQuery in state' -Level 'DEBUG'; return }
    $targets = @()
    foreach ($s in @('Intune','Autopilot','Entra','AD','SCCM')) {
        $chk = Get-Variable -Name "chkTarget$s" -ValueOnly -Scope Script -ErrorAction SilentlyContinue
        if (-not $chk) { continue }
        $r = $Global:LookupState.Results[$s]
        if ($chk.IsChecked -and $r -and $r.Found) { $targets += $s }
    }
    Write-DebugLog "[Decommission] candidate targets after filter: $($targets -join ',')" -Level 'DEBUG'
    if ($targets.Count -eq 0) { Write-DebugLog '[Decommission] aborted: no checked-and-found targets' -Level 'DEBUG'; Show-Toast 'No matching targets selected' -Type Warning; return }

    # Build pre-flight summary (#2) and safety warnings (#3,#4,#5)
    $preflight = Build-PreflightSummary -DeviceQuery $q -Targets $targets
    $warnings  = Get-SafetyWarnings -Targets $targets
    if ($warnings.Count -gt 0) {
        Write-DebugLog "[Decommission] safety warnings: $($warnings.Count)" -Level 'WARN'
        # Note that the operator saw specific warning categories — used by achievements
        $warningJoined = ($warnings -join ' ')
        if ($warningJoined -match 'LAPS')      { Unlock-Achievement 'laps_aware' }
        if ($warningJoined -match 'BitLocker') { Unlock-Achievement 'bitlocker_aware' }
        $Global:LastWarnedDevice = $q
    } else {
        $Global:LastWarnedDevice = $null
    }

    $dry = [bool]$chkDryRun.IsChecked
    Write-DebugLog "[Decommission] dryRun=$dry" -Level 'DEBUG'
    if ($dry) {
        # Dry-run: nothing destructive; skip the type-name confirmation to reduce friction.
        $list = ($targets -join ', ')
        $body = "Dry-run for '$q' — will validate connectivity, scopes, and lookup completeness for $list. No destructive cmdlet will be executed."
        Show-ModalConfirm -Title 'Confirm dry-run' -Icon ([string][char]0xE7BA) `
            -Message $body `
            -DetailItems $preflight -Warnings $warnings `
            -ConfirmLabel 'Run dry-run' -OnConfirm {
                Hide-Modal
                Invoke-DecommissionWork -DeviceQuery $q -Targets $targets -DryRun
            }.GetNewClosure()
        return
    }

    $body = "This will PERMANENTLY remove '$q' from " + ($targets -join ', ') + "."
    Show-ModalInput -Title 'Confirm decommission' -Message $body -ConfirmLabel 'Decommission' `
                    -Icon ([string][char]0xE74D) `
                    -DetailItems $preflight -Warnings $warnings `
                    -InputPrompt "Type the device name ($q) to confirm:" -OnConfirm {
        $typed = $txtModalInput.Text.Trim()
        Hide-Modal
        if ([string]::Compare($typed, $q, [System.StringComparison]::OrdinalIgnoreCase) -ne 0) {
            Show-Toast 'Confirmation text did not match - aborted.' -Type Error
            $Global:CorrectFirstTryCount = 0
            return
        }
        # Track first-try correct typing for the 'no_typos' achievement (10 in a row)
        $Global:CorrectFirstTryCount = [int]$Global:CorrectFirstTryCount + 1
        if ($Global:CorrectFirstTryCount -ge 10) { Unlock-Achievement 'no_typos' }
        Invoke-DecommissionWork -DeviceQuery $q -Targets $targets
    }.GetNewClosure()
}

function Invoke-DecommissionWork {
    <#
    .SYNOPSIS
        Executes the per-system removal cmdlets in a background runspace.
    .DESCRIPTION
        Iterates Targets in order Intune → Autopilot → Entra → AD → SCCM
        (chosen so cloud removals happen before AD; if AD replication is
        slow it doesn't block the Graph operations). Each system runs in
        its own try/catch so a single failure doesn't abort the others.
    .PARAMETER DryRun
        When set, validates connectivity / scopes / object existence but
        does NOT call any Remove-* cmdlet. Cards flip to 'Would remove'.
    #>
    param([string]$DeviceQuery,[string[]]$Targets,[switch]$DryRun)
    $tag = if ($DryRun) { 'DRY-RUN' } else { 'DECOMMISSION' }
    Write-DebugLog "=== $tag START - device='$DeviceQuery' targets=$($Targets -join ',') ===" -Level 'WARN'
    $Global:DecommStartTime = Get-Date
    Set-Status -Text $(if ($DryRun) { 'Dry-run in progress' } else { 'Decommissioning' }) -Detail $DeviceQuery -State Busy
    Show-GlobalProgress -Text $(if ($DryRun) { "Validating '$DeviceQuery'..." } else { "Decommissioning '$DeviceQuery'..." })
    $btnDecommission.IsEnabled = $false
    $btnLookup.IsEnabled       = $false

    # Snapshot inputs to pass to background runspace
    $payload = @{
        DeviceQuery = $DeviceQuery
        Targets     = $Targets
        Results     = $Global:LookupState.Results
        Settings    = $Global:Settings
        CredAD      = $Global:Creds.AD
        CredSCCM    = $Global:Creds.SCCM
        DryRun      = [bool]$DryRun
    }

    Start-BackgroundWork -Variables $payload -Work {
        $log = New-Object System.Collections.ArrayList
        $stepResults = @{}

        function _Log { param($m,$lvl='INFO') [void]$log.Add(@{ Time=(Get-Date).ToString('HH:mm:ss.fff'); Level=$lvl; Msg=$m }) }

        # ----- Intune -----
        if ($Targets -contains 'Intune') {
            try {
                _Log $(if ($DryRun) { 'Intune: DRY-RUN — validating prerequisites' } else { 'Intune: removing managed device' })
                if (-not $Results.Intune -or -not $Results.Intune.Raw -or -not $Results.Intune.Raw.Id) {
                    throw 'Intune lookup result is incomplete — missing device id.'
                }
                Import-Module Microsoft.Graph.DeviceManagement -ErrorAction Stop
                $needScope = 'DeviceManagementManagedDevices.ReadWrite.All'
                $ctx = $null; try { $ctx = Get-MgContext -ErrorAction SilentlyContinue } catch {}
                if (-not $ctx) {
                    throw 'Not signed in to Microsoft Graph. Sign in via the toolbar first.'
                }
                if (-not ($ctx.Scopes -contains $needScope)) {
                    throw "Missing Graph scope '$needScope'. Sign out and back in to acquire it."
                }
                $id = $Results.Intune.Raw.Id
                if ($DryRun) {
                    $stepResults.Intune = @{ Success=$true; DryRun=$true; Message="Would call Remove-MgDeviceManagementManagedDevice -ManagedDeviceId $id" }
                    _Log "Intune DRY-RUN OK (id=$id)" 'SUCCESS'
                } else {
                    Remove-MgDeviceManagementManagedDevice -ManagedDeviceId $id -ErrorAction Stop
                    $stepResults.Intune = @{ Success=$true;  Message="Removed managed device id=$id" }
                    _Log "Intune: removed (id=$id)" 'SUCCESS'
                }
            } catch {
                $stepResults.Intune = @{ Success=$false; Message=$_.Exception.Message }
                _Log "Intune FAILED: $($_.Exception.Message)" 'ERROR'
            }
        }

        # ----- Entra -----
        if ($Targets -contains 'Entra') {
            try {
                _Log $(if ($DryRun) { 'Entra: DRY-RUN — validating prerequisites' } else { 'Entra: removing device object' })
                if (-not $Results.Entra -or -not $Results.Entra.Raw -or -not $Results.Entra.Raw.Id) {
                    throw 'Entra lookup result is incomplete — missing object id.'
                }
                Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
                $needScope = 'Device.ReadWrite.All'
                $ctx = $null; try { $ctx = Get-MgContext -ErrorAction SilentlyContinue } catch {}
                if (-not $ctx) {
                    throw 'Not signed in to Microsoft Graph. Sign in via the toolbar first.'
                }
                if (-not ($ctx.Scopes -contains $needScope)) {
                    throw "Missing Graph scope '$needScope'. Sign out and back in to acquire it."
                }
                $id = $Results.Entra.Raw.Id
                if ($DryRun) {
                    $stepResults.Entra = @{ Success=$true; DryRun=$true; Message="Would call Remove-MgDevice -DeviceId $id" }
                    _Log "Entra DRY-RUN OK (id=$id)" 'SUCCESS'
                } else {
                    Remove-MgDevice -DeviceId $id -ErrorAction Stop
                    $stepResults.Entra = @{ Success=$true;  Message="Removed Entra device id=$id" }
                    _Log "Entra: removed (id=$id)" 'SUCCESS'
                }
            } catch {
                $stepResults.Entra = @{ Success=$false; Message=$_.Exception.Message }
                _Log "Entra FAILED: $($_.Exception.Message)" 'ERROR'
            }
        }

        # ----- AD -----
        if ($Targets -contains 'AD') {
            try {
                _Log $(if ($DryRun) { 'AD: DRY-RUN — verifying object exists' } else { 'AD: removing computer object' })
                if (-not $Results.AD -or -not $Results.AD.Raw -or -not $Results.AD.Raw.DistinguishedName) {
                    throw 'AD lookup result is incomplete — missing distinguishedName.'
                }
                Import-Module ActiveDirectory -ErrorAction Stop
                $params = @{ ErrorAction='Stop'; Confirm=$false }
                if ($Settings.AD.Server) { $params.Server = $Settings.AD.Server }
                if ($CredAD)             { $params.Credential = $CredAD }
                $dn = $Results.AD.Raw.DistinguishedName
                if ($DryRun) {
                    # Verify object still exists; do not remove.
                    $getParams = @{ Identity=$dn; Properties='DistinguishedName'; ErrorAction='Stop' }
                    if ($Settings.AD.Server) { $getParams.Server = $Settings.AD.Server }
                    if ($CredAD)             { $getParams.Credential = $CredAD }
                    Get-ADObject @getParams | Out-Null
                    $stepResults.AD = @{ Success=$true; DryRun=$true; Message="Would call Remove-ADObject -Identity '$dn' -Recursive" }
                    _Log "AD DRY-RUN OK ($dn)" 'SUCCESS'
                } else {
                    # Remove leaf and any child objects (e.g., msDS-DeviceRegistration leftovers).
                    try {
                        Get-ADObject -Identity $dn -Properties DistinguishedName -ErrorAction Stop | Out-Null
                        Remove-ADObject -Identity $dn -Recursive @params
                        $stepResults.AD = @{ Success=$true;  Message="Removed AD object $dn" }
                        _Log "AD: removed ($dn)" 'SUCCESS'
                    } catch {
                        $msg = $_.Exception.Message
                        if ($msg -match 'Cannot find an object|does not exist|cannot be found') {
                            $stepResults.AD = @{ Success=$true; Message="AD object already absent ($dn)" }
                            _Log "AD: object already absent (treated as success)" 'SUCCESS'
                        } else { throw }
                    }
                }
            } catch {
                $stepResults.AD = @{ Success=$false; Message=$_.Exception.Message }
                _Log "AD FAILED: $($_.Exception.Message)" 'ERROR'
            }
        }

        # ----- Autopilot -----
        if ($Targets -contains 'Autopilot') {
            try {
                _Log $(if ($DryRun) { 'Autopilot: DRY-RUN — validating prerequisites' } else { 'Autopilot: removing device identity' })
                if (-not $Results.Autopilot -or -not $Results.Autopilot.Raw -or -not $Results.Autopilot.Raw.Id) {
                    throw 'Autopilot lookup result is incomplete — missing device identity id.'
                }
                Import-Module Microsoft.Graph.DeviceManagement.Enrollment -ErrorAction Stop
                $needScope = 'DeviceManagementServiceConfig.ReadWrite.All'
                $ctx = $null; try { $ctx = Get-MgContext -ErrorAction SilentlyContinue } catch {}
                if (-not $ctx) {
                    throw 'Not signed in to Microsoft Graph. Sign in via the toolbar first.'
                }
                if (-not ($ctx.Scopes -contains $needScope)) {
                    throw "Missing Graph scope '$needScope'. Sign out and back in to acquire it."
                }
                $id = $Results.Autopilot.Raw.Id
                if ($DryRun) {
                    $stepResults.Autopilot = @{ Success=$true; DryRun=$true; Message="Would call Remove-MgDeviceManagementWindowsAutopilotDeviceIdentity -WindowsAutopilotDeviceIdentityId $id" }
                    _Log "Autopilot DRY-RUN OK (id=$id)" 'SUCCESS'
                } else {
                    Remove-MgDeviceManagementWindowsAutopilotDeviceIdentity -WindowsAutopilotDeviceIdentityId $id -ErrorAction Stop
                    $stepResults.Autopilot = @{ Success=$true; Message="Removed Autopilot device identity id=$id" }
                    _Log "Autopilot: removed (id=$id)" 'SUCCESS'
                }
            } catch {
                $stepResults.Autopilot = @{ Success=$false; Message=$_.Exception.Message }
                _Log "Autopilot FAILED: $($_.Exception.Message)" 'ERROR'
            }
        }

        # ----- SCCM -----
        if ($Targets -contains 'SCCM') {
            try {
                _Log $(if ($DryRun) { 'SCCM: DRY-RUN — validating prerequisites' } else { 'SCCM: removing device record' })
                if (-not $Results.SCCM -or -not $Results.SCCM.Raw -or -not $Results.SCCM.Raw.ResourceID) {
                    throw 'SCCM lookup result is incomplete — missing ResourceID.'
                }
                $cmModule = Join-Path $env:SMS_ADMIN_UI_PATH '..\ConfigurationManager.psd1'
                if (-not (Test-Path $cmModule)) {
                    $cmModule = "${env:ProgramFiles(x86)}\Microsoft Endpoint Manager\AdminConsole\bin\ConfigurationManager.psd1"
                }
                Import-Module $cmModule -ErrorAction Stop
                $sc = $Settings.SCCM.SiteCode
                if (-not (Get-PSDrive -Name $sc -ErrorAction SilentlyContinue)) {
                    $np = @{ Name=$sc; PSProvider='CMSite'; Root=$Settings.SCCM.Server; ErrorAction='Stop' }
                    if ($CredSCCM) { $np.Credential = $CredSCCM }
                    New-PSDrive @np | Out-Null
                }
                Push-Location ($sc + ':')
                try {
                    if ($DryRun) {
                        $rid = $Results.SCCM.Raw.ResourceID
                        # Verify the resource still exists.
                        $cur = Get-CMDevice -ResourceId $rid -ErrorAction SilentlyContinue
                        if (-not $cur) { throw "SCCM resource $rid no longer exists." }
                        $stepResults.SCCM = @{ Success=$true; DryRun=$true; Message="Would call Remove-CMDevice -ResourceId $rid -Force" }
                        _Log "SCCM DRY-RUN OK (ResourceID=$rid)" 'SUCCESS'
                    } else {
                        Remove-CMDevice -ResourceId $Results.SCCM.Raw.ResourceID -Force -ErrorAction Stop
                        $stepResults.SCCM = @{ Success=$true; Message="Removed SCCM resource $($Results.SCCM.Raw.ResourceID)" }
                        _Log "SCCM: removed (ResourceID=$($Results.SCCM.Raw.ResourceID))" 'SUCCESS'
                    }
                } finally { Pop-Location }
            } catch {
                $stepResults.SCCM = @{ Success=$false; Message=$_.Exception.Message }
                _Log "SCCM FAILED: $($_.Exception.Message)" 'ERROR'
            }
        }

        return @{ Steps=$stepResults; Log=$log; IsDryRun=$DryRun; Device=$DeviceQuery; Targets=$Targets }
    } -OnComplete {
        param($Results,$Errors)
        $r = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { $null }
        if ($r -and $r.Log) {
            foreach ($entry in $r.Log) { Write-DebugLog $entry.Msg -Level $entry.Level }
        }
        # Recover device + targets from the result hashtable. The OnComplete runs in the poll
        # timer scope where the Invoke-DecommissionWork parameters ($DeviceQuery, $Targets) are
        # no longer in scope — we can't use .GetNewClosure() because $btnDecommission and other
        # FindName Script-scope variables would also become invisible.
        $deviceName = if ($r) { [string]$r.Device } else { '' }
        $targetList = if ($r) { @($r.Targets) } else { @() }

        $okCount = 0; $failCount = 0; $isDry = [bool]($r -and $r.IsDryRun)
        if ($r -and $r.Steps) {
            foreach ($k in $r.Steps.Keys) {
                $st = $r.Steps[$k]
                if ($st.Success) {
                    $cardState = if ($st.DryRun) { 'WouldRemove' } else { 'Removed' }
                    Set-CardStatus -System $k -State $cardState -Detail $st.Message
                    $okCount++
                } else {
                    Set-CardStatus -System $k -State Error -Detail $st.Message
                    $failCount++
                }
            }
        }
        Hide-GlobalProgress
        Set-LookupButtonsIdle
        $btnDecommission.IsEnabled = $false
        $verb = if ($isDry) { 'Dry-run' } else { 'Decommission' }
        $verbed = if ($isDry) { 'validated' } else { 'removed' }

        # #8 — Per-system result icons in toast
        $iconParts = @()
        foreach ($sys in @('AD','Entra','Intune','Autopilot','SCCM')) {
            if ($r -and $r.Steps -and $r.Steps[$sys]) {
                $st = $r.Steps[$sys]
                $icon = if ($st.Success) { [char]0x2713 } else { [char]0x2717 }
                $iconParts += "$sys $icon"
            }
        }
        $iconSummary = if ($iconParts.Count -gt 0) { " ($($iconParts -join '  '))" } else { '' }

        if ($failCount -eq 0) {
            Set-Status -Text "$verb complete" -Detail "$okCount target(s) $verbed" -State Success
            Show-Toast "$verb OK$iconSummary" -Type Success
        } elseif ($okCount -gt 0) {
            Set-Status -Text "$verb completed with errors" -Detail "$okCount $verbed, $failCount failed" -State Warning
            Show-Toast "$verb partial$iconSummary" -Type Warning
        } else {
            Set-Status -Text "$verb failed" -Detail 'All targets failed' -State Error
            Show-Toast "$verb failed$iconSummary" -Type Error
        }
        Write-DebugLog "=== $($verb.ToUpper()) END - ok=$okCount fail=$failCount device='$deviceName' ===" -Level 'WARN'

        # Audit trail (#1) — persists to decommission-history.json
        $stepHash = if ($r -and $r.Steps) { $r.Steps } else { @{} }
        Save-AuditEntry -DeviceQuery $deviceName -Targets $targetList -DryRun $isDry -StepResults $stepHash
        # Achievements: evaluate after audit so counts include this run
        Check-DecommissionAchievements -DryRun:$isDry -Targets $targetList -Steps $stepHash -DeviceName $deviceName

        # Re-check: auto re-run lookup to confirm objects are gone (real decommissions only)
        if (-not $isDry -and $okCount -gt 0 -and $deviceName) {
            Write-DebugLog "[Decommission] scheduling re-check lookup for '$deviceName' in 3s" -Level 'INFO'
            # Stash the device name in a Global so the Tick handler can read it without closure capture.
            $Global:RecheckDevice = $deviceName
            if ($Global:RecheckTimer) { try { $Global:RecheckTimer.Stop() } catch {} }
            $Global:RecheckTimer = New-Object System.Windows.Threading.DispatcherTimer
            $Global:RecheckTimer.Interval = [TimeSpan]::FromSeconds(3)
            $Global:RecheckTimer.Add_Tick({
                try { $Global:RecheckTimer.Stop() } catch {}
                $q = $Global:RecheckDevice
                Write-DebugLog "[Recheck] re-running lookup for '$q' to verify removal" -Level 'INFO'
                if ($q) { Start-DeviceLookup -Query $q }
            })
            $Global:RecheckTimer.Start()
        }
    }
}

# ============================================================================
# SETTINGS DIALOG WIRING
# ============================================================================
function Set-SettingsTab {
    param([string]$Tab)
    $tabs = @{
        'General'    = @($setTabGeneral,    $setPaneGeneral)
        'AD'         = @($setTabAD,         $setPaneAD)
        'Entra'      = @($setTabEntra,      $setPaneEntra)
        'Intune'     = @($setTabIntune,     $setPaneIntune)
        'SCCM'       = @($setTabSCCM,       $setPaneSCCM)
        'Appearance' = @($setTabAppearance, $setPaneAppearance)
    }
    $accent  = $Window.Resources['ThemeAccent']
    $primary = $Window.Resources['ThemeTextPrimary']
    $muted   = $Window.Resources['ThemeTextDim']
    foreach ($k in $tabs.Keys) {
        $btn  = $tabs[$k][0]
        $pane = $tabs[$k][1]
        if ($k -eq $Tab) {
            $btn.BorderBrush = $accent
            $btn.Foreground  = $primary
            $btn.FontWeight  = 'SemiBold'
            $pane.Visibility = 'Visible'
        } else {
            $btn.BorderBrush = [System.Windows.Media.Brushes]::Transparent
            $btn.Foreground  = $muted
            $btn.FontWeight  = 'Normal'
            $pane.Visibility = 'Collapsed'
        }
    }
}
function Hide-SettingsDialog {
    $pnlSettingsView.Visibility = 'Collapsed'
    $svMainContent.Visibility   = 'Visible'
}
function Show-SettingsDialog {
    $s = $Global:Settings
    $txtSetADServer.Text     = $s.AD.Server
    $txtSetADSearchBase.Text = $s.AD.SearchBase
    $chkSetADEnabled.IsChecked = [bool]$s.AD.Enabled
    $txtSetEntraTenant.Text  = $s.Entra.TenantId
    $chkSetEntraEnabled.IsChecked  = [bool]$s.Entra.Enabled
    $chkSetIntuneEnabled.IsChecked = [bool]$s.Intune.Enabled
    if ($chkSetAutopilotEnabled) { $chkSetAutopilotEnabled.IsChecked = [bool]$s.Autopilot.Enabled }
    $txtSetSCCMServer.Text   = $s.SCCM.Server
    $txtSetSCCMSite.Text     = $s.SCCM.SiteCode
    $chkSetSCCMEnabled.IsChecked  = [bool]$s.SCCM.Enabled
    $rbThemeLight.IsChecked  = ($s.Theme -eq 'Light')
    $rbThemeDark.IsChecked   = ($s.Theme -ne 'Light')
    if ($txtSetRecentDays) { $txtSetRecentDays.Text = [string]$s.RecentActivityDays }
    Set-SettingsTab 'General'
    $svMainContent.Visibility   = 'Collapsed'
    if ($pnlBulkView)         { $pnlBulkView.Visibility         = 'Collapsed' }
    if ($pnlHistoryView)      { $pnlHistoryView.Visibility      = 'Collapsed' }
    if ($pnlAchievementsView) { $pnlAchievementsView.Visibility = 'Collapsed' }
    $pnlSettingsView.Visibility = 'Visible'
}
function Apply-SettingsFromDialog {
    Write-DebugLog '[Settings] save clicked — validating + applying' -Level 'DEBUG'
    $siteCode = $txtSetSCCMSite.Text.Trim().ToUpper()
    if ($siteCode -and -not ($siteCode -match '^[A-Z0-9]{1,3}$')) {
        Write-DebugLog "[Settings] save aborted: invalid SCCM site code '$siteCode'" -Level 'DEBUG'
        Set-SettingsTab 'SCCM'
        Show-Toast 'SCCM site code must be 1-3 alphanumeric characters' -Type Error
        return
    }
    $Global:Settings.AD.Server        = $txtSetADServer.Text.Trim()
    $Global:Settings.AD.SearchBase    = $txtSetADSearchBase.Text.Trim()
    $Global:Settings.AD.Enabled       = [bool]$chkSetADEnabled.IsChecked
    $Global:Settings.Entra.TenantId   = $txtSetEntraTenant.Text.Trim()
    $Global:Settings.Entra.Enabled    = [bool]$chkSetEntraEnabled.IsChecked
    $Global:Settings.Intune.Enabled   = [bool]$chkSetIntuneEnabled.IsChecked
    if ($chkSetAutopilotEnabled) { $Global:Settings.Autopilot.Enabled = [bool]$chkSetAutopilotEnabled.IsChecked }
    $Global:Settings.SCCM.Server      = $txtSetSCCMServer.Text.Trim()
    $Global:Settings.SCCM.SiteCode    = $siteCode
    $Global:Settings.SCCM.Enabled     = [bool]$chkSetSCCMEnabled.IsChecked
    $Global:Settings.Theme            = if ($rbThemeLight.IsChecked) { 'Light' } else { 'Dark' }
    if ($txtSetRecentDays -and $txtSetRecentDays.Text.Trim()) {
        $daysVal = 7
        if ([int]::TryParse($txtSetRecentDays.Text.Trim(), [ref]$daysVal) -and $daysVal -ge 0) {
            $Global:Settings.RecentActivityDays = $daysVal
        }
    }
    Save-Settings
    & $ApplyTheme ($Global:Settings.Theme -eq 'Light')
    Apply-DefaultTargetSelection
    Hide-SettingsDialog
    Show-Toast 'Settings saved' -Type Success
    Write-DebugLog "[Settings] applied: AD.Enabled=$($Global:Settings.AD.Enabled) Entra.Enabled=$($Global:Settings.Entra.Enabled) Intune.Enabled=$($Global:Settings.Intune.Enabled) SCCM.Enabled=$($Global:Settings.SCCM.Enabled) Theme=$($Global:Settings.Theme)" -Level 'DEBUG'
}
function Apply-DefaultTargetSelection {
    $chkTargetAD.IsChecked       = [bool]$Global:Settings.AD.Enabled
    $chkTargetEntra.IsChecked    = [bool]$Global:Settings.Entra.Enabled
    $chkTargetIntune.IsChecked   = [bool]$Global:Settings.Intune.Enabled
    if ($chkTargetAutopilot) { $chkTargetAutopilot.IsChecked = [bool]$Global:Settings.Autopilot.Enabled }
    $chkTargetSCCM.IsChecked     = [bool]$Global:Settings.SCCM.Enabled
}

# ============================================================================
# EVENT WIRING
# ============================================================================
# Title bar — drag is handled automatically by WindowChrome (CaptionHeight=32);
# child elements set WindowChrome.IsHitTestVisibleInChrome=True to be clickable.
$btnMinimize.Add_Click({ $Window.WindowState = 'Minimized' })
$btnMaximize.Add_Click({
    $Window.WindowState = if ($Window.WindowState -eq 'Maximized') { 'Normal' } else { 'Maximized' }
})
$btnClose.Add_Click({ $Window.Close() })
$btnHelp.Add_Click({
    Show-ModalConfirm -Title 'Device Decommissioner' -Icon ([string][char]0xE897) -HideCancel `
        -Message ("Removes a device from AD, Entra ID, Intune and SCCM in one action.`n`n" +
                  "1. Enter the device hostname (or AzureAD ObjectId) and click Look up.`n" +
                  "2. Review the four discovery cards and uncheck any system you do NOT want to touch.`n" +
                  "3. Click Decommission and type the device name to confirm.`n`n" +
                  "AD/SCCM use stored DPAPI-encrypted credentials (per Windows user). " +
                  "Entra ID and Intune use interactive Microsoft Graph sign-in.") `
        -ConfirmLabel 'Got it' -OnConfirm { Hide-Modal }
})
$btnThemeToggle.Add_Click({
    $newLight = -not $Global:IsLightMode
    Write-DebugLog "[UI] theme toggle → $(if ($newLight) { 'Light' } else { 'Dark' })" -Level 'DEBUG'
    & $ApplyTheme $newLight
    $Global:Settings.Theme = if ($newLight) { 'Light' } else { 'Dark' }
    Save-Settings
    Unlock-Achievement 'theme_toggle'
})
$btnSettings.Add_Click({ Write-DebugLog '[UI] btnSettings clicked' -Level 'DEBUG'; Show-SettingsDialog })

# ============================================================================
# ENTRA AUTH TOOLBAR
# ============================================================================
# Required Graph scopes for the lookup + decommission paths.
$Global:RequiredGraphScopes = @('Device.ReadWrite.All','DeviceManagementManagedDevices.ReadWrite.All','DeviceManagementServiceConfig.ReadWrite.All')

function Update-EntraAuthIndicator {
    # Reads the cached MgContext silently (no interactive prompt) and updates the toolbar.
    $ctx = $null
    try { if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) { $ctx = Get-MgContext -ErrorAction SilentlyContinue } } catch {}
    if (-not $ctx) {
        $authDot.Fill         = $Window.Resources['ThemeError']
        $lblAuthStatus.Text   = 'Not signed in'
        $bdrTenantInfo.Visibility = 'Collapsed'
        $btnEntraSignIn.Visibility  = 'Visible'
        $btnEntraSignOut.Visibility = 'Collapsed'
        if ($pnlEntraCallout) { $pnlEntraCallout.Visibility = 'Visible' }
        Write-DebugLog 'EntraAuth: no cached MgContext' -Level 'DEBUG'
        return
    }
    $haveScopes = $true
    foreach ($s in $Global:RequiredGraphScopes) { if (-not ($ctx.Scopes -contains $s)) { $haveScopes = $false; break } }
    $account = if ($ctx.Account) { [string]$ctx.Account } else { '(unknown account)' }
    if ($haveScopes) {
        $authDot.Fill        = $Window.Resources['ThemeSuccess']
        $lblAuthStatus.Text  = $account
    } else {
        $authDot.Fill        = $Window.Resources['ThemeWarning']
        $lblAuthStatus.Text  = "$account (limited scopes)"
    }
    $tenantId = if ($ctx.TenantId) { [string]$ctx.TenantId } else { '' }
    if ($tenantId) {
        $lblTenantInfo.Text   = "Tenant: $tenantId"
        $bdrTenantInfo.Visibility = 'Visible'
    } else {
        $bdrTenantInfo.Visibility = 'Collapsed'
    }
    $btnEntraSignIn.Visibility  = 'Collapsed'
    $btnEntraSignOut.Visibility = 'Visible'
    if ($pnlEntraCallout) { $pnlEntraCallout.Visibility = if ($haveScopes) { 'Collapsed' } else { 'Visible' } }
    Write-DebugLog "EntraAuth: cached context for $account (scopes complete=$haveScopes)" -Level 'DEBUG'
}

$btnEntraSignIn.Add_Click({
    Write-DebugLog '[UI] btnEntraSignIn clicked - launching interactive Connect-MgGraph' -Level 'DEBUG'
    if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
        Show-Toast 'Microsoft.Graph PowerShell SDK is not installed. See README - Prerequisites.' -Type Error
        return
    }
    $btnEntraSignIn.IsEnabled        = $false
    if ($btnEntraSignInCallout) { $btnEntraSignInCallout.IsEnabled = $false }
    Set-Status -Text 'Signing in to Entra...' -State Busy
    $Window.Cursor = [System.Windows.Input.Cursors]::Wait
    # Pump the dispatcher so status + cursor updates render before the blocking Graph call.
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([action]{}, 'Background')
    # Connect-MgGraph MUST run on the UI thread — it pops a browser/WAM dialog.
    # Background runspaces can't show interactive prompts and their session state
    # is isolated, so the cached MgContext wouldn't be visible to lookup runspaces.
    try {
        $cp = @{ Scopes=$Global:RequiredGraphScopes; NoWelcome=$true; ErrorAction='Stop' }
        if ($Global:Settings.Entra.TenantId) { $cp.TenantId = $Global:Settings.Entra.TenantId }
        Connect-MgGraph @cp | Out-Null
        Show-Toast 'Signed in to Entra. Lookups will reuse this session.' -Type Success
        Set-Status -Text 'Ready' -State Idle
        Unlock-Achievement 'first_signin'
    } catch {
        Write-DebugLog "EntraAuth: Connect-MgGraph failed: $($_.Exception.Message)" -Level 'ERROR'
        Show-Toast "Sign-in failed: $($_.Exception.Message)" -Type Error
        Set-Status -Text 'Sign-in failed' -State Error
    } finally {
        $Window.Cursor = $null
        $btnEntraSignIn.IsEnabled = $true
        if ($btnEntraSignInCallout) { $btnEntraSignInCallout.IsEnabled = $true }
        Update-EntraAuthIndicator
    }
})

$btnEntraSignOut.Add_Click({
    Write-DebugLog '[UI] btnEntraSignOut clicked' -Level 'DEBUG'
    try {
        if (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }
        Show-Toast 'Signed out of Entra.' -Type Info
    } catch {
        Write-DebugLog "EntraAuth: Disconnect-MgGraph error: $($_.Exception.Message)" -Level 'WARN'
    }
    Update-EntraAuthIndicator
})

# Mirror the toolbar Sign-In button (same handler) on the inline callout card
if ($btnEntraSignInCallout) {
    $btnEntraSignInCallout.Add_Click({
        $btnEntraSignIn.RaiseEvent((New-Object System.Windows.RoutedEventArgs ([System.Windows.Controls.Button]::ClickEvent)))
    })
}

# ============================================================================
# RAIL + SIDEBAR (Phase 4)
# ============================================================================
function Refresh-RecentDevicesList {
    if (-not $lstRecentDevices) { return }
    $lstRecentDevices.Items.Clear()
    $items = @($Global:RecentDevices | Where-Object { $_ })
    foreach ($name in $items) { [void]$lstRecentDevices.Items.Add([string]$name) }
    $lblRecentEmpty.Visibility = if ($items.Count -eq 0) { 'Visible' } else { 'Collapsed' }
}

function Add-RecentDevice {
    param([string]$Name)
    if (-not $Name) { return }
    $trim = $Name.Trim()
    if (-not $trim) { return }
    $existing = @($Global:RecentDevices | Where-Object { $_ -and ([string]$_).ToLowerInvariant() -ne $trim.ToLowerInvariant() })
    $merged = ,$trim + $existing
    if ($merged.Count -gt 20) { $merged = $merged[0..19] }
    $Global:RecentDevices = @($merged)
    Save-RecentDevices
    Refresh-RecentDevicesList
}

function Apply-SidebarVisibility {
    if (-not $colSidebar) { return }
    if ($Global:Settings.SidebarVisible) {
        $colSidebar.Width = New-Object System.Windows.GridLength 260
        $pnlSidebar.Visibility = 'Visible'
    } else {
        $colSidebar.Width = New-Object System.Windows.GridLength 0
        $pnlSidebar.Visibility = 'Collapsed'
    }
}

function Apply-LogPanelVisibility {
    if (-not $rowLog) { return }
    if ($Global:Settings.LogPanelVisible) {
        $rowSplitter.Height = New-Object System.Windows.GridLength 4
        $rowLog.Height      = New-Object System.Windows.GridLength 160
        if ($btnRestoreLog) { $btnRestoreLog.Visibility = 'Collapsed' }
    } else {
        $rowSplitter.Height = New-Object System.Windows.GridLength 0
        $rowLog.Height      = New-Object System.Windows.GridLength 0
        if ($btnRestoreLog) { $btnRestoreLog.Visibility = 'Visible' }
    }
}

if ($btnToggleLog) {
    $btnToggleLog.Add_Click({
        $Global:Settings.LogPanelVisible = $false
        Apply-LogPanelVisibility
        Save-Settings
    })
}
if ($btnRestoreLog) {
    $btnRestoreLog.Add_Click({
        $Global:Settings.LogPanelVisible = $true
        Apply-LogPanelVisibility
        Save-Settings
    })
}

# ============================================================================
# PREREQUISITE DETECTION & INSTALL
# ============================================================================
function Update-PrereqBanner {
    if (-not $pnlPrereqs) { return }
    $missing = New-Object System.Collections.Generic.List[string]
    $graphMissing = $false
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory))                              { $missing.Add('ActiveDirectory (RSAT)') | Out-Null }
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Identity.DirectoryManagement)) { $missing.Add('Microsoft.Graph.Identity.DirectoryManagement') | Out-Null; $graphMissing = $true }
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.DeviceManagement))             { $missing.Add('Microsoft.Graph.DeviceManagement')             | Out-Null; $graphMissing = $true }
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.DeviceManagement.Enrollment))  { $missing.Add('Microsoft.Graph.DeviceManagement.Enrollment')  | Out-Null; $graphMissing = $true }
    $sccmModule = "${env:ProgramFiles(x86)}\Microsoft Endpoint Manager\AdminConsole\bin\ConfigurationManager.psd1"
    $sccmModuleAlt = if ($env:SMS_ADMIN_UI_PATH) { Join-Path $env:SMS_ADMIN_UI_PATH '..\ConfigurationManager.psd1' } else { $null }
    if (-not (Test-Path $sccmModule) -and -not ($sccmModuleAlt -and (Test-Path $sccmModuleAlt))) {
        if ($Global:Settings.SCCM.Enabled) { $missing.Add('ConfigurationManager (SCCM admin console)') | Out-Null }
    }
    if ($missing.Count -eq 0) {
        $pnlPrereqs.Visibility = 'Collapsed'
        return
    }
    $lblPrereqsDetail.Text  = 'Missing: ' + ($missing -join ', ') + '. Lookups against these systems will be skipped until the modules are installed.'
    $btnInstallGraph.Visibility = if ($graphMissing) { 'Visible' } else { 'Collapsed' }
    $pnlPrereqs.Visibility = 'Visible'
}

if ($btnInstallGraph) {
    $btnInstallGraph.Add_Click({
        Write-DebugLog '[UI] btnInstallGraph clicked - Install-Module Microsoft.Graph (CurrentUser)' -Level 'INFO'
        $btnInstallGraph.IsEnabled = $false
        Set-Status -Text 'Installing Microsoft.Graph (CurrentUser)...' -State Busy
        Show-Toast 'Installing Microsoft.Graph (this can take a couple of minutes)' -Type Info
        Start-BackgroundWork -Variables @{} -Work {
            try {
                if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
                    Install-PackageProvider -Name NuGet -Force -Scope CurrentUser -ErrorAction Stop | Out-Null
                }
                Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Install-Module Microsoft.Graph.DeviceManagement             -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Install-Module Microsoft.Graph.DeviceManagement.Enrollment  -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                [pscustomobject]@{ Success=$true; Error=$null }
            } catch {
                [pscustomobject]@{ Success=$false; Error=$_.Exception.Message }
            }
        } -OnComplete {
            param($Result)
            $btnInstallGraph.IsEnabled = $true
            if ($Result -and $Result.Success) {
                Show-Toast 'Microsoft.Graph installed successfully' -Type Success
                Set-Status -Text 'Ready' -State Idle
            } else {
                $msg = if ($Result) { $Result.Error } else { 'Unknown error' }
                Show-Toast "Install failed: $msg" -Type Error
                Set-Status -Text 'Install failed' -State Error
            }
            Update-PrereqBanner
        }
    })
}
if ($btnPrereqHelp) {
    $btnPrereqHelp.Add_Click({
        $cmds = @(
            [pscustomobject]@{
                Label       = 'Microsoft Graph SDK (no admin required)'
                Description = "Run as the same user that runs this tool. Uses your existing PSGallery trust. If your environment blocks PSGallery, use Save-Module on a connected machine and copy the modules to a path under `$env:PSModulePath."
                Command     = 'Install-Module Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.DeviceManagement, Microsoft.Graph.DeviceManagement.Enrollment -Scope CurrentUser -AllowClobber -Force'
            }
            [pscustomobject]@{
                Label       = 'Active Directory (RSAT) — needs admin'
                Description = "DISM is the most reliable on Windows 10/11. Run from an elevated PowerShell."
                Command     = 'Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"'
            }
            [pscustomobject]@{
                Label       = 'SCCM Console (ConfigurationManager module) — admin install'
                Description = "Download the Configuration Manager console MSI from your site server's \\\\<SiteServer>\\SMS_<SiteCode>\\Tools\\ConsoleSetup share, run ConsoleSetup.exe, then point Settings → SCCM at your site code."
                Command     = '\\<SiteServer>\SMS_<SiteCode>\Tools\ConsoleSetup\ConsoleSetup.exe'
            }
        )
        Show-ModalConfirm -Title 'Installing prerequisites' -Icon ([string][char]0xE7BA) -HideCancel `
            -Message 'Copy any of the commands below and run them in a regular (or elevated, where noted) PowerShell prompt.' `
            -Commands $cmds `
            -ConfirmLabel 'Got it' -OnConfirm { Hide-Modal }
    })
}

# Export audit trail as CSV
if ($btnExportAudit) {
    $btnExportAudit.Add_Click({
        try {
            # Force-refresh from disk (the user may have manually appended/copied the file)
            $Global:AuditHistoryCache = $null
            $raw = Read-AuditHistory
            if (-not $raw -or @($raw).Count -eq 0) {
                Show-Toast 'Audit trail is empty. Run a decommission first.' -Type Warning
                return
            }
            $dlg = New-Object Microsoft.Win32.SaveFileDialog
            $dlg.Title  = 'Export audit trail'
            $dlg.Filter = 'CSV files (*.csv)|*.csv'
            $dlg.FileName = "decommission-history_$(Get-Date -Format 'yyyyMMdd').csv"
            $dlg.InitialDirectory = $Global:Root
            if ($dlg.ShowDialog() -ne $true) { return }
            $rows = @()
            foreach ($entry in @($raw)) {
                $systems = if ($entry.Targets) { ($entry.Targets -join ';') } else { '' }
                $resultParts = @()
                if ($entry.Results) {
                    foreach ($prop in $entry.Results.PSObject.Properties) {
                        $v = $prop.Value
                        $icon = if ($v.Success) { 'OK' } else { 'FAIL' }
                        $resultParts += "$($prop.Name)=$icon"
                    }
                }
                $rows += [pscustomobject]@{
                    Timestamp = $entry.Timestamp
                    Operator  = $entry.Operator
                    Machine   = $entry.Machine
                    Device    = $entry.Device
                    DryRun    = $entry.DryRun
                    Systems   = $systems
                    Results   = ($resultParts -join '; ')
                }
            }
            $rows | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
            Show-Toast "Exported $($rows.Count) entries to CSV" -Type Success
            Write-DebugLog "[Settings] audit trail exported to $($dlg.FileName)" -Level 'INFO'
            Unlock-Achievement 'exported_audit'
        } catch {
            Show-Toast "Export failed: $($_.Exception.Message)" -Type Error
            Write-DebugLog "[Settings] audit export error: $($_.Exception.Message)" -Level 'ERROR'
        }
    })
}

$btnRailHamburger.Add_Click({
    $Global:Settings.SidebarVisible = -not [bool]$Global:Settings.SidebarVisible
    Apply-SidebarVisibility
    Save-Settings
    Write-DebugLog "[UI] Sidebar toggled (visible=$($Global:Settings.SidebarVisible))" -Level 'DEBUG'
})
$btnRailDevice.Add_Click({
    # Home: dismiss any view-swap, return to the main decommission flow.
    if ($pnlSettingsView -and $pnlSettingsView.Visibility -eq 'Visible') { Hide-SettingsDialog }
    if ($pnlBulkView    -and $pnlBulkView.Visibility    -eq 'Visible') { Hide-BulkView }
    if ($pnlHistoryView -and $pnlHistoryView.Visibility -eq 'Visible') { Hide-HistoryView }
    if ($pnlAchievementsView -and $pnlAchievementsView.Visibility -eq 'Visible') { Hide-AchievementsView }
    $svMainContent.Visibility = 'Visible'
    if ($txtDeviceId) { $txtDeviceId.Focus() | Out-Null }
})
$btnRailSettings.Add_Click({ Write-DebugLog '[UI] btnRailSettings clicked' -Level 'DEBUG'; Show-SettingsDialog })
$btnRailHelp.Add_Click({ if ($btnHelp -and $btnHelp.Command) { $btnHelp.Command.Execute($null) } else { $btnHelp.RaiseEvent((New-Object System.Windows.RoutedEventArgs ([System.Windows.Controls.Button]::ClickEvent))) } })

# ============================================================================
# BULK LOOKUP
# ============================================================================
function Show-BulkView {
    $svMainContent.Visibility   = 'Collapsed'
    $pnlSettingsView.Visibility = 'Collapsed'
    if ($pnlHistoryView) { $pnlHistoryView.Visibility = 'Collapsed' }
    if ($pnlAchievementsView) { $pnlAchievementsView.Visibility = 'Collapsed' }
    if ($pnlBulkView) { $pnlBulkView.Visibility = 'Visible' }
}
function Hide-BulkView {
    if ($pnlBulkView) { $pnlBulkView.Visibility = 'Collapsed' }
    $svMainContent.Visibility   = 'Visible'
}

if ($btnRailBulk) {
    $btnRailBulk.Add_Click({ Show-BulkView })
}
if ($btnBulkBack) {
    $btnBulkBack.Add_Click({ Hide-BulkView })
}
# Bulk input placeholder hint — hide when user types
if ($txtBulkInput -and $lblBulkInputHint) {
    $txtBulkInput.Add_TextChanged({
        $lblBulkInputHint.Visibility = if ([string]::IsNullOrEmpty($txtBulkInput.Text)) { 'Visible' } else { 'Collapsed' }
    })
}
if ($btnBulkRun -and $dgBulkResults) {
    $btnBulkRun.Add_Click({
        $lines = @($txtBulkInput.Text -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        if ($lines.Count -eq 0) { Show-Toast 'Paste hostnames first (one per line)' -Type Warning; return }
        if ($lines.Count -gt 100) { Show-Toast 'Max 100 devices per batch' -Type Warning; return }
        Write-DebugLog "[BulkLookup] starting for $($lines.Count) device(s)" -Level 'INFO'
        Unlock-Achievement 'first_inventory'
        $btnBulkRun.IsEnabled = $false
        Set-Status -Text 'Bulk lookup in progress' -Detail "$($lines.Count) device(s)" -State Busy
        $dgBulkResults.ItemsSource = $null

        # Build a list of jobs — one per device, all systems queried
        $settings = $Global:Settings
        $credAD   = $Global:Creds.AD
        Start-BackgroundWork -Variables @{
            Devices=$lines; ADServer=$settings.AD.Server; ADSearchBase=$settings.AD.SearchBase; CredAD=$credAD
            EntraTenantId=$settings.Entra.TenantId
        } -Work {
            $matrix = New-Object System.Collections.ArrayList
            foreach ($dev in $Devices) {
                $row = @{ Device=$dev; AD='-'; Entra='-'; Intune='-'; Autopilot='-'; SCCM='-' }

                # AD
                try {
                    if (Get-Module -ListAvailable -Name ActiveDirectory) {
                        Import-Module ActiveDirectory -ErrorAction Stop
                        $p = @{ ErrorAction='Stop' }
                        if ($ADServer)     { $p.Server     = $ADServer }
                        if ($ADSearchBase) { $p.SearchBase = $ADSearchBase }
                        if ($CredAD)       { $p.Credential = $CredAD }
                        $q = $dev.Replace("'","''")
                        $p.Filter = "Name -eq '$q' -or sAMAccountName -eq '$q' -or sAMAccountName -eq '$($q)$'"
                        $obj = Get-ADComputer @p -Properties Name | Select-Object -First 1
                        $row.AD = if ($obj) { 'Found' } else { 'Not found' }
                    } else { $row.AD = 'N/A' }
                } catch { $row.AD = 'Error' }

                # Entra + Intune + Autopilot (Graph)
                try {
                    $ctx = $null; try { $ctx = Get-MgContext -ErrorAction SilentlyContinue } catch {}
                    if ($ctx) {
                        # Entra
                        try {
                            Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
                            $d = @(Get-MgDevice -Filter "displayName eq '$dev'" -All -ErrorAction Stop)
                            $row.Entra = if ($d.Count -gt 0) { 'Found' } else { 'Not found' }
                        } catch { $row.Entra = 'Error' }
                        # Intune
                        try {
                            Import-Module Microsoft.Graph.DeviceManagement -ErrorAction Stop
                            $m = @(Get-MgDeviceManagementManagedDevice -Filter "deviceName eq '$dev'" -All -ErrorAction Stop)
                            $row.Intune = if ($m.Count -gt 0) { 'Found' } else { 'Not found' }
                        } catch { $row.Intune = 'Error' }
                        # Autopilot
                        try {
                            Import-Module Microsoft.Graph.DeviceManagement.Enrollment -ErrorAction Stop
                            $ap = @(Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -All -Top 200 -ErrorAction Stop) | Where-Object { $_.DisplayName -eq $dev -or $_.SerialNumber -eq $dev }
                            $row.Autopilot = if ($ap) { 'Found' } else { 'Not found' }
                        } catch { $row.Autopilot = 'Error' }
                    } else {
                        $row.Entra = 'Sign in'; $row.Intune = 'Sign in'; $row.Autopilot = 'Sign in'
                    }
                } catch { $row.Entra = 'Error'; $row.Intune = 'Error'; $row.Autopilot = 'Error' }

                [void]$matrix.Add([pscustomobject]$row)
            }
            return $matrix
        } -OnComplete {
            param($Results,$Errors)
            $btnBulkRun.IsEnabled = $true
            if ($Results -and $Results.Count -gt 0) {
                $data = @($Results[0])
                $dgBulkResults.ItemsSource = $data
                if ($pnlBulkEmpty) { $pnlBulkEmpty.Visibility = if ($data.Count -eq 0) { 'Visible' } else { 'Collapsed' } }
                $found = ($data | Where-Object { $_.AD -eq 'Found' -or $_.Entra -eq 'Found' -or $_.Intune -eq 'Found' -or $_.Autopilot -eq 'Found' }).Count
                Set-Status -Text 'Inventory check complete' -Detail "$($data.Count) device(s), $found with matches" -State Success
                Write-DebugLog "[Inventory] complete: $($data.Count) rows, $found with matches" -Level 'SUCCESS'
            } else {
                Set-Status -Text 'Inventory check failed' -State Error
                Show-Toast 'Inventory check returned no results' -Type Error
            }
        }
    })
}
if ($btnBulkCopy -and $dgBulkResults) {
    $btnBulkCopy.Add_Click({
        $data = $dgBulkResults.ItemsSource
        if (-not $data) { Show-Toast 'No results to copy' -Type Warning; return }
        $lines = @('Device`tAD`tEntra`tIntune`tAutopilot`tSCCM')
        foreach ($r in $data) { $lines += "$($r.Device)`t$($r.AD)`t$($r.Entra)`t$($r.Intune)`t$($r.Autopilot)`t$($r.SCCM)" }
        [System.Windows.Clipboard]::SetText(($lines -join "`n"))
        Show-Toast "Copied $($data.Count) rows (tab-separated)" -Type Info
    })
}

# ============================================================================
# DECOMMISSION HISTORY VIEW
# ============================================================================
$Global:HistoryRows = @()

function Show-HistoryView {
    $svMainContent.Visibility   = 'Collapsed'
    $pnlSettingsView.Visibility = 'Collapsed'
    if ($pnlBulkView)         { $pnlBulkView.Visibility         = 'Collapsed' }
    if ($pnlAchievementsView) { $pnlAchievementsView.Visibility = 'Collapsed' }
    if ($pnlHistoryView) { $pnlHistoryView.Visibility = 'Visible' }
    Refresh-HistoryView
}
function Hide-HistoryView {
    if ($pnlHistoryView) { $pnlHistoryView.Visibility = 'Collapsed' }
    $svMainContent.Visibility = 'Visible'
}

function Show-AchievementsView {
    $svMainContent.Visibility   = 'Collapsed'
    $pnlSettingsView.Visibility = 'Collapsed'
    if ($pnlBulkView)    { $pnlBulkView.Visibility    = 'Collapsed' }
    if ($pnlHistoryView) { $pnlHistoryView.Visibility = 'Collapsed' }
    if ($pnlAchievementsView) { $pnlAchievementsView.Visibility = 'Visible' }
    Render-Achievements
}
function Hide-AchievementsView {
    if ($pnlAchievementsView) { $pnlAchievementsView.Visibility = 'Collapsed' }
    $svMainContent.Visibility = 'Visible'
}

function ConvertTo-HistorySymbol {
    param($systemResult)
    if (-not $systemResult) { return '-' }
    if ($systemResult.Success) { return [string][char]0x2713 }   # ✓
    return [string][char]0x2717                                    # ✗
}

function Refresh-HistoryView {
    <#
    .SYNOPSIS
        Reloads the History DataGrid from decommission-history.json. Forces a
        cache refresh by invalidating $Global:AuditHistoryCache so the user
        gets the latest data even if another process touched the file.
    #>
    if (-not $dgHistory) { return }
    # Force the next Read-AuditHistory to re-parse from disk
    $Global:AuditHistoryCache = $null
    $rawEntries = Read-AuditHistory
    $rows = New-Object System.Collections.ArrayList
    foreach ($entry in $rawEntries) {
        $when = ''
        try { $when = ([datetime]::Parse($entry.Timestamp)).ToString('yyyy-MM-dd HH:mm:ss') } catch { $when = [string]$entry.Timestamp }
        $r = $entry.Results
        [void]$rows.Add([pscustomobject]@{
            When      = $when
            Device    = if ([string]::IsNullOrWhiteSpace([string]$entry.Device)) { '(unknown)' } else { [string]$entry.Device }
            Operator  = [string]$entry.Operator
            Mode      = if ($entry.DryRun) { 'Dry-run' } else { 'Real' }
            AD        = ConvertTo-HistorySymbol $r.AD
            Entra     = ConvertTo-HistorySymbol $r.Entra
            Intune    = ConvertTo-HistorySymbol $r.Intune
            Autopilot = ConvertTo-HistorySymbol $r.Autopilot
            SCCM      = ConvertTo-HistorySymbol $r.SCCM
            IsDryRun  = [bool]$entry.DryRun
        })
    }
    # Newest first
    $Global:HistoryRows = @($rows | Sort-Object -Property When -Descending)
    Apply-HistoryFilter
}

function Apply-HistoryFilter {
    if (-not $dgHistory) { return }
    $filter = if ($txtHistoryFilter) { $txtHistoryFilter.Text.Trim() } else { '' }
    $hideDry = if ($chkHistoryHideDryRun) { [bool]$chkHistoryHideDryRun.IsChecked } else { $false }
    $filtered = $Global:HistoryRows
    if ($hideDry) { $filtered = @($filtered | Where-Object { -not $_.IsDryRun }) }
    if ($filter) {
        $filtered = @($filtered | Where-Object {
            $_.Device -like "*$filter*" -or $_.Operator -like "*$filter*" -or $_.When -like "*$filter*"
        })
    }
    $dgHistory.ItemsSource = $filtered
    # Empty-state overlay
    if ($pnlHistoryEmpty) {
        $pnlHistoryEmpty.Visibility = if (@($filtered).Count -eq 0 -and @($Global:HistoryRows).Count -eq 0) { 'Visible' } else { 'Collapsed' }
    }
    # Filter hint visibility
    if ($lblHistoryFilterHint -and $txtHistoryFilter) {
        $lblHistoryFilterHint.Visibility = if ([string]::IsNullOrEmpty($txtHistoryFilter.Text)) { 'Visible' } else { 'Collapsed' }
    }
    if ($lblHistoryDetail) {
        $totalCount = @($Global:HistoryRows).Count
        $shownCount = @($filtered).Count
        if ($totalCount -eq 0) {
            $lblHistoryDetail.Text = 'No decommissions recorded yet.'
        } elseif ($shownCount -eq $totalCount) {
            $lblHistoryDetail.Text = "$totalCount entries from $Global:AuditFile"
        } else {
            $lblHistoryDetail.Text = "$shownCount of $totalCount entries shown (filtered)"
        }
    }
}

if ($btnRailHistory) { $btnRailHistory.Add_Click({ Show-HistoryView }) }
if ($btnHistoryBack) { $btnHistoryBack.Add_Click({ Hide-HistoryView }) }
if ($btnRailAchievements) { $btnRailAchievements.Add_Click({ Show-AchievementsView }) }
if ($btnAchievementsBack) { $btnAchievementsBack.Add_Click({ Hide-AchievementsView }) }
if ($btnHistoryRefresh) { $btnHistoryRefresh.Add_Click({ Refresh-HistoryView }) }
if ($txtHistoryFilter) {
    $txtHistoryFilter.Add_TextChanged({ Apply-HistoryFilter })
}
if ($chkHistoryHideDryRun) {
    $chkHistoryHideDryRun.Add_Checked({   Apply-HistoryFilter })
    $chkHistoryHideDryRun.Add_Unchecked({ Apply-HistoryFilter })
}
if ($btnHistoryLookup -and $dgHistory) {
    $btnHistoryLookup.Add_Click({
        $sel = $dgHistory.SelectedItem
        if (-not $sel -or -not $sel.Device) { Show-Toast 'Select a row first' -Type Warning; return }
        $name = [string]$sel.Device
        Hide-HistoryView
        $txtDeviceId.Text = $name
        Start-DeviceLookup -Query $name
    })
    # Double-click row also triggers lookup
    $dgHistory.Add_MouseDoubleClick({
        $sel = $dgHistory.SelectedItem
        if (-not $sel -or -not $sel.Device) { return }
        $name = [string]$sel.Device
        Hide-HistoryView
        $txtDeviceId.Text = $name
        Start-DeviceLookup -Query $name
    })
}

$lstRecentDevices.Add_SelectionChanged({
    if ($lstRecentDevices.SelectedItem) {
        $name = [string]$lstRecentDevices.SelectedItem
        Write-DebugLog "[UI] Recent device clicked: $name" -Level 'DEBUG'
        $txtDeviceId.Text = $name
        # Defer SelectedIndex reset and lookup kickoff so click visual settles before the lookup locks UI.
        $lstRecentDevices.SelectedIndex = -1
        Start-DeviceLookup -Query $name
    }
})

$btnClearRecent.Add_Click({
    if ($Global:RecentDevices.Count -eq 0) { return }
    $Global:RecentDevices = @()
    Save-RecentDevices
    Refresh-RecentDevicesList
    Show-Toast 'Recent devices cleared' -Type Info
})

# Verbose DEBUG-level logging is on by default; can be turned off in Settings
# if needed. The previous title-bar toggle was removed (always-on is more useful
# for an alpha tool — operators want to see what happened when something fails).
$Global:DebugOverlayEnabled = $true

# Device-ID input watermark / placeholder
if ($lblDeviceIdHint) {
    $updateHint = { $lblDeviceIdHint.Visibility = if ([string]::IsNullOrEmpty($txtDeviceId.Text)) { 'Visible' } else { 'Collapsed' } }
    $txtDeviceId.Add_TextChanged($updateHint)
    & $updateHint
}

# (Log panel buttons are wired further below alongside the other UI handlers.)

# Lookup
$btnLookup.Add_Click({ Write-DebugLog '[UI] btnLookup clicked' -Level 'DEBUG'; Start-DeviceLookup -Query $txtDeviceId.Text.Trim() })
if ($btnCancelLookup) {
    $btnCancelLookup.Add_Click({
        if (-not $Global:LookupInFlight) { return }
        Write-DebugLog '[UI] btnCancelLookup clicked - cancelling in-flight lookup' -Level 'WARN'
        $btnCancelLookup.IsEnabled = $false
        # Bump generation so any straggling background result is dropped by Save-LookupResult.
        $Global:LookupGen = [int]$Global:LookupGen + 1
        if ($Global:LookupTimeoutTimer) { try { $Global:LookupTimeoutTimer.Stop() } catch {} }
        # Best-effort runspace shutdown (the runspace will keep running its current cmdlet, but its result is now ignored).
        foreach ($job in @($Global:BgJobs)) {
            try { $job.PS.Stop() } catch {}
        }
        foreach ($s in $Script:AllSystems) {
            if (-not $Global:LookupState.Results[$s]) { Set-CardStatus -System $s -State Warning -Detail 'Cancelled.' }
        }
        Hide-GlobalProgress
        Set-Status -Text 'Lookup cancelled' -State Warning
        $Global:LookupInFlight = $false
        Set-LookupButtonsIdle
        Show-Toast 'Lookup cancelled' -Type Info
        Unlock-Achievement 'cancelled_lookup'
    })
}
$txtDeviceId.Add_KeyDown({
    if ($_.Key -eq 'Return') { Write-DebugLog '[UI] Enter pressed in device input' -Level 'DEBUG'; Start-DeviceLookup -Query $txtDeviceId.Text.Trim() }
})
$btnClearLookup.Add_Click({
    # Invalidate any in-flight lookup callbacks (generation bump) and release the UI lock.
    $Global:LookupGen = [int]$Global:LookupGen + 1
    $Global:LookupInFlight = $false
    if ($Global:LookupTimeoutTimer) { try { $Global:LookupTimeoutTimer.Stop() } catch {} }
    Set-LookupButtonsIdle
    Hide-GlobalProgress
    $txtDeviceId.Clear()
    Reset-AllCards
    Set-Status -Text 'Ready' -State Idle
})
$btnRefresh.Add_Click({
    if ($Global:LookupState.DeviceQuery) { Write-DebugLog "[UI] btnRefresh clicked (re-running '$($Global:LookupState.DeviceQuery)')" -Level 'DEBUG'; Start-DeviceLookup -Query $Global:LookupState.DeviceQuery }
})
foreach ($s in $Script:AllSystems) {
    $chk = Get-Variable -Name "chkTarget$s" -ValueOnly -Scope Script
    $chk.Add_Checked({   Update-DecommissionButton })
    $chk.Add_Unchecked({ Update-DecommissionButton })
}

# Copy result buttons (#7) — copies the detail text from the matching card
foreach ($s in $Script:AllSystems) {
    $btn = Get-Variable -Name "btnCopy$s" -ValueOnly -Scope Script -ErrorAction SilentlyContinue
    if ($btn) {
        $sysCapture = $s
        $btn.Add_Click({
            $det = Get-Variable -Name "lbl${sysCapture}Detail" -ValueOnly -Scope Script
            if ($det -and $det.Text) {
                [System.Windows.Clipboard]::SetText($det.Text)
                Show-Toast "$sysCapture result copied" -Type Info -DurationMs 1500
            }
        }.GetNewClosure())
    }
}
$btnDecommission.Add_Click({ Start-Decommission })
$chkDryRun.Add_Checked({   Update-DecommissionButton })
$chkDryRun.Add_Unchecked({ Update-DecommissionButton })

# Credentials
$btnEditCredsAD.Add_Click({   Write-DebugLog '[UI] btnEditCredsAD clicked' -Level 'DEBUG'; Edit-StoredCreds -Target 'AD' })
$btnEditCredsSCCM.Add_Click({ Write-DebugLog '[UI] btnEditCredsSCCM clicked' -Level 'DEBUG'; Edit-StoredCreds -Target 'SCCM' })

# Modal buttons (single shared handler)
$btnModalConfirm.Add_Click({
    if ($Global:ModalCallback) {
        try { & $Global:ModalCallback }
        catch {
            Write-DebugLog "Modal callback error: $($_.Exception.Message)" -Level 'ERROR'
            Hide-Modal
        }
    } else { Hide-Modal }
})
$btnModalCancel.Add_Click({
    # If the operator cancels a confirm modal that had safety warnings shown, that's a 'heeded_warning' moment.
    if ($Global:LastWarnedDevice -and $pnlModalWarnings -and $pnlModalWarnings.ItemsSource) {
        Unlock-Achievement 'heeded_warning'
    }
    Hide-Modal
})

# Settings view (in-app tabbed)
$btnSettingsCancel.Add_Click({ Hide-SettingsDialog })
$btnSettingsBack.Add_Click({ Hide-SettingsDialog })
$btnSettingsSave.Add_Click({ Apply-SettingsFromDialog })
$setTabGeneral.Add_Click({    Set-SettingsTab 'General' })
$setTabAD.Add_Click({         Set-SettingsTab 'AD' })
$setTabEntra.Add_Click({      Set-SettingsTab 'Entra' })
$setTabIntune.Add_Click({     Set-SettingsTab 'Intune' })
$setTabSCCM.Add_Click({       Set-SettingsTab 'SCCM' })
$setTabAppearance.Add_Click({ Set-SettingsTab 'Appearance' })

# Click-outside-to-dismiss for the modal overlay
$pnlModalOverlay.Add_MouseLeftButtonDown({
    if ($_.OriginalSource -eq $pnlModalOverlay) { Hide-Modal }
})

# Enter / Escape inside modal text + password inputs
$txtModalInput.Add_KeyDown({
    if ($_.Key -eq 'Return') {
        if ($Global:ModalCallback) {
            try { & $Global:ModalCallback }
            catch { Write-DebugLog "Modal callback error: $($_.Exception.Message)" -Level 'ERROR'; Hide-Modal }
        }
        $_.Handled = $true
    } elseif ($_.Key -eq 'Escape') { Hide-Modal; $_.Handled = $true }
})
$pwdModalInput.Add_KeyDown({
    if ($_.Key -eq 'Return') {
        if ($Global:ModalCallback) {
            try { & $Global:ModalCallback }
            catch { Write-DebugLog "Modal callback error: $($_.Exception.Message)" -Level 'ERROR'; Hide-Modal }
        }
        $_.Handled = $true
    } elseif ($_.Key -eq 'Escape') { Hide-Modal; $_.Handled = $true }
})

# Global Escape — close any visible overlay
$Window.Add_PreviewKeyDown({
    if ($_.Key -ne 'Escape') { return }
    if ($pnlModalOverlay.Visibility -eq 'Visible') { Hide-Modal; $_.Handled = $true; return }
    if ($pnlSettingsView.Visibility  -eq 'Visible') { Hide-SettingsDialog; $_.Handled = $true }
})

# Log panel buttons
$btnClearLog.Add_Click({
    $paraLog.Inlines.Clear()
    $Global:DebugLineCount = 0
})
$btnCopyLog.Add_Click({
    $tr = New-Object System.Windows.Documents.TextRange ($rtbLog.Document.ContentStart, $rtbLog.Document.ContentEnd)
    [System.Windows.Clipboard]::SetText($tr.Text)
    Show-Toast 'Log copied to clipboard' -Type Info
})

# ============================================================================
# BACKGROUND POLLING TIMER
# ============================================================================
$pollTimer = New-Object System.Windows.Threading.DispatcherTimer
$pollTimer.Interval = [TimeSpan]::FromMilliseconds($Script:TIMER_INTERVAL_MS)
$pollTimer.Add_Tick({
    if ($Global:BgJobs.Count -eq 0) { return }
    for ($i = $Global:BgJobs.Count - 1; $i -ge 0; $i--) {
        $job = $Global:BgJobs[$i]
        if ($job.AsyncResult.IsCompleted) {
            $elapsed = [math]::Round(((Get-Date) - $job.StartedAt).TotalSeconds, 1)
            Write-DebugLog "BgWork: job #$($i) completed after ${elapsed}s (label=$($job.Label))" -Level 'INFO'
            $results = @(); $errs = @()
            try { $results = @($job.PS.EndInvoke($job.AsyncResult)) } catch { $errs += $_ }
            if ($job.PS.Streams.Error.Count -gt 0) {
                Write-DebugLog "BgWork: job #$($i) had $($job.PS.Streams.Error.Count) stream error(s): $($job.PS.Streams.Error[0])" -Level 'WARN'
                $errs += @($job.PS.Streams.Error)
            }
            try { & $job.OnComplete $results $errs } catch {
                Write-DebugLog "OnComplete handler error: $($_.Exception.Message)`n$($_.ScriptStackTrace)" -Level 'ERROR'
            }
            try { $job.PS.Dispose() } catch {}
            try { $job.Runspace.Dispose() } catch {}
            $Global:BgJobs.RemoveAt($i)
        }
    }
})

# ============================================================================
# INITIAL STATE
# ============================================================================
Load-Settings
Load-RecentDevices
Load-Creds
Load-Achievements
& $ApplyTheme ($Global:Settings.Theme -eq 'Light')
Apply-DefaultTargetSelection
Update-CredIndicators
Reset-AllCards
Set-Status -Text 'Ready' -State Idle
Update-EntraAuthIndicator   # silent cache check (no interactive prompt at launch)
Refresh-RecentDevicesList
Apply-SidebarVisibility
Apply-LogPanelVisibility
Update-PrereqBanner
Write-DebugLog "$Global:AppTitle v$Global:AppVersion started" -Level 'SUCCESS'
Write-DebugLog "Settings: $Global:SettingsFile" -Level 'DEBUG'
Write-DebugLog "Credentials: $Global:CredsFile" -Level 'DEBUG'
Write-DebugLog "Log file: $Global:DebugLogFile" -Level 'DEBUG'

$pollTimer.Start()
$Window.Add_Closing({
    if ($Global:LookupInFlight -or ($Global:BgJobs -and $Global:BgJobs.Count -gt 0)) {
        $res = [System.Windows.MessageBox]::Show(
            'A lookup or decommission is still in progress. Close anyway?',
            'Device Decommissioner', 'YesNo', 'Warning')
        if ($res -ne 'Yes') { $_.Cancel = $true }
    }
})
$Window.Add_Closed({
    try { $pollTimer.Stop() } catch {}
    try { if ($Global:LookupTimeoutTimer) { $Global:LookupTimeoutTimer.Stop() } } catch {}
    foreach ($job in @($Global:BgJobs)) {
        try { if (-not $job.AsyncResult.IsCompleted) { $job.PS.Stop() } } catch {}
        try { $job.PS.Dispose() } catch {}
        try { $job.Runspace.Dispose() } catch {}
    }
    $Global:BgJobs.Clear()
})

# ============================================================================
# SHOW WINDOW
# ============================================================================
[void]$Window.ShowDialog()
