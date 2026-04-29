#Requires -Version 5.1
<#
.SYNOPSIS
    AVD Rewind - Backup, visualize, and selectively restore Azure Virtual Desktop environments.
.DESCRIPTION
    PowerShell/WPF tool that captures AVD metadata (workspaces, host pools, application groups,
    session hosts, scaling plans, RBAC, diagnostics), visualizes the topology as an interactive
    tree, and enables selective or full restore - for accidental deletion recovery or cross-region
    migration of AVD control-plane objects.
.NOTES
    Author : Anton Romanyuk
    Version: 0.1.0
    Date   : 2026-03-26
#>

# ===============================================================================
# SECTION 1: PRE-LOAD & INITIALIZATION
# Loads WPF/WinForms assemblies, sets DPI awareness, bootstraps storage directories,
# defines named constants (timer interval, toast duration, cache TTL), and caches a BrushConverter.
# ===============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'

# Strip OneDrive user-profile module path to prevent old Az.Accounts versions
$env:PSModulePath = ($env:PSModulePath -split ';' |
    Where-Object { $_ -notlike '*OneDrive*' }) -join ';'

$Global:Root = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Global:Root)) { $Global:Root = $PWD.Path }

$Global:AppVersion     = "0.1.0-alpha"
$Global:AppTitle       = "AVD Rewind v$($Global:AppVersion)"
$Global:PrefsPath      = Join-Path $Global:Root "user_prefs.json"
$Global:AchievementsFile = Join-Path $Global:Root "achievements.json"
$Global:BackupDir      = Join-Path $Global:Root "backups"
$Global:ReportDir      = Join-Path $Global:Root "reports"
$Global:DebugLogFile   = Join-Path $env:TEMP "AvdRewind_debug.log"

# Ensure storage directories exist
foreach ($dir in @($Global:BackupDir, $Global:ReportDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
        Write-Host "[INIT] Created directory: $dir" -ForegroundColor DarkGray
    }
}

# Named constants
$Script:LOG_MAX_LINES       = 500
$Script:TIMER_INTERVAL_MS   = 50
$Script:TOAST_DURATION_MS   = 4000
$Script:CONFETTI_COUNT       = 60
$Script:CLEANUP_DELAY_MS     = 5000
$Script:TOPOLOGY_CACHE_TTL_S = 300

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
# Validates six Az modules (Accounts, DesktopVirtualization, Resources, Compute, Network,
# Monitor) at minimum versions. Renders a status grid and copy-to-clipboard install command.
# ===============================================================================

<#
.SYNOPSIS
    Validates that all required Azure PowerShell modules are installed at minimum versions.
.DESCRIPTION
    Checks six Az modules (Az.Accounts, Az.DesktopVirtualization, Az.Resources, Az.Compute,
    Az.Network, Az.Monitor) against their minimum required versions. Returns a result object
    indicating overall success, missing modules, and full status for each module.
.OUTPUTS
    PSCustomObject with Success (bool), Missing (array), and All (array) properties.
#>
function Test-Prerequisites {
    $Modules = @(
        @{ Name='Az.Accounts';              Required='2.7.5';  Purpose='Authentication & subscription context' }
        @{ Name='Az.DesktopVirtualization'; Required='4.0.0';  Purpose='AVD host pools, app groups, workspaces' }
        @{ Name='Az.Resources';             Required='6.0.0';  Purpose='RBAC role assignments, resource groups' }
        @{ Name='Az.Compute';               Required='5.0.0';  Purpose='VM metadata, Invoke-AzVMRunCommand' }
        @{ Name='Az.Network';               Required='5.0.0';  Purpose='NIC/subnet/NSG mapping per session host' }
        @{ Name='Az.Monitor';               Required='4.0.0';  Purpose='Diagnostic settings backup/restore' }
    )
    $Results = foreach ($M in $Modules) {
        $Installed = Get-Module -ListAvailable -Name $M.Name -ErrorAction SilentlyContinue |
                     Sort-Object Version -Descending | Select-Object -First 1
        [PSCustomObject]@{
            Name      = $M.Name
            Required  = $M.Required
            Purpose   = $M.Purpose
            Installed = if ($Installed) { $Installed.Version.ToString() } else { $null }
            Available = ($null -ne $Installed) -and ([version]$Installed.Version -ge [version]$M.Required)
        }
    }
    $Missing = @($Results | Where-Object { -not $_.Available })
    return [PSCustomObject]@{
        Success = ($Missing.Count -eq 0)
        Missing = $Missing
        All     = $Results
    }
}

<#
.SYNOPSIS
    Builds a one-line Install-Module command string for all missing prerequisite modules.
.PARAMETER MissingModules
    Array of module result objects from Test-Prerequisites that lack the required version.
.OUTPUTS
    String containing semicolon-separated Install-Module commands, or empty string if none missing.
#>
function Get-RemediationCommand {
    param($MissingModules)
    $Missing = $MissingModules
    if (-not $Missing) { return '' }
    $InstallCmds = $Missing | ForEach-Object { "Install-Module $($_.Name) -MinimumVersion $($_.Required) -Force" }
    return ($InstallCmds -join '; ')
}

<#
.SYNOPSIS
    Renders the prerequisite module status grid in the Settings tab.
.DESCRIPTION
    Clears and rebuilds the prerequisite status panel with one row per module showing
    name, purpose, required version, installed version, and a color-coded status badge.
    Shows or hides the Copy Install Command button based on whether any modules are missing.
.PARAMETER PrereqResult
    Result object from Test-Prerequisites containing All and Missing arrays.
#>
function Render-PrereqStatus {
    param($PrereqResult)
    $pnlPrereqStatus.Children.Clear()
    $AnyMissing = $false
    foreach ($M in $PrereqResult.All) {
        $Row = [System.Windows.Controls.Grid]::new()
        $Row.Margin = [System.Windows.Thickness]::new(0,0,0,6)
        $C0 = [System.Windows.Controls.ColumnDefinition]::new(); $C0.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $C1 = [System.Windows.Controls.ColumnDefinition]::new(); $C1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Auto)
        [void]$Row.ColumnDefinitions.Add($C0)
        [void]$Row.ColumnDefinitions.Add($C1)

        # Module name + purpose
        $NamePanel = [System.Windows.Controls.StackPanel]::new()
        $NameLbl = [System.Windows.Controls.TextBlock]::new()
        $NameLbl.Text = $M.Name
        $NameLbl.FontSize = 12
        $NameLbl.FontWeight = [System.Windows.FontWeights]::SemiBold
        $NameLbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
        [void]$NamePanel.Children.Add($NameLbl)

        $PurposeLbl = [System.Windows.Controls.TextBlock]::new()
        $PurposeLbl.Text = "$($M.Purpose) (>= v$($M.Required))"
        $PurposeLbl.FontSize = 10
        $PurposeLbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
        $PurposeLbl.TextWrapping = 'Wrap'
        [void]$NamePanel.Children.Add($PurposeLbl)

        [System.Windows.Controls.Grid]::SetColumn($NamePanel, 0)
        [void]$Row.Children.Add($NamePanel)

        # Status badge
        $Badge = [System.Windows.Controls.Border]::new()
        $Badge.CornerRadius = [System.Windows.CornerRadius]::new(4)
        $Badge.Padding = [System.Windows.Thickness]::new(8,2,8,2)
        $Badge.VerticalAlignment = 'Center'
        $StatusLbl = [System.Windows.Controls.TextBlock]::new()
        $StatusLbl.FontSize = 10
        $StatusLbl.FontWeight = [System.Windows.FontWeights]::SemiBold

        if ($M.Available) {
            $StatusLbl.Text = "v$($M.Installed)"
            $Badge.Background = $Global:CachedBC.ConvertFromString($(if($Global:IsLightMode){'#2016A34A'}else{'#1A00C853'}))
            $StatusLbl.Foreground = $Global:CachedBC.ConvertFromString($(if($Global:IsLightMode){'#15803D'}else{'#4ADE80'}))
            $Badge.ToolTip = "Installed: v$($M.Installed) (requires >= v$($M.Required))"
        } else {
            $AnyMissing = $true
            $StatusLbl.Text = 'MISSING'
            $Badge.Background = $Global:CachedBC.ConvertFromString($(if($Global:IsLightMode){'#20DC2626'}else{'#20FF5000'}))
            $StatusLbl.Foreground = $Global:CachedBC.ConvertFromString($(if($Global:IsLightMode){'#DC2626'}else{'#FF5000'}))
            $Badge.ToolTip = "Not installed or below minimum version v$($M.Required). Run: Install-Module $($M.Name) -MinimumVersion $($M.Required) -Force"
        }
        $Badge.Child = $StatusLbl
        [System.Windows.Controls.Grid]::SetColumn($Badge, 1)
        [void]$Row.Children.Add($Badge)
        [void]$pnlPrereqStatus.Children.Add($Row)
    }

    # Show/hide install command button
    if ($AnyMissing) {
        $btnCopyInstallCmd.Visibility = 'Visible'
        $Cmd = Get-RemediationCommand -MissingModules $PrereqResult.Missing
        $Global:PrereqInstallCmd = $Cmd
    } else {
        $btnCopyInstallCmd.Visibility = 'Collapsed'
    }
}

# ===============================================================================
# SECTION 3: THREAD SYNCHRONIZATION BRIDGE
# Creates a synchronized hashtable with StatusQueue and LogQueue for safe cross-thread
# communication, tracks background runspace jobs in $Global:BgJobs ArrayList.
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
    and function definitions, prepends PSModulePath cleanup (strips OneDrive paths), and begins
    asynchronous invocation. The job is tracked in $Global:BgJobs and polled by the dispatcher
    timer for completion. On completion, the OnComplete callback runs on the UI thread.
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
.PARAMETER Functions
    Array of function names to inject into the runspace (copies their ScriptBlock definitions).
#>
function Start-BackgroundWork {
    param(
        [string]$Name = 'BgWork',
        [ScriptBlock]$ScriptBlock,
        [ScriptBlock]$OnComplete,
        [array]$Arguments = @(),
        [hashtable]$Variables = @{},
        [hashtable]$Context   = @{},
        [string[]]$Functions  = @()
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
    foreach ($k in $Variables.Keys) {
        $PS.Runspace.SessionStateProxy.SetVariable($k, $Variables[$k])
    }
    $PS.Runspace.SessionStateProxy.SetVariable('__CleanModulePath', $CleanModulePath)

    # Build function definitions to inject into the runspace
    $FuncDefs = ''
    foreach ($fn in $Functions) {
        $cmdInfo = Get-Command $fn -ErrorAction SilentlyContinue
        if ($cmdInfo -and $cmdInfo.ScriptBlock) {
            $FuncDefs += "function $fn {`n$($cmdInfo.ScriptBlock.ToString())`n}`n`n"
        }
    }

    # Build script: prepend PSModulePath cleanup as a separate line
    $Preamble = '$env:PSModulePath = $__CleanModulePath'
    $ScriptText = $Work.ToString().TrimStart()

    # If script has a param block, we must inject the preamble AFTER it
    if ($ScriptText -match '(?s)^(param\s*\([^)]*\)\s*)(.*)$') {
        $ParamPart = $Matches[1]
        $BodyPart  = $Matches[2]
        $FullScript = $ParamPart + "`n" + $Preamble + "`n" + $FuncDefs + $BodyPart
    } else {
        $FullScript = $Preamble + "`n" + $FuncDefs + $ScriptText
    }

    [void]$PS.AddScript($FullScript)
    foreach ($arg in $Arguments) {
        [void]$PS.AddArgument($arg)
    }

    Write-DebugLog "BgWork '$Name': launching runspace" -Level 'DEBUG'
    $Async = $PS.BeginInvoke()
    [void]$Global:BgJobs.Add(@{
        PS          = $PS
        Runspace    = $RS
        Handle      = $Async
        OnComplete  = $OnComplete
        Context     = $Context
        Name        = $Name
        StartTime   = (Get-Date)
    })
    Write-DebugLog "BgWork: queued job #$($Global:BgJobs.Count)" -Level 'DEBUG'
}

# ===============================================================================
# SECTION 4: XAML GUI LOAD
# Reads AvdRewind_UI.xaml, strips designer-only attributes, parses with XamlReader,
# and creates the main WPF Window object. Fatal errors show a MessageBox and exit.
# ===============================================================================

$XamlPath = Join-Path $Global:Root "AvdRewind_UI.xaml"
if (-not (Test-Path $XamlPath)) {
    [System.Windows.MessageBox]::Show(
        "XAML file not found:`n$XamlPath`n`nEnsure AvdRewind_UI.xaml is in the same folder as this script.",
        "AVD Rewind - Fatal Error", 'OK', 'Error') | Out-Null
    exit 1
}

try {
    $XamlContent = [System.IO.File]::ReadAllText($XamlPath)
    # Remove attributes incompatible with XmlNodeReader / PS 5.1
    $XamlContent = $XamlContent -replace 'x:Class="[^"]*"', ''
    $XamlContent = $XamlContent -replace 'mc:Ignorable="[^"]*"', ''
    $XamlContent = $XamlContent -replace 'xmlns:mc="[^"]*"', ''
    $XamlContent = $XamlContent -replace 'PresentationOptions:Freeze="[^"]*"', ''
    $XamlContent = $XamlContent -replace 'xmlns:PresentationOptions="[^"]*"', ''
    $XamlDoc = [xml]$XamlContent
    $Reader = New-Object System.Xml.XmlNodeReader $XamlDoc
    $Window = [System.Windows.Markup.XamlReader]::Load($Reader)
    Write-Host "[INIT] XAML parsed successfully" -ForegroundColor DarkGray
} catch {
    [System.Windows.MessageBox]::Show(
        "Failed to parse XAML:`n$($_.Exception.Message)",
        "AVD Rewind - Fatal Error", 'OK', 'Error') | Out-Null
    exit 1
}

# ===============================================================================
# SECTION 5: ELEMENT REFERENCES
# Binds 100+ named XAML elements (title bar, auth bar, icon rail, sidebar panels, tab panels,
# dashboard cards, topology tree, backup list, restore tree, settings controls, splash overlay,
# debug log, status bar) to PowerShell variables via FindName().
# ===============================================================================

# Title bar
$titleBar          = $Window.FindName("titleBar")
$lblTitle          = $Window.FindName("lblTitle")
$lblTitleVersion   = $Window.FindName("lblTitleVersion")
$btnHelp           = $Window.FindName("btnHelp")
$btnThemeToggle    = $Window.FindName("btnThemeToggle")
$btnMinimize       = $Window.FindName("btnMinimize")
$btnMaximize       = $Window.FindName("btnMaximize")
$btnClose          = $Window.FindName("btnClose")

# Auth bar
$authDot           = $Window.FindName("authDot")
$lblAuthStatus     = $Window.FindName("lblAuthStatus")
$cmbSubscription   = $Window.FindName("cmbSubscription")
$btnLogin          = $Window.FindName("btnLogin")
$btnLogout         = $Window.FindName("btnLogout")

# Icon rail
$btnHamburger      = $Window.FindName("btnHamburger")
$railDashboard     = $Window.FindName("railDashboard")
$railTopology      = $Window.FindName("railTopology")
$railBackups       = $Window.FindName("railBackups")
$railRestore       = $Window.FindName("railRestore")
$railSettings      = $Window.FindName("railSettings")
$railOutput        = $Window.FindName("railOutput")

# Rail indicators
$railDashboardIndicator = $Window.FindName("railDashboardIndicator")
$railTopologyIndicator  = $Window.FindName("railTopologyIndicator")
$railBackupsIndicator   = $Window.FindName("railBackupsIndicator")
$railRestoreIndicator   = $Window.FindName("railRestoreIndicator")
$railSettingsIndicator  = $Window.FindName("railSettingsIndicator")

# Sidebar panels
$pnlSidebar            = $Window.FindName("pnlSidebar")
$pnlSidebarDashboard   = $Window.FindName("pnlSidebarDashboard")
$pnlSidebarTopology    = $Window.FindName("pnlSidebarTopology")
$pnlSidebarBackups     = $Window.FindName("pnlSidebarBackups")
$pnlSidebarRestore     = $Window.FindName("pnlSidebarRestore")
$pnlSidebarSettings    = $Window.FindName("pnlSidebarSettings")

# Sidebar: Dashboard
$lblSidebarHostPools    = $Window.FindName("lblSidebarHostPools")
$lblSidebarSessionHosts = $Window.FindName("lblSidebarSessionHosts")
$lblSidebarLastBackup   = $Window.FindName("lblSidebarLastBackup")

# Sidebar: Topology
$cmbTopoResourceGroup   = $Window.FindName("cmbTopoResourceGroup")
$txtTopoSearch          = $Window.FindName("txtTopoSearch")

# Sidebar: Backups
$lblSidebarBackupCount  = $Window.FindName("lblSidebarBackupCount")
$btnCompareWithCurrent  = $Window.FindName("btnCompareWithCurrent")

# Sidebar: Restore
$cmbRestoreSource       = $Window.FindName("cmbRestoreSource")
$rdoSameRegion          = $Window.FindName("rdoSameRegion")
$rdoDiffRegion          = $Window.FindName("rdoDiffRegion")
$cmbTargetRegion        = $Window.FindName("cmbTargetRegion")

# Sidebar: Restore - RG Mapping & VM Discovery
$chkAutoCreateRg        = $Window.FindName("chkAutoCreateRg")
$pnlRgMapping           = $Window.FindName("pnlRgMapping")
$lblRgMappingHint       = $Window.FindName("lblRgMappingHint")
$chkDiscoverVms         = $Window.FindName("chkDiscoverVms")
$lblDiscoverVmsHint     = $Window.FindName("lblDiscoverVmsHint")

# Achievements
$lblAchievementCount    = $Window.FindName("lblAchievementCount")
$pnlAchievements        = $Window.FindName("pnlAchievements")
$pnlAchievementsSettings = $Window.FindName("pnlAchievementsSettings")

# Main content area
$bdrDotGrid        = $Window.FindName("bdrDotGrid")
$bdrGradientGlow   = $Window.FindName("bdrGradientGlow")
$bdrShimmer        = $Window.FindName("bdrShimmer")
$pnlGlobalProgress = $Window.FindName("pnlGlobalProgress")
$prgGlobal         = $Window.FindName("prgGlobal")
$lblGlobalProgress = $Window.FindName("lblGlobalProgress")
$pnlToastHost      = $Window.FindName("pnlToastHost")

# Tab headers
$tabDashboard      = $Window.FindName("tabDashboard")
$tabTopology       = $Window.FindName("tabTopology")
$tabBackups        = $Window.FindName("tabBackups")
$tabRestore        = $Window.FindName("tabRestore")
$tabSettings       = $Window.FindName("tabSettings")

# Tab bar buttons
$btnBackupNow      = $Window.FindName("btnBackupNow")
$btnExportReport   = $Window.FindName("btnExportReport")
$btnRefreshTopology = $Window.FindName("btnRefreshTopology")

# Tab panels
$pnlTabDashboard   = $Window.FindName("pnlTabDashboard")
$pnlTabTopology    = $Window.FindName("pnlTabTopology")
$pnlTabBackups     = $Window.FindName("pnlTabBackups")
$pnlTabRestore     = $Window.FindName("pnlTabRestore")
$pnlTabSettings    = $Window.FindName("pnlTabSettings")

# Dashboard content
$pnlDashboardCards  = $Window.FindName("pnlDashboardCards")
$lblDashHostPools   = $Window.FindName("lblDashHostPools")
$lblDashAppGroups   = $Window.FindName("lblDashAppGroups")
$lblDashWorkspaces  = $Window.FindName("lblDashWorkspaces")
$lblDashSessionHosts = $Window.FindName("lblDashSessionHosts")
$lblDashScalingPlans = $Window.FindName("lblDashScalingPlans")
$pnlDriftBanner    = $Window.FindName("pnlDriftBanner")
$lblDriftMessage    = $Window.FindName("lblDriftMessage")
$btnQuickBackup    = $Window.FindName("btnQuickBackup")
$btnQuickRestore   = $Window.FindName("btnQuickRestore")
$btnQuickCompare   = $Window.FindName("btnQuickCompare")
$pnlRecentActivity = $Window.FindName("pnlRecentActivity")

# Topology
$tvTopology        = $Window.FindName("tvTopology")
$lblTopologyEmpty  = $Window.FindName("lblTopologyEmpty")

# Backups
$lstBackups        = $Window.FindName("lstBackups")
$lblBackupsEmpty   = $Window.FindName("lblBackupsEmpty")
$pnlDiffOverlay    = $Window.FindName("pnlDiffOverlay")
$pnlDiffResults    = $Window.FindName("pnlDiffResults")

# Restore
$tvRestoreTree     = $Window.FindName("tvRestoreTree")
$lblRestoreEmpty   = $Window.FindName("lblRestoreEmpty")
$pnlPreFlight      = $Window.FindName("pnlPreFlight")
$pnlPreFlightResults = $Window.FindName("pnlPreFlightResults")
$btnExecuteRestore = $Window.FindName("btnExecuteRestore")
$btnDryRun         = $Window.FindName("btnDryRun")
$btnExportPlan     = $Window.FindName("btnExportPlan")

# Settings
$chkDarkMode       = $Window.FindName("chkDarkMode")
$chkAutoBackup     = $Window.FindName("chkAutoBackup")
$chkRetentionEnabled = $Window.FindName("chkRetentionEnabled")
$pnlRetentionDays  = $Window.FindName("pnlRetentionDays")
$txtRetentionDays  = $Window.FindName("txtRetentionDays")
$btnResetSettings  = $Window.FindName("btnResetSettings")
$pnlPrereqStatus   = $Window.FindName("pnlPrereqStatus")
$btnCopyInstallCmd  = $Window.FindName("btnCopyInstallCmd")
$lblCopyInstallCmd  = $Window.FindName("lblCopyInstallCmd")

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

# Debug log
$logScroller       = $Window.FindName("logScroller")
$paraLog           = $Window.FindName("paraLog")
$btnClearLog       = $Window.FindName("btnClearLog")

# Status bar
$lblStatusBar      = $Window.FindName("lblStatusBar")
$lblStatusRight    = $Window.FindName("lblStatusRight")

# Layout references
$colLeftPanel      = $Window.FindName("colLeftPanel")
$rowBottomPanel    = $Window.FindName("rowBottomPanel")

Write-Host "[INIT] Element references bound ($((Get-Variable -Scope Script | Where-Object { $_.Value -is [System.Windows.FrameworkElement] -or $_.Value -is [System.Windows.Documents.Paragraph] }).Count) elements)" -ForegroundColor DarkGray

# ===============================================================================
# SECTION 6: GLOBAL STATE
# Initialises all script-level state variables: auth context, topology cache, UI flags,
# achievement tracking, logging ring buffers, and sidebar visibility.
# ===============================================================================

$Global:IsLightMode          = $false
$Global:ActiveTabName        = 'Dashboard'
$Global:ActiveRailName       = 'Dashboard'
$Global:CurrentSubId         = ''
$Global:CurrentSubName       = ''
$Global:Subscriptions        = @()
$Global:SubsSeen             = @{}
$Global:CurrentTopology      = $null
$Global:TopologyCacheTime    = [datetime]::MinValue
$Global:SelectedBackup       = $null
$Global:Achievements         = @()
$Global:WindowRendered       = $false
$Global:SplashDismissQueued  = $false
$Global:DebugOverlayEnabled  = $false
$Global:ThemeToggleCount     = 0
$Global:DiscoveryStartTime   = $null
$Global:DiscoveryInProgress  = $false
$Global:SidebarVisible       = $true

$Global:DebugLineCount       = 0
$Global:DebugMaxLines        = $Script:LOG_MAX_LINES
$Global:FullLogSB            = [System.Text.StringBuilder]::new()
$Global:FullLogLineCount     = 0
$Global:FullLogMaxLines      = 1000

# ===============================================================================
# SECTION 7: WRITE-DEBUGLOG (3-destination output)
# Centralized logging function writing to console, rotating disk file, in-memory ring buffer,
# and the UI activity log RichTextBox with level-based color coding (INFO/DEBUG/WARN/ERROR/SUCCESS).
# ===============================================================================

<#
.SYNOPSIS
    Writes a timestamped log entry to three destinations: console, disk file, and UI RichTextBox.
.DESCRIPTION
    Formats a log line as [HH:mm:ss.fff] [LEVEL] Message and outputs to:
    (1) PowerShell console via Write-Host, (2) rotating disk log file at %TEMP%\AvdRewind_debug.log
    (auto-rotates at 2 MB), (3) in-memory ring buffer (1000 lines), and (4) the activity log
    RichTextBox panel with level-based color coding. DEBUG-level messages are suppressed in the
    UI unless $Global:DebugOverlayEnabled is true.
.PARAMETER Message
    The log message text. Supports string interpolation.
.PARAMETER Level
    Log severity: INFO, DEBUG, WARN, ERROR, or SUCCESS. Defaults to INFO.
#>
function Write-DebugLog {
    param([string]$Message, [string]$Level = 'INFO')

    $ts   = (Get-Date).ToString('HH:mm:ss.fff')
    $line = "[$ts] [$Level] $Message"

    # 1. Console output (always)
    Write-Host $line -ForegroundColor DarkGray

    # 2. Disk log (always, rotate at 2 MB)
    try {
        $diskLogPath = $Global:DebugLogFile
        if (-not [string]::IsNullOrWhiteSpace($diskLogPath)) {
            if ((Test-Path $diskLogPath) -and (Get-Item $diskLogPath).Length -gt 2MB) {
                $Rotated = $diskLogPath + '.old'
                if (Test-Path $Rotated) { Remove-Item $Rotated -Force }
                Rename-Item $diskLogPath $Rotated -Force
                Write-Host "[LOG] Rotated disk log (>2 MB)" -ForegroundColor DarkGray
            }
            [System.IO.File]::AppendAllText($diskLogPath, $line + "`r`n")
        }
    } catch { }

    # 2b. Full log ring buffer (memory, 1000 lines)
    if ($Global:FullLogSB) {
        [void]$Global:FullLogSB.AppendLine($line)
        $Global:FullLogLineCount++
        if ($Global:FullLogLineCount -gt $Global:FullLogMaxLines) {
            $idx = $Global:FullLogSB.ToString().IndexOf("`n")
            if ($idx -ge 0) { [void]$Global:FullLogSB.Remove(0, $idx + 1) }
            $Global:FullLogLineCount--
        }
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
# SECTION 8: THEME SYSTEM
# Dark and Light palettes with 40+ color keys each (backgrounds, accents, borders, text, status).
# $ApplyTheme scriptblock swaps all Window.Resources brushes and regenerates procedural backgrounds.
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
    ThemeSuccessBg    = "#1A00C853"; ThemeSuccessText   = "#4ADE80"
    ThemeWarningBg    = "#1AF59E0B"; ThemeWarningText   = "#FBBF24"
    ThemeError         = "#FF5000";  ThemeErrorDim      = "#20FF5000"
    ThemeProgressEdge  = "#18FFFFFF"
    ThemeSidebarBg     = "#111113";  ThemeSidebarBorder = "#00000000"
}

$Global:ThemeLight = @{
    ThemeAppBg         = "#F5F5F5";  ThemePanelBg       = "#FAFAFA"
    ThemeCardBg        = "#FFFFFF";  ThemeCardAltBg     = "#F8F8F8"
    ThemeInputBg       = "#F0F0F0";  ThemeDeepBg        = "#EEEEEE"
    ThemeOutputBg      = "#F0F0F0";  ThemeSurfaceBg     = "#F2F2F5"
    ThemeHoverBg       = "#E8E8EC";  ThemeSelectedBg    = "#E0E0E0"
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
    ThemeSuccessBg    = "#2016A34A"; ThemeSuccessText   = "#15803D"
    ThemeWarningBg    = "#20EA580C"; ThemeWarningText   = "#C2410C"
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
# SECTION 9: TOAST NOTIFICATIONS
# Auto-dismissing overlay notifications with fade-in/fade-out animations. Four visual types:
# Success (green), Error (red), Warning (amber), Info (blue) with theme-aware color palettes.
# ===============================================================================

<#
.SYNOPSIS
    Displays an auto-dismissing toast notification in the top-right corner of the window.
.DESCRIPTION
    Creates a themed border with icon and message text, fades it in over 200ms, then
    auto-dismisses after the specified duration with a fade-out animation. Supports four
    types (Success, Error, Warning, Info) with theme-aware color palettes for both dark
    and light modes. Multiple toasts stack vertically.
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
    [void]$SP.Children.Add($IconTB)

    $MsgTB = [System.Windows.Controls.TextBlock]::new()
    $MsgTB.Text       = $Message
    $MsgTB.FontSize   = 12
    $MsgTB.Foreground = $Global:CachedBC.ConvertFromString($Colors.Text)
    $MsgTB.VerticalAlignment = 'Center'
    $MsgTB.TextWrapping = 'Wrap'
    $MsgTB.MaxWidth   = 400
    [void]$SP.Children.Add($MsgTB)

    $Toast.Child = $SP
    $pnlToastHost.Children.Insert(0, $Toast)

    # Fade in
    $FadeIn = [System.Windows.Media.Animation.DoubleAnimation]::new()
    $FadeIn.From     = 0; $FadeIn.To = 1
    $FadeIn.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(200))
    $Toast.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $FadeIn)

    # Auto-dismiss timer
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

    Write-DebugLog "[TOAST] [$Type] $Message (duration=${DurationMs}ms)" -Level 'DEBUG'
}

# ===============================================================================
# SECTION 10: SHOW-THEMEDDIALOG
# Chromeless modal dialog window inheriting parent theme resources. Supports icon, title,
# message, optional text input with match validation, and configurable button array.
# ===============================================================================

<#
.SYNOPSIS
    Shows a modal themed dialog window with icon, title, message, optional text input, and custom buttons.
.DESCRIPTION
    Creates a chromeless WPF dialog that inherits the parent window's theme resources. Supports
    configurable icon, title, message body, optional text input field with match validation (e.g.
    type `RESTORE` to confirm), and an array of button definitions. The dialog is draggable and
    returns the result tag of the clicked button, or 'Cancel' if dismissed.
.PARAMETER Title
    Dialog header text. Defaults to 'Confirm'.
.PARAMETER Message
    Body text displayed below the header separator.
.PARAMETER Icon
    Segoe Fluent Icons character code for the header icon badge.
.PARAMETER IconColor
    Theme resource key for the icon foreground color.
.PARAMETER Buttons
    Array of hashtables defining buttons. Each has Text, IsAccent (bool), and Result (string tag).
.PARAMETER InputPrompt
    If non-empty, shows a labeled text input field above the buttons.
.PARAMETER InputMatch
    If non-empty, the accept button is disabled until the input text matches this exact string.
.OUTPUTS
    String result tag of clicked button, or hashtable with Result and Input if InputPrompt was used.
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

    Write-DebugLog "[DIALOG] Opening: Title='$Title'" -Level 'DEBUG'
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
        $MsgBlock.Margin     = [System.Windows.Thickness]::new(0,0,0, $(if ($InputPrompt) { 14 } else { 24 }))
        $MsgBlock.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextSecondary')
        [void]$MainStack.Children.Add($MsgBlock)
    }

    # Optional input field
    $InputBox = $null
    if ($InputPrompt) {
        $InputLabel = [System.Windows.Controls.TextBlock]::new()
        $InputLabel.Text = $InputPrompt
        $InputLabel.FontSize = 11
        $InputLabel.Margin   = [System.Windows.Thickness]::new(0,0,0,6)
        $InputLabel.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextSecondary')
        [void]$MainStack.Children.Add($InputLabel)

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
        [void]$MainStack.Children.Add($InputBorder)
    }

    # InputMatch: track the accept button for enabling/disabling
    $AcceptBtnRef = $null

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
            if ($InputMatch -and $InputBox) {
                $Btn.IsEnabled = $false
                $BtnBorder.Opacity = 0.4
                $AcceptBtnRef = @{ Btn = $Btn; Border = $BtnBorder }
            }
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

    # Wire InputMatch validation
    if ($InputMatch -and $InputBox -and $AcceptBtnRef) {
        $matchPattern = $InputMatch
        $btnRef = $AcceptBtnRef
        $InputBox.Add_TextChanged({
            $text = $InputBox.Text.Trim()
            $isMatch = $text -eq $matchPattern
            $btnRef.Btn.IsEnabled = $isMatch
            $btnRef.Border.Opacity = if ($isMatch) { 1.0 } else { 0.4 }
        }.GetNewClosure())
    }

    $OuterBorder.Add_MouseLeftButtonDown({
        try { $Dlg.DragMove() } catch { }
    }.GetNewClosure())

    $Dlg.ShowDialog() | Out-Null
    $DlgResult = if ($Dlg.Tag) { $Dlg.Tag } else { 'Cancel' }
    Write-DebugLog "[DIALOG] Closed: Result='$DlgResult'" -Level 'DEBUG'

    if ($InputBox) {
        return @{ Result = $DlgResult; Input = $InputBox.Text }
    }
    return $DlgResult
}

# ===============================================================================
# SECTION 11: SHIMMER PROGRESS ANIMATION
# Global progress bar supporting both determinate (percentage) and indeterminate (shimmer/marquee)
# modes. Start-Shimmer / Stop-Shimmer convenience wrappers for background operations.
# ===============================================================================

$Global:ShimmerRunning = $false

<#
.SYNOPSIS
    Shows the global determinate progress bar with a percentage value and status text.
.PARAMETER Value
    Current progress value (0 to Maximum).
.PARAMETER Maximum
    Maximum value for the progress bar. Defaults to 100.
.PARAMETER Text
    Status text displayed next to the progress bar.
#>
function Show-GlobalProgress {
    param([int]$Value = 0, [int]$Maximum = 100, [string]$Text = '')
    $prgGlobal.Maximum = $Maximum
    $prgGlobal.Value   = $Value
    $prgGlobal.IsIndeterminate = $false
    $prgGlobal.Visibility = 'Visible'
    $bdrShimmer.Visibility = 'Collapsed'
    $lblGlobalProgress.Text = $Text
    $pnlGlobalProgress.Visibility = 'Visible'
}

<#
.SYNOPSIS
    Shows the global progress bar in indeterminate shimmer/marquee mode for unknown-duration operations.
.PARAMETER Text
    Status text displayed next to the shimmer animation.
#>
function Show-GlobalProgressIndeterminate {
    param([string]$Text = '')
    $prgGlobal.Visibility = 'Collapsed'
    $bdrShimmer.Visibility = 'Visible'
    $lblGlobalProgress.Text = $Text
    $pnlGlobalProgress.Visibility = 'Visible'
}

<#
.SYNOPSIS
    Updates the global progress bar value and optional status text.
#>
function Update-GlobalProgress {
    <# Updates progress bar value and label text #>
    param([int]$Value, [string]$Text)
    $prgGlobal.Value = $Value
    if ($Text) { $lblGlobalProgress.Text = $Text }
}

<#
.SYNOPSIS
    Hides the global progress bar and shimmer animation, collapsing the progress panel.
#>
function Hide-GlobalProgress {
    <# Hides the global progress bar #>
    $pnlGlobalProgress.Visibility = 'Collapsed'
    $prgGlobal.Value = 0
    $prgGlobal.IsIndeterminate = $false
    $prgGlobal.Visibility = 'Collapsed'
    $bdrShimmer.Visibility = 'Collapsed'
    $lblGlobalProgress.Text = ''
}

function Start-Shimmer { Show-GlobalProgressIndeterminate }

function Stop-Shimmer { Hide-GlobalProgress }

# ===============================================================================
# SECTION 11b: SPLASH STEP HELPERS
# ===============================================================================

<#
.SYNOPSIS
    Updates the splash screen step indicator dots and labels during startup initialization.
.DESCRIPTION
    Sets the visual state (pending, running, done, error, skipped) for one of the three
    startup steps shown on the splash overlay. Each step has a colored dot and text label.
#>
function Set-SplashStep {
    param([int]$Step, [string]$State, [string]$Text)
    $Dot = Get-Variable -Name "dotStep$Step" -ValueOnly -Scope Script
    $Lbl = Get-Variable -Name "lblStep$Step" -ValueOnly -Scope Script
    if ($Text) { $Lbl.Text = $Text }
    switch ($State) {
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
    Dismisses the startup splash overlay with a fade-out animation.
#>
function Hide-Splash {
    Write-DebugLog "[SPLASH] Dismissing splash overlay" -Level 'DEBUG'
    $pnlSplash.Visibility = 'Collapsed'
}

<#
.SYNOPSIS
    Schedules the splash screen to auto-hide after a short delay, or immediately if the window is already rendered.
#>
function Start-SplashDismissTimer {
    param([int]$DelayMs = 3000)
    $Script:SplashDismissDelayMs = $DelayMs
    if ($Global:WindowRendered) {
        $ST = [System.Windows.Threading.DispatcherTimer]::new()
        $ST.Interval = [TimeSpan]::FromMilliseconds($DelayMs)
        $STRef = $ST
        $HideSplashFn = { Hide-Splash }
        $ST.Add_Tick({ $STRef.Stop(); & $HideSplashFn }.GetNewClosure())
        $ST.Start()
    } else {
        $Global:SplashDismissQueued = $true
    }
}

# ===============================================================================
# SECTION 12: TAB & NAVIGATION MANAGEMENT
# ===============================================================================

$AllTabs = @('Dashboard','Topology','Backups','Restore','Settings')
$AllTabHeaders = @{
    Dashboard  = $tabDashboard
    Topology   = $tabTopology
    Backups    = $tabBackups
    Restore    = $tabRestore
    Settings   = $tabSettings
}
$AllTabPanels = @{
    Dashboard  = $pnlTabDashboard
    Topology   = $pnlTabTopology
    Backups    = $pnlTabBackups
    Restore    = $pnlTabRestore
    Settings   = $pnlTabSettings
}
$AllRailIndicators = @{
    Dashboard  = $railDashboardIndicator
    Topology   = $railTopologyIndicator
    Backups    = $railBackupsIndicator
    Restore    = $railRestoreIndicator
    Settings   = $railSettingsIndicator
}
$AllSidebarPanels = @{
    Dashboard  = $pnlSidebarDashboard
    Topology   = $pnlSidebarTopology
    Backups    = $pnlSidebarBackups
    Restore    = $pnlSidebarRestore
    Settings   = $pnlSidebarSettings
}

<#
.SYNOPSIS
    Switches the active tab in the main content area with fade animation and sidebar panel swap.
.DESCRIPTION
    Updates the tab header highlight, content panel visibility, icon rail indicator, and sidebar
    panel for the selected tab. Applies a 150ms fade-in animation to the newly visible panel.
    When switching to the Restore tab, auto-populates the restore source dropdown from backups.
.PARAMETER TabName
    Target tab name: Dashboard, Topology, Backups, Restore, or Settings.
#>
function Switch-Tab {
    param([string]$TabName)

    # Block navigation to Restore while discovery is running
    if ($TabName -eq 'Restore' -and $Global:DiscoveryInProgress) {
        Show-Toast -Message 'Discovery is running — Restore is unavailable until it completes.' -Type 'Warning'
        return
    }

    $Global:ActiveTabName = $TabName
    $Global:ActiveRailName = $TabName

    foreach ($t in $AllTabs) {
        $hdr = $AllTabHeaders[$t]
        if ($hdr) {
            if ($t -eq $TabName) {
                $hdr.BorderBrush = $Global:CachedBC.ConvertFromString('#0078D4')
                $hdr.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeHoverBg')
            } else {
                $hdr.BorderBrush = [System.Windows.Media.Brushes]::Transparent
                $hdr.Background  = [System.Windows.Media.Brushes]::Transparent
            }
        }

        $pnl = $AllTabPanels[$t]
        if ($pnl) {
            $pnl.Visibility = if ($t -eq $TabName) { 'Visible' } else { 'Collapsed' }
        }

        $ind = $AllRailIndicators[$t]
        if ($ind) {
            if ($t -eq $TabName) {
                $ind.SetResourceReference([System.Windows.Shapes.Rectangle]::FillProperty, 'ThemeAccentLight')
            } else {
                $ind.Fill = [System.Windows.Media.Brushes]::Transparent
            }
        }

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

    Write-DebugLog "[NAV] Switched to tab: $TabName" -Level 'DEBUG'

    # Topology will be loaded on-demand via Refresh button

    # Populate restore source dropdown when switching to Restore tab
    if ($TabName -eq 'Restore') {
        $cmbRestoreSource.Items.Clear()
        $Backups = Get-BackupList
        foreach ($bk in $Backups) {
            $item = [System.Windows.Controls.ComboBoxItem]::new()
            $item.Content = if ($bk.Label) { $bk.Label } else { $bk.FileName }
            $item.Tag = $bk
            [void]$cmbRestoreSource.Items.Add($item)
        }
        if ($cmbRestoreSource.Items.Count -gt 0) { $cmbRestoreSource.SelectedIndex = 0 }
    }
}

# ===============================================================================
# SECTION 13: USER PREFERENCES
# JSON-based persistence (user_prefs.json) for theme, auto-backup, retention, default
# subscription, and achievements. Retention cleanup deletes backups older than N days.
# ===============================================================================

<#
.SYNOPSIS
    Loads user preferences from user_prefs.json and applies them to UI controls.
.DESCRIPTION
    Reads the preferences file and sets theme mode, auto-backup toggle, retention settings,
    default subscription ID, and unlocked achievements. Logs success or failure.
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
            if ($null -ne $prefs.autoBackupOnLaunch) {
                $chkAutoBackup.IsChecked = [bool]$prefs.autoBackupOnLaunch
            }
            if ($null -ne $prefs.backupRetentionDays) {
                $txtRetentionDays.Text = $prefs.backupRetentionDays.ToString()
            }
            if ($null -ne $prefs.retentionEnabled) {
                $chkRetentionEnabled.IsChecked = [bool]$prefs.retentionEnabled
                $pnlRetentionDays.Visibility = if ([bool]$prefs.retentionEnabled) { 'Visible' } else { 'Collapsed' }
            }
            if ($null -ne $prefs.defaultSubscriptionId) {
                $Global:CurrentSubId = $prefs.defaultSubscriptionId
            }
            if ($null -ne $prefs.achievements) {
                $Global:Achievements = @($prefs.achievements)
                Write-DebugLog "[PREFS] Restored $($Global:Achievements.Count) achievement(s)" -Level 'DEBUG'
            }
            Write-DebugLog "[PREFS] Loaded: theme=$($prefs.theme), autoBackup=$($prefs.autoBackupOnLaunch), retention=$($prefs.backupRetentionDays)d" -Level 'SUCCESS'
        } catch {
            Write-DebugLog "[PREFS] Failed to load: $($_.Exception.Message)" -Level 'WARN'
        }
    }
}

<#
.SYNOPSIS
    Persists current UI settings and achievement state to user_prefs.json.
.DESCRIPTION
    Serializes theme mode, auto-backup, retention, subscription ID, and achievements
    to a JSON file at $Global:PrefsPath.
#>
function Save-UserPrefs {
    $prefs = @{
        theme                 = if ($Global:IsLightMode) { 'light' } else { 'dark' }
        autoBackupOnLaunch    = [bool]$chkAutoBackup.IsChecked
        retentionEnabled      = [bool]$chkRetentionEnabled.IsChecked
        backupRetentionDays   = [int]$txtRetentionDays.Text
        defaultSubscriptionId = $Global:CurrentSubId
        achievements          = @($Global:Achievements)
    }
    try {
        $prefs | ConvertTo-Json -Depth 4 | Set-Content $Global:PrefsPath -Encoding UTF8
        Write-DebugLog "[PREFS] Saved: theme=$($prefs.theme), sub=$($prefs.defaultSubscriptionId)" -Level 'DEBUG'
    } catch {
        Write-DebugLog "[PREFS] Save failed: $($_.Exception.Message)" -Level 'WARN'
    }
}

<#
.SYNOPSIS
    Deletes backup files older than the configured retention period.
.DESCRIPTION
    When retention is enabled, scans the backups directory for JSON files with a LastWriteTime
    older than the retention threshold (default 30 days) and removes them along with their
    SHA256 sidecar files. Unlocks the CleanSlate achievement if any backups were removed.
#>
function Invoke-RetentionCleanup {
    if (-not [bool]$chkRetentionEnabled.IsChecked) {
        Write-DebugLog "[RETENTION] Cleanup disabled -- skipping" -Level 'DEBUG'
        return
    }
    $days = 30
    if ([int]::TryParse($txtRetentionDays.Text, [ref]$days) -and $days -gt 0) { } else { $days = 30 }
    $cutoff = (Get-Date).AddDays(-$days)
    $removed = 0
    if (Test-Path $Global:BackupDir) {
        Get-ChildItem $Global:BackupDir -Filter '*.json' -File | Where-Object { $_.LastWriteTime -lt $cutoff } | ForEach-Object {
            try {
                Remove-Item $_.FullName -Force
                $sha = $_.FullName + '.sha256'
                if (Test-Path $sha) { Remove-Item $sha -Force }
                $removed++
                Write-DebugLog "[RETENTION] Deleted: $($_.Name)" -Level 'DEBUG'
            } catch {
                Write-DebugLog "[RETENTION] Failed to delete $($_.Name): $($_.Exception.Message)" -Level 'WARN'
            }
        }
    }
    if ($removed -gt 0) {
        Write-DebugLog "[RETENTION] Cleaned up $removed backup(s) older than $days days" -Level 'INFO'
        Check-Achievement 'CleanSlate'
    }
}

# ===============================================================================
# SECTION 14: AVD DISCOVERY ENGINE
# Full enumeration of AVD control-plane objects: workspaces, host pools, app groups, applications,
# RBAC assignments, session hosts (with VM metadata), scaling plans, and diagnostic settings.
# Uses batched RBAC and VM caching per resource group to minimize API calls.
# ===============================================================================

function Get-AvdTopology {
    <#
    .SYNOPSIS
        Enumerates all AVD objects in the current subscription and builds topology tree.
        Designed to run INSIDE a background runspace via Start-BackgroundWork.
    #>
    param([string]$SubscriptionId, [string]$SubscriptionName)

    $ProgressPreference = 'Continue'
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $Topology = [PSCustomObject]@{
        SchemaVersion   = '1.0'
        Timestamp       = (Get-Date).ToUniversalTime().ToString('o')
        SubscriptionId  = $SubscriptionId
        SubscriptionName = $SubscriptionName
        Workspaces      = @()
        HostPools       = @()
        ResourceGroups  = @()
    }

    # Enumerate workspaces
    Write-Progress -Activity 'AVD Discovery' -Status 'Discovering workspaces...' -PercentComplete 5
    $Workspaces = @(Get-AzWvdWorkspace -ErrorAction SilentlyContinue)
    $Topology.Workspaces = foreach ($ws in $Workspaces) {
        [PSCustomObject]@{
            ResourceId              = $ws.Id
            Name                    = $ws.Name
            ResourceGroupName       = ($ws.Id -split '/')[4]
            Location                = $ws.Location
            FriendlyName            = $ws.FriendlyName
            Description             = $ws.Description
            ApplicationGroupReferences = @($ws.ApplicationGroupReference)
            Tags                    = if ($ws.Tag -is [hashtable]) { $ws.Tag } elseif ($ws.Tag) { $ht = @{}; foreach ($k in $ws.Tag.Keys) { $ht[$k] = $ws.Tag[$k] }; $ht } else { @{} }
        }
    }

    # Enumerate host pools
    Write-Progress -Activity 'AVD Discovery' -Status "Found $(@($Workspaces).Count) workspace(s). Discovering host pools..." -PercentComplete 15
    $HostPools = @(Get-AzWvdHostPool -ErrorAction SilentlyContinue)

    # Prefetch shared collections once to avoid repeated Azure API calls per host pool.
    Write-Progress -Activity 'AVD Discovery' -Status "Prefetching app groups and scaling plans..." -PercentComplete 18
    $AllAppGroups = @(Get-AzWvdApplicationGroup -ErrorAction SilentlyContinue)
    $AppGroupsByHostPool = @{}
    foreach ($ag in $AllAppGroups) {
        if (-not $ag.HostPoolArmPath) { continue }
        $hpKey = $ag.HostPoolArmPath.ToLower()
        if (-not $AppGroupsByHostPool.ContainsKey($hpKey)) {
            $AppGroupsByHostPool[$hpKey] = [System.Collections.ArrayList]::new()
        }
        [void]$AppGroupsByHostPool[$hpKey].Add($ag)
    }

    $AllScalingPlans = @(Get-AzWvdScalingPlan -ErrorAction SilentlyContinue)
    $ScalingPlansByHostPool = @{}
    foreach ($sp in $AllScalingPlans) {
        foreach ($hpRef in @($sp.HostPoolReference)) {
            if (-not $hpRef -or -not $hpRef.HostPoolArmPath) { continue }
            $hpKey = $hpRef.HostPoolArmPath.ToLower()
            if (-not $ScalingPlansByHostPool.ContainsKey($hpKey)) {
                $ScalingPlansByHostPool[$hpKey] = [System.Collections.ArrayList]::new()
            }
            [void]$ScalingPlansByHostPool[$hpKey].Add($sp)
        }
    }

    $hpIdx = 0; $hpTotal = @($HostPools).Count
    $Topology.HostPools = foreach ($hp in $HostPools) {
        $hpIdx++
        $hpPct = [math]::Min(20 + [int](($hpIdx / [math]::Max($hpTotal,1)) * 60), 80)
        $hpRG = ($hp.Id -split '/')[4]

        # App Groups for this host pool
        Write-Progress -Activity 'AVD Discovery' -Status "[$hpIdx/$hpTotal] $($hp.Name) — loading app groups..." -PercentComplete $hpPct
        $hpKey = $hp.Id.ToLower()
        $AppGroups = if ($AppGroupsByHostPool.ContainsKey($hpKey)) { @($AppGroupsByHostPool[$hpKey]) } else { @() }

        # Batch-load RBAC for all scopes in RG once (avoids per-app-group calls)
        if (-not $script:rbacCache) { $script:rbacCache = @{} }
        if (-not $script:rbacCache.ContainsKey($hpRG)) {
            try {
                $rgRbac = @(Get-AzRoleAssignment -ResourceGroupName $hpRG -ErrorAction SilentlyContinue)
                $script:rbacCache[$hpRG] = $rgRbac
            } catch { $script:rbacCache[$hpRG] = @() }
        }
        $rgRbacAll = $script:rbacCache[$hpRG]

        $AppGroupData = foreach ($ag in $AppGroups) {
            $agRG = ($ag.Id -split '/')[4]

            # Applications (RemoteApp type only)
            $Apps = @()
            if ($ag.ApplicationGroupType -eq 'RemoteApp') {
                $Apps = @(Get-AzWvdApplication -ResourceGroupName $agRG -GroupName $ag.Name -ErrorAction SilentlyContinue | ForEach-Object {
                    [PSCustomObject]@{
                        ResourceId    = $_.Id
                        Name          = $_.Name
                        FriendlyName  = $_.FriendlyName
                        FilePath      = $_.FilePath
                        IconPath      = $_.IconPath
                        CommandLineSetting = $_.CommandLineSetting
                    }
                })
            }

            # RBAC assignments scoped to this app group (filtered from batch)
            $agScope = $ag.Id.ToLower()
            $Rbac = @($rgRbacAll | Where-Object { $_.Scope.ToLower() -eq $agScope } | ForEach-Object {
                [PSCustomObject]@{
                    RoleDefinitionName = $_.RoleDefinitionName
                    RoleDefinitionId   = $_.RoleDefinitionId
                    PrincipalId        = $_.ObjectId
                    PrincipalType      = $_.ObjectType
                    DisplayName        = $_.DisplayName
                    Scope              = $_.Scope
                }
            })

            [PSCustomObject]@{
                ResourceId           = $ag.Id
                Name                 = $ag.Name
                ResourceGroupName    = $agRG
                Location             = $ag.Location
                ApplicationGroupType = $ag.ApplicationGroupType
                FriendlyName         = $ag.FriendlyName
                Description          = $ag.Description
                HostPoolArmPath      = $ag.HostPoolArmPath
                Applications         = $Apps
                RbacAssignments      = $Rbac
                Tags                 = if ($ag.Tag -is [hashtable]) { $ag.Tag } elseif ($ag.Tag) { $ht = @{}; foreach ($k in $ag.Tag.Keys) { $ht[$k] = $ag.Tag[$k] }; $ht } else { @{} }
            }
        }

        # Session hosts
        Write-Progress -Activity 'AVD Discovery' -Status "[$hpIdx/$hpTotal] $($hp.Name) - loading session hosts..." -PercentComplete $hpPct
        $shPhase = [System.Diagnostics.Stopwatch]::StartNew()

        # Batch-load all VMs in the RG once (avoids per-host Get-AzVM calls)
        if (-not $script:vmCache) { $script:vmCache = @{} }
        if (-not $script:vmCache.ContainsKey($hpRG)) {
            try {
                $rgVms = @(Get-AzVM -ResourceGroupName $hpRG -ErrorAction SilentlyContinue)
                $script:vmCache[$hpRG] = @{}
                foreach ($v in $rgVms) { $script:vmCache[$hpRG][$v.Id.ToLower()] = $v }
            } catch { $script:vmCache[$hpRG] = @{} }
        }
        $vmLookup = $script:vmCache[$hpRG]

        $SessionHosts = @(Get-AzWvdSessionHost -ResourceGroupName $hpRG -HostPoolName $hp.Name -ErrorAction SilentlyContinue | ForEach-Object {
            $shName = $_.Name -replace "^$($hp.Name)/", ''
            $vmResId = $_.ResourceId
            $VmMeta = $null
            $NicInfo = $null

            if ($vmResId -and $vmLookup.ContainsKey($vmResId.ToLower())) {
                $vm = $vmLookup[$vmResId.ToLower()]
                $VmMeta = [PSCustomObject]@{
                    VmSize       = $vm.HardwareProfile.VmSize
                    OsDiskName   = $vm.StorageProfile.OsDisk.Name
                    OsType       = $vm.StorageProfile.OsDisk.OsType
                    ImageRef     = if ($vm.StorageProfile.ImageReference) {
                        "$($vm.StorageProfile.ImageReference.Publisher)/$($vm.StorageProfile.ImageReference.Offer)/$($vm.StorageProfile.ImageReference.Sku)"
                    } else { $null }
                    Tags         = $vm.Tags
                }
            }

            [PSCustomObject]@{
                Name              = $shName
                SessionHostPath   = $_.Name
                ResourceId        = $vmResId
                Status            = $_.Status
                AllowNewSession   = $_.AllowNewSession
                AssignedUser      = $_.AssignedUser
                LastHeartBeat     = $_.LastHeartBeat
                UpdateState       = $_.UpdateState
                VmMetadata        = $VmMeta
                NetworkConfig     = $NicInfo
            }
        })
        Write-Progress -Activity 'AVD Discovery' -Status "[$hpIdx/$hpTotal] $($hp.Name) - $($SessionHosts.Count) hosts ($([int]$shPhase.Elapsed.TotalSeconds)s)" -PercentComplete $hpPct

        # Scaling plans linked to this pool
        Write-Progress -Activity 'AVD Discovery' -Status "[$hpIdx/$hpTotal] $($hp.Name) — loading scaling plans..." -PercentComplete $hpPct
        $LinkedScalingPlans = if ($ScalingPlansByHostPool.ContainsKey($hpKey)) { @($ScalingPlansByHostPool[$hpKey]) } else { @() }
        $ScalingPlans = @($LinkedScalingPlans | ForEach-Object {
                [PSCustomObject]@{
                    ResourceId        = $_.Id
                    Name              = $_.Name
                    ResourceGroupName = ($_.Id -split '/')[4]
                    Location          = $_.Location
                    TimeZone          = $_.TimeZone
                    HostPoolType      = $_.HostPoolType
                    Schedules         = @($_.Schedule | ForEach-Object {
                        [PSCustomObject]@{
                            Name                  = $_.Name
                            DaysOfWeek            = $_.DaysOfWeek
                            RampUpStartTime       = $_.RampUpStartTime
                            PeakStartTime         = $_.PeakStartTime
                            RampDownStartTime     = $_.RampDownStartTime
                            OffPeakStartTime      = $_.OffPeakStartTime
                            RampUpLoadBalancingAlgorithm  = $_.RampUpLoadBalancingAlgorithm
                            PeakLoadBalancingAlgorithm    = $_.PeakLoadBalancingAlgorithm
                            RampUpMinimumHostsPct = $_.RampUpMinimumHostsPct
                            RampUpCapacityThresholdPct = $_.RampUpCapacityThresholdPct
                        }
                    })
                    HostPoolReferences = @($_.HostPoolReference)
                    Tags              = if ($_.Tag -is [hashtable]) { $_.Tag } elseif ($_.Tag) { $ht = @{}; foreach ($k in $_.Tag.Keys) { $ht[$k] = $_.Tag[$k] }; $ht } else { @{} }
                }
            })

        # Diagnostic settings on host pool
        Write-Progress -Activity 'AVD Discovery' -Status "[$hpIdx/$hpTotal] $($hp.Name) — loading diagnostics..." -PercentComplete $hpPct
        $DiagSettings = @()
        try {
            $DiagSettings = @(Get-AzDiagnosticSetting -ResourceId $hp.Id -ErrorAction SilentlyContinue | ForEach-Object {
                [PSCustomObject]@{
                    Name                     = $_.Name
                    WorkspaceId              = $_.WorkspaceId
                    StorageAccountId         = $_.StorageAccountId
                    EventHubAuthorizationRuleId = $_.EventHubAuthorizationRuleId
                    EventHubName             = $_.EventHubName
                    Logs                     = @($_.Log | ForEach-Object {
                        [PSCustomObject]@{
                            Category = $_.Category
                            Enabled  = $_.Enabled
                        }
                    })
                }
            })
        } catch { }

        [PSCustomObject]@{
            ResourceId          = $hp.Id
            Name                = $hp.Name
            ResourceGroupName   = $hpRG
            Location            = $hp.Location
            HostPoolType        = $hp.HostPoolType
            LoadBalancerType    = $hp.LoadBalancerType
            PreferredAppGroupType = $hp.PreferredAppGroupType
            MaxSessionLimit     = $hp.MaxSessionLimit
            ValidationEnvironment = $hp.ValidationEnvironment
            CustomRdpProperty   = $hp.CustomRdpProperty
            FriendlyName        = $hp.FriendlyName
            Description         = $hp.Description
            StartVMOnConnect    = $hp.StartVMOnConnect
            Tags                = if ($hp.Tag -is [hashtable]) { $hp.Tag } elseif ($hp.Tag) { $ht = @{}; foreach ($k in $hp.Tag.Keys) { $ht[$k] = $hp.Tag[$k] }; $ht } else { @{} }
            AppGroups           = @($AppGroupData)
            SessionHosts        = @($SessionHosts)
            ScalingPlans        = @($ScalingPlans)
            DiagnosticSettings  = @($DiagSettings)
        }
    }

    # Collect unique resource groups
    Write-Progress -Activity 'AVD Discovery' -Status 'Finalizing topology...' -PercentComplete 95
    $AllRGs = @()
    $AllRGs += $Topology.Workspaces | ForEach-Object { $_.ResourceGroupName }
    $AllRGs += $Topology.HostPools  | ForEach-Object { $_.ResourceGroupName }
    $Topology.ResourceGroups = @($AllRGs | Sort-Object -Unique)

    $shTotal = @($Topology.HostPools | ForEach-Object { $_.SessionHosts } | Where-Object { $_ }).Count
    Write-Progress -Activity 'AVD Discovery' -Status "Done: $(@($Topology.HostPools).Count) HP, $(@($Topology.Workspaces).Count) WS, $shTotal SH in $([int]$sw.Elapsed.TotalSeconds)s" -PercentComplete 100

    return $Topology
}

# ===============================================================================
# SECTION 15: BACKUP ENGINE
# Serializes topology to versioned JSON with SHA256 integrity sidecar. Loads and validates
# all backups from disk with integrity verification and resource count summaries.
# ===============================================================================

<#
.SYNOPSIS
    Serializes the current AVD topology to a versioned JSON backup with SHA256 integrity sidecar.
.DESCRIPTION
    Wraps the topology in a bundle with schema version, tool version, timestamp, subscription
    metadata, and optional label. Writes JSON to the backups directory with a filename encoding
    the subscription ID and timestamp. Generates a SHA256 hash sidecar file for integrity
    verification on load.
.PARAMETER Topology
    The AVD topology object returned by Get-AvdTopology.
.PARAMETER Label
    Optional human-readable label for the backup (e.g. 'pre-migration snapshot').
.OUTPUTS
    String path to the saved backup file.
#>
function Save-AvdBackup {
    param(
        [Parameter(Mandatory)]$Topology,
        [string]$Label = ''
    )

    $SubSafe = ($Topology.SubscriptionId -replace '-','').Substring(0,12)
    $Ts      = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $FileName = "backup_${SubSafe}_${Ts}.json"
    $FilePath = Join-Path $Global:BackupDir $FileName

    $Bundle = [PSCustomObject]@{
        schemaVersion    = '1.0'
        toolVersion      = $Global:AppVersion
        timestamp        = (Get-Date).ToUniversalTime().ToString('o')
        subscriptionId   = $Topology.SubscriptionId
        subscriptionName = $Topology.SubscriptionName
        label            = $Label
        topology         = $Topology
    }

    $Json = $Bundle | ConvertTo-Json -Depth 20 -Compress:$false
    [System.IO.File]::WriteAllText($FilePath, $Json, [System.Text.Encoding]::UTF8)

    # SHA256 sidecar
    $Hash = [System.Security.Cryptography.SHA256]::Create()
    $Bytes = [System.Text.Encoding]::UTF8.GetBytes($Json)
    $HashBytes = $Hash.ComputeHash($Bytes)
    $HashHex = ($HashBytes | ForEach-Object { $_.ToString('x2') }) -join ''
    [System.IO.File]::WriteAllText("$FilePath.sha256", $HashHex, [System.Text.Encoding]::UTF8)

    Write-DebugLog "[BACKUP] Saved: $FileName ($([math]::Round($Json.Length / 1KB, 1)) KB, SHA256=$($HashHex.Substring(0,16))...)" -Level 'SUCCESS'
    return $FilePath
}

<#
.SYNOPSIS
    Loads and validates all backup files from the backups directory.
.DESCRIPTION
    Scans for backup_*.json files, parses each one, verifies SHA256 integrity against the
    sidecar file, and returns metadata including resource counts, file size, and verification
    status. Results are sorted newest-first.
.OUTPUTS
    Array of PSCustomObject with FileName, FilePath, Timestamp, resource counts, SizeKB,
    Verified (bool), and the deserialized Topology object.
#>
function Get-BackupList {
    if (-not (Test-Path $Global:BackupDir)) { return @() }
    $Backups = @(Get-ChildItem $Global:BackupDir -Filter 'backup_*.json' -File | Sort-Object LastWriteTime -Descending | ForEach-Object {
        try {
            $Meta = Get-Content $_.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
            # Verify SHA256 integrity
            $Verified = $false
            $ShaPath = $_.FullName + '.sha256'
            if (Test-Path $ShaPath) {
                $ExpectedHash = (Get-Content $ShaPath -Raw).Trim()
                $Hash = [System.Security.Cryptography.SHA256]::Create()
                $Bytes = [System.IO.File]::ReadAllBytes($_.FullName)
                $ActualHash = ($Hash.ComputeHash($Bytes) | ForEach-Object { $_.ToString('x2') }) -join ''
                $Verified = ($ActualHash -eq $ExpectedHash)
            }

            $hpCount = @($Meta.topology.HostPools).Count
            $agCount = @($Meta.topology.HostPools | ForEach-Object { $_.AppGroups } | Where-Object { $_ }).Count
            $wsCount = @($Meta.topology.Workspaces).Count
            $shCount = @($Meta.topology.HostPools | ForEach-Object { $_.SessionHosts } | Where-Object { $_ }).Count
            $spCount = @($Meta.topology.HostPools | ForEach-Object { $_.ScalingPlans } | Where-Object { $_ }).Count

            [PSCustomObject]@{
                FileName         = $_.Name
                FilePath         = $_.FullName
                Timestamp        = $Meta.timestamp
                SubscriptionId   = $Meta.subscriptionId
                SubscriptionName = $Meta.subscriptionName
                Label            = $Meta.label
                HostPools        = $hpCount
                AppGroups        = $agCount
                Workspaces       = $wsCount
                SessionHosts     = $shCount
                ScalingPlans     = $spCount
                SizeKB           = [math]::Round($_.Length / 1KB, 1)
                Verified         = $Verified
                Topology         = $Meta.topology
            }
        } catch {
            Write-DebugLog "[BACKUP] Failed to parse $($_.Name): $($_.Exception.Message)" -Level 'WARN'
        }
    })
    return $Backups
}

# ===============================================================================
# SECTION 16: DIFF ENGINE
# Compares baseline (backup) vs. current topology by flattening to resource-ID-keyed maps.
# Classifies objects as Added/Removed/Modified/Unchanged, ignores volatile properties,
# and produces per-property change summaries with human-readable before/after values.
# ===============================================================================

<#
.SYNOPSIS
    Compares a baseline topology snapshot against the current live topology to detect drift.
.DESCRIPTION
    Flattens both topologies into resource-ID-keyed lookup tables, then classifies each object
    as Added (in current only), Removed (in baseline only), Modified (property differences
    excluding volatile fields like LastHeartBeat/Status), or Unchanged. Also detects session
    host count changes per host pool. For modified objects, produces per-property change
    summaries with human-readable before/after values.
.PARAMETER Baseline
    The reference topology object (typically from a backup).
.PARAMETER Current
    The current live-discovered topology object.
.OUTPUTS
    Hashtable with Added, Removed, Modified, and Unchanged ArrayLists of diff result objects.
#>
function Compare-AvdTopology {
    param($Baseline, $Current)

    $Results = @{
        Added    = [System.Collections.ArrayList]::new()
        Removed  = [System.Collections.ArrayList]::new()
        Modified = [System.Collections.ArrayList]::new()
        Unchanged = [System.Collections.ArrayList]::new()
    }

    # Build lookup tables by resource ID
    $BaselineObjects = @{}
    $CurrentObjects  = @{}

    # Flatten topology into object list
    $FlattenTopo = {
        param($Topo)
        $Objects = @{}
        foreach ($ws in @($Topo.Workspaces)) {
            $Objects[$ws.ResourceId] = @{ Type='Workspace'; Name=$ws.Name; Data=$ws }
        }
        foreach ($hp in @($Topo.HostPools)) {
            $Objects[$hp.ResourceId] = @{ Type='HostPool'; Name=$hp.Name; Data=$hp }
            foreach ($ag in @($hp.AppGroups)) {
                $Objects[$ag.ResourceId] = @{ Type='AppGroup'; Name=$ag.Name; Data=$ag; ParentHP=$hp.Name }
                foreach ($app in @($ag.Applications)) {
                    $Objects[$app.ResourceId] = @{ Type='Application'; Name=$app.Name; Data=$app; ParentAG=$ag.Name }
                }
            }
            foreach ($sp in @($hp.ScalingPlans)) {
                $Objects[$sp.ResourceId] = @{ Type='ScalingPlan'; Name=$sp.Name; Data=$sp; ParentHP=$hp.Name }
            }
        }
        return $Objects
    }

    $BaselineObjects = & $FlattenTopo $Baseline
    $CurrentObjects  = & $FlattenTopo $Current

    # Volatile properties to ignore in comparison
    $VolatileProps = @('LastHeartBeat','Status','UpdateState','Timestamp','provisioningState')

    # Find removed (in baseline, not in current)
    foreach ($rid in $BaselineObjects.Keys) {
        if (-not $CurrentObjects.ContainsKey($rid)) {
            [void]$Results.Removed.Add([PSCustomObject]@{
                ObjectType = $BaselineObjects[$rid].Type
                ResourceId = $rid
                ObjectName = $BaselineObjects[$rid].Name
                Details    = "Object exists in backup but not in current environment"
            })
        }
    }

    # Find added (in current, not in baseline)
    foreach ($rid in $CurrentObjects.Keys) {
        if (-not $BaselineObjects.ContainsKey($rid)) {
            [void]$Results.Added.Add([PSCustomObject]@{
                ObjectType = $CurrentObjects[$rid].Type
                ResourceId = $rid
                ObjectName = $CurrentObjects[$rid].Name
                Details    = "Object exists in current environment but not in backup"
            })
        }
    }

    # Find modified (in both, compare key properties)
    foreach ($rid in $BaselineObjects.Keys) {
        if ($CurrentObjects.ContainsKey($rid)) {
            $bObj = $BaselineObjects[$rid].Data
            $cObj = $CurrentObjects[$rid].Data
            $objName = $BaselineObjects[$rid].Name
            $objType = $BaselineObjects[$rid].Type

            Write-DebugLog "[DIFF] Comparing $objType '$objName'" -Level 'DEBUG'

            # Per-property comparison (excluding volatile + child-collection props)
            $ChangedProps = [System.Collections.ArrayList]::new()
            # Child collections are either compared individually (AppGroups, ScalingPlans) or
            # contain volatile nested data (SessionHosts heartbeat/status). Exclude from parent diff.
            $ChildProps = @('SessionHosts','AppGroups','ScalingPlans','DiagnosticSettings','Applications','RbacAssignments')
            $ExcludeProps = $VolatileProps + $ChildProps
            $bProps = $bObj | Select-Object -Property * -ExcludeProperty $ExcludeProps
            $cProps = $cObj | Select-Object -Property * -ExcludeProperty $ExcludeProps

            $allPropNames = @($bProps.PSObject.Properties.Name)
            Write-DebugLog "[DIFF]   Properties to compare ($($allPropNames.Count)): $($allPropNames -join ', ')" -Level 'DEBUG'

            foreach ($prop in $bProps.PSObject.Properties) {
                $pn = $prop.Name
                $bVal = $prop.Value
                $cVal = $cProps.$pn
                # Normalize Tags: Azure API returns TrackedResourceTags, JSON round-trip returns PSCustomObject/hashtable
                if ($pn -eq 'Tags') {
                    $normTag = {
                        param($v)
                        if ($null -eq $v) { return '' }
                        $pairs = @()
                        if ($v -is [hashtable]) {
                            foreach ($k in ($v.Keys | Sort-Object)) { $pairs += "$k=$($v[$k])" }
                        } elseif ($v -is [System.Collections.IDictionary]) {
                            # TrackedResourceTags implements IDictionary
                            foreach ($k in ($v.Keys | Sort-Object)) { $pairs += "$k=$($v[$k])" }
                        } else {
                            # PSCustomObject from JSON — only include real tag properties, skip framework props
                            $tagProps = @($v.PSObject.Properties | Where-Object {
                                $_.Name -notin @('AdditionalProperties','Count','Keys','Values','IsReadOnly','IsFixedSize','SyncRoot','IsSynchronized')
                            } | Sort-Object Name)
                            foreach ($p in $tagProps) { $pairs += "$($p.Name)=$($p.Value)" }
                        }
                        $pairs -join '|'
                    }
                    $bJson = & $normTag $bVal
                    $cJson = & $normTag $cVal
                } else {
                    # Detect collection metadata objects (old backups where List[String] was
                    # serialized as {"Length":N} instead of a JSON array) and treat as unknown array.
                    $collMetaNames = @('Length','LongLength','Count','Capacity','Rank','IsReadOnly','IsFixedSize','SyncRoot','IsSynchronized')
                    $normVal = {
                        param($v)
                        if ($v -is [PSCustomObject]) {
                            $pn2 = @($v.PSObject.Properties.Name)
                            if ($pn2.Count -gt 0 -and ($pn2 | Where-Object { $_ -notin $collMetaNames }).Count -eq 0) {
                                # Collection metadata — can't recover items; return marker JSON
                                $n = if ($v.PSObject.Properties.Match('Count').Count) { $v.Count }
                                     elseif ($v.PSObject.Properties.Match('Length').Count) { $v.Length }
                                     else { 0 }
                                return "{`"__collectionMeta`":true,`"count`":$n}"
                            }
                        }
                        return (ConvertTo-Json -InputObject $v -Depth 10 -Compress -ErrorAction SilentlyContinue) -replace '^\s*$',''
                    }
                    # Use -InputObject to prevent pipeline enumeration of arrays
                    $bJson = & $normVal $bVal
                    $cJson = & $normVal $cVal
                }
                if ($bJson -ne $cJson) {
                    # Debug: log the raw JSON for each differing property
                    $bSnip = if ($bJson.Length -gt 120) { $bJson.Substring(0,120) + '...' } else { $bJson }
                    $cSnip = if ($cJson.Length -gt 120) { $cJson.Substring(0,120) + '...' } else { $cJson }
                    Write-DebugLog "[DIFF]   CHANGED '$pn': baseline=[$bSnip] current=[$cSnip]" -Level 'DEBUG'
                    # Also log types for troubleshooting
                    $bType = if ($null -ne $bVal) { $bVal.GetType().Name } else { 'null' }
                    $cType = if ($null -ne $cVal) { $cVal.GetType().Name } else { 'null' }
                    Write-DebugLog "[DIFF]     Types: baseline=$bType current=$cType" -Level 'DEBUG'
                    $beforeExact = if ([string]::IsNullOrWhiteSpace($bJson)) { '(empty)' } else { $bJson }
                    $afterExact  = if ([string]::IsNullOrWhiteSpace($cJson)) { '(empty)' } else { $cJson }
                    if ($beforeExact.Length -gt 600) { $beforeExact = $beforeExact.Substring(0,600) + '...' }
                    if ($afterExact.Length -gt 600)  { $afterExact  = $afterExact.Substring(0,600) + '...' }
                    # Build a human-readable summary
                    $summary = switch -Wildcard ($pn) {
                        'ApplicationGroupReferences' {
                            # Normalize: if either value is a collection metadata PSCustomObject
                            # (from old backup serialization), treat as empty ref list for summary
                            $bIsMeta = ($bVal -is [PSCustomObject]) -and (@($bVal.PSObject.Properties.Name) | Where-Object { $_ -notin @('Length','LongLength','Count','Capacity','Rank','IsReadOnly','IsFixedSize','SyncRoot','IsSynchronized') }).Count -eq 0
                            $bRefs = if ($bIsMeta) { @() } else { @($bVal) }
                            $cRefs = @($cVal)
                            $added = @($cRefs | Where-Object { $_ -notin $bRefs }) | ForEach-Object { ($_ -split '/')[-1] }
                            $removed = @($bRefs | Where-Object { $_ -notin $cRefs }) | ForEach-Object { ($_ -split '/')[-1] }
                            $parts = @()
                            if ($added.Count -gt 0) { $parts += "+$($added -join ', ')" }
                            if ($removed.Count -gt 0) { $parts += "-$($removed -join ', ')" }
                            if ($parts.Count -gt 0) { $parts -join '; ' } else { "$($bRefs.Count) -> $($cRefs.Count) refs" }
                        }
                        'Tags' {
                            $bCount = if ($bVal) { ($bVal | Get-Member -MemberType NoteProperty).Count } else { 0 }
                            $cCount = if ($cVal) { ($cVal | Get-Member -MemberType NoteProperty).Count } else { 0 }
                            "$bCount -> $cCount tags"
                        }
                        default {
                            # Scalar values — show short before/after
                            $bStr = if ($null -eq $bVal) { '(null)' } else { "$bVal" }
                            $cStr = if ($null -eq $cVal) { '(null)' } else { "$cVal" }
                            if ($bStr.Length -gt 40) { $bStr = $bStr.Substring(0,37) + '...' }
                            if ($cStr.Length -gt 40) { $cStr = $cStr.Substring(0,37) + '...' }
                            "$bStr -> $cStr"
                        }
                    }
                    [void]$ChangedProps.Add([PSCustomObject]@{
                        Property = $pn
                        Summary  = $summary
                        Before   = $beforeExact
                        After    = $afterExact
                    })
                }
            }

            if ($ChangedProps.Count -gt 0) {
                Write-DebugLog "[DIFF]   RESULT: $objType '$objName' -> MODIFIED ($($ChangedProps.Count) props: $($ChangedProps | ForEach-Object { $_.Property }))" -Level 'INFO'
                [void]$Results.Modified.Add([PSCustomObject]@{
                    ObjectType       = $BaselineObjects[$rid].Type
                    ResourceId       = $rid
                    ObjectName       = $BaselineObjects[$rid].Name
                    Details          = ($ChangedProps | ForEach-Object { "$($_.Property): $($_.Summary)" }) -join '; '
                    ChangedProperties = $ChangedProps
                })
            } else {
                Write-DebugLog "[DIFF]   RESULT: $objType '$objName' -> UNCHANGED" -Level 'DEBUG'
                [void]$Results.Unchanged.Add([PSCustomObject]@{
                    ObjectType = $BaselineObjects[$rid].Type
                    ResourceId = $rid
                    ObjectName = $BaselineObjects[$rid].Name
                })
            }
        }
    }

    # Also check session host count changes per host pool
    foreach ($bHP in @($Baseline.HostPools)) {
        $cHP = $Current.HostPools | Where-Object { $_.ResourceId -eq $bHP.ResourceId }
        if ($cHP) {
            $bSHCount = @($bHP.SessionHosts).Count
            $cSHCount = @($cHP.SessionHosts).Count
            if ($bSHCount -ne $cSHCount) {
                [void]$Results.Modified.Add([PSCustomObject]@{
                    ObjectType = 'SessionHostCount'
                    ResourceId = $bHP.ResourceId
                    ObjectName = "$($bHP.Name) session hosts"
                    Details    = "Count changed: $bSHCount -> $cSHCount"
                })
            }
        }
    }

    Write-DebugLog "[DIFF] Results: Added=$($Results.Added.Count), Removed=$($Results.Removed.Count), Modified=$($Results.Modified.Count), Unchanged=$($Results.Unchanged.Count)" -Level 'INFO'
    return $Results
}

# ===============================================================================
# SECTION 17: RESTORE ENGINE
# Builds dependency-ordered restore manifests from diff results, executes Create/Update
# actions via Az cmdlets with dry-run validation, and tracks success/failure per item.
# Supports cross-region location override and resource group remapping.
# ===============================================================================

# Restore dependency order: RG -> HostPool -> AppGroup -> Applications -> Workspace -> RBAC -> ScalingPlan -> DiagSettings
$Global:RestoreOrder = @('ResourceGroup','HostPool','AppGroup','Application','Workspace','RbacAssignment','ScalingPlan','DiagnosticSetting','SessionHost')

<#
.SYNOPSIS
    Builds a dependency-ordered restore action manifest from diff results.
.DESCRIPTION
    Converts Removed items to Create actions and Modified items to Update actions, assigns
    dependency sort order (RG -> HostPool -> AppGroup -> Application -> Workspace -> ScalingPlan),
    and enriches each entry with the corresponding backup data for the restore cmdlets.
.PARAMETER DiffResults
    Diff results hashtable from Compare-AvdTopology.
.PARAMETER BackupTopology
    The backup topology object used to look up full resource definitions for restore.
.OUTPUTS
    ArrayList of manifest entry objects with Action, ObjectType, ResourceId, SortOrder,
    Selected flag, BackupData, and optional ChangedProperties.
#>
function Build-RestoreManifest {
    param($DiffResults, $BackupTopology)

    $Manifest = [System.Collections.ArrayList]::new()
    $TypeMap = @{
        'Workspace'    = 6  # after RBAC
        'HostPool'     = 1
        'AppGroup'     = 2
        'Application'  = 3
        'ScalingPlan'  = 4
        'SessionHostCount' = 5
    }

    # Removed items need restore (re-create)
    foreach ($item in $DiffResults.Removed) {
        $ord = if ($TypeMap.ContainsKey($item.ObjectType)) { $TypeMap[$item.ObjectType] } else { 99 }
        [void]$Manifest.Add([PSCustomObject]@{
            Action      = 'Create'
            ObjectType  = $item.ObjectType
            ResourceId  = $item.ResourceId
            ObjectName  = $item.ObjectName
            SortOrder   = $ord
            Selected    = $true
            BackupData  = $null
        })
    }

    # Modified items need update
    foreach ($item in $DiffResults.Modified) {
        $ord = if ($TypeMap.ContainsKey($item.ObjectType)) { $TypeMap[$item.ObjectType] } else { 99 }
        [void]$Manifest.Add([PSCustomObject]@{
            Action            = 'Update'
            ObjectType        = $item.ObjectType
            ResourceId        = $item.ResourceId
            ObjectName        = $item.ObjectName
            SortOrder         = $ord
            Selected          = $true
            BackupData        = $null
            ChangedProperties = if ($item.PSObject.Properties.Match('ChangedProperties').Count) { $item.ChangedProperties } else { @() }
        })
    }

    # Enrich manifest entries with backup data
    foreach ($entry in $Manifest) {
        switch ($entry.ObjectType) {
            'HostPool' {
                $entry.BackupData = $BackupTopology.HostPools | Where-Object { $_.ResourceId -eq $entry.ResourceId }
            }
            'AppGroup' {
                foreach ($hp in $BackupTopology.HostPools) {
                    $ag = $hp.AppGroups | Where-Object { $_.ResourceId -eq $entry.ResourceId }
                    if ($ag) { $entry.BackupData = $ag; break }
                }
            }
            'Application' {
                foreach ($hp in $BackupTopology.HostPools) {
                    foreach ($ag in $hp.AppGroups) {
                        $app = $ag.Applications | Where-Object { $_.ResourceId -eq $entry.ResourceId }
                        if ($app) { $entry.BackupData = $app; break }
                    }
                }
            }
            'Workspace' {
                $entry.BackupData = $BackupTopology.Workspaces | Where-Object { $_.ResourceId -eq $entry.ResourceId }
            }
            'ScalingPlan' {
                foreach ($hp in $BackupTopology.HostPools) {
                    $sp = $hp.ScalingPlans | Where-Object { $_.ResourceId -eq $entry.ResourceId }
                    if ($sp) { $entry.BackupData = $sp; break }
                }
            }
        }
    }

    # Sort by dependency order
    $Manifest = [System.Collections.ArrayList]@($Manifest | Sort-Object SortOrder)
    return $Manifest
}

<#
.SYNOPSIS
    Exports the selected restore manifest entries as a JSON plan file via SaveFileDialog.
.DESCRIPTION
    Filters the manifest to selected items, wraps them in a schema-versioned plan with export
    timestamp and action counts, and writes to the user-chosen file path. Unlocks the Exporter
    achievement on success.
.PARAMETER Manifest
    The restore manifest ArrayList from Build-RestoreManifest.
#>
function Export-RestorePlan {
    param([System.Collections.ArrayList]$Manifest)

    $Selected = @($Manifest | Where-Object Selected)
    if ($Selected.Count -eq 0) {
        Show-Toast -Message 'No items selected for export.' -Type 'Warning'
        return
    }

    $Dialog = [Microsoft.Win32.SaveFileDialog]::new()
    $Dialog.Title = 'Export Restore Plan'
    $Dialog.InitialDirectory = $Global:BackupDir
    $Dialog.Filter = 'JSON Files (*.json)|*.json|Text Files (*.txt)|*.txt'
    $Dialog.DefaultExt = '.json'
    $Dialog.FileName = "RestorePlan_$(Get-Date -Format 'yyyyMMdd_HHmmss')"

    if ($Dialog.ShowDialog() -ne $true) { return }

    try {
        $Plan = [PSCustomObject]@{
            schemaVersion = '1.0'
            exportedAt    = (Get-Date).ToUniversalTime().ToString('o')
            totalItems    = $Selected.Count
            addCount      = @($Selected | Where-Object Action -eq 'Create').Count
            updateCount   = @($Selected | Where-Object Action -eq 'Update').Count
            items         = @($Selected | Sort-Object SortOrder | ForEach-Object {
                [PSCustomObject]@{
                    order      = $_.SortOrder
                    action     = $_.Action
                    type       = $_.ObjectType
                    name       = $_.ObjectName
                    resourceId = $_.ResourceId
                    targetRegion = $_.TargetRegion
                    changedProperties = if ($_.Action -eq 'Update' -and $_.ChangedProperties -and $_.ChangedProperties.Count -gt 0) {
                        @($_.ChangedProperties | ForEach-Object {
                            [PSCustomObject]@{
                                property = $_.Property
                                summary  = $_.Summary
                                before   = $_.Before
                                after    = $_.After
                            }
                        })
                    } else { @() }
                    hasBackupData = ($null -ne $_.BackupData)
                    backupData = if ($_.BackupData) {
                        $bd = @{}
                        if ($_.BackupData.PSObject.Properties.Match('ResourceGroupName').Count) { $bd.resourceGroup = $_.BackupData.ResourceGroupName }
                        if ($_.BackupData.PSObject.Properties.Match('Location').Count) { $bd.location = $_.BackupData.Location }
                        if ($_.BackupData.PSObject.Properties.Match('HostPoolType').Count) { $bd.hostPoolType = $_.BackupData.HostPoolType }
                        if ($_.BackupData.PSObject.Properties.Match('ApplicationGroupType').Count) { $bd.appGroupType = $_.BackupData.ApplicationGroupType }
                        if ($_.BackupData.PSObject.Properties.Match('FriendlyName').Count) { $bd.friendlyName = $_.BackupData.FriendlyName }
                        $bd
                    } else { $null }
                }
            })
        }

        $Json = $Plan | ConvertTo-Json -Depth 10
        [System.IO.File]::WriteAllText($Dialog.FileName, $Json, [System.Text.Encoding]::UTF8)
        Write-DebugLog "[EXPORT] Restore plan saved: $($Dialog.FileName)" -Level 'SUCCESS'
        Show-Toast -Message "Plan exported: $(Split-Path $Dialog.FileName -Leaf)" -Type 'Success'
        Check-Achievement 'Exporter'
    } catch {
        Write-DebugLog "[EXPORT] Plan export failed: $($_.Exception.Message)" -Level 'ERROR'
        Show-Toast -Message "Export failed: $($_.Exception.Message)" -Type 'Error'
    }
}

<#
.SYNOPSIS
    Executes a single restore action (Create or Update) for one AVD resource, with dry-run support.
.DESCRIPTION
    Dispatches the restore operation by ObjectType (HostPool, AppGroup, Application, Workspace,
    ScalingPlan, SessionHostCount) using the appropriate Az cmdlet (New-AzWvd* or Update-AzWvd*).
    In DryRun mode, validates required fields without making API calls. Supports cross-region
    location override via ManifestEntry.TargetRegion.
.PARAMETER ManifestEntry
    A single manifest entry object from Build-RestoreManifest.
.PARAMETER DryRun
    When set, performs validation only without making Azure API calls.
.OUTPUTS
    PSCustomObject with Success (bool), Message (string), and Item (original entry).
#>
function Invoke-RestoreItem {
    param(
        [PSCustomObject]$ManifestEntry,
        [switch]$DryRun
    )

    $Type   = $ManifestEntry.ObjectType
    $Action = $ManifestEntry.Action
    $Data   = $ManifestEntry.BackupData
    $Prefix = '[RESTORE]'

    # Override location if cross-region restore is selected
    if ($ManifestEntry.TargetRegion -and $Data -and $Data.PSObject.Properties.Match("Location").Count) {
        $Data.Location = $ManifestEntry.TargetRegion
    }

    if (-not $Data -and $Type -ne 'SessionHostCount') {
        Write-DebugLog "$Prefix Skipping $Type '$($ManifestEntry.ObjectName)' - no backup data" -Level 'WARN'
        return [PSCustomObject]@{ Success=$false; Message="No backup data"; Item=$ManifestEntry }
    }

    Write-DebugLog "$Prefix $Action $Type '$($ManifestEntry.ObjectName)' ..." -Level 'INFO'

    if ($DryRun) {
        # Validate the manifest entry would succeed
        $Issues = [System.Collections.ArrayList]::new()
        if (-not $Data -and $Type -ne 'SessionHostCount') {
            [void]$Issues.Add('Missing backup data')
        }
        if ($Data -and $Data.PSObject.Properties.Match('ResourceGroupName').Count -and -not $Data.ResourceGroupName) {
            [void]$Issues.Add('Missing ResourceGroupName')
        }
        if ($Data -and $Data.PSObject.Properties.Match('Location').Count -and -not $Data.Location) {
            [void]$Issues.Add('Missing Location')
        }
        if ($Type -eq 'HostPool' -and $Data) {
            if (-not $Data.HostPoolType) { [void]$Issues.Add('Missing HostPoolType') }
            if (-not $Data.LoadBalancerType) { [void]$Issues.Add('Missing LoadBalancerType') }
            if ($Action -eq 'Create' -and -not $Data.PreferredAppGroupType) { [void]$Issues.Add('Missing PreferredAppGroupType') }
        }
        if ($Type -eq 'AppGroup' -and $Data) {
            if (-not $Data.ApplicationGroupType) { [void]$Issues.Add('Missing ApplicationGroupType') }
            if (-not $Data.HostPoolArmPath) { [void]$Issues.Add('Missing HostPoolArmPath') }
        }
        if ($Type -eq 'Application' -and $Data) {
            if (-not $Data.ResourceId) { [void]$Issues.Add('Missing ResourceId for AppGroup parsing') }
            if (-not $Data.Name) { [void]$Issues.Add('Missing application Name') }
        }
        if ($Type -eq 'ScalingPlan' -and $Data -and $Action -eq 'Create') {
            if (-not $Data.TimeZone) { [void]$Issues.Add('Missing TimeZone') }
        }

        if ($Issues.Count -gt 0) {
            $msg = "Dry-run FAIL: $Action $Type '$($ManifestEntry.ObjectName)' - $($Issues -join '; ')"
            Write-DebugLog $msg -Level 'WARN'
            return [PSCustomObject]@{ Success=$false; Message=($Issues -join '; '); Item=$ManifestEntry }
        }
        $msg = "Dry-run OK: would $Action $Type '$($ManifestEntry.ObjectName)'"
        if ($Data.ResourceGroupName) { $msg += " in RG $($Data.ResourceGroupName)" }
        if ($Data.Location) { $msg += " ($($Data.Location))" }
        Write-DebugLog $msg -Level 'DEBUG'
        return [PSCustomObject]@{ Success=$true; Message=$msg; Item=$ManifestEntry }
    }

    try {
        switch ($Type) {
            'HostPool' {
                if ($Action -eq 'Create') {
                    $Params = @{
                        ResourceGroupName     = $Data.ResourceGroupName
                        Name                  = $Data.Name
                        Location              = $Data.Location
                        HostPoolType          = $Data.HostPoolType
                        LoadBalancerType      = $Data.LoadBalancerType
                        PreferredAppGroupType = $Data.PreferredAppGroupType
                    }
                    if ($null -ne $Data.MaxSessionLimit) { $Params.MaxSessionLimit = $Data.MaxSessionLimit }
                    if ($Data.ValidationEnvironment) { $Params.ValidationEnvironment = $Data.ValidationEnvironment }
                    if ($Data.CustomRdpProperty) { $Params.CustomRdpProperty = $Data.CustomRdpProperty }
                    if ($Data.FriendlyName) { $Params.FriendlyName = $Data.FriendlyName }
                    if ($Data.Description) { $Params.Description = $Data.Description }
                    if ($Data.StartVMOnConnect) { $Params.StartVMOnConnect = $Data.StartVMOnConnect }
                    $null = New-AzWvdHostPool @Params -ErrorAction Stop
                } else {
                    $Params = @{
                        ResourceGroupName = $Data.ResourceGroupName
                        Name              = $Data.Name
                    }
                    if ($Data.LoadBalancerType) { $Params.LoadBalancerType = $Data.LoadBalancerType }
                    if ($Data.PreferredAppGroupType) { $Params.PreferredAppGroupType = $Data.PreferredAppGroupType }
                    if ($null -ne $Data.MaxSessionLimit) { $Params.MaxSessionLimit = $Data.MaxSessionLimit }
                    if ($Data.ValidationEnvironment) { $Params.ValidationEnvironment = $Data.ValidationEnvironment }
                    if ($Data.CustomRdpProperty) { $Params.CustomRdpProperty = $Data.CustomRdpProperty }
                    if ($Data.FriendlyName) { $Params.FriendlyName = $Data.FriendlyName }
                    if ($Data.Description) { $Params.Description = $Data.Description }
                    $null = Update-AzWvdHostPool @Params -ErrorAction Stop
                }
            }
            'AppGroup' {
                if ($Action -eq 'Create') {
                    $Params = @{
                        ResourceGroupName    = $Data.ResourceGroupName
                        Name                 = $Data.Name
                        Location             = $Data.Location
                        ApplicationGroupType = $Data.ApplicationGroupType
                        HostPoolArmPath      = $Data.HostPoolArmPath
                    }
                    if ($Data.FriendlyName) { $Params.FriendlyName = $Data.FriendlyName }
                    if ($Data.Description) { $Params.Description = $Data.Description }
                    $null = New-AzWvdApplicationGroup @Params -ErrorAction Stop
                } else {
                    $Params = @{
                        ResourceGroupName = $Data.ResourceGroupName
                        Name              = $Data.Name
                    }
                    if ($Data.FriendlyName) { $Params.FriendlyName = $Data.FriendlyName }
                    if ($Data.Description) { $Params.Description = $Data.Description }
                    $null = Update-AzWvdApplicationGroup @Params -ErrorAction Stop
                }
            }
            'Application' {
                # Parse app group from resource ID
                $Parts = $Data.ResourceId -split '/'
                $agRG   = $Parts[4]
                $agName = $Parts[8]
                $Params = @{
                    ResourceGroupName = $agRG
                    GroupName         = $agName
                    Name              = $Data.Name
                    CommandLineSetting = if ($Data.CommandLineSetting) { $Data.CommandLineSetting } else { 'DoNotAllow' }
                }
                if ($Data.FilePath) { $Params.FilePath = $Data.FilePath }
                if ($Data.FriendlyName) { $Params.FriendlyName = $Data.FriendlyName }
                if ($Data.IconPath) { $Params.IconPath = $Data.IconPath }

                if ($Action -eq 'Create') {
                    $null = New-AzWvdApplication @Params -ErrorAction Stop
                } else {
                    $null = Update-AzWvdApplication @Params -ErrorAction Stop
                }
            }
            'Workspace' {
                $Params = @{
                    ResourceGroupName = $Data.ResourceGroupName
                    Name              = $Data.Name
                }
                if ($Action -eq 'Create' -and $Data.Location) { $Params.Location = $Data.Location }
                if ($Data.FriendlyName) { $Params.FriendlyName = $Data.FriendlyName }
                if ($Data.Description) { $Params.Description = $Data.Description }
                if ($Data.ApplicationGroupReferences) {
                    # Drop null/blank refs to avoid Az cmdlet path validation errors.
                    $cleanRefs = @($Data.ApplicationGroupReferences | ForEach-Object {
                        if ($null -eq $_) { return }
                        $val = [string]$_
                        if (-not [string]::IsNullOrWhiteSpace($val)) { $val.Trim() }
                    })
                    if ($cleanRefs.Count -gt 0) {
                        $validRefs = [System.Collections.ArrayList]::new()
                        foreach ($ref in $cleanRefs) {
                            if ($ref -notmatch '^/subscriptions/[^/]+/resourcegroups/[^/]+/providers/Microsoft\.DesktopVirtualization/applicationgroups/[^/]+$') {
                                Write-DebugLog "$Prefix Workspace '$($ManifestEntry.ObjectName)' skipping malformed app-group reference: $ref" -Level 'WARN'
                                continue
                            }
                            $parts = $ref -split '/'
                            $agRg = $parts[4]
                            $agName = $parts[8]
                            $ag = $null
                            try {
                                $ag = Get-AzWvdApplicationGroup -ResourceGroupName $agRg -Name $agName -ErrorAction SilentlyContinue
                            } catch { }
                            if ($ag -and $ag.Id) {
                                [void]$validRefs.Add($ag.Id)
                            } else {
                                Write-DebugLog "$Prefix Workspace '$($ManifestEntry.ObjectName)' app-group not found, skipping reference: $ref" -Level 'WARN'
                            }
                        }

                        if ($validRefs.Count -gt 0) {
                            $Params.ApplicationGroupReference = @($validRefs)
                            Write-DebugLog "$Prefix Workspace '$($ManifestEntry.ObjectName)' using $($validRefs.Count)/$($cleanRefs.Count) valid app-group reference(s)" -Level 'INFO'
                        } else {
                            return [PSCustomObject]@{
                                Success = $false
                                Message = 'No valid ApplicationGroupReferences remain after validation'
                                Item    = $ManifestEntry
                            }
                        }
                    } else {
                        Write-DebugLog "$Prefix Workspace '$($ManifestEntry.ObjectName)' has no valid ApplicationGroupReferences after cleanup" -Level 'WARN'
                    }
                }

                if ($Action -eq 'Create') {
                    $null = New-AzWvdWorkspace @Params -ErrorAction Stop
                } else {
                    $null = Update-AzWvdWorkspace @Params -ErrorAction Stop
                }
            }
            'ScalingPlan' {
                if ($Action -eq 'Create') {
                    $Params = @{
                        ResourceGroupName = $Data.ResourceGroupName
                        Name              = $Data.Name
                        Location          = $Data.Location
                        TimeZone          = $Data.TimeZone
                    }
                    if ($Data.HostPoolType) { $Params.HostPoolType = $Data.HostPoolType }
                    if ($Data.Schedules) { $Params.Schedule = $Data.Schedules }
                    if ($Data.HostPoolReferences) { $Params.HostPoolReference = $Data.HostPoolReferences }
                    if ($Data.FriendlyName) { $Params.FriendlyName = $Data.FriendlyName }
                    if ($Data.Description) { $Params.Description = $Data.Description }
                    $null = New-AzWvdScalingPlan @Params -ErrorAction Stop
                } else {
                    $Params = @{
                        ResourceGroupName = $Data.ResourceGroupName
                        Name              = $Data.Name
                    }
                    if ($Data.TimeZone) { $Params.TimeZone = $Data.TimeZone }
                    if ($Data.Schedules) { $Params.Schedule = $Data.Schedules }
                    if ($Data.HostPoolReferences) { $Params.HostPoolReference = $Data.HostPoolReferences }
                    if ($Data.FriendlyName) { $Params.FriendlyName = $Data.FriendlyName }
                    if ($Data.Description) { $Params.Description = $Data.Description }
                    $null = Update-AzWvdScalingPlan @Params -ErrorAction Stop
                }
            }
            'SessionHostCount' {
                # Session host count differs - re-register missing hosts
                Write-DebugLog "$Prefix Session host count mismatch for $($Data.Name)" -Level 'INFO'
                if ($Action -eq 'Update' -and $Data.SessionHosts) {
                    $hpRG = $Data.ResourceGroupName
                    $hpName = $Data.Name
                    foreach ($sh in $Data.SessionHosts) {
                        if ($sh.ResourceId) {
                            $regResult = Register-SessionHost -HostPoolName $hpName `
                                -ResourceGroupName $hpRG -VmResourceId $sh.ResourceId -DryRun:$DryRun
                            Write-DebugLog "$Prefix SessionHost $($sh.Name): $($regResult.Message)" `
                                -Level $(if ($regResult.Success) { 'SUCCESS' } else { 'WARN' })
                        }
                    }
                }
            }
            default {
                Write-DebugLog "$Prefix Unsupported type: $Type" -Level 'WARN'
                return [PSCustomObject]@{ Success=$false; Message="Unsupported type: $Type"; Item=$ManifestEntry }
            }
        }
        Write-DebugLog "$Prefix Successfully $($Action)d $Type '$($ManifestEntry.ObjectName)'" -Level 'SUCCESS'
        return [PSCustomObject]@{ Success=$true; Message="$Action completed"; Item=$ManifestEntry }
    } catch {
        Write-DebugLog "$Prefix FAILED $Type '$($ManifestEntry.ObjectName)': $($_.Exception.Message)" -Level 'ERROR'
        return [PSCustomObject]@{ Success=$false; Message=$_.Exception.Message; Item=$ManifestEntry }
    }
}

<#
.SYNOPSIS
    Executes all selected manifest items in dependency order with progress tracking.
.DESCRIPTION
    Iterates through selected manifest entries sorted by SortOrder, invokes each via
    Invoke-RestoreItem, tracks success/failure counts, and logs progress percentages.
.PARAMETER Manifest
    The restore manifest ArrayList from Build-RestoreManifest.
.PARAMETER DryRun
    When set, all items are validated without making Azure API calls.
.OUTPUTS
    ArrayList of result objects from Invoke-RestoreItem.
#>
function Invoke-FullRestore {
    param(
        [System.Collections.ArrayList]$Manifest,
        [switch]$DryRun
    )

    $TotalCount = @($Manifest | Where-Object Selected).Count
    $Current = 0
    $Results = [System.Collections.ArrayList]::new()
    $FailCount = 0

    Write-DebugLog "[RESTORE] Starting restore: $TotalCount items (DryRun=$DryRun)" -Level 'INFO'

    foreach ($entry in ($Manifest | Where-Object Selected | Sort-Object SortOrder)) {
        $Current++
        $PctComplete = [math]::Round(($Current / $TotalCount) * 100)
        Write-DebugLog "[RESTORE] [$Current/$TotalCount] ($PctComplete%) Processing $($entry.ObjectType) '$($entry.ObjectName)'" -Level 'INFO'

        $result = Invoke-RestoreItem -ManifestEntry $entry -DryRun:$DryRun

        # Normalize any accidental multi-object output into a single restore result record.
        $normalized = $null
        if ($null -eq $result) {
            $normalized = [PSCustomObject]@{
                Success = $false
                Message = 'Restore item returned no result object'
                Item    = $entry
            }
        } elseif ($result -is [array]) {
            $candidates = @($result | Where-Object { $_ -and $_.PSObject.Properties.Match('Success').Count -gt 0 })
            if ($candidates.Count -gt 0) {
                $normalized = $candidates[-1]
            } else {
                $normalized = [PSCustomObject]@{
                    Success = $false
                    Message = "Unexpected restore result type(s): $((@($result | ForEach-Object { $_.GetType().FullName }) -join ', '))"
                    Item    = $entry
                }
            }
        } elseif ($result.PSObject.Properties.Match('Success').Count -gt 0) {
            $normalized = $result
        } else {
            $normalized = [PSCustomObject]@{
                Success = $false
                Message = "Unexpected restore result type: $($result.GetType().FullName)"
                Item    = $entry
            }
        }

        [void]$Results.Add($normalized)

        if (-not $normalized.Success) { $FailCount++ }
    }

    Write-DebugLog "[RESTORE] Complete: $($TotalCount - $FailCount)/$TotalCount succeeded, $FailCount failed" -Level $(if ($FailCount -gt 0) { 'WARN' } else { 'SUCCESS' })
    return $Results
}

# ===============================================================================
# SECTION 18: SESSION HOST REGISTRATION
# Generates AVD registration tokens and pushes them to VMs via Invoke-AzVMRunCommand
# to re-register session hosts with their host pools after restore operations.
# ===============================================================================

<#
.SYNOPSIS
    Generates a registration token and re-registers a VM as a session host via Invoke-AzVMRunCommand.
.DESCRIPTION
    Creates a 4-hour registration token for the specified host pool, then pushes a PowerShell
    script to the target VM that writes the token to the RDInfraAgent registry key, resets the
    IsRegistered flag, and restarts the RDAgentBootLoader service.
.PARAMETER HostPoolName
    Name of the target host pool.
.PARAMETER ResourceGroupName
    Resource group containing the host pool.
.PARAMETER VmResourceId
    Full ARM resource ID of the VM to register.
.PARAMETER DryRun
    When set, returns validation result without executing the registration.
.OUTPUTS
    PSCustomObject with Success (bool) and Message (string).
#>
function Register-SessionHost {
    param(
        [Parameter(Mandatory)][string]$HostPoolName,
        [Parameter(Mandatory)][string]$ResourceGroupName,
        [Parameter(Mandatory)][string]$VmResourceId,
        [switch]$DryRun
    )

    $Prefix = if ($DryRun) { '[DRYRUN]' } else { '[REG]' }
    Write-DebugLog "$Prefix Registering VM to host pool $HostPoolName" -Level 'INFO'

    if ($DryRun) {
        return [PSCustomObject]@{ Success=$true; Message="Dry-run: would register VM" }
    }

    try {
        # Generate new registration token (valid 4 hours)
        $Token = New-AzWvdRegistrationInfo -ResourceGroupName $ResourceGroupName -HostPoolName $HostPoolName `
            -ExpirationTime ((Get-Date).AddHours(4)) -ErrorAction Stop

        if (-not $Token.Token) {
            throw "Failed to generate registration token"
        }

        $RegToken = $Token.Token

        # Install/re-register agent via Run Command
        $ScriptContent = @"
`$RegPath = 'HKLM:\SOFTWARE\Microsoft\RDInfraAgent'
if (Test-Path `$RegPath) {
    Set-ItemProperty -Path `$RegPath -Name RegistrationToken -Value '$RegToken'
    Set-ItemProperty -Path `$RegPath -Name IsRegistered -Value 0
    Restart-Service -Name RDAgentBootLoader -Force
    Write-Output 'Re-registration initiated via existing agent'
} else {
    Write-Output 'RDInfraAgent not installed. Manual agent installation required.'
}
"@

        $vmName = ($VmResourceId -split '/')[-1]
        $vmRG   = ($VmResourceId -split '/')[4]

        $RunResult = Invoke-AzVMRunCommand -ResourceGroupName $vmRG -VMName $vmName `
            -CommandId 'RunPowerShellScript' -ScriptString $ScriptContent -ErrorAction Stop

        $Output = $RunResult.Value | Where-Object { $_.Code -eq 'ComponentStatus/StdOut/succeeded' } | Select-Object -ExpandProperty Message

        Write-DebugLog "$Prefix VM $vmName registration result: $Output" -Level 'SUCCESS'
        return [PSCustomObject]@{ Success=$true; Message=$Output }
    } catch {
        Write-DebugLog "$Prefix Registration failed for $vmName : $($_.Exception.Message)" -Level 'ERROR'
        return [PSCustomObject]@{ Success=$false; Message=$_.Exception.Message }
    }
}


# ===============================================================================
# SECTION 18b: RG MAPPING & VM PATTERN DISCOVERY
# Resource group mapping UI for cross-environment restores (source RG -> target RG textboxes).
# VM pattern discovery finds session host VMs matching naming conventions beyond backup data.
# ===============================================================================

<#
.SYNOPSIS
    Builds the resource group mapping UI grid for cross-environment restores.
.DESCRIPTION
    Collects unique resource group names from all host pools, app groups, scaling plans, and
    workspaces in the backup. Renders a source-label to editable-textbox row for each RG,
    allowing users to remap source RGs to different target RGs before restore execution.
.PARAMETER BackupTopology
    The topology object from a loaded backup file.
#>
function Populate-RgMappingGrid {
    param($BackupTopology)
    $Window.Dispatcher.Invoke([Action]{
        $pnlRgMapping.Children.Clear()
        $Global:RgMappings = @{}

        # Collect unique RGs from backup
        $rgs = [System.Collections.ArrayList]::new()
        foreach ($hp in @($BackupTopology.HostPools)) {
            if ($hp.ResourceGroupName -and $rgs -notcontains $hp.ResourceGroupName) { [void]$rgs.Add($hp.ResourceGroupName) }
            foreach ($ag in @($hp.AppGroups)) {
                $agRG = if ($ag.ResourceGroupName) { $ag.ResourceGroupName } elseif ($ag.ResourceId) { ($ag.ResourceId -split "/")[4] } else { $null }
                if ($agRG -and $rgs -notcontains $agRG) { [void]$rgs.Add($agRG) }
            }
            foreach ($sp in @($hp.ScalingPlans)) {
                if ($sp.ResourceGroupName -and $rgs -notcontains $sp.ResourceGroupName) { [void]$rgs.Add($sp.ResourceGroupName) }
            }
        }
        foreach ($ws in @($BackupTopology.Workspaces)) {
            $wsRG = if ($ws.ResourceGroupName) { $ws.ResourceGroupName } elseif ($ws.ResourceId) { ($ws.ResourceId -split "/")[4] } else { $null }
            if ($wsRG -and $rgs -notcontains $wsRG) { [void]$rgs.Add($wsRG) }
        }

        if ($rgs.Count -eq 0) {
            $lblRgMappingHint.Text = "No resource groups found in backup."
            $lblRgMappingHint.Visibility = "Visible"
            return
        }

        $lblRgMappingHint.Visibility = "Collapsed"
        foreach ($rg in ($rgs | Sort-Object)) {
            $Global:RgMappings[$rg] = $rg  # default: same name

            $row = [System.Windows.Controls.Grid]::new()
            $col1 = [System.Windows.Controls.ColumnDefinition]::new(); $col1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
            $col2 = [System.Windows.Controls.ColumnDefinition]::new(); $col2.Width = [System.Windows.GridLength]::new(14, [System.Windows.GridUnitType]::Pixel)
            $col3 = [System.Windows.Controls.ColumnDefinition]::new(); $col3.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
            [void]$row.ColumnDefinitions.Add($col1); [void]$row.ColumnDefinitions.Add($col2); [void]$row.ColumnDefinitions.Add($col3)
            $row.Margin = [System.Windows.Thickness]::new(0,0,0,4)

            # Source RG label
            $srcLbl = [System.Windows.Controls.TextBlock]::new()
            $srcLbl.Text = $rg
            $srcLbl.FontSize = 10
            $srcLbl.TextTrimming = [System.Windows.TextTrimming]::CharacterEllipsis
            $srcLbl.VerticalAlignment = "Center"
            $srcLbl.ToolTip = $rg
            $srcLbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, "ThemeTextDim")
            [System.Windows.Controls.Grid]::SetColumn($srcLbl, 0)
            [void]$row.Children.Add($srcLbl)

            # Arrow
            $arrow = [System.Windows.Controls.TextBlock]::new()
            $arrow.Text = [string]([char]0xE72A)
            $arrow.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
            $arrow.FontSize = 9
            $arrow.HorizontalAlignment = "Center"; $arrow.VerticalAlignment = "Center"
            $arrow.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, "ThemeTextDisabled")
            [System.Windows.Controls.Grid]::SetColumn($arrow, 1)
            [void]$row.Children.Add($arrow)

            # Target RG text box
            $tgtBox = [System.Windows.Controls.TextBox]::new()
            $tgtBox.Text = $rg
            $tgtBox.Tag = $rg  # original RG name as key
            $tgtBox.FontSize = 10
            $tgtBox.Padding = [System.Windows.Thickness]::new(4,2,4,2)
            $tgtBox.SetResourceReference([System.Windows.Controls.TextBox]::BackgroundProperty, "ThemeInputBg")
            $tgtBox.SetResourceReference([System.Windows.Controls.TextBox]::ForegroundProperty, "ThemeTextBody")
            $tgtBox.SetResourceReference([System.Windows.Controls.TextBox]::BorderBrushProperty, "ThemeBorderCard")
            $tgtBox.BorderThickness = [System.Windows.Thickness]::new(1)
            $tgtBox.ToolTip = "Edit to map $rg to a different target resource group"
            [System.Windows.Controls.Grid]::SetColumn($tgtBox, 2)
            [void]$row.Children.Add($tgtBox)

            # Track changes
            $tgtBox.Add_LostFocus({
                param($s, $e)
                $sourceRg = $s.Tag
                $targetRg = $s.Text.Trim()
                if ($targetRg) { $Global:RgMappings[$sourceRg] = $targetRg }
            }.GetNewClosure())

            [void]$pnlRgMapping.Children.Add($row)
        }
        Write-DebugLog "[RG-MAP] Populated $($rgs.Count) resource group mapping(s)" -Level "DEBUG"
    })
}

<#
.SYNOPSIS
    Applies resource group remappings to all manifest entries before restore execution.
.DESCRIPTION
    Flushes pending textbox changes from the UI, then rewrites ResourceGroupName, ResourceId
    paths, HostPoolArmPath, ApplicationGroupReferences, and HostPoolReferences in each manifest
    entry BackupData according to the $Global:RgMappings lookup table.
.PARAMETER Manifest
    The restore manifest ArrayList whose BackupData objects will be mutated in-place.
#>
function Apply-RgMappings {
    param([System.Collections.ArrayList]$Manifest)

    if (-not $Global:RgMappings -or $Global:RgMappings.Count -eq 0) { return }

    # Flush any pending text box changes
    $Window.Dispatcher.Invoke([Action]{
        foreach ($child in $pnlRgMapping.Children) {
            if ($child -is [System.Windows.Controls.Grid]) {
                foreach ($gc in $child.Children) {
                    if ($gc -is [System.Windows.Controls.TextBox]) {
                        $srcRg = $gc.Tag
                        $tgtRg = $gc.Text.Trim()
                        if ($tgtRg) { $Global:RgMappings[$srcRg] = $tgtRg }
                    }
                }
            }
        }
    })

    $changed = 0
    foreach ($entry in $Manifest) {
        if (-not $entry.BackupData) { continue }
        $bd = $entry.BackupData
        # Map ResourceGroupName
        if ($bd.PSObject.Properties.Match("ResourceGroupName").Count -and $bd.ResourceGroupName) {
            $src = $bd.ResourceGroupName
            if ($Global:RgMappings.ContainsKey($src) -and $Global:RgMappings[$src] -ne $src) {
                $bd.ResourceGroupName = $Global:RgMappings[$src]
                $changed++
                Write-DebugLog "[RG-MAP] $($entry.ObjectName): $src -> $($bd.ResourceGroupName)" -Level "DEBUG"
            }
        }
        # Map ResourceId path (for Application type which parses RG from ResourceId)
        if ($bd.PSObject.Properties.Match("ResourceId").Count -and $bd.ResourceId) {
            foreach ($mapping in $Global:RgMappings.GetEnumerator()) {
                if ($mapping.Key -ne $mapping.Value -and $bd.ResourceId -match "/resourceGroups/$($mapping.Key)/") {
                    $bd.ResourceId = $bd.ResourceId -replace "/resourceGroups/$([regex]::Escape($mapping.Key))/", "/resourceGroups/$($mapping.Value)/"
                    Write-DebugLog "[RG-MAP] ResourceId remapped for $($entry.ObjectName)" -Level "DEBUG"
                }
            }
        }
        # Map HostPoolArmPath for AppGroups
        if ($bd.PSObject.Properties.Match("HostPoolArmPath").Count -and $bd.HostPoolArmPath) {
            foreach ($mapping in $Global:RgMappings.GetEnumerator()) {
                if ($mapping.Key -ne $mapping.Value -and $bd.HostPoolArmPath -match "/resourceGroups/$($mapping.Key)/") {
                    $bd.HostPoolArmPath = $bd.HostPoolArmPath -replace "/resourceGroups/$([regex]::Escape($mapping.Key))/", "/resourceGroups/$($mapping.Value)/"
                    Write-DebugLog "[RG-MAP] HostPoolArmPath remapped for $($entry.ObjectName)" -Level "DEBUG"
                }
            }
        }
        # Map ApplicationGroupReferences for Workspaces
        if ($bd.PSObject.Properties.Match("ApplicationGroupReferences").Count -and $bd.ApplicationGroupReferences) {
            $newRefs = @($bd.ApplicationGroupReferences | ForEach-Object {
                $ref = $_
                foreach ($mapping in $Global:RgMappings.GetEnumerator()) {
                    if ($mapping.Key -ne $mapping.Value -and $ref -match "/resourceGroups/$($mapping.Key)/") {
                        $ref = $ref -replace "/resourceGroups/$([regex]::Escape($mapping.Key))/", "/resourceGroups/$($mapping.Value)/"
                    }
                }
                $ref
            })
            $bd.ApplicationGroupReferences = $newRefs
        }
        # Map HostPoolReferences for ScalingPlans
        if ($bd.PSObject.Properties.Match("HostPoolReferences").Count -and $bd.HostPoolReferences) {
            $newRefs = @($bd.HostPoolReferences | ForEach-Object {
                $ref = $_
                if ($ref.PSObject.Properties.Match("HostPoolArmPath").Count -and $ref.HostPoolArmPath) {
                    foreach ($mapping in $Global:RgMappings.GetEnumerator()) {
                        if ($mapping.Key -ne $mapping.Value -and $ref.HostPoolArmPath -match "/resourceGroups/$($mapping.Key)/") {
                            $ref.HostPoolArmPath = $ref.HostPoolArmPath -replace "/resourceGroups/$([regex]::Escape($mapping.Key))/", "/resourceGroups/$($mapping.Value)/"
                        }
                    }
                }
                $ref
            })
            $bd.HostPoolReferences = $newRefs
        }
    }
    if ($changed -gt 0) { Write-DebugLog "[RG-MAP] Applied RG mapping to $changed entries" -Level "INFO" }
}

<#
.SYNOPSIS
    Creates any missing Azure resource groups required by the restore manifest.
.DESCRIPTION
    Collects unique target resource group names from selected manifest entries, checks each
    for existence, and creates missing ones using the location from the first matching manifest
    entry or the supplied default. Designed to run inside a background runspace.
.PARAMETER Manifest
    The restore manifest ArrayList containing BackupData with ResourceGroupName fields.
.PARAMETER DefaultLocation
    Fallback Azure region if no location can be inferred from manifest entries.
.OUTPUTS
    Array of resource group names that were created.
#>
function Ensure-ResourceGroups {
    param(
        [System.Collections.ArrayList]$Manifest,
        [string]$DefaultLocation
    )

    # Collect unique target RGs from manifest
    $targetRgs = [System.Collections.ArrayList]::new()
    foreach ($entry in ($Manifest | Where-Object Selected)) {
        $bd = $entry.BackupData
        if (-not $bd) { continue }
        $rgName = $null
        if ($bd.PSObject.Properties.Match("ResourceGroupName").Count) { $rgName = $bd.ResourceGroupName }
        if ($rgName -and $targetRgs -notcontains $rgName) { [void]$targetRgs.Add($rgName) }
    }

    if ($targetRgs.Count -eq 0) { return @() }

    $created = [System.Collections.ArrayList]::new()
    foreach ($rg in $targetRgs) {
        try {
            $existing = Get-AzResourceGroup -Name $rg -ErrorAction SilentlyContinue
            if ($existing) {
                Write-DebugLog "[RG-ENSURE] RG '$rg' exists ($($existing.Location))" -Level "DEBUG"
                continue
            }
            # Determine location: use first manifest entry with this RG, or default
            $loc = $DefaultLocation
            $sample = $Manifest | Where-Object { $_.BackupData -and $_.BackupData.PSObject.Properties.Match("ResourceGroupName").Count -and $_.BackupData.ResourceGroupName -eq $rg -and $_.BackupData.PSObject.Properties.Match("Location").Count } | Select-Object -First 1
            if ($sample -and $sample.BackupData.Location) { $loc = $sample.BackupData.Location }

            Write-DebugLog "[RG-ENSURE] Creating RG '$rg' in $loc" -Level "INFO"
            New-AzResourceGroup -Name $rg -Location $loc -ErrorAction Stop | Out-Null
            Write-DebugLog "[RG-ENSURE] Created RG '$rg' in $loc" -Level "SUCCESS"
            [void]$created.Add($rg)
        } catch {
            Write-DebugLog "[RG-ENSURE] Failed to create RG '$rg': $($_.Exception.Message)" -Level "ERROR"
        }
    }
    return $created
}

<#
.SYNOPSIS
    Discovers VMs matching the host pool naming pattern that are not present in the backup.
.DESCRIPTION
    Extracts a common naming prefix from backup session host names (stripping trailing digits
    and domain suffixes), lists all VMs in the resource group matching that prefix, and returns
    those not already in the backup. Useful for finding VMs added after the backup was taken.
.PARAMETER HostPoolName
    Name of the host pool (for context logging).
.PARAMETER ResourceGroupName
    Resource group to scan for VMs.
.PARAMETER BackupSessionHosts
    Array of session host objects from the backup topology.
.OUTPUTS
    Array of PSCustomObject with Name, ResourceId, Status='Discovered', IsDiscovered=true.
#>
function Find-SessionHostsByPattern {
    param(
        [string]$HostPoolName,
        [string]$ResourceGroupName,
        [array]$BackupSessionHosts
    )

    if (-not $BackupSessionHosts -or $BackupSessionHosts.Count -eq 0) {
        return @()
    }

    # Extract naming pattern from backup session hosts
    # Common patterns: "prefix-001", "prefix001", "prefix-1"
    $names = @($BackupSessionHosts | ForEach-Object {
        $n = $_.Name
        # Strip domain suffix if present (e.g. vm-001.contoso.com -> vm-001)
        if ($n -match "^([^.]+)\.") { $n = $Matches[1] }
        $n
    })

    if ($names.Count -eq 0) { return @() }

    # Find common prefix by stripping trailing digits/separators
    $prefix = $names[0] -replace "[\-_]?\d+$", ""
    if (-not $prefix) { $prefix = $names[0] }

    Write-DebugLog "[VM-DISCOVER] Pattern prefix: '$prefix' from $($names.Count) backup host(s) in RG '$ResourceGroupName'" -Level "DEBUG"

    # Get all VMs in the RG
    try {
        $allVms = @(Get-AzVM -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue)
    } catch {
        Write-DebugLog "[VM-DISCOVER] Failed to list VMs in $ResourceGroupName : $($_.Exception.Message)" -Level "WARN"
        return @()
    }

    # Filter to VMs matching the prefix pattern
    $matching = @($allVms | Where-Object { $_.Name -like "$prefix*" })

    # Find VMs NOT in the backup
    $backupVmIds = @($BackupSessionHosts | Where-Object { $_.ResourceId } | ForEach-Object { $_.ResourceId.ToLower() })
    $discovered = @($matching | Where-Object { $_.Id.ToLower() -notin $backupVmIds })

    Write-DebugLog "[VM-DISCOVER] Found $($matching.Count) VMs matching '$prefix*', $($discovered.Count) not in backup" -Level "INFO"

    # Return as session-host-like objects
    $result = @($discovered | ForEach-Object {
        [PSCustomObject]@{
            Name       = $_.Name
            ResourceId = $_.Id
            Status     = "Discovered"
            IsDiscovered = $true
        }
    })

    return $result
}

# ===============================================================================
# SECTION 19: UI UPDATE HELPERS
# Dashboard stat card updates, interactive topology tree builder with rich node headers
# (icons, type badges, location, status dots), backup list renderer, restore manifest
# tree with tri-state checkboxes, and drift detection banner management.
# ===============================================================================

<#
.SYNOPSIS
    Updates the five dashboard stat cards (Host Pools, App Groups, Workspaces, Session Hosts, Scaling Plans).
.PARAMETER Topology
    The current AVD topology object, or $null to reset cards to placeholder dashes.
#>
function Update-DashboardCards {
    param($Topology)
    $Window.Dispatcher.Invoke([Action]{
        if (-not $Topology) {
            $lblDashHostPools.Text    = '--'
            $lblDashAppGroups.Text    = '--'
            $lblDashWorkspaces.Text   = '--'
            $lblDashSessionHosts.Text = '--'
            $lblDashScalingPlans.Text = '--'
            return
        }
        $hpCount = @($Topology.HostPools).Count
        $agCount = @($Topology.HostPools | ForEach-Object { $_.AppGroups } | Where-Object { $_ }).Count
        $wsCount = @($Topology.Workspaces).Count
        $shCount = @($Topology.HostPools | ForEach-Object { $_.SessionHosts } | Where-Object { $_ }).Count
        $spCount = @($Topology.HostPools | ForEach-Object { $_.ScalingPlans } | Where-Object { $_ }).Count

        $lblDashHostPools.Text    = $hpCount
        $lblDashAppGroups.Text    = $agCount
        $lblDashWorkspaces.Text   = $wsCount
        $lblDashSessionHosts.Text = $shCount
        $lblDashScalingPlans.Text = $spCount
    })
}

<#
.SYNOPSIS
    Creates a small styled label badge for resource type indicators in the topology tree.
.PARAMETER Text
    Badge display text (e.g. 'Pooled', 'RemoteApp').
.PARAMETER BgResource
    Theme resource key for the badge background brush.
.PARAMETER FgResource
    Theme resource key for the badge foreground text brush.
.OUTPUTS
    System.Windows.Controls.Border containing the styled TextBlock.
#>
function New-TypeBadge {
    param([string]$Text, [string]$BgResource = 'ThemeInputBg', [string]$FgResource = 'ThemeTextMuted')
    $Badge = [System.Windows.Controls.Border]::new()
    $Badge.CornerRadius = [System.Windows.CornerRadius]::new(3)
    $Badge.Padding = [System.Windows.Thickness]::new(5,1,5,1)
    $Badge.Margin = [System.Windows.Thickness]::new(6,0,0,0)
    $Badge.VerticalAlignment = 'Center'
    $Badge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, $BgResource)
    $Lbl = [System.Windows.Controls.TextBlock]::new()
    $Lbl.Text = $Text; $Lbl.FontSize = 9; $Lbl.FontWeight = [System.Windows.FontWeights]::SemiBold
    $Lbl.VerticalAlignment = 'Center'
    $Lbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $FgResource)
    $Badge.Child = $Lbl; return $Badge
}

<#
.SYNOPSIS
    Creates a small colored ellipse used as a status indicator dot in tree nodes.
.PARAMETER Color
    Hex color string for the dot fill brush.
.OUTPUTS
    System.Windows.Shapes.Ellipse (6x6 pixels).
#>
function New-StatusDot {
    param([string]$Color = '#16A34A')
    $Dot = [System.Windows.Shapes.Ellipse]::new()
    $Dot.Width = 6; $Dot.Height = 6; $Dot.Margin = [System.Windows.Thickness]::new(6,0,0,0)
    $Dot.VerticalAlignment = 'Center'
    $Dot.Fill = $Global:CachedBC.ConvertFromString($Color)
    return $Dot
}

<#
.SYNOPSIS
    Builds a rich horizontal StackPanel header for a TreeViewItem with icon, text, badges, location, and status.
.PARAMETER Icon
    Segoe Fluent Icons character for the leading icon.
.PARAMETER IconColor
    Theme resource key for the icon foreground.
.PARAMETER Text
    Primary display text (resource name).
.PARAMETER Bold
    When set, renders the text in SemiBold weight.
.PARAMETER Badges
    Array of badge strings rendered as type indicators after the name.
.PARAMETER Location
    Azure region text displayed in muted style.
.PARAMETER StatusColor
    Hex color for the status dot indicator.
.PARAMETER StatusText
    Optional status label text next to the dot.
.PARAMETER Count
    If >= 0, renders a count badge after the name.
.OUTPUTS
    System.Windows.Controls.StackPanel containing the composed header elements.
#>
function New-TreeHeader {
    param([string]$Icon, [string]$IconColor, [string]$Text, [switch]$Bold, [string[]]$Badges, [string]$Location, [string]$StatusColor, [string]$StatusText, [int]$Count = -1)
    $sp = [System.Windows.Controls.StackPanel]::new(); $sp.Orientation = 'Horizontal'
    # Icon
    $ic = [System.Windows.Controls.TextBlock]::new(); $ic.Text = $Icon
    $ic.FontFamily = [System.Windows.Media.FontFamily]::new('Segoe Fluent Icons, Segoe MDL2 Assets')
    $ic.FontSize = 13; $ic.Margin = [System.Windows.Thickness]::new(0,0,6,0); $ic.VerticalAlignment = 'Center'
    $ic.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $IconColor)
    [void]$sp.Children.Add($ic)
    # Name
    $nm = [System.Windows.Controls.TextBlock]::new(); $nm.Text = $Text; $nm.VerticalAlignment = 'Center'
    if ($Bold) { $nm.FontWeight = [System.Windows.FontWeights]::SemiBold }
    $nm.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
    [void]$sp.Children.Add($nm)
    # Count badge
    if ($Count -ge 0) {
        $cb = New-TypeBadge -Text $Count -BgResource 'ThemeAccentDim' -FgResource 'ThemeAccentText'
        [void]$sp.Children.Add($cb)
    }
    # Type badges
    foreach ($b in $Badges) { [void]$sp.Children.Add((New-TypeBadge -Text $b)) }
    # Location
    if ($Location) {
        $loc = [System.Windows.Controls.TextBlock]::new(); $loc.Text = $Location; $loc.FontSize = 10
        $loc.Margin = [System.Windows.Thickness]::new(8,0,0,0); $loc.VerticalAlignment = 'Center'
        $loc.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
        [void]$sp.Children.Add($loc)
    }
    # Status dot + text
    if ($StatusColor) {
        [void]$sp.Children.Add((New-StatusDot -Color $StatusColor))
        if ($StatusText) {
            $st = [System.Windows.Controls.TextBlock]::new(); $st.Text = $StatusText; $st.FontSize = 10
            $st.Margin = [System.Windows.Thickness]::new(4,0,0,0); $st.VerticalAlignment = 'Center'
            $st.Foreground = $Global:CachedBC.ConvertFromString($StatusColor)
            [void]$sp.Children.Add($st)
        }
    }
    return $sp
}

<#
.SYNOPSIS
    Rebuilds the interactive topology TreeView from the current AVD topology object.
.DESCRIPTION
    Clears and repopulates the Topology tab TreeView with a hierarchical structure:
    Host Pools -> App Groups -> Applications/RBAC, Session Hosts (with status dots),
    Scaling Plans, and standalone Workspaces with app group references. Each node has
    rich headers with icons, type badges, location hints, and status indicators.
.PARAMETER Topology
    The AVD topology object from Get-AvdTopology.
#>
function Update-TopologyTree {
    param($Topology)
    $Window.Dispatcher.Invoke([Action]{
        $tvTopology.Items.Clear()
        $lblTopologyEmpty.Visibility = if ($Topology) { 'Collapsed' } else { 'Visible' }
        if (-not $Topology) { return }

        foreach ($hp in $Topology.HostPools) {
            $hpNode = [System.Windows.Controls.TreeViewItem]::new()
            $hpNode.Header = New-TreeHeader -Icon ([string]([char]0xE968)) -IconColor 'ThemeAccent' `
                -Text $hp.Name -Bold -Badges @($hp.HostPoolType, $hp.LoadBalancerType) `
                -Location $hp.Location -Count (@($hp.AppGroups).Count + @($hp.SessionHosts).Count)
            $hpNode.Tag = $hp
            $hpNode.IsExpanded = $true
            $hpNode.Foreground = $Window.FindResource('ThemeTextBody')

            # App Groups
            foreach ($ag in @($hp.AppGroups)) {
                $agNode = [System.Windows.Controls.TreeViewItem]::new()
                $agNode.Header = New-TreeHeader -Icon ([string]([char]0xE71D)) -IconColor 'ThemeGreenAccent' `
                    -Text $ag.Name -Badges @($ag.ApplicationGroupType) -Location $ag.Location
                $agNode.Tag = $ag
                $agNode.Foreground = $Window.FindResource('ThemeTextBody')

                foreach ($app in @($ag.Applications)) {
                    $appNode = [System.Windows.Controls.TreeViewItem]::new()
                    $appLabel = if ($app.FriendlyName) { $app.FriendlyName } else { $app.Name }
                    $appNode.Header = New-TreeHeader -Icon ([string]([char]0xE74C)) -IconColor 'ThemeTextMuted' -Text $appLabel -Badges @('remoteapp')
                    $appNode.Tag = $app
                    $appNode.Foreground = $Window.FindResource('ThemeTextMuted')
                    [void]$agNode.Items.Add($appNode)
                }

                # RBAC
                foreach ($rbac in @($ag.RbacAssignments)) {
                    $rbNode = [System.Windows.Controls.TreeViewItem]::new()
                    $rbNode.Header = New-TreeHeader -Icon ([string]([char]0xE8D7)) -IconColor 'ThemeTextDim' -Text $rbac.DisplayName -Badges @($rbac.RoleDefinitionName)
                    $rbNode.Tag = $rbac
                    $rbNode.Foreground = $Window.FindResource('ThemeTextMuted')
                    [void]$agNode.Items.Add($rbNode)
                }

                [void]$hpNode.Items.Add($agNode)
            }

            # Session Hosts
            foreach ($sh in @($hp.SessionHosts)) {
                $shNode = [System.Windows.Controls.TreeViewItem]::new()
                $shStatusMap = switch ($sh.Status) {
                    'Available'    { @{ Color='#16A34A'; Text='Available' } }
                    'Unavailable'  { @{ Color='#DC2626'; Text='Unavailable' } }
                    'Shutdown'     { @{ Color='#71717A'; Text='Shutdown' } }
                    'Disconnected' { @{ Color='#F59E0B'; Text='Disconnected' } }
                    default        { @{ Color='#71717A'; Text=$sh.Status } }
                }
                $shNode.Header = New-TreeHeader -Icon ([string]([char]0xE7F4)) -IconColor 'ThemeTextMuted' `
                    -Text $sh.Name -Badges @('session-host') `
                    -StatusColor $shStatusMap.Color -StatusText $shStatusMap.Text
                $shNode.Tag = $sh
                $shNode.Foreground = $Window.FindResource('ThemeTextMuted')
                [void]$hpNode.Items.Add($shNode)
            }

            # Scaling Plans
            foreach ($sp in @($hp.ScalingPlans)) {
                $spNode = [System.Windows.Controls.TreeViewItem]::new()
                $spNode.Header = New-TreeHeader -Icon ([string]([char]0xE9D5)) -IconColor 'ThemeTextMuted' -Text $sp.Name -Badges @('scaling-plan') -Location $sp.Location
                $spNode.Tag = $sp
                $spNode.Foreground = $Window.FindResource('ThemeTextMuted')
                [void]$hpNode.Items.Add($spNode)
            }

            # Diagnostic Settings
            foreach ($ds in @($hp.DiagnosticSettings)) {
                $dsNode = [System.Windows.Controls.TreeViewItem]::new()
                $dsBadges = @('diagnostic')
                if ($ds.WorkspaceId)    { $dsBadges += 'Log Analytics' }
                if ($ds.StorageAccountId) { $dsBadges += 'Storage' }
                if ($ds.EventHubName)   { $dsBadges += 'Event Hub' }
                $dsNode.Header = New-TreeHeader -Icon ([string]([char]0xEA80)) -IconColor 'ThemeTextMuted' -Text $ds.Name -Badges $dsBadges
                $dsNode.Tag = $ds
                $dsNode.Foreground = $Window.FindResource('ThemeTextMuted')

                # Show log categories under diag setting
                foreach ($log in @($ds.Logs)) {
                    $logNode = [System.Windows.Controls.TreeViewItem]::new()
                    $enabledStatus = if ($log.Enabled) { @{ Color='#16A34A'; Text='Enabled' } } else { @{ Color='#71717A'; Text='Disabled' } }
                    $logNode.Header = New-TreeHeader -Icon ([string]([char]0xE7C3)) -IconColor 'ThemeTextDim' -Text $log.Category `
                        -StatusColor $enabledStatus.Color -StatusText $enabledStatus.Text
                    $logNode.Foreground = $Window.FindResource('ThemeTextMuted')
                    [void]$dsNode.Items.Add($logNode)
                }
                [void]$hpNode.Items.Add($dsNode)
            }

            [void]$tvTopology.Items.Add($hpNode)
        }

        # Workspaces as top-level nodes
        foreach ($ws in $Topology.Workspaces) {
            $wsNode = [System.Windows.Controls.TreeViewItem]::new()
            $wsNode.Header = New-TreeHeader -Icon ([string]([char]0xE8A1)) -IconColor 'ThemeAccentLight' `
                -Text $ws.Name -Bold -Badges @('workspace') `
                -Location $ws.Location -Count (@($ws.ApplicationGroupReferences).Count)
            $wsNode.Tag = $ws
            $wsNode.IsExpanded = $false
            $wsNode.Foreground = $Window.FindResource('ThemeTextBody')

            foreach ($agRef in @($ws.ApplicationGroupReferences)) {
                $refNode = [System.Windows.Controls.TreeViewItem]::new()
                $refNode.Header = New-TreeHeader -Icon ([string]([char]0xE71B)) -IconColor 'ThemeTextDim' -Text ($agRef -split '/')[-1] -Badges @('app-group-ref')
                $refNode.Foreground = $Window.FindResource('ThemeTextMuted')
                [void]$wsNode.Items.Add($refNode)
            }
            [void]$tvTopology.Items.Add($wsNode)
        }
    })
}

<#
.SYNOPSIS
    Populates the Backups tab ListBox with all backup files and their integrity status.
.DESCRIPTION
    Calls Get-BackupList, renders each backup as a styled panel with verification badge,
    label/filename, subscription, timestamp, resource counts, and file size. Updates the
    sidebar backup count and empty-state visibility.
#>
function Update-BackupList {
    $Window.Dispatcher.Invoke([Action]{
        $lstBackups.Items.Clear()
        $Backups = Get-BackupList
        $Global:BackupCache = $Backups

        foreach ($b in $Backups) {
            $Verified = if ($b.Verified) { [char]0x2705 } else { [char]0x26A0 }
            $Header = "$Verified $($b.SubscriptionName) | $($b.Timestamp) | HP:$($b.HostPools) AG:$($b.AppGroups) SH:$($b.SessionHosts) ($($b.SizeKB) KB)"
            if ($b.Label) { $Header = "$Verified [$($b.Label)] $Header" }

            $Item = [System.Windows.Controls.ListBoxItem]::new()
            $Item.Content = $Header
            $Item.Tag = $b
            $Item.Foreground = $Window.FindResource('ThemeTextBody')
            $Item.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
            $Item.FontSize = 12
            [void]$lstBackups.Items.Add($Item)
        }

        $lblSidebarBackupCount.Text = "$($Backups.Count) backup(s)"
        $lblBackupsEmpty.Visibility = if ($Backups.Count -gt 0) { 'Collapsed' } else { 'Visible' }
    })
}

<#
.SYNOPSIS
    Builds the restore manifest TreeView with grouped items, action badges, and tri-state checkboxes.
.DESCRIPTION
    Groups manifest entries by ObjectType, creates parent nodes with type icons and child counts,
    renders per-item badges (Create=green, Update=blue) with property-level diff details for
    Update actions, and wires tri-state parent checkboxes that toggle all children.
#>
function Update-RestoreTree {
    param($DiffResults, $BackupTopology)
    $Window.Dispatcher.Invoke([Action]{
        $tvRestoreTree.Items.Clear()
        if (-not $DiffResults) { return }

        $Manifest = @(Build-RestoreManifest -DiffResults $DiffResults -BackupTopology $BackupTopology)
        $Global:RestoreManifest = $Manifest

        # Group by type
        $Groups = $Manifest | Group-Object ObjectType
        foreach ($g in $Groups) {
            # Icon/color map for restore tree
            $restoreIconMap = @{
                'HostPool'         = @{ Icon = [string]([char]0xE968); Color = 'ThemeAccent' }
                'AppGroup'         = @{ Icon = [string]([char]0xE71D); Color = 'ThemeTextMuted' }
                'Application'      = @{ Icon = [string]([char]0xE74C); Color = 'ThemeTextMuted' }
                'Workspace'        = @{ Icon = [string]([char]0xE8A1); Color = 'ThemeAccent' }
                'ScalingPlan'      = @{ Icon = [string]([char]0xE9D5); Color = 'ThemeTextMuted' }
                'SessionHostCount' = @{ Icon = [string]([char]0xE7F4); Color = 'ThemeTextMuted' }
            }
            $parentCb = [System.Windows.Controls.CheckBox]::new()
            $parentCb.IsChecked = $true
            $parentCb.VerticalContentAlignment = 'Center'
            # Build styled content for parent checkbox
            $parentSp = [System.Windows.Controls.StackPanel]::new(); $parentSp.Orientation = 'Horizontal'
            $typeInfo = $restoreIconMap[$g.Name]
            if ($typeInfo) {
                $ico = [System.Windows.Controls.TextBlock]::new(); $ico.Text = $typeInfo.Icon
                $ico.FontFamily = [System.Windows.Media.FontFamily]::new('Segoe Fluent Icons, Segoe MDL2 Assets')
                $ico.FontSize = 13; $ico.Margin = [System.Windows.Thickness]::new(0,0,6,0); $ico.VerticalAlignment = 'Center'
                $ico.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $typeInfo.Color)
                [void]$parentSp.Children.Add($ico)
            }
            $typeLbl = [System.Windows.Controls.TextBlock]::new(); $typeLbl.Text = $g.Name
            $typeLbl.FontWeight = [System.Windows.FontWeights]::SemiBold; $typeLbl.VerticalAlignment = 'Center'
            $typeLbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
            [void]$parentSp.Children.Add($typeLbl)
            [void]$parentSp.Children.Add((New-TypeBadge -Text $g.Count -BgResource 'ThemeAccentDim' -FgResource 'ThemeAccentText'))
            $parentCb.Content = $parentSp

            $parentNode = [System.Windows.Controls.TreeViewItem]::new()
            $parentNode.Header = $parentCb
            $parentNode.IsExpanded = $true

            foreach ($item in $g.Group) {
                $cb = [System.Windows.Controls.CheckBox]::new()
                $cb.IsChecked = $true
                $cb.Tag = $item
                $cb.VerticalContentAlignment = 'Center'
                # Build styled content for child checkbox
                $childSp = [System.Windows.Controls.StackPanel]::new(); $childSp.Orientation = 'Horizontal'
                $childLbl = [System.Windows.Controls.TextBlock]::new(); $childLbl.Text = $item.ObjectName
                $childLbl.VerticalAlignment = 'Center'
                $childLbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
                [void]$childSp.Children.Add($childLbl)
                # Action badge (color-coded)
                $actionColors = switch ($item.Action) {
                    'Create' { @{ Bg='ThemeSuccessBg'; Fg='ThemeSuccessText' } }
                    'Update' { @{ Bg='ThemeWarningBg'; Fg='ThemeWarningText' } }
                    default  { @{ Bg='ThemeInputBg';   Fg='ThemeTextMuted'   } }
                }
                [void]$childSp.Children.Add((New-TypeBadge -Text $item.Action -BgResource $actionColors.Bg -FgResource $actionColors.Fg))
                # Type badge
                $typeLabel = ($item.ObjectType -creplace '([A-Z])', ' $1').Trim().ToLower()
                [void]$childSp.Children.Add((New-TypeBadge -Text $typeLabel))
                $cb.Content = $childSp
                $cb.Add_Checked({ $this.Tag.Selected = $true })
                $cb.Add_Unchecked({ $this.Tag.Selected = $false })

                $childNode = [System.Windows.Controls.TreeViewItem]::new()
                $childNode.Header = $cb

                # Show property-level changes for Update items
                if ($item.Action -eq 'Update' -and $item.ChangedProperties -and $item.ChangedProperties.Count -gt 0) {
                    foreach ($cp in $item.ChangedProperties) {
                        $propSp = [System.Windows.Controls.StackPanel]::new(); $propSp.Orientation = 'Horizontal'
                        $propIcon = [System.Windows.Controls.TextBlock]::new(); $propIcon.Text = [string]([char]0xE8AB)
                        $propIcon.FontFamily = [System.Windows.Media.FontFamily]::new('Segoe Fluent Icons, Segoe MDL2 Assets')
                        $propIcon.FontSize = 10; $propIcon.Margin = [System.Windows.Thickness]::new(0,0,5,0); $propIcon.VerticalAlignment = 'Center'
                        $propIcon.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
                        [void]$propSp.Children.Add($propIcon)
                        $propName = [System.Windows.Controls.TextBlock]::new(); $propName.Text = $cp.Property
                        $propName.FontSize = 11; $propName.FontWeight = [System.Windows.FontWeights]::SemiBold
                        $propName.VerticalAlignment = 'Center'; $propName.Margin = [System.Windows.Thickness]::new(0,0,6,0)
                        $propName.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
                        [void]$propSp.Children.Add($propName)
                        $propVal = [System.Windows.Controls.TextBlock]::new(); $propVal.Text = $cp.Summary
                        $propVal.FontSize = 11; $propVal.VerticalAlignment = 'Center'
                        $propVal.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
                        [void]$propSp.Children.Add($propVal)
                        $propNode = [System.Windows.Controls.TreeViewItem]::new()
                        $propNode.Header = $propSp
                        if ($cp.PSObject.Properties.Match('Before').Count -and $cp.PSObject.Properties.Match('After').Count) {
                            $propNode.ToolTip = "Before: $($cp.Before)`nAfter:  $($cp.After)"
                        }
                        $propNode.Padding = [System.Windows.Thickness]::new(0)
                        $propNode.Margin  = [System.Windows.Thickness]::new(0)
                        [void]$childNode.Items.Add($propNode)
                    }
                }

                [void]$parentNode.Items.Add($childNode)
            }

            # Parent checkbox tri-state behaviour
            $parentCb.Add_Checked({
                param($s,$e)
                foreach ($child in $parentNode.Items) { ($child.Header -as [System.Windows.Controls.CheckBox]).IsChecked = $true }
            }.GetNewClosure())
            $parentCb.Add_Unchecked({
                param($s,$e)
                foreach ($child in $parentNode.Items) { ($child.Header -as [System.Windows.Controls.CheckBox]).IsChecked = $false }
            }.GetNewClosure())

            [void]$tvRestoreTree.Items.Add($parentNode)
        }

        # Pre-flight summary
        $AddedCount   = @($DiffResults.Added).Count
        $RemovedCount = @($DiffResults.Removed).Count
        $ModifiedCount= @($DiffResults.Modified).Count
        $lblRestoreEmpty.Text = "Pre-flight: $AddedCount added, $RemovedCount removed, $ModifiedCount modified - $($Manifest.Count) restore action(s)"
        Update-RestoreActionButtons
    })
}

<#+
.SYNOPSIS
    Enables/disables restore action buttons based on topology, manifest, and discovery state.
#>
function Update-RestoreActionButtons {
    $Window.Dispatcher.Invoke([Action]{
        $hasManifest = ($Global:RestoreManifest -and @($Global:RestoreManifest).Count -gt 0)
        $hasTopology = ($null -ne $Global:CurrentTopology)
        $canRun = ($hasManifest -and $hasTopology -and -not $Global:DiscoveryInProgress)
        $btnExecuteRestore.IsEnabled = $canRun
        $btnDryRun.IsEnabled = $canRun
        $btnExportPlan.IsEnabled = $canRun
    })
}

<#
.SYNOPSIS
    Shows or hides the drift-detected warning banner on the Dashboard tab.
.DESCRIPTION
    Compares diff result counts to determine if Added, Removed, or Modified objects exist.
    Shows the drift banner with a summary message if changes are detected, otherwise hides it.
#>
function Update-DriftBanner {
    param($DiffResults)
    $Window.Dispatcher.Invoke([Action]{
        if (-not $DiffResults -or (
            $DiffResults.Added.Count -eq 0 -and
            $DiffResults.Removed.Count -eq 0 -and
            $DiffResults.Modified.Count -eq 0)) {
            $pnlDriftBanner.Visibility = 'Collapsed'
            return
        }
        $Total = $DiffResults.Added.Count + $DiffResults.Removed.Count + $DiffResults.Modified.Count
        $lblDriftMessage.Text = "$Total drift(s) detected since last backup"
        $pnlDriftBanner.Visibility = 'Visible'
    })
}

<#
.SYNOPSIS
    Updates the status bar text at the bottom of the window.
#>
function Update-StatusBar {
    param([string]$Text)
    $Window.Dispatcher.Invoke([Action]{
        $lblStatusBar.Text = $Text
    })
}

# ===============================================================================
# SECTION 20: EXPORT ENGINE
# ===============================================================================

<#
.SYNOPSIS
    Exports the current topology to a JSON or CSV file via SaveFileDialog.
.DESCRIPTION
    Presents a save dialog with JSON/CSV format options. For JSON, serializes the full topology
    at depth 20. For CSV, flattens host pools and session hosts into tabular rows. Unlocks the
    Exporter achievement on success.
#>
function Export-TopologyReport {
    param(
        [string]$Format = 'JSON',
        $Topology
    )

    if (-not $Topology) {
        Show-Toast -Message "No topology loaded. Run a discovery first." -Type 'Warning'
        return
    }

    $Dialog = [Microsoft.Win32.SaveFileDialog]::new()
    $Dialog.Title = "Export Topology"
    $Dialog.InitialDirectory = $Global:BackupDir
    switch ($Format) {
        'JSON' {
            $Dialog.Filter = "JSON Files (*.json)|*.json"
            $Dialog.DefaultExt = ".json"
        }
        'CSV' {
            $Dialog.Filter = "CSV Files (*.csv)|*.csv"
            $Dialog.DefaultExt = ".csv"
        }
    }
    $Dialog.FileName = "AvdRewind_Export_$(Get-Date -Format 'yyyyMMdd_HHmmss')"

    if ($Dialog.ShowDialog() -ne $true) { return }

    try {
        switch ($Format) {
            'JSON' {
                $Topology | ConvertTo-Json -Depth 20 | Out-File -FilePath $Dialog.FileName -Encoding UTF8
            }
            'CSV' {
                # Flatten host pools + session hosts for CSV
                $Rows = foreach ($hp in $Topology.HostPools) {
                    foreach ($sh in @($hp.SessionHosts)) {
                        [PSCustomObject]@{
                            HostPool      = $hp.Name
                            Location      = $hp.Location
                            PoolType      = $hp.HostPoolType
                            SessionHost   = $sh.Name
                            Status        = $sh.Status
                            AllowNew      = $sh.AllowNewSession
                            AssignedUser  = $sh.AssignedUser
                            VmSize        = $sh.VmMetadata.VmSize
                            PrivateIp     = $sh.NetworkConfig.PrivateIp
                        }
                    }
                    if (@($hp.SessionHosts).Count -eq 0) {
                        [PSCustomObject]@{
                            HostPool      = $hp.Name
                            Location      = $hp.Location
                            PoolType      = $hp.HostPoolType
                            SessionHost   = ''
                            Status        = ''
                            AllowNew      = ''
                            AssignedUser  = ''
                            VmSize        = ''
                            PrivateIp     = ''
                        }
                    }
                }
                $Rows | Export-Csv -Path $Dialog.FileName -NoTypeInformation -Encoding UTF8
            }
        }
        Write-DebugLog "[EXPORT] Saved $Format report: $($Dialog.FileName)" -Level 'SUCCESS'
        Show-Toast -Message "Exported to $($Dialog.FileName)" -Type 'Success'
        Check-Achievement 'Exporter'
    } catch {
        Write-DebugLog "[EXPORT] Failed: $($_.Exception.Message)" -Level 'ERROR'
        Show-Toast -Message "Export failed: $($_.Exception.Message)" -Type 'Error'
    }
}

# ===============================================================================
# SECTION 20b: AUTH UI HELPERS
# ===============================================================================

<#
.SYNOPSIS
    Authenticates to Azure and populates the subscription dropdown.
.DESCRIPTION
    Validates any cached Azure context first; if expired or missing, initiates a fresh
    Connect-AzAccount login. On success, enumerates all accessible subscriptions, populates
    the subscription combobox, restores the previously selected subscription if available,
    and updates the auth status indicator (green dot + connected label).
#>
function Connect-ToAzure {
    Write-DebugLog "[AUTH] Connect-ToAzure: >>> ENTER" -Level 'DEBUG'
    Update-StatusBar "Authenticating..."
    $authDot.Fill = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(255,165,0))  # orange = connecting

    Start-BackgroundWork -Name 'AzLogin' -ScriptBlock {
        try {
            $ctx = Get-AzContext -ErrorAction SilentlyContinue
            if ($ctx -and $ctx.Account) {
                # Validate that cached token is still valid
                try {
                    $null = Get-AzTenant -ErrorAction Stop | Select-Object -First 1
                    return @{ Status='Cached'; Context=$ctx }
                } catch {
                    # Token expired, need re-auth
                }
            }
            Connect-AzAccount -ErrorAction Stop | Out-Null
            $ctx = Get-AzContext
            return @{ Status='Fresh'; Context=$ctx }
        } catch {
            return @{ Status='Failed'; Error=$_.Exception.Message }
        }
    } -OnComplete {
        param($Result)
        if ($Result.Status -eq 'Failed') {
            Write-DebugLog "[AUTH] Login failed: $($Result.Error)" -Level 'ERROR'
            $authDot.Fill = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(255,59,48))  # red
            $lblAuthStatus.Text = "Not connected"
            $btnLogin.Visibility = 'Visible'
            $btnLogout.Visibility = 'Collapsed'
            $cmbSubscription.Visibility = 'Collapsed'
            Show-Toast -Message "Login failed: $($Result.Error)" -Type 'Error'
            Update-StatusBar "Authentication failed"
            return
        }

        $ctx = $Result.Context
        Write-DebugLog "[AUTH] Logged in as $($ctx.Account.Id) ($($Result.Status))" -Level 'SUCCESS'
        $authDot.Fill = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(52,199,89))  # green
        $lblAuthStatus.Text = $ctx.Account.Id
        $Global:CurrentSubId = $ctx.Subscription.Id
        $btnLogin.Visibility = 'Collapsed'
        $btnLogout.Visibility = 'Visible'
        $cmbSubscription.Visibility = 'Visible'

        # Populate subscription combo
        Start-BackgroundWork -Name 'LoadSubs' -ScriptBlock {
            @(Get-AzSubscription -ErrorAction SilentlyContinue | Where-Object State -eq 'Enabled' | Select-Object Name, Id)
        } -OnComplete {
            param($Subs)
            $cmbSubscription.Items.Clear()
            foreach ($s in @($Subs)) {
                $item = [System.Windows.Controls.ComboBoxItem]::new()
                $item.Content = $s.Name
                $item.Tag = $s.Id
                [void]$cmbSubscription.Items.Add($item)
                if ($s.Id -eq $Global:CurrentSubId) {
                    $cmbSubscription.SelectedItem = $item
                }
            }
            Update-StatusBar "Connected - $(@($Subs).Count) subscription(s)"
        }
    }
}

<#
.SYNOPSIS
    Disconnects from Azure and resets the UI to the unauthenticated state.
.DESCRIPTION
    Calls Disconnect-AzAccount, clears the subscription dropdown, resets the auth status
    indicator, and clears cached topology and subscription state.
#>
function Disconnect-FromAzure {
    Write-DebugLog "[AUTH] Disconnecting..." -Level 'INFO'
    try {
        Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null
    } catch { }
    $authDot.Fill = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(255,59,48))  # red
    $lblAuthStatus.Text = "Not connected"
    $btnLogin.Visibility = 'Visible'
    $btnLogout.Visibility = 'Collapsed'
    $cmbSubscription.Visibility = 'Collapsed'
    $cmbSubscription.Items.Clear()
    $Global:CurrentSubId = $null
    $Global:CurrentTopology = $null
    Update-DashboardCards -Topology $null
    Update-StatusBar "Disconnected"
    Write-DebugLog "[AUTH] Disconnected" -Level 'INFO'
}

<#
.SYNOPSIS
    Triggers a background AVD topology discovery for the active subscription.
.DESCRIPTION
    Launches Get-AvdTopology in a background runspace, shows shimmer progress, and on completion
    updates the dashboard cards, topology tree, sidebar stats, and auto-diffs against the latest
    backup. Unlocks discovery-related achievements (FirstDiscovery, SpeedDemon, BigEnvironment).
    Triggers auto-backup if enabled in settings.
#>
function Start-AvdDiscovery {
    if (-not $Global:CurrentSubId) {
        Show-Toast -Message "Connect to Azure first." -Type 'Warning'
        return
    }
    $Global:DiscoveryInProgress = $true
    Update-RestoreActionButtons
    # If user is currently on the Restore tab, redirect to Dashboard
    if ($Global:ActiveTabName -eq 'Restore') {
        Switch-Tab 'Dashboard'
        Show-Toast -Message 'Switched away from Restore — discovery in progress.' -Type 'Warning'
    }
    Write-DebugLog "[DISCOVER] Discovery triggered" -Level 'INFO'
    Update-StatusBar "Discovering AVD topology..."
    Show-GlobalProgressIndeterminate -Text "Discovering AVD topology..."

    $Global:DiscoveryStartTime = Get-Date
    Start-BackgroundWork -Name 'Discovery' -Functions @('Get-AvdTopology') -ScriptBlock {
        param($SubId, $SubName)
        Get-AvdTopology -SubscriptionId $SubId -SubscriptionName $SubName
    } -Arguments @($Global:CurrentSubId, ($cmbSubscription.SelectedItem.Content)) -OnComplete {
        param($Topo)
        Stop-Shimmer
        $Global:DiscoveryInProgress = $false
        Update-RestoreActionButtons
        if (-not $Topo -or -not $Topo.HostPools) {
            Write-DebugLog "[DISCOVER] Discovery returned no data" -Level 'WARN'
            Update-StatusBar "Discovery failed - no data returned"
            return
        }
        $Global:CurrentTopology = $Topo
        $Global:TopologyCacheTime = Get-Date
        Update-DashboardCards -Topology $Topo
        Update-TopologyTree -Topology $Topo
        Update-StatusBar "Discovery complete - $(@($Topo.HostPools).Count) host pool(s)"
        Show-Toast -Message "Topology discovered successfully" -Type 'Success'
        Check-Achievement 'FirstDiscovery'

        # Populate target region dropdown from discovered topology locations
        $cmbTargetRegion.Items.Clear()
        $regions = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($hp in @($Topo.HostPools)) {
            if ($hp.Location) { [void]$regions.Add($hp.Location) }
            foreach ($ag in @($hp.AppGroups)) { if ($ag.Location) { [void]$regions.Add($ag.Location) } }
            foreach ($sp in @($hp.ScalingPlans)) { if ($sp.Location) { [void]$regions.Add($sp.Location) } }
        }
        foreach ($ws in @($Topo.Workspaces)) { if ($ws.Location) { [void]$regions.Add($ws.Location) } }
        $commonRegions = @('eastus','eastus2','westus','westus2','westus3','centralus','northcentralus','southcentralus','westcentralus','northeurope','westeurope','uksouth','ukwest','francecentral','germanywestcentral','switzerlandnorth','swedencentral','norwayeast','eastasia','southeastasia','japaneast','japanwest','australiaeast','australiasoutheast','koreacentral','canadacentral','canadaeast','brazilsouth','southafricanorth','uaenorth')
        foreach ($r in $commonRegions) { [void]$regions.Add($r) }
        foreach ($r in ($regions | Sort-Object)) {
            $item = [System.Windows.Controls.ComboBoxItem]::new()
            $item.Content = $r
            [void]$cmbTargetRegion.Items.Add($item)
        }
        if ($cmbTargetRegion.Items.Count -gt 0) { $cmbTargetRegion.SelectedIndex = 0 }

        # Populate Resource Group filter
        $cmbTopoResourceGroup.Items.Clear()
        $allItem = [System.Windows.Controls.ComboBoxItem]::new()
        $allItem.Content = 'All Resource Groups'
        [void]$cmbTopoResourceGroup.Items.Add($allItem)
        foreach ($rg in @($Topo.ResourceGroups)) {
            $rgItem = [System.Windows.Controls.ComboBoxItem]::new()
            $rgItem.Content = $rg
            [void]$cmbTopoResourceGroup.Items.Add($rgItem)
        }
        $cmbTopoResourceGroup.SelectedIndex = 0
        if ($Global:DiscoveryStartTime -and ((Get-Date) - $Global:DiscoveryStartTime).TotalSeconds -lt 10) { Check-Achievement 'SpeedDemon' }
        if (@($Topo.HostPools).Count -ge 10) { Check-Achievement 'BigEnvironment' }

        # Auto-diff against latest backup
        $Backups = Get-BackupList
        if ($Backups.Count -gt 0) {
            $LatestBackup = $Backups[0]
            $DiffResult = Compare-AvdTopology -Baseline $LatestBackup.Topology -Current $Topo
            $Global:LastDiffResult = $DiffResult
            Update-DriftBanner -DiffResults $DiffResult
        }
    }
}

<#
.SYNOPSIS
    Switches the active Azure subscription context and triggers a fresh discovery.
.DESCRIPTION
    Sets the Azure context to the selected subscription via Set-AzContext, updates global
    state, saves preferences, records the subscription for the MultiSub achievement tracker,
    and initiates Start-AvdDiscovery for the new subscription.
#>
function Switch-Subscription {
    param([string]$SubId, [string]$SubName)
    if (-not $SubId) { return }

    Write-DebugLog "[AUTH] Switching subscription to: $SubId" -Level 'INFO'
    Update-StatusBar "Switching subscription..."

    Start-BackgroundWork -Name 'SwitchSub' -ScriptBlock {
        param($sid, $sname)
        Set-AzContext -SubscriptionId $sid -ErrorAction Stop | Out-Null
        return @{ SubId = $sid; SubName = $sname }
    } -Arguments @($SubId, $SubName) -OnComplete {
        param($Result)
        $Global:CurrentSubId = $Result.SubId
        $Global:CurrentTopology = $null
        $Global:TopologyCacheTime = $null
        Write-DebugLog "[AUTH] Subscription switched to $($Result.SubName) ($($Result.SubId))" -Level 'SUCCESS'
        Update-StatusBar "Subscription: $($Result.SubName)"
        $Global:SubsSeen[$Result.SubId] = $true
        if ($Global:SubsSeen.Count -ge 3) { Check-Achievement 'MultiSub' }
    }
}

# ===============================================================================
# SECTION 21: EVENT WIRING
# ===============================================================================

# --- Title bar ---
$btnMinimize.Add_Click({ $Window.WindowState = 'Minimized' })
$btnMaximize.Add_Click({
    if ($Window.WindowState -eq 'Maximized') { $Window.WindowState = 'Normal' }
    else { $Window.WindowState = 'Maximized' }
})
$btnClose.Add_Click({ $Window.Close() })
$titleBar.Add_MouseLeftButtonDown({
    param($s, $e)
    if ($e.ChangedButton -eq 'Left') {
        if ($e.ClickCount -eq 2) {
            if ($Window.WindowState -eq 'Maximized') { $Window.WindowState = 'Normal' }
            else { $Window.WindowState = 'Maximized' }
        } else { $Window.DragMove() }
    }
})

# --- Theme toggle ---
$btnThemeToggle.Add_Click({
    $Global:IsLightMode = -not $Global:IsLightMode
    & $ApplyTheme $Global:IsLightMode
    Save-UserPrefs
    $Global:ThemeToggleCount++
    if ($Global:ThemeToggleCount -ge 5) { Check-Achievement 'ThemeSwitcher' }
    Write-DebugLog "Theme toggled to $(if ($Global:IsLightMode) {'Light'} else {'Dark'}) mode" -Level 'INFO'
})

# --- Auth buttons ---
$btnLogin.Add_Click({ Connect-ToAzure })
$btnLogout.Add_Click({ Disconnect-FromAzure })

# --- Subscription combo ---
$cmbSubscription.Add_SelectionChanged({
    $sel = $cmbSubscription.SelectedItem
    if ($sel -and $sel.Tag) {
        Switch-Subscription -SubId $sel.Tag -SubName $sel.Content
    }
})

# --- Icon rail navigation ---
$railDashboard.Add_Click({ Switch-Tab 'Dashboard' })
$railTopology.Add_Click({ Switch-Tab 'Topology' })
$railBackups.Add_Click({ Switch-Tab 'Backups' })
$railRestore.Add_Click({ Switch-Tab 'Restore' })
$railSettings.Add_Click({ Switch-Tab 'Settings' })

# --- Tab headers ---
$tabDashboard.Add_MouseLeftButtonDown({ Switch-Tab 'Dashboard' })
$tabTopology.Add_MouseLeftButtonDown({ Switch-Tab 'Topology' })
$tabBackups.Add_MouseLeftButtonDown({ Switch-Tab 'Backups' })
$tabRestore.Add_MouseLeftButtonDown({ Switch-Tab 'Restore' })
$tabSettings.Add_MouseLeftButtonDown({ Switch-Tab 'Settings' })

# --- Backup Now button ---
$btnBackupNow.Add_Click({
    if (-not $Global:CurrentTopology) {
        Show-Toast -Message "No topology loaded. Run a discovery first." -Type 'Warning'
        return
    }
    Write-DebugLog "[BACKUP] Manual backup triggered" -Level 'INFO'
    Update-StatusBar "Saving backup..."
    Show-GlobalProgressIndeterminate -Text "Saving backup..."
    $Label = ''
    $result = Show-ThemedDialog -Title 'Backup Label' -Message 'Enter an optional label for this backup:' `
        -InputPrompt 'Label (optional)' `
        -Buttons @( @{ Text='Save'; IsAccent=$true; Result='Save' }, @{ Text='Cancel'; IsAccent=$false; Result='Cancel' } )
    if ($result.Result -eq 'Save') {
        $Label = $result.Input
    } else { Hide-GlobalProgress; return }

    try {
        $Path = Save-AvdBackup -Topology $Global:CurrentTopology -Label $Label
        Update-BackupList
        Show-Toast -Message "Backup saved: $(Split-Path $Path -Leaf)" -Type 'Success'
        Update-StatusBar "Backup saved"
        Hide-GlobalProgress
        Check-Achievement 'FirstBackup'
        if ((Get-BackupList).Count -ge 10) { Check-Achievement 'BackupHoarder' }
    } catch {
        Write-DebugLog "[BACKUP] Save failed: $($_.Exception.Message)" -Level 'ERROR'
        Show-Toast -Message "Backup failed: $($_.Exception.Message)" -Type 'Error'
        Update-StatusBar "Backup failed"
        Hide-GlobalProgress
    }
})

# --- Refresh button ---
$btnRefreshTopology.Add_Click({ Start-AvdDiscovery })

# --- Export button ---
$btnExportReport.Add_Click({
    Export-TopologyReport -Format 'JSON' -Topology $Global:CurrentTopology
})

# --- Backup list selection ---
$lstBackups.Add_SelectionChanged({
    $sel = $lstBackups.SelectedItem
    if ($sel -and $sel.Tag) {
        $Global:SelectedBackup = $sel.Tag
        Write-DebugLog "[BACKUP] Selected: $($sel.Tag.FileName)" -Level 'DEBUG'
    }
})

# --- Compare button (Backups tab) ---
$btnCompareWithCurrent.Add_Click({
    if (-not $Global:SelectedBackup) {
        Show-Toast -Message "Select a backup to compare against." -Type 'Warning'
        return
    }
    if (-not $Global:CurrentTopology) {
        Show-Toast -Message "No live topology. Run a discovery first." -Type 'Warning'
        return
    }

    Write-DebugLog "[DIFF] Comparing backup '$($Global:SelectedBackup.FileName)' vs current" -Level 'INFO'
    Update-StatusBar "Comparing..."

    $DiffResult = Compare-AvdTopology -Baseline $Global:SelectedBackup.Topology -Current $Global:CurrentTopology
    $Global:LastDiffResult = $DiffResult

    # Show diff overlay
    $Window.Dispatcher.Invoke([Action]{
        $pnlDiffOverlay.Visibility = 'Visible'
        $pnlDiffResults.Children.Clear()
        $reportRoot = [System.Windows.Controls.StackPanel]::new()
        $reportRoot.Margin = [System.Windows.Thickness]::new(0)

        $shortArmId = {
            param([string]$id)
            if ([string]::IsNullOrWhiteSpace($id)) { return $id }
            if ($id -match '/providers/[^/]+/([^/]+)/([^/]+)$') {
                return "$($Matches[1])/$($Matches[2])"
            }
            return $id
        }

        $formatDiffValue = {
            param($raw)

            $text = if ($null -eq $raw) { '(null)' } else { [string]$raw }
            if ([string]::IsNullOrWhiteSpace($text)) { return '(empty)' }

            $parsed = $null
            $isJson = $false
            if (($text.StartsWith('[') -or $text.StartsWith('{') -or $text.StartsWith('"')) -and $text.Length -gt 1) {
                try {
                    $parsed = $text | ConvertFrom-Json -ErrorAction Stop
                    # ConvertFrom-Json unwraps single-element arrays via pipeline — re-wrap
                    if ($text.StartsWith('[') -and $parsed -isnot [System.Array]) {
                        $parsed = @($parsed)
                    }
                    $isJson = $true
                } catch { }
            }

            if ($isJson -and $parsed -is [System.Array]) {
                if ($parsed.Count -eq 0) { return '(empty list)' }
                $items = @()
                foreach ($v in $parsed) {
                    $itemText = if ($null -eq $v) { '(null)' } else { [string]$v }
                    if ($itemText -match '^/subscriptions/.+') { $itemText = & $shortArmId $itemText }
                    $items += "- $itemText"
                }
                return ($items -join "`n")
            }

            if ($isJson -and $parsed -is [PSCustomObject]) {
                # Detect collection metadata marker from normalization layer
                if ($parsed.PSObject.Properties.Match('__collectionMeta').Count) {
                    $n = if ($parsed.PSObject.Properties.Match('count').Count) { $parsed.count } else { '?' }
                    return "($n item(s) - content not available in this backup)"
                }
                # Detect collection metadata from broken array serialization (e.g. List[String]
                # serialized as {"Length":N} instead of as a JSON array)
                $pNames = @($parsed.PSObject.Properties.Name)
                $metaNames = @('Length','LongLength','Count','Capacity','Rank','IsReadOnly','IsFixedSize','SyncRoot','IsSynchronized')
                if (($pNames | Where-Object { $_ -notin $metaNames }).Count -eq 0) {
                    $n = if ($parsed.PSObject.Properties.Match('Count').Count) { $parsed.Count }
                         elseif ($parsed.PSObject.Properties.Match('Length').Count) { $parsed.Length }
                         else { '?' }
                    return "($n item(s) - content not available in this backup)"
                }
                $pairs = @()
                foreach ($p in ($parsed.PSObject.Properties | Sort-Object Name)) {
                    $pairs += "$($p.Name)=$($p.Value)"
                }
                if ($pairs.Count -gt 0) { return ($pairs -join '; ') }
            }

            if ($text -match '^".*"$') {
                try { $text = [string]($text | ConvertFrom-Json -ErrorAction Stop) } catch { }
            }
            if ($text -match '^/subscriptions/.+') { $text = & $shortArmId $text }
            if ($text.Length -gt 220) { $text = $text.Substring(0,220) + '...' }
            return $text
        }

        $newBadge = {
            param(
                [string]$label,
                [int]$count,
                [string]$brushKey,
                [string]$bgKey,
                [string]$icon
            )
            $badge = [System.Windows.Controls.Border]::new()
            $badge.CornerRadius = [System.Windows.CornerRadius]::new(8)
            $badge.BorderThickness = [System.Windows.Thickness]::new(1)
            $badge.Padding = [System.Windows.Thickness]::new(10,6,10,6)
            $badge.Margin = [System.Windows.Thickness]::new(0,0,8,8)
            $badge.BorderBrush = $Window.FindResource($brushKey)
            $badge.Background = $Window.FindResource($bgKey)

            $sp = [System.Windows.Controls.StackPanel]::new()
            $sp.Orientation = 'Horizontal'

            $iconTb = [System.Windows.Controls.TextBlock]::new()
            $iconTb.Text = $icon
            $iconTb.FontSize = 12
            $iconTb.FontWeight = [System.Windows.FontWeights]::Bold
            $iconTb.Margin = [System.Windows.Thickness]::new(0,0,6,0)
            $iconTb.Foreground = $Window.FindResource($brushKey)
            [void]$sp.Children.Add($iconTb)

            $countTb = [System.Windows.Controls.TextBlock]::new()
            $countTb.Text = "$count"
            $countTb.FontSize = 13
            $countTb.FontWeight = [System.Windows.FontWeights]::SemiBold
            $countTb.Margin = [System.Windows.Thickness]::new(0,0,6,0)
            $countTb.Foreground = $Window.FindResource('ThemeTextPrimary')
            [void]$sp.Children.Add($countTb)

            $labelTb = [System.Windows.Controls.TextBlock]::new()
            $labelTb.Text = $label
            $labelTb.FontSize = 11
            $labelTb.Foreground = $Window.FindResource('ThemeTextSecondary')
            [void]$sp.Children.Add($labelTb)

            $badge.Child = $sp
            return $badge
        }

        $hero = [System.Windows.Controls.Border]::new()
        $hero.CornerRadius = [System.Windows.CornerRadius]::new(8)
        $hero.BorderThickness = [System.Windows.Thickness]::new(1)
        $hero.Padding = [System.Windows.Thickness]::new(12,10,12,10)
        $hero.Margin = [System.Windows.Thickness]::new(0,0,0,10)
        $hero.BorderBrush = $Window.FindResource('ThemeAccent')
        $hero.Background = $Window.FindResource('ThemeAccentDim')

        $heroGrid = [System.Windows.Controls.Grid]::new()
        [void]$heroGrid.ColumnDefinitions.Add([System.Windows.Controls.ColumnDefinition]::new())
        [void]$heroGrid.ColumnDefinitions.Add([System.Windows.Controls.ColumnDefinition]::new())

        $heroLeft = [System.Windows.Controls.StackPanel]::new()
        $heroTitle = [System.Windows.Controls.TextBlock]::new()
        $heroTitle.Text = 'Drift Intelligence Report'
        $heroTitle.FontSize = 14
        $heroTitle.FontWeight = [System.Windows.FontWeights]::SemiBold
        $heroTitle.Foreground = $Window.FindResource('ThemeTextPrimary')
        [void]$heroLeft.Children.Add($heroTitle)

        $heroSub = [System.Windows.Controls.TextBlock]::new()
        $heroSub.Text = "Backup: $($Global:SelectedBackup.FileName)"
        $heroSub.FontSize = 11
        $heroSub.Foreground = $Window.FindResource('ThemeTextSecondary')
        [void]$heroLeft.Children.Add($heroSub)
        [System.Windows.Controls.Grid]::SetColumn($heroLeft, 0)
        [void]$heroGrid.Children.Add($heroLeft)

        $heroRight = [System.Windows.Controls.TextBlock]::new()
        $heroRight.Text = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        $heroRight.FontSize = 11
        $heroRight.HorizontalAlignment = 'Right'
        $heroRight.VerticalAlignment = 'Center'
        $heroRight.Foreground = $Window.FindResource('ThemeTextMuted')
        [System.Windows.Controls.Grid]::SetColumn($heroRight, 1)
        [void]$heroGrid.Children.Add($heroRight)

        $hero.Child = $heroGrid
        [void]$reportRoot.Children.Add($hero)

        $badgeWrap = [System.Windows.Controls.WrapPanel]::new()
        $badgeWrap.Margin = [System.Windows.Thickness]::new(0,0,0,10)
        [void]$badgeWrap.Children.Add((& $newBadge 'Added' $DiffResult.Added.Count 'ThemeSuccess' 'ThemeSuccessBg' '+'))
        [void]$badgeWrap.Children.Add((& $newBadge 'Removed' $DiffResult.Removed.Count 'ThemeError' 'ThemeErrorDim' '-'))
        [void]$badgeWrap.Children.Add((& $newBadge 'Modified' $DiffResult.Modified.Count 'ThemeWarning' 'ThemeWarningBg' '~'))
        [void]$badgeWrap.Children.Add((& $newBadge 'Unchanged' $DiffResult.Unchanged.Count 'ThemeTextMuted' 'ThemeCardBg' '='))
        [void]$reportRoot.Children.Add($badgeWrap)

        $newSectionHeader = {
            param([string]$title, [string]$brushKey)
            $tb = [System.Windows.Controls.TextBlock]::new()
            $tb.Text = $title
            $tb.FontSize = 11
            $tb.FontWeight = [System.Windows.FontWeights]::Bold
            $tb.Margin = [System.Windows.Thickness]::new(0,8,0,6)
            $tb.Foreground = $Window.FindResource($brushKey)
            return $tb
        }

        if ($DiffResult.Added.Count -gt 0) {
            [void]$reportRoot.Children.Add((& $newSectionHeader "ADDED ($($DiffResult.Added.Count))" 'ThemeSuccess'))
            foreach ($a in $DiffResult.Added) {
                $tb = [System.Windows.Controls.TextBlock]::new()
                $tb.Text = "+ [$($a.ObjectType)] $($a.ObjectName)"
                $tb.FontFamily = [System.Windows.Media.FontFamily]::new('Cascadia Mono, Consolas')
                $tb.FontSize = 12
                $tb.Margin = [System.Windows.Thickness]::new(4,0,0,2)
                $tb.Foreground = $Window.FindResource('ThemeTextBody')
                [void]$reportRoot.Children.Add($tb)
            }
        }
        if ($DiffResult.Removed.Count -gt 0) {
            [void]$reportRoot.Children.Add((& $newSectionHeader "REMOVED ($($DiffResult.Removed.Count))" 'ThemeError'))
            foreach ($r in $DiffResult.Removed) {
                $tb = [System.Windows.Controls.TextBlock]::new()
                $tb.Text = "- [$($r.ObjectType)] $($r.ObjectName)"
                $tb.FontFamily = [System.Windows.Media.FontFamily]::new('Cascadia Mono, Consolas')
                $tb.FontSize = 12
                $tb.Margin = [System.Windows.Thickness]::new(4,0,0,2)
                $tb.Foreground = $Window.FindResource('ThemeTextBody')
                [void]$reportRoot.Children.Add($tb)
            }
        }
        if ($DiffResult.Modified.Count -gt 0) {
            [void]$reportRoot.Children.Add((& $newSectionHeader "MODIFIED ($($DiffResult.Modified.Count))" 'ThemeWarning'))
            foreach ($m in $DiffResult.Modified) {
                $card = [System.Windows.Controls.Border]::new()
                $card.CornerRadius = [System.Windows.CornerRadius]::new(8)
                $card.BorderThickness = [System.Windows.Thickness]::new(1)
                $card.Padding = [System.Windows.Thickness]::new(10,8,10,8)
                $card.Margin = [System.Windows.Thickness]::new(0,0,0,8)
                $card.Background = $Window.FindResource('ThemeCardBg')
                $card.BorderBrush = $Window.FindResource('ThemeBorderMedium')

                $cardSp = [System.Windows.Controls.StackPanel]::new()
                $head = [System.Windows.Controls.TextBlock]::new()
                $head.Text = "~ [$($m.ObjectType)] $($m.ObjectName)"
                $head.FontFamily = [System.Windows.Media.FontFamily]::new('Cascadia Mono, Consolas')
                $head.FontWeight = [System.Windows.FontWeights]::SemiBold
                $head.FontSize = 12
                $head.Foreground = $Window.FindResource('ThemeTextPrimary')
                $head.Margin = [System.Windows.Thickness]::new(0,0,0,4)
                [void]$cardSp.Children.Add($head)

                if ($m.ChangedProperties -and $m.ChangedProperties.Count -gt 0) {
                    foreach ($cp in $m.ChangedProperties) {
                        $summaryText = if ($cp.Summary) { $cp.Summary } else { 'changed' }
                        $prop = [System.Windows.Controls.TextBlock]::new()
                        $prop.Text = "* $($cp.Property): $summaryText"
                        $prop.FontSize = 11
                        $prop.Foreground = $Window.FindResource('ThemeTextSecondary')
                        $prop.Margin = [System.Windows.Thickness]::new(0,2,0,2)
                        [void]$cardSp.Children.Add($prop)

                        if ($cp.PSObject.Properties.Match('Before').Count -and $cp.PSObject.Properties.Match('After').Count) {
                            $beforeText = & $formatDiffValue $cp.Before
                            $afterText  = & $formatDiffValue $cp.After

                            $beforeLbl = [System.Windows.Controls.TextBlock]::new()
                            $beforeLbl.Text = 'before'
                            $beforeLbl.FontSize = 10
                            $beforeLbl.FontWeight = [System.Windows.FontWeights]::Bold
                            $beforeLbl.Foreground = $Window.FindResource('ThemeError')
                            $beforeLbl.Margin = [System.Windows.Thickness]::new(10,1,0,0)
                            [void]$cardSp.Children.Add($beforeLbl)
                            foreach ($line in @($beforeText -split "`n")) {
                                $beforeLine = [System.Windows.Controls.TextBlock]::new()
                                $beforeLine.Text = "  $line"
                                $beforeLine.FontFamily = [System.Windows.Media.FontFamily]::new('Cascadia Mono, Consolas')
                                $beforeLine.FontSize = 10
                                $beforeLine.TextWrapping = 'Wrap'
                                $beforeLine.Foreground = $Window.FindResource('ThemeTextBody')
                                $beforeLine.Margin = [System.Windows.Thickness]::new(14,0,0,0)
                                [void]$cardSp.Children.Add($beforeLine)
                            }

                            $afterLbl = [System.Windows.Controls.TextBlock]::new()
                            $afterLbl.Text = 'after'
                            $afterLbl.FontSize = 10
                            $afterLbl.FontWeight = [System.Windows.FontWeights]::Bold
                            $afterLbl.Foreground = $Window.FindResource('ThemeSuccess')
                            $afterLbl.Margin = [System.Windows.Thickness]::new(10,1,0,0)
                            [void]$cardSp.Children.Add($afterLbl)
                            foreach ($line in @($afterText -split "`n")) {
                                $afterLine = [System.Windows.Controls.TextBlock]::new()
                                $afterLine.Text = "  $line"
                                $afterLine.FontFamily = [System.Windows.Media.FontFamily]::new('Cascadia Mono, Consolas')
                                $afterLine.FontSize = 10
                                $afterLine.TextWrapping = 'Wrap'
                                $afterLine.Foreground = $Window.FindResource('ThemeTextBody')
                                $afterLine.Margin = [System.Windows.Thickness]::new(14,0,0,0)
                                [void]$cardSp.Children.Add($afterLine)
                            }
                        } elseif ($cp.Summary) {
                            $changeLine = [System.Windows.Controls.TextBlock]::new()
                            $changeLine.Text = "  change: $($cp.Summary)"
                            $changeLine.FontFamily = [System.Windows.Media.FontFamily]::new('Cascadia Mono, Consolas')
                            $changeLine.FontSize = 10
                            $changeLine.Foreground = $Window.FindResource('ThemeTextDim')
                            $changeLine.Margin = [System.Windows.Thickness]::new(14,0,0,0)
                            [void]$cardSp.Children.Add($changeLine)
                        }
                    }
                } elseif ($m.Details) {
                    $detailTb = [System.Windows.Controls.TextBlock]::new()
                    $detailTb.Text = "details: $($m.Details)"
                    $detailTb.FontSize = 11
                    $detailTb.Foreground = $Window.FindResource('ThemeTextSecondary')
                    [void]$cardSp.Children.Add($detailTb)
                }

                $card.Child = $cardSp
                [void]$reportRoot.Children.Add($card)
            }
        }

        if ($DiffResult.Added.Count -eq 0 -and $DiffResult.Removed.Count -eq 0 -and $DiffResult.Modified.Count -eq 0) {
            $allGood = [System.Windows.Controls.TextBlock]::new()
            $allGood.Text = 'No drift detected. Backup and live topology are aligned.'
            $allGood.Margin = [System.Windows.Thickness]::new(0,6,0,2)
            $allGood.FontSize = 12
            $allGood.Foreground = $Window.FindResource('ThemeSuccessText')
            [void]$reportRoot.Children.Add($allGood)
        }

        [void]$pnlDiffResults.Children.Add($reportRoot)
    })

    Update-DriftBanner -DiffResults $DiffResult
    Update-RestoreTree -DiffResults $DiffResult -BackupTopology $Global:SelectedBackup.Topology
    Update-StatusBar "Comparison complete"
    Show-Toast -Message "Diff complete: $($DiffResult.Added.Count + $DiffResult.Removed.Count + $DiffResult.Modified.Count) change(s)" -Type 'Info'
    Check-Achievement 'FirstDiff'
})

# --- Restore buttons ---
$btnDryRun.Add_Click({
    if ($Global:DiscoveryInProgress) {
        Show-Toast -Message "Discovery is still running. Wait for it to complete before dry-run." -Type 'Warning'
        return
    }
    if (-not $Global:CurrentTopology) {
        Show-Toast -Message "Run Refresh first to discover the live topology before restoring." -Type 'Warning'
        return
    }
    if (-not $Global:RestoreManifest -or $Global:RestoreManifest.Count -eq 0) {
        Show-Toast -Message "No restore items. Run a diff first." -Type 'Warning'
        return
    }

    $Selected = @($Global:RestoreManifest | Where-Object Selected)
    if ($Selected.Count -eq 0) {
        Show-Toast -Message "No items selected for restore." -Type 'Warning'
        return
    }

    Write-DebugLog "[RESTORE] Dry run: $($Selected.Count) items" -Level 'INFO'
    Update-StatusBar "Running dry-run..."
    Show-GlobalProgressIndeterminate -Text "Running dry-run validation..."
    $Results = Invoke-FullRestore -Manifest $Global:RestoreManifest -DryRun
    $Passed = @($Results | Where-Object Success).Count
    Show-Toast -Message "Dry-run complete: $Passed/$($Selected.Count) would succeed" -Type 'Info'
    Update-StatusBar "Dry-run complete"
    Hide-GlobalProgress
    Check-Achievement 'FirstDryRun'
})

$btnExportPlan.Add_Click({
    if ($Global:DiscoveryInProgress) {
        Show-Toast -Message "Discovery is still running. Wait for it to complete before exporting a restore plan." -Type 'Warning'
        return
    }
    if (-not $Global:CurrentTopology) {
        Show-Toast -Message "Run Refresh first to discover the live topology before exporting a restore plan." -Type 'Warning'
        return
    }
    if (-not $Global:RestoreManifest -or $Global:RestoreManifest.Count -eq 0) {
        Show-Toast -Message 'No restore items. Load a backup and run a diff first.' -Type 'Warning'
        return
    }
    Export-RestorePlan -Manifest $Global:RestoreManifest
})

$btnCopyInstallCmd.Add_Click({
    if ($Global:PrereqInstallCmd) {
        [System.Windows.Clipboard]::SetText($Global:PrereqInstallCmd)
        Show-Toast -Message 'Install command copied to clipboard!' -Type 'Success'
        Write-DebugLog "[SETTINGS] Prereq install command copied: $($Global:PrereqInstallCmd)" -Level 'INFO'
    }
})

$cmbRestoreSource.Add_SelectionChanged({
    $sel = $cmbRestoreSource.SelectedItem
    if ($sel -and $sel.Tag) {
        $Global:SelectedRestoreBackup = $sel.Tag
        Populate-RgMappingGrid -BackupTopology $sel.Tag.Topology
        # Load and diff against current topology
        if ($Global:CurrentTopology) {
            $DiffResult = Compare-AvdTopology -Baseline $sel.Tag.Topology -Current $Global:CurrentTopology
            Update-RestoreTree -DiffResults $DiffResult -BackupTopology $sel.Tag.Topology
        } else {
            # No current topology - show all backup items as creatable
            $backupTopo = $sel.Tag.Topology
            $fakeRemoved = [System.Collections.ArrayList]::new()
            foreach ($hp in @($backupTopo.HostPools)) {
                [void]$fakeRemoved.Add(@{ ObjectType='HostPool'; ObjectName=$hp.Name; ResourceId=$hp.ResourceId; Data=$hp })
                foreach ($ag in @($hp.AppGroups)) {
                    [void]$fakeRemoved.Add(@{ ObjectType='AppGroup'; ObjectName=$ag.Name; ResourceId=$ag.ResourceId; Data=$ag })
                }
            }
            foreach ($ws in @($backupTopo.Workspaces)) {
                [void]$fakeRemoved.Add(@{ ObjectType='Workspace'; ObjectName=$ws.Name; ResourceId=$ws.ResourceId; Data=$ws })
            }
            $fakeDiff = @{ Added=[System.Collections.ArrayList]::new(); Removed=$fakeRemoved; Modified=[System.Collections.ArrayList]::new(); Unchanged=[System.Collections.ArrayList]::new() }
            Update-RestoreTree -DiffResults $fakeDiff -BackupTopology $backupTopo
            $lblRestoreEmpty.Text = 'No live topology loaded. Showing all backup items as restorable. Run Refresh for precise diff.'
        }
    }
})

$rdoDiffRegion.Add_Checked({ $cmbTargetRegion.Visibility = 'Visible' })
$rdoSameRegion.Add_Checked({ $cmbTargetRegion.Visibility = 'Collapsed' })

$chkDiscoverVms.Add_Checked({
    # Re-trigger restore source selection to rebuild tree with discovery
    if ($cmbRestoreSource.SelectedItem -and $cmbRestoreSource.SelectedItem.Tag) {
        Write-DebugLog "[VM-DISCOVER] Discovery enabled - refreshing restore tree" -Level "INFO"
        $lblDiscoverVmsHint.Text = "Querying Azure for additional VMs..."
        $lblDiscoverVmsHint.Visibility = "Visible"
        $sel = $cmbRestoreSource.SelectedItem
        if ($Global:CurrentTopology) {
            $DiffResult = Compare-AvdTopology -Baseline $sel.Tag.Topology -Current $Global:CurrentTopology
        } else {
            $backupTopo = $sel.Tag.Topology
            $fakeRemoved = [System.Collections.ArrayList]::new()
            foreach ($hp in @($backupTopo.HostPools)) {
                [void]$fakeRemoved.Add(@{ ObjectType="HostPool"; ObjectName=$hp.Name; ResourceId=$hp.ResourceId; Data=$hp })
                foreach ($ag in @($hp.AppGroups)) { [void]$fakeRemoved.Add(@{ ObjectType="AppGroup"; ObjectName=$ag.Name; ResourceId=$ag.ResourceId; Data=$ag }) }
            }
            foreach ($ws in @($backupTopo.Workspaces)) { [void]$fakeRemoved.Add(@{ ObjectType="Workspace"; ObjectName=$ws.Name; ResourceId=$ws.ResourceId; Data=$ws }) }
            $DiffResult = @{ Added=[System.Collections.ArrayList]::new(); Removed=$fakeRemoved; Modified=[System.Collections.ArrayList]::new(); Unchanged=[System.Collections.ArrayList]::new() }
        }
        # Discover additional VMs for each host pool — save originals first to avoid mutating backup
        $Global:OriginalSessionHosts = @{}
        foreach ($hp in @($sel.Tag.Topology.HostPools)) {
            $Global:OriginalSessionHosts[$hp.Name] = @($hp.SessionHosts)
            if ($hp.SessionHosts -and $hp.SessionHosts.Count -gt 0) {
                $hpRG = $hp.ResourceGroupName
                # Apply RG mapping if any
                if ($Global:RgMappings -and $Global:RgMappings.ContainsKey($hpRG)) { $hpRG = $Global:RgMappings[$hpRG] }
                try {
                    $extra = Find-SessionHostsByPattern -HostPoolName $hp.Name -ResourceGroupName $hpRG -BackupSessionHosts $hp.SessionHosts
                    if ($extra.Count -gt 0) {
                        $hp.SessionHosts = @($hp.SessionHosts) + @($extra)
                        Write-DebugLog "[VM-DISCOVER] Added $($extra.Count) discovered VM(s) to $($hp.Name)" -Level "INFO"
                    }
                } catch {
                    Write-DebugLog "[VM-DISCOVER] Error discovering VMs for $($hp.Name): $($_.Exception.Message)" -Level "WARN"
                }
            }
        }
        Update-RestoreTree -DiffResults $DiffResult -BackupTopology $sel.Tag.Topology
        $lblDiscoverVmsHint.Text = "Discovery complete."
    }
})

$chkDiscoverVms.Add_Unchecked({
    # Restore original session hosts and reload
    if ($cmbRestoreSource.SelectedItem -and $cmbRestoreSource.SelectedItem.Tag) {
        Write-DebugLog "[VM-DISCOVER] Discovery disabled - reverting to backup data" -Level "INFO"
        $lblDiscoverVmsHint.Visibility = "Collapsed"
        # Restore original session host arrays
        $sel = $cmbRestoreSource.SelectedItem
        if ($Global:OriginalSessionHosts) {
            foreach ($hp in @($sel.Tag.Topology.HostPools)) {
                if ($Global:OriginalSessionHosts.ContainsKey($hp.Name)) {
                    $hp.SessionHosts = $Global:OriginalSessionHosts[$hp.Name]
                }
            }
            $Global:OriginalSessionHosts = $null
        }
        # Force re-select to reload original data
        $idx = $cmbRestoreSource.SelectedIndex
        $cmbRestoreSource.SelectedIndex = -1
        $cmbRestoreSource.SelectedIndex = $idx
    }
})

$btnExecuteRestore.Add_Click({
    if ($Global:DiscoveryInProgress) {
        Show-Toast -Message "Discovery is still running. Wait for it to complete before executing restore." -Type 'Warning'
        return
    }
    if (-not $Global:CurrentTopology) {
        Show-Toast -Message "Run Refresh first to discover the live topology before restoring." -Type 'Warning'
        return
    }
    if (-not $Global:RestoreManifest -or $Global:RestoreManifest.Count -eq 0) {
        Show-Toast -Message "No restore items. Run a diff first." -Type 'Warning'
        return
    }
    $Selected = @($Global:RestoreManifest | Where-Object Selected)
    if ($Selected.Count -eq 0) {
        Show-Toast -Message "No items selected for restore." -Type 'Warning'
        return
    }

    $confirm = Show-ThemedDialog -Title 'Confirm Restore' `
        -Message "You are about to restore $($Selected.Count) item(s). This will CREATE or UPDATE resources in your Azure environment.`n`nThis action cannot be easily undone. Type RESTORE below to confirm." `
        -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' `
        -InputPrompt 'Type RESTORE to confirm' -InputMatch 'RESTORE' `
        -Buttons @( @{ Text='Execute Restore'; IsAccent=$true; Result='Confirmed' }, @{ Text='Cancel'; IsAccent=$false; Result='Cancel' } )

    if ($confirm.Result -ne 'Confirmed') { return }

    # Apply target region override if cross-region restore selected
        if ($rdoDiffRegion.IsChecked -and $cmbTargetRegion.SelectedItem) {
            $targetRgn = $cmbTargetRegion.SelectedItem.Content
            Write-DebugLog "[RESTORE] Cross-region restore: target=$targetRgn" -Level 'INFO'
            foreach ($entry in $Global:RestoreManifest) {
                $entry | Add-Member -NotePropertyName 'TargetRegion' -NotePropertyValue $targetRgn -Force
            }
        }
        Write-DebugLog "[RESTORE] User confirmed restore of $($Selected.Count) items" -Level 'WARN'
    Update-StatusBar "Restoring $($Selected.Count) item(s)..."
    Show-GlobalProgressIndeterminate -Text "Restoring $($Selected.Count) item(s)..."


    # Apply RG mappings to manifest
    Apply-RgMappings -Manifest $Global:RestoreManifest

    # Determine default location for RG creation
    $defaultLoc = if ($rdoDiffRegion.IsChecked -and $cmbTargetRegion.SelectedItem) { $cmbTargetRegion.SelectedItem.Content } else {
        $sample = $Global:RestoreManifest | Where-Object { $_.BackupData -and $_.BackupData.PSObject.Properties.Match("Location").Count -and $_.BackupData.Location } | Select-Object -First 1
        if ($sample) { $sample.BackupData.Location } else { "eastus" }
    }
    $autoCreateRg = $chkAutoCreateRg.IsChecked

    Start-BackgroundWork -Name 'RestoreOp' -Functions @('Invoke-FullRestore','Invoke-RestoreItem','Register-SessionHost','Ensure-ResourceGroups','Write-DebugLog') -ScriptBlock {
        param($Manifest, $AutoCreateRg, $DefaultLoc)
        if ($AutoCreateRg) { Ensure-ResourceGroups -Manifest $Manifest -DefaultLocation $DefaultLoc }
        Invoke-FullRestore -Manifest $Manifest
    } -Arguments @($Global:RestoreManifest, $autoCreateRg, $defaultLoc) -OnComplete {
        param($Results)
        Stop-Shimmer
        $Passed = @($Results | Where-Object Success).Count
        $Failed = @($Results | Where-Object { -not $_.Success }).Count
        if ($Failed -gt 0) {
            foreach ($fr in @($Results | Where-Object { -not $_.Success })) {
                $objType = if ($fr.Item -and $fr.Item.ObjectType) { $fr.Item.ObjectType } else { 'UnknownType' }
                $objName = if ($fr.Item -and $fr.Item.ObjectName) { $fr.Item.ObjectName } else { 'UnknownName' }
                $reason  = if ($fr.Message) { $fr.Message } else { 'Unknown error' }
                Write-DebugLog "[RESTORE] FAILED [$objType] '$objName' - $reason" -Level 'ERROR'
                if ($objType -eq 'UnknownType' -or $objName -eq 'UnknownName') {
                    $raw = try { $fr | ConvertTo-Json -Depth 8 -Compress } catch { "<serialize-failed: $($_.Exception.Message)>" }
                    Write-DebugLog "[RESTORE] FAILED RAW: $raw" -Level 'DEBUG'
                }
            }
        }
        $Level = if ($Failed -gt 0) { 'WARN' } else { 'SUCCESS' }
        Write-DebugLog "[RESTORE] Complete: $Passed succeeded, $Failed failed" -Level $Level
        Show-Toast -Message "Restore complete: $Passed succeeded, $Failed failed" -Type $(if ($Failed -gt 0) { 'Warning' } else { 'Success' })
        Update-StatusBar "Restore complete"
        Hide-GlobalProgress
        Check-Achievement 'FirstRestore'
        if ($Failed -eq 0 -and $Passed -ge 5) { Check-Achievement 'PerfectRestore' }
    }
})

# --- Settings toggles ---
$chkDarkMode.Add_Checked({
    $Global:IsLightMode = $false
    & $ApplyTheme $false
    Save-UserPrefs
})
$chkDarkMode.Add_Unchecked({
    $Global:IsLightMode = $true
    & $ApplyTheme $true
    Save-UserPrefs
})
$chkAutoBackup.Add_Checked({
    $Global:AutoBackupEnabled = $true
    Save-UserPrefs
    Write-DebugLog "[SETTINGS] Auto-backup enabled" -Level 'INFO'
})
$chkAutoBackup.Add_Unchecked({
    $Global:AutoBackupEnabled = $false
    Save-UserPrefs
    Write-DebugLog "[SETTINGS] Auto-backup disabled" -Level 'INFO'
})
$chkRetentionEnabled.Add_Checked({
    $pnlRetentionDays.Visibility = 'Visible'
    Save-UserPrefs
    Write-DebugLog "[SETTINGS] Retention enabled" -Level 'DEBUG'
})
$chkRetentionEnabled.Add_Unchecked({
    $pnlRetentionDays.Visibility = 'Collapsed'
    Save-UserPrefs
    Write-DebugLog "[SETTINGS] Retention disabled" -Level 'DEBUG'
})
$txtRetentionDays.Add_LostFocus({
    $val = $txtRetentionDays.Text -replace '[^\d]',''
    if ($val -and [int]$val -gt 0) {
        $Global:RetentionDays = [int]$val
        Save-UserPrefs
        Write-DebugLog "[SETTINGS] Retention set to $($Global:RetentionDays) days" -Level 'DEBUG'
    }
})

# --- Window closing ---
$Window.Add_Closing({
    Write-DebugLog "[APP] Window closing - saving preferences" -Level 'INFO'
    Save-UserPrefs
})

$Window.Add_ContentRendered({
    $Global:WindowRendered = $true
    Write-DebugLog "[APP] Window content rendered" -Level 'DEBUG'
    if ($Global:SplashDismissQueued) {
        $Global:SplashDismissQueued = $false
        Start-SplashDismissTimer
    }
})

# ===============================================================================
# SECTION 22: FILTER & SEARCH
# ===============================================================================

$txtTopoSearch.Add_TextChanged({
    $query = $txtTopoSearch.Text.Trim()
    Filter-TopologyTree -Query $query
})

<#
.SYNOPSIS
    Performs live text search on the topology TreeView, showing only matching nodes and their ancestors.
.DESCRIPTION
    Takes the search text from the topology search box, recursively filters all TreeViewItem
    nodes by name match, and expands ancestor nodes of matches for visibility.
#>
function Filter-TopologyTree {
    param([string]$Query)
    $Window.Dispatcher.Invoke([Action]{
        if ([string]::IsNullOrWhiteSpace($Query)) {
            # Show all
            foreach ($node in $tvTopology.Items) {
                $node.Visibility = 'Visible'
                Set-ChildVisibility -Node $node -Visible $true
            }
            return
        }

        $lowerQ = $Query.ToLower()
        foreach ($node in $tvTopology.Items) {
            $match = Apply-FilterRecursive -Node $node -Query $lowerQ
            $node.Visibility = if ($match) { 'Visible' } else { 'Collapsed' }
        }
    })
}

<#
.SYNOPSIS
    Recursively applies filter visibility to TreeViewItem nodes, propagating match state to children.
#>
function Apply-FilterRecursive {
    param($Node, [string]$Query)
    $headerText = if ($Node.Header -is [string]) { $Node.Header } else { $Node.Header.ToString() }
    $selfMatch = $headerText.ToLower().Contains($Query)
    $childMatch = $false

    foreach ($child in $Node.Items) {
        $cm = Apply-FilterRecursive -Node $child -Query $Query
        if ($cm) { $childMatch = $true }
        $child.Visibility = if ($cm -or $selfMatch) { 'Visible' } else { 'Collapsed' }
    }

    if ($selfMatch -or $childMatch) {
        $Node.IsExpanded = $true
    }
    return ($selfMatch -or $childMatch)
}

<#
.SYNOPSIS
    Toggles the Visibility property of all descendant TreeViewItems recursively.
#>
function Set-ChildVisibility {
    param($Node, [bool]$Visible)
    foreach ($child in $Node.Items) {
        $child.Visibility = if ($Visible) { 'Visible' } else { 'Collapsed' }
        Set-ChildVisibility -Node $child -Visible $Visible
    }
}

$cmbTopoResourceGroup.Add_SelectionChanged({
    $sel = $cmbTopoResourceGroup.SelectedItem
    if (-not $sel -or $sel.Content -eq 'All Resource Groups') {
        Filter-TopologyTree -Query ''
        return
    }
    $rgName = $sel.Content
    $Window.Dispatcher.Invoke([Action]{
        foreach ($node in $tvTopology.Items) {
            $tag = $node.Tag
            $match = $false
            if ($tag -and $tag.ResourceGroupName -eq $rgName) { $match = $true }
            $node.Visibility = if ($match) { 'Visible' } else { 'Collapsed' }
        }
    })
})

# ===============================================================================
# SECTION 23: ACHIEVEMENT SYSTEM
# ===============================================================================

$Global:AchievementDefs = @(
    @{ Id='FirstDiscovery';  Icon=[char]::ConvertFromUtf32(0x1F50D); Title='Explorer';       Desc='Ran your first AVD discovery' }
    @{ Id='FirstBackup';     Icon=[char]::ConvertFromUtf32(0x1F4BE); Title='Archivist';      Desc='Saved your first backup' }
    @{ Id='FirstDiff';       Icon=[char]::ConvertFromUtf32(0x1F50E); Title='Detective';      Desc='Ran your first topology comparison' }
    @{ Id='FirstDryRun';     Icon=[char]::ConvertFromUtf32(0x1F9EA); Title='Cautious One';   Desc='Ran your first dry-run restore' }
    @{ Id='FirstRestore';    Icon=[char]::ConvertFromUtf32(0x23EA);  Title='Time Traveler';  Desc='Performed your first restore' }
    @{ Id='PerfectRestore';  Icon=[char]::ConvertFromUtf32(0x2728);  Title='Flawless';       Desc='Restored 5+ items with zero failures' }
    @{ Id='SpeedDemon';      Icon=[char]::ConvertFromUtf32(0x26A1);  Title='Speed Demon';    Desc='Completed a discovery in under 10s' }
    @{ Id='BackupHoarder';   Icon=[char]::ConvertFromUtf32(0x1F4E6); Title='Hoarder';        Desc='Accumulated 10+ backups' }
    @{ Id='NightOwl';        Icon=[char]::ConvertFromUtf32(0x1F989); Title='Night Owl';      Desc='Used AVD Rewind after midnight' }
    @{ Id='ThemeSwitcher';   Icon=[char]::ConvertFromUtf32(0x1F3A8); Title='Indecisive';     Desc='Toggled theme 5+ times in one session' }
    @{ Id='MultiSub';        Icon=[char]::ConvertFromUtf32(0x1F310); Title='Multi-Tenant';   Desc='Switched between 3+ subscriptions' }
    @{ Id='BigEnvironment';  Icon=[char]::ConvertFromUtf32(0x1F40B); Title='Whale';          Desc='Discovered 10+ host pools in one sub' }
    @{ Id='CleanSlate';      Icon=[char]::ConvertFromUtf32(0x1F9F9); Title='Clean Slate';    Desc='Ran retention cleanup successfully' }
    @{ Id='Exporter';        Icon=[char]::ConvertFromUtf32(0x1F4E4); Title='Exporter';       Desc='Exported topology to file' }
    @{ Id='Dedicated';       Icon=[char]::ConvertFromUtf32(0x2B50);  Title='Dedicated';      Desc='Opened AVD Rewind 10+ times' }
)

<#
.SYNOPSIS
    Checks and unlocks a specific achievement if not already unlocked.
.DESCRIPTION
    If the achievement ID is not in the unlocked list, adds it, shows a celebratory toast
    notification, re-renders the achievement badge grid, and saves preferences to persist
    the unlock.
#>
function Check-Achievement {
    param([string]$Key)
    $def = $Global:AchievementDefs | Where-Object { $_.Id -eq $Key }
    if (-not $def) { return }
    if ($Global:Achievements -contains $Key) { return }

    $Global:Achievements += $Key
    Write-DebugLog "[ACHIEVEMENT] Unlocked: $($def.Title) - $($def.Desc)" -Level 'INFO'
    Show-Toast -Message "Achievement unlocked: $($def.Title)!" -Type 'Success'
    $lblAchievementCount.Text = "$($Global:Achievements.Count)/$($Global:AchievementDefs.Count)"
    Render-Achievements
    Save-UserPrefs
}

<#
.SYNOPSIS
    Renders the achievement badge grid in the Settings tab showing locked and unlocked achievements.
.DESCRIPTION
    Iterates all 15 achievement definitions, rendering each as a themed badge with the icon
    (or locked placeholder), title tooltip, and description. Updates the X/15 unlocked counter.
#>
function Render-Achievements {
    $Window.Dispatcher.Invoke([Action]{
        $pnlAchievements.Children.Clear()
        $Unlocked = 0
        foreach ($Def in $Global:AchievementDefs) {
            $IsUnlocked = $Global:Achievements -contains $Def.Id
            if ($IsUnlocked) { $Unlocked++ }

            $Badge = [System.Windows.Controls.Border]::new()
            $Badge.Width        = 32
            $Badge.Height       = 32
            $Badge.CornerRadius = [System.Windows.CornerRadius]::new(6)
            $Badge.Margin       = [System.Windows.Thickness]::new(0, 0, 4, 4)
            $Badge.ToolTip      = if ($IsUnlocked) { "$($Def.Title): $($Def.Desc)" } else { '???' }

            if ($IsUnlocked) {
                $Badge.Background      = $Window.FindResource('ThemeSelectedBg')
                $Badge.BorderBrush     = $Window.FindResource('ThemeBorderElevated')
                $Badge.BorderThickness = [System.Windows.Thickness]::new(1)
                $Lbl = [System.Windows.Controls.TextBlock]::new()
                $Lbl.Text              = $Def.Icon
                $Lbl.FontFamily        = [System.Windows.Media.FontFamily]::new('Segoe UI Emoji')
                $Lbl.FontSize          = 16
                $Lbl.HorizontalAlignment = 'Center'
                $Lbl.VerticalAlignment   = 'Center'
            } else {
                $Badge.Background      = $Window.FindResource('ThemeDeepBg')
                $Badge.BorderBrush     = $Window.FindResource('ThemeBorder')
                $Badge.BorderThickness = [System.Windows.Thickness]::new(1)
                $Lbl = [System.Windows.Controls.TextBlock]::new()
                $Lbl.Text              = '?'
                $Lbl.FontSize          = 14
                $Lbl.Foreground        = $Window.FindResource('ThemeTextDisabled')
                $Lbl.HorizontalAlignment = 'Center'
                $Lbl.VerticalAlignment   = 'Center'
            }
            $Badge.Child = $Lbl
            [void]$pnlAchievements.Children.Add($Badge)
        }
        $lblAchievementCount.Text = "$Unlocked/$($Global:AchievementDefs.Count)"
    })
}

<#
.SYNOPSIS
    Loads persisted achievement state from preferences and renders the badge grid on startup.
#>
function Load-Achievements {
    # Achievements already loaded from prefs in Load-UserPrefs; just render the badges
    Write-DebugLog "[ACHIEVEMENTS] Rendering $($Global:Achievements.Count) unlocked achievement(s)" -Level 'DEBUG'
    Render-Achievements
}
# ===============================================================================
# SECTION 24: DISPATCHER TIMER (Background Job Polling)
# ===============================================================================

$Global:DispatcherTimer = [System.Windows.Threading.DispatcherTimer]::new()
$Global:DispatcherTimer.Interval = [TimeSpan]::FromMilliseconds($TIMER_INTERVAL_MS)
$Global:DispatcherTimer.Add_Tick({
    try {
        # Poll background jobs
        $jobCount = $Global:BgJobs.Count
        for ($i = $jobCount - 1; $i -ge 0; $i--) {
            $job = $Global:BgJobs[$i]
            if (-not $job -or -not $job.Handle) { continue }
            # Stream progress updates to the shimmer label for running jobs
            if (-not $job.Handle.IsCompleted -and $job.PS.Streams.Progress.Count -gt 0) {
                $latestProg = $job.PS.Streams.Progress[$job.PS.Streams.Progress.Count - 1]
                if ($latestProg.StatusDescription) {
                    $lblGlobalProgress.Text = $latestProg.StatusDescription
                }
            }
            if ($job.Handle.IsCompleted) {
                $elapsed = ((Get-Date) - $job.StartTime).TotalSeconds
                try {
                    $result = $job.PS.EndInvoke($job.Handle)
                    if ($job.PS.Streams.Error.Count -gt 0) {
                        $errMsg = $job.PS.Streams.Error[0].Exception.Message
                        Write-DebugLog "BgJob[$i] '$($job.Name)' ERROR after $([math]::Round($elapsed,1))s: $errMsg" -Level 'ERROR'
                        if ($job.OnComplete) {
                            try { & $job.OnComplete $null } catch { Write-DebugLog "BgJob[$i] OnComplete crash: $($_.Exception.Message)" -Level 'ERROR' }
                        }
                    } else {
                        Write-DebugLog "BgJob[$i] '$($job.Name)' COMPLETED after $([math]::Round($elapsed,1))s" -Level 'DEBUG'
                        if ($job.OnComplete) {
                            try {
                                $data = if ($result -and $result.Count -gt 0) { $result[0] } else { $result }
                                & $job.OnComplete $data
                            } catch {
                                Write-DebugLog "BgJob[$i] '$($job.Name)' OnComplete crash: $($_.Exception.Message)`n$($_.ScriptStackTrace)" -Level 'ERROR'
                            }
                        }
                    }
                } catch {
                    Write-DebugLog "BgJob[$i] '$($job.Name)' EndInvoke crash: $($_.Exception.Message)" -Level 'ERROR'
                } finally {
                    try { $job.PS.Dispose() } catch { }
                    try { $job.Runspace.Dispose() } catch { }
                    $Global:BgJobs.RemoveAt($i)
                }
            }
        }

        # Check for time-based achievements
        if ((Get-Date).Hour -ge 0 -and (Get-Date).Hour -lt 5) {
            Check-Achievement 'NightOwl'
        }
    } catch {
        Write-Host "[TIMER CRASH] $($_.Exception.Message)" -ForegroundColor Red
    }
})

# ===============================================================================


# --- Help dialog ---
$btnHelp.Add_Click({
    $HelpText = @"
AVD Rewind - Azure Virtual Desktop Disaster Recovery Tool

PREREQUISITES
  Az.Accounts              Authentication & subscription context
  Az.DesktopVirtualization Host pools, app groups, workspaces
  Az.Resources             RBAC role assignments, resource groups
  Az.Compute               VM metadata, Invoke-AzVMRunCommand
  Az.Network               NIC/subnet/NSG mapping per session host
  Az.Monitor               Diagnostic settings backup/restore

  Install all:
  Install-Module Az.Accounts, Az.DesktopVirtualization, Az.Resources, Az.Compute, Az.Network, Az.Monitor -Force

TABS
  Dashboard   Overview counters and quick actions
  Topology    Interactive tree of your AVD resources
  Backups     Browse and compare backup snapshots
  Restore     Selectively restore resources from backup
  Settings    Theme, auto-backup, retention, prerequisites

WORKFLOW
  1. Sign In to Azure (top bar)
  2. Click Refresh to discover the AVD topology
  3. Click Backup Now to create a snapshot
  4. On Restore tab, select a backup and check items
  5. Use Dry Run to validate, then Execute Restore

KEYBOARD SHORTCUTS
  Ctrl+R      Refresh topology
  Ctrl+B      Backup now
  Ctrl+T      Toggle theme
"@

    Show-ThemedDialog -Title 'AVD Rewind - Help' `
        -Message $HelpText `
        -Buttons @( @{ Text='Close'; IsAccent=$true; Result='OK' } )
})
# SECTION 25: STARTUP SEQUENCE
# Splash screen initialization, prerequisite validation, preference loading, theme
# application, auto-login attempt, initial tab setup, and Window.ShowDialog() launch.
# ===============================================================================

Write-DebugLog "=== AVD Rewind v$Global:AppVersion ===" -Level 'DEBUG'
Write-DebugLog "Startup sequence initiated" -Level 'DEBUG'

# Step 1: Prerequisites
Set-SplashStep -Step 1 -State 'running'
$PrereqResult = Test-Prerequisites
if (-not $PrereqResult.Success) {
    Set-SplashStep -Step 1 -State 'error'
    $Missing = ($PrereqResult.Missing | ForEach-Object { $_.Name }) -join ', '
    $Remediation = Get-RemediationCommand -MissingModules $PrereqResult.Missing
    Show-ThemedDialog -Title 'Missing Prerequisites' `
        -Message "The following modules are required but not installed:`n`n$Missing`n`nRun this command to install them:`n$Remediation" `
        -Buttons @( @{ Text='OK'; IsAccent=$true; Result='OK' } )
    Write-DebugLog "[STARTUP] Missing modules: $Missing" -Level 'ERROR'
} else {
    Set-SplashStep -Step 1 -State 'done'
    Write-DebugLog "[STARTUP] All prerequisites satisfied" -Level 'SUCCESS'
}
Render-PrereqStatus -PrereqResult $PrereqResult

# Step 2: Load preferences & theme
Set-SplashStep -Step 2 -State 'running'
Load-UserPrefs
& $ApplyTheme $Global:IsLightMode
Load-Achievements
Set-SplashStep -Step 2 -State 'done'
Write-DebugLog "[STARTUP] Preferences and theme loaded" -Level 'DEBUG'

# Step 3: Auto-connect
Set-SplashStep -Step 3 -State 'running'
$Global:DispatcherTimer.Start()
Write-DebugLog "[STARTUP] Dispatcher timer started (${TIMER_INTERVAL_MS}ms interval)" -Level 'DEBUG'

# Check for cached Azure context
$CachedCtx = Get-AzContext -ErrorAction SilentlyContinue
if ($CachedCtx -and $CachedCtx.Account) {
    Write-DebugLog "[STARTUP] Found cached Azure context: $($CachedCtx.Account.Id)" -Level 'INFO'
    Set-SplashStep -Step 3 -State 'done'
    Connect-ToAzure
} else {
    Write-DebugLog "[STARTUP] No cached Azure context - manual login required" -Level 'INFO'
    Set-SplashStep -Step 3 -State 'skipped'
}

# Track app opens for achievement
$Global:AppOpenCount++
if ($Global:AppOpenCount -ge 10) { Check-Achievement 'Dedicated' }

# Default tab
Switch-Tab 'Dashboard'
Update-BackupList

# Kick splash dismiss
Start-SplashDismissTimer

# Show window
$Window.ShowDialog() | Out-Null

# Cleanup on exit
$Global:DispatcherTimer.Stop()
Write-DebugLog "=== AVD Rewind shutdown ===" -Level 'INFO'
