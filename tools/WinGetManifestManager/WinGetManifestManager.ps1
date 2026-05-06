<#
.SYNOPSIS
    WinGet Manifest Manager - GUI for creating, editing, and publishing WinGet package manifests.

.DESCRIPTION
    A WPF-based GUI tool that authenticates against Azure using Entra ID,
    manages WinGet multi-file manifests (version + installer + defaultLocale),
    uploads installer binaries to Azure Blob Storage, and publishes manifests
    to a private WinGet REST source via Add-WinGetManifest.

    ARCHITECTURE (mirrors AIBLogMonitor):
    1. GUI Thread (STA): Renders WPF, handles user input, Connect-AzAccount.
    2. Worker Runspaces (Background): Az API calls, SHA256 hashing, blob uploads.
    3. Synchronized Queue: Thread-safe bridge between worker and GUI.
    4. DispatcherTimer (50ms): Drains the queue and updates UI elements.

.NOTES
    Version:        0.1.0-alpha
    Status:         Work in Progress
    Author:         Anton Romanyuk
    Creation Date:  03/10/2026
    Dependencies:   Az.Accounts, Az.Resources, Az.Storage, Microsoft.WinGet.RestSource

.EXAMPLE
    .\WinGetManifestManager.ps1
    # Launches the GUI. Sign in, create manifests, upload packages, publish.
#>

# ==============================================================================
# SECTION 1: PRE-LOAD & INITIALIZATION
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'

$Global:Root = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Global:Root)) { $Global:Root = $PWD.Path }

$Global:AppVersion = "0.1.0-alpha"
$Global:AppTitle = "WinGet Manifest Manager v$($Global:AppVersion)"
$Global:PrefsPath = Join-Path $Global:Root "user_prefs.json"

# ─── Named Constants ──────────────────────────────────────────────────────
$Script:LOG_MAX_LINES       = 500    # Debug log ring buffer capacity
$Script:RBAC_MAX_POLLS      = 20     # Max polling iterations for role propagation
$Script:RBAC_POLL_SEC       = 15     # Seconds between RBAC propagation polls
$Script:STREAK_HOURS        = 24     # Hours within which publishes count as a streak
$Script:CONFETTI_COUNT      = 60     # Number of confetti particles
$Script:CLEANUP_DELAY_MS    = 5000   # Delay before cleaning temp files after publish
$Script:TIMER_INTERVAL_MS   = 50     # Dispatcher timer tick interval
$Script:TOAST_DURATION_MS   = 4000   # Default toast auto-dismiss time

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

# Thread-safe process output helper — pure .NET event handlers that don't need a
# PowerShell runspace, avoiding "no Runspace available" crashes on thread pool threads.
try {
    Add-Type -TypeDefinition @"
using System;
using System.Collections.Concurrent;
using System.Diagnostics;

public class ProcessOutputHelper {
    private ConcurrentQueue<string> _queue;
    public ProcessOutputHelper(ConcurrentQueue<string> queue) { _queue = queue; }
    public void OnOutput(object s, DataReceivedEventArgs e) {
        if (e.Data != null) _queue.Enqueue("OUT:" + e.Data);
    }
    public void OnError(object s, DataReceivedEventArgs e) {
        if (e.Data != null) _queue.Enqueue("ERR:" + e.Data);
    }
    public void OnExited(object s, EventArgs e) {
        var p = s as Process;
        if (p != null) _queue.Enqueue("__DEPLOY_EXIT__:" + p.ExitCode);
    }
}
"@ -ErrorAction SilentlyContinue
} catch { <# Already loaded #> }

# ==============================================================================
# SECTION 2: PREREQUISITE MODULE CHECK (No Auto-Install)
# ==============================================================================

$RequiredModules = @(
    @{ Name = 'Az.Accounts';                MinVersion = '2.7.5';  Purpose = 'Azure authentication (Connect-AzAccount)' }
    @{ Name = 'Az.Resources';               MinVersion = '6.0.0';  Purpose = 'Resource group discovery (Get-AzResourceGroup)' }
    @{ Name = 'Az.Storage';                  MinVersion = '5.0.0';  Purpose = 'Blob storage operations (Set-AzStorageBlobContent)' }
    @{ Name = 'Microsoft.WinGet.RestSource'; MinVersion = '1.10.0'; Purpose = 'Manifest publishing (Add-WinGetManifest)' }
)

function Test-Prerequisites {
    <#
    .SYNOPSIS
        Checks if all required modules are installed. Returns a results array.
    #>
    $Results = @()
    foreach ($Mod in $RequiredModules) {
        $Installed = Get-Module -ListAvailable -Name $Mod.Name -ErrorAction SilentlyContinue |
                     Sort-Object Version -Descending | Select-Object -First 1
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
    StatusQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())
    DownloadProgress = -1
    DownloadBytes    = 0
    DownloadTotal    = 0
    UploadProgress   = -1
    UploadBytes      = 0
    UploadTotal      = 0
    BatchTotal       = 0
    BatchDone        = 0
    BatchCurrent     = ''
})

# Background job tracker — polled by the DispatcherTimer to avoid UI-thread hangs
$Global:BgJobs = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
$Global:TimerProcessing = $false

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
        [hashtable]$Variables = @{},
        [hashtable]$Context = @{}
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

    Write-DebugLog "BgWork: launching runspace (state=$($RS.RunspaceStateInfo.State), vars=$($Variables.Keys -join ','))" -Level 'DEBUG'
    $Async = $PS.BeginInvoke()
    $Global:BgJobs.Add(@{
        PS          = $PS
        Runspace    = $RS
        AsyncResult = $Async
        OnComplete  = $OnComplete
        Context     = $Context
        StartedAt   = [datetime]::Now
    }) | Out-Null
    Write-DebugLog "BgWork: queued job #$($Global:BgJobs.Count)" -Level 'DEBUG'
}

# ==============================================================================
# SECTION 4: XAML GUI LOAD
# ==============================================================================

$XamlPath = Join-Path $Global:Root "WinGetManifestManager_UI.xaml"
if (-not (Test-Path $XamlPath)) {
    [System.Windows.MessageBox]::Show(
        "CRITICAL: 'WinGetManifestManager_UI.xaml' not found in:`n$Global:Root",
        "WinGet Manifest Manager", 'OK', 'Error') | Out-Null
    exit 1
}

$XamlContent = Get-Content $XamlPath -Raw -Encoding UTF8
$XamlContent = $XamlContent -replace 'Title="WinGet Manifest Manager"', "Title=`"$Global:AppTitle`""

# Parse XAML
try {
    $Window = [Windows.Markup.XamlReader]::Parse($XamlContent)
} catch {
    [System.Windows.MessageBox]::Show(
        "XAML Parse Error:`n$($_.Exception.Message)",
        "WinGet Manifest Manager", 'OK', 'Error') | Out-Null
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

# Left panel
$pnlLeftSidebar     = $Window.FindName("pnlLeftSidebar")
$colLeftPanel       = $Window.FindName("colLeftPanel")
$splitterLeft       = $Window.FindName("splitterLeft")
$btnHamburger       = $Window.FindName("btnHamburger")
$railConfigs        = $Window.FindName("railConfigs")
$railStorage        = $Window.FindName("railStorage")
$railSettings       = $Window.FindName("railSettings")
$grpLeftConfigs     = $Window.FindName("grpLeftConfigs")
$grpLeftStorage     = $Window.FindName("grpLeftStorage")
$grpLeftSettings    = $Window.FindName("grpLeftSettings")
$lstConfigs         = $Window.FindName("lstConfigs")
$lblConfigEmpty     = $Window.FindName("lblConfigEmpty")
$lblConfigCount     = $Window.FindName("lblConfigCount")
$btnNewConfig       = $Window.FindName("btnNewConfig")
$btnRefreshConfigs  = $Window.FindName("btnRefreshConfigs")
$btnRefreshStorage  = $Window.FindName("btnRefreshStorage")
$btnNewStorage      = $Window.FindName("btnNewStorage")
$tvStorage          = $Window.FindName("tvStorage")
$lblStorageEmpty    = $Window.FindName("lblStorageEmpty")
$txtRestSourceName  = $Window.FindName("txtRestSourceName")
$cmbRestSource      = $Window.FindName("cmbRestSource")
$txtManifestPath    = $Window.FindName("txtManifestPath")
$btnBrowseManifestPath = $Window.FindName("btnBrowseManifestPath")
$cmbManifestVersion = $Window.FindName("cmbManifestVersion")
$chkDarkMode        = $Window.FindName("chkDarkMode")
$chkDebug           = $Window.FindName("chkDebug")
$txtDefaultContainer   = $Window.FindName("txtDefaultContainer")
$txtDiscoveryPattern   = $Window.FindName("txtDiscoveryPattern")
$txtCacheTTL           = $Window.FindName("txtCacheTTL")
$cmbDefaultSubscription = $Window.FindName("cmbDefaultSubscription")
$chkAutoSave           = $Window.FindName("chkAutoSave")
$chkConfirmUpload      = $Window.FindName("chkConfirmUpload")
$chkDisableAnimations  = $Window.FindName("chkDisableAnimations")
$lblSettingsVersion    = $Window.FindName("lblSettingsVersion")
$lblSettingsCommit     = $Window.FindName("lblSettingsCommit")
$lblSettingsPrefsPath  = $Window.FindName("lblSettingsPrefsPath")

# Center content
$pnlPrereqError    = $Window.FindName("pnlPrereqError")
$pnlModuleStatus   = $Window.FindName("pnlModuleStatus")
$txtPrereqCommand  = $Window.FindName("txtPrereqCommand")
$btnCopyCommand    = $Window.FindName("btnCopyCommand")
$btnRetryPrereq    = $Window.FindName("btnRetryPrereq")
$lblPrereqMessage  = $Window.FindName("lblPrereqMessage")
$pnlMainContent    = $Window.FindName("pnlMainContent")

# Tab buttons
$tabCreate  = $Window.FindName("tabCreate")
$tabEdit    = $Window.FindName("tabEdit")
$tabUpload  = $Window.FindName("tabUpload")
$tabManage  = $Window.FindName("tabManage")
$tabDeploy  = $Window.FindName("tabDeploy")
$tabRemove  = $Window.FindName("tabRemove")
$tabBackup  = $Window.FindName("tabBackup")

# Tab panels
$pnlTabCreate   = $Window.FindName("pnlTabCreate")
$pnlTabEdit     = $Window.FindName("pnlTabEdit")
$pnlTabUpload   = $Window.FindName("pnlTabUpload")
$pnlTabManage   = $Window.FindName("pnlTabManage")
$pnlTabDeploy   = $Window.FindName("pnlTabDeploy")
$pnlTabRemove   = $Window.FindName("pnlTabRemove")
$pnlTabBackup   = $Window.FindName("pnlTabBackup")

# Global Progress Bar
$prgGlobal         = $Window.FindName("prgGlobal")
$pnlGlobalProgress = $Window.FindName("pnlGlobalProgress")
$brdGlobalShimmer  = $Window.FindName("brdGlobalShimmer")
$lblGlobalProgress = $Window.FindName("lblGlobalProgress")

# Create tab fields — Package Identity
$txtPkgId              = $Window.FindName("txtPkgId")
$txtPkgVersion         = $Window.FindName("txtPkgVersion")
$cmbCreateManifestVer  = $Window.FindName("cmbCreateManifestVer")

# Create tab fields — Metadata
$txtPkgName       = $Window.FindName("txtPkgName")
$txtPublisher     = $Window.FindName("txtPublisher")
$txtLicense       = $Window.FindName("txtLicense")
$txtShortDesc     = $Window.FindName("txtShortDesc")
$txtDescription   = $Window.FindName("txtDescription")
$txtAuthor        = $Window.FindName("txtAuthor")
$txtMoniker       = $Window.FindName("txtMoniker")
$txtPackageUrl    = $Window.FindName("txtPackageUrl")
$txtPublisherUrl  = $Window.FindName("txtPublisherUrl")
$txtLicenseUrl    = $Window.FindName("txtLicenseUrl")
$txtSupportUrl    = $Window.FindName("txtSupportUrl")
$txtCopyright     = $Window.FindName("txtCopyright")
$txtPrivacyUrl    = $Window.FindName("txtPrivacyUrl")
$txtTags          = $Window.FindName("txtTags")
$txtReleaseNotes    = $Window.FindName("txtReleaseNotes")
$txtReleaseNotesUrl = $Window.FindName("txtReleaseNotesUrl")

# Create tab fields — Primary installer
$pnlInstallers      = $Window.FindName("pnlInstallers")
$pnlInstaller0      = $Window.FindName("pnlInstaller0")
$cmbInstallerPreset = $Window.FindName("cmbInstallerPreset")
$btnApplyPreset     = $Window.FindName("btnApplyPreset")
$cmbInstallerType0  = $Window.FindName("cmbInstallerType0")
$cmbArch0           = $Window.FindName("cmbArch0")
$cmbScope0          = $Window.FindName("cmbScope0")
$txtLocalFile0      = $Window.FindName("txtLocalFile0")
$btnBrowseFile0     = $Window.FindName("btnBrowseFile0")
$lblHash0           = $Window.FindName("lblHash0")
$txtInstallerUrl0   = $Window.FindName("txtInstallerUrl0")
$txtSilent0         = $Window.FindName("txtSilent0")
$txtSilentProgress0 = $Window.FindName("txtSilentProgress0")
$txtCustomSwitch0   = $Window.FindName("txtCustomSwitch0")
$txtInteractive0    = $Window.FindName("txtInteractive0")
$txtLog0            = $Window.FindName("txtLog0")
$txtRepair0         = $Window.FindName("txtRepair0")
$chkModeInteractive = $Window.FindName("chkModeInteractive")
$chkModeSilent      = $Window.FindName("chkModeSilent")
$chkModeSilentProgress = $Window.FindName("chkModeSilentProgress")
$chkInstallLocationRequired = $Window.FindName("chkInstallLocationRequired")
$chkPlatformDesktop  = $Window.FindName("chkPlatformDesktop")
$chkPlatformUniversal = $Window.FindName("chkPlatformUniversal")
$txtMinOSVersion0   = $Window.FindName("txtMinOSVersion0")
$txtExpectedReturnCodes0 = $Window.FindName("txtExpectedReturnCodes0")
$txtAFDisplayName0  = $Window.FindName("txtAFDisplayName0")
$txtAFPublisher0    = $Window.FindName("txtAFPublisher0")
$txtAFDisplayVersion0 = $Window.FindName("txtAFDisplayVersion0")
$txtAFProductCode0  = $Window.FindName("txtAFProductCode0")
$cmbUpgrade0        = $Window.FindName("cmbUpgrade0")
$cmbElevation0      = $Window.FindName("cmbElevation0")
$txtProductCode0    = $Window.FindName("txtProductCode0")
$txtCommands0       = $Window.FindName("txtCommands0")
$txtFileExt0        = $Window.FindName("txtFileExt0")
$pnlNestedInstaller = $Window.FindName("pnlNestedInstaller")
$cmbNestedType0     = $Window.FindName("cmbNestedType0")
$txtPortableAlias0  = $Window.FindName("txtPortableAlias0")
$txtNestedPath0     = $Window.FindName("txtNestedPath0")
$btnAddInstaller    = $Window.FindName("btnAddInstaller")

# Create tab actions
$btnImportPSADT = $Window.FindName("btnImportPSADT")
$btnSaveToDisk  = $Window.FindName("btnSaveToDisk")
$btnPublish     = $Window.FindName("btnPublish")
$btnSaveConfig  = $Window.FindName("btnSaveConfig")
$txtYamlPreview = $Window.FindName("txtYamlPreview")

# Edit tab
$btnLoadFromDisk   = $Window.FindName("btnLoadFromDisk")
$btnLoadFromConfig = $Window.FindName("btnLoadFromConfig")
$pnlEditFields     = $Window.FindName("pnlEditFields")
$lblEditPkgId      = $Window.FindName("lblEditPkgId")
$lblEditPkgVersion = $Window.FindName("lblEditPkgVersion")
$txtEditYaml       = $Window.FindName("txtEditYaml")
$btnEditSave       = $Window.FindName("btnEditSave")
$btnEditPublish    = $Window.FindName("btnEditPublish")

# Manage tab
$lblManageSourceName = $Window.FindName("lblManageSourceName")
$lblManagePkgCount   = $Window.FindName("lblManagePkgCount")
$btnRefreshManage    = $Window.FindName("btnRefreshManage")
$lblManageEmpty      = $Window.FindName("lblManageEmpty")
$lstManagePackages   = $Window.FindName("lstManagePackages")
$pnlManageDetails    = $Window.FindName("pnlManageDetails")
$lblManagePkgId      = $Window.FindName("lblManagePkgId")
$lblManageVersions   = $Window.FindName("lblManageVersions")
$cmbManageVersion    = $Window.FindName("cmbManageVersion")
$chkDeleteBlob       = $Window.FindName("chkDeleteBlob")
$chkDeleteConfig     = $Window.FindName("chkDeleteConfig")
$btnRemoveVersion    = $Window.FindName("btnRemoveVersion")
$btnRemovePackage    = $Window.FindName("btnRemovePackage")
$prgManage           = $Window.FindName("prgManage")
$lblManageStatus     = $Window.FindName("lblManageStatus")
$txtManageSearch     = $Window.FindName("txtManageSearch")
$btnNewVersion       = $Window.FindName("btnNewVersion")
$btnDiffVersions     = $Window.FindName("btnDiffVersions")

# Toast & Stepper
$pnlToastHost        = $Window.FindName("pnlToastHost")
$pnlStepper          = $Window.FindName("pnlStepper")
$stepIdentity        = $Window.FindName("stepIdentity")
$stepMetadata        = $Window.FindName("stepMetadata")
$stepInstaller       = $Window.FindName("stepInstaller")
$stepReview          = $Window.FindName("stepReview")

# Collapsible sections
$secIdentityHeader   = $Window.FindName("secIdentityHeader")
$secIdentityChevron  = $Window.FindName("secIdentityChevron")
$secIdentityBody     = $Window.FindName("secIdentityBody")
$secMetadataHeader   = $Window.FindName("secMetadataHeader")
$secMetadataChevron  = $Window.FindName("secMetadataChevron")
$secMetadataBody     = $Window.FindName("secMetadataBody")
$secInstallerHeader  = $Window.FindName("secInstallerHeader")
$secInstallerChevron = $Window.FindName("secInstallerChevron")
$secInstallerBody    = $Window.FindName("secInstallerBody")
$secDepsHeader       = $Window.FindName("secDepsHeader")
$secDepsChevron      = $Window.FindName("secDepsChevron")
$secDepsBody         = $Window.FindName("secDepsBody")
$secUninstallHeader  = $Window.FindName("secUninstallHeader")
$secUninstallChevron = $Window.FindName("secUninstallChevron")
$secUninstallBody    = $Window.FindName("secUninstallBody")

# Dependencies & Uninstall fields
$txtPackageDeps      = $Window.FindName("txtPackageDeps")
$txtWindowsFeatures  = $Window.FindName("txtWindowsFeatures")
$txtWindowsLibs      = $Window.FindName("txtWindowsLibs")
$txtExternalDeps     = $Window.FindName("txtExternalDeps")
$txtUninstallCmd     = $Window.FindName("txtUninstallCmd")
$txtUninstallSilent  = $Window.FindName("txtUninstallSilent")

# Drag-drop zones & URL test
$dropZoneInstaller0  = $Window.FindName("dropZoneInstaller0")
$dropZoneUpload      = $Window.FindName("dropZoneUpload")
$btnTestUrl0         = $Window.FindName("btnTestUrl0")

# Confetti, quality, streak, achievements
$cnvConfetti         = $Window.FindName("cnvConfetti")
$bdrDotGrid          = $Window.FindName("bdrDotGrid")
$bdrGradientGlow     = $Window.FindName("bdrGradientGlow")
$barQuality          = $Window.FindName("barQuality")
$lblQualityPct       = $Window.FindName("lblQualityPct")
$lblStreak           = $Window.FindName("lblStreak")
$pnlAchievements    = $Window.FindName("pnlAchievements")
$lblAchievementCount = $Window.FindName("lblAchievementCount")
$btnToggleAchievements = $Window.FindName("btnToggleAchievements")
$txtAchievementChevron = $Window.FindName("txtAchievementChevron")

# Upload tab
$txtUploadFile     = $Window.FindName("txtUploadFile")
$btnBrowseUpload   = $Window.FindName("btnBrowseUpload")
$lblUploadHash     = $Window.FindName("lblUploadHash")
$lblUploadSize     = $Window.FindName("lblUploadSize")
$cmbUploadStorage  = $Window.FindName("cmbUploadStorage")
$cmbUploadContainer = $Window.FindName("cmbUploadContainer")
$txtBlobPath       = $Window.FindName("txtBlobPath")
$prgUpload         = $Window.FindName("prgUpload")
$lblUploadProgress = $Window.FindName("lblUploadProgress")
$btnUpload         = $Window.FindName("btnUpload")
$btnCopyHash       = $Window.FindName("btnCopyHash")
$btnCopyYaml       = $Window.FindName("btnCopyYaml")

# Installer reorder/edit buttons
$btnEditInstaller     = $Window.FindName("btnEditInstaller")
$btnMoveInstallerUp   = $Window.FindName("btnMoveInstallerUp")
$btnMoveInstallerDown = $Window.FindName("btnMoveInstallerDown")

# Blob browser
$lstBlobBrowser    = $Window.FindName("lstBlobBrowser")
$lblBlobCount      = $Window.FindName("lblBlobCount")
$btnRefreshBlobs   = $Window.FindName("btnRefreshBlobs")
$btnDeleteBlob     = $Window.FindName("btnDeleteBlob")
$pnlBlobBrowser    = $Window.FindName("pnlBlobBrowser")

# Batch installer operations
$pnlBatchOps       = $Window.FindName("pnlBatchOps")
$lblBatchCount     = $Window.FindName("lblBatchCount")
$lstBatchInstallers = $Window.FindName("lstBatchInstallers")
$prgBatch          = $Window.FindName("prgBatch")
$lblBatchProgress  = $Window.FindName("lblBatchProgress")
$btnDownloadAll    = $Window.FindName("btnDownloadAll")
$btnUploadAll      = $Window.FindName("btnUploadAll")
$btnCancelBatch    = $Window.FindName("btnCancelBatch")
$lblMultiInstallerHint = $Window.FindName("lblMultiInstallerHint")
$Global:BatchCancelled = $false
$Global:BatchDownloadedFiles = @{}

# Bottom log panel
$pnlBottomPanel      = $Window.FindName("pnlBottomPanel")
$splitterBottom      = $Window.FindName("splitterBottom")
$rowBottomPanel      = $Window.FindName("rowBottomPanel")
$btnClearLog         = $Window.FindName("btnClearLog")
$btnToggleBottomSize = $Window.FindName("btnToggleBottomSize")
$icoToggleBottomSize = $Window.FindName("icoToggleBottomSize")
$btnHideBottom       = $Window.FindName("btnHideBottom")
$logScroller         = $Window.FindName("logScroller")
$rtbActivityLog      = $Window.FindName("rtbActivityLog")
$docActivityLog      = $Window.FindName("docActivityLog")
$paraLog             = $Window.FindName("paraLog")
$lblStorageInfo      = $Window.FindName("lblStorageInfo")
$pnlStorageInfoRows  = $Window.FindName("pnlStorageInfoRows")

# Status bar
$statusDot      = $Window.FindName("statusDot")
$lblStatus      = $Window.FindName("lblStatus")
$lblPackageCount = $Window.FindName("lblPackageCount")
$lblRegion      = $Window.FindName("lblRegion")
$lblVersion     = $Window.FindName("lblVersion")

$lblVersion.Text = "v$($Global:AppVersion)"
$lblTitleVersion.Text = "v$($Global:AppVersion)"
$lblSettingsVersion.Text = "v$($Global:AppVersion)"
$lblSettingsPrefsPath.Text = $Global:PrefsPath

# Resolve build info for About section (build.json first, git fallback for dev)
$BuildJsonPath = Join-Path $Global:Root 'build.json'
$CommitHash = $null
if (Test-Path $BuildJsonPath) {
    try {
        $BuildInfo = Get-Content $BuildJsonPath -Raw | ConvertFrom-Json
        $CommitHash = $BuildInfo.commit
    } catch { <# build.json parse error — fall back to git #> }
}
if (-not $CommitHash) {
    try { $CommitHash = & git -C $Global:Root rev-parse --short HEAD 2>$null } catch { }
}
if ($CommitHash) { $lblSettingsCommit.Text = "commit $CommitHash" }

# ==============================================================================
# SECTION 5B: DEBUG LOG FUNCTION
# ==============================================================================

$Global:DebugLineCount = 0
$Global:DebugMaxLines  = $Script:LOG_MAX_LINES
$Global:DebugSB        = [System.Text.StringBuilder]::new(4096)
$Global:FullLogSB      = [System.Text.StringBuilder]::new(8192)
$Global:FullLogLines   = 0

# ── File-based debug log (survives crashes) ──────────────────────────────────
$Global:DebugLogFile = Join-Path $env:TEMP 'WinGetMM_debug.log'
try {
    # Rotate if over 2 MB — keep one .prev backup
    if ((Test-Path $Global:DebugLogFile) -and (Get-Item $Global:DebugLogFile).Length -gt 2MB) {
        $PrevLog = $Global:DebugLogFile + '.prev'
        if (Test-Path $PrevLog) { Remove-Item $PrevLog -Force -ErrorAction SilentlyContinue }
        Rename-Item $Global:DebugLogFile $PrevLog -Force -ErrorAction SilentlyContinue
    }
    # Session header
    $Header = "`n=== WinGet Manifest Manager v$Global:AppVersion — $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
    [System.IO.File]::AppendAllText($Global:DebugLogFile, $Header + "`r`n")
} catch { <# best-effort — don't block startup #> }

function Write-DebugLog {
    <#
    .SYNOPSIS
        Writes a timestamped line to the PS console and the activity log panel.
        Level INFO/WARN/ERROR always appear in the log.
        Level DEBUG only appears when Debug Overlay is enabled.
        Output is color-coded by level in the bottom panel RichTextBox.
    #>
    param([string]$Message, [string]$Level = 'INFO')

    $ts   = (Get-Date).ToString('HH:mm:ss.fff')
    $line = "[$ts] [$Level] $Message"

    # Always write to PS console — visible even when GUI hangs
    Write-Host $line -ForegroundColor DarkGray

    # Append to disk log (survives crashes)
    try { [System.IO.File]::AppendAllText($Global:DebugLogFile, $line + "`r`n") } catch { <# best-effort #> }

    # Full log buffer (ring buffer — keeps everything for debug replay)
    $Global:FullLogLines++
    if ($Global:FullLogLines -gt ($Global:DebugMaxLines * 2)) {
        $text = $Global:FullLogSB.ToString()
        $nl   = $text.IndexOf("`n")
        if ($nl -ge 0) {
            $Global:FullLogSB.Clear()
            $Global:FullLogSB.Append($text.Substring($nl + 1)) | Out-Null
            $Global:FullLogLines--
        }
    }
    $Global:FullLogSB.AppendLine($line) | Out-Null

    # Skip DEBUG-level messages from visible log unless overlay is enabled
    if ($Level -eq 'DEBUG' -and -not $Global:DebugOverlayEnabled) { return }

    # Append color-coded Run to RichTextBox paragraph
    if ($paraLog) {
        $Color = if ($Global:IsLightMode) {
            switch ($Level) {
                'ERROR'   { '#CC0000' }
                'WARN'    { '#B86E00' }
                'SUCCESS' { '#008A2E' }
                'DEBUG'   { '#8888AA' }
                default   { '#444444' }
            }
        } else {
            switch ($Level) {
                'ERROR'   { '#FF4040' }
                'WARN'    { '#FF9100' }
                'SUCCESS' { '#16C60C' }
                'DEBUG'   { '#B8860B' }
                default   { '#888888' }
            }
        }
        if ($paraLog.Inlines.Count -gt 0) {
            $paraLog.Inlines.Add((New-Object System.Windows.Documents.LineBreak))
        }
        $Run = New-Object System.Windows.Documents.Run($line)
        $Run.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Color)
        $paraLog.Inlines.Add($Run)

        # Collect minimap/heatmap data for ERROR/WARN/SUCCESS
        $Global:MinimapLineTotal = $Global:DebugLineCount + 1
        if ($Level -eq 'ERROR' -or $Level -eq 'WARN' -or $Level -eq 'SUCCESS') {
            $Global:MinimapDots.Add([PSCustomObject]@{ Color = $Color; Line = $Global:DebugLineCount })
            $Global:HeatmapData.Add([PSCustomObject]@{ Color = $Color; Line = $Global:DebugLineCount })
        }

        # Trim old lines from ring buffer
        $Global:DebugLineCount++
        if ($Global:DebugLineCount -gt $Global:DebugMaxLines) {
            # Remove first Run + its LineBreak
            if ($paraLog.Inlines.Count -ge 2) {
                $paraLog.Inlines.Remove($paraLog.Inlines.FirstInline)
                if ($paraLog.Inlines.Count -gt 0) {
                    $paraLog.Inlines.Remove($paraLog.Inlines.FirstInline)
                }
                $Global:DebugLineCount--
            }
        }
        $logScroller.ScrollToEnd()
    }
}

# ── Minimap + Heatmap globals ─────────────────────────────────────────────────
$Global:MinimapDots              = [System.Collections.Generic.List[PSObject]]::new()
$Global:MinimapLineTotal         = 0
$Global:MinimapLastRenderedIdx   = 0
$Global:HeatmapData              = [System.Collections.Generic.List[PSObject]]::new()
$Global:HeatmapLastRenderedIdx   = 0

function Render-Minimap {
    <# Incrementally renders color-coded dots on the vertical minimap canvas #>
    $cnv = $Window.FindName('cnvMinimap')
    if (-not $cnv -or $cnv.ActualHeight -le 0) { return }
    $H = $cnv.ActualHeight
    $Total = [Math]::Max($Global:MinimapLineTotal, 1)
    # Render new dots since last call
    for ($i = $Global:MinimapLastRenderedIdx; $i -lt $Global:MinimapDots.Count; $i++) {
        $Dot = $Global:MinimapDots[$i]
        $Y = ($Dot.Line / $Total) * $H
        $Rect = New-Object System.Windows.Shapes.Rectangle
        $Rect.Width = 10; $Rect.Height = 2
        $Rect.Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Dot.Color)
        [System.Windows.Controls.Canvas]::SetLeft($Rect, 2)
        [System.Windows.Controls.Canvas]::SetTop($Rect, $Y)
        $cnv.Children.Add($Rect) | Out-Null
    }
    $Global:MinimapLastRenderedIdx = $Global:MinimapDots.Count
    # Viewport indicator
    $Tag = 'minimapViewport'
    $Old = $cnv.Children | Where-Object { $_.Tag -eq $Tag }
    if ($Old) { $cnv.Children.Remove($Old) }
    if ($logScroller.ExtentHeight -gt 0) {
        $VpTop = ($logScroller.VerticalOffset / $logScroller.ExtentHeight) * $H
        $VpH   = [Math]::Max(($logScroller.ViewportHeight / $logScroller.ExtentHeight) * $H, 4)
        $Vp = New-Object System.Windows.Shapes.Rectangle
        $Vp.Width = 14; $Vp.Height = $VpH
        $Vp.Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#30FFFFFF')
        $Vp.Tag = $Tag
        [System.Windows.Controls.Canvas]::SetLeft($Vp, 0)
        [System.Windows.Controls.Canvas]::SetTop($Vp, $VpTop)
        $cnv.Children.Add($Vp) | Out-Null
    }
}

function Render-Heatmap {
    <# Incrementally renders color marks on the horizontal heatmap timeline canvas #>
    $cnv = $Window.FindName('cnvHeatmap')
    if (-not $cnv -or $cnv.ActualWidth -le 0) { return }
    $W = $cnv.ActualWidth
    $Total = [Math]::Max($Global:MinimapLineTotal, 1)
    for ($i = $Global:HeatmapLastRenderedIdx; $i -lt $Global:HeatmapData.Count; $i++) {
        $Dot = $Global:HeatmapData[$i]
        $X = ($Dot.Line / $Total) * $W
        $Rect = New-Object System.Windows.Shapes.Rectangle
        $Rect.Width = 2; $Rect.Height = 10
        $Rect.Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Dot.Color)
        [System.Windows.Controls.Canvas]::SetLeft($Rect, $X)
        [System.Windows.Controls.Canvas]::SetTop($Rect, 2)
        $cnv.Children.Add($Rect) | Out-Null
    }
    $Global:HeatmapLastRenderedIdx = $Global:HeatmapData.Count
    # Viewport indicator
    $Tag = 'heatmapViewport'
    $Old = $cnv.Children | Where-Object { $_.Tag -eq $Tag }
    if ($Old) { $cnv.Children.Remove($Old) }
    if ($logScroller.ExtentHeight -gt 0) {
        $VpLeft = ($logScroller.VerticalOffset / $logScroller.ExtentHeight) * $W
        $VpW   = [Math]::Max(($logScroller.ViewportHeight / $logScroller.ExtentHeight) * $W, 4)
        $Vp = New-Object System.Windows.Shapes.Rectangle
        $Vp.Width = $VpW; $Vp.Height = 14
        $Vp.Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#30FFFFFF')
        $Vp.Tag = $Tag
        [System.Windows.Controls.Canvas]::SetLeft($Vp, $VpLeft)
        [System.Windows.Controls.Canvas]::SetTop($Vp, 0)
        $cnv.Children.Add($Vp) | Out-Null
    }
}

# ==============================================================================
# SECTION 6: THEME ENGINE (Microsoft Color Scheme)
# ==============================================================================

$Global:IsLightMode = $false
$Global:AnimationsDisabled = $false
$Global:SuppressThemeHandler = $false
$Global:IsAuthenticated = $false
$Global:SuppressSubChangeHandler = $false

$Global:ThemeDark = @{
    # Bagel Commander design language — luminance-layered surfaces, zero-border cards
    ThemeAppBg         = "#111113"; ThemePanelBg       = "#18181B"
    ThemeCardBg        = "#1E1E1E"; ThemeCardAltBg     = "#1A1A1A"
    ThemeInputBg       = "#141414"; ThemeDeepBg        = "#0D0D0D"
    ThemeOutputBg      = "#0A0A0A"; ThemeSurfaceBg     = "#1F1F23"
    ThemeHoverBg       = "#27272B"; ThemeSelectedBg    = "#2A2A2A"
    ThemePressedBg     = "#1A1A1A"
    ThemeAccent        = "#0078D4"; ThemeAccentHover    = "#1A8AD4"
    ThemeAccentLight   = "#60CDFF"; ThemeAccentDim     = "#1A0078D4"
    ThemeGreenAccent   = "#00C853"
    ThemeTextPrimary   = "#FFFFFF"; ThemeTextBody      = "#E0E0E0"
    ThemeTextSecondary = "#A1A1AA"; ThemeTextMuted     = "#71717A"
    ThemeTextDim       = "#8B8B93"; ThemeTextDisabled  = "#383838"
    ThemeTextFaintest  = "#6B6B73"
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
    ThemeGreenAccent   = "#00873D"
    ThemeTextPrimary   = "#111111"; ThemeTextBody      = "#222222"
    ThemeTextSecondary = "#555555"; ThemeTextMuted     = "#666666"
    ThemeTextDim       = "#767676"; ThemeTextDisabled  = "#CCCCCC"
    ThemeTextFaintest  = "#808080"
    ThemeBorder        = "#0A000000"; ThemeBorderCard  = "#0A000000"
    ThemeBorderElevated = "#0A000000"; ThemeBorderHover = "#BBBBBB"
    ThemeBorderMedium  = "#1A000000"
    ThemeBorderSubtle  = "#D0D0D0"
    ThemeErrorDim      = "#20DC2626"
    ThemeProgressEdge  = "#18000000"
    ThemeScrollThumb   = "#C0C0C0"
    ThemeScrollTrack   = "#0A000000"
    ThemeSuccess       = "#16A34A"; ThemeWarning       = "#EA580C"
    ThemeError         = "#DC2626"; ThemeSidebarBg     = "#EBEBEB"
    ThemeSidebarBorder = "#00000000"
}

$ApplyTheme = {
    param([bool]$IsLight)
    $Palette = if ($IsLight) { $Global:ThemeLight } else { $Global:ThemeDark }

    foreach ($Key in $Palette.Keys) {
        $NewColor = [System.Windows.Media.ColorConverter]::ConvertFromString($Palette[$Key])
        $NewBrush = [System.Windows.Media.SolidColorBrush]::new($NewColor)
        $Window.Resources[$Key] = $NewBrush
    }
    $Global:IsLightMode = $IsLight

    # Repaint active tab with new theme colors
    if ($Global:ActiveTabName) { Switch-Tab $Global:ActiveTabName }

    # Update dot grid pattern for theme (DrawingBrush can't use DynamicResource)
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
        # Rebuild gradient with theme-appropriate accent colors
        $Grad = [System.Windows.Media.LinearGradientBrush]::new()
        $Grad.StartPoint = [System.Windows.Point]::new(0.5, 0)
        $Grad.EndPoint   = [System.Windows.Point]::new(0.5, 1)
        if ($IsLight) {
            $Grad.GradientStops.Add([System.Windows.Media.GradientStop]::new(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#180078D4'), 0.0))
            $Grad.GradientStops.Add([System.Windows.Media.GradientStop]::new(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#0C0078D4'), 0.35))
            $Grad.GradientStops.Add([System.Windows.Media.GradientStop]::new(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#060078D4'), 0.6))
            $Grad.GradientStops.Add([System.Windows.Media.GradientStop]::new(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#00FFFFFF'), 1.0))
        } else {
            $Grad.GradientStops.Add([System.Windows.Media.GradientStop]::new(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#200078D4'), 0.0))
            $Grad.GradientStops.Add([System.Windows.Media.GradientStop]::new(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#100078D4'), 0.35))
            $Grad.GradientStops.Add([System.Windows.Media.GradientStop]::new(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#080078D4'), 0.6))
            $Grad.GradientStops.Add([System.Windows.Media.GradientStop]::new(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#00000000'), 1.0))
        }
        $bdrGradientGlow.Background = $Grad
        $bdrGradientGlow.Opacity = 1.0
    }
}

function Show-ThemedDialog {
    <#
    .SYNOPSIS
        Shows a themed modal dialog matching the app's dark/light mode.
    .PARAMETER Title
        Dialog window title.
    .PARAMETER Message
        Body text for the dialog.
    .PARAMETER Icon
        Segoe Fluent Icons character (default: question mark E897).
    .PARAMETER IconColor
        Theme resource key for icon color (default: ThemeAccentLight).
    .PARAMETER Buttons
        Array of hashtables: @{ Text='Yes'; IsAccent=$true; Result='Yes' }
        If omitted, defaults to a single 'OK' button.
    .OUTPUTS
        The Result string of the clicked button, or 'Cancel' if closed.
    #>
    param(
        [string]$Title   = 'Confirm',
        [string]$Message = '',
        [string]$Icon    = [string]([char]0xE897),
        [string]$IconColor = 'ThemeAccentLight',
        [array]$Buttons  = @( @{ Text='OK'; IsAccent=$true; Result='OK' } ),
        [string]$InputPrompt = '',
        [string]$InputMatch  = ''
    )

    $Palette = if ($Global:IsLightMode) { $Global:ThemeLight } else { $Global:ThemeDark }
    $Br = { param([string]$Key)
        [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.ColorConverter]::ConvertFromString($Palette[$Key])
        )
    }

    $Dlg = New-Object System.Windows.Window
    $Dlg.Title = $Title
    $Dlg.SizeToContent = 'WidthAndHeight'
    $Dlg.MinWidth = 380
    $Dlg.MaxWidth = 520
    $Dlg.ResizeMode = 'NoResize'
    $Dlg.WindowStartupLocation = 'CenterOwner'
    $Dlg.Owner = $Window
    $Dlg.Background = (& $Br 'ThemeCardBg')
    $Dlg.Foreground = (& $Br 'ThemeTextBody')
    $Dlg.FontFamily = [System.Windows.Media.FontFamily]::new('Inter, Segoe UI Variable Display, Segoe UI')
    $Dlg.WindowStyle = 'None'
    $Dlg.AllowsTransparency = $true

    # Main border with rounded corners and shadow
    $OuterBorder = New-Object System.Windows.Controls.Border
    $OuterBorder.Background = (& $Br 'ThemeCardBg')
    $OuterBorder.BorderBrush = (& $Br 'ThemeBorderElevated')
    $OuterBorder.BorderThickness = [System.Windows.Thickness]::new(1)
    $OuterBorder.CornerRadius = [System.Windows.CornerRadius]::new(12)
    $OuterBorder.Padding = [System.Windows.Thickness]::new(28, 24, 28, 20)
    $OuterBorder.Effect = [System.Windows.Media.Effects.DropShadowEffect]@{
        Color = [System.Windows.Media.Colors]::Black
        Direction = 270; ShadowDepth = 8; BlurRadius = 28; Opacity = 0.45
    }

    $Stack = New-Object System.Windows.Controls.StackPanel

    # ── Header row: icon badge + title ──
    $HeaderGrid = New-Object System.Windows.Controls.Grid
    $HeaderGrid.Margin = [System.Windows.Thickness]::new(0, 0, 0, 16)
    $Col1 = New-Object System.Windows.Controls.ColumnDefinition
    $Col1.Width = [System.Windows.GridLength]::new(0, [System.Windows.GridUnitType]::Auto)
    $Col2 = New-Object System.Windows.Controls.ColumnDefinition
    $Col2.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $HeaderGrid.ColumnDefinitions.Add($Col1) | Out-Null
    $HeaderGrid.ColumnDefinitions.Add($Col2) | Out-Null

    $IconBadge = New-Object System.Windows.Controls.Border
    $IconBadge.Width = 36; $IconBadge.Height = 36
    $IconBadge.CornerRadius = [System.Windows.CornerRadius]::new(8)
    $IconBadge.Background = (& $Br 'ThemeAccentDim')
    $IconBadge.Margin = [System.Windows.Thickness]::new(0, 0, 14, 0)
    $IconTB = New-Object System.Windows.Controls.TextBlock
    $IconTB.Text = $Icon
    $IconTB.FontFamily = [System.Windows.Media.FontFamily]::new('Segoe Fluent Icons, Segoe MDL2 Assets')
    $IconTB.FontSize = 17
    $IconTB.Foreground = (& $Br $IconColor)
    $IconTB.HorizontalAlignment = 'Center'
    $IconTB.VerticalAlignment = 'Center'
    $IconBadge.Child = $IconTB
    [System.Windows.Controls.Grid]::SetColumn($IconBadge, 0)
    $HeaderGrid.Children.Add($IconBadge) | Out-Null

    $TitleTB = New-Object System.Windows.Controls.TextBlock
    $TitleTB.Text = $Title
    $TitleTB.FontSize = 15
    $TitleTB.FontWeight = 'Bold'
    $TitleTB.Foreground = (& $Br 'ThemeTextPrimary')
    $TitleTB.VerticalAlignment = 'Center'
    [System.Windows.Controls.Grid]::SetColumn($TitleTB, 1)
    $HeaderGrid.Children.Add($TitleTB) | Out-Null

    $Stack.Children.Add($HeaderGrid) | Out-Null

    # ── Separator ──
    $Sep = New-Object System.Windows.Controls.Border
    $Sep.Height = 1
    $Sep.Background = (& $Br 'ThemeBorder')
    $Sep.Margin = [System.Windows.Thickness]::new(0, 0, 0, 18)
    $Stack.Children.Add($Sep) | Out-Null

    # ── Message body ──
    $MsgTB = New-Object System.Windows.Controls.TextBlock
    $MsgTB.Text = $Message
    $MsgTB.FontSize = 12
    $MsgTB.Foreground = (& $Br 'ThemeTextSecondary')
    $MsgTB.TextWrapping = 'Wrap'
    $MsgTB.Margin = [System.Windows.Thickness]::new(0, 0, 0, $(if ($InputPrompt) { '14' } else { '24' }))
    $MsgTB.LineHeight = 20
    $Stack.Children.Add($MsgTB) | Out-Null

    # ── Optional input field (for type-to-confirm) ──
    $InputBox = $null
    $DangerBtn = $null
    if ($InputPrompt) {
        $InputLabel = New-Object System.Windows.Controls.TextBlock
        $InputLabel.Text = $InputPrompt
        $InputLabel.FontSize = 11
        $InputLabel.Foreground = (& $Br 'ThemeTextSecondary')
        $InputLabel.Margin = [System.Windows.Thickness]::new(0, 0, 0, 6)
        $Stack.Children.Add($InputLabel) | Out-Null

        $InputBox = New-Object System.Windows.Controls.TextBox
        $InputBox.FontSize = 13
        $InputBox.FontFamily = [System.Windows.Media.FontFamily]::new('Cascadia Code, Consolas, Segoe UI')
        $InputBox.Padding = [System.Windows.Thickness]::new(10, 8, 10, 8)
        $InputBox.Margin = [System.Windows.Thickness]::new(0, 0, 0, 20)
        $InputBox.Background = (& $Br 'ThemeInputBg')
        $InputBox.Foreground = (& $Br 'ThemeTextPrimary')
        $InputBox.BorderBrush = (& $Br 'ThemeBorder')
        $InputBox.BorderThickness = [System.Windows.Thickness]::new(1)
        $InputBox.CaretBrush = (& $Br 'ThemeTextPrimary')

        # Round the border via ControlTemplate
        $InputTemplateXaml = @"
<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                 xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                 TargetType="TextBox">
    <Border Background="{TemplateBinding Background}"
            BorderBrush="{TemplateBinding BorderBrush}"
            BorderThickness="{TemplateBinding BorderThickness}"
            CornerRadius="6">
        <ScrollViewer x:Name="PART_ContentHost" Margin="0"/>
    </Border>
</ControlTemplate>
"@
        $InputBox.Template = [System.Windows.Markup.XamlReader]::Parse($InputTemplateXaml)
        $Stack.Children.Add($InputBox) | Out-Null
    }

    # ── Button row ──
    $BtnPanel = New-Object System.Windows.Controls.StackPanel
    $BtnPanel.Orientation = 'Horizontal'
    $BtnPanel.HorizontalAlignment = 'Right'

    $Dlg.Tag = 'Cancel'   # default result if user closes via Alt+F4, etc.

    foreach ($BDef in $Buttons) {
        $Btn = New-Object System.Windows.Controls.Button
        $Btn.Margin = [System.Windows.Thickness]::new(8, 0, 0, 0)
        $Btn.Cursor = [System.Windows.Input.Cursors]::Hand
        $Btn.Tag = $BDef.Result
        $Btn.Focusable = $true

        # Build visual content as a Border + TextBlock
        $BtnBorder = New-Object System.Windows.Controls.Border
        $BtnBorder.CornerRadius = [System.Windows.CornerRadius]::new(8)
        $BtnBorder.Padding = [System.Windows.Thickness]::new(20, 8, 20, 8)
        $BtnLabel = New-Object System.Windows.Controls.TextBlock
        $BtnLabel.Text = $BDef.Text
        $BtnLabel.HorizontalAlignment = 'Center'
        $BtnLabel.VerticalAlignment = 'Center'
        $BtnLabel.FontSize = 12
        $BtnLabel.FontWeight = [System.Windows.FontWeights]::SemiBold

        if ($BDef.IsAccent) {
            $NormalBg = (& $Br 'ThemeAccent')
            $HoverBg  = (& $Br 'ThemeAccentLight')
            $BtnBorder.Background = $NormalBg
            $BtnLabel.Foreground = [System.Windows.Media.Brushes]::White
        } elseif ($BDef.IsDanger) {
            $NormalBg = [System.Windows.Media.Brushes]::Transparent
            $HoverBg  = [System.Windows.Media.SolidColorBrush]::new(
                [System.Windows.Media.Color]::FromArgb(30, 255, 80, 80))
            $BtnBorder.Background = $NormalBg
            $BtnBorder.BorderBrush = (& $Br 'ThemeBorderCard')
            $BtnBorder.BorderThickness = [System.Windows.Thickness]::new(1)
            $BtnLabel.Foreground = (& $Br 'ThemeError')
            # Track the danger button so we can disable it when InputMatch is set
            $DangerBtn = $Btn
        } else {
            $NormalBg = [System.Windows.Media.Brushes]::Transparent
            $HoverBg  = [System.Windows.Media.SolidColorBrush]::new(
                [System.Windows.Media.Color]::FromArgb(25, 255, 255, 255))
            $BtnBorder.Background = $NormalBg
            $BtnBorder.BorderBrush = (& $Br 'ThemeBorderCard')
            $BtnBorder.BorderThickness = [System.Windows.Thickness]::new(1)
            $BtnLabel.Foreground = (& $Br 'ThemeTextSecondary')
        }

        # Store normal/hover brushes on Tag for the hover handlers
        $BtnBorder.Tag = @{ Normal = $NormalBg; Hover = $HoverBg }
        $BtnBorder.Child = $BtnLabel

        # Hover effects — lighten accent, subtle fill for outline buttons
        $BtnBorder.Add_MouseEnter({
            $this.Background = $this.Tag.Hover
        })
        $BtnBorder.Add_MouseLeave({
            $this.Background = $this.Tag.Normal
        })

        # Replace default Button chrome with a flat template that still receives clicks
        $TemplateXaml = @"
<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                 TargetType="Button">
    <Border Background="Transparent">
        <ContentPresenter/>
    </Border>
</ControlTemplate>
"@
        $Btn.Template = [System.Windows.Markup.XamlReader]::Parse($TemplateXaml)
        $Btn.Content = $BtnBorder

        # No .GetNewClosure() — navigate from sender to parent Window instead
        $Btn.Add_Click({
            $SenderBtn = $this
            $ParentWin = [System.Windows.Window]::GetWindow($SenderBtn)
            $ParentWin.Tag = $SenderBtn.Tag
            $ParentWin.Close()
        })

        $BtnPanel.Children.Add($Btn) | Out-Null
    }

    $Stack.Children.Add($BtnPanel) | Out-Null
    $OuterBorder.Child = $Stack
    $Dlg.Content = $OuterBorder

    # ── Wire up InputMatch live validation ──
    if ($InputMatch -and $DangerBtn -and $InputBox) {
        $DangerBtn.IsEnabled = $false
        $DangerBtn.Opacity = 0.4
        $InputBox.Tag = @{ MatchText = $InputMatch; TargetBtn = $DangerBtn }
        $InputBox.Add_TextChanged({
            $Ctx = $this.Tag
            $Match = $this.Text.Trim() -eq $Ctx.MatchText
            $Ctx.TargetBtn.IsEnabled = $Match
            $Ctx.TargetBtn.Opacity = if ($Match) { 1.0 } else { 0.4 }
        })
    }

    # Allow dragging the dialog
    $Dlg.Add_MouseLeftButtonDown({ param($s,$e); if ($e.ChangedButton -eq 'Left') { $s.DragMove() } })

    $Dlg.ShowDialog() | Out-Null
    return $Dlg.Tag
}

# ==============================================================================
# SECTION 7A: MANIFEST CREATOR - YAML GENERATION
# ==============================================================================

function Get-ComboBoxSelectedText {
    <# Gets the selected text from a ComboBox (handles ComboBoxItem content) #>
    param($ComboBox)
    if ($null -eq $ComboBox -or $null -eq $ComboBox.SelectedItem) { return "" }
    if ($ComboBox.SelectedItem -is [System.Windows.Controls.ComboBoxItem]) {
        return $ComboBox.SelectedItem.Content.ToString()
    }
    return $ComboBox.SelectedItem.ToString()
}

function Set-ComboBoxByContent {
    <# Selects a ComboBox item by matching its content text #>
    param($ComboBox, [string]$Value)
    if (-not $Value) { return }
    for ($i = 0; $i -lt $ComboBox.Items.Count; $i++) {
        $ItemContent = if ($ComboBox.Items[$i] -is [System.Windows.Controls.ComboBoxItem]) {
            $ComboBox.Items[$i].Content.ToString()
        } else { $ComboBox.Items[$i].ToString() }
        if ($ItemContent -eq $Value) { $ComboBox.SelectedIndex = $i; return }
    }
}

# Cached BrushConverter instance — avoids repeated allocations across the UI
$Global:CachedBC = [System.Windows.Media.BrushConverter]::new()

function Format-YamlValue {
    <# Wraps a YAML scalar in single quotes if it contains special characters #>
    param([string]$Value)
    if (-not $Value) { return $Value }
    if ($Value -match '[{}\[\]:\''",&#*?|><!%@`]') {
        return "'" + ($Value -replace "'", "''") + "'"
    }
    return $Value
}

function New-StyledListViewItem {
    <# Creates a themed ListViewItem card with icon badge + name + subtitle #>
    param([string]$IconChar, [string]$PrimaryText, [string]$SubtitleText, $TagValue)
    $Item = New-Object System.Windows.Controls.ListViewItem
    $Row  = New-Object System.Windows.Controls.StackPanel
    $Row.Orientation = 'Horizontal'
    $Row.VerticalAlignment = 'Center'
    # Icon badge
    $IconBorder = New-Object System.Windows.Controls.Border
    $IconBorder.Width = 28; $IconBorder.Height = 28
    $IconBorder.CornerRadius = [System.Windows.CornerRadius]::new(6)
    $IconBorder.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeAccentDim')
    $IconBorder.Margin = [System.Windows.Thickness]::new(0, 0, 10, 0)
    $IconText = New-Object System.Windows.Controls.TextBlock
    $IconText.Text = $IconChar
    $IconText.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
    $IconText.FontSize = 13
    $IconText.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeAccentLight')
    $IconText.HorizontalAlignment = 'Center'; $IconText.VerticalAlignment = 'Center'
    $IconBorder.Child = $IconText
    # Text stack
    $TextStack = New-Object System.Windows.Controls.StackPanel
    $LblName = New-Object System.Windows.Controls.TextBlock
    $LblName.Text = $PrimaryText; $LblName.FontSize = 11
    $LblName.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextBody')
    $LblName.TextTrimming = 'CharacterEllipsis'
    $LblSub = New-Object System.Windows.Controls.TextBlock
    $LblSub.Text = $SubtitleText; $LblSub.FontSize = 9
    $LblSub.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
    $LblSub.Margin = [System.Windows.Thickness]::new(0, 1, 0, 0)
    $TextStack.Children.Add($LblName) | Out-Null
    $TextStack.Children.Add($LblSub) | Out-Null
    $Row.Children.Add($IconBorder) | Out-Null
    $Row.Children.Add($TextStack) | Out-Null
    $Item.Content = $Row
    $Item.Tag = $TagValue
    return $Item
}

# ─── Toast Notification System ─────────────────────────────────────────────
function Show-Toast {
    <# Shows a slide-in toast notification with auto-dismiss #>
    param(
        [string]$Message,
        [ValidateSet('Success','Error','Warning','Info')][string]$Type = 'Info',
        [int]$DurationMs = $Script:TOAST_DURATION_MS
    )
    if ($Global:IsLightMode) {
        $Colors = @{
            Success = @{ Bg='#E8F5E9'; Border='#16A34A'; TextColor='#14532D'; Icon=[char]0xE73E }
            Error   = @{ Bg='#FEE2E2'; Border='#DC2626'; TextColor='#7F1D1D'; Icon=[char]0xEA39 }
            Warning = @{ Bg='#FFF7ED'; Border='#EA580C'; TextColor='#7C2D12'; Icon=[char]0xE7BA }
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
    $MsgTB.MaxWidth    = 300
    $MsgTB.VerticalAlignment = 'Center'
    $SP.Children.Add($IconTB) | Out-Null
    $SP.Children.Add($MsgTB) | Out-Null
    $Toast.Child = $SP

    $pnlToastHost.Children.Insert(0, $Toast)

    # Fade In
    if ($Global:AnimationsDisabled) {
        $Toast.Opacity = 1
    } else {
        $FadeIn = New-Object System.Windows.Media.Animation.DoubleAnimation
        $FadeIn.From     = 0
        $FadeIn.To       = 1
        $FadeIn.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(200))
        $Toast.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $FadeIn)
    }

    # Auto-dismiss timer
    $DismissTimer = New-Object System.Windows.Threading.DispatcherTimer
    $DismissTimer.Interval = [TimeSpan]::FromMilliseconds($DurationMs)
    $ToastRef = $Toast
    $HostRef  = $pnlToastHost
    $TimerRef = $DismissTimer
    $AnimDisabledRef = $Global:AnimationsDisabled

    # Separate removal timer — fires 250ms after fade starts (avoids nested .GetNewClosure() crash in PS 5.1)
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
            $FadeOut.From     = 1
            $FadeOut.To       = 0
            $FadeOut.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(200))
            $ToastRef.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $FadeOut)
            $RemoveTimerRef.Start()
        }
        Write-DebugLog "[ToastTimer] DismissTimer completed OK (fade+remove scheduled)" -Level 'DEBUG'
        } catch {
            Write-DebugLog "[ToastTimer] CRASH: $($_.Exception.Message)" -Level 'ERROR'
            Write-DebugLog "[ToastTimer] ScriptStackTrace: $($_.ScriptStackTrace)" -Level 'ERROR'
        }
    }.GetNewClosure())
    $DismissTimer.Start()
    Write-DebugLog "Toast [$Type]: $Message"
}

# ─── Tab Transition (opacity fade) ────────────────────────────────────────
function Invoke-TabFade {
    <# Fades in the target panel with a 150ms opacity animation #>
    param([System.Windows.UIElement]$Panel)
    $Panel.Visibility = 'Visible'
    if ($Global:AnimationsDisabled) {
        $Panel.Opacity = 1
        return
    }
    $Panel.Opacity = 0
    $Fade = New-Object System.Windows.Media.Animation.DoubleAnimation
    $Fade.From     = 0
    $Fade.To       = 1
    $Fade.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(150))
    $Panel.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $Fade)
}

# ─── Collapsible Section Toggle ───────────────────────────────────────────
function Toggle-Section {
    <# Toggles a collapsible section body and updates chevron icon #>
    param(
        [System.Windows.Controls.TextBlock]$Chevron,
        [System.Windows.UIElement]$Body
    )
    if ($Body.Visibility -eq 'Visible') {
        $Body.Visibility = 'Collapsed'
        $Chevron.Text = [char]0xE76C  # right chevron (collapsed)
    } else {
        $Body.Visibility = 'Visible'
        $Chevron.Text = [char]0xE70D  # down chevron (expanded)
    }
}

# ─── Inline Field Validation ──────────────────────────────────────────────
function Add-FieldValidation {
    <# Adds a LostFocus handler to a TextBox for required-field validation #>
    param(
        [System.Windows.Controls.TextBox]$Field,
        [string]$FieldLabel
    )
    $Field.Add_LostFocus({
        param($s, $e)
        if ([string]::IsNullOrWhiteSpace($s.Text)) {
            $s.BorderBrush = $Window.Resources['ThemeError']
            $s.ToolTip = "$FieldLabel is required"
        } else {
            $s.BorderBrush = $Window.Resources['ThemeBorder']
            $s.ToolTip = $null
        }
    }.GetNewClosure())
}

# ─── URL Format Validation ────────────────────────────────────────────────
function Add-UrlValidation {
    <# Adds a LostFocus handler to validate URL format #>
    param(
        [System.Windows.Controls.TextBox]$Field,
        [string]$FieldLabel
    )
    $Field.Add_LostFocus({
        param($s, $e)
        $Val = $s.Text.Trim()
        if ($Val -and $Val -notmatch '^https?://') {
            $s.BorderBrush = $Window.Resources['ThemeWarning']
            $s.ToolTip = "$FieldLabel must start with http:// or https://"
        } elseif ([string]::IsNullOrWhiteSpace($Val)) {
            $s.BorderBrush = $Window.Resources['ThemeBorder']
            $s.ToolTip = $null
        } else {
            $s.BorderBrush = $Window.Resources['ThemeBorder']
            $s.ToolTip = $null
        }
    }.GetNewClosure())
}

# ─── Progress Ring on Buttons ─────────────────────────────────────────────
function Set-ButtonBusy {
    <# Swaps a button's content to show a spinning progress indicator #>
    param(
        [System.Windows.Controls.Button]$Button,
        [string]$Text = 'Working...'
    )
    $Button.Tag = $Button.Content
    $Button.IsEnabled = $false
    $SP = New-Object System.Windows.Controls.StackPanel
    $SP.Orientation = 'Horizontal'
    $Ring = New-Object System.Windows.Controls.TextBlock
    $Ring.Text       = [char]0xF16A  # Progress ring icon
    $Ring.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe Fluent Icons, Segoe MDL2 Assets")
    $Ring.FontSize   = 12
    $Ring.Foreground = [System.Windows.Media.Brushes]::White
    $Ring.VerticalAlignment = 'Center'
    $Ring.Margin     = [System.Windows.Thickness]::new(0,0,6,0)
    # Spin animation
    if (-not $Global:AnimationsDisabled) {
        $Spin = New-Object System.Windows.Media.Animation.DoubleAnimation
        $Spin.From = 0; $Spin.To = 360
        $Spin.Duration       = [System.Windows.Duration]::new([TimeSpan]::FromSeconds(1))
        $Spin.RepeatBehavior = [System.Windows.Media.Animation.RepeatBehavior]::Forever
        $RT = New-Object System.Windows.Media.RotateTransform
        $Ring.RenderTransform = $RT
        $Ring.RenderTransformOrigin = [System.Windows.Point]::new(0.5, 0.5)
        $RT.BeginAnimation([System.Windows.Media.RotateTransform]::AngleProperty, $Spin)
    }
    $LblTB = New-Object System.Windows.Controls.TextBlock
    $LblTB.Text = $Text; $LblTB.FontSize = 12; $LblTB.Foreground = [System.Windows.Media.Brushes]::White
    $LblTB.VerticalAlignment = 'Center'
    $SP.Children.Add($Ring) | Out-Null
    $SP.Children.Add($LblTB) | Out-Null
    $Button.Content = $SP
}

function Reset-ButtonBusy {
    <# Restores a button to its original content #>
    param([System.Windows.Controls.Button]$Button)
    if ($null -ne $Button.Tag) {
        $Button.Content = $Button.Tag
        $Button.Tag = $null
    }
    $Button.IsEnabled = $true
}

# ─── Drag & Drop Handler ──────────────────────────────────────────────────
function Register-DropZone {
    <# Registers drag-and-drop events on a Border with drop-zone visual feedback #>
    param(
        [System.Windows.Controls.Border]$Zone,
        [scriptblock]$OnFileDrop  # receives [string]$FilePath
    )
    $ZoneRef = $Zone
    $Zone.Add_DragEnter({
        param($s, $e)
        if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
            $e.Effects = [System.Windows.DragDropEffects]::Copy
            $s.BorderBrush = $Window.Resources['ThemeAccent']
            $s.BorderThickness = [System.Windows.Thickness]::new(2)
        }
    })
    $Zone.Add_DragLeave({
        param($s, $e)
        $s.BorderBrush = $Window.Resources['ThemeBorderCard']
        $s.BorderThickness = [System.Windows.Thickness]::new(1)
    })
    $Zone.Add_Drop({
        param($s, $e)
        $s.BorderBrush = $Window.Resources['ThemeBorderCard']
        $s.BorderThickness = [System.Windows.Thickness]::new(1)
        if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
            $Files = $e.Data.GetData([System.Windows.DataFormats]::FileDrop)
            if ($Files -and $Files.Count -gt 0) {
                & $OnFileDrop $Files[0]
            }
        }
    }.GetNewClosure())
}

# ─── URL Reachability Check ───────────────────────────────────────────────
function Test-InstallerUrl {
    <# Tests if an installer URL is reachable via HEAD request #>
    param([string]$Url)
    if (-not $Url -or $Url -notmatch '^https?://') {
        Show-Toast "Please enter a valid URL first" -Type Warning
        return
    }
    Show-Toast "Testing URL reachability..." -Type Info -DurationMs 2000
    Start-BackgroundWork -Variables @{ TestUrl = $Url } -Work {
        try {
            $Resp = Invoke-WebRequest -Uri $TestUrl -Method Head -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
            return @{ Ok = $true; Status = $Resp.StatusCode; Size = $Resp.Headers['Content-Length'] }
        } catch {
            return @{ Ok = $false; Error = $_.Exception.Message }
        }
    } -OnComplete {
        param($Results, $Errors)
        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ Ok = $false; Error = 'No result' } }
        if ($R.Ok) {
            $SizeStr = if ($R.Size -is [array]) { $R.Size[0] } else { $R.Size }
            $SizeInfo = if ($SizeStr) { " ($([math]::Round([int64]$SizeStr/1MB, 1)) MB)" } else { '' }
            Show-Toast "URL reachable — HTTP $($R.Status)$SizeInfo" -Type Success
        } else {
            Show-Toast "URL unreachable: $($R.Error)" -Type Error -DurationMs 6000
        }
    }
}

# ─── Full Manifest Schema Validation ──────────────────────────────────────
function Validate-ManifestSchema {
    <# Performs comprehensive schema validation on all manifest fields #>
    $Errors = @()

    # Package Identifier: Publisher.Package format (hyphens allowed per WinGet schema)
    $PkgId = $txtPkgId.Text.Trim()
    if (-not $PkgId) { $Errors += 'Package Identifier is required' }
    elseif ($PkgId -notmatch '^[\w\-]+(\.[\w\-]+)+$') { $Errors += 'Package Identifier must be in Publisher.Package format (e.g., 7zip.7-Zip)' }

    # Version
    $Ver = $txtPkgVersion.Text.Trim()
    if (-not $Ver) { $Errors += 'Package Version is required' }
    elseif ($Ver -notmatch '^\d+(\.\d+)*([a-zA-Z0-9\.\-+]*)$') { $Errors += 'Package Version format appears invalid' }

    # Name
    if (-not $txtPkgName.Text.Trim()) { $Errors += 'Package Name is required' }

    # Publisher
    if (-not $txtPublisher.Text.Trim()) { $Errors += 'Publisher is required' }

    # License
    if (-not $txtLicense.Text.Trim()) { $Errors += 'License is required' }

    # Short description
    $SD = $txtShortDesc.Text.Trim()
    if (-not $SD) { $Errors += 'Short Description is required' }
    elseif ($SD.Length -gt 256) { $Errors += "Short Description exceeds 256 characters ($($SD.Length))" }

    # Installer Type
    $InsType = Get-ComboBoxSelectedText $cmbInstallerType0
    if (-not $InsType) { $Errors += 'Installer Type is required' }

    # SHA256 hash format
    $Hash = $Global:CurrentHash0
    if ($Hash -and $Hash -notmatch '^[A-Fa-f0-9]{64}$') { $Errors += 'SHA256 hash must be 64 hex characters' }

    # Installer URL format
    $InsUrl = $txtInstallerUrl0.Text.Trim()
    if ($InsUrl -and $InsUrl -notmatch '^https?://') { $Errors += 'Installer URL must start with http:// or https://' }

    # URL fields validation
    foreach ($UrlField in @(
        @{ Ctrl=$txtPackageUrl;    Name='Package URL' }
        @{ Ctrl=$txtPublisherUrl;  Name='Publisher URL' }
        @{ Ctrl=$txtLicenseUrl;    Name='License URL' }
        @{ Ctrl=$txtSupportUrl;    Name='Support URL' }
        @{ Ctrl=$txtPrivacyUrl;    Name='Privacy URL' }
        @{ Ctrl=$txtReleaseNotesUrl; Name='Release Notes URL' }
    )) {
        $V = $UrlField.Ctrl.Text.Trim()
        if ($V -and $V -notmatch '^https?://') { $Errors += "$($UrlField.Name) must start with http:// or https://" }
    }

    # Enum validations
    $ValidTypes = @('exe','msi','msix','wix','zip','burn','inno','nullsoft','portable')
    if ($InsType -and $InsType -notin $ValidTypes) { $Errors += "Invalid Installer Type: $InsType" }

    $Upgrade = Get-ComboBoxSelectedText $cmbUpgrade0
    if ($Upgrade -and $Upgrade -notin @('install','uninstallPrevious','deny')) {
        $Errors += "Invalid Upgrade Behavior: $Upgrade"
    }

    return $Errors
}

# ─── Stepper Navigation ───────────────────────────────────────────────────
function Update-Stepper {
    <# Updates the stepper UI to highlight the given step #>
    param([int]$ActiveStep)  # 1=Identity, 2=Metadata, 3=Installer, 4=Review
    $Steps = @(
        @{ Btn=$stepIdentity;  Num='1' }
        @{ Btn=$stepMetadata;  Num='2' }
        @{ Btn=$stepInstaller; Num='3' }
        @{ Btn=$stepReview;    Num='4' }
    )
    foreach ($i in 0..3) {
        $S = $Steps[$i]
        $SP = $S.Btn.Content
        if ($SP -is [System.Windows.Controls.StackPanel]) {
            $Circle = $SP.Children[0]  # Border
            $Label  = $SP.Children[1]  # TextBlock
            $NumTB  = $Circle.Child     # TextBlock inside circle
            if (($i + 1) -le $ActiveStep) {
                $Circle.Background = $Window.Resources['ThemeAccent']
                $NumTB.Foreground  = [System.Windows.Media.Brushes]::White
                $Label.Foreground  = $Window.Resources['ThemeTextPrimary']
                $Label.FontWeight  = [System.Windows.FontWeights]::SemiBold
            } else {
                $Circle.Background = $Window.Resources['ThemeSurfaceBg']
                $NumTB.Foreground  = $Window.Resources['ThemeTextMuted']
                $Label.Foreground  = $Window.Resources['ThemeTextMuted']
                $Label.FontWeight  = [System.Windows.FontWeights]::Normal
            }
        }
    }
}

# ─── Retry Logic Wrapper ──────────────────────────────────────────────────
function Invoke-WithRetry {
    <# Retries a scriptblock with exponential backoff #>
    param(
        [scriptblock]$Action,
        [int]$MaxRetries = 3,
        [int]$BaseDelayMs = 500,
        [string]$OperationName = 'Operation'
    )
    $Attempt = 0
    while ($true) {
        $Attempt++
        try {
            return (& $Action)
        } catch {
            if ($Attempt -ge $MaxRetries) {
                throw $_
            }
            $Delay = $BaseDelayMs * [math]::Pow(2, $Attempt - 1)
            Write-DebugLog "$OperationName failed (attempt $Attempt/$MaxRetries), retrying in ${Delay}ms: $($_.Exception.Message)" -Level 'WARN'
            Start-Sleep -Milliseconds $Delay
        }
    }
}

# ─── Temp File Cleanup ────────────────────────────────────────────────────
function Clear-TempFiles {
    <# Cleans up stale temp files from previous sessions #>
    $TempPattern = Join-Path ([System.IO.Path]::GetTempPath()) 'WinGetMM_*'
    $StaleFiles = Get-ChildItem -Path $TempPattern -Recurse -Force -ErrorAction SilentlyContinue |
                  Where-Object { $_.LastWriteTime -lt (Get-Date).AddHours(-1) }
    $Cleaned = 0
    foreach ($F in $StaleFiles) {
        try {
            if ($F.PSIsContainer) { Remove-Item $F.FullName -Recurse -Force -ErrorAction Stop }
            else { Remove-Item $F.FullName -Force -ErrorAction Stop }
            $Cleaned++
        } catch { Write-DebugLog "Clear-TempFiles: failed to remove '$($F.Name)': $($_.Exception.Message)" -Level 'WARN' }
    }
    if ($Cleaned -gt 0) { Write-DebugLog "Cleaned $Cleaned stale temp files" -Level 'INFO' }
}

# ─── Concurrent Publish Guard ─────────────────────────────────────────────
$Script:PublishInProgress = $false

# ==============================================================================
# SECTION: CONFETTI ANIMATION
# ==============================================================================

function Start-ConfettiAnimation {
    <# Robinhood-style confetti celebration on successful publish/milestone #>
    if ($Global:AnimationsDisabled) { return }
    if (-not $cnvConfetti) { return }
    if ($cnvConfetti.Visibility -eq 'Visible') { return }
    $W = $Window.ActualWidth
    $H = $Window.ActualHeight
    if ($W -le 0 -or $H -le 0) { return }

    $cnvConfetti.Children.Clear()
    $cnvConfetti.Visibility = 'Visible'

    $Colors = @('#FF4444','#FFD700','#00C853','#60CDFF','#0078D4','#B388FF','#FF6D00','#E040FB')
    $Rand   = [System.Random]::new()

    for ($i = 0; $i -lt $Script:CONFETTI_COUNT; $i++) {
        $Size = $Rand.Next(4, 10)
        $Rect = [System.Windows.Shapes.Rectangle]::new()
        $Rect.Width  = $Size
        $Rect.Height = $Size * ($Rand.NextDouble() * 1.5 + 0.5)
        $Rect.Fill   = $Global:CachedBC.ConvertFromString($Colors[$Rand.Next($Colors.Count)])
        $Rect.RadiusX = if ($Rand.Next(3) -eq 0) { $Size / 2 } else { 1 }
        $Rect.RadiusY = $Rect.RadiusX
        $Rect.Opacity = 0.9
        $Rect.RenderTransform = [System.Windows.Media.RotateTransform]::new($Rand.Next(360))

        $X0 = $Rand.NextDouble() * $W
        [System.Windows.Controls.Canvas]::SetLeft($Rect, $X0)
        [System.Windows.Controls.Canvas]::SetTop($Rect, -20)

        $cnvConfetti.Children.Add($Rect) | Out-Null

        # Fall animation
        $Fall = [System.Windows.Media.Animation.DoubleAnimation]::new()
        $Fall.From     = $Rand.Next(-40, -10)
        $Fall.To       = $H + 20
        $Fall.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds($Rand.Next(2000, 4500)))
        $Ease = [System.Windows.Media.Animation.CubicEase]::new()
        $Ease.EasingMode = 'EaseIn'
        $Fall.EasingFunction = $Ease

        # Drift animation (horizontal sway)
        $Drift = [System.Windows.Media.Animation.DoubleAnimation]::new()
        $Drift.From     = $X0
        $Drift.To       = $X0 + $Rand.Next(-120, 120)
        $Drift.Duration = $Fall.Duration

        # Spin animation
        $Spin = [System.Windows.Media.Animation.DoubleAnimation]::new()
        $Spin.From     = 0
        $Spin.To       = $Rand.Next(-720, 720)
        $Spin.Duration = $Fall.Duration

        $Rect.BeginAnimation([System.Windows.Controls.Canvas]::TopProperty, $Fall)
        $Rect.BeginAnimation([System.Windows.Controls.Canvas]::LeftProperty, $Drift)
        $Rect.RenderTransform.BeginAnimation([System.Windows.Media.RotateTransform]::AngleProperty, $Spin)
    }

    # Auto-cleanup after 5 seconds
    $CleanTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $CleanTimer.Interval = [TimeSpan]::FromMilliseconds($Script:CLEANUP_DELAY_MS)
    $CleanTimer.Tag = $cnvConfetti
    $CleanTimer.Add_Tick({
        Write-DebugLog "[ConfettiTimer] Cleanup fired — this is null? $($null -eq $this), Tag is null? $($null -eq $this.Tag)" -Level 'DEBUG'
        try {
            $this.Tag.Children.Clear()
            $this.Tag.Visibility = 'Collapsed'
            $this.Stop()
            Write-DebugLog "[ConfettiTimer] Cleanup OK" -Level 'DEBUG'
        } catch {
            Write-DebugLog "[ConfettiTimer] CRASH: $($_.Exception.Message)" -Level 'ERROR'
            Write-DebugLog "[ConfettiTimer] ScriptStackTrace: $($_.ScriptStackTrace)" -Level 'ERROR'
        }
    })
    $CleanTimer.Start()
}

# ==============================================================================
# SECTION: PUBLISH STREAK TRACKER
# ==============================================================================

$Script:StreakFile = Join-Path $Global:Root 'publish_streak.json'

function Load-PublishStreak {
    if (Test-Path $Script:StreakFile) {
        try {
            $Data = Get-Content $Script:StreakFile -Raw | ConvertFrom-Json
            $Script:PublishStreak = [int]$Data.streak
            $Script:LastPublishDate = [datetime]$Data.lastDate
            $Script:TotalPublishes = if ($null -ne $Data.total) { [int]$Data.total } else { 0 }
        } catch {
            $Script:PublishStreak   = 0
            $Script:LastPublishDate = [datetime]::MinValue
            $Script:TotalPublishes  = 0
        }
    } else {
        $Script:PublishStreak   = 0
        $Script:LastPublishDate = [datetime]::MinValue
        $Script:TotalPublishes  = 0
    }
}

function Save-PublishStreak {
    @{ streak = $Script:PublishStreak; lastDate = $Script:LastPublishDate.ToString('o'); total = $Script:TotalPublishes } |
        ConvertTo-Json | Set-Content $Script:StreakFile -Force
}

function Update-StreakDisplay {
    if (-not $lblStreak) { return }
    if ($Script:PublishStreak -le 0) {
        $lblStreak.Text = ''
        return
    }
    # Fire emoji: Unicode 0x1F525
    $Fire = [char]::ConvertFromUtf32(0x1F525)
    $lblStreak.Text = "$Fire $($Script:PublishStreak) publish streak!"

    if ($Script:PublishStreak -ge 10) {
        $lblStreak.Foreground = $Window.FindResource('ThemeError')
    } elseif ($Script:PublishStreak -ge 5) {
        $lblStreak.Foreground = $Window.FindResource('ThemeWarning')
    } else {
        $lblStreak.Foreground = $Window.FindResource('ThemeTextMuted')
    }
}

function Record-PublishSuccess {
    $Now = Get-Date
    $Script:TotalPublishes++
    # If last publish was within 24 hours, increment streak; else reset
    if (($Now - $Script:LastPublishDate).TotalHours -le $Script:STREAK_HOURS) {
        $Script:PublishStreak++
    } else {
        $Script:PublishStreak = 1
    }
    $Script:LastPublishDate = $Now
    Save-PublishStreak
    Update-StreakDisplay

    # Trigger confetti at streak milestones or every publish
    if ($Script:PublishStreak -ge 3) {
        Start-ConfettiAnimation
    }
}

# ==============================================================================
# SECTION: ACHIEVEMENT SYSTEM
# ==============================================================================

$Script:AchievementsFile = Join-Path $Global:Root 'achievements.json'
$Script:Achievements     = @{}

$Script:AchievementDefs = @(
    @{ Id='first_publish';  Icon=[char]::ConvertFromUtf32(0x1F389); Name='First Publish';  Desc='Published your first package' }
    @{ Id='five_packages';  Icon=[char]::ConvertFromUtf32(0x1F4E6); Name='Five Star';      Desc='Published 5 packages total' }
    @{ Id='ten_packages';   Icon=[char]::ConvertFromUtf32(0x1F3C6); Name='Package Pro';    Desc='Published 10 packages total' }
    @{ Id='twentyfive_pkg';  Icon=[char]::ConvertFromUtf32(0x1F48E); Name='Diamond Shipper'; Desc='Published 25 packages total' }
    @{ Id='streak_3';       Icon=[char]::ConvertFromUtf32(0x1F525); Name='On a Roll';      Desc='3 publishes in 24 hours' }
    @{ Id='streak_5';       Icon=[char]::ConvertFromUtf32(0x1F525); Name='On Fire';        Desc='5 publishes in 24 hours' }
    @{ Id='streak_10';      Icon=[char]::ConvertFromUtf32(0x1F4AA); Name='Unstoppable';    Desc='10 publishes in 24 hours' }
    @{ Id='perfect_yaml';   Icon=[char]::ConvertFromUtf32(0x2728);  Name='Perfectionist';  Desc='Published with 100% quality score' }
    @{ Id='night_owl';      Icon=[char]::ConvertFromUtf32(0x1F989); Name='Night Owl';      Desc='Published between midnight and 5am' }
    @{ Id='early_bird';     Icon=[char]::ConvertFromUtf32(0x1F305); Name='Early Bird';     Desc='Published between 5am and 7am' }
    @{ Id='speed_demon';    Icon=[char]::ConvertFromUtf32(0x26A1);  Name='Speed Demon';    Desc='Published in under 10 seconds' }
    @{ Id='weekend_warrior'; Icon=[char]::ConvertFromUtf32(0x1F6E1); Name='Weekend Warrior'; Desc='Published on a weekend' }
    @{ Id='psadt_master';   Icon=[char]::ConvertFromUtf32(0x1F4E6); Name='PSADT Master';   Desc='Published a PSADT package' }
    @{ Id='centurion';      Icon=[char]::ConvertFromUtf32(0x1F4AF); Name='Centurion';      Desc='Published 100 packages total' }
    @{ Id='first_import';   Icon=[char]::ConvertFromUtf32(0x1F310); Name='Community Curator'; Desc='Imported a package from WinGet community repo' }
    @{ Id='import_10';      Icon=[char]::ConvertFromUtf32(0x1F30D); Name='Globetrotter';   Desc='Imported 10 packages from community repo' }
    @{ Id='first_upload';   Icon=[char]::ConvertFromUtf32(0x2601);  Name='Cloud First';    Desc='Uploaded your first installer to blob storage' }
    @{ Id='batch_download'; Icon=[char]::ConvertFromUtf32(0x1F4E5); Name='Batch Commander'; Desc='Batch downloaded all installers for a package' }
    @{ Id='batch_upload';   Icon=[char]::ConvertFromUtf32(0x1F680); Name='Cloud Fleet';    Desc='Batch uploaded all installers to blob storage' }
    @{ Id='multi_arch';     Icon=[char]::ConvertFromUtf32(0x1F3D7); Name='Multi-Arch';     Desc='Configured a package with 5+ installers' }
    @{ Id='ten_configs';    Icon=[char]::ConvertFromUtf32(0x1F4DA); Name='Librarian';      Desc='Saved 10 package configs' }
    @{ Id='first_backup';   Icon=[char]::ConvertFromUtf32(0x1F4BE); Name='Safety Net';     Desc='Completed your first backup' }
    @{ Id='first_restore';  Icon=[char]::ConvertFromUtf32(0x1F504); Name='Phoenix';        Desc='Restored from a backup' }
    @{ Id='first_delete';   Icon=[char]::ConvertFromUtf32(0x1F5D1); Name='Spring Cleaning'; Desc='Removed a package from REST source' }
    @{ Id='fifty_packages'; Icon=[char]::ConvertFromUtf32(0x2B50);  Name='Half Century';   Desc='Published 50 packages total' }
    @{ Id='theme_toggle';   Icon=[char]::ConvertFromUtf32(0x1F3A8); Name='Chameleon';      Desc='Toggled the theme for the first time' }
    @{ Id='diff_viewer';    Icon=[char]::ConvertFromUtf32(0x1F50D); Name='Diff Detective'; Desc='Compared two package versions' }
    @{ Id='yaml_export';    Icon=[char]::ConvertFromUtf32(0x1F4C4); Name='Disk Jockey';    Desc='Saved manifest YAML to disk' }
    @{ Id='shortcut_user';  Icon=[char]::ConvertFromUtf32(0x2328);  Name='Keyboard Ninja'; Desc='Used a keyboard shortcut' }
    @{ Id='blob_cleanup';   Icon=[char]::ConvertFromUtf32(0x1F9F9); Name='Tidy Cloud';     Desc='Deleted a blob from storage' }
)

function Load-Achievements {
    if (Test-Path $Script:AchievementsFile) {
        try {
            $Raw = Get-Content $Script:AchievementsFile -Raw | ConvertFrom-Json
            $Script:Achievements = @{}
            foreach ($Prop in $Raw.PSObject.Properties) {
                $Script:Achievements[$Prop.Name] = $Prop.Value
            }
        } catch { <# Corrupted achievements file — reset to empty #> $Script:Achievements = @{} }
    }
}

function Save-Achievements {
    $Script:Achievements | ConvertTo-Json | Set-Content $Script:AchievementsFile -Force
}

function Unlock-Achievement {
    param([string]$Id)
    if ($Script:Achievements.ContainsKey($Id)) { return }

    $Def = $Script:AchievementDefs | Where-Object { $_.Id -eq $Id }
    if (-not $Def) { return }

    $Script:Achievements[$Id] = (Get-Date).ToString('o')
    Save-Achievements

    Show-Toast "$($Def.Icon) Achievement Unlocked: $($Def.Name) — $($Def.Desc)" -Type Success -DurationMs 5000
    Start-ConfettiAnimation
    Render-Achievements
    Write-DebugLog "Achievement unlocked: $($Def.Name)" -Level 'INFO'
}

function Render-Achievements {
    if (-not $pnlAchievements) { return }
    $pnlAchievements.Children.Clear()
    $Unlocked = 0
    foreach ($Def in $Script:AchievementDefs) {
        $IsUnlocked = $Script:Achievements.ContainsKey($Def.Id)
        if ($IsUnlocked) { $Unlocked++ }

        $Badge = [System.Windows.Controls.Border]::new()
        $Badge.Width        = 32
        $Badge.Height       = 32
        $Badge.CornerRadius = [System.Windows.CornerRadius]::new(6)
        $Badge.Margin       = [System.Windows.Thickness]::new(0, 0, 4, 4)
        $Badge.ToolTip      = if ($IsUnlocked) { "$($Def.Name): $($Def.Desc)" } else { '???' }

        if ($IsUnlocked) {
            $Badge.Background = $Window.FindResource('ThemeSelectedBg')
            $Badge.BorderBrush = $Window.FindResource('ThemeBorderElevated')
            $Badge.BorderThickness = [System.Windows.Thickness]::new(1)
            $Lbl = [System.Windows.Controls.TextBlock]::new()
            $Lbl.Text = $Def.Icon
            $Lbl.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI Emoji")
            $Lbl.FontSize = 16
            $Lbl.HorizontalAlignment = 'Center'
            $Lbl.VerticalAlignment   = 'Center'
        } else {
            $Badge.Background = $Window.FindResource('ThemeDeepBg')
            $Badge.BorderBrush = $Window.FindResource('ThemeBorder')
            $Badge.BorderThickness = [System.Windows.Thickness]::new(1)
            $Lbl = [System.Windows.Controls.TextBlock]::new()
            $Lbl.Text = '?'
            $Lbl.FontSize   = 14
            $Lbl.Foreground = $Window.FindResource('ThemeTextDisabled')
            $Lbl.HorizontalAlignment = 'Center'
            $Lbl.VerticalAlignment   = 'Center'
        }
        $Badge.Child = $Lbl
        $pnlAchievements.Children.Add($Badge) | Out-Null
    }
    if ($lblAchievementCount) { $lblAchievementCount.Text = "$Unlocked/$($Script:AchievementDefs.Count)" }
}

# Achievements collapse/expand toggle
$Global:AchievementsCollapsed = $false
if ($btnToggleAchievements) {
    $btnToggleAchievements.Add_Click({
        $Global:AchievementsCollapsed = -not $Global:AchievementsCollapsed
        if ($Global:AchievementsCollapsed) {
            $pnlAchievements.Visibility = 'Collapsed'
            $txtAchievementChevron.RenderTransform = [System.Windows.Media.RotateTransform]::new(0)
        } else {
            $pnlAchievements.Visibility = 'Visible'
            $txtAchievementChevron.RenderTransform = [System.Windows.Media.RotateTransform]::new(180)
        }
    })
}

function Check-PublishAchievements {
    # First publish
    if (-not $Script:Achievements.ContainsKey('first_publish')) {
        Unlock-Achievement 'first_publish'
    }

    # Publish count milestones
    if ($Script:TotalPublishes -ge 5)   { Unlock-Achievement 'five_packages' }
    if ($Script:TotalPublishes -ge 10)  { Unlock-Achievement 'ten_packages' }
    if ($Script:TotalPublishes -ge 25)  { Unlock-Achievement 'twentyfive_pkg' }
    if ($Script:TotalPublishes -ge 50)  { Unlock-Achievement 'fifty_packages' }
    if ($Script:TotalPublishes -ge 100) { Unlock-Achievement 'centurion' }

    # Streak milestones
    if ($Script:PublishStreak -ge 3)  { Unlock-Achievement 'streak_3' }
    if ($Script:PublishStreak -ge 5)  { Unlock-Achievement 'streak_5' }
    if ($Script:PublishStreak -ge 10) { Unlock-Achievement 'streak_10' }

    # Perfect YAML — read current quality score from UI
    $QualityPct = 0
    if ($lblQualityPct -and $lblQualityPct.Text -match '(\d+)') {
        $QualityPct = [int]$Matches[1]
    }
    if ($QualityPct -ge 100) { Unlock-Achievement 'perfect_yaml' }

    # Time-based achievements
    $Hour = (Get-Date).Hour
    if ($Hour -ge 0 -and $Hour -lt 5) { Unlock-Achievement 'night_owl' }
    if ($Hour -ge 5 -and $Hour -lt 7) { Unlock-Achievement 'early_bird' }

    # Weekend warrior
    $DayOfWeek = (Get-Date).DayOfWeek
    if ($DayOfWeek -eq 'Saturday' -or $DayOfWeek -eq 'Sunday') { Unlock-Achievement 'weekend_warrior' }

    # Speed demon — check elapsed time from publish start
    if ($Script:PublishStartTime) {
        $Elapsed = ((Get-Date) - $Script:PublishStartTime).TotalSeconds
        if ($Elapsed -lt 10) { Unlock-Achievement 'speed_demon' }
    }

    # PSADT master — check if the current package has PSADT tag
    if ($txtTags -and $txtTags.Text -match 'PSADT') { Unlock-Achievement 'psadt_master' }
}

# ==============================================================================
# SECTION: MANIFEST QUALITY SCORE
# ==============================================================================

function Update-QualityMeter {
    <# Calculates and displays a manifest completeness/quality percentage #>
    $Score  = 0
    $Total  = 0

    # Required fields (weighted heavier)
    $RequiredFields = @(
        @{ Field=$txtPkgId; Weight=10 }
        @{ Field=$txtPkgVersion; Weight=10 }
        @{ Field=$txtPkgName; Weight=10 }
        @{ Field=$txtPublisher; Weight=10 }
        @{ Field=$txtLicense; Weight=10 }
        @{ Field=$txtShortDesc; Weight=10 }
    )

    foreach ($RF in $RequiredFields) {
        $Total += $RF.Weight
        if ($RF.Field -and $RF.Field.Text.Trim().Length -gt 0) { $Score += $RF.Weight }
    }

    # Optional metadata fields (bonus)
    $OptionalFields = @(
        @{ Field=$txtPublisherUrl; Weight=5 }
        @{ Field=$txtPackageUrl; Weight=5 }
        @{ Field=$txtLicenseUrl; Weight=5 }
        @{ Field=$txtSupportUrl; Weight=3 }
        @{ Field=$txtPrivacyUrl; Weight=3 }
        @{ Field=$txtDescription; Weight=5 }
        @{ Field=$txtReleaseNotes; Weight=3 }
        @{ Field=$txtReleaseNotesUrl; Weight=2 }
        @{ Field=$txtAuthor; Weight=3 }
        @{ Field=$txtTags; Weight=3 }
        @{ Field=$txtMoniker; Weight=2 }
    )

    foreach ($OF in $OptionalFields) {
        $Total += $OF.Weight
        if ($OF.Field -and $OF.Field.Text.Trim().Length -gt 0) { $Score += $OF.Weight }
    }

    # Installer URL check
    $Total += 10
    if ($txtInstallerUrl0 -and $txtInstallerUrl0.Text -match '^https?://') { $Score += 10 }

    # SHA256 presence
    $Total += 5
    if ($Global:CurrentHash0 -match '^[A-Fa-f0-9]{64}$') { $Score += 5 }

    $Pct = if ($Total -gt 0) { [math]::Round(($Score / $Total) * 100) } else { 0 }

    # Update visual
    $MaxWidth = 60
    $BarWidth = [math]::Round($Pct / 100 * $MaxWidth)
    if ($barQuality) { $barQuality.Width = $BarWidth }
    if ($lblQualityPct) { $lblQualityPct.Text = "${Pct}%" }

    # Color code: red < 40, yellow < 70, green >= 70
    $Color = if ($Pct -ge 70) { '#00C853' } elseif ($Pct -ge 40) { '#FFB900' } else { '#D13438' }
    $Brush = $Global:CachedBC.ConvertFromString($Color)
    if ($barQuality) { $barQuality.Background  = $Brush }
    if ($lblQualityPct) { $lblQualityPct.Foreground = $Brush }

    return $Pct
}

# ==============================================================================
# SECTION: SHAKE ANIMATION (validation error feedback)
# ==============================================================================

function Invoke-ShakeAnimation {
    <# Applies a quick horizontal shake to a UI element on validation failure #>
    param([System.Windows.UIElement]$Element)
    if (-not $Element) { return }
    if ($Global:AnimationsDisabled) { return }

    $Transform = [System.Windows.Media.TranslateTransform]::new()
    $Element.RenderTransform = $Transform

    # Keyframe animation: 0→-6→6→-4→4→-2→0 over 400ms
    $Anim = [System.Windows.Media.Animation.DoubleAnimationUsingKeyFrames]::new()
    $Anim.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(400))

    $Ease = [System.Windows.Media.Animation.CubicEase]::new()
    $Ease.EasingMode = 'EaseOut'

    $Frames = @(
        @{ Time=50;  Val=-6 }
        @{ Time=100; Val=6 }
        @{ Time=175; Val=-4 }
        @{ Time=250; Val=4 }
        @{ Time=325; Val=-2 }
        @{ Time=400; Val=0 }
    )

    foreach ($F in $Frames) {
        $KF = [System.Windows.Media.Animation.EasingDoubleKeyFrame]::new()
        $KF.KeyTime = [System.Windows.Media.Animation.KeyTime]::FromTimeSpan([TimeSpan]::FromMilliseconds($F.Time))
        $KF.Value   = $F.Val
        $KF.EasingFunction = $Ease
        $Anim.KeyFrames.Add($KF) | Out-Null
    }

    $Transform.BeginAnimation([System.Windows.Media.TranslateTransform]::XProperty, $Anim)
}

# ==============================================================================
# SECTION: STATUS DOT PULSE ANIMATION
# ==============================================================================

function Start-StatusPulse {
    <# Pulses the status dot to indicate an operation is in progress #>
    if (-not $statusDot) { return }
    $statusDot.Fill = $Global:CachedBC.ConvertFromString('#FFB900')
    if ($Global:AnimationsDisabled) { return }
    $Pulse = [System.Windows.Media.Animation.DoubleAnimation]::new()
    $Pulse.From       = 1.0
    $Pulse.To         = 0.3
    $Pulse.Duration   = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(600))
    $Pulse.AutoReverse = $true
    $Pulse.RepeatBehavior = [System.Windows.Media.Animation.RepeatBehavior]::Forever
    $statusDot.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $Pulse)
    $statusDot.Fill = $Global:CachedBC.ConvertFromString('#FFB900')
}

function Stop-StatusPulse {
    <# Stops the status dot animation and resets to idle #>
    if (-not $statusDot) { return }
    $statusDot.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $null)
    $statusDot.Opacity = 1.0
}

function Reset-StatusDot {
    <# Resets status dot to default idle state #>
    if (-not $statusDot) { return }
    $statusDot.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $null)
    $statusDot.Opacity = 1.0
    $statusDot.Fill = $Window.FindResource('ThemeTextDim')
}

function Build-VersionYaml {
    <# Generates the version manifest YAML #>
    param([string]$PkgId, [string]$PkgVersion, [string]$ManifestVersion, [string]$Locale)
    $Yaml = @"
# yaml-language-server: `$schema=https://aka.ms/winget-manifest.version.$ManifestVersion.schema.json
PackageIdentifier: $PkgId
PackageVersion: $PkgVersion
DefaultLocale: $Locale
ManifestType: version
ManifestVersion: $ManifestVersion
"@
    return $Yaml
}

function ConvertTo-InstallerEntryYaml {
    <# Renders a single installer entry as YAML — shared by Build-InstallerYaml and Build-MergedManifestYaml #>
    param([hashtable]$Inst)
    $Yaml = "- Architecture: $($Inst.Architecture)"
    $Yaml += "`n  InstallerType: $($Inst.InstallerType)"
    $Yaml += "`n  InstallerUrl: $($Inst.InstallerUrl)"
    $Yaml += "`n  InstallerSha256: $($Inst.InstallerSha256)"
    if ($Inst.Scope) { $Yaml += "`n  Scope: $($Inst.Scope)" }
    if ($Inst.NestedInstallerType) {
        $Yaml += "`n  NestedInstallerType: $($Inst.NestedInstallerType)"
        if ($Inst.NestedInstallerFiles -and $Inst.NestedInstallerFiles.Count -gt 0) {
            $Yaml += "`n  NestedInstallerFiles:"
            foreach ($nf in $Inst.NestedInstallerFiles) {
                $Yaml += "`n  - RelativeFilePath: $($nf.RelativeFilePath)"
                if ($nf.PortableCommandAlias) { $Yaml += "`n    PortableCommandAlias: $($nf.PortableCommandAlias)" }
            }
        }
    }
    if ($Inst.InstallerSwitches) {
        $HasSwitches = $false
        foreach ($sw in @('Silent','SilentWithProgress','Interactive','Custom','InstallLocation','Log','Repair')) {
            if ($Inst.InstallerSwitches[$sw]) { $HasSwitches = $true; break }
        }
        if ($HasSwitches) {
            $Yaml += "`n  InstallerSwitches:"
            if ($Inst.InstallerSwitches.Silent)             { $Yaml += "`n    Silent: $(Format-YamlValue $Inst.InstallerSwitches.Silent)" }
            if ($Inst.InstallerSwitches.SilentWithProgress) { $Yaml += "`n    SilentWithProgress: $(Format-YamlValue $Inst.InstallerSwitches.SilentWithProgress)" }
            if ($Inst.InstallerSwitches.Interactive)        { $Yaml += "`n    Interactive: $(Format-YamlValue $Inst.InstallerSwitches.Interactive)" }
            if ($Inst.InstallerSwitches.Custom)             { $Yaml += "`n    Custom: $(Format-YamlValue $Inst.InstallerSwitches.Custom)" }
            if ($Inst.InstallerSwitches.InstallLocation)    { $Yaml += "`n    InstallLocation: $(Format-YamlValue $Inst.InstallerSwitches.InstallLocation)" }
            if ($Inst.InstallerSwitches.Log)                { $Yaml += "`n    Log: $(Format-YamlValue $Inst.InstallerSwitches.Log)" }
            if ($Inst.InstallerSwitches.Repair)             { $Yaml += "`n    Repair: $(Format-YamlValue $Inst.InstallerSwitches.Repair)" }
        }
    }
    if ($Inst.InstallLocationRequired) { $Yaml += "`n  InstallLocationRequired: true" }
    if ($Inst.MinimumOSVersion) { $Yaml += "`n  MinimumOSVersion: $($Inst.MinimumOSVersion)" }
    if ($Inst.Platform -and $Inst.Platform.Count -gt 0) {
        $Yaml += "`n  Platform:"
        foreach ($plat in $Inst.Platform) { $Yaml += "`n  - $plat" }
    }
    if ($Inst.ExpectedReturnCodes -and $Inst.ExpectedReturnCodes.Count -gt 0) {
        $Yaml += "`n  ExpectedReturnCodes:"
        foreach ($erc in $Inst.ExpectedReturnCodes) {
            $Yaml += "`n  - InstallerReturnCode: $($erc.InstallerReturnCode)"
            $Yaml += "`n    ReturnResponse: $($erc.ReturnResponse)"
        }
    }
    if ($Inst.AppsAndFeaturesEntries -and $Inst.AppsAndFeaturesEntries.Count -gt 0) {
        $Yaml += "`n  AppsAndFeaturesEntries:"
        foreach ($afe in $Inst.AppsAndFeaturesEntries) {
            $first = $true
            if ($afe.DisplayName)    { $Yaml += "`n  - DisplayName: $(Format-YamlValue $afe.DisplayName)"; $first = $false }
            if ($afe.Publisher)      { if ($first) { $Yaml += "`n  - Publisher: $(Format-YamlValue $afe.Publisher)"; $first = $false } else { $Yaml += "`n    Publisher: $(Format-YamlValue $afe.Publisher)" } }
            if ($afe.DisplayVersion) { if ($first) { $Yaml += "`n  - DisplayVersion: $(Format-YamlValue $afe.DisplayVersion)"; $first = $false } else { $Yaml += "`n    DisplayVersion: $(Format-YamlValue $afe.DisplayVersion)" } }
            if ($afe.ProductCode)    { if ($first) { $Yaml += "`n  - ProductCode: $(Format-YamlValue $afe.ProductCode)"; $first = $false } else { $Yaml += "`n    ProductCode: $(Format-YamlValue $afe.ProductCode)" } }
        }
    }
    if ($Inst.UpgradeBehavior)      { $Yaml += "`n  UpgradeBehavior: $($Inst.UpgradeBehavior)" }
    if ($Inst.ElevationRequirement) { $Yaml += "`n  ElevationRequirement: $($Inst.ElevationRequirement)" }
    if ($Inst.ProductCode)          { $Yaml += "`n  ProductCode: $(Format-YamlValue $Inst.ProductCode)" }
    if ($Inst.Commands -and $Inst.Commands.Count -gt 0) {
        $Yaml += "`n  Commands:"
        foreach ($cmd in $Inst.Commands) { $Yaml += "`n  - $cmd" }
    }
    if ($Inst.FileExtensions -and $Inst.FileExtensions.Count -gt 0) {
        $Yaml += "`n  FileExtensions:"
        foreach ($ext in $Inst.FileExtensions) { $Yaml += "`n  - $ext" }
    }
    if ($Inst.InstallModes -and $Inst.InstallModes.Count -gt 0) {
        $Yaml += "`n  InstallModes:"
        foreach ($mode in $Inst.InstallModes) { $Yaml += "`n  - $mode" }
    }
    return $Yaml
}

function Build-InstallerYaml {
    <# Generates the installer manifest YAML #>
    param([string]$PkgId, [string]$PkgVersion, [string]$ManifestVersion, [array]$Installers)
    $Yaml = @"
# yaml-language-server: `$schema=https://aka.ms/winget-manifest.installer.$ManifestVersion.schema.json
PackageIdentifier: $PkgId
PackageVersion: $PkgVersion
Installers:
"@
    foreach ($Inst in $Installers) {
        $Yaml += "`n$(ConvertTo-InstallerEntryYaml $Inst)"
    }

    # Dependencies (package-level)
    $DepYaml = Build-DependenciesYaml
    if ($DepYaml) { $Yaml += "`n$DepYaml" }

    $Yaml += "`nManifestType: installer"
    $Yaml += "`nManifestVersion: $ManifestVersion"
    return $Yaml
}

function Build-DefaultLocaleYaml {
    <# Generates the defaultLocale manifest YAML #>
    param([string]$PkgId, [string]$PkgVersion, [string]$ManifestVersion, [hashtable]$Meta)
    $Locale = if ($Meta.PackageLocale) { $Meta.PackageLocale } else { "en-US" }
    $Yaml = @"
# yaml-language-server: `$schema=https://aka.ms/winget-manifest.defaultLocale.$ManifestVersion.schema.json
PackageIdentifier: $PkgId
PackageVersion: $PkgVersion
PackageLocale: $Locale
"@
    if ($Meta.Publisher)         { $Yaml += "`nPublisher: $($Meta.Publisher)" }
    if ($Meta.PublisherUrl)      { $Yaml += "`nPublisherUrl: $($Meta.PublisherUrl)" }
    if ($Meta.PublisherSupportUrl) { $Yaml += "`nPublisherSupportUrl: $($Meta.PublisherSupportUrl)" }
    if ($Meta.PrivacyUrl)        { $Yaml += "`nPrivacyUrl: $($Meta.PrivacyUrl)" }
    if ($Meta.Author)            { $Yaml += "`nAuthor: $($Meta.Author)" }
    if ($Meta.PackageName)       { $Yaml += "`nPackageName: $($Meta.PackageName)" }
    if ($Meta.PackageUrl)        { $Yaml += "`nPackageUrl: $($Meta.PackageUrl)" }
    if ($Meta.License)           { $Yaml += "`nLicense: $($Meta.License)" }
    if ($Meta.LicenseUrl)        { $Yaml += "`nLicenseUrl: $($Meta.LicenseUrl)" }
    if ($Meta.Copyright)         { $Yaml += "`nCopyright: $($Meta.Copyright)" }
    if ($Meta.CopyrightUrl)      { $Yaml += "`nCopyrightUrl: $($Meta.CopyrightUrl)" }
    if ($Meta.ShortDescription)  {
        $sd = ($Meta.ShortDescription -replace '\r?\n', ' ').Trim()
        $Yaml += "`nShortDescription: $sd"
    }
    if ($Meta.Description) {
        $descText = $Meta.Description.TrimEnd()
        if ($descText -match '\r?\n') {
            $Yaml += "`nDescription: |"
            foreach ($dLine in ($descText -split '\r?\n')) {
                $Yaml += "`n  $($dLine.TrimEnd())"
            }
        } else {
            $Yaml += "`nDescription: $descText"
        }
    }
    if ($Meta.Moniker)           { $Yaml += "`nMoniker: $($Meta.Moniker)" }
    if ($Meta.Tags -and $Meta.Tags.Count -gt 0) {
        $Yaml += "`nTags:"
        foreach ($tag in $Meta.Tags) { $Yaml += "`n- $tag" }
    }
    if ($Meta.ReleaseNotes) {
        $Yaml += "`nReleaseNotes: |"
        foreach ($rnLine in ($Meta.ReleaseNotes -split "`n")) {
            $Yaml += "`n  $($rnLine.TrimEnd())"
        }
    }
    if ($Meta.ReleaseNotesUrl)   { $Yaml += "`nReleaseNotesUrl: $($Meta.ReleaseNotesUrl)" }
    $Yaml += "`nManifestType: defaultLocale"
    $Yaml += "`nManifestVersion: $ManifestVersion"
    return $Yaml
}

function Build-DependenciesYaml {
    <# Generates YAML for Dependencies section from UI fields #>
    $HasDeps = $false
    $Yaml = ''

    # Package Dependencies
    $PkgDepsText = $txtPackageDeps.Text.Trim()
    if ($PkgDepsText) {
        $Lines = $PkgDepsText -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        if ($Lines.Count -gt 0) {
            $HasDeps = $true
            $Yaml += "Dependencies:`n  PackageDependencies:"
            foreach ($Line in $Lines) {
                $Parts = $Line -split '\|' | ForEach-Object { $_.Trim() }
                $Yaml += "`n  - PackageIdentifier: $($Parts[0])"
                if ($Parts.Count -gt 1 -and $Parts[1]) { $Yaml += "`n    MinimumVersion: $($Parts[1])" }
            }
        }
    }

    # Windows Features
    $WinFeatures = $txtWindowsFeatures.Text.Trim()
    if ($WinFeatures) {
        $Items = $WinFeatures -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        if ($Items.Count -gt 0) {
            if (-not $HasDeps) { $Yaml += "Dependencies:"; $HasDeps = $true }
            $Yaml += "`n  WindowsFeatures:"
            foreach ($Item in $Items) { $Yaml += "`n  - $Item" }
        }
    }

    # Windows Libraries
    $WinLibs = $txtWindowsLibs.Text.Trim()
    if ($WinLibs) {
        $Items = $WinLibs -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        if ($Items.Count -gt 0) {
            if (-not $HasDeps) { $Yaml += "Dependencies:"; $HasDeps = $true }
            $Yaml += "`n  WindowsLibraries:"
            foreach ($Item in $Items) { $Yaml += "`n  - $Item" }
        }
    }

    # External Dependencies
    $ExtDeps = $txtExternalDeps.Text.Trim()
    if ($ExtDeps) {
        $Items = $ExtDeps -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        if ($Items.Count -gt 0) {
            if (-not $HasDeps) { $Yaml += "Dependencies:"; $HasDeps = $true }
            $Yaml += "`n  ExternalDependencies:"
            foreach ($Item in $Items) { $Yaml += "`n  - $Item" }
        }
    }

    if ($HasDeps) { return $Yaml } else { return '' }
}

function Build-UninstallYaml {
    <# Generates YAML for Uninstall info from UI fields #>
    $UninstallCmd = $txtUninstallCmd.Text.Trim()
    $UninstallSilent = $txtUninstallSilent.Text.Trim()
    if (-not $UninstallCmd -and -not $UninstallSilent) { return '' }

    $Yaml = ''
    if ($UninstallCmd) { $Yaml += "UninstallString: $(Format-YamlValue $UninstallCmd)" }
    if ($UninstallSilent) {
        if ($Yaml) { $Yaml += "`n" }
        $Yaml += "QuietUninstallString: $(Format-YamlValue "$UninstallCmd $UninstallSilent")"
    }
    return $Yaml
}

function Build-MergedManifestYaml {
    <# Generates a single 'singleton' manifest YAML for REST source publishing #>
    param([string]$PkgId, [string]$PkgVersion, [string]$ManifestVersion, [hashtable]$Meta, [array]$Installers)
    $Locale = if ($Meta.PackageLocale) { $Meta.PackageLocale } else { "en-US" }
    $Yaml = @"
# yaml-language-server: `$schema=https://aka.ms/winget-manifest.singleton.$ManifestVersion.schema.json
PackageIdentifier: $PkgId
PackageVersion: $PkgVersion
PackageLocale: $Locale
"@
    # ── Metadata (defaultLocale fields) ──
    if ($Meta.Publisher)            { $Yaml += "`nPublisher: $($Meta.Publisher)" }
    if ($Meta.PublisherUrl)         { $Yaml += "`nPublisherUrl: $($Meta.PublisherUrl)" }
    if ($Meta.PublisherSupportUrl)  { $Yaml += "`nPublisherSupportUrl: $($Meta.PublisherSupportUrl)" }
    if ($Meta.PrivacyUrl)           { $Yaml += "`nPrivacyUrl: $($Meta.PrivacyUrl)" }
    if ($Meta.Author)               { $Yaml += "`nAuthor: $($Meta.Author)" }
    if ($Meta.PackageName)          { $Yaml += "`nPackageName: $($Meta.PackageName)" }
    if ($Meta.PackageUrl)           { $Yaml += "`nPackageUrl: $($Meta.PackageUrl)" }
    if ($Meta.License)              { $Yaml += "`nLicense: $($Meta.License)" }
    if ($Meta.LicenseUrl)           { $Yaml += "`nLicenseUrl: $($Meta.LicenseUrl)" }
    if ($Meta.Copyright)            { $Yaml += "`nCopyright: $($Meta.Copyright)" }
    if ($Meta.CopyrightUrl)         { $Yaml += "`nCopyrightUrl: $($Meta.CopyrightUrl)" }
    if ($Meta.ShortDescription) {
        $sd = ($Meta.ShortDescription -replace '\r?\n', ' ').Trim()
        $Yaml += "`nShortDescription: $sd"
    }
    if ($Meta.Description) {
        $descText = $Meta.Description.TrimEnd()
        if ($descText -match '\r?\n') {
            $Yaml += "`nDescription: |"
            foreach ($dLine in ($descText -split '\r?\n')) {
                $Yaml += "`n  $($dLine.TrimEnd())"
            }
        } else {
            $Yaml += "`nDescription: $descText"
        }
    }
    if ($Meta.Moniker)              { $Yaml += "`nMoniker: $($Meta.Moniker)" }
    if ($Meta.Tags -and $Meta.Tags.Count -gt 0) {
        $Yaml += "`nTags:"
        foreach ($tag in $Meta.Tags) { $Yaml += "`n- $tag" }
    }
    if ($Meta.ReleaseNotesUrl)      { $Yaml += "`nReleaseNotesUrl: $($Meta.ReleaseNotesUrl)" }

    # ── Installers ──
    $Yaml += "`nInstallers:"
    foreach ($Inst in $Installers) {
        $Yaml += "`n$(ConvertTo-InstallerEntryYaml $Inst)"
    }

    # Dependencies
    $DepYaml = Build-DependenciesYaml
    if ($DepYaml) { $Yaml += "`n$DepYaml" }

    # Uninstall info
    $UninstYaml = Build-UninstallYaml
    if ($UninstYaml) { $Yaml += "`n$UninstYaml" }

    $Yaml += "`nManifestType: singleton"
    $Yaml += "`nManifestVersion: $ManifestVersion"
    return $Yaml
}

# ── Multi-Installer State ────────────────────────────────────────────────────
$Global:InstallerEntries = [System.Collections.Generic.List[hashtable]]::new()

function Get-AllInstallerData {
    <# Returns all installer entries: previously added + current card (if non-empty) #>
    $All = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($e in $Global:InstallerEntries) { $All.Add($e) }
    $Current = Get-InstallerDataFromUI
    # Only add current card if it has a URL or local hash (i.e., not blank)
    if ($Current.InstallerUrl -or $Global:CurrentHash0) {
        $All.Add($Current)
    }
    if ($All.Count -eq 0) { $All.Add($Current) }  # fallback: always at least one
    return @($All)
}

function Update-InstallerSummaryList {
    <# Refreshes the summary ListView showing added installer entries #>
    $lstInstallerEntries = $Window.FindName("lstInstallerEntries")
    $pnlInstallerSummary = $Window.FindName("pnlInstallerSummary")
    $btnRemoveLastInstaller = $Window.FindName("btnRemoveLastInstaller")
    $lstInstallerEntries.Items.Clear()
    if ($Global:InstallerEntries.Count -gt 0) {
        $pnlInstallerSummary.Visibility = 'Visible'
        $btnRemoveLastInstaller.Visibility = 'Visible'
        $btnEditInstaller.Visibility = 'Visible'
        $btnMoveInstallerUp.Visibility = if ($Global:InstallerEntries.Count -gt 1) { 'Visible' } else { 'Collapsed' }
        $btnMoveInstallerDown.Visibility = if ($Global:InstallerEntries.Count -gt 1) { 'Visible' } else { 'Collapsed' }
        $idx = 1
        foreach ($e in $Global:InstallerEntries) {
            $Desc = "#$idx  $($e.Architecture) / $($e.InstallerType)"
            if ($e.Scope) { $Desc += " / $($e.Scope)" }
            if ($e.InstallerUrl) { $Desc += " — $([System.IO.Path]::GetFileName(([uri]$e.InstallerUrl).LocalPath))" }
            $Item = New-StyledListViewItem -IconChar ([char]0xE7B8) -PrimaryText $Desc `
                -SubtitleText "SHA256: $(if($e.InstallerSha256){$e.InstallerSha256.Substring(0,16)+'...'}else{'(none)'})" -TagValue $e
            $lstInstallerEntries.Items.Add($Item) | Out-Null
            $idx++
        }
    } else {
        $pnlInstallerSummary.Visibility = 'Collapsed'
        $btnRemoveLastInstaller.Visibility = 'Collapsed'
        $btnEditInstaller.Visibility = 'Collapsed'
        $btnMoveInstallerUp.Visibility = 'Collapsed'
        $btnMoveInstallerDown.Visibility = 'Collapsed'
    }
    # Update batch panel visibility + multi-installer hint
    Update-BatchPanel
}

function Get-InstallerDataFromUI {
    <# Reads installer fields from the primary installer panel #>
    $Modes = @()
    if ($chkModeInteractive.IsChecked)    { $Modes += 'interactive' }
    if ($chkModeSilent.IsChecked)          { $Modes += 'silent' }
    if ($chkModeSilentProgress.IsChecked)  { $Modes += 'silentWithProgress' }

    $Inst = @{
        InstallerType      = Get-ComboBoxSelectedText $cmbInstallerType0
        Architecture       = Get-ComboBoxSelectedText $cmbArch0
        Scope              = Get-ComboBoxSelectedText $cmbScope0
        InstallerUrl       = $txtInstallerUrl0.Text.Trim()
        InstallerSha256    = if ($Global:CurrentHash0) { $Global:CurrentHash0 } else { "" }
        InstallerSwitches  = @{
            Silent             = $txtSilent0.Text.Trim()
            SilentWithProgress = $txtSilentProgress0.Text.Trim()
            Custom             = $txtCustomSwitch0.Text.Trim()
            Interactive        = $txtInteractive0.Text.Trim()
            Log                = $txtLog0.Text.Trim()
            Repair             = $txtRepair0.Text.Trim()
        }
        UpgradeBehavior      = Get-ComboBoxSelectedText $cmbUpgrade0
        ElevationRequirement = Get-ComboBoxSelectedText $cmbElevation0
        ProductCode          = $txtProductCode0.Text.Trim()
        Commands             = @()
        FileExtensions       = @()
        InstallModes         = $Modes
        InstallLocationRequired = [bool]$chkInstallLocationRequired.IsChecked
        MinimumOSVersion     = $txtMinOSVersion0.Text.Trim()
        Platform             = @()
        ExpectedReturnCodes  = @()
    }
    # Platform
    if ($chkPlatformDesktop.IsChecked)  { $Inst.Platform += 'Windows.Desktop' }
    if ($chkPlatformUniversal.IsChecked) { $Inst.Platform += 'Windows.Universal' }
    # ExpectedReturnCodes (parse "code:behavior, code:behavior")
    if ($txtExpectedReturnCodes0.Text.Trim()) {
        $Inst.ExpectedReturnCodes = @($txtExpectedReturnCodes0.Text.Trim() -split ',' | ForEach-Object {
            $parts = $_.Trim() -split ':'
            if ($parts.Count -ge 2) {
                @{ InstallerReturnCode = [int]$parts[0].Trim(); ReturnResponse = $parts[1].Trim() }
            }
        } | Where-Object { $_ })
    }
    # AppsAndFeaturesEntries
    $afEntry = @{}
    if ($txtAFDisplayName0.Text.Trim())    { $afEntry.DisplayName    = $txtAFDisplayName0.Text.Trim() }
    if ($txtAFPublisher0.Text.Trim())      { $afEntry.Publisher      = $txtAFPublisher0.Text.Trim() }
    if ($txtAFDisplayVersion0.Text.Trim()) { $afEntry.DisplayVersion = $txtAFDisplayVersion0.Text.Trim() }
    if ($txtAFProductCode0.Text.Trim())    { $afEntry.ProductCode    = $txtAFProductCode0.Text.Trim() }
    if ($afEntry.Count -gt 0) { $Inst.AppsAndFeaturesEntries = @($afEntry) }
    if ($txtCommands0.Text.Trim()) {
        $Inst.Commands = @($txtCommands0.Text.Trim() -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }
    if ($txtFileExt0.Text.Trim()) {
        $Inst.FileExtensions = @($txtFileExt0.Text.Trim() -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    # Nested installer (only when InstallerType = zip)
    if ($Inst.InstallerType -eq 'zip') {
        $NestedType = Get-ComboBoxSelectedText $cmbNestedType0
        $NestedPath = $txtNestedPath0.Text.Trim()
        $PortableAlias = $txtPortableAlias0.Text.Trim()
        if ($NestedType -and $NestedPath) {
            $Inst.NestedInstallerType = $NestedType
            $NestedFile = @{ RelativeFilePath = $NestedPath }
            if ($PortableAlias) { $NestedFile.PortableCommandAlias = $PortableAlias }
            $Inst.NestedInstallerFiles = @($NestedFile)
        }
    }
    return $Inst
}

function Get-MetadataFromUI {
    <# Reads defaultLocale metadata fields from the UI #>
    $Meta = @{
        PackageLocale      = "en-US"
        Publisher          = $txtPublisher.Text.Trim()
        PublisherUrl       = $txtPublisherUrl.Text.Trim()
        PublisherSupportUrl = $txtSupportUrl.Text.Trim()
        PrivacyUrl         = $txtPrivacyUrl.Text.Trim()
        Author             = $txtAuthor.Text.Trim()
        PackageName        = $txtPkgName.Text.Trim()
        PackageUrl         = $txtPackageUrl.Text.Trim()
        License            = $txtLicense.Text.Trim()
        LicenseUrl         = $txtLicenseUrl.Text.Trim()
        Copyright          = $txtCopyright.Text.Trim()
        ShortDescription   = $txtShortDesc.Text.Trim()
        Description        = $txtDescription.Text.Trim()
        Moniker            = $txtMoniker.Text.Trim()
        Tags               = @()
        ReleaseNotes       = $txtReleaseNotes.Text.Trim()
        ReleaseNotesUrl    = $txtReleaseNotesUrl.Text.Trim()
    }
    if ($txtTags.Text.Trim()) {
        $Meta.Tags = @($txtTags.Text.Trim() -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }
    return $Meta
}

function Update-YamlPreview {
    <# Regenerates the YAML preview in the Create tab #>
    $PkgId    = $txtPkgId.Text.Trim()
    $PkgVer   = $txtPkgVersion.Text.Trim()
    $ManVer   = Get-ComboBoxSelectedText $cmbCreateManifestVer
    if (-not $ManVer) { $ManVer = "1.9.0" }

    if (-not $PkgId -or -not $PkgVer) {
        $txtYamlPreview.Text = "# Fill in Package Identifier and Version to see preview..."
        return
    }

    $Meta = Get-MetadataFromUI
    $AllInstallers = Get-AllInstallerData

    $VersionYaml = Build-VersionYaml -PkgId $PkgId -PkgVersion $PkgVer -ManifestVersion $ManVer -Locale "en-US"
    $InstallerYaml = Build-InstallerYaml -PkgId $PkgId -PkgVersion $PkgVer -ManifestVersion $ManVer -Installers $AllInstallers
    $LocaleYaml = Build-DefaultLocaleYaml -PkgId $PkgId -PkgVersion $PkgVer -ManifestVersion $ManVer -Meta $Meta

    $Preview = @"
# ===== $PkgId.yaml (version) =====
$VersionYaml

# ===== $PkgId.installer.yaml =====
$InstallerYaml

# ===== $PkgId.locale.en-US.yaml (defaultLocale) =====
$LocaleYaml
"@
    $txtYamlPreview.Text = $Preview
    $null = Update-QualityMeter
}

# ==============================================================================
# SECTION 7B: MANIFEST EDITOR / LOADER
# ==============================================================================

function Load-ManifestFromConfig {
    <# Loads a package config JSON and populates the Edit tab #>
    param([string]$ConfigPath)
    if (-not (Test-Path $ConfigPath)) {
        Write-DebugLog "Config not found: $ConfigPath" -Level 'WARN'
        return
    }
    try {
        $Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
        $lblEditPkgId.Text = $Config.packageIdentifier
        $lblEditPkgVersion.Text = "v$($Config.packageVersion)"

        # Build YAML from config and display
        $PkgId  = $Config.packageIdentifier
        $PkgVer = $Config.packageVersion
        $ManVer = if ($Config.manifestVersion) { $Config.manifestVersion } else { "1.9.0" }

        $Meta = @{
            PackageLocale      = "en-US"
            Publisher          = $Config.defaultLocale.publisher
            PublisherUrl       = $Config.defaultLocale.publisherUrl
            PackageName        = $Config.defaultLocale.packageName
            PackageUrl         = $Config.defaultLocale.packageUrl
            License            = $Config.defaultLocale.license
            LicenseUrl         = $Config.defaultLocale.licenseUrl
            Copyright          = $Config.defaultLocale.copyright
            ShortDescription   = $Config.defaultLocale.shortDescription
            Description        = $Config.defaultLocale.description
            Moniker            = $Config.defaultLocale.moniker
            Tags               = @($Config.defaultLocale.tags)
            ReleaseNotes       = $Config.defaultLocale.releaseNotes
            ReleaseNotesUrl    = $Config.defaultLocale.releaseNotesUrl
            Author             = $Config.defaultLocale.author
            PublisherSupportUrl = $Config.defaultLocale.publisherSupportUrl
            PrivacyUrl         = $Config.defaultLocale.privacyUrl
        }

        $Installers = @()
        foreach ($inst in $Config.installers) {
            $InstData = @{
                Architecture       = $inst.architecture
                InstallerType      = $inst.installerType
                InstallerUrl       = $inst.installerUrl
                InstallerSha256    = $inst.installerSha256
                Scope              = $inst.scope
                UpgradeBehavior    = $inst.upgradeBehavior
                ElevationRequirement = $inst.elevationRequirement
                ProductCode        = $inst.productCode
                Commands           = @($inst.commands)
                FileExtensions     = @($inst.fileExtensions)
                InstallModes       = @($inst.installModes)
                InstallerSwitches  = @{
                    Silent             = $inst.installerSwitches.silent
                    SilentWithProgress = $inst.installerSwitches.silentWithProgress
                    Custom             = $inst.installerSwitches.custom
                    Interactive        = $inst.installerSwitches.interactive
                    Log                = $inst.installerSwitches.log
                    Repair             = $inst.installerSwitches.repair
                }
                InstallLocationRequired = [bool]$inst.installLocationRequired
                MinimumOSVersion   = $inst.minimumOSVersion
                Platform           = @($inst.platform)
                ExpectedReturnCodes = @()
                AppsAndFeaturesEntries = @()
            }
            if ($inst.expectedReturnCodes) {
                $InstData.ExpectedReturnCodes = @($inst.expectedReturnCodes | ForEach-Object {
                    @{ InstallerReturnCode = $_.installerReturnCode; ReturnResponse = $_.returnResponse }
                })
            }
            if ($inst.appsAndFeaturesEntries) {
                $InstData.AppsAndFeaturesEntries = @($inst.appsAndFeaturesEntries | ForEach-Object {
                    $afe = @{}
                    if ($_.displayName)    { $afe.DisplayName    = $_.displayName }
                    if ($_.publisher)      { $afe.Publisher      = $_.publisher }
                    if ($_.displayVersion) { $afe.DisplayVersion = $_.displayVersion }
                    if ($_.productCode)    { $afe.ProductCode    = $_.productCode }
                    $afe
                })
            }
            if ($inst.nestedInstallerType) {
                $InstData.NestedInstallerType = $inst.nestedInstallerType
                if ($inst.nestedInstallerFiles) {
                    $InstData.NestedInstallerFiles = @($inst.nestedInstallerFiles | ForEach-Object {
                        $nf = @{ RelativeFilePath = $_.relativeFilePath }
                        if ($_.portableCommandAlias) { $nf.PortableCommandAlias = $_.portableCommandAlias }
                        $nf
                    })
                }
            }
            $Installers += $InstData
        }

        $VersionYaml   = Build-VersionYaml -PkgId $PkgId -PkgVersion $PkgVer -ManifestVersion $ManVer -Locale "en-US"
        $InstallerYaml = Build-InstallerYaml -PkgId $PkgId -PkgVersion $PkgVer -ManifestVersion $ManVer -Installers $Installers
        $LocaleYaml    = Build-DefaultLocaleYaml -PkgId $PkgId -PkgVersion $PkgVer -ManifestVersion $ManVer -Meta $Meta

        $txtEditYaml.Text = "$VersionYaml`n`n$InstallerYaml`n`n$LocaleYaml"
        $pnlEditFields.Visibility = 'Visible'
        Write-DebugLog "Loaded config: $($Config.packageIdentifier) v$($Config.packageVersion)"
    } catch {
        Write-DebugLog "Error loading config: $($_.Exception.Message)" -Level 'ERROR'
    }
}

function Populate-CreateTabFromConfig {
    <# Loads a package config JSON and populates all Create tab fields for editing #>
    param([string]$ConfigPath)
    if (-not (Test-Path $ConfigPath)) {
        Write-DebugLog "Config not found: $ConfigPath" -Level 'WARN'
        return
    }
    try {
        $Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

        # Package identity
        $txtPkgId.Text      = $Config.packageIdentifier
        $txtPkgVersion.Text = $Config.packageVersion
        if ($Config.manifestVersion) { Set-ComboBoxByContent $cmbCreateManifestVer $Config.manifestVersion }

        # Metadata
        $dl = $Config.defaultLocale
        if ($dl) {
            $txtPkgName.Text       = if ($dl.packageName) { $dl.packageName } else { "" }
            $txtPublisher.Text     = if ($dl.publisher) { $dl.publisher } else { "" }
            $txtLicense.Text       = if ($dl.license) { $dl.license } else { "" }
            $txtShortDesc.Text     = if ($dl.shortDescription) { $dl.shortDescription } else { "" }
            $txtDescription.Text   = if ($dl.description) { $dl.description } else { "" }
            $txtAuthor.Text        = if ($dl.author) { $dl.author } else { "" }
            $txtMoniker.Text       = if ($dl.moniker) { $dl.moniker } else { "" }
            $txtPackageUrl.Text    = if ($dl.packageUrl) { $dl.packageUrl } else { "" }
            $txtPublisherUrl.Text  = if ($dl.publisherUrl) { $dl.publisherUrl } else { "" }
            $txtLicenseUrl.Text    = if ($dl.licenseUrl) { $dl.licenseUrl } else { "" }
            $txtSupportUrl.Text    = if ($dl.publisherSupportUrl) { $dl.publisherSupportUrl } else { "" }
            $txtCopyright.Text     = if ($dl.copyright) { $dl.copyright } else { "" }
            $txtPrivacyUrl.Text    = if ($dl.privacyUrl) { $dl.privacyUrl } else { "" }
            $txtReleaseNotes.Text   = if ($dl.releaseNotes) { $dl.releaseNotes } else { "" }
            $txtReleaseNotesUrl.Text = if ($dl.releaseNotesUrl) { $dl.releaseNotesUrl } else { "" }
            if ($dl.tags -and $dl.tags.Count -gt 0) {
                $txtTags.Text = ($dl.tags -join ', ')
            } else { $txtTags.Text = "" }
        }

        # First installer
        if ($Config.installers -and $Config.installers.Count -gt 0) {
            $inst = $Config.installers[0]
            Set-ComboBoxByContent $cmbInstallerType0 $inst.installerType
            Set-ComboBoxByContent $cmbArch0 $inst.architecture
            Set-ComboBoxByContent $cmbScope0 $inst.scope
            $txtInstallerUrl0.Text = if ($inst.installerUrl) { $inst.installerUrl } else { "" }
            if ($inst.installerSha256) { $Global:CurrentHash0 = $inst.installerSha256; $lblHash0.Text = "SHA256: $($inst.installerSha256)" }
            Set-ComboBoxByContent $cmbUpgrade0 $inst.upgradeBehavior
            Set-ComboBoxByContent $cmbElevation0 $inst.elevationRequirement
            $txtProductCode0.Text = if ($inst.productCode) { $inst.productCode } else { "" }
            $txtCommands0.Text = if ($inst.commands) { ($inst.commands -join ', ') } else { "" }
            $txtFileExt0.Text = if ($inst.fileExtensions) { ($inst.fileExtensions -join ', ') } else { "" }

            # Switches
            if ($inst.installerSwitches) {
                $txtSilent0.Text         = if ($inst.installerSwitches.silent) { $inst.installerSwitches.silent } else { "" }
                $txtSilentProgress0.Text = if ($inst.installerSwitches.silentWithProgress) { $inst.installerSwitches.silentWithProgress } else { "" }
                $txtCustomSwitch0.Text   = if ($inst.installerSwitches.custom) { $inst.installerSwitches.custom } else { "" }
                $txtInteractive0.Text    = if ($inst.installerSwitches.interactive) { $inst.installerSwitches.interactive } else { "" }
                $txtLog0.Text            = if ($inst.installerSwitches.log) { $inst.installerSwitches.log } else { "" }
                $txtRepair0.Text         = if ($inst.installerSwitches.repair) { $inst.installerSwitches.repair } else { "" }
            }

            # Install modes
            $chkModeInteractive.IsChecked    = $inst.installModes -and ('interactive' -in @($inst.installModes))
            $chkModeSilent.IsChecked          = $inst.installModes -and ('silent' -in @($inst.installModes))
            $chkModeSilentProgress.IsChecked  = $inst.installModes -and ('silentWithProgress' -in @($inst.installModes))
            $chkInstallLocationRequired.IsChecked = [bool]$inst.installLocationRequired

            # Platform & MinimumOSVersion
            $chkPlatformDesktop.IsChecked  = $inst.platform -and ('Windows.Desktop' -in @($inst.platform))
            $chkPlatformUniversal.IsChecked = $inst.platform -and ('Windows.Universal' -in @($inst.platform))
            $txtMinOSVersion0.Text = if ($inst.minimumOSVersion) { $inst.minimumOSVersion } else { "" }

            # ExpectedReturnCodes
            if ($inst.expectedReturnCodes -and $inst.expectedReturnCodes.Count -gt 0) {
                $txtExpectedReturnCodes0.Text = ($inst.expectedReturnCodes | ForEach-Object {
                    "$($_.installerReturnCode):$($_.returnResponse)"
                }) -join ', '
            } else {
                $txtExpectedReturnCodes0.Text = ""
            }

            # AppsAndFeaturesEntries
            if ($inst.appsAndFeaturesEntries -and $inst.appsAndFeaturesEntries.Count -gt 0) {
                $afe = $inst.appsAndFeaturesEntries[0]
                $txtAFDisplayName0.Text    = if ($afe.displayName)    { $afe.displayName }    else { "" }
                $txtAFPublisher0.Text      = if ($afe.publisher)      { $afe.publisher }      else { "" }
                $txtAFDisplayVersion0.Text = if ($afe.displayVersion) { $afe.displayVersion } else { "" }
                $txtAFProductCode0.Text    = if ($afe.productCode)    { $afe.productCode }    else { "" }
            } else {
                $txtAFDisplayName0.Text = ""
                $txtAFPublisher0.Text = ""
                $txtAFDisplayVersion0.Text = ""
                $txtAFProductCode0.Text = ""
            }

            # Nested installer
            if ($inst.nestedInstallerType) {
                Set-ComboBoxByContent $cmbNestedType0 $inst.nestedInstallerType
                $pnlNestedInstaller.Visibility = 'Visible'
                if ($inst.nestedInstallerFiles -and $inst.nestedInstallerFiles.Count -gt 0) {
                    $txtNestedPath0.Text     = if ($inst.nestedInstallerFiles[0].relativeFilePath) { $inst.nestedInstallerFiles[0].relativeFilePath } else { "" }
                    $txtPortableAlias0.Text  = if ($inst.nestedInstallerFiles[0].portableCommandAlias) { $inst.nestedInstallerFiles[0].portableCommandAlias } else { "" }
                }
            } else {
                $pnlNestedInstaller.Visibility = 'Collapsed'
                $txtNestedPath0.Text = ""
                $txtPortableAlias0.Text = ""
            }
        }

        # Load additional installers (beyond first) into InstallerEntries list
        $Global:InstallerEntries.Clear()
        if ($Config.installers -and $Config.installers.Count -gt 1) {
            for ($idx = 1; $idx -lt $Config.installers.Count; $idx++) {
                $ei = $Config.installers[$idx]
                $entry = @{
                    Architecture       = $ei.architecture
                    InstallerType      = $ei.installerType
                    InstallerUrl       = $ei.installerUrl
                    InstallerSha256    = $ei.installerSha256
                    Scope              = $ei.scope
                    UpgradeBehavior    = $ei.upgradeBehavior
                    ElevationRequirement = $ei.elevationRequirement
                    ProductCode        = $ei.productCode
                    Commands           = @($ei.commands)
                    FileExtensions     = @($ei.fileExtensions)
                    InstallModes       = @($ei.installModes)
                    InstallerSwitches  = @{
                        Silent             = $ei.installerSwitches.silent
                        SilentWithProgress = $ei.installerSwitches.silentWithProgress
                        Custom             = $ei.installerSwitches.custom
                        Interactive        = $ei.installerSwitches.interactive
                        Log                = $ei.installerSwitches.log
                        Repair             = $ei.installerSwitches.repair
                    }
                    InstallLocationRequired = [bool]$ei.installLocationRequired
                    MinimumOSVersion   = $ei.minimumOSVersion
                    Platform           = @($ei.platform)
                    ExpectedReturnCodes = @()
                    AppsAndFeaturesEntries = @()
                }
                if ($ei.expectedReturnCodes) {
                    $entry.ExpectedReturnCodes = @($ei.expectedReturnCodes | ForEach-Object {
                        @{ InstallerReturnCode = $_.installerReturnCode; ReturnResponse = $_.returnResponse }
                    })
                }
                if ($ei.appsAndFeaturesEntries) {
                    $entry.AppsAndFeaturesEntries = @($ei.appsAndFeaturesEntries | ForEach-Object {
                        $afe2 = @{}
                        if ($_.displayName)    { $afe2.DisplayName    = $_.displayName }
                        if ($_.publisher)      { $afe2.Publisher      = $_.publisher }
                        if ($_.displayVersion) { $afe2.DisplayVersion = $_.displayVersion }
                        if ($_.productCode)    { $afe2.ProductCode    = $_.productCode }
                        $afe2
                    })
                }
                if ($ei.nestedInstallerType) {
                    $entry.NestedInstallerType = $ei.nestedInstallerType
                    if ($ei.nestedInstallerFiles) {
                        $entry.NestedInstallerFiles = @($ei.nestedInstallerFiles | ForEach-Object {
                            $nf2 = @{ RelativeFilePath = $_.relativeFilePath }
                            if ($_.portableCommandAlias) { $nf2.PortableCommandAlias = $_.portableCommandAlias }
                            $nf2
                        })
                    }
                }
                $Global:InstallerEntries.Add($entry)
            }
        }
        Update-InstallerSummaryList

        # Reset upload fields first so stale values don't persist
        $txtBlobPath.Text = ''
        $txtUploadFile.Text = ''
        $lblUploadHash.Text = 'Select a file to compute hash...'
        $lblUploadProgress.Text = ''
        # Restore upload/storage tab values from config
        if ($Config.storage) {
            if ($Config.storage.blobPath) { $txtBlobPath.Text = $Config.storage.blobPath }
            if ($Config.storage.localFile -and (Test-Path $Config.storage.localFile)) {
                $txtUploadFile.Text = $Config.storage.localFile
            }
            # Auto-select storage account if it matches one in the combo
            if ($Config.storage.storageAccountName -and $cmbUploadStorage.Items.Count -gt 0) {
                for ($si = 0; $si -lt $cmbUploadStorage.Items.Count; $si++) {
                    if ($cmbUploadStorage.Items[$si].Tag -and $cmbUploadStorage.Items[$si].Tag.Name -eq $Config.storage.storageAccountName) {
                        $cmbUploadStorage.SelectedIndex = $si
                        break
                    }
                }
                # Container will be loaded async — store desired container to select after load
                if ($Config.storage.containerName) {
                    $Global:PendingContainerSelect = $Config.storage.containerName
                }
            }
        }

        # Reset preset to (none) since we loaded custom values
        $cmbInstallerPreset.SelectedIndex = 0

        Update-YamlPreview
        Write-DebugLog "Populated Create tab from config: $($Config.packageIdentifier) v$($Config.packageVersion)"
        $lblStatus.Text = "Loaded: $($Config.packageIdentifier) v$($Config.packageVersion)"
    } catch {
        Write-DebugLog "Error populating Create tab: $($_.Exception.Message)" -Level 'ERROR'
        $lblStatus.Text = "Error loading config"
    }
}

# ==============================================================================
# SECTION 7C: PACKAGE CONFIG (JSON) LOAD/SAVE
# ==============================================================================

$Global:ConfigDir = Join-Path $Global:Root "package_configs"
$Global:PendingContainerSelect = $null

function Save-PackageConfig {
    <# Saves current form data as a package config JSON file. Use -Force to skip overwrite confirmation. #>
    param([switch]$Force)
    $PkgId  = $txtPkgId.Text.Trim()
    $PkgVer = $txtPkgVersion.Text.Trim()
    if (-not $PkgId -or -not $PkgVer) {
        Write-DebugLog "Cannot save config: PackageIdentifier and Version are required" -Level 'WARN'
        $lblStatus.Text = "Cannot save: fill in Package ID and Version"
        return
    }

    if (-not (Test-Path $Global:ConfigDir)) {
        New-Item -Path $Global:ConfigDir -ItemType Directory -Force | Out-Null
    }

    # Check for existing config and prompt before overwriting
    $FileName = "$PkgId.json"
    $FilePath = Join-Path $Global:ConfigDir $FileName
    if (-not $Force -and (Test-Path $FilePath)) {
        $Confirm = Show-ThemedDialog -Title 'Overwrite Config' `
            -Message "A config for '$PkgId' already exists.`n`nOverwrite it with the current form data?" `
            -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' `
            -Buttons @(
                @{ Text='Cancel'; IsAccent=$false; Result='Cancel' }
                @{ Text='Overwrite'; IsAccent=$true; Result='Overwrite' }
            )
        if ($Confirm -ne 'Overwrite') {
            Write-DebugLog "Save-PackageConfig: user cancelled overwrite for $FileName"
            return
        }
    }

    $Meta = Get-MetadataFromUI
    $AllInstallers = Get-AllInstallerData

    $InstallerObjects = @($AllInstallers | ForEach-Object {
        $i = $_
        [PSCustomObject]@{
            architecture       = $i.Architecture
            installerType      = $i.InstallerType
            installerUrl       = $i.InstallerUrl
            installerSha256    = $i.InstallerSha256
            scope              = $i.Scope
            installModes       = $i.InstallModes
            installerSwitches  = [PSCustomObject]@{
                silent             = $i.InstallerSwitches.Silent
                silentWithProgress = $i.InstallerSwitches.SilentWithProgress
                custom             = $i.InstallerSwitches.Custom
                interactive        = $i.InstallerSwitches.Interactive
                log                = $i.InstallerSwitches.Log
                repair             = $i.InstallerSwitches.Repair
            }
            upgradeBehavior      = $i.UpgradeBehavior
            elevationRequirement = $i.ElevationRequirement
            productCode          = $i.ProductCode
            commands             = $i.Commands
            fileExtensions       = $i.FileExtensions
            installLocationRequired = $i.InstallLocationRequired
            minimumOSVersion     = $i.MinimumOSVersion
            platform             = $i.Platform
            expectedReturnCodes  = if ($i.ExpectedReturnCodes) { @($i.ExpectedReturnCodes | ForEach-Object { [PSCustomObject]$_ }) } else { $null }
            appsAndFeaturesEntries = if ($i.AppsAndFeaturesEntries) { @($i.AppsAndFeaturesEntries | ForEach-Object { [PSCustomObject]$_ }) } else { $null }
            nestedInstallerType  = if ($i.NestedInstallerType) { $i.NestedInstallerType } else { $null }
            nestedInstallerFiles = if ($i.NestedInstallerFiles) { @($i.NestedInstallerFiles | ForEach-Object { [PSCustomObject]$_ }) } else { $null }
        }
    })

    $Config = [PSCustomObject]@{
        '$schema'          = "WinGetManifestManager/1.0"
        packageIdentifier  = $PkgId
        packageVersion     = $PkgVer
        manifestVersion    = Get-ComboBoxSelectedText $cmbCreateManifestVer
        defaultLocale      = [PSCustomObject]@{
            packageLocale      = "en-US"
            publisher          = $Meta.Publisher
            publisherUrl       = $Meta.PublisherUrl
            publisherSupportUrl = $Meta.PublisherSupportUrl
            privacyUrl         = $Meta.PrivacyUrl
            author             = $Meta.Author
            packageName        = $Meta.PackageName
            packageUrl         = $Meta.PackageUrl
            license            = $Meta.License
            licenseUrl         = $Meta.LicenseUrl
            copyright          = $Meta.Copyright
            shortDescription   = $Meta.ShortDescription
            description        = $Meta.Description
            moniker            = $Meta.Moniker
            tags               = $Meta.Tags
            releaseNotes       = $Meta.ReleaseNotes
            releaseNotesUrl    = $Meta.ReleaseNotesUrl
        }
        installers = $InstallerObjects
        storage = [PSCustomObject]@{
            storageAccountName = if ($cmbUploadStorage.SelectedItem -and $cmbUploadStorage.SelectedItem.Tag) { $cmbUploadStorage.SelectedItem.Tag.Name } else { "" }
            containerName      = if ($cmbUploadContainer.SelectedItem) { $cmbUploadContainer.SelectedItem.Content.ToString() } else { "" }
            blobPath           = $txtBlobPath.Text.Trim()
            localFile          = $txtUploadFile.Text.Trim()
        }
        lastModified  = (Get-Date).ToString('o')
        lastPublished = $null
    }

    $Config | ConvertTo-Json -Depth 10 | Set-Content $FilePath -Force -Encoding UTF8
    Write-DebugLog "Saved package config: $FileName"
    $lblStatus.Text = "Config saved: $FileName"
    Refresh-ConfigList
    # Config count achievement
    $CfgCount = @(Get-ChildItem $Global:ConfigDir -Filter '*.json' -ErrorAction SilentlyContinue).Count
    if ($CfgCount -ge 10) { Unlock-Achievement 'ten_configs' }
    # Multi-arch achievement
    if ($InstallerObjects.Count -ge 5) { Unlock-Achievement 'multi_arch' }
}

function Refresh-ConfigList {
    <# Refreshes the left panel config list #>
    $lstConfigs.Items.Clear()
    if (-not (Test-Path $Global:ConfigDir)) {
        $lblConfigEmpty.Visibility = 'Visible'
        $lblConfigCount.Text = "0 configs"
        return
    }
    $Files = Get-ChildItem $Global:ConfigDir -Filter "*.json" -ErrorAction SilentlyContinue
    if ($Files.Count -eq 0) {
        $lblConfigEmpty.Visibility = 'Visible'
        $lblConfigCount.Text = "0 configs"
        return
    }
    $lblConfigEmpty.Visibility = 'Collapsed'
    foreach ($F in $Files) {
        $Item = New-StyledListViewItem -IconChar ([char]0xE7C3) -PrimaryText $F.BaseName `
            -SubtitleText ($F.LastWriteTime.ToString('yyyy-MM-dd HH:mm')) -TagValue $F.FullName
        $lstConfigs.Items.Add($Item) | Out-Null
    }
    $lblConfigCount.Text = "$($Files.Count) config$(if ($Files.Count -ne 1) {'s'})"
}

# ==============================================================================
# SECTION 7D: AZURE STORAGE BROWSER & UPLOADER
# ==============================================================================

$Global:StorageAccounts = @()
$Global:CurrentStorageContext = $null

function Discover-StorageAccounts {
    <# Discovers Azure Storage Accounts in background #>
    Write-DebugLog "Discovering storage accounts..."
    $lblStatus.Text = "Discovering storage accounts..."
    Start-BackgroundWork -Work {
        $Results = @()
        try {
            $Accounts = Get-AzStorageAccount -ErrorAction Stop
            foreach ($A in $Accounts) {
                $Results += [PSCustomObject]@{
                    Name          = $A.StorageAccountName
                    ResourceGroup = $A.ResourceGroupName
                    Location      = $A.PrimaryLocation
                    Kind          = $A.Kind
                    Sku           = $A.Sku.Name
                }
            }
        } catch {
            # Error handled in OnComplete
        }
        return $Results
    } -OnComplete {
        param($Results, $Errors)
        if ($Errors.Count -gt 0) {
            Write-DebugLog "Storage discovery error: $($Errors[0])" -Level 'ERROR'
        }
        $Accounts = if ($Results -and $Results.Count -gt 0) { $Results } else { @() }
        $Global:StorageAccounts = $Accounts
        Write-DebugLog "Found $($Accounts.Count) storage account(s)"

        # Populate storage list
        $tvStorage.Items.Clear()
        $cmbUploadStorage.Items.Clear()
        if ($Accounts.Count -eq 0) {
            $lblStorageEmpty.Text = "No storage accounts found"
            $lblStorageEmpty.Visibility = 'Visible'
        } else {
            $lblStorageEmpty.Visibility = 'Collapsed'
            foreach ($A in $Accounts) {
                # Build styled card item with cloud icon + name + region
                $Item = New-StyledListViewItem -IconChar ([char]0xE753) -PrimaryText $A.Name `
                    -SubtitleText $A.Location -TagValue $A
                $tvStorage.Items.Add($Item) | Out-Null

                # Upload combo item
                $CmbItem = New-Object System.Windows.Controls.ComboBoxItem
                $CmbItem.Content = "$($A.Name) ($($A.Location))"
                $CmbItem.Tag = $A
                $cmbUploadStorage.Items.Add($CmbItem) | Out-Null
            }
            # Auto-select preferred storage account (or first) to trigger container loading
            $MatchIdx = -1
            if ($Global:PreferredStorageAccount) {
                for ($si = 0; $si -lt $cmbUploadStorage.Items.Count; $si++) {
                    if ($cmbUploadStorage.Items[$si].Tag.Name -eq $Global:PreferredStorageAccount) {
                        $MatchIdx = $si; break
                    }
                }
            }
            $cmbUploadStorage.SelectedIndex = if ($MatchIdx -ge 0) { $MatchIdx } else { 0 }
        }
        $lblStatus.Text = "Found $($Accounts.Count) storage account(s)"
        # Refresh backup/restore storage combos with new data
        if (Get-Command Populate-BackupRestoreCombos -ErrorAction SilentlyContinue) { Populate-BackupRestoreCombos }
    }
}

function Discover-RestSources {
    <# Discovers WinGet REST source Function Apps, Key Vaults, Storage Accounts, and APIM instances
       in background. Populates the source combo and all Remove tab override ComboBoxes.
       Uses Get-AzResource (Az.Resources) to avoid extra module dependencies. #>
    Write-DebugLog "Discovering REST sources and related Azure resources..."
    $DiscPattern = if ($txtDiscoveryPattern -and $txtDiscoveryPattern.Text) { $txtDiscoveryPattern.Text } else { '^func-' }
    Start-BackgroundWork -Variables @{ DiscoveryPattern = $DiscPattern } -Work {
        $WinGetFuncApps = @()
        $AllFuncApps    = @()
        $KeyVaults      = @()
        $StorageAccts   = @()
        $APIMs          = @()
        try {
            # All Function Apps (for override combo) and WinGet-filtered subset (for source combo)
            $RawFuncApps = @(Get-AzResource -ResourceType 'Microsoft.Web/sites' -ErrorAction Stop |
                Where-Object { $_.Kind -and $_.Kind -match 'functionapp' })
            foreach ($FA in $RawFuncApps) {
                $Obj = [PSCustomObject]@{
                    Name            = $FA.Name
                    ResourceGroup   = $FA.ResourceGroupName
                    Location        = $FA.Location
                    DefaultHostName = "$($FA.Name).azurewebsites.net"
                }
                $AllFuncApps += $Obj
                if ($FA.Name -match 'winget' -or $FA.Name -match $DiscoveryPattern -or
                    ($FA.Tags -and ($FA.Tags.ContainsKey('WinGetSource') -or $FA.Tags.ContainsKey('winget-restsource')))) {
                    $WinGetFuncApps += $Obj
                }
            }
        } catch { }
        try {
            $KeyVaults = @(Get-AzResource -ResourceType 'Microsoft.KeyVault/vaults' -ErrorAction Stop | ForEach-Object {
                [PSCustomObject]@{ Name = $_.Name; ResourceGroup = $_.ResourceGroupName; Location = $_.Location }
            })
        } catch { }
        try {
            $StorageAccts = @(Get-AzResource -ResourceType 'Microsoft.Storage/storageAccounts' -ErrorAction Stop | ForEach-Object {
                [PSCustomObject]@{ Name = $_.Name; ResourceGroup = $_.ResourceGroupName; Location = $_.Location }
            })
        } catch { }
        try {
            $APIMs = @(Get-AzResource -ResourceType 'Microsoft.ApiManagement/service' -ErrorAction Stop | ForEach-Object {
                [PSCustomObject]@{ Name = $_.Name; ResourceGroup = $_.ResourceGroupName; Location = $_.Location }
            })
        } catch { }
        return @{
            WinGetFuncApps = $WinGetFuncApps
            AllFuncApps    = $AllFuncApps
            KeyVaults      = $KeyVaults
            StorageAccounts = $StorageAccts
            APIMs          = $APIMs
        }
    } -OnComplete {
        param($Results, $Errors)
        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ WinGetFuncApps = @(); AllFuncApps = @(); KeyVaults = @(); StorageAccounts = @(); APIMs = @() } }
        $FuncApps = @($R.WinGetFuncApps)
        Write-DebugLog "Discovered: $($FuncApps.Count) REST sources, $(@($R.AllFuncApps).Count) FuncApps, $(@($R.KeyVaults).Count) KeyVaults, $(@($R.StorageAccounts).Count) Storage, $(@($R.APIMs).Count) APIM"
        Populate-RestSourceCombo -FuncApps $FuncApps
        Populate-RemoveSourceCombo -FuncApps $FuncApps
        Populate-RemoveOverrideCombos -FuncApps @($R.AllFuncApps) -KeyVaults @($R.KeyVaults) -StorageAccounts @($R.StorageAccounts) -APIMs @($R.APIMs)
        if (Get-Command Populate-BackupRestoreCombos -ErrorAction SilentlyContinue) { Populate-BackupRestoreCombos }
    }
}

# --- Global Progress Bar helpers (Bagel Commander style) ---------------------
function Show-GlobalProgress {
    <# Shows the global progress bar below the tab menu with value and label #>
    param([int]$Value = 0, [int]$Maximum = 100, [string]$Text = '')
    $prgGlobal.Maximum = $Maximum
    $prgGlobal.Value   = $Value
    $prgGlobal.IsIndeterminate = $false
    $prgGlobal.Visibility = 'Visible'
    $brdGlobalShimmer.Visibility = 'Collapsed'
    $lblGlobalProgress.Text = $Text
    $pnlGlobalProgress.Visibility = 'Visible'
}

function Show-GlobalProgressIndeterminate {
    <# Shows the global progress bar in indeterminate (shimmer) mode #>
    param([string]$Text = '')
    $prgGlobal.Visibility = 'Collapsed'
    $brdGlobalShimmer.Visibility = 'Visible'
    $lblGlobalProgress.Text = $Text
    $pnlGlobalProgress.Visibility = 'Visible'
}

function Update-GlobalProgress {
    <# Updates progress bar value and label text #>
    param([int]$Value, [string]$Text)
    $prgGlobal.Value = $Value
    if ($Text) { $lblGlobalProgress.Text = $Text }
}

function Hide-GlobalProgress {
    <# Hides the global progress bar #>
    $pnlGlobalProgress.Visibility = 'Collapsed'
    $prgGlobal.Value = 0
    $prgGlobal.IsIndeterminate = $false
    $prgGlobal.Visibility = 'Collapsed'
    $brdGlobalShimmer.Visibility = 'Collapsed'
    $lblGlobalProgress.Text = ''
}

# --- Remove tab ComboBox population helpers ----------------------------------
function Populate-RemoveSourceCombo {
    <# Populates the Remove tab source name ComboBox from discovered Function Apps #>
    param([array]$FuncApps = @())
    Write-DebugLog "Populate-RemoveSourceCombo: received $($FuncApps.Count) FuncApp(s)" -Level 'DEBUG'
    if ($null -eq $cmbRemoveName) { return }
    $Current = $cmbRemoveName.Text
    $cmbRemoveName.Items.Clear()
    foreach ($FA in $FuncApps) {
        $SourceName = $FA.Name -replace '^func-', ''
        $Item = New-Object System.Windows.Controls.ComboBoxItem
        $Item.Content = $SourceName
        $Item.Tag = $FA
        $cmbRemoveName.Items.Add($Item) | Out-Null
    }
    if ($Current) { $cmbRemoveName.Text = $Current }
}

function Populate-RemoveRGCombo {
    <# Populates the Remove tab resource group ComboBox from discovered RGs #>
    param([array]$ResourceGroups = @())
    Write-DebugLog "Populate-RemoveRGCombo: received $($ResourceGroups.Count) RG(s)" -Level 'DEBUG'
    if ($null -eq $cmbRemoveRG) { return }
    $Current = $cmbRemoveRG.Text
    $cmbRemoveRG.Items.Clear()
    foreach ($RG in $ResourceGroups) {
        $Item = New-Object System.Windows.Controls.ComboBoxItem
        $Item.Content = $RG.Name
        $Item.Tag = $RG
        $cmbRemoveRG.Items.Add($Item) | Out-Null
    }
    if ($Current) { $cmbRemoveRG.Text = $Current }
}

function Populate-RemoveOverrideCombos {
    <# Populates the four Remove override ComboBoxes from discovered Azure resources #>
    param(
        [array]$FuncApps       = @(),
        [array]$KeyVaults      = @(),
        [array]$StorageAccounts = @(),
        [array]$APIMs          = @()
    )
    # Function Apps
    if ($cmbRemoveOverrideFuncApp) {
        $Cur = $cmbRemoveOverrideFuncApp.Text
        $cmbRemoveOverrideFuncApp.Items.Clear()
        foreach ($R in $FuncApps) {
            $Item = New-Object System.Windows.Controls.ComboBoxItem
            $Item.Content = $R.Name
            $Item.Tag = $R
            $cmbRemoveOverrideFuncApp.Items.Add($Item) | Out-Null
        }
        if ($Cur) { $cmbRemoveOverrideFuncApp.Text = $Cur }
    }
    # Key Vaults
    if ($cmbRemoveOverrideKeyVault) {
        $Cur = $cmbRemoveOverrideKeyVault.Text
        $cmbRemoveOverrideKeyVault.Items.Clear()
        foreach ($R in $KeyVaults) {
            $Item = New-Object System.Windows.Controls.ComboBoxItem
            $Item.Content = $R.Name
            $Item.Tag = $R
            $cmbRemoveOverrideKeyVault.Items.Add($Item) | Out-Null
        }
        if ($Cur) { $cmbRemoveOverrideKeyVault.Text = $Cur }
    }
    # Storage Accounts
    if ($cmbRemoveOverrideStorage) {
        $Cur = $cmbRemoveOverrideStorage.Text
        $cmbRemoveOverrideStorage.Items.Clear()
        foreach ($R in $StorageAccounts) {
            $Item = New-Object System.Windows.Controls.ComboBoxItem
            $Item.Content = $R.Name
            $Item.Tag = $R
            $cmbRemoveOverrideStorage.Items.Add($Item) | Out-Null
        }
        if ($Cur) { $cmbRemoveOverrideStorage.Text = $Cur }
    }
    # APIM
    if ($cmbRemoveOverrideAPIM) {
        $Cur = $cmbRemoveOverrideAPIM.Text
        $cmbRemoveOverrideAPIM.Items.Clear()
        foreach ($R in $APIMs) {
            $Item = New-Object System.Windows.Controls.ComboBoxItem
            $Item.Content = $R.Name
            $Item.Tag = $R
            $cmbRemoveOverrideAPIM.Items.Add($Item) | Out-Null
        }
        if ($Cur) { $cmbRemoveOverrideAPIM.Text = $Cur }
    }
    Write-DebugLog "Populate-RemoveOverrideCombos: FA=$($FuncApps.Count) KV=$($KeyVaults.Count) ST=$($StorageAccounts.Count) APIM=$($APIMs.Count)" -Level 'DEBUG'
}

# Left panel storage list → sync selection to upload combo (triggers container + info refresh)
$tvStorage.Add_SelectionChanged({
    $Sel = $tvStorage.SelectedItem
    if (-not $Sel -or -not $Sel.Tag) { return }
    $SelName = $Sel.Tag.Name
    for ($i = 0; $i -lt $cmbUploadStorage.Items.Count; $i++) {
        if ($cmbUploadStorage.Items[$i].Tag -and $cmbUploadStorage.Items[$i].Tag.Name -eq $SelName) {
            if ($cmbUploadStorage.SelectedIndex -ne $i) {
                $cmbUploadStorage.SelectedIndex = $i
            }
            break
        }
    }
    # Ensure Storage section is visible on left panel
    Show-LeftSection 'Storage'
})

# ==============================================================================
# SECTION 7E: STORAGE ACCOUNT PROVISIONER
# ==============================================================================

# Element references for new storage overlay
$pnlNewStorage        = $Window.FindName("pnlNewStorage")
$btnCloseNewStorage   = $Window.FindName("btnCloseNewStorage")
$cmbNewStorageRG      = $Window.FindName("cmbNewStorageRG")
$btnRefreshRGs        = $Window.FindName("btnRefreshRGs")
$txtNewStorageName    = $Window.FindName("txtNewStorageName")
$lblStorageNameHint   = $Window.FindName("lblStorageNameHint")
$cmbNewStorageLocation = $Window.FindName("cmbNewStorageLocation")
$cmbNewStorageSku     = $Window.FindName("cmbNewStorageSku")
$cmbNewStorageKind    = $Window.FindName("cmbNewStorageKind")
$chkNewStoragePublicAccess = $Window.FindName("chkNewStoragePublicAccess")
$lblNewStorageStatus  = $Window.FindName("lblNewStorageStatus")
$prgNewStorage        = $Window.FindName("prgNewStorage")
$btnCancelNewStorage  = $Window.FindName("btnCancelNewStorage")
$btnCreateStorage     = $Window.FindName("btnCreateStorage")

function Show-NewStoragePanel {
    <# Opens the storage creation overlay and populates resource groups #>
    $pnlNewStorage.Visibility = 'Visible'
    $txtNewStorageName.Text = ""
    $lblNewStorageStatus.Text = ""
    $prgNewStorage.Visibility = 'Collapsed'
    $btnCreateStorage.IsEnabled = $true
    Refresh-ResourceGroups
}

function Hide-NewStoragePanel {
    <# Closes the storage creation overlay #>
    $pnlNewStorage.Visibility = 'Collapsed'
}

function Refresh-ResourceGroups {
    <# Discovers resource groups and populates the RG combo box #>
    Write-DebugLog "Discovering resource groups..."
    Start-BackgroundWork -Work {
        try {
            $RGs = @(Get-AzResourceGroup -ErrorAction Stop | Sort-Object ResourceGroupName | ForEach-Object {
                [PSCustomObject]@{ Name = $_.ResourceGroupName; Location = $_.Location }
            })
            return @{ RGs = $RGs; Error = $null }
        } catch {
            return @{ RGs = @(); Error = $_.Exception.Message }
        }
    } -OnComplete {
        param($Results, $Errors)
        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ RGs = @() } }
        $cmbNewStorageRG.Items.Clear()
        foreach ($RG in $R.RGs) {
            $Item = New-Object System.Windows.Controls.ComboBoxItem
            $Item.Content = $RG.Name
            $Item.Tag = $RG
            $cmbNewStorageRG.Items.Add($Item) | Out-Null
        }
        if ($R.RGs.Count -gt 0) { $cmbNewStorageRG.SelectedIndex = 0 }
        # Also populate Remove tab RG combo
        Populate-RemoveRGCombo -ResourceGroups $R.RGs
        Write-DebugLog "Found $($R.RGs.Count) resource group(s)"
    }
}

# Function source code (as STRING) for the current-user OID resolver. Background
# runspaces use CreateDefault() ISS so they DON'T inherit script-scope functions.
# We can't pass a scriptblock either — PowerShell scriptblocks carry their original
# session state with them, and invoking a main-thread-bound scriptblock from a
# background runspace causes concurrent access to the main session state which
# corrupts WPF/dispatcher scope tracking ('Global scope cannot be removed').
# So instead, each background runspace recreates the function locally via
# `Invoke-Expression $ResolveOidFnSrc` before calling Resolve-CurrentUserObjectId.
#
# Robust against guest users (B2B): for a guest signed in to a customer's tenant,
# the email-style UPN ('alice@contoso.com') won't match the tenant directory —
# the actual UPN there is 'alice_contoso.com#EXT#@<tenant>.onmicrosoft.com'.
# Strategy:
#   1. Get-AzADUser -SignedIn  (Az.Resources >= 6.x; tenant-aware, handles guests)
#   2. Microsoft Graph /me
#   3. Get-AzADUser -UserPrincipalName <email>  (member users)
#   4. Get-AzADUser -Mail <email>
# Returns $null if all strategies fail.
$Script:ResolveOidFnSrc = @'
function Resolve-CurrentUserObjectId {
    param([string]$Email)
    if (-not $Email) { try { $Email = (Get-AzContext).Account.Id } catch {} }
    try {
        $u = Get-AzADUser -SignedIn -ErrorAction Stop
        if ($u -and $u.Id) { return $u.Id }
    } catch {}
    try {
        $tok = (Get-AzAccessToken -ResourceUrl 'https://graph.microsoft.com/' -ErrorAction Stop).Token
        if ($tok) {
            $me = Invoke-RestMethod -Uri 'https://graph.microsoft.com/v1.0/me' -Headers @{ Authorization = "Bearer $tok" } -ErrorAction Stop
            if ($me -and $me.id) { return $me.id }
        }
    } catch {}
    if ($Email) {
        try {
            $u = Get-AzADUser -UserPrincipalName $Email -ErrorAction Stop
            if ($u -and $u.Id) { return $u.Id }
        } catch {}
        try {
            $u = Get-AzADUser -Mail $Email -ErrorAction Stop
            if ($u -and $u.Id) { return $u.Id }
        } catch {}
    }
    return $null
}
'@

# Main-thread copy of the function (define once for any main-thread callers).
Invoke-Expression $Script:ResolveOidFnSrc

function Set-StoragePublicAccess {
    <# Toggles anonymous public read access on a storage account + container.
       Sets account-level AllowBlobPublicAccess and container ACL ('Blob' = anonymous read on
       blobs only; 'Off' = private). Requires Storage Account Contributor (control plane) for
       the account flag and either an account key or Storage Blob Data Contributor (data plane)
       for the container ACL. Refreshes the storage info panel on completion. #>
    param(
        [Parameter(Mandatory)][string]$StorageAccountName,
        [Parameter(Mandatory)][string]$ResourceGroupName,
        [Parameter(Mandatory)][string]$ContainerName,
        [Parameter(Mandatory)][ValidateSet('Enable','Disable')][string]$Mode
    )
    $Enable = ($Mode -eq 'Enable')
    $Verb   = if ($Enable) { 'Enabling' } else { 'Disabling' }
    Write-DebugLog "$Verb public read access on $StorageAccountName/$ContainerName..."
    $lblStatus.Text = "$Verb public access on $StorageAccountName/$ContainerName..."
    Show-GlobalProgressIndeterminate -Text "$Verb public read access..."

    Start-BackgroundWork -Variables @{
        AcctName  = $StorageAccountName
        RG        = $ResourceGroupName
        Container = $ContainerName
        Enable    = $Enable
    } -Work {
        $Result = @{ Success = $false; Steps = @(); Error = $null }
        try {
            # Step 1: account-level AllowBlobPublicAccess
            Set-AzStorageAccount -ResourceGroupName $RG -Name $AcctName -AllowBlobPublicAccess $Enable -ErrorAction Stop | Out-Null
            $Result.Steps += "Account flag AllowBlobPublicAccess set to $Enable"

            # Step 2: container ACL
            $Ctx = New-AzStorageContext -StorageAccountName $AcctName -UseConnectedAccount -ErrorAction Stop
            $AclMode = if ($Enable) { 'Blob' } else { 'Off' }
            Set-AzStorageContainerAcl -Name $Container -Context $Ctx -Permission $AclMode -ErrorAction Stop | Out-Null
            $Result.Steps += "Container '$Container' ACL set to $AclMode"

            $Result.Success = $true
            return $Result
        } catch {
            $Result.Error = $_.Exception.Message
            return $Result
        }
    } -OnComplete {
        param($Results, $Errors)
        Hide-GlobalProgress
        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { $null }
        if ($R -and $R.Success) {
            foreach ($Step in $R.Steps) { Write-DebugLog "  $Step" -Level 'SUCCESS' }
            $lblStatus.Text = if ($Enable) { "Public read access ENABLED — installers downloadable anonymously" } else { "Public read access DISABLED" }
            # Refresh the storage info panel by re-triggering the storage selection
            $SelIdx = $cmbUploadStorage.SelectedIndex
            if ($SelIdx -ge 0) {
                $cmbUploadStorage.SelectedIndex = -1
                $cmbUploadStorage.SelectedIndex = $SelIdx
            }
            if (Get-Command Show-Toast -ErrorAction SilentlyContinue) {
                $Type = if ($Enable) { 'Warning' } else { 'Success' }
                $Msg  = if ($Enable) { "Public read enabled — REMOVE before going to production" } else { "Public read disabled" }
                Show-Toast -Message $Msg -Type $Type
            }
        } else {
            $ErrMsg = if ($R) { $R.Error } elseif ($Errors -and $Errors.Count -gt 0) { $Errors[0].ToString() } else { 'Unknown error' }
            Write-DebugLog "Public access toggle failed: $ErrMsg" -Level 'ERROR'
            $lblStatus.Text = "Public access toggle failed"
            Show-ThemedDialog -Title 'Operation Failed' -Message "Failed to update public access:`n`n$ErrMsg" `
                -Icon ([string]([char]0xEA39)) -IconColor 'ThemeError' | Out-Null
        }
    }
}

function Assign-StorageBlobRole {
    <# Assigns Storage Blob Data Contributor role to the current user on a storage account,
       then polls until the role has propagated (Azure RBAC can take up to 5-10 min). #>
    param(
        [string]$StorageAccountId,
        [string]$StorageAccountName
    )
    Write-DebugLog "Assigning Storage Blob Data Contributor role..."
    $lblStatus.Text = "Assigning storage role..."
    $statusDot.Fill = $Window.Resources['ThemeWarning']

    Start-BackgroundWork -Variables @{ SaId = $StorageAccountId; ResolveOidFnSrc = $Script:ResolveOidFnSrc } -Work {
        try {
            # Recreate the helper inside this runspace (avoid cross-thread session state).
            Invoke-Expression $ResolveOidFnSrc
            $CurrentUser = (Get-AzContext).Account.Id
            $ObjectId = Resolve-CurrentUserObjectId -Email $CurrentUser
            if (-not $ObjectId) {
                throw "Could not resolve directory ObjectId for '$CurrentUser' (tried -SignedIn, Graph /me, UPN, and Mail). If you're a guest user, ensure you're consented to the customer tenant."
            }
            New-AzRoleAssignment -ObjectId $ObjectId -RoleDefinitionName 'Storage Blob Data Contributor' -Scope $SaId -ErrorAction Stop | Out-Null
            return @{ Success = $true; Error = $null }
        } catch {
            # If role already exists, treat as success
            if ($_.Exception.Message -match 'RoleAssignmentExists|Conflict') {
                return @{ Success = $true; Error = $null }
            }
            return @{ Success = $false; Error = $_.Exception.Message }
        }
    } -OnComplete {
        param($Results, $Errors)
        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ Success = $false; Error = 'No result' } }
        if ($R.Success) {
            Write-DebugLog "Role assignment created — waiting for propagation..."
            $lblStatus.Text = "Role assigned — waiting for propagation (this can take a few minutes)..."
            # Start propagation polling
            Poll-RolePropagation -StorageAccountId $StorageAccountId -StorageAccountName $StorageAccountName
        } else {
            Write-DebugLog "Role assignment failed: $($R.Error)" -Level 'ERROR'
            $lblStatus.Text = "Role assignment failed — assign manually"
            $statusDot.Fill = $Window.Resources['ThemeError']
        }
    }
}

function Poll-RolePropagation {
    <# Polls Azure storage with OAuth to verify the RBAC role has propagated.
       Attempts every 15 seconds for up to 5 minutes (20 attempts). #>
    param(
        [string]$StorageAccountId,
        [string]$StorageAccountName
    )
    # Extract account name from ARM id if not provided
    if (-not $StorageAccountName -and $StorageAccountId -match '/storageAccounts/(.+)$') {
        $StorageAccountName = $Matches[1]
    }
    $Global:RolePollAttempt = 0
    $Global:RolePollMaxAttempts = $Script:RBAC_MAX_POLLS
    $Global:RolePollAcctName = $StorageAccountName
    $Global:RolePollAcctId = $StorageAccountId
    $Global:RolePollActive = $true

    Write-DebugLog "Starting role propagation poll for $StorageAccountName (max $($Global:RolePollMaxAttempts) attempts, 15s interval)" -Level 'DEBUG'
    $lblStatus.Text = "Waiting for role propagation... (attempt 1/$($Global:RolePollMaxAttempts))"

    # Launch first check immediately
    Start-RolePropagationCheck
}

function Start-RolePropagationCheck {
    <# Launches a single background role-propagation check #>
    if (-not $Global:RolePollActive) { return }
    $Global:RolePollAttempt++
    $Attempt = $Global:RolePollAttempt
    $MaxAttempts = $Global:RolePollMaxAttempts
    $AcctName = $Global:RolePollAcctName

    Write-DebugLog "Role propagation check attempt $Attempt/$MaxAttempts for $AcctName" -Level 'DEBUG'
    $lblStatus.Text = "Waiting for role propagation... (attempt $Attempt/$MaxAttempts)"

    Start-BackgroundWork -Variables @{
        AccName     = $AcctName
        AttemptNum  = $Attempt
        MaxAttempts = $MaxAttempts
        PollSec     = $Script:RBAC_POLL_SEC
    } -Work {
        # Wait 15 seconds between checks (except first attempt)
        if ($AttemptNum -gt 1) {
            $sleepSec = if ($PollSec -and $PollSec -gt 0) { [int]$PollSec } else { 15 }
            Start-Sleep -Seconds $sleepSec
        }

        try {
            $Ctx = New-AzStorageContext -StorageAccountName $AccName -UseConnectedAccount -ErrorAction Stop
            # Try listing containers — this requires the data-plane role
            $null = Get-AzStorageContainer -Context $Ctx -MaxCount 1 -ErrorAction Stop
            return @{ Propagated = $true; Attempt = $AttemptNum; Error = $null }
        } catch {
            if ($_.Exception.Message -match 'AuthorizationPermissionMismatch|403|AuthenticationFailed') {
                return @{ Propagated = $false; Attempt = $AttemptNum; Error = $_.Exception.Message }
            }
            # Other errors (network, etc.) — treat as not propagated but log
            return @{ Propagated = $false; Attempt = $AttemptNum; Error = $_.Exception.Message }
        }
    } -OnComplete {
        param($Results, $Errors)
        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ Propagated = $false; Attempt = 0 } }
        if ($R.Propagated) {
            $Global:RolePollActive = $false
            Write-DebugLog "Role propagated after attempt $($R.Attempt)!"
            $lblStatus.Text = "Role propagated — uploads are now enabled"
            $statusDot.Fill = $Window.Resources['ThemeSuccess']
            # Re-trigger container load by re-selecting the storage account
            $SelIdx = $cmbUploadStorage.SelectedIndex
            if ($SelIdx -ge 0) {
                $cmbUploadStorage.SelectedIndex = -1
                $cmbUploadStorage.SelectedIndex = $SelIdx
            }
        } elseif ($R.Attempt -ge $Global:RolePollMaxAttempts) {
            $Global:RolePollActive = $false
            Write-DebugLog "Role propagation timed out after $($R.Attempt) attempts" -Level 'WARN'
            $lblStatus.Text = "Role propagation timed out — try switching storage account or wait a few minutes"
            $statusDot.Fill = $Window.Resources['ThemeWarning']
        } else {
            # Schedule next check
            Write-DebugLog "Role not yet propagated (attempt $($R.Attempt)), retrying..."
            $lblStatus.Text = "Waiting for role propagation... (attempt $($R.Attempt + 1)/$($Global:RolePollMaxAttempts))"
            Start-RolePropagationCheck
        }
    }
}

function New-StorageAccountFromUI {
    <# Validates form and creates a new Azure Storage Account in a background runspace #>
    # Get form values
    $AcctName = $txtNewStorageName.Text.Trim().ToLower()
    $RGName   = if ($cmbNewStorageRG.Text) { $cmbNewStorageRG.Text.Trim() }
                elseif ($cmbNewStorageRG.SelectedItem) { (Get-ComboBoxSelectedText $cmbNewStorageRG) }
                else { "" }
    $Location = Get-ComboBoxSelectedText $cmbNewStorageLocation
    $Sku      = Get-ComboBoxSelectedText $cmbNewStorageSku
    $Kind     = Get-ComboBoxSelectedText $cmbNewStorageKind
    $PublicAccess = [bool]$chkNewStoragePublicAccess.IsChecked

    # Validate
    $ValidationErrors = @()
    if (-not $AcctName)  { $ValidationErrors += "Storage account name is required" }
    if (-not $RGName)    { $ValidationErrors += "Resource group is required" }
    if (-not $Location)  { $ValidationErrors += "Location is required" }
    if ($AcctName -and ($AcctName.Length -lt 3 -or $AcctName.Length -gt 24)) {
        $ValidationErrors += "Name must be 3-24 characters"
    }
    if ($AcctName -and $AcctName -notmatch '^[a-z0-9]+$') {
        $ValidationErrors += "Name must contain only lowercase letters and numbers"
    }
    if ($ValidationErrors.Count -gt 0) {
        $lblNewStorageStatus.Text = $ValidationErrors -join "; "
        $lblNewStorageStatus.Foreground = $Window.Resources['ThemeError']
        return
    }

    # Disable form while creating
    $btnCreateStorage.IsEnabled = $false
    $prgNewStorage.Visibility = 'Visible'
    $lblNewStorageStatus.Text = "Creating storage account '$AcctName' in $Location..."
    $lblNewStorageStatus.Foreground = $Window.Resources['ThemeTextMuted']
    Write-DebugLog "Creating storage account: $AcctName (RG=$RGName, Location=$Location, SKU=$Sku, Kind=$Kind)"
    $lblStatus.Text = "Creating storage account..."

    Start-BackgroundWork -Variables @{
        AN  = $AcctName
        RG  = $RGName
        Loc = $Location
        SK  = $Sku
        KD  = $Kind
        PA  = $PublicAccess
        ResolveOidFnSrc = $Script:ResolveOidFnSrc
    } -Work {
        try {
            # Ensure resource group exists (create if typed a new name)
            $ExistingRG = Get-AzResourceGroup -Name $RG -ErrorAction SilentlyContinue
            if (-not $ExistingRG) {
                New-AzResourceGroup -Name $RG -Location $Loc -ErrorAction Stop | Out-Null
            }
            # Create storage account
            $Params = @{
                ResourceGroupName  = $RG
                Name               = $AN
                Location           = $Loc
                SkuName            = $SK
                Kind               = $KD
                AllowBlobPublicAccess = $PA
                MinimumTlsVersion  = 'TLS1_2'
                ErrorAction        = 'Stop'
            }
            $SA = New-AzStorageAccount @Params

            # Post-creation: assign Storage Blob Data Contributor to current user
            $PostSteps = @()
            try {
                # Recreate the helper inside this runspace (avoid cross-thread session state).
                Invoke-Expression $ResolveOidFnSrc
                $CurrentUser = (Get-AzContext).Account.Id
                $ObjectId = Resolve-CurrentUserObjectId -Email $CurrentUser
                if ($ObjectId) {
                    $ExistingRole = Get-AzRoleAssignment -ObjectId $ObjectId -Scope $SA.Id -RoleDefinitionName 'Storage Blob Data Contributor' -ErrorAction SilentlyContinue
                    if (-not $ExistingRole) {
                        New-AzRoleAssignment -ObjectId $ObjectId -RoleDefinitionName 'Storage Blob Data Contributor' -Scope $SA.Id -ErrorAction Stop | Out-Null
                        $PostSteps += 'Assigned Storage Blob Data Contributor role'
                    } else {
                        $PostSteps += 'Storage Blob Data Contributor role already assigned'
                    }
                } else {
                    $PostSteps += "Role assignment skipped: could not resolve directory ObjectId for '$CurrentUser' (guest user without consent?)"
                }
            } catch {
                $PostSteps += "Role assignment skipped: $($_.Exception.Message)"
            }

            # Post-creation: create 'packages' container with Blob-level public access
            try {
                $Ctx = New-AzStorageContext -StorageAccountName $AN -UseConnectedAccount -ErrorAction Stop
                New-AzStorageContainer -Name 'packages' -Context $Ctx -Permission Blob -ErrorAction Stop | Out-Null
                $PostSteps += "Created 'packages' container with blob-level public read access"
            } catch {
                $PostSteps += "Container creation skipped: $($_.Exception.Message)"
            }

            return @{
                Success   = $true
                Name      = $SA.StorageAccountName
                RG        = $RG
                Location  = $SA.PrimaryLocation
                PostSteps = $PostSteps
                Error     = $null
            }
        } catch {
            return @{ Success = $false; Name = $AN; Error = $_.Exception.Message }
        }
    } -OnComplete {
        param($Results, $Errors)
        $prgNewStorage.Visibility = 'Collapsed'
        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ Success = $false; Error = 'No result' } }
        if ($R.Success) {
            $lblNewStorageStatus.Text = "Storage account '$($R.Name)' created successfully!"
            $lblNewStorageStatus.Foreground = $Window.Resources['ThemeSuccess']
            Write-DebugLog "Storage account created: $($R.Name) in $($R.Location)"
            # Log post-creation steps
            if ($R.PostSteps) {
                foreach ($Step in $R.PostSteps) { Write-DebugLog "  Post-setup: $Step" }
            }
            $lblStatus.Text = "Storage account '$($R.Name)' created + configured"
            $statusDot.Fill = $Window.Resources['ThemeSuccess']
            # Refresh storage list and auto-close panel after brief delay
            Discover-StorageAccounts
            Hide-NewStoragePanel
        } else {
            $lblNewStorageStatus.Text = "Failed: $($R.Error)"
            $lblNewStorageStatus.Foreground = $Window.Resources['ThemeError']
            $btnCreateStorage.IsEnabled = $true
            Write-DebugLog "Storage creation failed: $($R.Error)" -Level 'ERROR'
            $lblStatus.Text = "Storage creation failed"
            $statusDot.Fill = $Window.Resources['ThemeError']
        }
    }
}

# ==============================================================================
# SECTION 7F: REST SOURCE PUBLISHER
# ==============================================================================

function Publish-ManifestToRestSource {
    <# Publishes manifest files to the WinGet REST source #>
    param([string]$ManifestPath, [string]$FunctionName, [switch]$SkipDuplicateCheck)
    if (-not (Test-Path $ManifestPath)) {
        Write-DebugLog "Manifest path not found: $ManifestPath" -Level 'ERROR'
        return
    }
    Write-DebugLog "Publishing manifests from: $ManifestPath (SkipDupCheck=$SkipDuplicateCheck)"
    $lblStatus.Text = "Publishing manifests..."
    Show-GlobalProgressIndeterminate -Text 'Publishing manifests to REST source...'
    $SkipDup = [bool]$SkipDuplicateCheck
    Start-BackgroundWork -Variables @{
        ManPath  = $ManifestPath
        FuncName = $FunctionName
        SkipDup  = $SkipDup
    } -Context @{
        ManPath  = $ManifestPath
        FuncName = $FunctionName
    } -Work {
        try {
            # Duplicate check — runs in background so UI stays responsive
            if (-not $SkipDup) {
                $ManFile = Get-ChildItem $ManPath -Filter '*.yaml' | Select-Object -First 1
                if ($ManFile) {
                    $YamlContent = Get-Content $ManFile.FullName -Raw
                    $PubPkgId = $null; $PubPkgVer = $null
                    if ($YamlContent -match '(?m)^PackageIdentifier:\s*(.+)$') { $PubPkgId = $Matches[1].Trim() }
                    if ($YamlContent -match '(?m)^PackageVersion:\s*(.+)$')    { $PubPkgVer = $Matches[1].Trim() }
                    if ($PubPkgId) {
                        try {
                            $Existing = Get-WinGetManifest -FunctionName $FuncName -PackageIdentifier $PubPkgId -ErrorAction SilentlyContinue
                            if ($Existing) {
                                $ExVer = if ($Existing.Versions) { $Existing.Versions | ForEach-Object {
                                    if ($_.PSObject.Properties['PackageVersion']) { $_.PackageVersion }
                                    elseif ($_.PSObject.Properties['Version']) { $_.Version }
                                }} else { @() }
                                if ($PubPkgVer -in $ExVer) {
                                    return @{ Success = $false; Error = $null; PackageInfo = "$PubPkgId v$PubPkgVer"; DuplicateFound = $true }
                                }
                            }
                        } catch { <# duplicate check failed -- proceed with publish #> }
                    }
                }
            }

            # Capture Add-WinGetManifest output (PackageIdentifier + Versions table)
            $CmdResult = Add-WinGetManifest -FunctionName $FuncName -Path $ManPath -ErrorAction Stop
            $PkgInfo = ''
            if ($CmdResult) {
                $Items = @($CmdResult)
                foreach ($Item in $Items) {
                    $PkgId  = if ($Item.PSObject.Properties['PackageIdentifier'])  { $Item.PackageIdentifier }
                              elseif ($Item.PSObject.Properties['PackageIdentidier']) { $Item.PackageIdentidier }  # known typo in module
                              else { '' }
                    $PkgVer = if ($Item.PSObject.Properties['Versions']) { ($Item.Versions | ForEach-Object { $_.ToString() }) -join ', ' }
                              elseif ($Item.PSObject.Properties['PackageVersion']) { $Item.PackageVersion }
                              else { '' }
                    if ($PkgId) { $PkgInfo = "$PkgId v$PkgVer" }
                }
            }
            return @{ Success = $true; Error = $null; PackageInfo = $PkgInfo }
        } catch {
            # Capture full error detail including inner exceptions and response body
            $ErrMsg = $_.Exception.Message
            $Inner = $_.Exception.InnerException
            while ($Inner) {
                $ErrMsg += " -> $($Inner.Message)"
                $Inner = $Inner.InnerException
            }
            # Try to extract HTTP response body for REST API errors
            if ($_.Exception.Response) {
                try {
                    $RespStream = $_.Exception.Response.GetResponseStream()
                    $Reader = New-Object System.IO.StreamReader($RespStream)
                    $Body = $Reader.ReadToEnd()
                    $Reader.Dispose()
                    if ($Body) { $ErrMsg += " | Response: $Body" }
                } catch { <# ignore stream read errors #> }
            }
            return @{ Success = $false; Error = $ErrMsg; PackageInfo = '' }
        }
    } -OnComplete {
        param($Results, $Errors, $Ctx)
        Write-DebugLog "[PublishCB] OnComplete ENTERED (Results=$($Results.Count), Errors=$($Errors.Count), ManPath=$($Ctx.ManPath))" -Level 'DEBUG'
        try {
        # Find our result hashtable (last item) — earlier items may be command output
        $R = $null
        if ($Results -and $Results.Count -gt 0) {
            for ($ri = $Results.Count - 1; $ri -ge 0; $ri--) {
                if ($Results[$ri] -is [hashtable] -and $Results[$ri].ContainsKey('Success')) {
                    $R = $Results[$ri]; break
                }
            }
        }
        if (-not $R) { $R = @{ Success = $false; Error = 'No result returned from background job'; PackageInfo = '' } }
        Write-DebugLog "[PublishCB] Result: Success=$($R.Success), Error=$($R.Error), PkgInfo=$($R.PackageInfo)" -Level 'DEBUG'

        # Handle duplicate detection — background found existing version, prompt user on UI thread
        if ($R.DuplicateFound) {
            Write-DebugLog "[PublishCB] Duplicate detected: $($R.PackageInfo) — prompting user" -Level 'DEBUG'
            $DupAnswer = Show-ThemedDialog -Title 'Package Version Exists' `
                -Message "$($R.PackageInfo) already exists in the REST source.`n`nPublishing will overwrite the existing version." `
                -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' `
                -Buttons @(
                    @{ Text='Cancel'; IsAccent=$false; Result='Cancel' }
                    @{ Text='Overwrite'; IsAccent=$true; Result='Overwrite' }
                )
            if ($DupAnswer -eq 'Overwrite') {
                Write-DebugLog "[PublishCB] User chose to overwrite — re-publishing (skip dup check)" -Level 'WARN'
                # Re-publish with duplicate check skipped to avoid infinite loop
                Publish-ManifestToRestSource -ManifestPath $Ctx.ManPath -FunctionName $Ctx.FuncName -SkipDuplicateCheck
            } else {
                Write-DebugLog "[PublishCB] User cancelled overwrite" -Level 'DEBUG'
                Hide-GlobalProgress
                Stop-StatusPulse
                $Script:PublishInProgress = $false
                Reset-ButtonBusy -Button $btnPublish
                # Clean up temp dir
                if ($Ctx.ManPath) { Remove-Item $Ctx.ManPath -Recurse -Force -ErrorAction SilentlyContinue }
            }
            return
        }

        if ($R.Success) {
            $Msg = "Manifest published successfully!"
            if ($R.PackageInfo) { $Msg = "Published: $($R.PackageInfo)" }
            Write-DebugLog $Msg
            $lblStatus.Text = $Msg
            $statusDot.Fill = $Window.Resources['ThemeSuccess']
            Show-Toast $Msg -Type Success
            Write-DebugLog "[PublishCB] Calling Stop-StatusPulse, Record-PublishSuccess, Check-PublishAchievements" -Level 'DEBUG'
            Stop-StatusPulse
            Record-PublishSuccess
            Check-PublishAchievements
        } else {
            # Combine result error with any error stream messages for full context
            $ErrMsg = $R.Error
            if ((-not $ErrMsg) -and $Errors -and $Errors.Count -gt 0) {
                $ErrMsg = ($Errors | ForEach-Object { $_.ToString() }) -join '; '
            }
            if (-not $ErrMsg) { $ErrMsg = 'Unknown error — check the Activity Log for details' }
            Write-DebugLog "Publish failed: $ErrMsg" -Level 'ERROR'
            $lblStatus.Text = "Publish failed — see Activity Log"
            $statusDot.Fill = $Window.Resources['ThemeError']
            Show-Toast "Publish failed — see Activity Log" -Type Error -DurationMs 6000
            Stop-StatusPulse
            Show-ThemedDialog -Title 'Publish Error' -Message "Publish failed:`n`n$ErrMsg" `
                -Icon ([string]([char]0xEA39)) -IconColor 'ThemeError' | Out-Null
        }
        # Reset publish guard, button, and clean up temp dir
        Write-DebugLog "[PublishCB] Reset publish guard and button" -Level 'DEBUG'
        Hide-GlobalProgress
        $Script:PublishInProgress = $false
        Reset-ButtonBusy -Button $btnPublish
        if ($Ctx.ManPath) { Remove-Item $Ctx.ManPath -Recurse -Force -ErrorAction SilentlyContinue }
        Write-DebugLog "[PublishCB] OnComplete EXIT (success)" -Level 'DEBUG'
        } catch {
            Write-DebugLog "[PublishCB] CRASH: $($_.Exception.Message)" -Level 'ERROR'
            Write-DebugLog "[PublishCB] ScriptStackTrace: $($_.ScriptStackTrace)" -Level 'ERROR'
        }
    }
}

# ==============================================================================
# SECTION 7G: SHA256 HASH CALCULATOR (Local File)
# ==============================================================================

$Global:CurrentHash0 = ""

function Compute-FileHash {
    <# Computes SHA256 of a local file in a background runspace #>
    param([string]$FilePath, [string]$TargetField)
    if (-not (Test-Path $FilePath)) {
        Write-DebugLog "File not found for hashing: $FilePath" -Level 'WARN'
        return
    }
    Write-DebugLog "Computing SHA256 for: $(Split-Path $FilePath -Leaf)" -Level 'DEBUG'
    $lblStatus.Text = "Computing SHA256 hash..."
    Start-BackgroundWork -Variables @{
        FPath = $FilePath
        Field = $TargetField
        SyncH = $Global:SyncHash
    } -Work {
        $FileInfo = Get-Item $FPath
        $Hash = (Get-FileHash -Path $FPath -Algorithm SHA256).Hash
        $SyncH.StatusQueue.Enqueue(@{
            Type  = 'HashResult'
            Value = $Hash
            Size  = $FileInfo.Length
            Name  = $FileInfo.Name
            Field = $Field
        })
        return @{ Hash = $Hash; Size = $FileInfo.Length; Name = $FileInfo.Name; Field = $Field }
    } -OnComplete {
        param($Results, $Errors)
        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { $null }
        if ($R -and $R.Hash) {
            $SizeMB = [math]::Round($R.Size / 1MB, 2)
            $Msg = "SHA256: $($R.Hash) ($SizeMB MB, $($R.Name))"
            Write-DebugLog $Msg
            if ($R.Field -eq 'installer0') {
                $lblHash0.Text = "SHA256: $($R.Hash)"
                $Global:CurrentHash0 = $R.Hash
                Update-YamlPreview
            } elseif ($R.Field -eq 'upload') {
                $lblUploadHash.Text = $R.Hash
                $lblUploadSize.Text = "$SizeMB MB - $($R.Name)"
                $btnCopyHash.IsEnabled = $true
            }
            $lblStatus.Text = "Hash computed: $($R.Name)"
        } else {
            Write-DebugLog "Hash computation failed" -Level 'ERROR'
            $lblStatus.Text = "Hash computation failed"
        }
    }
}

# ==============================================================================
# SECTION 7G2: SHA256 HASH FROM URL
# ==============================================================================

function Compute-UrlHash {
    <# Downloads an installer from a URL, stages it for upload, and computes SHA256 if not already set #>
    param([string]$Url)
    if (-not $Url -or $Url -notmatch '^https?://') {
        Show-ThemedDialog -Title 'Invalid URL' `
            -Message 'Enter a valid HTTP/HTTPS installer URL first.' `
            -Icon ([string]([char]0xE946)) -IconColor 'ThemeWarning' | Out-Null
        return
    }
    Write-DebugLog "Downloading installer: $Url" -Level 'DEBUG'
    $lblStatus.Text = "Downloading installer..."
    $lblHash0.Text = "Downloading..."
    $btnHashUrl0.IsEnabled = $false
    $btnCancelDownload.Visibility = 'Visible'
    $Global:DownloadCancelled = $false
    $Global:SyncHash.DownloadProgress = 0
    $Global:SyncHash.DownloadBytes    = 0
    $Global:SyncHash.DownloadTotal    = 0
    Start-BackgroundWork -Variables @{
        DlUrl  = $Url
        SyncH  = $Global:SyncHash
    } -Work {
        $TempDir = Join-Path $env:TEMP 'WinGetMM_Downloads'
        if (-not (Test-Path $TempDir)) { New-Item -ItemType Directory -Path $TempDir -Force | Out-Null }
        $FileName = [System.IO.Path]::GetFileName(([uri]$DlUrl).LocalPath)
        if (-not $FileName -or $FileName.Length -lt 2) { $FileName = 'installer_download' }
        $TempFile = Join-Path $TempDir $FileName
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        $Request  = $null
        $Response = $null
        $Stream   = $null
        $FileOut  = $null
        try {
            # Stream-based download with progress — avoids WebClient event threading issues
            $Request = [System.Net.HttpWebRequest]::Create([uri]$DlUrl)
            $Request.Timeout = 120000
            $Response = $Request.GetResponse()
            $TotalBytes = $Response.ContentLength   # -1 if server doesn't send Content-Length
            $Stream  = $Response.GetResponseStream()
            $FileOut = [System.IO.File]::Create($TempFile)
            $Buffer  = New-Object byte[] 65536
            $Downloaded = [long]0
            $LastReport = [DateTime]::MinValue
            while (($BytesRead = $Stream.Read($Buffer, 0, $Buffer.Length)) -gt 0) {
                $FileOut.Write($Buffer, 0, $BytesRead)
                $Downloaded += $BytesRead
                # Throttle progress updates to every 200ms to avoid overhead
                $Now = [DateTime]::UtcNow
                if (($Now - $LastReport).TotalMilliseconds -ge 200) {
                    $SyncH.DownloadBytes = $Downloaded
                    $SyncH.DownloadTotal = $TotalBytes
                    if ($TotalBytes -gt 0) {
                        $SyncH.DownloadProgress = [int]([math]::Floor($Downloaded * 100 / $TotalBytes))
                    } else {
                        $SyncH.DownloadProgress = 0
                    }
                    $LastReport = $Now
                }
            }
            $FileOut.Close(); $FileOut = $null
            $SyncH.DownloadProgress = 100
            $SyncH.DownloadBytes    = $Downloaded
            if ($Downloaded -eq 0) {
                throw "Download failed — file is empty"
            }
            $SyncH.DownloadProgress = -2  # signal: hashing phase
            $Hash = (Get-FileHash -Path $TempFile -Algorithm SHA256).Hash
            return @{ Hash = $Hash; Size = $Downloaded; Name = $FileName; Path = $TempFile }
        } catch {
            if (Test-Path $TempFile) { Remove-Item $TempFile -Force -ErrorAction SilentlyContinue }
            throw
        } finally {
            if ($FileOut)  { try { $FileOut.Dispose()  } catch {} }
            if ($Stream)   { try { $Stream.Dispose()   } catch {} }
            if ($Response) { try { $Response.Dispose()  } catch {} }
            $SyncH.DownloadProgress = -1
        }
    } -OnComplete {
        param($Results, $Errors)
        $btnHashUrl0.IsEnabled = $true
        $btnCancelDownload.Visibility = 'Collapsed'
        $Global:SyncHash.DownloadProgress = -1
        if ($Global:DownloadCancelled) {
            $lblHash0.Text = "Download cancelled"
            $lblStatus.Text = "Download cancelled"
            return
        }
        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { $null }
        if ($R -and $R.Path) {
            $SizeMB = [math]::Round($R.Size / 1MB, 2)
            $Global:CurrentHash0 = $R.Hash
            $lblHash0.Text = "SHA256: $($R.Hash)"
            # Stage for upload
            $txtUploadFile.Text = $R.Path
            $txtLocalFile0.Text = $R.Path
            Compute-FileHash -FilePath $R.Path -TargetField 'upload'
            Update-YamlPreview
            Write-DebugLog "Downloaded: $($R.Name) ($SizeMB MB) → $($R.Path)"
            $lblStatus.Text = "Downloaded $($R.Name) ($SizeMB MB) — staged for upload"
            Show-Toast "Downloaded $($R.Name) — ready for upload to storage" -Type Success -DurationMs 4000
        } else {
            $ErrMsg = if ($Errors -and $Errors.Count -gt 0) { $Errors[0].ToString() } else { 'Unknown error' }
            Write-DebugLog "Download failed: $ErrMsg" -Level 'ERROR'
            $lblHash0.Text = "Download failed — check URL"
            $lblStatus.Text = "Download from URL failed"
        }
    }
    # Store reference to the download job so it can be cancelled
    $Global:DownloadJob = $Global:BgJobs[$Global:BgJobs.Count - 1]
}

# ==============================================================================
# SECTION 7H-MANAGE: PACKAGE MANAGEMENT (Query, Remove, Cleanup)
# ==============================================================================

$Global:ManagePackages = @()   # cached list from last query

function Show-RestSourceConnectCommand {
    <# Resolves the WinGet REST source URL (APIM gateway preferred, falls back to Function App
       hostname) and shows a themed dialog with the full 'winget source add' command and a Copy
       button. Discovery runs in a background runspace to keep the UI responsive. #>
    $FuncName = if ($txtRestSourceName -and $txtRestSourceName.Text) { $txtRestSourceName.Text.Trim() } else { Get-ActiveRestSource }
    if (-not $FuncName) {
        Show-ThemedDialog -Title 'REST Source Not Set' `
            -Message 'Enter the REST Source Function name in Settings first.' `
            -Icon ([string]([char]0xE946)) -IconColor 'ThemeWarning' | Out-Null
        return
    }

    Show-GlobalProgressIndeterminate -Text "Resolving connection URL for $FuncName..."
    Write-DebugLog "Resolving connect command for $FuncName"

    Start-BackgroundWork -Variables @{
        FuncName = $FuncName
    } -Work {
        $Result = @{
            FuncName  = $FuncName
            FuncUrl   = $null
            ApimUrl   = $null
            Source    = 'unknown'
            Error     = $null
        }
        try {
            # Locate the function app
            $Fa = Get-AzResource -ResourceType 'Microsoft.Web/sites' -Name $FuncName -ErrorAction SilentlyContinue
            if (-not $Fa) {
                # Try without filter (cross-RG)
                $Fa = @(Get-AzResource -ResourceType 'Microsoft.Web/sites' -ErrorAction SilentlyContinue |
                        Where-Object { $_.Name -eq $FuncName })[0]
            }
            if (-not $Fa) {
                $Result.Error = "Function App '$FuncName' not found in current subscription."
                return $Result
            }
            $Result.FuncUrl = "https://$($Fa.Name).azurewebsites.net/api/"

            # Look for an APIM in the same RG fronting it (best-effort heuristic)
            $Apims = @(Get-AzResource -ResourceType 'Microsoft.ApiManagement/service' `
                        -ResourceGroupName $Fa.ResourceGroupName -ErrorAction SilentlyContinue)
            if ($Apims.Count -gt 0) {
                try {
                    $Apim = $Apims[0]
                    $Token = (Get-AzAccessToken -ResourceUrl 'https://management.azure.com/' -ErrorAction Stop).Token
                    $Uri = "https://management.azure.com$($Apim.ResourceId)?api-version=2023-05-01-preview"
                    $Resp = Invoke-RestMethod -Uri $Uri -Headers @{ Authorization = "Bearer $Token" } -Method GET -ErrorAction Stop
                    if ($Resp.properties.gatewayUrl) {
                        $Result.ApimUrl = "$($Resp.properties.gatewayUrl)/winget/"
                    }
                } catch {
                    # APIM lookup is best-effort; ignore
                }
            }
            $Result.Source = if ($Result.ApimUrl) { 'apim' } else { 'function' }
            return $Result
        } catch {
            $Result.Error = $_.Exception.Message
            return $Result
        }
    } -OnComplete {
        param($Results, $Errors)
        Hide-GlobalProgress
        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { $null }
        if (-not $R -or $R.Error) {
            $Msg = if ($R -and $R.Error) { $R.Error } else { 'Unknown error resolving REST source URL.' }
            Show-ThemedDialog -Title 'Resolve Failed' -Message $Msg `
                -Icon ([string]([char]0xEA39)) -IconColor 'ThemeError' | Out-Null
            return
        }

        $Url      = if ($R.ApimUrl) { $R.ApimUrl } else { $R.FuncUrl }
        $UrlLabel = if ($R.ApimUrl) { "APIM gateway" } else { "Function App (direct)" }
        $Cmd      = "winget source add --name `"$($R.FuncName)`" --arg `"$Url`" --type `"Microsoft.Rest`" --accept-source-agreements"

        $Body = @"
Source: $UrlLabel
URL:    $Url

Run the command below from an ELEVATED PowerShell window on the target machine:

$Cmd

Tip: verify the endpoint first with
  Invoke-RestMethod '$($Url.TrimEnd('/'))/information'

If APIM fronts the source and requires a subscription key, append:
  --header "Ocp-Apim-Subscription-Key:<your-key>"
"@

        $Choice = Show-ThemedDialog -Title 'Connect Command' -Message $Body `
            -Icon ([string]([char]0xE71B)) -IconColor 'ThemeAccentLight' `
            -Buttons @(
                @{ Text='Copy command'; IsAccent=$true; Result='Copy' }
                @{ Text='Close';        IsAccent=$false; Result='Close' }
            )
        if ($Choice -eq 'Copy') {
            try {
                [System.Windows.Clipboard]::SetText($Cmd)
                if (Get-Command Show-Toast -ErrorAction SilentlyContinue) {
                    Show-Toast -Message "Command copied to clipboard" -Type 'Success'
                }
                Write-DebugLog "Connect command copied to clipboard ($UrlLabel)"
            } catch {
                Write-DebugLog "Clipboard copy failed: $($_.Exception.Message)" -Level 'WARN'
            }
        }
    }
}

function Test-RestSourceConnection {
    <# Tests connectivity to the WinGet REST source by performing a lightweight query #>
    $FuncName = Get-ActiveRestSource
    if (-not $FuncName) {
        Show-ThemedDialog -Title 'REST Source Not Set' `
            -Message 'Enter the REST Source Function name in Settings first.' `
            -Icon ([string]([char]0xE946)) -IconColor 'ThemeWarning' | Out-Null
        return
    }
    $lblRestSourceStatus = $Window.FindName("lblRestSourceStatus")
    $lblRestSourceStatus.Text = "Testing connection..."
    $lblRestSourceStatus.Foreground = $Window.FindResource("ThemeTextMuted")
    $lblStatus.Text = "Testing REST source..."
    Show-GlobalProgressIndeterminate -Text 'Testing REST source connection...'
    Start-BackgroundWork -Variables @{
        FuncName = $FuncName
    } -Work {
        try {
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $Result = Find-WinGetManifest -FunctionName $FuncName -ErrorAction Stop
            $sw.Stop()
            $Count = if ($Result) { @($Result).Count } else { 0 }
            return @{ Success = $true; Count = $Count; ElapsedMs = $sw.ElapsedMilliseconds; Error = $null }
        } catch {
            return @{ Success = $false; Count = 0; ElapsedMs = 0; Error = $_.Exception.Message }
        }
    } -OnComplete {
        param($Results, $Errors)
        $lblRSS = $Window.FindName("lblRestSourceStatus")
        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { $null }
        if ($R -and $R.Success) {
            $lblRSS.Text = [char]0xE73E + " Connected — $($R.Count) packages ($($R.ElapsedMs)ms)"
            $lblRSS.Foreground = $Window.FindResource("ThemeGreenAccent")
            $lblStatus.Text = "REST source OK"
            Write-DebugLog "REST source test passed: $($R.Count) packages in $($R.ElapsedMs)ms"
        } else {
            $ErrMsg = if ($R -and $R.Error) { $R.Error } elseif ($Errors -and $Errors.Count -gt 0) { $Errors[0].ToString() } else { 'Unknown error' }
            $lblRSS.Text = [char]0xEA39 + " Connection failed: $ErrMsg"
            $lblRSS.Foreground = $Window.FindResource("ThemeRedAccent")
            $lblStatus.Text = "REST source test failed"
            Write-DebugLog "REST source test failed: $ErrMsg" -Level 'ERROR'
        }
        Hide-GlobalProgress
    }
}

$Global:ManageCacheTime   = $null  # timestamp of last successful query
$Global:ManageCacheSource = ''     # function name that was queried
$Global:ManageCacheTTL    = 300    # seconds before cache expires (5 min)

function Update-ManageCacheIndicator {
    <# Updates the cache age label on the Manage tab #>
    $lblCacheAge = $Window.FindName('lblManageCacheAge')
    if (-not $lblCacheAge) { return }
    if (-not $Global:ManageCacheTime) {
        $lblCacheAge.Text = ''
        return
    }
    $AgeSec = [int]((Get-Date) - $Global:ManageCacheTime).TotalSeconds
    $TTL    = [int]$Global:ManageCacheTTL
    if ($TTL -le 0) {
        $lblCacheAge.Text = "Cached $([math]::Floor($AgeSec / 60))m $($AgeSec % 60)s ago (caching disabled)"
        return
    }
    $Remaining = [math]::Max(0, $TTL - $AgeSec)
    if ($AgeSec -lt 60) {
        $AgeStr = "${AgeSec}s ago"
    } else {
        $AgeStr = "$([math]::Floor($AgeSec / 60))m $($AgeSec % 60)s ago"
    }
    if ($Remaining -le 0) {
        $lblCacheAge.Text = "Cache expired \u2014 will refresh on next query"
    } else {
        $lblCacheAge.Text = "Cached $AgeStr \u00B7 expires in $([math]::Floor($Remaining / 60))m $($Remaining % 60)s"
    }
}

function Refresh-ManagePackageList {
    <# Queries the WinGet REST source for published packages and populates the list #>
    param([switch]$ForceRefresh)
    $FuncName = Get-ActiveRestSource
    if (-not $FuncName) {
        Show-ThemedDialog -Title 'REST Source Not Set' `
            -Message 'Enter the REST Source Function name in Settings first.' `
            -Icon ([string]([char]0xE946)) -IconColor 'ThemeWarning' | Out-Null
        return
    }

    # Check TTL cache unless forced
    if (-not $ForceRefresh -and $Global:ManageCacheSource -eq $FuncName -and $Global:ManageCacheTime -and
        ((Get-Date) - $Global:ManageCacheTime).TotalSeconds -lt $Global:ManageCacheTTL -and $Global:ManagePackages.Count -gt 0) {
        Write-DebugLog "Manage: using cached results ($($Global:ManagePackages.Count) packages, age $([int]((Get-Date) - $Global:ManageCacheTime).TotalSeconds)s)"
        Update-ManageCacheIndicator
        return
    }

    $lblManageSourceName.Text = "REST Source: $FuncName"
    $lblManageStatus.Text = "Querying packages..."
    $lblStatus.Text = "Querying REST source..."
    $prgManage.Visibility = 'Visible'
    $prgManage.IsIndeterminate = $true
    Show-GlobalProgressIndeterminate -Text 'Querying REST source packages...'
    Write-DebugLog "Manage: querying packages from $FuncName"

    Start-BackgroundWork -Variables @{ FuncName = $FuncName } -Work {
        try {
            $Pkgs = Find-WinGetManifest -FunctionName $FuncName -ErrorAction Stop
            $Result = @()
            foreach ($P in @($Pkgs)) {
                $PkgId  = if ($P.PSObject.Properties['PackageIdentifier'])  { $P.PackageIdentifier }
                          elseif ($P.PSObject.Properties['PackageIdentidier']) { $P.PackageIdentidier }  # known typo
                          else { '' }
                $Versions = @()
                if ($P.PSObject.Properties['Versions']) {
                    $Versions = @($P.Versions | ForEach-Object { $_.ToString() })
                } elseif ($P.PSObject.Properties['PackageVersion']) {
                    $Versions = @($P.PackageVersion.ToString())
                }
                if ($PkgId) {
                    $Result += @{ Id = $PkgId; Versions = $Versions }
                }
            }
            return @{ Success = $true; Packages = $Result; Error = $null }
        } catch {
            return @{ Success = $false; Packages = @(); Error = $_.Exception.Message }
        }
    } -OnComplete {
        param($Results, $Errors)
        $prgManage.Visibility = 'Collapsed'
        $prgManage.IsIndeterminate = $false
        Hide-GlobalProgress
        $R = $null
        if ($Results -and $Results.Count -gt 0) {
            for ($ri = $Results.Count - 1; $ri -ge 0; $ri--) {
                if ($Results[$ri] -is [hashtable] -and $Results[$ri].ContainsKey('Success')) {
                    $R = $Results[$ri]; break
                }
            }
        }
        if (-not $R) { $R = @{ Success = $false; Packages = @(); Error = 'No result from query' } }

        if ($R.Success) {
            $Global:ManagePackages = $R.Packages
            $Global:ManageCacheTime = Get-Date
            $Global:ManageCacheSource = $FuncName
            Update-ManageCacheIndicator
            $lstManagePackages.Items.Clear()
            if ($R.Packages.Count -eq 0) {
                $lblManageEmpty.Text = "No packages found in the REST source."
                $lblManageEmpty.Visibility = 'Visible'
                $lstManagePackages.Visibility = 'Collapsed'
            } else {
                $lblManageEmpty.Visibility = 'Collapsed'
                $lstManagePackages.Visibility = 'Visible'
                foreach ($Pkg in $R.Packages) {
                    $Item = New-StyledListViewItem -IconChar ([char]0xE7C3) -PrimaryText $Pkg.Id `
                        -SubtitleText "Versions: $(($Pkg.Versions -join ', '))" -TagValue $Pkg
                    $lstManagePackages.Items.Add($Item) | Out-Null
                }
            }
            $lblManagePkgCount.Text = "$($R.Packages.Count) package$(if ($R.Packages.Count -ne 1){'s'})"
            $lblManageStatus.Text = "Found $($R.Packages.Count) package$(if ($R.Packages.Count -ne 1){'s'})"
            $lblStatus.Text = "Package query complete"
            $statusDot.Fill = $Window.Resources['ThemeSuccess']
            Write-DebugLog "Manage: found $($R.Packages.Count) packages"
        } else {
            $ErrMsg = $R.Error
            if ((-not $ErrMsg) -and $Errors -and $Errors.Count -gt 0) {
                $ErrMsg = ($Errors | ForEach-Object { $_.ToString() }) -join '; '
            }
            Write-DebugLog "Manage: query failed — $ErrMsg" -Level 'ERROR'
            $lblManageStatus.Text = "Query failed — see Activity Log"
            $lblManageEmpty.Text = "Query failed: $ErrMsg"
            $lblManageEmpty.Visibility = 'Visible'
            $lstManagePackages.Visibility = 'Collapsed'
            $lblStatus.Text = "Package query failed"
            $statusDot.Fill = $Window.Resources['ThemeError']
        }
        $pnlManageDetails.Visibility = 'Collapsed'
    }
}

function Remove-PackageFromRestSource {
    <# Removes a package or version from the WinGet REST source #>
    param(
        [string]$PackageId,
        [string]$Version,          # empty = remove entire package
        [bool]$DeleteBlob   = $false,
        [bool]$DeleteConfig = $false
    )
    $FuncName = Get-ActiveRestSource
    if (-not $FuncName -or -not $PackageId) { return }

    $VerLabel = if ($Version) { "v$Version" } else { '(all versions)' }
    Write-DebugLog "Manage: removing $PackageId $VerLabel from $FuncName"
    $lblManageStatus.Text = "Removing $PackageId $VerLabel..."
    $lblStatus.Text = "Removing package..."
    $prgManage.Visibility = 'Visible'
    $prgManage.IsIndeterminate = $true
    Show-GlobalProgressIndeterminate -Text "Removing $PackageId..."

    # Build variables for background work
    $Vars = @{
        FuncName     = $FuncName
        PkgId        = $PackageId
        PkgVer       = $Version
        DelBlob      = $DeleteBlob
        DelConfig    = $DeleteConfig
        ConfigDir    = $Global:ConfigDir
        BlobTargets  = @()  # populated below from local config
    }

    # Pre-read local config to extract blob coordinates for deletion
    if ($DeleteBlob) {
        $CfgFile = Join-Path $Global:ConfigDir "$PackageId.json"
        if (Test-Path $CfgFile) {
            try {
                $Cfg = Get-Content $CfgFile -Raw | ConvertFrom-Json
                $Targets = @()
                # Prefer the explicit storage block
                $SA   = $Cfg.storage.storageAccountName
                $Cont = $Cfg.storage.containerName
                $Blob = $Cfg.storage.blobPath
                if ($SA -and $Cont -and $Blob) {
                    $Targets += @{ Account = $SA; Container = $Cont; Blob = $Blob }
                }
                # Fallback: parse each installer URL
                foreach ($Inst in @($Cfg.installers)) {
                    $Url = $Inst.installerUrl
                    if ($Url -match '^https://([^.]+)\.blob\.core\.windows\.net/([^/]+)/(.+)$') {
                        $Parsed = @{ Account = $Matches[1]; Container = $Matches[2]; Blob = $Matches[3] }
                        # Avoid duplicates
                        $Dup = $Targets | Where-Object { $_.Account -eq $Parsed.Account -and $_.Container -eq $Parsed.Container -and $_.Blob -eq $Parsed.Blob }
                        if (-not $Dup) { $Targets += $Parsed }
                    }
                }
                $Vars.BlobTargets = $Targets
                Write-DebugLog "Manage: found $($Targets.Count) blob target(s) for deletion"
            } catch {
                Write-DebugLog "Manage: could not read config for blob info — $($_.Exception.Message)" -Level 'WARN'
            }
        } else {
            Write-DebugLog "Manage: no local config found for $PackageId — blob deletion skipped" -Level 'WARN'
        }
    }

    Start-BackgroundWork -Variables $Vars -Work {
        $Errors = @()
        $Removed = $false
        try {
            if ($PkgVer) {
                Remove-WinGetManifest -FunctionName $FuncName -PackageIdentifier $PkgId -PackageVersion $PkgVer -ErrorAction Stop
            } else {
                Remove-WinGetManifest -FunctionName $FuncName -PackageIdentifier $PkgId -ErrorAction Stop
            }
            $Removed = $true
        } catch {
            $Errors += "REST remove: $($_.Exception.Message)"
        }

        # Delete blob if requested
        $BlobDeleted = $false
        if ($DelBlob -and $Removed -and $BlobTargets.Count -gt 0) {
            foreach ($BT in $BlobTargets) {
                try {
                    $Ctx = New-AzStorageContext -StorageAccountName $BT.Account -UseConnectedAccount -ErrorAction Stop
                    Remove-AzStorageBlob -Container $BT.Container -Blob $BT.Blob -Context $Ctx -Force -ErrorAction Stop
                    $BlobDeleted = $true
                } catch {
                    $Errors += "Blob delete ($($BT.Account)/$($BT.Container)/$($BT.Blob)): $($_.Exception.Message)"
                }
            }
        }

        # Delete local config if requested
        $ConfigDeleted = $false
        if ($DelConfig -and $Removed) {
            try {
                $ConfigFile = Join-Path $ConfigDir "$PkgId.json"
                if (Test-Path $ConfigFile) {
                    Remove-Item $ConfigFile -Force
                    $ConfigDeleted = $true
                }
            } catch {
                $Errors += "Config delete: $($_.Exception.Message)"
            }
        }

        return @{
            Success       = $Removed
            PkgId         = $PkgId
            Version       = $PkgVer
            BlobDeleted   = $BlobDeleted
            ConfigDeleted = $ConfigDeleted
            Errors        = $Errors
        }
    } -OnComplete {
        param($Results, $Errors)
        $prgManage.Visibility = 'Collapsed'
        $prgManage.IsIndeterminate = $false
        Hide-GlobalProgress
        $R = $null
        if ($Results -and $Results.Count -gt 0) {
            for ($ri = $Results.Count - 1; $ri -ge 0; $ri--) {
                if ($Results[$ri] -is [hashtable] -and $Results[$ri].ContainsKey('Success')) {
                    $R = $Results[$ri]; break
                }
            }
        }
        if (-not $R) { $R = @{ Success = $false; Errors = @('No result'); PkgId = $PackageId } }

        if ($R.Success) {
            $VerStr = if ($R.Version) { "v$($R.Version)" } else { '(all versions)' }
            $Msg = "Removed $($R.PkgId) $VerStr"
            if ($R.BlobDeleted)   { $Msg += " + blob" }
            if ($R.ConfigDeleted) { $Msg += " + local config" }
            Write-DebugLog $Msg
            $lblManageStatus.Text = $Msg
            $lblStatus.Text = $Msg
            $statusDot.Fill = $Window.Resources['ThemeSuccess']
            Unlock-Achievement 'first_delete'
            $pnlManageDetails.Visibility = 'Collapsed'
            # Refresh both lists
            Refresh-ManagePackageList
            Refresh-ConfigList
        } else {
            $ErrMsg = ($R.Errors -join '; ')
            if ((-not $ErrMsg) -and $Errors -and $Errors.Count -gt 0) {
                $ErrMsg = ($Errors | ForEach-Object { $_.ToString() }) -join '; '
            }
            Write-DebugLog "Remove failed: $ErrMsg" -Level 'ERROR'
            $lblManageStatus.Text = "Remove failed — see Activity Log"
            $lblStatus.Text = "Remove failed"
            $statusDot.Fill = $Window.Resources['ThemeError']
            Show-ThemedDialog -Title 'Remove Failed' `
                -Message "Failed to remove package:`n`n$ErrMsg" `
                -Icon ([string]([char]0xEA39)) -IconColor 'ThemeError' | Out-Null
        }
    }
}

# ==============================================================================
# SECTION 7H: MANIFEST VALIDATOR
# ==============================================================================

function Validate-ManifestFields {
    <# Validates required manifest fields and optionally runs full schema validation #>
    param(
        [switch]$ForPublish,
        [switch]$FullSchema
    )
    $Errors = @()
    if (-not $txtPkgId.Text.Trim())      { $Errors += "Package Identifier is required" }
    if (-not $txtPkgVersion.Text.Trim())  { $Errors += "Package Version is required" }
    if (-not $txtPkgName.Text.Trim())     { $Errors += "Package Name is required" }
    if (-not $txtPublisher.Text.Trim())   { $Errors += "Publisher is required" }
    if (-not $txtLicense.Text.Trim())     { $Errors += "License is required" }
    if (-not $txtShortDesc.Text.Trim())   { $Errors += "Short Description is required" }
    if (-not (Get-ComboBoxSelectedText $cmbInstallerType0)) { $Errors += "Installer Type is required" }

    # Validate PackageIdentifier format
    $PkgId = $txtPkgId.Text.Trim()
    if ($PkgId -and $PkgId -notmatch '^[A-Za-z0-9][\w.-]+\.[A-Za-z0-9][\w.-]+$') {
        $Errors += "Package Identifier must be in 'Publisher.Package' format"
    }

    # Extra checks when publishing to REST source
    if ($ForPublish) {
        if (-not $txtInstallerUrl0.Text.Trim()) {
            $Errors += "Installer URL is required — upload the package first or enter a URL"
        }
        if (-not $Global:CurrentHash0) {
            $Errors += "Installer SHA256 hash is required — browse a local file to compute it"
        }
    }

    # Full schema validation (URL formats, version format, enums, field lengths)
    if ($FullSchema -or $ForPublish) {
        $SchemaErrors = Validate-ManifestSchema
        $Errors += $SchemaErrors | Where-Object { $_ -notin $Errors }
    }

    return $Errors
}

function Save-ManifestToDisk {
    <# Saves the 3-file manifest set to disk #>
    $ValidationErrors = Validate-ManifestFields
    if ($ValidationErrors.Count -gt 0) {
        $Msg = "Validation errors:`n" + ($ValidationErrors -join "`n")
        Write-DebugLog $Msg -Level 'WARN'
        Show-ThemedDialog -Title 'Validation Error' -Message $Msg `
            -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' | Out-Null
        return
    }

    $PkgId  = $txtPkgId.Text.Trim()
    $PkgVer = $txtPkgVersion.Text.Trim()
    $ManVer = Get-ComboBoxSelectedText $cmbCreateManifestVer
    if (-not $ManVer) { $ManVer = "1.9.0" }

    # Determine base path
    $BasePath = $txtManifestPath.Text.Trim()
    if (-not $BasePath) { $BasePath = Join-Path $Global:Root "manifests" }

    # Create folder structure: manifests/<FirstChar>/<PkgId>/<Version>/
    $FirstChar = $PkgId.Substring(0,1).ToUpper()
    $ManifestDir = Join-Path $BasePath $FirstChar $PkgId $PkgVer
    if (-not (Test-Path $ManifestDir)) {
        New-Item -Path $ManifestDir -ItemType Directory -Force | Out-Null
    }

    $Meta = Get-MetadataFromUI
    $AllInstallers = Get-AllInstallerData

    # Generate YAML files
    $VersionYaml   = Build-VersionYaml -PkgId $PkgId -PkgVersion $PkgVer -ManifestVersion $ManVer -Locale "en-US"
    $InstallerYaml = Build-InstallerYaml -PkgId $PkgId -PkgVersion $PkgVer -ManifestVersion $ManVer -Installers $AllInstallers
    $LocaleYaml    = Build-DefaultLocaleYaml -PkgId $PkgId -PkgVersion $PkgVer -ManifestVersion $ManVer -Meta $Meta

    # Write files
    $VersionYaml   | Set-Content (Join-Path $ManifestDir "$PkgId.yaml") -Force -Encoding UTF8
    $InstallerYaml | Set-Content (Join-Path $ManifestDir "$PkgId.installer.yaml") -Force -Encoding UTF8
    $LocaleYaml    | Set-Content (Join-Path $ManifestDir "$PkgId.locale.en-US.yaml") -Force -Encoding UTF8

    Write-DebugLog "Manifests saved to: $ManifestDir"
    $lblStatus.Text = "Manifests saved to: $ManifestDir"
    $statusDot.Fill = $Window.Resources['ThemeSuccess']
    Show-Toast "Manifests saved to disk" -Type Success
    Unlock-Achievement 'yaml_export'
}

# ==============================================================================
# SECTION 7I: PSADT PACKAGE IMPORTER
# ==============================================================================

function Import-PSADTPackage {
    <# Parses a PSADT Deploy-Application.ps1 and populates the Create tab fields #>
    param([string]$FolderPath)

    $ScriptPath = Join-Path $FolderPath "Deploy-Application.ps1"
    if (-not (Test-Path $ScriptPath)) {
        Show-ThemedDialog -Title 'Import Error' -Message "Deploy-Application.ps1 not found in the selected folder.`n`nSelect the root of a PSADT package." `
            -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' | Out-Null
        return $null
    }

    $FilesDir = Join-Path $FolderPath "Files"
    $ScriptContent = Get-Content $ScriptPath -Raw -ErrorAction SilentlyContinue
    if (-not $ScriptContent) {
        Show-ThemedDialog -Title 'Import Error' -Message "Could not read Deploy-Application.ps1" `
            -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' | Out-Null
        return $null
    }

    Write-DebugLog "Importing PSADT package from: $FolderPath"

    # ── Parse PSADT variables ──
    $Parsed = @{
        AppName    = ''
        AppVersion = ''
        AppVendor  = ''
        AppArch    = ''
        AppAuthor  = ''
    }

    # Match: [string]$appName = 'value'  or  [string]$appName = "value"
    $VarPatterns = @{
        AppName    = '\$appName\s*=\s*[''"]([^''"]*)[''"]'
        AppVersion = '\$appVersion\s*=\s*[''"]([^''"]*)[''"]'
        AppVendor  = '\$appVendor\s*=\s*[''"]([^''"]*)[''"]'
        AppArch    = '\$appArch\s*=\s*[''"]([^''"]*)[''"]'
        AppAuthor  = '\$appScriptAuthor\s*=\s*[''"]([^''"]*)[''"]'
    }
    foreach ($Key in $VarPatterns.Keys) {
        if ($ScriptContent -match $VarPatterns[$Key]) {
            $Parsed[$Key] = $Matches[1].Trim()
        }
    }

    # ── Detect installer files ──
    $InstallerFile = $null
    $InstallerType = ''
    $InstallerPath = ''

    if (Test-Path $FilesDir) {
        $InstallerFiles = Get-ChildItem $FilesDir -File | Where-Object {
            $_.Extension -match '\.(msi|exe|msix|appx)$'
        }
        if ($InstallerFiles.Count -gt 0) {
            $InstallerFile = $InstallerFiles[0]
            $InstallerPath = $InstallerFile.FullName
            switch -Regex ($InstallerFile.Extension) {
                '\.msi$'  { $InstallerType = 'msi' }
                '\.exe$'  { $InstallerType = 'exe' }
                '\.msix$' { $InstallerType = 'msix' }
                '\.appx$' { $InstallerType = 'msix' }
            }
        }
    }

    # ── Detect architecture from filename if $appArch is empty ──
    $DetectedArch = 'x64'  # default
    $ArchSource = if ($Parsed.AppArch) { $Parsed.AppArch } elseif ($InstallerFile) { $InstallerFile.Name } else { '' }
    if ($ArchSource) {
        if ($ArchSource -match 'x64|amd64|64bit') { $DetectedArch = 'x64' }
        elseif ($ArchSource -match 'x86|32bit|i386') { $DetectedArch = 'x86' }
        elseif ($ArchSource -match 'arm64|aarch64') { $DetectedArch = 'arm64' }
    }

    # ── Parse silent switches from Execute-MSI / Execute-Process calls ──
    $SilentSwitch = ''
    $SilentProgressSwitch = ''
    $ProductCode = ''

    # Execute-MSI with -parameters
    if ($ScriptContent -match 'Execute-MSI\s+.*-[Pp]ath\s+[''"][^''"]*[''"].*-[Pp]arameters\s+[''"]([^''"]*)[''"]') {
        $MsiParams = $Matches[1].Trim()
        if ($MsiParams -match '/qn') { $SilentSwitch = '/qn' }
        if ($MsiParams -match '/qb') { $SilentProgressSwitch = '/qb' }
    }
    # MSI without explicit params = silent by default
    if (-not $SilentSwitch -and $InstallerType -eq 'msi') {
        $SilentSwitch = '/qn'
        $SilentProgressSwitch = '/qb'
    }

    # Execute-Process with -Parameters for EXE
    if ($InstallerType -eq 'exe' -and $ScriptContent -match 'Execute-Process\s+.*-[Pp]arameters\s+[''"]([^''"]*)[''"]') {
        $ExeParams = $Matches[1].Trim()
        $SilentSwitch = $ExeParams
    }

    # Product code from uninstall section
    if ($ScriptContent -match 'Execute-MSI\s+.*-Action\s+[''"]Uninstall[''"].*-Path\s+[''"](\{[0-9A-Fa-f\-]+\})[''"]') {
        $ProductCode = $Matches[1]
    }

    # Fallback: extract MSI GUID from installer filename (e.g. {GUID}.msi)
    if (-not $ProductCode -and $InstallerPath -and $InstallerPath -match '(\{[0-9A-Fa-f\-]+\})') {
        $ProductCode = $Matches[1]
    }

    # ── Parse CloseApps from Show-InstallationWelcome → Commands ──
    $CloseApps = @()
    if ($ScriptContent -match 'Show-InstallationWelcome\s+.*-CloseApps\s+[''"]([^''"]*)[''"]') {
        $CloseApps = @($Matches[1] -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    # ── Parse uninstall commands from the Uninstall section ──
    $UninstallCmds = @()
    $HasUninstallSection = $false
    if ($ScriptContent -match '(?si)deploymentType\s+-ieq\s+[''"]Uninstall[''"].*?\{(.+?)(?:ElseIf|\z)') {
        $UninstallBlock = $Matches[1]
        $HasUninstallSection = $true

        # Execute-MSI -Action Uninstall
        $MsiUninstalls = [regex]::Matches($UninstallBlock, 'Execute-MSI\s+.*-Action\s+[''"]Uninstall[''"].*')
        foreach ($m in $MsiUninstalls) { $UninstallCmds += $m.Value.Trim() }

        # Execute-Process in uninstall block
        $ExeUninstalls = [regex]::Matches($UninstallBlock, 'Execute-Process\s+.*')
        foreach ($m in $ExeUninstalls) { $UninstallCmds += $m.Value.Trim() }

        # Remove-MSIApplications
        $RemoveMsi = [regex]::Matches($UninstallBlock, 'Remove-MSIApplications\s+.*')
        foreach ($m in $RemoveMsi) { $UninstallCmds += $m.Value.Trim() }
    }

    # ── Detect repair section ──
    $HasRepairSection = $false
    if ($ScriptContent -match '(?si)deploymentType\s+-ieq\s+[''"]Repair[''"].*?\{(.+?)(?:ElseIf|\z)') {
        $RepairBlock = $Matches[1]
        # Check it has actual commands, not just comments
        if ($RepairBlock -match '(Execute-MSI|Execute-Process|Remove-MSIApplications)') {
            $HasRepairSection = $true
        }
    }

    # ── File extensions from Files directory ──
    $FileExtensions = @()
    $FilesDir = Join-Path $FolderPath 'Files'
    if (Test-Path $FilesDir) {
        $FileExtensions = @(Get-ChildItem -Path $FilesDir -File -Recurse -ErrorAction SilentlyContinue |
            ForEach-Object { $_.Extension.TrimStart('.').ToLower() } |
            Where-Object { $_ -and $_ -notin @('log','txt','md','ini','cfg') } |
            Sort-Object -Unique)
    }

    # ── DeferTimes from Show-InstallationWelcome ──
    $DeferTimes = 0
    if ($ScriptContent -match 'Show-InstallationWelcome\s+.*-DeferTimes\s+(\d+)') {
        $DeferTimes = [int]$Matches[1]
    }

    # ── AllowRebootPassThru ──
    $AllowReboot = $false
    if ($ScriptContent -match 'AllowRebootPassThru') {
        $AllowReboot = $true
    }

    # ── Build result object ──
    return [PSCustomObject]@{
        AppName              = $Parsed.AppName
        AppVersion           = $Parsed.AppVersion
        AppVendor            = $Parsed.AppVendor
        AppArch              = $DetectedArch
        AppAuthor            = $Parsed.AppAuthor
        AppLang              = $Parsed.AppLang
        AppRevision          = $Parsed.AppRevision
        InstallName          = $Parsed.InstallName
        InstallTitle         = $Parsed.InstallTitle
        ScriptVersion        = $Parsed.ScriptVersion
        ScriptDate           = $Parsed.ScriptDate
        InstallerType        = $InstallerType
        InstallerFilePath    = $InstallerPath
        SilentSwitch         = $SilentSwitch
        SilentProgressSwitch = $SilentProgressSwitch
        ProductCode          = $ProductCode
        Commands             = $CloseApps
        UninstallCmds        = $UninstallCmds
        HasUninstallSection  = $HasUninstallSection
        HasRepairSection     = $HasRepairSection
        FileExtensions       = $FileExtensions
        DeferTimes           = $DeferTimes
        AllowReboot          = $AllowReboot
        SourceFolder         = $FolderPath
    }
}

function Expand-PSADTZip {
    <# Extracts a PSADT .zip to a temp folder, finds the root containing Deploy-Application.ps1 #>
    param([string]$ZipPath)

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $TempDir = Join-Path ([System.IO.Path]::GetTempPath()) "PSADT_Import_$([System.IO.Path]::GetRandomFileName())"
    New-Item -Path $TempDir -ItemType Directory -Force | Out-Null
    Write-DebugLog "Extracting PSADT zip to: $TempDir"

    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $TempDir)
    } catch {
        Write-DebugLog "Failed to extract zip: $($_.Exception.Message)" -Level 'ERROR'
        Show-ThemedDialog -Title 'Import Error' -Message "Failed to extract ZIP file:`n$($_.Exception.Message)" `
            -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeError' | Out-Null
        Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue
        return $null
    }

    # Find Deploy-Application.ps1 — could be at root or one level deep
    $ScriptFile = Get-ChildItem $TempDir -Filter 'Deploy-Application.ps1' -Recurse -Depth 2 | Select-Object -First 1
    if (-not $ScriptFile) {
        Write-DebugLog "No Deploy-Application.ps1 found in zip" -Level 'WARN'
        Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue
        Show-ThemedDialog -Title 'Import Error' -Message "No Deploy-Application.ps1 found inside the ZIP.`nThis doesn't appear to be a valid PSADT package." `
            -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' | Out-Null
        return $null
    }

    return [PSCustomObject]@{
        FolderPath = $ScriptFile.DirectoryName
        TempRoot   = $TempDir
    }
}

function New-PSADTZipForUpload {
    <# Creates a zip of the entire PSADT package folder for WinGet upload #>
    param(
        [string]$SourceFolder,
        [string]$AppName,
        [string]$AppVersion
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $CleanName = ($AppName -replace '[^a-zA-Z0-9\-_]', '_').Trim('_')
    $ZipName = "${CleanName}_${AppVersion}_PSADT.zip"
    $OutputDir = Join-Path $Global:Root "psadt_packages"
    if (-not (Test-Path $OutputDir)) {
        New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
    }
    $ZipPath = Join-Path $OutputDir $ZipName

    if (Test-Path $ZipPath) { Remove-Item $ZipPath -Force }

    Write-DebugLog "Creating PSADT zip: $ZipPath from $SourceFolder"
    try {
        [System.IO.Compression.ZipFile]::CreateFromDirectory(
            $SourceFolder, $ZipPath,
            [System.IO.Compression.CompressionLevel]::Optimal, $false
        )
        $SizeMB = [math]::Round((Get-Item $ZipPath).Length / 1MB, 2)
        Write-DebugLog "PSADT zip created: $ZipPath ($SizeMB MB)"
        return $ZipPath
    } catch {
        Write-DebugLog "Failed to create zip: $($_.Exception.Message)" -Level 'ERROR'
        Show-ThemedDialog -Title 'Zip Error' -Message "Failed to create PSADT zip:`n$($_.Exception.Message)" `
            -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeError' | Out-Null
        return $null
    }
}

function Populate-CreateTabFromPSADT {
    <# Takes parsed PSADT data and fills the Create tab form fields #>
    param(
        [PSCustomObject]$PsadtData,
        [ValidateSet('raw','psadt')][string]$Mode = 'raw'
    )

    # Package identity
    $Vendor = if ($PsadtData.AppVendor) { $PsadtData.AppVendor } else { "MyOrg" }
    $CleanName = $PsadtData.AppName -replace '[^a-zA-Z0-9\-]', ''
    $txtPkgId.Text      = "$Vendor.$CleanName"
    $txtPkgVersion.Text = $PsadtData.AppVersion

    # Metadata
    $txtPkgName.Text    = $PsadtData.AppName
    $txtPublisher.Text  = $Vendor
    $txtAuthor.Text     = if ($PsadtData.AppAuthor) { $PsadtData.AppAuthor } else { "" }
    $txtShortDesc.Text  = "$($PsadtData.AppName) $($PsadtData.AppVersion)"
    $txtLicense.Text    = "Proprietary"

    if ($Mode -eq 'psadt') {
        # ── PSADT zip mode: InstallerType=zip, nested=exe pointing to Deploy-Application.exe ──
        $ZipPath = New-PSADTZipForUpload -SourceFolder $PsadtData.SourceFolder `
                                         -AppName $PsadtData.AppName `
                                         -AppVersion $PsadtData.AppVersion
        if (-not $ZipPath) { return }

        Set-ComboBoxByContent $cmbInstallerType0 'zip'
        Set-ComboBoxByContent $cmbArch0 $PsadtData.AppArch
        Set-ComboBoxByContent $cmbScope0 'machine'

        # Configure nested installer — Deploy-Application.exe runs the full PSADT flow
        Set-ComboBoxByContent $cmbNestedType0 'exe'
        $txtNestedPath0.Text = 'Deploy-Application.exe'
        $pnlNestedInstaller.Visibility = 'Visible'

        # Silent switches apply to Deploy-Application.exe, not the inner MSI
        $txtSilent0.Text         = '-DeploymentType "Install" -DeployMode "Silent"'
        $txtSilentProgress0.Text = '-DeploymentType "Install" -DeployMode "NonInteractive"'
        $txtInteractive0.Text    = '-DeploymentType "Install" -DeployMode "Interactive"'
        $txtCustomSwitch0.Text   = ''
        $txtLog0.Text            = ''
        $txtRepair0.Text         = ''
        $txtProductCode0.Text    = $PsadtData.ProductCode

        # Set zip as the upload file
        $txtLocalFile0.Text = $ZipPath
        $txtUploadFile.Text = $ZipPath
        Compute-FileHash -FilePath $ZipPath -TargetField 'installer0'
        Compute-FileHash -FilePath $ZipPath -TargetField 'upload'

        $lblStatus.Text = "Imported PSADT (zip mode): $($PsadtData.AppName) v$($PsadtData.AppVersion)"
    } else {
        # ── Raw mode: just the MSI/EXE from Files/ ──
        if ($PsadtData.InstallerType) { Set-ComboBoxByContent $cmbInstallerType0 $PsadtData.InstallerType }
        Set-ComboBoxByContent $cmbArch0 $PsadtData.AppArch
        Set-ComboBoxByContent $cmbScope0 'machine'

        $txtSilent0.Text         = $PsadtData.SilentSwitch
        $txtSilentProgress0.Text = $PsadtData.SilentProgressSwitch
        $txtProductCode0.Text    = $PsadtData.ProductCode

        # Hide nested installer panel
        $pnlNestedInstaller.Visibility = 'Collapsed'
        $txtNestedPath0.Text = ''

        if ($PsadtData.InstallerFilePath -and (Test-Path $PsadtData.InstallerFilePath)) {
            $txtLocalFile0.Text = $PsadtData.InstallerFilePath
            $txtUploadFile.Text = $PsadtData.InstallerFilePath
            Compute-FileHash -FilePath $PsadtData.InstallerFilePath -TargetField 'installer0'
            Compute-FileHash -FilePath $PsadtData.InstallerFilePath -TargetField 'upload'
        }

        $lblStatus.Text = "Imported PSADT (raw installer): $($PsadtData.AppName) v$($PsadtData.AppVersion)"
    }

    # Common fields for both modes
    if ($PsadtData.Commands.Count -gt 0) {
        $txtCommands0.Text = ($PsadtData.Commands -join ', ')
    }
    $chkModeInteractive.IsChecked   = $true
    $chkModeSilent.IsChecked         = $true
    $chkModeSilentProgress.IsChecked = $true
    $cmbInstallerPreset.SelectedIndex = 0

    # ── Uninstall commands ──
    if ($Mode -eq 'psadt') {
        # In PSADT zip mode, uninstall goes through Deploy-Application.exe
        $txtUninstallCmd.Text    = 'Deploy-Application.exe -DeploymentType "Uninstall" -DeployMode "Interactive"'
        $txtUninstallSilent.Text = 'Deploy-Application.exe -DeploymentType "Uninstall" -DeployMode "Silent"'
    } elseif ($PsadtData.UninstallCmds.Count -gt 0) {
        # In raw mode, use the actual uninstall commands parsed from the script
        $txtUninstallCmd.Text    = $PsadtData.UninstallCmds[0]
        $txtUninstallSilent.Text = ''
    }

    # ── Upgrade behavior: if PSADT has an uninstall section, use uninstallPrevious ──
    if ($PsadtData.HasUninstallSection) {
        Set-ComboBoxByContent $cmbUpgrade0 'uninstallPrevious'
    }

    # ── PSADT always requires elevation ──
    Set-ComboBoxByContent $cmbElevation0 'elevationRequired'

    # ── Copyright from vendor ──
    if ($PsadtData.AppVendor) {
        $txtCopyright.Text = "Copyright (c) $($PsadtData.AppVendor)"
    }

    # ── Description: build a richer description from available PSADT metadata ──
    $DescParts = @()
    $DescTitle = if ($PsadtData.InstallTitle) { $PsadtData.InstallTitle.Trim() } else { $PsadtData.AppName }
    $DescParts += $DescTitle
    if ($PsadtData.AppVendor) { $DescParts += "by $($PsadtData.AppVendor)" }
    if ($PsadtData.HasRepairSection) { $DescParts += '(supports repair)' }
    $txtDescription.Text = ($DescParts -join ' ')

    # ── Tags from app name, vendor, and language ──
    $TagList = @()
    if ($PsadtData.AppName)   { $TagList += $PsadtData.AppName -replace '\s+', '-' }
    if ($PsadtData.AppVendor) { $TagList += $PsadtData.AppVendor -replace '\s+', '-' }
    $TagList += 'PSADT'
    if ($PsadtData.AppLang)   { $TagList += $PsadtData.AppLang }
    $txtTags.Text = ($TagList -join ', ')

    # ── Moniker: lowercase short name ──
    if ($PsadtData.AppName) {
        $txtMoniker.Text = ($PsadtData.AppName -replace '[^a-zA-Z0-9]', '').ToLower()
    }

    # ── File extensions ──
    if ($PsadtData.FileExtensions.Count -gt 0) {
        $txtFileExt0.Text = ($PsadtData.FileExtensions -join ', ')
    }

    Update-YamlPreview
    Write-DebugLog "Populated Create tab from PSADT ($Mode): $($PsadtData.AppName) v$($PsadtData.AppVersion)"
    $statusDot.Fill = $Window.Resources['ThemeSuccess']
}

# ==============================================================================
# SECTION 8: PREREQUISITE CHECK UI
# ==============================================================================

$Global:PrereqsPassed = $false

function Show-PrereqResults {
    <# Displays prerequisite check results in the UI #>
    param([array]$Results)
    $pnlModuleStatus.Children.Clear()
    $AllOK = $true
    foreach ($R in $Results) {
        $SP = New-Object System.Windows.Controls.StackPanel
        $SP.Orientation = 'Horizontal'
        $SP.Margin = [System.Windows.Thickness]::new(0, 2, 0, 2)

        $Icon = New-Object System.Windows.Controls.TextBlock
        $Icon.FontSize = 12
        $Icon.Margin = [System.Windows.Thickness]::new(0, 0, 8, 0)
        $Icon.VerticalAlignment = 'Center'

        $Label = New-Object System.Windows.Controls.TextBlock
        $Label.FontSize = 12
        $Label.VerticalAlignment = 'Center'

        if ($R.Available) {
            $Icon.Text = [char]0x2713  # checkmark
            $Icon.Foreground = $Window.Resources['ThemeSuccess']
            $Label.Text = "$($R.Name) v$($R.Installed)"
            $Label.Foreground = $Window.Resources['ThemeTextBody']
        } else {
            $Icon.Text = [char]0x2717  # cross
            $Icon.Foreground = $Window.Resources['ThemeError']
            $Label.Text = "$($R.Name) — NOT FOUND"
            $Label.Foreground = $Window.Resources['ThemeError']
            $AllOK = $false
        }
        $SP.Children.Add($Icon) | Out-Null
        $SP.Children.Add($Label) | Out-Null
        $pnlModuleStatus.Children.Add($SP) | Out-Null
    }

    if ($AllOK) {
        $Global:PrereqsPassed = $true
        $pnlPrereqError.Visibility = 'Collapsed'
        $pnlMainContent.Visibility = 'Visible'
        $btnLogin.IsEnabled = $true
        $btnPublish.IsEnabled = $false  # Enabled after auth
        Write-DebugLog "All prerequisites met"
    } else {
        $Global:PrereqsPassed = $false
        $pnlPrereqError.Visibility = 'Visible'
        $pnlMainContent.Visibility = 'Collapsed'
        $btnLogin.IsEnabled = $false
        $Cmd = Get-RemediationCommand -ModuleResults $Results
        $txtPrereqCommand.Text = $Cmd
        Write-DebugLog "Missing prerequisites — publishing disabled" -Level 'WARN'
    }
}

# ==============================================================================
# SECTION 9: AZURE AUTHENTICATION (Ported from AIBLogMonitor)
# ==============================================================================

function Get-ActiveRestSource {
    <# Returns the currently selected REST source function name from the global combo or settings fallback #>
    $Sel = $cmbRestSource.SelectedItem
    if ($Sel -and $Sel.Tag) { return [string]$Sel.Tag }
    $Fallback = $txtRestSourceName.Text.Trim()
    if ($Fallback) { Write-DebugLog "Get-ActiveRestSource: using Settings fallback '$Fallback'" -Level 'DEBUG' }
    return $Fallback
}

function Populate-RestSourceCombo {
    <# Populates the global REST source ComboBox from discovered Function Apps and saved sources #>
    param([array]$FuncApps = @())
    $cmbRestSource.Items.Clear()
    $Selected = $null

    # Add discovered Function Apps
    foreach ($FA in $FuncApps) {
        $Item = New-Object System.Windows.Controls.ComboBoxItem
        $Item.Content = "$($FA.Name)"
        $Item.Tag = $FA.Name
        $cmbRestSource.Items.Add($Item) | Out-Null
    }

    # Add saved sources not already in discovered list
    $DiscoveredNames = @($FuncApps | ForEach-Object { $_.Name })
    if ($cmbSavedSources) {
        foreach ($SavedItem in $cmbSavedSources.Items) {
            $SName = if ($SavedItem.Tag) { $SavedItem.Tag } elseif ($SavedItem.Content) { $SavedItem.Content } else { $null }
            if ($SName -and $SName -notin $DiscoveredNames) {
                $Item = New-Object System.Windows.Controls.ComboBoxItem
                $Item.Content = "$SName"
                $Item.Tag = $SName
                $cmbRestSource.Items.Add($Item) | Out-Null
            }
        }
    }

    # Also add current settings textbox value if not already present
    $Current = $txtRestSourceName.Text.Trim()
    if ($Current) {
        $AllTags = @()
        foreach ($ci in $cmbRestSource.Items) { if ($ci.Tag) { $AllTags += $ci.Tag } }
        if ($Current -notin $AllTags) {
            $Item = New-Object System.Windows.Controls.ComboBoxItem
            $Item.Content = "$Current"
            $Item.Tag = $Current
            $cmbRestSource.Items.Add($Item) | Out-Null
        }
    }

    # Select the one matching the settings textbox value (or first)
    if ($cmbRestSource.Items.Count -gt 0) {
        $cmbRestSource.Visibility = 'Visible'
        for ($i = 0; $i -lt $cmbRestSource.Items.Count; $i++) {
            if ($cmbRestSource.Items[$i].Tag -eq $Current) {
                $cmbRestSource.SelectedIndex = $i
                $Selected = $Current
                break
            }
        }
        if (-not $Selected) { $cmbRestSource.SelectedIndex = 0 }
    }
}

function Set-AuthUIConnected {
    <# Updates UI elements to the "connected" state #>
    param(
        [string]$AccountId,
        [string]$SubName,
        [string]$SubId,
        [array]$Subscriptions = @()
    )
    Write-DebugLog "Set-AuthUIConnected: acct=$AccountId sub=$SubName subs=$($Subscriptions.Count)"
    $Global:IsAuthenticated = $true
    $authDot.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, 'ThemeSuccess')
    $lblAuthStatus.Text = $AccountId
    $lblAuthStatus.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextPrimary')

    $cmbSubscription.Items.Clear()
    $Global:SuppressSubChangeHandler = $true
    $DefaultSubName = if ($cmbDefaultSubscription -and $cmbDefaultSubscription.Text) { $cmbDefaultSubscription.Text.Trim() } else { '' }
    $SelectedIdx = -1
    if ($Subscriptions.Count -gt 0) {
        # Populate the settings default-subscription dropdown with discovered subs
        $PrevDefault = $cmbDefaultSubscription.Text
        $cmbDefaultSubscription.Items.Clear()
        for ($i = 0; $i -lt $Subscriptions.Count; $i++) {
            $DI = New-Object System.Windows.Controls.ComboBoxItem
            $DI.Content = [string]$Subscriptions[$i].Name
            $DI.Tag     = [string]$Subscriptions[$i].Id
            $cmbDefaultSubscription.Items.Add($DI) | Out-Null
        }
        if ($PrevDefault) { $cmbDefaultSubscription.Text = $PrevDefault }

        for ($i = 0; $i -lt $Subscriptions.Count; $i++) {
            $Item = New-Object System.Windows.Controls.ComboBoxItem
            $Item.Content = [string]$Subscriptions[$i].Name
            $Item.Tag     = [string]$Subscriptions[$i].Id
            $cmbSubscription.Items.Add($Item) | Out-Null
            # Prefer default subscription from settings, then fall back to Azure context
            if ($DefaultSubName -and $Subscriptions[$i].Name -ieq $DefaultSubName) {
                $SelectedIdx = $i
            } elseif ($SelectedIdx -lt 0 -and $Subscriptions[$i].Id -eq $SubId) {
                $SelectedIdx = $i
            }
        }
        if ($SelectedIdx -ge 0) { $cmbSubscription.SelectedIndex = $SelectedIdx }
    } else {
        $Item = New-Object System.Windows.Controls.ComboBoxItem
        $Item.Content = [string]$SubName
        $Item.Tag     = [string]$SubId
        $cmbSubscription.Items.Add($Item) | Out-Null
        $cmbSubscription.SelectedIndex = 0
    }
    $Global:SuppressSubChangeHandler = $false

    # If default subscription from settings differs from Azure context, switch to it
    $SelItem = $cmbSubscription.SelectedItem
    if ($SelItem -and $SelItem.Tag -and $SelItem.Tag -ne $SubId) {
        Write-DebugLog "Default subscription '$($SelItem.Content)' differs from context — switching"
        Switch-Subscription -SubId $SelItem.Tag
    }

    $cmbSubscription.Visibility = 'Visible'
    $btnLogin.Visibility = 'Collapsed'
    $btnLogout.Visibility = 'Visible'
    $lblStatus.Text = "Connected as $AccountId"
    $statusDot.Fill = $Window.Resources['ThemeSuccess']

    # Update deploy/remove target labels with current subscription
    $lblDeployTargetSub.Text = "$SubName ($SubId)"
    $lblRemoveTargetSub.Text = "$SubName ($SubId)"

    # Enable publishing
    $btnPublish.IsEnabled = $true
    $btnEditPublish.IsEnabled = $true
    $btnUpload.IsEnabled = $true
    Write-DebugLog "Set-AuthUIConnected: done"
}

function Set-AuthUIDisconnected {
    <# Updates UI elements to the "disconnected" state #>
    param([string]$Reason = "Not connected")
    Write-DebugLog "Set-AuthUIDisconnected: reason=$Reason"
    $Global:IsAuthenticated = $false
    $authDot.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, 'ThemeError')
    $lblAuthStatus.Text = $Reason
    $lblAuthStatus.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
    $cmbSubscription.Items.Clear()
    $cmbSubscription.Visibility = 'Collapsed'
    $cmbRestSource.Items.Clear()
    $cmbRestSource.Visibility = 'Collapsed'
    $btnLogin.Visibility = 'Visible'
    $btnLogout.Visibility = 'Collapsed'
    $statusDot.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, 'ThemeTextDim')

    # Reset deploy/remove target labels
    $lblDeployTargetSub.Text = 'Not connected'
    $lblRemoveTargetSub.Text = 'Not connected'

    # Disable publishing
    $btnPublish.IsEnabled = $false
    $btnEditPublish.IsEnabled = $false
    $btnUpload.IsEnabled = $false
}

function Connect-ToAzure {
    <# Interactive Azure login via browser auth (WAM disabled) — runs on UI thread #>
    $lblAuthStatus.Text = "Signing in..."
    $lblStatus.Text = "Authenticating with Entra ID..."
    $Window.Cursor = [System.Windows.Input.Cursors]::Wait
    $btnLogin.IsEnabled = $false
    Write-DebugLog "Connect-ToAzure: >>> ENTER"

    try {
        $Context = Get-AzContext -ErrorAction SilentlyContinue

        $NeedLogin = $true
        if ($Context -and $Context.Account) {
            $lblAuthStatus.Text = "Validating token for $($Context.Account.Id)..."
            Write-DebugLog "Connect-ToAzure: cached context exists for $($Context.Account.Id), validating..."
            try {
                $null = Get-AzSubscription -SubscriptionId $Context.Subscription.Id -ErrorAction Stop -WarningAction SilentlyContinue
                Write-DebugLog "Connect-ToAzure: cached token is valid, skipping re-auth"
                $NeedLogin = $false
            } catch {
                Write-DebugLog "Connect-ToAzure: cached token expired/invalid: $($_.Exception.Message)"
            }
        }

        if ($NeedLogin) {
            Write-DebugLog "Connect-ToAzure: disabling WAM, using interactive browser..."
            $lblAuthStatus.Text = "Opening browser for sign-in..."
            $lblStatus.Text = "Browser sign-in (check your default browser)..."
            try { Update-AzConfig -EnableLoginByWam $false -ErrorAction SilentlyContinue | Out-Null } catch { Write-DebugLog "WAM disable skipped: $($_.Exception.Message)" -Level 'DEBUG' }
            $ConnectParams = @{ WarningAction = 'SilentlyContinue'; ErrorAction = 'Stop' }
            if ($Context -and $Context.Tenant.Id) {
                $ConnectParams['TenantId'] = $Context.Tenant.Id
            }
            Connect-AzAccount @ConnectParams | Out-Null
        }

        $Context = Get-AzContext
        if ($Context -and $Context.Account) {
            $lblAuthStatus.Text = $Context.Account.Id
            $lblStatus.Text = "Signed in — loading subscriptions..."

            $PostAcctId  = $Context.Account.Id
            $PostSubName = $Context.Subscription.Name
            $PostSubId   = $Context.Subscription.Id

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
                } catch { Write-DebugLog "Subscription query failed: $($_.Exception.Message)" -Level 'WARN' }
                return $R
            } -OnComplete {
                param($Results, $Errors)
                $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ AccountId = ''; SubName = ''; SubId = ''; Subs = @() } }
                Write-DebugLog "Connect-ToAzure OnComplete: $($R.Subs.Count) subscription(s) found"
                Set-AuthUIConnected -AccountId $R.AccountId -SubName $R.SubName -SubId $R.SubId -Subscriptions $R.Subs
                # Discover resources after auth
                Discover-StorageAccounts
                Discover-RestSources
                $Window.Cursor = [System.Windows.Input.Cursors]::Arrow
            }
        } else {
            $lblAuthStatus.Text = "Sign-in cancelled"
            $lblStatus.Text = "Sign-in was cancelled"
            $Window.Cursor = [System.Windows.Input.Cursors]::Arrow
        }
    } catch {
        $authDot.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, 'ThemeError')
        $lblAuthStatus.Text = "Sign-in failed"
        $lblAuthStatus.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeError')
        $lblStatus.Text = "Auth error: $($_.Exception.Message)"
        Write-DebugLog "Connect-ToAzure EXCEPTION: $($_.Exception.Message)" -Level 'ERROR'
        $Window.Cursor = [System.Windows.Input.Cursors]::Arrow
    } finally {
        $btnLogin.IsEnabled = $true
    }
}

function Disconnect-FromAzure {
    <# Logs out of Azure and resets UI state #>
    try { Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null } catch { Write-DebugLog "Disconnect-AzAccount: $($_.Exception.Message)" -Level 'DEBUG' }
    Set-AuthUIDisconnected -Reason "Not connected"
    $lblStatus.Text = "Disconnected"
    $tvStorage.Items.Clear()
    $cmbUploadStorage.Items.Clear()
    $cmbUploadContainer.Items.Clear()
    $Global:StorageAccounts = @()
}

# ==============================================================================
# SECTION 10: SUBSCRIPTION & RESOURCE DISCOVERY
# ==============================================================================

# Subscription change handler — refresh storage when user picks a different sub
function Switch-Subscription {
    param([string]$SubId)
    if (-not $SubId) { return }
    Write-DebugLog "Switching subscription to: $SubId"
    $lblStatus.Text = "Switching subscription..."
    Show-GlobalProgressIndeterminate -Text 'Switching subscription...'
    Start-BackgroundWork -Variables @{ SId = $SubId } -Work {
        Set-AzContext -SubscriptionId $SId -ErrorAction Stop | Out-Null
        return @{ SubId = $SId }
    } -OnComplete {
        param($Results, $Errors)
        Write-DebugLog "Subscription switched, refreshing resources..."
        Hide-GlobalProgress
        Discover-StorageAccounts
        Discover-RestSources
    }
}

# ==============================================================================
# SECTION 11: USER PREFERENCES (Load/Save)
# ==============================================================================

$Global:LeftPanelOpen = $true
$Global:RightPanelOpen = $true  # kept for prefs compat

function Load-UserPrefs {
    <# Loads saved preferences and applies to UI #>
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
        if ($null -ne $P.DebugEnabled)    { $chkDebug.IsChecked = $P.DebugEnabled }
        if ($null -ne $P.AutoSave)           { $chkAutoSave.IsChecked = $P.AutoSave }
        if ($null -ne $P.ConfirmUpload)      { $chkConfirmUpload.IsChecked = $P.ConfirmUpload }
        if ($null -ne $P.DisableAnimations) {
            $chkDisableAnimations.IsChecked = $P.DisableAnimations
            $Global:AnimationsDisabled = $P.DisableAnimations
        }
        if ($null -ne $P.RestSourceName)  { $txtRestSourceName.Text = $P.RestSourceName }
        if ($null -ne $P.SavedSources -and $P.SavedSources.Count -gt 0) {
            $cmbSavedSources.Items.Clear()
            foreach ($Src in $P.SavedSources) {
                $Item = New-Object System.Windows.Controls.ComboBoxItem
                $Item.Content = $Src
                $Item.Tag = $Src
                $cmbSavedSources.Items.Add($Item) | Out-Null
            }
        }
        # Seed global REST source combo from saved sources (discovery will enrich later)
        Populate-RestSourceCombo
        if ($null -ne $P.ManifestPath)    { $txtManifestPath.Text = $P.ManifestPath }
        if ($null -ne $P.DefaultContainer)   { $txtDefaultContainer.Text = $P.DefaultContainer }
        if ($null -ne $P.DiscoveryPattern)    { $txtDiscoveryPattern.Text = $P.DiscoveryPattern }
        if ($null -ne $P.CacheTTL) {
            $txtCacheTTL.Text = [string]$P.CacheTTL
            $Global:ManageCacheTTL = [int]$P.CacheTTL
        }
        if ($null -ne $P.DefaultSubscription) { $cmbDefaultSubscription.Text = $P.DefaultSubscription }
        if ($null -ne $P.LastStorageAccount) { $Global:PreferredStorageAccount = $P.LastStorageAccount }
        if ($null -ne $P.ManifestVersion) {
            for ($i = 0; $i -lt $cmbManifestVersion.Items.Count; $i++) {
                if ($cmbManifestVersion.Items[$i].Content -eq $P.ManifestVersion) {
                    $cmbManifestVersion.SelectedIndex = $i
                    break
                }
            }
        }

        # Backup/Restore preferences
        if ($null -ne $P.BackupStorageAcct -and $P.BackupStorageAcct) { $cmbBackupStorageAcct.Text = $P.BackupStorageAcct }
        if ($null -ne $P.BackupFolder -and $P.BackupFolder) { $txtBackupFolder.Text = $P.BackupFolder }

        # Achievements collapsed state
        if ($null -ne $P.AchievementsCollapsed -and $P.AchievementsCollapsed) {
            $Global:AchievementsCollapsed = $true
            if ($pnlAchievements) { $pnlAchievements.Visibility = 'Collapsed' }
            if ($txtAchievementChevron) { $txtAchievementChevron.RenderTransform = [System.Windows.Media.RotateTransform]::new(0) }
        }

        # Left panel state
        if ($null -ne $P.LeftPanelOpen -and $P.LeftPanelOpen -eq $false) {
            Toggle-LeftPanel $false
        }
        # Right panel state
        if ($null -ne $P.BottomPanelOpen -and $P.BottomPanelOpen -eq $false) {
            Toggle-BottomPanel $false
        }

        # Window position/size (validate against virtual screen bounds)
        if ($null -ne $P.WindowState) {
            if ($P.WindowState -ne 'Maximized' -and $null -ne $P.WindowLeft) {
                $VW = [System.Windows.SystemParameters]::VirtualScreenWidth
                $VH = [System.Windows.SystemParameters]::VirtualScreenHeight
                $VL = [System.Windows.SystemParameters]::VirtualScreenLeft
                $VT = [System.Windows.SystemParameters]::VirtualScreenTop
                $PL = $P.WindowLeft; $PT = $P.WindowTop; $PW = $P.WindowWidth; $PH = $P.WindowHeight
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

        Write-DebugLog "User preferences loaded" -Level 'DEBUG'
    } catch {
        Write-DebugLog "Failed to load preferences: $($_.Exception.Message)" -Level 'WARN'
    }
}

function Save-UserPrefs {
    <# Persists current UI state to JSON #>
    try {
        $P = [PSCustomObject]@{
            IsLightMode      = $Global:IsLightMode
            DebugEnabled     = [bool]$chkDebug.IsChecked
            AutoSave         = [bool]$chkAutoSave.IsChecked
            ConfirmUpload    = [bool]$chkConfirmUpload.IsChecked
            DisableAnimations = $Global:AnimationsDisabled
            RestSourceName   = $txtRestSourceName.Text
            SavedSources     = @($cmbSavedSources.Items | ForEach-Object { $_.Tag })
            ManifestPath     = $txtManifestPath.Text
            ManifestVersion  = Get-ComboBoxSelectedText $cmbManifestVersion
            DefaultContainer = $txtDefaultContainer.Text
            DiscoveryPattern = $txtDiscoveryPattern.Text
            CacheTTL         = $Global:ManageCacheTTL
            DefaultSubscription = $cmbDefaultSubscription.Text
            LastStorageAccount = if ($cmbUploadStorage.SelectedItem -and $cmbUploadStorage.SelectedItem.Tag) {
                $cmbUploadStorage.SelectedItem.Tag.Name
            } else { '' }
            LeftPanelOpen    = $Global:LeftPanelOpen
            BottomPanelOpen  = $Global:BottomPanelOpen
            AchievementsCollapsed = $Global:AchievementsCollapsed
            BackupStorageAcct = $cmbBackupStorageAcct.Text
            BackupFolder     = $txtBackupFolder.Text
            WindowState      = $Window.WindowState.ToString()
            WindowLeft       = $Window.Left
            WindowTop        = $Window.Top
            WindowWidth      = $Window.Width
            WindowHeight     = $Window.Height
        }
        $P | ConvertTo-Json | Set-Content $Global:PrefsPath -Force
        Write-DebugLog "User preferences saved" -Level 'DEBUG'
    } catch {
        Write-DebugLog "Failed to save preferences: $($_.Exception.Message)" -Level 'WARN'
    }
}

# ==============================================================================
# SECTION 12: DISPATCHER TIMER (Queue → UI Bridge)
# ==============================================================================

$Timer = New-Object System.Windows.Threading.DispatcherTimer
$Timer.Interval = [TimeSpan]::FromMilliseconds($Script:TIMER_INTERVAL_MS)
$Timer.Add_Tick({
    if ($Global:TimerProcessing) { return }
    $Global:TimerProcessing = $true
    try {
    # Use [datetime]::Now instead of Get-Date — something in the Az/Graph module
    # load chain occasionally breaks cmdlet resolution in the WPF dispatcher thread,
    # causing 'Get-Date is not recognized' to spam every 50ms and prevent background
    # job completion callbacks from firing (uploads appear to hang).
    $TickNow = [datetime]::Now

    # ── Background job completion polling ───────────────────────────────────
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
                Write-DebugLog "BgJob[$bi]: invoking OnComplete callback" -Level 'DEBUG'
                & $Job.OnComplete $BgResult $BgErrors $Job.Context
                Write-DebugLog "BgJob[$bi]: OnComplete callback finished OK" -Level 'DEBUG'
            } catch {
                Write-DebugLog "BgJob[$bi]: callback EXCEPTION: $($_.Exception.Message)" -Level 'ERROR'
                Write-DebugLog "BgJob[$bi]: ScriptStackTrace: $($_.ScriptStackTrace)" -Level 'ERROR'
                Write-DebugLog "BgJob[$bi]: InnerException: $($_.Exception.InnerException)" -Level 'ERROR'
            }
            try { $Job.PS.Dispose() } catch { Write-DebugLog "BgJob dispose PS: $($_.Exception.Message)" -Level 'DEBUG' }
            try { $Job.Runspace.Dispose() } catch { Write-DebugLog "BgJob dispose Runspace: $($_.Exception.Message)" -Level 'DEBUG' }
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

    # ── Download progress update ─────────────────────────────────────────────
    $DlPct = $Global:SyncHash.DownloadProgress
    if ($DlPct -ge 0) {
        $DlBytes = $Global:SyncHash.DownloadBytes
        $DlTotal = $Global:SyncHash.DownloadTotal
        $DlMB = [math]::Round($DlBytes / 1MB, 1)
        if ($DlTotal -gt 0) {
            $TotalMB = [math]::Round($DlTotal / 1MB, 1)
            $lblHash0.Text = "Downloading... $DlPct% ($DlMB / $TotalMB MB)"
            $lblStatus.Text = "Downloading installer... $DlPct%"
        } else {
            $lblHash0.Text = "Downloading... $DlMB MB"
            $lblStatus.Text = "Downloading installer..."
        }
    } elseif ($DlPct -eq -2) {
        $lblHash0.Text = "Computing SHA256 hash..."
        $lblStatus.Text = "Computing SHA256 hash..."
    }

    # ── Upload progress update ───────────────────────────────────────────────
    $UlPct = $Global:SyncHash.UploadProgress
    if ($UlPct -ge 0) {
        $UlBytes = $Global:SyncHash.UploadBytes
        $UlTotal = $Global:SyncHash.UploadTotal
        $UlMB = [math]::Round($UlBytes / 1MB, 1)
        $TotalMB = [math]::Round($UlTotal / 1MB, 1)
        $prgUpload.Value = $UlPct
        $lblUploadProgress.Text = "$UlMB / $TotalMB MB ($UlPct%)"
        $lblStatus.Text = "Uploading... $UlMB / $TotalMB MB"
    }

    # ── Batch progress update ────────────────────────────────────────────────
    if ($Global:SyncHash.BatchTotal -and $Global:SyncHash.BatchTotal -gt 0) {
        $Done  = $Global:SyncHash.BatchDone
        $Total = $Global:SyncHash.BatchTotal
        $Current = $Global:SyncHash.BatchCurrent
        if ($Done -lt $Total) {
            $prgBatch.Value = $Done
            $lblBatchProgress.Text = "Processing $($Done+1)/$Total — $Current"
            $lblStatus.Text = "Batch: $($Done+1)/$Total — $Current"
        }
    }

    # ── Status queue drain ──────────────────────────────────────────────────
    while ($Global:SyncHash.StatusQueue.Count -gt 0) {
        $Msg = $Global:SyncHash.StatusQueue.Dequeue()
        if ($Msg.Type -eq 'HashResult') {
            $SizeMB = [math]::Round($Msg.Size / 1MB, 2)
            if ($Msg.Field -eq 'installer0') {
                $lblHash0.Text = "SHA256: $($Msg.Value)"
                $Global:CurrentHash0 = $Msg.Value
            } elseif ($Msg.Field -eq 'upload') {
                $lblUploadHash.Text = $Msg.Value
                $lblUploadSize.Text = "$SizeMB MB - $($Msg.Name)"
                $btnCopyHash.IsEnabled = $true
            }
        }
    }

    # ── Deploy output queue drain ───────────────────────────────────────────
    if ($Global:DeployOutputQueue) {
        $Msg = $null
        while ($Global:DeployOutputQueue.TryDequeue([ref]$Msg)) {
            if ($Msg -match '^__DEPLOY_EXIT__:(.*)$') {
                $ExitCode = $Matches[1]
                $Level = if ($ExitCode -eq '0') { 'SUCCESS' } else { 'ERROR' }
                Write-DebugLog "Deploy process exited with code $ExitCode" -Level $Level
                if ($Global:DeployContext -eq 'Remove') {
                    # Remove operation finished
                    $lblRemoveStatus.Text = if ($ExitCode -eq '0') { 'Removal completed successfully' } else { "Removal failed (exit code $ExitCode)" }
                    $prgRemove.Visibility = 'Collapsed'
                    $btnRemoveCancel.Visibility = 'Collapsed'
                    $btnRunRemove.IsEnabled = $true
                } else {
                    $lblDeployStatus.Text = if ($ExitCode -eq '0') { 'Completed successfully' } else { "Failed (exit code $ExitCode)" }
                    $prgDeploy.Visibility = 'Collapsed'
                    $btnDeployCancel.Visibility = 'Collapsed'
                    $btnRunDeploy.IsEnabled = $true
                }
                $Global:DeployProcess = $null
                $Global:DeployContext = $null
                Hide-GlobalProgress
                # Auto-save deployed source to saved sources list
                if ($ExitCode -eq '0' -and $txtDeployName.Text.Trim()) {
                    $FuncName = "func-$($txtDeployName.Text.Trim())"
                    $Existing = $cmbSavedSources.Items | Where-Object { $_.Tag -eq $FuncName }
                    if (-not $Existing) {
                        $Item = New-Object System.Windows.Controls.ComboBoxItem
                        $Item.Content = $FuncName
                        $Item.Tag = $FuncName
                        $cmbSavedSources.Items.Add($Item) | Out-Null
                        Write-DebugLog "Auto-added '$FuncName' to saved sources" -Level 'INFO'
                        Save-UserPrefs
                    }
                }
            } elseif ($Msg -match '^ERR:(.*)$') {
                Write-DebugLog $Matches[1] -Level 'ERROR'
            } elseif ($Msg -match '^OUT:(.*)$') {
                $Line = $Matches[1]
                $Level = if ($Line -match '\[ERROR\]|FAIL|Exception') { 'ERROR' }
                         elseif ($Line -match '\[WARNING\]|WARN') { 'WARN' }
                         elseif ($Line -match '\[SUCCESS\]|COMPLETED|✓') { 'SUCCESS' }
                         elseif ($Line -match '\[DEBUG\]') { 'DEBUG' }
                         else { 'INFO' }
                Write-DebugLog $Line -Level $Level
            }
        }
    }

    } catch {
        # CRITICAL: a single tick error must not propagate to the dispatcher —
        # it triggers 'Global scope cannot be removed' which kills the entire
        # message loop (the app appears frozen and Add-Tick never runs again).
        try { Write-DebugLog "Timer tick exception: $($_.Exception.Message)" -Level 'ERROR' } catch {}
        try { Write-DebugLog "  ScriptStackTrace: $($_.ScriptStackTrace)" -Level 'DEBUG' } catch {}
    } finally { $Global:TimerProcessing = $false }
})

# ==============================================================================
# SECTION 13: EVENT HANDLERS
# ==============================================================================

# ── Title Bar ─────────────────────────────────────────────────────────────────
$btnMinimize.Add_Click({ $Window.WindowState = 'Minimized' })
$btnMaximize.Add_Click({
    if ($Window.WindowState -eq 'Maximized') { $Window.WindowState = 'Normal' }
    else { $Window.WindowState = 'Maximized' }
})
$btnClose.Add_Click({ $Window.Close() })

$btnThemeToggle.Add_Click({
    if ($Global:SuppressThemeHandler) { return }
    $NewLight = -not $Global:IsLightMode
    & $ApplyTheme $NewLight
    $btnThemeToggle.Content = if ($NewLight) { [char]0x263E } else { [char]0x2600 }
    $Global:SuppressThemeHandler = $true
    if ($chkDarkMode) { $chkDarkMode.IsChecked = -not $NewLight }
    $Global:SuppressThemeHandler = $false
    Unlock-Achievement 'theme_toggle'
    Render-Achievements
})

$chkDarkMode.Add_Checked({
    if ($Global:SuppressThemeHandler) { return }
    & $ApplyTheme $false
    $btnThemeToggle.Content = [char]0x2600
    Render-Achievements
})
$chkDarkMode.Add_Unchecked({
    if ($Global:SuppressThemeHandler) { return }
    & $ApplyTheme $true
    $btnThemeToggle.Content = [char]0x263E
    Render-Achievements
})

# ── Purge Downloads ─────────────────────────────────────────────────────
$btnPurgeDownloads = $Window.FindName("btnPurgeDownloads")
$btnPurgeDownloads.Add_Click({
    $DlDir = Join-Path $env:TEMP 'WinGetMM_Downloads'
    if (-not (Test-Path $DlDir)) {
        Show-Toast "No download cache found" -Type Info -DurationMs 2000
        return
    }
    $Files = Get-ChildItem $DlDir -Force -ErrorAction SilentlyContinue
    $Count = ($Files | Measure-Object).Count
    $SizeMB = [math]::Round(($Files | Measure-Object -Property Length -Sum).Sum / 1MB, 1)
    if ($Count -eq 0) {
        Show-Toast "Download cache is empty" -Type Info -DurationMs 2000
        return
    }
    $Confirm = Show-ThemedDialog -Title 'Purge Downloads' `
        -Message "Delete $Count cached file(s) ($SizeMB MB) from:`n$DlDir" `
        -Icon ([string]([char]0xE74D)) -IconColor 'ThemeWarning' -ShowCancel
    if ($Confirm -eq 'OK') {
        Remove-Item $DlDir -Recurse -Force -ErrorAction SilentlyContinue
        Write-DebugLog "Purged download cache: $Count files ($SizeMB MB)" -Level 'INFO'
        $lblStatus.Text = "Purged $Count downloaded file(s) ($SizeMB MB)"
        Show-Toast "Purged $Count file(s) ($SizeMB MB)" -Type Success -DurationMs 3000
    }
})

# ── Cache TTL ────────────────────────────────────────────────────────────────
$txtCacheTTL.Add_LostFocus({
    $Val = 0
    if ([int]::TryParse($txtCacheTTL.Text, [ref]$Val) -and $Val -ge 0) {
        $Global:ManageCacheTTL = $Val
        Write-DebugLog "Cache TTL set to ${Val}s" -Level 'DEBUG'
    } else {
        $txtCacheTTL.Text = [string]$Global:ManageCacheTTL
        Show-Toast "TTL must be a non-negative integer" -Type Warning -DurationMs 2000
    }
})

# ── Restore Defaults ─────────────────────────────────────────────────────────
$btnRestoreDefaults = $Window.FindName("btnRestoreDefaults")
if ($btnRestoreDefaults) {
    $btnRestoreDefaults.Add_Click({
        $Confirm = Show-ThemedDialog -Title 'Restore Defaults' `
            -Message 'Reset all settings to factory defaults? This cannot be undone.' `
            -Icon ([string]([char]0xE777)) -IconColor 'ThemeWarning' `
            -Buttons @(
                @{ Text='Cancel'; IsAccent=$false; Result='Cancel' }
                @{ Text='Reset'; IsDanger=$true; Result='Reset' }
            )
        if ($Confirm -ne 'Reset') { return }

        $txtRestSourceName.Text   = 'func-corpwinget'
        $txtManifestPath.Text     = ''
        $cmbManifestVersion.SelectedIndex = 0
        $txtDefaultContainer.Text = 'packages'
        $txtDiscoveryPattern.Text = '^func-|winget'
        $txtCacheTTL.Text         = '300'
        $Global:ManageCacheTTL    = 300
        $cmbDefaultSubscription.Text = ''
        $chkDarkMode.IsChecked    = $true
        $chkDebug.IsChecked       = $false
        $chkAutoSave.IsChecked    = $true
        $chkConfirmUpload.IsChecked = $true
        # Invalidate package cache
        $Global:ManageCacheTime   = $null
        $Global:ManageCacheSource = ''
        Update-ManageCacheIndicator
        Save-UserPrefs
        Write-DebugLog "Settings restored to defaults" -Level 'INFO'
        Show-Toast "Settings restored to defaults" -Type Success -DurationMs 3000
    })
}

# ── Auth ──────────────────────────────────────────────────────────────────────
$btnLogin.Add_Click({ Connect-ToAzure })
$btnLogout.Add_Click({ Disconnect-FromAzure })

$cmbSubscription.Add_SelectionChanged({
    if ($Global:SuppressSubChangeHandler) { return }
    $Sel = $cmbSubscription.SelectedItem
    if ($Sel -and $Sel.Tag) {
        Switch-Subscription -SubId $Sel.Tag
        $lblDeployTargetSub.Text = "$($Sel.Content) ($($Sel.Tag))"
        $lblRemoveTargetSub.Text = "$($Sel.Content) ($($Sel.Tag))"
    }
})

# ── Tab Navigation ────────────────────────────────────────────────────────────
function Switch-Tab {
    param([string]$TabName)
    $Global:ActiveTabName = $TabName
    # Hide all panels
    $pnlTabCreate.Visibility  = 'Collapsed'
    $pnlTabEdit.Visibility    = 'Collapsed'
    $pnlTabUpload.Visibility  = 'Collapsed'
    $pnlTabManage.Visibility  = 'Collapsed'
    $pnlTabDeploy.Visibility  = 'Collapsed'
    $pnlTabRemove.Visibility  = 'Collapsed'
    $pnlTabBackup.Visibility  = 'Collapsed'

    # Reset all tab visuals
    foreach ($btn in @($tabCreate, $tabEdit, $tabUpload, $tabManage, $tabDeploy, $tabRemove, $tabBackup)) {
        $btn.BorderBrush = [System.Windows.Media.Brushes]::Transparent
        $btn.Foreground = $Window.Resources['ThemeTextMuted']
        $btn.FontWeight = [System.Windows.FontWeights]::Normal
    }

    # Activate selected tab with fade transition
    switch ($TabName) {
        'Create' {
            Invoke-TabFade $pnlTabCreate
            $tabCreate.BorderBrush = $Window.Resources['ThemeAccent']
            $tabCreate.Foreground = $Window.Resources['ThemeTextPrimary']
            $tabCreate.FontWeight = [System.Windows.FontWeights]::SemiBold
        }
        'Edit' {
            Invoke-TabFade $pnlTabEdit
            $tabEdit.BorderBrush = $Window.Resources['ThemeAccent']
            $tabEdit.Foreground = $Window.Resources['ThemeTextPrimary']
            $tabEdit.FontWeight = [System.Windows.FontWeights]::SemiBold
        }
        'Upload' {
            Invoke-TabFade $pnlTabUpload
            $tabUpload.BorderBrush = $Window.Resources['ThemeAccent']
            $tabUpload.Foreground = $Window.Resources['ThemeTextPrimary']
            $tabUpload.FontWeight = [System.Windows.FontWeights]::SemiBold
            Update-BatchPanel
        }
        'Manage' {
            Invoke-TabFade $pnlTabManage
            $tabManage.BorderBrush = $Window.Resources['ThemeAccent']
            $tabManage.Foreground = $Window.Resources['ThemeTextPrimary']
            $tabManage.FontWeight = [System.Windows.FontWeights]::SemiBold
            # Update REST source label
            $FN = Get-ActiveRestSource
            $lblManageSourceName.Text = if ($FN) { "REST Source: $FN" } else { 'REST Source: (not set)' }
        }
        'Deploy' {
            Invoke-TabFade $pnlTabDeploy
            $tabDeploy.BorderBrush = $Window.Resources['ThemeAccent']
            $tabDeploy.Foreground = $Window.Resources['ThemeTextPrimary']
            $tabDeploy.FontWeight = [System.Windows.FontWeights]::SemiBold
        }
        'Remove' {
            Invoke-TabFade $pnlTabRemove
            $tabRemove.BorderBrush = $Window.Resources['ThemeAccent']
            $tabRemove.Foreground = $Window.Resources['ThemeTextPrimary']
            $tabRemove.FontWeight = [System.Windows.FontWeights]::SemiBold
        }
        'Backup' {
            Invoke-TabFade $pnlTabBackup
            $tabBackup.BorderBrush = $Window.Resources['ThemeAccent']
            $tabBackup.Foreground = $Window.Resources['ThemeTextPrimary']
            $tabBackup.FontWeight = [System.Windows.FontWeights]::SemiBold
            # Refresh backup/restore combos with latest discovered data
            Populate-BackupRestoreCombos
        }
    }
    Write-DebugLog "Switched to tab: $TabName"
}

$tabCreate.Add_Click({ Switch-Tab 'Create' })
$tabEdit.Add_Click({ Switch-Tab 'Edit' })
$tabUpload.Add_Click({ Switch-Tab 'Upload' })
$tabManage.Add_Click({ Switch-Tab 'Manage' })
$tabDeploy.Add_Click({ Switch-Tab 'Deploy' })
$tabRemove.Add_Click({ Switch-Tab 'Remove' })
$tabBackup.Add_Click({ Switch-Tab 'Backup' })

# ── Collapsible Section Handlers ──────────────────────────────────────────────
$secIdentityHeader.Add_MouseLeftButtonDown({
    Toggle-Section -Chevron $secIdentityChevron -Body $secIdentityBody
})
$secMetadataHeader.Add_MouseLeftButtonDown({
    Toggle-Section -Chevron $secMetadataChevron -Body $secMetadataBody
})
$secInstallerHeader.Add_MouseLeftButtonDown({
    Toggle-Section -Chevron $secInstallerChevron -Body $secInstallerBody
})
$secDepsHeader.Add_MouseLeftButtonDown({
    Toggle-Section -Chevron $secDepsChevron -Body $secDepsBody
})
$secUninstallHeader.Add_MouseLeftButtonDown({
    Toggle-Section -Chevron $secUninstallChevron -Body $secUninstallBody
})

# ── Stepper Click Handlers ────────────────────────────────────────────────────
$stepIdentity.Add_Click({
    Update-Stepper 1
    # Scroll to top of Create tab
    $pnlTabCreate.ScrollToTop()
})
$stepMetadata.Add_Click({
    Update-Stepper 2
    $secMetadataHeader.BringIntoView()
})
$stepInstaller.Add_Click({
    Update-Stepper 3
    $secInstallerHeader.BringIntoView()
})
$stepReview.Add_Click({
    Update-Stepper 4
    # Scroll to YAML Preview / Actions area
    $txtYamlPreview.BringIntoView()
})

# ── Inline Validation on Required Fields ──────────────────────────────────────
Add-FieldValidation -Field $txtPkgId      -FieldLabel 'Package Identifier'
Add-FieldValidation -Field $txtPkgVersion -FieldLabel 'Package Version'
Add-FieldValidation -Field $txtPkgName    -FieldLabel 'Package Name'
Add-FieldValidation -Field $txtPublisher  -FieldLabel 'Publisher'
Add-FieldValidation -Field $txtLicense    -FieldLabel 'License'
Add-FieldValidation -Field $txtShortDesc  -FieldLabel 'Short Description'

# URL format validation
Add-UrlValidation -Field $txtInstallerUrl0 -FieldLabel 'Installer URL'
Add-UrlValidation -Field $txtPackageUrl    -FieldLabel 'Package URL'
Add-UrlValidation -Field $txtPublisherUrl  -FieldLabel 'Publisher URL'
Add-UrlValidation -Field $txtLicenseUrl    -FieldLabel 'License URL'
Add-UrlValidation -Field $txtSupportUrl    -FieldLabel 'Support URL'
Add-UrlValidation -Field $txtPrivacyUrl    -FieldLabel 'Privacy URL'
Add-UrlValidation -Field $txtReleaseNotesUrl -FieldLabel 'Release Notes URL'

# ── Drag & Drop Zones ────────────────────────────────────────────────────────
Register-DropZone -Zone $dropZoneInstaller0 -OnFileDrop {
    param([string]$FilePath)
    $txtLocalFile0.Text = $FilePath
    Compute-FileHash -FilePath $FilePath -TargetField 'installer0'
    # Sync to upload tab
    $txtUploadFile.Text = $FilePath
    Compute-FileHash -FilePath $FilePath -TargetField 'upload'
    Show-Toast "File loaded: $(Split-Path $FilePath -Leaf)" -Type Success
}
Register-DropZone -Zone $dropZoneUpload -OnFileDrop {
    param([string]$FilePath)
    $txtUploadFile.Text = $FilePath
    Compute-FileHash -FilePath $FilePath -TargetField 'upload'
    # Sync to Create tab
    $txtLocalFile0.Text = $FilePath
    Compute-FileHash -FilePath $FilePath -TargetField 'installer0'
    Show-Toast "File loaded: $(Split-Path $FilePath -Leaf)" -Type Success
}

# ── URL Test Button ───────────────────────────────────────────────────────────
$btnTestUrl0.Add_Click({
    Test-InstallerUrl -Url $txtInstallerUrl0.Text.Trim()
})

# ── Hash from URL Button ─────────────────────────────────────────────────────
$btnHashUrl0 = $Window.FindName("btnHashUrl0")
$btnHashUrl0.Add_Click({
    Compute-UrlHash -Url $txtInstallerUrl0.Text.Trim()
})

# ── Cancel Download Button ────────────────────────────────────────────────────
$btnCancelDownload = $Window.FindName("btnCancelDownload")
$btnCancelDownload.Add_Click({
    $Global:DownloadCancelled = $true
    $Job = $Global:DownloadJob
    if ($Job) {
        try { $Job.PS.Stop() } catch { <# already stopped #> }
        try { $Job.PS.Dispose() } catch { <# already disposed #> }
        try { $Job.Runspace.Dispose() } catch { <# already disposed #> }
        $Global:BgJobs.Remove($Job)
        $Global:DownloadJob = $null
    }
    $btnHashUrl0.IsEnabled = $true
    $btnCancelDownload.Visibility = 'Collapsed'
    $Global:SyncHash.DownloadProgress = -1
    $lblHash0.Text = "Download cancelled"
    $lblStatus.Text = "Download cancelled"
    Write-DebugLog "Download cancelled by user" -Level 'WARN'
})

# ── Manage Tab Search/Filter ──────────────────────────────────────────────────
$txtManageSearch.Add_TextChanged({
    $SearchText = $txtManageSearch.Text.Trim().ToLower()
    if (-not $lstManagePackages.Items) { return }
    foreach ($Item in $lstManagePackages.Items) {
        if (-not $Item.Tag) { continue }
        $Match = (-not $SearchText) -or
                 ($Item.Tag.Id -and $Item.Tag.Id.ToLower().Contains($SearchText)) -or
                 ($Item.Content -and $Item.Content.ToString().ToLower().Contains($SearchText))
        if ($Item -is [System.Windows.Controls.ListViewItem]) {
            $Item.Visibility = if ($Match) { 'Visible' } else { 'Collapsed' }
        }
    }
})

# ── Sort Package List ─────────────────────────────────────────────────────────
$cmbManageSort = $Window.FindName("cmbManageSort")
$cmbManageSort.Add_SelectionChanged({
    if (-not $Global:ManagePackages -or $Global:ManagePackages.Count -eq 0) { return }
    $SortIdx = $cmbManageSort.SelectedIndex
    $Sorted = switch ($SortIdx) {
        0 { $Global:ManagePackages | Sort-Object { $_.Id } }
        1 { $Global:ManagePackages | Sort-Object { $_.Id } -Descending }
        2 { $Global:ManagePackages | Sort-Object { @($_.Versions).Count } -Descending }
        3 { $Global:ManagePackages | Sort-Object { @($_.Versions).Count } }
        default { $Global:ManagePackages }
    }
    $lstManagePackages.Items.Clear()
    foreach ($Pkg in $Sorted) {
        $Item = New-StyledListViewItem -IconChar ([char]0xE7C3) -PrimaryText $Pkg.Id `
            -SubtitleText "Versions: $(($Pkg.Versions -join ', '))" -TagValue $Pkg
        $lstManagePackages.Items.Add($Item) | Out-Null
    }
})

# ── New Version Button (Version Bump workflow) ────────────────────────────────
$btnNewVersion.Add_Click({
    $Sel = $lstManagePackages.SelectedItem
    if (-not $Sel -or -not $Sel.Tag) {
        Show-Toast "Select a package first" -Type Warning
        return
    }
    $PkgId = $Sel.Tag.Id
    $LatestVer = ($Sel.Tag.Versions | Sort-Object -Descending | Select-Object -First 1)

    # Pre-fill Create tab with existing package info
    $txtPkgId.Text = $PkgId
    if ($LatestVer) {
        # Try to bump version
        $Parts = $LatestVer -split '\.'
        if ($Parts.Count -gt 0 -and $Parts[-1] -match '^\d+$') {
            $Parts[-1] = [string]([int]$Parts[-1] + 1)
        }
        $txtPkgVersion.Text = $Parts -join '.'
    }

    Switch-Tab 'Create'
    Update-Stepper 1
    Show-Toast "Creating new version for $PkgId — update the version and fill in details" -Type Info -DurationMs 5000
})

# ── Import to Create Button ──────────────────────────────────────────────────
$btnImportToCreate = $Window.FindName("btnImportToCreate")
$btnImportToCreate.Add_Click({
    $Sel = $lstManagePackages.SelectedItem
    if (-not $Sel -or -not $Sel.Tag) {
        Show-Toast "Select a package first" -Type Warning
        return
    }
    $PkgId = $Sel.Tag.Id
    $SelVer = if ($cmbManageVersion.SelectedItem) { $cmbManageVersion.SelectedItem.ToString() } else { $null }
    if (-not $SelVer) {
        $SelVer = ($Sel.Tag.Versions | Sort-Object -Descending | Select-Object -First 1)
    }
    if (-not $SelVer) {
        Show-Toast "No version available to import" -Type Warning
        return
    }
    $FuncName = Get-ActiveRestSource
    if (-not $FuncName) {
        Show-Toast "Set REST Source Function Name in Settings first" -Type Warning
        return
    }

    Show-Toast "Fetching manifest for $PkgId v$SelVer..." -Type Info -DurationMs 3000

    Start-BackgroundWork -Variables @{
        FuncName = $FuncName
        ImportPkgId = $PkgId
        ImportVer = $SelVer
    } -Work {
        try {
            Import-Module Microsoft.WinGet.RestSource -ErrorAction Stop
            $M = Get-WinGetManifest -FunctionName $FuncName -PackageIdentifier $ImportPkgId -Version $ImportVer -ErrorAction Stop
            return @{ Manifest = ($M | ConvertTo-Json -Depth 10); Error = $null }
        } catch {
            return @{ Error = $_.Exception.Message }
        }
    } -OnComplete {
        param($Results, $Errors)
        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ Error = 'No result' } }
        if ($R.Error) {
            Show-Toast "Import failed: $($R.Error)" -Type Error
            return
        }
        try {
            $M = $R.Manifest | ConvertFrom-Json

            # Identity
            $txtPkgId.Text = if ($M.PackageIdentifier) { $M.PackageIdentifier } else { $ImportPkgId }
            $txtPkgVersion.Text = if ($M.PackageVersion) { $M.PackageVersion } else { $ImportVer }

            # Default locale metadata
            $dl = $M.DefaultLocale
            if (-not $dl -and $M.Locales) { $dl = $M.Locales | Where-Object { $_.PackageLocale -eq 'en-US' } | Select-Object -First 1 }
            if ($dl) {
                $txtPkgName.Text       = if ($dl.PackageName) { $dl.PackageName } else { "" }
                $txtPublisher.Text     = if ($dl.Publisher) { $dl.Publisher } else { "" }
                $txtLicense.Text       = if ($dl.License) { $dl.License } else { "" }
                $txtShortDesc.Text     = if ($dl.ShortDescription) { $dl.ShortDescription } else { "" }
                $txtDescription.Text   = if ($dl.Description) { $dl.Description } else { "" }
                $txtAuthor.Text        = if ($dl.Author) { $dl.Author } else { "" }
                $txtMoniker.Text       = if ($dl.Moniker) { $dl.Moniker } else { "" }
                $txtPackageUrl.Text    = if ($dl.PackageUrl) { $dl.PackageUrl } else { "" }
                $txtPublisherUrl.Text  = if ($dl.PublisherUrl) { $dl.PublisherUrl } else { "" }
                $txtLicenseUrl.Text    = if ($dl.LicenseUrl) { $dl.LicenseUrl } else { "" }
                $txtSupportUrl.Text    = if ($dl.PublisherSupportUrl) { $dl.PublisherSupportUrl } else { "" }
                $txtCopyright.Text     = if ($dl.Copyright) { $dl.Copyright } else { "" }
                $txtPrivacyUrl.Text    = if ($dl.PrivacyUrl) { $dl.PrivacyUrl } else { "" }
                $txtReleaseNotes.Text  = if ($dl.ReleaseNotes) { $dl.ReleaseNotes } else { "" }
                $txtReleaseNotesUrl.Text = if ($dl.ReleaseNotesUrl) { $dl.ReleaseNotesUrl } else { "" }
                if ($dl.Tags -and $dl.Tags.Count -gt 0) { $txtTags.Text = ($dl.Tags -join ', ') } else { $txtTags.Text = "" }
            }

            # Installers
            $Global:InstallerEntries.Clear()
            if ($M.Installers -and $M.Installers.Count -gt 0) {
                $first = $M.Installers[0]
                if ($first.Architecture)    { Set-ComboBoxByContent $cmbArch0 $first.Architecture }
                if ($first.InstallerType)   { Set-ComboBoxByContent $cmbInstallerType0 $first.InstallerType }
                if ($first.Scope)           { Set-ComboBoxByContent $cmbScope0 $first.Scope }
                $txtInstallerUrl0.Text = if ($first.InstallerUrl) { $first.InstallerUrl } else { "" }
                if ($first.InstallerSha256) {
                    $Global:CurrentHash0 = $first.InstallerSha256
                    $lblHash0.Text = "SHA256: $($first.InstallerSha256)"
                }
                if ($first.UpgradeBehavior)      { Set-ComboBoxByContent $cmbUpgrade0 $first.UpgradeBehavior }
                if ($first.ElevationRequirement) { Set-ComboBoxByContent $cmbElevation0 $first.ElevationRequirement }
                $txtProductCode0.Text = if ($first.ProductCode) { $first.ProductCode } else { "" }
                $txtCommands0.Text = if ($first.Commands) { ($first.Commands -join ', ') } else { "" }
                $txtFileExt0.Text = if ($first.FileExtensions) { ($first.FileExtensions -join ', ') } else { "" }

                # Switches
                if ($first.InstallerSwitches) {
                    $sw = $first.InstallerSwitches
                    $txtSilent0.Text         = if ($sw.Silent) { $sw.Silent } else { "" }
                    $txtSilentProgress0.Text = if ($sw.SilentWithProgress) { $sw.SilentWithProgress } else { "" }
                    $txtCustomSwitch0.Text   = if ($sw.Custom) { $sw.Custom } else { "" }
                    $txtInteractive0.Text    = if ($sw.Interactive) { $sw.Interactive } else { "" }
                    $txtLog0.Text            = if ($sw.Log) { $sw.Log } else { "" }
                    $txtRepair0.Text         = if ($sw.Repair) { $sw.Repair } else { "" }
                }

                # Install modes
                $modes = @($first.InstallModes)
                $chkModeInteractive.IsChecked    = 'interactive' -in $modes
                $chkModeSilent.IsChecked          = 'silent' -in $modes
                $chkModeSilentProgress.IsChecked  = 'silentWithProgress' -in $modes
                $chkInstallLocationRequired.IsChecked = [bool]$first.InstallLocationRequired

                # Platform & MinOS
                $plat = @($first.Platform)
                $chkPlatformDesktop.IsChecked  = 'Windows.Desktop' -in $plat
                $chkPlatformUniversal.IsChecked = 'Windows.Universal' -in $plat
                $txtMinOSVersion0.Text = if ($first.MinimumOSVersion) { $first.MinimumOSVersion } else { "" }

                # Additional installers → InstallerEntries
                for ($idx = 1; $idx -lt $M.Installers.Count; $idx++) {
                    $ei = $M.Installers[$idx]
                    $entry = @{
                        Architecture       = $ei.Architecture
                        InstallerType      = $ei.InstallerType
                        InstallerUrl       = $ei.InstallerUrl
                        InstallerSha256    = $ei.InstallerSha256
                        Scope              = $ei.Scope
                        UpgradeBehavior    = $ei.UpgradeBehavior
                        ElevationRequirement = $ei.ElevationRequirement
                        ProductCode        = $ei.ProductCode
                        Commands           = @($ei.Commands)
                        FileExtensions     = @($ei.FileExtensions)
                        InstallModes       = @($ei.InstallModes)
                        InstallerSwitches  = @{
                            Silent             = $ei.InstallerSwitches.Silent
                            SilentWithProgress = $ei.InstallerSwitches.SilentWithProgress
                            Custom             = $ei.InstallerSwitches.Custom
                            Interactive        = $ei.InstallerSwitches.Interactive
                            Log                = $ei.InstallerSwitches.Log
                            Repair             = $ei.InstallerSwitches.Repair
                        }
                        InstallLocationRequired = [bool]$ei.InstallLocationRequired
                        MinimumOSVersion   = $ei.MinimumOSVersion
                        Platform           = @($ei.Platform)
                    }
                    $Global:InstallerEntries.Add($entry)
                }
                Update-InstallerSummaryList
            }

            Switch-Tab 'Create'
            Update-Stepper 1
            Update-YamlPreview
            Show-Toast "Imported $($M.PackageIdentifier) v$($M.PackageVersion) — edit and publish" -Type Success -DurationMs 5000
        } catch {
            Show-Toast "Error parsing manifest: $($_.Exception.Message)" -Type Error
        }
    }
})

# ── Import from WinGet Community Repo ─────────────────────────────────────────
$txtCommunityPkgId   = $Window.FindName("txtCommunityPkgId")
$txtCommunityVersion = $Window.FindName("txtCommunityVersion")
$btnImportCommunity  = $Window.FindName("btnImportCommunity")

$btnImportCommunity.Add_Click({
    $PkgId = $txtCommunityPkgId.Text.Trim()
    if (-not $PkgId -or $PkgId -eq 'e.g. Microsoft.VisualStudioCode') {
        Show-Toast "Enter a package identifier" -Type Warning
        return
    }
    $ReqVersion = $txtCommunityVersion.Text.Trim()
    if ($ReqVersion -eq 'latest' -or $ReqVersion -eq 'Version (latest)') { $ReqVersion = '' }

    # Check if a config already exists for this package
    $ExistingConfig = Join-Path $Global:ConfigDir "$PkgId.json"
    if (Test-Path $ExistingConfig) {
        $Confirm = Show-ThemedDialog -Title 'Existing Config Found' `
            -Message "A local config for '$PkgId' already exists.`n`nImporting from the community repo will overwrite the current form data and auto-save the config." `
            -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' `
            -Buttons @(
                @{ Text='Cancel'; IsAccent=$false; Result='Cancel' }
                @{ Text='Import & Overwrite'; IsAccent=$true; Result='Import' }
            )
        if ($Confirm -ne 'Import') { return }
    }

    Show-Toast "Fetching from winget-pkgs repo..." -Type Info -DurationMs 5000

    # Build path: manifests/<first-char-lower>/<Publisher>/<Name>/...
    $Parts = $PkgId.Split('.')
    $FirstChar = $Parts[0].Substring(0,1).ToLower()
    $PathPrefix = "manifests/$FirstChar/$($Parts -join '/')"

    Start-BackgroundWork -Variables @{
        PathPrefix = $PathPrefix
        ReqVersion = $ReqVersion
        PkgId = $PkgId
    } -Work {
        try {
            $BaseUrl = "https://api.github.com/repos/microsoft/winget-pkgs/contents/$PathPrefix"
            $Headers = @{ 'User-Agent' = 'WinGetManifestManager'; Accept = 'application/vnd.github.v3+json' }

            # If no version specified, list directory to find latest
            $Version = $ReqVersion
            if (-not $Version) {
                try {
                    $Listing = Invoke-RestMethod -Uri $BaseUrl -Headers $Headers -ErrorAction Stop
                } catch {
                    if ($_.Exception.Message -match '404|Not Found') {
                        return @{ Error = "Package '$PkgId' not found in the winget-pkgs community repo.`nVerify the Package Identifier is correct (case-sensitive, e.g. Microsoft.VisualStudioCode)." }
                    }
                    throw
                }
                $Dirs = @($Listing | Where-Object { $_.type -eq 'dir' } | Select-Object -ExpandProperty name)
                # Filter to version-like directories (must start with a digit) — excludes "Beta", ".validation", etc.
                $Dirs = @($Dirs | Where-Object { $_ -match '^\d' })
                # If no version dirs found, check if we're one level too high (user entered short name like "WinMerge" instead of "WinMerge.WinMerge")
                if ($Dirs.Count -eq 0) {
                    $AllDirs = @($Listing | Where-Object { $_.type -eq 'dir' } | Select-Object -ExpandProperty name)
                    $NonSpecial = @($AllDirs | Where-Object { $_ -notmatch '^\.' })
                    if ($NonSpecial.Count -eq 1) {
                        # Single sub-package directory — drill down automatically
                        $BaseUrl = "$BaseUrl/$($NonSpecial[0])"
                        $PkgId = "$PkgId.$($NonSpecial[0])"
                        $SubListing = Invoke-RestMethod -Uri $BaseUrl -Headers $Headers -ErrorAction Stop
                        $Dirs = @($SubListing | Where-Object { $_.type -eq 'dir' } | Select-Object -ExpandProperty name)
                        $Dirs = @($Dirs | Where-Object { $_ -match '^\d' })
                    } elseif ($NonSpecial.Count -gt 1) {
                        $Candidates = ($NonSpecial | ForEach-Object { "$PkgId.$_" }) -join ", "
                        return @{ Error = "Package '$PkgId' is a publisher, not a package.`nDid you mean one of: $Candidates" }
                    }
                }
                if ($Dirs.Count -eq 0) { return @{ Error = "Package '$PkgId' exists but has no version directories." } }
                # Semver-aware sort: split on dots, pad each segment to 10 chars for proper numeric ordering
                $Version = ($Dirs | Sort-Object {
                    ($_ -split '[.\-]' | ForEach-Object { $_.PadLeft(10, '0') }) -join '.'
                } -Descending | Select-Object -First 1)
            }

            $VerUrl = "$BaseUrl/$Version"
            try {
                $Files = Invoke-RestMethod -Uri $VerUrl -Headers $Headers -ErrorAction Stop
            } catch {
                if ($_.Exception.Message -match '404|Not Found') {
                    return @{ Error = "Version '$Version' not found for '$PkgId'.`nAvailable versions may differ — try leaving version blank to auto-detect latest." }
                }
                throw
            }

            # If directory contains subdirs instead of YAML files (e.g. Beta/2.16.44-beta3/),
            # recurse into the latest sub-version
            $YamlFiles = @($Files | Where-Object { $_.name -match '\.yaml$' })
            if ($YamlFiles.Count -eq 0) {
                $SubDirs = @($Files | Where-Object { $_.type -eq 'dir' -and $_.name -match '^\d' } |
                    Select-Object -ExpandProperty name)
                if ($SubDirs.Count -gt 0) {
                    $SubVer = ($SubDirs | Sort-Object {
                        ($_ -split '[.\-]' | ForEach-Object { $_.PadLeft(10, '0') }) -join '.'
                    } -Descending | Select-Object -First 1)
                    $Version = "$Version/$SubVer"
                    $VerUrl = "$BaseUrl/$Version"
                    try {
                        $Files = Invoke-RestMethod -Uri $VerUrl -Headers $Headers -ErrorAction Stop
                    } catch { throw }
                    $YamlFiles = @($Files | Where-Object { $_.name -match '\.yaml$' })
                }
            }

            $Result = @{ Version = ($Version -split '/')[-1]; Yamls = @{} }

            foreach ($F in $YamlFiles) {
                $Content = Invoke-RestMethod -Uri $F.download_url -Headers $Headers -ErrorAction Stop
                $Result.Yamls[$F.name] = $Content
            }
            if ($Result.Yamls.Count -eq 0) { return @{ Error = "No YAML manifest files found for '$PkgId' v$Version" } }
            $Result.ResolvedPkgId = $PkgId
            return @{ Data = $Result; Error = $null }
        } catch {
            return @{ Error = $_.Exception.Message }
        }
    } -OnComplete {
        param($Results, $Errors)
        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ Error = 'No result' } }
        if ($R.Error) {
            Show-Toast "Community import failed: $($R.Error)" -Type Error
            return
        }
        try {
            $Data = $R.Data
            $Yamls = $Data.Yamls
            $ResolvedPkgId = if ($Data.ResolvedPkgId) { $Data.ResolvedPkgId } else { $PkgId }

            # If the package ID was resolved to a different name (e.g. WinMerge → WinMerge.WinMerge),
            # the pre-import check may have tested the wrong filename — verify the resolved config too
            if ($ResolvedPkgId -ne $PkgId) {
                $ResolvedConfig = Join-Path $Global:ConfigDir "$ResolvedPkgId.json"
                if (Test-Path $ResolvedConfig) {
                    $Confirm = Show-ThemedDialog -Title 'Existing Config Found' `
                        -Message "Package resolved to '$ResolvedPkgId'.`nA local config for this package already exists.`n`nImporting will overwrite the current form data and auto-save the config." `
                        -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' `
                        -Buttons @(
                            @{ Text='Cancel'; IsAccent=$false; Result='Cancel' }
                            @{ Text='Import & Overwrite'; IsAccent=$true; Result='Import' }
                        )
                    if ($Confirm -ne 'Import') {
                        Show-Toast "Import cancelled" -Type Info
                        return
                    }
                }
            }

            # ── Reset all form fields before populating ──────────────────────
            $txtPkgId.Text = ""; $txtPkgVersion.Text = ""; $txtPkgName.Text = ""
            $txtPublisher.Text = ""; $txtLicense.Text = ""; $txtShortDesc.Text = ""
            $txtDescription.Text = ""; $txtAuthor.Text = ""; $txtMoniker.Text = ""
            $txtPackageUrl.Text = ""; $txtPublisherUrl.Text = ""; $txtLicenseUrl.Text = ""
            $txtSupportUrl.Text = ""; $txtCopyright.Text = ""; $txtPrivacyUrl.Text = ""
            $txtTags.Text = ""; $txtReleaseNotes.Text = ""; $txtReleaseNotesUrl.Text = ""
            $txtLocalFile0.Text = ""; $lblHash0.Text = ""; $txtInstallerUrl0.Text = ""
            $txtSilent0.Text = ""; $txtSilentProgress0.Text = ""
            $txtProductCode0.Text = ""; $txtCommands0.Text = ""; $txtFileExt0.Text = ""
            $txtCustomSwitch0.Text = ""; $txtInteractive0.Text = ""
            $txtLog0.Text = ""; $txtRepair0.Text = ""
            $txtNestedPath0.Text = ""; $txtPortableAlias0.Text = ""
            $txtMinOSVersion0.Text = ""; $txtExpectedReturnCodes0.Text = ""
            $txtAFDisplayName0.Text = ""; $txtAFPublisher0.Text = ""
            $txtAFDisplayVersion0.Text = ""; $txtAFProductCode0.Text = ""
            $txtPackageDeps.Text = ""; $txtWindowsFeatures.Text = ""
            $txtUninstallCmd.Text = ""; $txtUninstallSilent.Text = ""
            $cmbInstallerType0.SelectedIndex = 0; $cmbArch0.SelectedIndex = 0
            $cmbScope0.SelectedIndex = 0; $cmbUpgrade0.SelectedIndex = 0
            $cmbElevation0.SelectedIndex = 0; $cmbNestedType0.SelectedIndex = 0
            $cmbInstallerPreset.SelectedIndex = 0
            $pnlNestedInstaller.Visibility = 'Collapsed'
            $chkModeInteractive.IsChecked = $true; $chkModeSilent.IsChecked = $true
            $chkModeSilentProgress.IsChecked = $false
            $chkInstallLocationRequired.IsChecked = $false
            $chkPlatformDesktop.IsChecked = $true; $chkPlatformUniversal.IsChecked = $false
            $Global:CurrentHash0 = ""
            $Global:InstallerEntries.Clear()
            $txtYamlPreview.Text = ""

            # Parse YAML files — community uses split format: .installer.yaml, .locale.*.yaml, main .yaml
            # Handles simple key: value, YAML lists (Tags), and block scalars (ReleaseNotes)
            $ParseYaml = {
                param([string]$Text)
                $Map = @{}
                $ListKey = $null; $ListItems = @()
                $BlockKey = $null; $BlockLines = @()
                foreach ($Line in ($Text -split "`n")) {
                    $L = $Line.TrimEnd()
                    # Flush block scalar on non-indented line
                    if ($BlockKey -and $L -match '^\S') {
                        $Map[$BlockKey] = ($BlockLines -join "`n").Trim()
                        $BlockKey = $null; $BlockLines = @()
                    }
                    if ($BlockKey) {
                        if ($L -match '^\s{2,}(.*)$') { $BlockLines += $Matches[1] }
                        continue
                    }
                    # Flush list on non-list, non-blank line
                    if ($ListKey -and $L -match '^\S') {
                        $Map[$ListKey] = $ListItems -join ', '
                        $ListKey = $null; $ListItems = @()
                    }
                    if ($ListKey -and $L -match '^\s*-\s+(.+)$') {
                        $ListItems += $Matches[1].Trim()
                        continue
                    }
                    # Key with block scalar indicator (|, |-, >, >-)
                    if ($L -match '^(\w[\w\.]*)\s*:\s*[|>]-?\s*$') {
                        $BlockKey = $Matches[1]; $BlockLines = @()
                        continue
                    }
                    # Key with no value — start of a YAML list (e.g. Tags:)
                    if ($L -match '^(\w[\w\.]*)\s*:\s*$') {
                        $ListKey = $Matches[1]; $ListItems = @()
                        continue
                    }
                    # Simple key: value
                    if ($L -match '^\s{0,2}(\w[\w\.]*)\s*:\s*(.+)$') {
                        $Map[$Matches[1]] = $Matches[2].Trim().Trim('"').Trim("'")
                    }
                }
                if ($ListKey -and $ListItems.Count -gt 0) { $Map[$ListKey] = $ListItems -join ', ' }
                if ($BlockKey -and $BlockLines.Count -gt 0) { $Map[$BlockKey] = ($BlockLines -join "`n").Trim() }
                return $Map
            }

            # Find the main version manifest
            $MainFile = $Yamls.Keys | Where-Object { $_ -notmatch '\.(installer|locale)\.' } | Select-Object -First 1
            if ($MainFile) {
                $Main = & $ParseYaml $Yamls[$MainFile]
                if ($Main.PackageIdentifier) { $txtPkgId.Text = $Main.PackageIdentifier }
                if ($Main.PackageVersion) { $txtPkgVersion.Text = $Main.PackageVersion }
                if ($Main.PackageName) { $txtPkgName.Text = $Main.PackageName }
                if ($Main.Publisher)   { $txtPublisher.Text = $Main.Publisher }
                if ($Main.License)     { $txtLicense.Text = $Main.License }
                if ($Main.ShortDescription) { $txtShortDesc.Text = $Main.ShortDescription }
                if ($Main.Description) { $txtDescription.Text = $Main.Description }
                if ($Main.Author)      { $txtAuthor.Text = $Main.Author }
                if ($Main.Moniker)     { $txtMoniker.Text = $Main.Moniker }
                if ($Main.PackageUrl)  { $txtPackageUrl.Text = $Main.PackageUrl }
                if ($Main.PublisherUrl) { $txtPublisherUrl.Text = $Main.PublisherUrl }
                if ($Main.LicenseUrl) { $txtLicenseUrl.Text = $Main.LicenseUrl }
                if ($Main.PublisherSupportUrl) { $txtSupportUrl.Text = $Main.PublisherSupportUrl }
                if ($Main.Copyright)   { $txtCopyright.Text = $Main.Copyright }
                if ($Main.PrivacyUrl)  { $txtPrivacyUrl.Text = $Main.PrivacyUrl }
                if ($Main.ReleaseNotes) { $txtReleaseNotes.Text = $Main.ReleaseNotes }
                if ($Main.ReleaseNotesUrl) { $txtReleaseNotesUrl.Text = $Main.ReleaseNotesUrl }
                if ($Main.Tags) { $txtTags.Text = $Main.Tags }
            }

            # Locale file — prefer en-US / defaultLocale over other locales
            $LocaleFile = $Yamls.Keys | Where-Object { $_ -match '\.locale\.en-US\.' } | Select-Object -First 1
            if (-not $LocaleFile) {
                $LocaleFile = $Yamls.Keys | Where-Object { $_ -match '\.locale\.' } | Sort-Object | Select-Object -First 1
            }
            if ($LocaleFile) {
                $Loc = & $ParseYaml $Yamls[$LocaleFile]
                if ($Loc.PackageName -and -not $txtPkgName.Text) { $txtPkgName.Text = $Loc.PackageName }
                if ($Loc.Publisher -and -not $txtPublisher.Text) { $txtPublisher.Text = $Loc.Publisher }
                if ($Loc.License -and -not $txtLicense.Text) { $txtLicense.Text = $Loc.License }
                if ($Loc.ShortDescription -and -not $txtShortDesc.Text) { $txtShortDesc.Text = $Loc.ShortDescription }
                if ($Loc.Description -and -not $txtDescription.Text) { $txtDescription.Text = $Loc.Description }
                if ($Loc.Author -and -not $txtAuthor.Text) { $txtAuthor.Text = $Loc.Author }
                if ($Loc.Moniker -and -not $txtMoniker.Text) { $txtMoniker.Text = $Loc.Moniker }
                if ($Loc.PackageUrl -and -not $txtPackageUrl.Text) { $txtPackageUrl.Text = $Loc.PackageUrl }
                if ($Loc.PublisherUrl -and -not $txtPublisherUrl.Text) { $txtPublisherUrl.Text = $Loc.PublisherUrl }
                if ($Loc.LicenseUrl -and -not $txtLicenseUrl.Text) { $txtLicenseUrl.Text = $Loc.LicenseUrl }
                if ($Loc.PublisherSupportUrl -and -not $txtSupportUrl.Text) { $txtSupportUrl.Text = $Loc.PublisherSupportUrl }
                if ($Loc.Copyright -and -not $txtCopyright.Text) { $txtCopyright.Text = $Loc.Copyright }
                if ($Loc.PrivacyUrl -and -not $txtPrivacyUrl.Text) { $txtPrivacyUrl.Text = $Loc.PrivacyUrl }
                if ($Loc.ReleaseNotes -and -not $txtReleaseNotes.Text) { $txtReleaseNotes.Text = $Loc.ReleaseNotes }
                if ($Loc.ReleaseNotesUrl -and -not $txtReleaseNotesUrl.Text) { $txtReleaseNotesUrl.Text = $Loc.ReleaseNotesUrl }
                if ($Loc.Tags -and -not $txtTags.Text) { $txtTags.Text = $Loc.Tags }
            }

            # Installer file — parse installers section (handles top-level defaults + 2-space indent)
            $InstallerFile = $Yamls.Keys | Where-Object { $_ -match '\.installer\.' } | Select-Object -First 1
            if ($InstallerFile) {
                $InsYaml = $Yamls[$InstallerFile]
                $Global:InstallerEntries.Clear()

                $Installers = @()
                $Current = $null
                $InSwitchBlock = $false
                $InInstallersList = $false
                $InSubBlock = $null
                $TopDefaults = @{}
                $TopInstallModes = @()
                $TopCommands = @()
                $TopFileExt = @()
                $TopPlatform = @()
                $TopDepPkgs = @()
                $TopDepFeatures = @()

                foreach ($Line in ($InsYaml -split "`n")) {
                    $L = $Line.TrimEnd()
                    if ($L -match '^\s*#' -or -not $L.Trim()) { continue }

                    # Top-level key (0 indent) — shared defaults like InstallerType, Scope
                    if ($L -match '^(\w[\w\.]*)\s*:\s*(.*)$') {
                        $K = $Matches[1]; $V = $Matches[2].Trim().Trim('"').Trim("'")
                        if ($K -eq 'Installers') { $InInstallersList = $true; $InSwitchBlock = $false; $InSubBlock = $null; continue }
                        if ($K -eq 'InstallerSwitches') { $InSwitchBlock = $true; $InSubBlock = $null; continue }
                        if ($K -eq 'InstallModes')   { $InSubBlock = 'TopModes'; $InSwitchBlock = $false; continue }
                        if ($K -eq 'Commands')        { $InSubBlock = 'TopCmds'; $InSwitchBlock = $false; continue }
                        if ($K -eq 'FileExtensions')  { $InSubBlock = 'TopFE'; $InSwitchBlock = $false; continue }
                        if ($K -eq 'Platform')        { $InSubBlock = 'TopPlat'; $InSwitchBlock = $false; continue }
                        if ($K -eq 'Dependencies')    { $InSubBlock = 'TopDeps'; $InSwitchBlock = $false; continue }
                        $InSwitchBlock = $false; $InSubBlock = $null
                        $TopDefaults[$K] = $V
                        continue
                    }

                    # Top-level list/sub-block items (before Installers block)
                    if (-not $InInstallersList -and $InSubBlock) {
                        if ($InSubBlock -eq 'TopModes' -and $L -match '^\s*-\s+(.+)$')  { $TopInstallModes += $Matches[1].Trim(); continue }
                        if ($InSubBlock -eq 'TopCmds' -and $L -match '^\s*-\s+(.+)$')   { $TopCommands += $Matches[1].Trim(); continue }
                        if ($InSubBlock -eq 'TopFE' -and $L -match '^\s*-\s+(.+)$')     { $TopFileExt += $Matches[1].Trim(); continue }
                        if ($InSubBlock -eq 'TopPlat' -and $L -match '^\s*-\s+(.+)$')   { $TopPlatform += $Matches[1].Trim(); continue }
                        if ($InSubBlock -eq 'TopDeps') {
                            if ($L -match '^\s+PackageDependencies\s*:') { $InSubBlock = 'TopDepPkgs'; continue }
                            if ($L -match '^\s+WindowsFeatures\s*:')    { $InSubBlock = 'TopDepFeats'; continue }
                        }
                        if ($InSubBlock -eq 'TopDepPkgs' -and $L -match '^\s*-\s+PackageIdentifier\s*:\s*(.+)$') { $TopDepPkgs += $Matches[1].Trim(); continue }
                        if ($InSubBlock -eq 'TopDepFeats' -and $L -match '^\s*-\s+(.+)$') { $TopDepFeatures += $Matches[1].Trim(); continue }
                    }

                    # Top-level InstallerSwitches children (2-space indent, before Installers block)
                    if ($InSwitchBlock -and -not $InInstallersList -and $L -match '^\s{2}(\w+)\s*:\s*(.+)$') {
                        $TopDefaults["Switch_$($Matches[1])"] = $Matches[2].Trim().Trim('"').Trim("'")
                        continue
                    }

                    if (-not $InInstallersList) { continue }

                    # New list item: "- Key: Value"
                    if ($L -match '^\s*-\s+(\w[\w\.]*)\s*:\s*(.*)$') {
                        $InSwitchBlock = $false
                        if ($Current) { $Installers += $Current }
                        $Current = @{}
                        $Current[$Matches[1]] = $Matches[2].Trim().Trim('"').Trim("'")
                        continue
                    }

                    # Properties within an installer entry
                    if ($Current) {
                        if ($L -match '^\s{2,3}InstallerSwitches\s*:') { $InSwitchBlock = $true; $InSubBlock = 'Switches'; continue }
                        if ($L -match '^\s{2,3}AppsAndFeaturesEntries\s*:') { $InSwitchBlock = $false; $InSubBlock = 'AandF'; continue }
                        if ($L -match '^\s{2,3}ExpectedReturnCodes\s*:') { $InSwitchBlock = $false; $InSubBlock = 'ERC'; continue }
                        if ($L -match '^\s{2,3}NestedInstallerFiles\s*:') { $InSwitchBlock = $false; $InSubBlock = 'NIF'; continue }
                        if ($L -match '^\s{2,3}Dependencies\s*:') { $InSwitchBlock = $false; $InSubBlock = 'Deps'; continue }
                        if ($L -match '^\s{2,3}InstallModes\s*:') { $InSwitchBlock = $false; $InSubBlock = 'Modes'; $Current['_InstallModes'] = @(); continue }
                        if ($L -match '^\s{2,3}Commands\s*:') { $InSwitchBlock = $false; $InSubBlock = 'CmdList'; $Current['_Commands'] = @(); continue }
                        if ($L -match '^\s{2,3}FileExtensions\s*:') { $InSwitchBlock = $false; $InSubBlock = 'FEList'; $Current['_FileExtensions'] = @(); continue }
                        if ($L -match '^\s{2,3}Platform\s*:') { $InSwitchBlock = $false; $InSubBlock = 'PlatList'; $Current['_Platform'] = @(); continue }
                        if ($L -match '^\s{2,3}Protocols\s*:') { $InSwitchBlock = $false; $InSubBlock = 'ProtoList'; $Current['_Protocols'] = @(); continue }

                        # Collect list items for yaml lists
                        if ($InSubBlock -eq 'Modes' -and $L -match '^\s*-\s+(.+)$')    { $Current['_InstallModes'] += $Matches[1].Trim(); continue }
                        if ($InSubBlock -eq 'CmdList' -and $L -match '^\s*-\s+(.+)$')  { $Current['_Commands'] += $Matches[1].Trim(); continue }
                        if ($InSubBlock -eq 'FEList' -and $L -match '^\s*-\s+(.+)$')   { $Current['_FileExtensions'] += $Matches[1].Trim(); continue }
                        if ($InSubBlock -eq 'PlatList' -and $L -match '^\s*-\s+(.+)$') { $Current['_Platform'] += $Matches[1].Trim(); continue }
                        if ($InSubBlock -eq 'ProtoList' -and $L -match '^\s*-\s+(.+)$') { $Current['_Protocols'] += $Matches[1].Trim(); continue }

                        # InstallerSwitches children
                        if ($InSwitchBlock -and $L -match '^\s{4,}(\w+)\s*:\s*(.+)$') {
                            $Current["Switch_$($Matches[1])"] = $Matches[2].Trim().Trim('"').Trim("'")
                            continue
                        }
                        # AppsAndFeaturesEntries (first entry only — take list item or sub-keys)
                        if ($InSubBlock -eq 'AandF') {
                            if ($L -match '^\s*-\s+(\w+)\s*:\s*(.+)$') { $Current["AF_$($Matches[1])"] = $Matches[2].Trim().Trim('"').Trim("'"); continue }
                            if ($L -match '^\s{4,}(\w+)\s*:\s*(.+)$')  { $Current["AF_$($Matches[1])"] = $Matches[2].Trim().Trim('"').Trim("'"); continue }
                        }
                        # ExpectedReturnCodes (collect code:response pairs)
                        if ($InSubBlock -eq 'ERC') {
                            if ($L -match '^\s*-\s+InstallerReturnCode\s*:\s*(\d+)')   { $Current['_ERC_Code'] = $Matches[1]; continue }
                            if ($L -match '^\s{4,}ReturnResponse\s*:\s*(.+)$')         { $Current["_ERC_$($Current['_ERC_Code'])"] = $Matches[1].Trim(); continue }
                            if ($L -match '^\s*-\s+ReturnResponse\s*:\s*(.+)$')        { continue }  # handled above
                        }
                        # NestedInstallerFiles
                        if ($InSubBlock -eq 'NIF') {
                            if ($L -match '^\s*-\s+RelativeFilePath\s*:\s*(.+)$')  { $Current['NestedRelPath'] = $Matches[1].Trim().Trim('"').Trim("'"); continue }
                            if ($L -match '^\s{4,}RelativeFilePath\s*:\s*(.+)$')   { $Current['NestedRelPath'] = $Matches[1].Trim().Trim('"').Trim("'"); continue }
                            if ($L -match '^\s{4,}PortableCommandAlias\s*:\s*(.+)$') { $Current['NestedAlias'] = $Matches[1].Trim().Trim('"').Trim("'"); continue }
                        }
                        # Dependencies sub-keys
                        if ($InSubBlock -eq 'Deps') {
                            if ($L -match '^\s{4,}(\w+)\s*:\s*(.+)$')  { $Current["Dep_$($Matches[1])"] = $Matches[2].Trim().Trim('"').Trim("'"); continue }
                            if ($L -match '^\s*-\s+(.+)$')             { continue }  # dependency list items
                        }

                        if ($L -match '^\s{2,3}(\w[\w\.]*)\s*:\s*(.+)$') {
                            $InSwitchBlock = $false; $InSubBlock = $null
                            $Current[$Matches[1]] = $Matches[2].Trim().Trim('"').Trim("'")
                        }
                    }
                }
                if ($Current) { $Installers += $Current }

                # Merge top-level defaults into each installer entry
                foreach ($inst in $Installers) {
                    foreach ($dk in $TopDefaults.Keys) {
                        if (-not $inst[$dk]) { $inst[$dk] = $TopDefaults[$dk] }
                    }
                }

                if ($Installers.Count -gt 0) {
                    $first = $Installers[0]
                    if ($first.Architecture)  { Set-ComboBoxByContent $cmbArch0 $first.Architecture }
                    if ($first.InstallerType) { Set-ComboBoxByContent $cmbInstallerType0 $first.InstallerType }
                    if ($first.Scope)          { Set-ComboBoxByContent $cmbScope0 $first.Scope }
                    $txtInstallerUrl0.Text = if ($first.InstallerUrl) { $first.InstallerUrl } else { '' }
                    if ($first.InstallerSha256) {
                        $Global:CurrentHash0 = $first.InstallerSha256
                        $lblHash0.Text = "SHA256: $($first.InstallerSha256)"
                    }
                    if ($first.UpgradeBehavior) { Set-ComboBoxByContent $cmbUpgrade0 $first.UpgradeBehavior }
                    if ($first.ProductCode) { $txtProductCode0.Text = $first.ProductCode }
                    if ($first.MinimumOSVersion) { $txtMinOSVersion0.Text = $first.MinimumOSVersion }
                    if ($first.ElevationRequirement) { Set-ComboBoxByContent $cmbElevation0 $first.ElevationRequirement }
                    # Commands, FileExtensions
                    if ($first['_Commands'] -and $first['_Commands'].Count -gt 0) { $txtCommands0.Text = $first['_Commands'] -join ', ' }
                    if ($first['_FileExtensions'] -and $first['_FileExtensions'].Count -gt 0) { $txtFileExt0.Text = $first['_FileExtensions'] -join ', ' }
                    # InstallModes
                    if ($first['_InstallModes']) {
                        $chkModeInteractive.IsChecked    = ('interactive' -in $first['_InstallModes'])
                        $chkModeSilent.IsChecked          = ('silent' -in $first['_InstallModes'])
                        $chkModeSilentProgress.IsChecked  = ('silentWithProgress' -in $first['_InstallModes'])
                    }
                    # Platform
                    if ($first['_Platform']) {
                        $chkPlatformDesktop.IsChecked  = ('Windows.Desktop' -in $first['_Platform'])
                        $chkPlatformUniversal.IsChecked = ('Windows.Universal' -in $first['_Platform'])
                    }
                    # ExpectedReturnCodes (format: "code:response, code:response")
                    $ercParts = @()
                    foreach ($ek in ($first.Keys | Where-Object { $_ -match '^_ERC_\d+$' })) {
                        $code = $ek -replace '^_ERC_', ''
                        $ercParts += "$($code):$($first[$ek])"
                    }
                    if ($ercParts.Count -gt 0) { $txtExpectedReturnCodes0.Text = $ercParts -join ', ' }
                    # AppsAndFeaturesEntries
                    if ($first.AF_DisplayName)    { $txtAFDisplayName0.Text    = $first.AF_DisplayName }
                    if ($first.AF_Publisher)      { $txtAFPublisher0.Text      = $first.AF_Publisher }
                    if ($first.AF_DisplayVersion) { $txtAFDisplayVersion0.Text = $first.AF_DisplayVersion }
                    if ($first.AF_ProductCode)    { $txtAFProductCode0.Text    = $first.AF_ProductCode }
                    # NestedInstaller
                    if ($first.NestedInstallerType) { Set-ComboBoxByContent $cmbNestedType0 $first.NestedInstallerType }
                    if ($first.NestedRelPath) { $txtNestedPath0.Text = $first.NestedRelPath }
                    if ($first.NestedAlias) { $txtPortableAlias0.Text = $first.NestedAlias }
                    # Switches
                    if ($first.Switch_Silent)             { $txtSilent0.Text = $first.Switch_Silent }
                    if ($first.Switch_SilentWithProgress) { $txtSilentProgress0.Text = $first.Switch_SilentWithProgress }
                    if ($first.Switch_Custom)             { $txtCustomSwitch0.Text = $first.Switch_Custom }
                    if ($first.Switch_Interactive)        { $txtInteractive0.Text = $first.Switch_Interactive }
                    if ($first.Switch_Log)                { $txtLog0.Text = $first.Switch_Log }
                    if ($first.Switch_Repair)             { $txtRepair0.Text = $first.Switch_Repair }

                    # Top-level shared lists (apply if per-installer fields aren't set)
                    if ($TopInstallModes.Count -gt 0 -and -not $first['_InstallModes']) {
                        $chkModeInteractive.IsChecked    = ('interactive' -in $TopInstallModes)
                        $chkModeSilent.IsChecked          = ('silent' -in $TopInstallModes)
                        $chkModeSilentProgress.IsChecked  = ('silentWithProgress' -in $TopInstallModes)
                    }
                    if ($TopCommands.Count -gt 0 -and -not $txtCommands0.Text) { $txtCommands0.Text = $TopCommands -join ', ' }
                    if ($TopFileExt.Count -gt 0 -and -not $txtFileExt0.Text) { $txtFileExt0.Text = $TopFileExt -join ', ' }
                    if ($TopPlatform.Count -gt 0 -and -not $first['_Platform']) {
                        $chkPlatformDesktop.IsChecked  = ('Windows.Desktop' -in $TopPlatform)
                        $chkPlatformUniversal.IsChecked = ('Windows.Universal' -in $TopPlatform)
                    }
                    # Dependencies
                    if ($TopDepPkgs.Count -gt 0) { $txtPackageDeps.Text = $TopDepPkgs -join ', ' }
                    if ($TopDepFeatures.Count -gt 0) { $txtWindowsFeatures.Text = $TopDepFeatures -join ', ' }
                    # Uninstall switches (from Switch_InstallForHandlerUninstall / Switch_Uninstall)
                    if ($first.Switch_Uninstall) { $txtUninstallCmd.Text = $first.Switch_Uninstall }
                    if ($first.Switch_UninstallSilent -or $first.Switch_UninstallQuiet) {
                        $txtUninstallSilent.Text = if ($first.Switch_UninstallSilent) { $first.Switch_UninstallSilent } else { $first.Switch_UninstallQuiet }
                    }

                    # Additional installers → InstallerEntries
                    for ($idx = 1; $idx -lt $Installers.Count; $idx++) {
                        $ei = $Installers[$idx]
                        $entry = @{
                            Architecture    = $ei.Architecture
                            InstallerType   = $ei.InstallerType
                            InstallerUrl    = $ei.InstallerUrl
                            InstallerSha256 = $ei.InstallerSha256
                            Scope           = $ei.Scope
                            UpgradeBehavior = $ei.UpgradeBehavior
                            ProductCode     = $ei.ProductCode
                            MinimumOSVersion = $ei.MinimumOSVersion
                            ElevationRequirement = $ei.ElevationRequirement
                            InstallerLocale = $ei.InstallerLocale
                            InstallerSwitches = @{
                                Silent             = $ei.Switch_Silent
                                SilentWithProgress = $ei.Switch_SilentWithProgress
                                Custom             = $ei.Switch_Custom
                                Interactive        = $ei.Switch_Interactive
                                Log                = $ei.Switch_Log
                                Repair             = $ei.Switch_Repair
                            }
                        }
                        # Array fields
                        if ($ei['_InstallModes'])    { $entry.InstallModes    = $ei['_InstallModes'] }
                        if ($ei['_Commands'])        { $entry.Commands        = $ei['_Commands'] }
                        if ($ei['_FileExtensions'])  { $entry.FileExtensions  = $ei['_FileExtensions'] }
                        if ($ei['_Platform'])        { $entry.Platform        = $ei['_Platform'] }
                        if ($ei['_Protocols'])       { $entry.Protocols       = $ei['_Protocols'] }
                        # AppsAndFeaturesEntries
                        $af = @{}
                        if ($ei.AF_DisplayName)    { $af.DisplayName    = $ei.AF_DisplayName }
                        if ($ei.AF_Publisher)      { $af.Publisher      = $ei.AF_Publisher }
                        if ($ei.AF_DisplayVersion) { $af.DisplayVersion = $ei.AF_DisplayVersion }
                        if ($ei.AF_ProductCode)    { $af.ProductCode    = $ei.AF_ProductCode }
                        if ($af.Count -gt 0) { $entry.AppsAndFeaturesEntries = @($af) }
                        # NestedInstaller
                        if ($ei.NestedInstallerType) { $entry.NestedInstallerType = $ei.NestedInstallerType }
                        if ($ei.NestedRelPath) {
                            $nif = @{ RelativeFilePath = $ei.NestedRelPath }
                            if ($ei.NestedAlias) { $nif.PortableCommandAlias = $ei.NestedAlias }
                            $entry.NestedInstallerFiles = @($nif)
                        }
                        $Global:InstallerEntries.Add($entry)
                    }
                    Update-InstallerSummaryList
                }
            }

            Update-Stepper 1
            Update-YamlPreview
            Show-Toast "Imported $ResolvedPkgId v$($Data.Version) from community repo" -Type Success -DurationMs 5000
            # Auto-save config (force — user already confirmed import)
            Save-PackageConfig -Force
            # Community import achievements
            Unlock-Achievement 'first_import'
            $Script:CommunityImportCount = if ($Script:CommunityImportCount) { $Script:CommunityImportCount + 1 } else { 1 }
            if ($Script:CommunityImportCount -ge 10) { Unlock-Achievement 'import_10' }
        } catch {
            Show-Toast "Error parsing community manifest: $($_.Exception.Message)" -Type Error
        }
    }
})

# ── Diff Versions Button ──────────────────────────────────────────────────────
$btnDiffVersions.Add_Click({
    $Sel = $lstManagePackages.SelectedItem
    if (-not $Sel -or -not $Sel.Tag) {
        Show-Toast "Select a package first" -Type Warning
        return
    }
    $Versions = @($Sel.Tag.Versions | Sort-Object -Descending)
    if ($Versions.Count -lt 2) {
        Show-Toast "Need at least 2 versions to compare" -Type Warning
        return
    }

    $PkgId = $Sel.Tag.Id
    $V1 = $Versions[0]
    $V2 = $Versions[1]

    Show-Toast "Loading diff for $PkgId ($V2 vs $V1)..." -Type Info -DurationMs 3000
    Unlock-Achievement 'diff_viewer'

    Start-BackgroundWork -Variables @{
        FuncName  = Get-ActiveRestSource
        DiffPkgId = $PkgId
        DiffV1 = $V1
        DiffV2 = $V2
    } -Work {
        try {
            Import-Module Microsoft.WinGet.RestSource -ErrorAction Stop
            $M1 = Get-WinGetManifest -FunctionName $FuncName -PackageIdentifier $DiffPkgId -Version $DiffV1 -ErrorAction Stop
            $M2 = Get-WinGetManifest -FunctionName $FuncName -PackageIdentifier $DiffPkgId -Version $DiffV2 -ErrorAction Stop
            return @{ M1 = ($M1 | ConvertTo-Json -Depth 10); M2 = ($M2 | ConvertTo-Json -Depth 10); V1 = $DiffV1; V2 = $DiffV2; Error = $null }
        } catch {
            return @{ Error = $_.Exception.Message }
        }
    } -OnComplete {
        param($Results, $Errors)
        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ Error = 'No result' } }
        if ($R.Error) {
            Show-Toast "Diff failed: $($R.Error)" -Type Error
            return
        }
        # Show diff in a themed dialog with side-by-side
        $DiffContent = "=== Version $($R.V2) ===`n$($R.M2)`n`n=== Version $($R.V1) ===`n$($R.M1)"
        Show-ThemedDialog -Title "Package Diff: $PkgId" -Message "Comparing $($R.V2) (older) vs $($R.V1) (newer)" `
            -Icon ([string]([char]0xE8FD)) -IconColor 'ThemeAccent' `
            -Buttons @( @{ Text='Close'; IsAccent=$false; Result='Close' } ) `
            -Width 700 -Height 500 -ExtraContent {
                param($SP)
                $DiffBox = New-Object System.Windows.Controls.TextBox
                $DiffBox.Text = $DiffContent
                $DiffBox.IsReadOnly = $true
                $DiffBox.AcceptsReturn = $true
                $DiffBox.TextWrapping = 'NoWrap'
                $DiffBox.FontFamily = New-Object System.Windows.Media.FontFamily("Cascadia Code, Consolas")
                $DiffBox.FontSize = 10
                $DiffBox.Background = $Window.Resources['ThemeOutputBg']
                $DiffBox.Foreground = $Window.Resources['ThemeTextBody']
                $DiffBox.BorderThickness = [System.Windows.Thickness]::new(1)
                $DiffBox.BorderBrush = $Window.Resources['ThemeBorderCard']
                $DiffBox.Padding = [System.Windows.Thickness]::new(8)
                $DiffBox.MaxHeight = 350
                $DiffBox.VerticalScrollBarVisibility = 'Auto'
                $DiffBox.HorizontalScrollBarVisibility = 'Auto'
                $SP.Children.Add($DiffBox) | Out-Null
            }
    }
})


function Toggle-LeftPanel {
    param([bool]$Show)
    $Global:LeftPanelOpen = $Show
    if ($Show) {
        $colLeftPanel.Width = [System.Windows.GridLength]::new(260)
        $pnlLeftSidebar.Visibility = 'Visible'
        $splitterLeft.Visibility = 'Visible'
    } else {
        $colLeftPanel.Width = [System.Windows.GridLength]::new(0)
        $pnlLeftSidebar.Visibility = 'Collapsed'
        $splitterLeft.Visibility = 'Collapsed'
    }
}

$btnHamburger.Add_Click({
    Toggle-LeftPanel (-not $Global:LeftPanelOpen)
})

# Left rail section toggles
$Global:LeftActiveSection = 'Configs'

function Show-LeftSection {
    param([string]$Section)
    $Global:LeftActiveSection = $Section
    $grpLeftConfigs.Visibility  = if ($Section -eq 'Configs')  { 'Visible' } else { 'Collapsed' }
    $grpLeftStorage.Visibility  = if ($Section -eq 'Storage')  { 'Visible' } else { 'Collapsed' }
    $grpLeftSettings.Visibility = if ($Section -eq 'Settings') { 'Visible' } else { 'Collapsed' }
    if (-not $Global:LeftPanelOpen) { Toggle-LeftPanel $true }
}

$railConfigs.Add_Click({ Show-LeftSection 'Configs' })
$railStorage.Add_Click({ Show-LeftSection 'Storage' })
$railSettings.Add_Click({ Show-LeftSection 'Settings' })

# ── Bottom Panel Toggle ───────────────────────────────────────────────────────
$Global:BottomPanelOpen = $true
$Global:BottomPanelSavedHeight = 160

function Toggle-BottomPanel {
    param([bool]$Show)
    $Global:BottomPanelOpen = $Show
    if ($Show) {
        $h = if ($Global:BottomPanelSavedHeight -gt 30) { $Global:BottomPanelSavedHeight } else { 160 }
        $rowBottomPanel.Height = [System.Windows.GridLength]::new($h)
        $pnlBottomPanel.Visibility = 'Visible'
        $splitterBottom.Visibility = 'Visible'
    } else {
        $Global:BottomPanelSavedHeight = $rowBottomPanel.Height.Value
        $rowBottomPanel.Height = [System.Windows.GridLength]::new(0)
        $pnlBottomPanel.Visibility = 'Collapsed'
        $splitterBottom.Visibility = 'Collapsed'
    }
}

$btnHideBottom.Add_Click({
    Toggle-BottomPanel $false
})

# Rail button to toggle output panel
$railOutput = $Window.FindName("railOutput")
$railOutput.Add_Click({
    Toggle-BottomPanel (-not $Global:BottomPanelOpen)
})

$Global:BottomMaximized = $false
$btnToggleBottomSize.Add_Click({
    if ($Global:BottomMaximized) {
        # Restore
        $h = if ($Global:BottomPanelSavedHeight -gt 30) { $Global:BottomPanelSavedHeight } else { 160 }
        $rowBottomPanel.Height = [System.Windows.GridLength]::new($h)
        $icoToggleBottomSize.Text = [string][char]0xE740
        $Global:BottomMaximized = $false
    } else {
        # Maximize — take most of the window
        $Global:BottomPanelSavedHeight = $rowBottomPanel.Height.Value
        $rowBottomPanel.Height = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $icoToggleBottomSize.Text = [string][char]0xE73F
        $Global:BottomMaximized = $true
    }
})

# ── Debug Overlay Toggle ──────────────────────────────────────────────────────
$Global:DebugOverlayEnabled = [bool]$chkDebug.IsChecked

function Rebuild-LogPanel {
    <# Rebuilds the RichTextBox content from the full log buffer #>
    param([bool]$IncludeDebug)
    if (-not $paraLog) { return }
    $paraLog.Inlines.Clear()
    $Global:DebugLineCount = 0
    # Reset minimap + heatmap for full rebuild
    $Global:MinimapDots.Clear()
    $Global:MinimapLineTotal = 0
    $Global:MinimapLastRenderedIdx = 0
    $Global:HeatmapData.Clear()
    $Global:HeatmapLastRenderedIdx = 0
    $cnvMM = $Window.FindName('cnvMinimap'); if ($cnvMM) { $cnvMM.Children.Clear() }
    $cnvHM = $Window.FindName('cnvHeatmap'); if ($cnvHM) { $cnvHM.Children.Clear() }
    $AllLines = $Global:FullLogSB.ToString() -split "`n" | Where-Object { $_.Trim() }
    if (-not $IncludeDebug) { $AllLines = $AllLines | Where-Object { $_ -notmatch '\[DEBUG\]' } }
    $Converter = [System.Windows.Media.BrushConverter]::new()
    $IsLight = $Global:IsLightMode
    foreach ($L in $AllLines) {
        if ($paraLog.Inlines.Count -gt 0) {
            $paraLog.Inlines.Add((New-Object System.Windows.Documents.LineBreak))
        }
        $Color = if ($IsLight) {
            if ($L -match '\[ERROR\]') { '#CC0000' }
            elseif ($L -match '\[WARN\]') { '#B86E00' }
            elseif ($L -match '\[SUCCESS\]') { '#008A2E' }
            elseif ($L -match '\[DEBUG\]') { '#8888AA' }
            else { '#444444' }
        } else {
            if ($L -match '\[ERROR\]') { '#FF4040' }
            elseif ($L -match '\[WARN\]') { '#FF9100' }
            elseif ($L -match '\[SUCCESS\]') { '#16C60C' }
            elseif ($L -match '\[DEBUG\]') { '#B8860B' }
            else { '#888888' }
        }
        $Run = New-Object System.Windows.Documents.Run($L.TrimEnd())
        $Run.Foreground = $Converter.ConvertFromString($Color)
        $paraLog.Inlines.Add($Run)
        $Global:DebugLineCount++
        # Rebuild minimap/heatmap data
        $Global:MinimapLineTotal = $Global:DebugLineCount
        if ($L -match '\[ERROR\]' -or $L -match '\[WARN\]' -or $L -match '\[SUCCESS\]') {
            $Global:MinimapDots.Add([PSCustomObject]@{ Color = $Color; Line = $Global:DebugLineCount })
            $Global:HeatmapData.Add([PSCustomObject]@{ Color = $Color; Line = $Global:DebugLineCount })
        }
    }
    $logScroller.ScrollToEnd()
    Render-Minimap
    Render-Heatmap
}

$chkDebug.Add_Checked({
    $Global:DebugOverlayEnabled = $true
    Rebuild-LogPanel -IncludeDebug $true
    if (-not $Global:BottomPanelOpen) { Toggle-BottomPanel $true }
    Write-DebugLog "Debug overlay enabled — verbose logging active"
})
$chkDebug.Add_Unchecked({
    $Global:DebugOverlayEnabled = $false
    Rebuild-LogPanel -IncludeDebug $false
})

$chkDisableAnimations.Add_Checked({
    $Global:AnimationsDisabled = $true
    Write-DebugLog "Animations disabled"
})
$chkDisableAnimations.Add_Unchecked({
    $Global:AnimationsDisabled = $false
    Write-DebugLog "Animations enabled"
})

# ── File Browse Buttons ──────────────────────────────────────────────────────
$btnBrowseFile0.Add_Click({
    $Dialog = New-Object System.Windows.Forms.OpenFileDialog
    $Dialog.Title = "Select Installer File"
    $Dialog.Filter = "Installer Files|*.exe;*.msi;*.msix;*.zip;*.appx;*.appxbundle|All Files|*.*"
    if ($Dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtLocalFile0.Text = $Dialog.FileName
        Compute-FileHash -FilePath $Dialog.FileName -TargetField 'installer0'
        # Sync to Upload tab if empty
        if (-not $txtUploadFile.Text) {
            $txtUploadFile.Text = $Dialog.FileName
            Compute-FileHash -FilePath $Dialog.FileName -TargetField 'upload'
        }
        # Auto-suggest blob path (container already provides the 'packages' namespace)
        $FileName = Split-Path $Dialog.FileName -Leaf
        $PkgId = $txtPkgId.Text.Trim()
        $PkgVer = $txtPkgVersion.Text.Trim()
        if ($PkgId -and $PkgVer) {
            $txtBlobPath.Text = "$PkgId/$PkgVer/$FileName"
        } elseif (-not $txtBlobPath.Text) {
            $txtBlobPath.Text = $FileName
        }
    }
})

$btnBrowseUpload.Add_Click({
    $Dialog = New-Object System.Windows.Forms.OpenFileDialog
    $Dialog.Title = "Select File to Upload"
    $Dialog.Filter = "All Files|*.*"
    if ($Dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtUploadFile.Text = $Dialog.FileName
        Compute-FileHash -FilePath $Dialog.FileName -TargetField 'upload'
        # Always recalculate blob path from current PkgId/PkgVer/FileName
        $FileName = Split-Path $Dialog.FileName -Leaf
        $PkgId = $txtPkgId.Text.Trim()
        $PkgVer = $txtPkgVersion.Text.Trim()
        if ($PkgId -and $PkgVer) {
            $txtBlobPath.Text = "$PkgId/$PkgVer/$FileName"
        } else {
            $txtBlobPath.Text = $FileName
        }
    }
})

$btnBrowseManifestPath.Add_Click({
    $Dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $Dialog.Description = "Select default manifest output folder"
    if ($Dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtManifestPath.Text = $Dialog.SelectedPath
    }
})

# ── Create Tab Actions ───────────────────────────────────────────────────────
$btnImportPSADT.Add_Click({
    # Ask the user: supply a folder or a zip?
    $SourceChoice = Show-ThemedDialog -Title 'Import PSADT Package' `
        -Message "Select the source of your PSADT package:" `
        -Icon ([string]([char]0xE8B5)) -IconColor 'ThemeAccentLight' `
        -Buttons @(
            @{ Text='Browse Folder'; IsAccent=$true; Result='Folder' },
            @{ Text='Select ZIP';    IsAccent=$false; Result='Zip' },
            @{ Text='Cancel';        IsAccent=$false; Result='Cancel' }
        )
    if ($SourceChoice -eq 'Cancel') { return }

    $PsadtFolder = $null
    $TempCleanup = $null

    if ($SourceChoice -eq 'Zip') {
        $OFD = New-Object System.Windows.Forms.OpenFileDialog
        $OFD.Title  = 'Select PSADT Package ZIP'
        $OFD.Filter = 'ZIP files (*.zip)|*.zip|All files (*.*)|*.*'
        if ($OFD.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $Extracted = Expand-PSADTZip -ZipPath $OFD.FileName
        if (-not $Extracted) { return }
        $PsadtFolder = $Extracted.FolderPath
        $TempCleanup = $Extracted.TempRoot
    } else {
        $FBD = New-Object System.Windows.Forms.FolderBrowserDialog
        $FBD.Description = 'Select PSADT package folder (containing Deploy-Application.ps1)'
        if ($FBD.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $PsadtFolder = $FBD.SelectedPath
    }

    $PsadtData = Import-PSADTPackage -FolderPath $PsadtFolder
    if (-not $PsadtData) {
        if ($TempCleanup) { Remove-Item $TempCleanup -Recurse -Force -ErrorAction SilentlyContinue }
        return
    }

    # Always use PSADT zip mode — the whole point of importing PSADT is to
    # preserve the full deployment flow (pre/post scripts, CloseApps, etc.)
    Populate-CreateTabFromPSADT -PsadtData $PsadtData -Mode 'psadt'

    # Clean up temp extraction
    if ($TempCleanup) { Remove-Item $TempCleanup -Recurse -Force -ErrorAction SilentlyContinue }
})
$btnSaveToDisk.Add_Click({ Save-ManifestToDisk })
$btnSaveConfig.Add_Click({ Save-PackageConfig })

$btnPublish.Add_Click({
  try {
    Write-DebugLog "[PUBLISH] === Publish button clicked ===" -Level 'DEBUG'
    # Concurrent publish guard
    if ($Script:PublishInProgress) {
        Show-Toast "A publish operation is already in progress" -Type Warning
        return
    }

    Write-DebugLog "[PUBLISH] Validating fields..." -Level 'DEBUG'
    $ValidationErrors = Validate-ManifestFields -ForPublish
    if ($ValidationErrors.Count -gt 0) {
        Show-ThemedDialog -Title 'Validation Error' -Message ($ValidationErrors -join "`n") `
            -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' | Out-Null
        Invoke-ShakeAnimation -Element $btnPublish
        return
    }
    # Save the 3-file split to disk for user reference
    Write-DebugLog "[PUBLISH] Calling Save-ManifestToDisk..." -Level 'DEBUG'
    Save-ManifestToDisk
    Write-DebugLog "[PUBLISH] Save-ManifestToDisk returned" -Level 'DEBUG'

    $FuncName = Get-ActiveRestSource
    if (-not $FuncName) {
        Show-ThemedDialog -Title 'Missing Setting' -Message 'REST Source Function Name is required. Set it in Settings.' `
            -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' | Out-Null
        return
    }

    # Check if InstallerUrl looks like a blob storage URL — warn if it's an external vendor URL
    $InstallerUrl0 = $txtInstallerUrl0.Text.Trim()
    $StorageSel = $cmbUploadStorage.SelectedItem
    $StorageName = if ($StorageSel -and $StorageSel.Tag) { $StorageSel.Tag.Name } else { '' }
    if ($InstallerUrl0 -and $StorageName -and $InstallerUrl0 -notmatch [regex]::Escape("$StorageName.blob.core.windows.net")) {
        $BlobAnswer = Show-ThemedDialog -Title 'Installer URL Warning' `
            -Message "The Installer URL does not point to your Azure Storage account ($StorageName).`n`nURL: $InstallerUrl0`n`nThe REST source API will fail if it cannot access the installer binary. Upload the package to blob storage first (Upload tab)." `
            -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' `
            -Buttons @(
                @{ Text='Cancel'; IsAccent=$false; Result='Cancel' }
                @{ Text='Publish Anyway'; IsAccent=$true; Result='PublishAnyway' }
            )
        if ($BlobAnswer -ne 'PublishAnyway') { return }
        Write-DebugLog "[PUBLISH] User chose to publish with non-blob URL: $InstallerUrl0" -Level 'WARN'
    }

    $Script:PublishInProgress = $true
    $Script:PublishStartTime = Get-Date
    Set-ButtonBusy -Button $btnPublish -Text 'Publishing...'
    Start-StatusPulse

    # Build split manifests (version + installer + defaultLocale) for REST source publishing
    # The REST source API requires split format, not singleton
    $PkgId  = $txtPkgId.Text.Trim()
    $PkgVer = $txtPkgVersion.Text.Trim()
    $ManVer = Get-ComboBoxSelectedText $cmbCreateManifestVer
    if (-not $ManVer) { $ManVer = "1.9.0" }

    $Meta = Get-MetadataFromUI
    $AllInstallers = Get-AllInstallerData

    $VersionYaml   = Build-VersionYaml   -PkgId $PkgId -PkgVersion $PkgVer -ManifestVersion $ManVer -Locale "en-US"
    $InstallerYaml = Build-InstallerYaml  -PkgId $PkgId -PkgVersion $PkgVer -ManifestVersion $ManVer -Installers $AllInstallers
    $LocaleYaml    = Build-DefaultLocaleYaml -PkgId $PkgId -PkgVersion $PkgVer -ManifestVersion $ManVer -Meta $Meta

    # Write split manifests to a temp directory for publishing
    $TempDir = Join-Path ([System.IO.Path]::GetTempPath()) "WinGetMM_publish_$([guid]::NewGuid().ToString('N'))"
    New-Item -Path $TempDir -ItemType Directory -Force | Out-Null
    $VersionYaml   | Set-Content (Join-Path $TempDir "$PkgId.yaml") -Force -Encoding UTF8
    $InstallerYaml | Set-Content (Join-Path $TempDir "$PkgId.installer.yaml") -Force -Encoding UTF8
    $LocaleYaml    | Set-Content (Join-Path $TempDir "$PkgId.locale.en-US.yaml") -Force -Encoding UTF8

    Write-DebugLog "[PUBLISH] Calling Publish-ManifestToRestSource (FuncName=$FuncName, Path=$TempDir)" -Level 'DEBUG'
    Publish-ManifestToRestSource -ManifestPath $TempDir -FunctionName $FuncName
    Write-DebugLog "[PUBLISH] Publish-ManifestToRestSource returned (bg job queued)" -Level 'DEBUG'
    Write-DebugLog "[PUBLISH] === Publish click handler finished OK ===" -Level 'DEBUG'
  } catch {
    Write-DebugLog "Publish click error: $($_.Exception.Message)" -Level 'ERROR'
    Write-DebugLog "Publish click ScriptStackTrace: $($_.ScriptStackTrace)" -Level 'ERROR'
    Show-Toast "Unexpected error — see Activity Log" -Type Error
    $Script:PublishInProgress = $false
    Reset-ButtonBusy -Button $btnPublish
    Stop-StatusPulse
  }
})

$btnAddInstaller.Add_Click({
    # Capture current installer card data and add to the entry list
    $Inst = Get-InstallerDataFromUI
    if (-not $Inst.InstallerUrl -and -not $Global:CurrentHash0) {
        Show-Toast "Fill in installer URL or select a local file first" -Type Warning
        return
    }
    $Inst.InstallerSha256 = $Global:CurrentHash0
    $Global:InstallerEntries.Add($Inst)
    Update-InstallerSummaryList
    Show-Toast "Installer #$($Global:InstallerEntries.Count) added — fill in the next one" -Type Success

    # Clear the installer card for the next entry
    $cmbInstallerType0.SelectedIndex = 0
    $cmbArch0.SelectedIndex = 0
    $cmbScope0.SelectedIndex = 0
    $txtInstallerUrl0.Text = ""
    $txtLocalFile0.Text = ""
    $txtSilent0.Text = ""
    $txtSilentProgress0.Text = ""
    $txtCustomSwitch0.Text = ""
    $txtInteractive0.Text = ""
    $txtLog0.Text = ""
    $txtRepair0.Text = ""
    $txtProductCode0.Text = ""
    $txtCommands0.Text = ""
    $txtFileExt0.Text = ""
    $txtMinOSVersion0.Text = ""
    $txtExpectedReturnCodes0.Text = ""
    $txtAFDisplayName0.Text = ""
    $txtAFPublisher0.Text = ""
    $txtAFDisplayVersion0.Text = ""
    $txtAFProductCode0.Text = ""
    $chkModeInteractive.IsChecked = $false
    $chkModeSilent.IsChecked = $false
    $chkModeSilentProgress.IsChecked = $false
    $chkInstallLocationRequired.IsChecked = $false
    $chkPlatformDesktop.IsChecked = $true
    $chkPlatformUniversal.IsChecked = $false
    $cmbUpgrade0.SelectedIndex = 0
    $cmbElevation0.SelectedIndex = 0
    $lblHash0.Text = ""
    $Global:CurrentHash0 = ""
    $pnlNestedInstaller.Visibility = 'Collapsed'
    Update-YamlPreview
})

$btnRemoveLastInstaller = $Window.FindName("btnRemoveLastInstaller")
$btnRemoveLastInstaller.Add_Click({
    if ($Global:InstallerEntries.Count -gt 0) {
        $Global:InstallerEntries.RemoveAt($Global:InstallerEntries.Count - 1)
        Update-InstallerSummaryList
        Show-Toast "Last installer entry removed" -Type Info
        Update-YamlPreview
    }
})

# ── Copy YAML to Clipboard ──────────────────────────────────────────────────
$btnCopyYaml.Add_Click({
    $Yaml = $txtYamlPreview.Text
    if ($Yaml) {
        [System.Windows.Clipboard]::SetText($Yaml)
        Show-Toast "YAML copied to clipboard" -Type Success -DurationMs 2000
    } else {
        Show-Toast "YAML preview is empty" -Type Warning
    }
})

# ── Installer Reorder & Edit ────────────────────────────────────────────────
function Set-InstallerUIFromEntry {
    <# Populates the installer card from a hashtable entry #>
    param([hashtable]$Entry)
    # Type
    for ($i = 0; $i -lt $cmbInstallerType0.Items.Count; $i++) {
        if ($cmbInstallerType0.Items[$i].Content -eq $Entry.InstallerType) { $cmbInstallerType0.SelectedIndex = $i; break }
    }
    # Architecture
    for ($i = 0; $i -lt $cmbArch0.Items.Count; $i++) {
        if ($cmbArch0.Items[$i].Content -eq $Entry.Architecture) { $cmbArch0.SelectedIndex = $i; break }
    }
    # Scope
    for ($i = 0; $i -lt $cmbScope0.Items.Count; $i++) {
        if ($cmbScope0.Items[$i].Content -eq $Entry.Scope) { $cmbScope0.SelectedIndex = $i; break }
    }
    # Upgrade behavior
    if ($Entry.UpgradeBehavior) {
        for ($i = 0; $i -lt $cmbUpgrade0.Items.Count; $i++) {
            if ($cmbUpgrade0.Items[$i].Content -eq $Entry.UpgradeBehavior) { $cmbUpgrade0.SelectedIndex = $i; break }
        }
    }
    # Elevation
    if ($Entry.ElevationRequirement) {
        for ($i = 0; $i -lt $cmbElevation0.Items.Count; $i++) {
            if ($cmbElevation0.Items[$i].Content -eq $Entry.ElevationRequirement) { $cmbElevation0.SelectedIndex = $i; break }
        }
    }
    $txtInstallerUrl0.Text       = if ($Entry.InstallerUrl) { $Entry.InstallerUrl } else { '' }
    $txtSilent0.Text             = if ($Entry.Silent) { $Entry.Silent } else { '' }
    $txtSilentProgress0.Text     = if ($Entry.SilentWithProgress) { $Entry.SilentWithProgress } else { '' }
    $txtCustomSwitch0.Text       = if ($Entry.Custom) { $Entry.Custom } else { '' }
    $txtInteractive0.Text        = if ($Entry.Interactive) { $Entry.Interactive } else { '' }
    $txtLog0.Text                = if ($Entry.Log) { $Entry.Log } else { '' }
    $txtRepair0.Text             = if ($Entry.Repair) { $Entry.Repair } else { '' }
    $txtProductCode0.Text        = if ($Entry.ProductCode) { $Entry.ProductCode } else { '' }
    $txtCommands0.Text           = if ($Entry.Commands) { $Entry.Commands } else { '' }
    $txtFileExt0.Text            = if ($Entry.FileExtensions) { $Entry.FileExtensions } else { '' }
    $txtMinOSVersion0.Text       = if ($Entry.MinimumOSVersion) { $Entry.MinimumOSVersion } else { '' }
    $txtExpectedReturnCodes0.Text = if ($Entry.ExpectedReturnCodes) { $Entry.ExpectedReturnCodes } else { '' }
    $txtAFDisplayName0.Text      = if ($Entry.AF_DisplayName) { $Entry.AF_DisplayName } else { '' }
    $txtAFPublisher0.Text        = if ($Entry.AF_Publisher) { $Entry.AF_Publisher } else { '' }
    $txtAFDisplayVersion0.Text   = if ($Entry.AF_DisplayVersion) { $Entry.AF_DisplayVersion } else { '' }
    $txtAFProductCode0.Text      = if ($Entry.AF_ProductCode) { $Entry.AF_ProductCode } else { '' }
    $chkModeInteractive.IsChecked      = $Entry.InstallerSuccessCodes -contains 'interactive'
    $chkModeSilent.IsChecked           = $Entry.InstallerSuccessCodes -contains 'silent'
    $chkModeSilentProgress.IsChecked   = $Entry.InstallerSuccessCodes -contains 'silentWithProgress'
    $lblHash0.Text = if ($Entry.InstallerSha256) { "SHA256: $($Entry.InstallerSha256)" } else { '' }
    $Global:CurrentHash0 = if ($Entry.InstallerSha256) { $Entry.InstallerSha256 } else { '' }
}

$btnEditInstaller.Add_Click({
    $Sel = $lstInstallerEntries.SelectedItem
    if (-not $Sel -or -not $Sel.Tag) {
        Show-Toast "Select an installer entry to edit" -Type Warning
        return
    }
    $SelIdx = $lstInstallerEntries.SelectedIndex
    if ($SelIdx -lt 0 -or $SelIdx -ge $Global:InstallerEntries.Count) { return }
    $Entry = $Global:InstallerEntries[$SelIdx]
    # Populate the installer card from the selected entry
    Set-InstallerUIFromEntry -Entry $Entry
    # Remove from list — user will re-add after editing
    $Global:InstallerEntries.RemoveAt($SelIdx)
    Update-InstallerSummaryList
    Show-Toast "Editing installer #$($SelIdx + 1) — modify and click Add to save" -Type Info
    Update-YamlPreview
})

$btnMoveInstallerUp.Add_Click({
    $SelIdx = $lstInstallerEntries.SelectedIndex
    if ($SelIdx -le 0) { return }
    $Item = $Global:InstallerEntries[$SelIdx]
    $Global:InstallerEntries.RemoveAt($SelIdx)
    $Global:InstallerEntries.Insert($SelIdx - 1, $Item)
    Update-InstallerSummaryList
    $lstInstallerEntries.SelectedIndex = $SelIdx - 1
    Update-YamlPreview
})

$btnMoveInstallerDown.Add_Click({
    $SelIdx = $lstInstallerEntries.SelectedIndex
    if ($SelIdx -lt 0 -or $SelIdx -ge ($Global:InstallerEntries.Count - 1)) { return }
    $Item = $Global:InstallerEntries[$SelIdx]
    $Global:InstallerEntries.RemoveAt($SelIdx)
    $Global:InstallerEntries.Insert($SelIdx + 1, $Item)
    Update-InstallerSummaryList
    $lstInstallerEntries.SelectedIndex = $SelIdx + 1
    Update-YamlPreview
})

# ── Installer Preset Logic ───────────────────────────────────────────────────
$btnApplyPreset.Add_Click({
    $Preset = Get-ComboBoxSelectedText $cmbInstallerPreset
    if (-not $Preset -or $Preset -eq '(none)') { return }

    Write-DebugLog "Applying installer preset: $Preset"

    switch -Wildcard ($Preset) {
        'PSADT (Deploy-Application.exe)' {
            Set-ComboBoxByContent $cmbInstallerType0 'zip'
            Set-ComboBoxByContent $cmbScope0 'machine'
            Set-ComboBoxByContent $cmbElevation0 'elevationRequired'
            Set-ComboBoxByContent $cmbUpgrade0 'install'
            $txtSilent0.Text = ''
            $txtSilentProgress0.Text = ''
            $txtCustomSwitch0.Text = '-DeploymentType "Install" -DeployMode "Silent"'
            $txtInteractive0.Text = '-DeploymentType "Install" -DeployMode "Interactive"'
            $txtLog0.Text = ''
            $txtRepair0.Text = ''
            $chkModeInteractive.IsChecked = $true
            $chkModeSilent.IsChecked = $true
            $chkModeSilentProgress.IsChecked = $false
            # Nested installer
            Set-ComboBoxByContent $cmbNestedType0 'exe'
            $txtNestedPath0.Text = 'Deploy-Application.exe'
            $txtPortableAlias0.Text = ''
        }
        'PSADT (PowerShell wrapper)' {
            Set-ComboBoxByContent $cmbInstallerType0 'zip'
            Set-ComboBoxByContent $cmbScope0 'machine'
            Set-ComboBoxByContent $cmbElevation0 'elevationRequired'
            Set-ComboBoxByContent $cmbUpgrade0 'install'
            $txtSilent0.Text = ''
            $txtSilentProgress0.Text = ''
            $txtCustomSwitch0.Text = '-DeploymentType "Install" -DeployMode "Silent"'
            $txtInteractive0.Text = '-DeploymentType "Install" -DeployMode "Interactive"'
            $txtLog0.Text = ''
            $txtRepair0.Text = ''
            $chkModeInteractive.IsChecked = $true
            $chkModeSilent.IsChecked = $true
            $chkModeSilentProgress.IsChecked = $false
            # Nested: PowerShell script (treated as exe with pwsh invocation)
            Set-ComboBoxByContent $cmbNestedType0 'exe'
            $txtNestedPath0.Text = 'Deploy-Application.exe'
            $txtPortableAlias0.Text = ''
        }
        'Standard MSI' {
            Set-ComboBoxByContent $cmbInstallerType0 'msi'
            Set-ComboBoxByContent $cmbScope0 'machine'
            Set-ComboBoxByContent $cmbElevation0 'elevationRequired'
            Set-ComboBoxByContent $cmbUpgrade0 'install'
            $txtSilent0.Text = '/qn /norestart'
            $txtSilentProgress0.Text = '/qb /norestart'
            $txtCustomSwitch0.Text = ''
            $txtInteractive0.Text = '/qf'
            $txtLog0.Text = '/l*v "%TEMP%\MSIInstall.log"'
            $txtRepair0.Text = '/f'
            $chkModeInteractive.IsChecked = $true
            $chkModeSilent.IsChecked = $true
            $chkModeSilentProgress.IsChecked = $true
        }
        'Standard EXE (Inno Setup)' {
            Set-ComboBoxByContent $cmbInstallerType0 'inno'
            Set-ComboBoxByContent $cmbElevation0 'elevationRequired'
            $txtSilent0.Text = '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-'
            $txtSilentProgress0.Text = '/SILENT /SUPPRESSMSGBOXES /NORESTART /SP-'
            $txtCustomSwitch0.Text = ''
            $txtInteractive0.Text = ''
            $txtLog0.Text = '/LOG="%TEMP%\InnoInstall.log"'
            $txtRepair0.Text = ''
            $chkModeInteractive.IsChecked = $true
            $chkModeSilent.IsChecked = $true
            $chkModeSilentProgress.IsChecked = $true
        }
        'Standard EXE (NSIS)' {
            Set-ComboBoxByContent $cmbInstallerType0 'nullsoft'
            Set-ComboBoxByContent $cmbElevation0 'elevationRequired'
            $txtSilent0.Text = '/S'
            $txtSilentProgress0.Text = '/S'
            $txtCustomSwitch0.Text = ''
            $txtInteractive0.Text = ''
            $txtLog0.Text = ''
            $txtRepair0.Text = ''
            $chkModeInteractive.IsChecked = $true
            $chkModeSilent.IsChecked = $true
            $chkModeSilentProgress.IsChecked = $false
        }
        'Burn Bundle' {
            Set-ComboBoxByContent $cmbInstallerType0 'burn'
            Set-ComboBoxByContent $cmbElevation0 'elevationRequired'
            Set-ComboBoxByContent $cmbUpgrade0 'install'
            $txtSilent0.Text = '/quiet /norestart'
            $txtSilentProgress0.Text = '/passive /norestart'
            $txtCustomSwitch0.Text = ''
            $txtInteractive0.Text = ''
            $txtLog0.Text = '/log "%TEMP%\BurnInstall.log"'
            $txtRepair0.Text = '/repair'
            $chkModeInteractive.IsChecked = $true
            $chkModeSilent.IsChecked = $true
            $chkModeSilentProgress.IsChecked = $true
        }
        'MSIX / AppX' {
            Set-ComboBoxByContent $cmbInstallerType0 'msix'
            Set-ComboBoxByContent $cmbScope0 'user'
            Set-ComboBoxByContent $cmbElevation0 'elevatesSelf'
            $txtSilent0.Text = ''
            $txtSilentProgress0.Text = ''
            $txtCustomSwitch0.Text = ''
            $txtInteractive0.Text = ''
            $txtLog0.Text = ''
            $txtRepair0.Text = ''
            $chkModeInteractive.IsChecked = $false
            $chkModeSilent.IsChecked = $true
            $chkModeSilentProgress.IsChecked = $false
        }
        'Portable (ZIP)' {
            Set-ComboBoxByContent $cmbInstallerType0 'zip'
            Set-ComboBoxByContent $cmbScope0 'user'
            $txtSilent0.Text = ''
            $txtSilentProgress0.Text = ''
            $txtCustomSwitch0.Text = ''
            $txtInteractive0.Text = ''
            $txtLog0.Text = ''
            $txtRepair0.Text = ''
            Set-ComboBoxByContent $cmbNestedType0 'portable'
            $txtNestedPath0.Text = ''
            $txtPortableAlias0.Text = ''
            $chkModeInteractive.IsChecked = $false
            $chkModeSilent.IsChecked = $true
            $chkModeSilentProgress.IsChecked = $false
        }
    }

    Update-YamlPreview
    $lblStatus.Text = "Preset applied: $Preset"
})

# ── Metadata Template Handler ────────────────────────────────────────────────
$cmbMetadataTemplate = $Window.FindName("cmbMetadataTemplate")
$btnApplyTemplate    = $Window.FindName("btnApplyTemplate")

$Global:MetadataTemplatesPath = Join-Path $Global:Root "metadata_templates.json"

$btnApplyTemplate.Add_Click({
    $Tmpl = Get-ComboBoxSelectedText $cmbMetadataTemplate
    if (-not $Tmpl -or $Tmpl -eq '(none)') { return }

    Write-DebugLog "Applying metadata template: $Tmpl"

    switch ($Tmpl) {
        'Internal LOB App' {
            $txtPublisher.Text    = $txtPublisher.Text
            $txtLicense.Text      = "Proprietary"
            $txtCopyright.Text    = "Copyright (c) $(Get-Date -Format yyyy) $($txtPublisher.Text)"
            $txtAuthor.Text       = $txtPublisher.Text
            $txtDescription.Text  = if (-not $txtDescription.Text.Trim()) { "Internal line-of-business application" } else { $txtDescription.Text }
        }
        'Open Source Tool' {
            $txtLicense.Text      = if (-not $txtLicense.Text.Trim()) { "MIT" } else { $txtLicense.Text }
            $txtCopyright.Text    = "Copyright (c) $(Get-Date -Format yyyy) $($txtPublisher.Text)"
        }
        'Commercial Software' {
            $txtLicense.Text      = if (-not $txtLicense.Text.Trim()) { "Proprietary" } else { $txtLicense.Text }
            $txtCopyright.Text    = "Copyright (c) $(Get-Date -Format yyyy) $($txtPublisher.Text)"
        }
        'Custom...' {
            if (-not (Test-Path $Global:MetadataTemplatesPath)) {
                # Create a sample file and open it
                $Sample = @{
                    templates = @(
                        @{
                            name        = "My Org Default"
                            publisher   = "Contoso Ltd."
                            publisherUrl = "https://contoso.com"
                            supportUrl  = "https://contoso.com/support"
                            privacyUrl  = "https://contoso.com/privacy"
                            license     = "Proprietary"
                            author      = "Contoso IT"
                            copyright   = "Copyright (c) {YEAR} Contoso Ltd."
                        }
                    )
                } | ConvertTo-Json -Depth 5
                $Sample | Set-Content $Global:MetadataTemplatesPath -Force -Encoding UTF8
                Show-Toast "Created metadata_templates.json — edit it, save, then re-apply" -Type Info -DurationMs 5000
                Start-Process $Global:MetadataTemplatesPath
                return
            }
            try {
                $CustomTemplates = (Get-Content $Global:MetadataTemplatesPath -Raw | ConvertFrom-Json).templates
                if (-not $CustomTemplates -or $CustomTemplates.Count -eq 0) {
                    Show-Toast "No templates found in metadata_templates.json" -Type Warning
                    return
                }
                $ct = $CustomTemplates[0]
                if ($ct.publisher)    { $txtPublisher.Text    = $ct.publisher }
                if ($ct.publisherUrl) { $txtPublisherUrl.Text = $ct.publisherUrl }
                if ($ct.supportUrl)   { $txtSupportUrl.Text   = $ct.supportUrl }
                if ($ct.privacyUrl)   { $txtPrivacyUrl.Text   = $ct.privacyUrl }
                if ($ct.license)      { $txtLicense.Text      = $ct.license }
                if ($ct.licenseUrl)   { $txtLicenseUrl.Text   = $ct.licenseUrl }
                if ($ct.author)       { $txtAuthor.Text       = $ct.author }
                $yearStr = (Get-Date -Format yyyy)
                if ($ct.copyright) {
                    $txtCopyright.Text = $ct.copyright -replace '\{YEAR\}', $yearStr
                }
                Show-Toast "Applied custom template: $($ct.name)" -Type Success
            } catch {
                Show-Toast "Error reading metadata_templates.json: $($_.Exception.Message)" -Type Error
            }
            Update-YamlPreview
            return
        }
    }

    Update-YamlPreview
    $lblStatus.Text = "Template applied: $Tmpl"
})

# Show/hide nested installer section based on InstallerType
$cmbInstallerType0.Add_SelectionChanged({
    $CurrentType = Get-ComboBoxSelectedText $cmbInstallerType0
    if ($CurrentType -eq 'zip') {
        $pnlNestedInstaller.Visibility = 'Visible'
    } else {
        $pnlNestedInstaller.Visibility = 'Collapsed'
    }

    # Auto-fill known silent switches when switch fields are still empty
    $SwitchesEmpty = (-not $txtSilent0.Text.Trim()) -and (-not $txtSilentProgress0.Text.Trim()) -and
                     (-not $txtCustomSwitch0.Text.Trim()) -and (-not $txtInteractive0.Text.Trim())
    if ($SwitchesEmpty) {
        switch ($CurrentType) {
            'msi' {
                $txtSilent0.Text = '/qn /norestart'
                $txtSilentProgress0.Text = '/qb /norestart'
                $txtInteractive0.Text = '/qf'
                $txtLog0.Text = '/l*v "%TEMP%\MSIInstall.log"'
                $txtRepair0.Text = '/f'
            }
            'inno' {
                $txtSilent0.Text = '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-'
                $txtSilentProgress0.Text = '/SILENT /SUPPRESSMSGBOXES /NORESTART /SP-'
                $txtLog0.Text = '/LOG="%TEMP%\InnoInstall.log"'
            }
            'nullsoft' {
                $txtSilent0.Text = '/S'
                $txtSilentProgress0.Text = '/S'
            }
            'burn' {
                $txtSilent0.Text = '/quiet /norestart'
                $txtSilentProgress0.Text = '/passive /norestart'
                $txtLog0.Text = '/log "%TEMP%\BurnInstall.log"'
                $txtRepair0.Text = '/repair'
            }
            'wix' {
                $txtSilent0.Text = '/qn /norestart'
                $txtSilentProgress0.Text = '/qb /norestart'
                $txtLog0.Text = '/l*v "%TEMP%\WixInstall.log"'
                $txtRepair0.Text = '/f'
            }
        }
    }

    Update-YamlPreview
})

# YAML preview auto-update on field changes
$PreviewFields = @($txtPkgId, $txtPkgVersion, $txtPkgName, $txtPublisher, $txtLicense,
                    $txtShortDesc, $txtDescription, $txtAuthor, $txtMoniker,
                    $txtPackageUrl, $txtPublisherUrl, $txtLicenseUrl, $txtSupportUrl,
                    $txtCopyright, $txtPrivacyUrl, $txtTags, $txtReleaseNotes, $txtReleaseNotesUrl,
                    $txtInstallerUrl0, $txtSilent0, $txtSilentProgress0,
                    $txtProductCode0, $txtCommands0, $txtFileExt0,
                    $txtCustomSwitch0, $txtInteractive0, $txtLog0, $txtRepair0,
                    $txtNestedPath0, $txtPortableAlias0, $txtMinOSVersion0, $txtExpectedReturnCodes0,
                    $txtAFDisplayName0, $txtAFPublisher0, $txtAFDisplayVersion0, $txtAFProductCode0)
foreach ($field in $PreviewFields) {
    $field.Add_TextChanged({ Update-YamlPreview })
}
$PreviewCombos = @($cmbInstallerType0, $cmbArch0, $cmbScope0, $cmbUpgrade0, $cmbElevation0, $cmbCreateManifestVer, $cmbNestedType0)
foreach ($combo in $PreviewCombos) {
    $combo.Add_SelectionChanged({ Update-YamlPreview })
}
# Checkboxes also update preview
$chkModeInteractive.Add_Checked({ Update-YamlPreview })
$chkModeInteractive.Add_Unchecked({ Update-YamlPreview })
$chkModeSilent.Add_Checked({ Update-YamlPreview })
$chkModeSilent.Add_Unchecked({ Update-YamlPreview })
$chkModeSilentProgress.Add_Checked({ Update-YamlPreview })
$chkModeSilentProgress.Add_Unchecked({ Update-YamlPreview })
$chkInstallLocationRequired.Add_Checked({ Update-YamlPreview })
$chkInstallLocationRequired.Add_Unchecked({ Update-YamlPreview })
$chkPlatformDesktop.Add_Checked({ Update-YamlPreview })
$chkPlatformDesktop.Add_Unchecked({ Update-YamlPreview })
$chkPlatformUniversal.Add_Checked({ Update-YamlPreview })
$chkPlatformUniversal.Add_Unchecked({ Update-YamlPreview })

# ── Edit Tab Actions ─────────────────────────────────────────────────────────
$btnEditSave.Add_Click({
    $YamlText = $txtEditYaml.Text
    if (-not $YamlText) {
        Show-ThemedDialog -Title 'Empty' -Message 'No YAML content to save.' `
            -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' | Out-Null
        return
    }
    $Dialog = New-Object System.Windows.Forms.SaveFileDialog
    $Dialog.Title = "Save Manifest YAML"
    $Dialog.Filter = "YAML Files|*.yaml|All Files|*.*"
    $Dialog.DefaultExt = ".yaml"
    if ($Dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $YamlText | Set-Content $Dialog.FileName -Force -Encoding UTF8
        Write-DebugLog "Edited YAML saved to: $($Dialog.FileName)"
        $lblStatus.Text = "YAML saved: $($Dialog.FileName)"
        $statusDot.Fill = $Window.Resources['ThemeSuccess']
    }
})

$btnEditPublish.Add_Click({
    $YamlText = $txtEditYaml.Text
    if (-not $YamlText) {
        Show-ThemedDialog -Title 'Empty' -Message 'No YAML content to publish.' `
            -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' | Out-Null
        return
    }
    $FuncName = Get-ActiveRestSource
    if (-not $FuncName) {
        Show-ThemedDialog -Title 'Missing Setting' -Message 'REST Source Function Name is required. Set it in Settings.' `
            -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' | Out-Null
        return
    }
    # Save YAML to temp directory, then publish
    $TempDir = Join-Path $env:TEMP "WinGetMMEdit_$(Get-Date -Format 'yyyyMMddHHmmss')"
    New-Item -Path $TempDir -ItemType Directory -Force | Out-Null
    $CurrentFile = ""; $CurrentContent = ""
    foreach ($Line in ($YamlText -split "`n")) {
        if ($Line -match '^# ===== (.+\.yaml)') {
            if ($CurrentFile -and $CurrentContent.Trim()) {
                $CurrentContent | Set-Content (Join-Path $TempDir $CurrentFile) -Force -Encoding UTF8
            }
            $CurrentFile = $Matches[1].Trim()
            $CurrentContent = ""
        } else {
            $CurrentContent += "$Line`n"
        }
    }
    if ($CurrentFile -and $CurrentContent.Trim()) {
        $CurrentContent | Set-Content (Join-Path $TempDir $CurrentFile) -Force -Encoding UTF8
    }
    if (-not (Get-ChildItem $TempDir -Filter "*.yaml" -ErrorAction SilentlyContinue)) {
        # No split sections found — save as single file
        $YamlText | Set-Content (Join-Path $TempDir "manifest.yaml") -Force -Encoding UTF8
    }
    Publish-ManifestToRestSource -ManifestPath $TempDir -FunctionName $FuncName
})

$btnLoadFromDisk.Add_Click({
    $Dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $Dialog.Description = "Select manifest folder (containing .yaml files)"
    if ($Dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $YamlFiles = Get-ChildItem $Dialog.SelectedPath -Filter "*.yaml" -ErrorAction SilentlyContinue
        if ($YamlFiles.Count -eq 0) {
            Show-ThemedDialog -Title 'No Manifests' -Message 'No .yaml files found in selected folder.' `
                -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' | Out-Null
            return
        }
        $AllYaml = ""
        foreach ($F in $YamlFiles) {
            $AllYaml += "# ===== $($F.Name) =====`n"
            $AllYaml += (Get-Content $F.FullName -Raw) + "`n`n"
        }
        $txtEditYaml.Text = $AllYaml
        $lblEditPkgId.Text = Split-Path $Dialog.SelectedPath -Leaf
        $lblEditPkgVersion.Text = ""
        $pnlEditFields.Visibility = 'Visible'
        Write-DebugLog "Loaded manifests from disk: $($Dialog.SelectedPath)"
    }
})

$btnLoadFromConfig.Add_Click({
    $Sel = $lstConfigs.SelectedItem
    if ($Sel -and $Sel.Tag) {
        Load-ManifestFromConfig -ConfigPath $Sel.Tag
        Populate-CreateTabFromConfig -ConfigPath $Sel.Tag
        Switch-Tab 'Create'
    } else {
        Show-ThemedDialog -Title 'No Selection' -Message 'Select a package config from the left panel first.' `
            -Icon ([string]([char]0xE946)) -IconColor 'ThemeAccentLight' | Out-Null
    }
})

# Config list double-click to load into Create tab
$lstConfigs.Add_MouseDoubleClick({
    $Sel = $lstConfigs.SelectedItem
    if ($Sel -and $Sel.Tag) {
        Load-ManifestFromConfig -ConfigPath $Sel.Tag
        Populate-CreateTabFromConfig -ConfigPath $Sel.Tag
        Switch-Tab 'Create'
    }
})

# ── Upload Tab Actions ───────────────────────────────────────────────────────
$btnUpload.Add_Click({
    $FilePath = $txtUploadFile.Text.Trim()
    if (-not $FilePath -or -not (Test-Path $FilePath)) {
        Show-ThemedDialog -Title 'No File' -Message 'Select a file to upload first.' `
            -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' | Out-Null
        return
    }
    $StorageSel = $cmbUploadStorage.SelectedItem
    if (-not $StorageSel) {
        Show-ThemedDialog -Title 'No Storage' -Message 'Select a storage account.' `
            -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' | Out-Null
        return
    }
    $ContainerSel = $cmbUploadContainer.SelectedItem
    $ContainerName = if ($ContainerSel) { $ContainerSel.Content.ToString() } else { "packages" }
    $BlobPath = $txtBlobPath.Text.Trim()
    if (-not $BlobPath) {
        $FileName = Split-Path $FilePath -Leaf
        $PkgId = $txtPkgId.Text.Trim()
        $PkgVer = $txtPkgVersion.Text.Trim()
        if ($PkgId -and $PkgVer) {
            $BlobPath = "$PkgId/$PkgVer/$FileName"
        } else {
            $BlobPath = $FileName
        }
        $txtBlobPath.Text = $BlobPath
    }

    $AcctName = $StorageSel.Tag.Name
    $AcctInfo = $StorageSel.Tag

    # Confirm upload if setting is enabled
    if ($chkConfirmUpload.IsChecked) {
        $Confirm = Show-ThemedDialog -Title 'Confirm Upload' `
            -Message "Upload file to Azure Storage?`n`nStorage: $AcctName`nContainer: $ContainerName`nBlob: $BlobPath" `
            -Icon ([string]([char]0xE898)) -IconColor 'ThemeAccentLight' `
            -Buttons @(
                @{ Text='Cancel'; IsAccent=$false; Result='Cancel' }
                @{ Text='Upload'; IsAccent=$true; Result='Upload' }
            )
        if ($Confirm -ne 'Upload') { return }
    }

    $FileSize = (Get-Item $FilePath).Length
    $FileSizeMB = [math]::Round($FileSize / 1MB, 1)
    Write-DebugLog "Uploading $FilePath ($FileSizeMB MB) to $AcctName/$ContainerName/$BlobPath"
    $lblStatus.Text = "Uploading... 0 / $FileSizeMB MB"
    $lblUploadProgress.Text = "0 / $FileSizeMB MB"
    $prgUpload.Visibility = 'Visible'
    $prgUpload.IsIndeterminate = $false
    $prgUpload.Maximum = 100
    $prgUpload.Value = 0
    $Global:SyncHash.UploadProgress = 0
    $Global:SyncHash.UploadBytes = 0
    $Global:SyncHash.UploadTotal = $FileSize
    Show-GlobalProgressIndeterminate -Text "Uploading 0 / $FileSizeMB MB..."
    Set-ButtonBusy -Button $btnUpload -Text 'Uploading...'

    Start-BackgroundWork -Variables @{
        FPath   = $FilePath
        AccName = $AcctName
        AccRG   = $AcctInfo.ResourceGroup
        CName   = $ContainerName
        BPath   = $BlobPath
        SyncH   = $Global:SyncHash
    } -Work {
        try {
            $SyncH.UploadProgress = 0
            $SyncH.UploadBytes = 0
            $SyncH.UploadTotal = (Get-Item $FPath).Length
            # Use OAuth context (Entra ID) instead of storage keys
            $Ctx = New-AzStorageContext -StorageAccountName $AccName -UseConnectedAccount -ErrorAction Stop
            # Ensure container exists
            try { New-AzStorageContainer -Name $CName -Context $Ctx -Permission Off -ErrorAction Stop | Out-Null } catch { <# Container already exists #> }
            # Upload with progress tracking via block upload
            $TotalSize = $SyncH.UploadTotal
            $BlockSize = 4MB
            if ($TotalSize -le $BlockSize) {
                # Small file — single upload, no chunking needed
                Set-AzStorageBlobContent -File $FPath -Container $CName -Blob $BPath -Context $Ctx -Force -ErrorAction Stop | Out-Null
                $SyncH.UploadBytes = $TotalSize
                $SyncH.UploadProgress = 100
            } else {
                # Large file — upload in blocks and report progress
                $BlobClient = Get-AzStorageBlobContent -Container $CName -Blob $BPath -Context $Ctx -ErrorAction SilentlyContinue
                # Get the CloudBlobContainer for the block blob API
                $CloudContainer = $Ctx.StorageAccount.CreateCloudBlobClient().GetContainerReference($CName)
                $CloudBlob = $CloudContainer.GetBlockBlobReference($BPath)
                $Stream = [System.IO.File]::OpenRead($FPath)
                try {
                    $Buffer = New-Object byte[] $BlockSize
                    $BlockIds = [System.Collections.ArrayList]::new()
                    $BytesSent = 0
                    $BlockNum = 0
                    while (($BytesRead = $Stream.Read($Buffer, 0, $BlockSize)) -gt 0) {
                        $BlockId = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("block-{0:D6}" -f $BlockNum))
                        $BlockStream = New-Object System.IO.MemoryStream($Buffer, 0, $BytesRead)
                        $CloudBlob.PutBlock($BlockId, $BlockStream, $null)
                        $BlockStream.Dispose()
                        [void]$BlockIds.Add($BlockId)
                        $BytesSent += $BytesRead
                        $BlockNum++
                        $SyncH.UploadBytes = $BytesSent
                        $SyncH.UploadProgress = [math]::Floor(($BytesSent / $TotalSize) * 100)
                    }
                    $CloudBlob.PutBlockList($BlockIds)
                    $SyncH.UploadProgress = 100
                } finally {
                    $Stream.Dispose()
                }
            }
            $BlobUrl = "https://$AccName.blob.core.windows.net/$CName/$BPath"
            return @{ Success = $true; BlobUrl = $BlobUrl; Error = $null }
        } catch {
            return @{ Success = $false; BlobUrl = $null; Error = $_.Exception.Message }
        }
    } -OnComplete {
        param($Results, $Errors)
        $prgUpload.Visibility = 'Collapsed'
        $prgUpload.Value = 0
        $Global:SyncHash.UploadProgress = -1
        $Global:SyncHash.UploadBytes = 0
        $Global:SyncHash.UploadTotal = 0
        Hide-GlobalProgress
        Reset-ButtonBusy -Button $btnUpload
        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ Success = $false; Error = 'No result' } }
        if ($R.Success) {
            Write-DebugLog "Upload complete: $($R.BlobUrl)"
            $lblUploadProgress.Text = "Uploaded: $($R.BlobUrl)"
            $lblStatus.Text = "Upload complete"
            $statusDot.Fill = $Window.Resources['ThemeSuccess']
            Show-Toast "Upload complete — URL copied to clipboard" -Type Success
            Unlock-Achievement 'first_upload'
            # Auto-fill the Installer URL on the Create tab
            $txtInstallerUrl0.Text = $R.BlobUrl
            [System.Windows.Clipboard]::SetText($R.BlobUrl)
            Write-DebugLog "Blob URL copied to clipboard and set as Installer URL"
            # Auto-save config so the blob URL is persisted
            $PkgId = $txtPkgId.Text.Trim()
            if ($PkgId) {
                Save-PackageConfig -Force
                Write-DebugLog "Config auto-saved with blob URL"
            }
        } else {
            Write-DebugLog "Upload failed: $($R.Error)" -Level 'ERROR'
            $lblUploadProgress.Text = "Upload failed: $($R.Error)"
            $lblStatus.Text = "Upload failed"
            $statusDot.Fill = $Window.Resources['ThemeError']
            Show-Toast "Upload failed: $($R.Error)" -Type Error -DurationMs 6000
        }
    }
})

$btnCopyHash.Add_Click({
    $Hash = $lblUploadHash.Text
    if ($Hash -and $Hash -ne "Select a file to compute hash...") {
        [System.Windows.Clipboard]::SetText($Hash)
        $lblStatus.Text = "Hash copied to clipboard"
    }
})

# ── Batch Installer Operations ────────────────────────────────────────────────
function Update-BatchPanel {
    <# Shows/hides batch panel on Upload tab + multi-installer hint on Create tab based on installer count #>
    $AllInstallers = Get-AllInstallerData
    # Only count installers that have a URL — blank/empty entries don't count for batch ops
    $ValidInstallers = @($AllInstallers | Where-Object { $_.InstallerUrl -and $_.InstallerUrl -match '^https?://' })
    $HasMulti = ($ValidInstallers.Count -gt 1)
    if ($pnlBatchOps)  { $pnlBatchOps.Visibility  = if ($HasMulti) { 'Visible' } else { 'Collapsed' } }
    if ($lblMultiInstallerHint) { $lblMultiInstallerHint.Visibility = if ($HasMulti) { 'Visible' } else { 'Collapsed' } }
    if ($HasMulti -and $lblBatchCount) {
        $lblBatchCount.Text = "$($ValidInstallers.Count) installers with URLs — download and upload all at once"
        # Refresh the list
        $lstBatchInstallers.Items.Clear()
        $idx = 0
        foreach ($inst in $AllInstallers) {
            $idx++
            $Arch = $inst.Architecture; $Type = $inst.InstallerType; $Scope = $inst.Scope
            $Url = $inst.InstallerUrl
            $FileName = if ($Url) { [System.IO.Path]::GetFileName(([uri]$Url).LocalPath) } else { '(no URL)' }
            $Desc = "#$idx  $Arch / $Type"
            if ($Scope) { $Desc += " / $Scope" }
            $Desc += " — $FileName"
            $DlFile = $Global:BatchDownloadedFiles["inst_$($idx-1)"]
            $StatusText = if ($DlFile -and (Test-Path $DlFile)) { "Downloaded: $([System.IO.Path]::GetFileName($DlFile))" }
                          elseif ($Url -and $Url -match '^https?://') { "Ready to download" }
                          else { "No URL" }
            $Item = New-StyledListViewItem -IconChar ([char]0xE7B8) -PrimaryText $Desc -SubtitleText $StatusText -TagValue @{ Index = ($idx-1); Url = $Url }
            $lstBatchInstallers.Items.Add($Item) | Out-Null
        }
    }
}

$btnDownloadAll.Add_Click({
    $AllInstallers = Get-AllInstallerData
    if ($AllInstallers.Count -lt 2) {
        Show-Toast "Only 1 installer — use the Download button on the Create tab" -Type Info
        return
    }
    # Collect URLs
    $UrlList = @()
    for ($i = 0; $i -lt $AllInstallers.Count; $i++) {
        $Url = $AllInstallers[$i].InstallerUrl
        if (-not $Url -or $Url -notmatch '^https?://') {
            Show-Toast "Installer #$($i+1) has no valid URL" -Type Warning
            return
        }
        $UrlList += @{ Index = $i; Url = $Url }
    }
    $Global:BatchCancelled = $false
    $Global:BatchDownloadedFiles = @{}
    $btnDownloadAll.IsEnabled = $false
    $btnUploadAll.IsEnabled = $false
    $btnCancelBatch.Visibility = 'Visible'
    $prgBatch.Visibility = 'Visible'
    $prgBatch.IsIndeterminate = $false
    $prgBatch.Maximum = $UrlList.Count
    $prgBatch.Value = 0
    $lblBatchProgress.Text = "Downloading 0/$($UrlList.Count)..."
    $lblStatus.Text = "Batch download starting..."

    $Global:SyncHash.BatchTotal = $UrlList.Count
    $Global:SyncHash.BatchDone  = 0
    $Global:SyncHash.BatchCurrent = ''

    Start-BackgroundWork -Variables @{
        Urls   = $UrlList
        SyncH  = $Global:SyncHash
    } -Work {
        $TempDir = Join-Path $env:TEMP 'WinGetMM_Downloads'
        if (-not (Test-Path $TempDir)) { New-Item -ItemType Directory -Path $TempDir -Force | Out-Null }
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        $Results = @{}
        foreach ($Entry in $Urls) {
            $Idx = $Entry.Index
            $Url = $Entry.Url
            $FileName = [System.IO.Path]::GetFileName(([uri]$Url).LocalPath)
            if (-not $FileName -or $FileName.Length -lt 2) { $FileName = "installer_$Idx" }
            $TempFile = Join-Path $TempDir $FileName
            $SyncH.BatchCurrent = $FileName
            $Request  = $null; $Response = $null; $Stream = $null; $FileOut = $null
            try {
                $Request = [System.Net.HttpWebRequest]::Create([uri]$Url)
                $Request.Timeout = 120000
                $Response = $Request.GetResponse()
                $TotalBytes = $Response.ContentLength
                $Stream  = $Response.GetResponseStream()
                $FileOut = [System.IO.File]::Create($TempFile)
                $Buffer  = New-Object byte[] 65536
                $Downloaded = [long]0
                while (($BytesRead = $Stream.Read($Buffer, 0, $Buffer.Length)) -gt 0) {
                    $FileOut.Write($Buffer, 0, $BytesRead)
                    $Downloaded += $BytesRead
                }
                $FileOut.Close(); $FileOut = $null
                $Hash = (Get-FileHash -Path $TempFile -Algorithm SHA256).Hash
                $Results["inst_$Idx"] = @{ Path = $TempFile; Hash = $Hash; Size = $Downloaded; Name = $FileName; Index = $Idx }
            } catch {
                $Results["inst_$Idx"] = @{ Error = $_.Exception.Message; Index = $Idx }
            } finally {
                if ($FileOut)  { try { $FileOut.Dispose()  } catch {} }
                if ($Stream)   { try { $Stream.Dispose()   } catch {} }
                if ($Response) { try { $Response.Dispose()  } catch {} }
            }
            $SyncH.BatchDone = $SyncH.BatchDone + 1
        }
        return $Results
    } -OnComplete {
        param($Results, $Errors)
        $prgBatch.Visibility = 'Collapsed'
        $btnCancelBatch.Visibility = 'Collapsed'
        $btnDownloadAll.IsEnabled = $true
        $Global:SyncHash.BatchTotal = 0
        if ($Global:BatchCancelled) {
            $lblBatchProgress.Text = "Download cancelled"
            $lblStatus.Text = "Batch download cancelled"
            return
        }
        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{} }
        $AllInstallers = Get-AllInstallerData
        $SuccessCount = 0; $FailCount = 0
        foreach ($Key in $R.Keys) {
            $Entry = $R[$Key]
            if ($Entry.Path -and (Test-Path $Entry.Path)) {
                $Global:BatchDownloadedFiles[$Key] = $Entry.Path
                $Idx = $Entry.Index
                # Update hash on installer entry
                if ($Idx -eq 0) {
                    $Global:CurrentHash0 = $Entry.Hash
                    $lblHash0.Text = "SHA256: $($Entry.Hash)"
                } elseif ($Idx -gt 0 -and $Idx -le $Global:InstallerEntries.Count) {
                    $Global:InstallerEntries[$Idx - 1].InstallerSha256 = $Entry.Hash
                }
                $SuccessCount++
            } elseif ($Entry.Error) {
                Write-DebugLog "Batch download #$($Entry.Index+1) failed: $($Entry.Error)" -Level 'ERROR'
                $FailCount++
            }
        }
        Update-InstallerSummaryList
        Update-BatchPanel
        $btnUploadAll.IsEnabled = ($SuccessCount -gt 0)
        if ($FailCount -eq 0) {
            $lblBatchProgress.Text = "All $SuccessCount installers downloaded successfully"
            $lblStatus.Text = "Downloaded $SuccessCount installers — ready to upload"
            Show-Toast "Downloaded $SuccessCount installers" -Type Success -DurationMs 4000
            Unlock-Achievement 'batch_download'
        } else {
            $lblBatchProgress.Text = "$SuccessCount downloaded, $FailCount failed"
            $lblStatus.Text = "Batch download: $SuccessCount OK, $FailCount failed"
            Show-Toast "$FailCount of $($SuccessCount+$FailCount) downloads failed — check log" -Type Warning
        }
        Update-YamlPreview
    }
})

$btnUploadAll.Add_Click({
    if ($Global:BatchDownloadedFiles.Count -eq 0) {
        Show-Toast "Download installers first" -Type Warning
        return
    }
    $StorageSel = $cmbUploadStorage.SelectedItem
    if (-not $StorageSel) {
        Show-ThemedDialog -Title 'No Storage' -Message 'Select a storage account on this tab first.' `
            -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' | Out-Null
        return
    }
    $ContainerSel = $cmbUploadContainer.SelectedItem
    $ContainerName = if ($ContainerSel) { $ContainerSel.Content.ToString() } else { 'packages' }
    $PkgId = $txtPkgId.Text.Trim()
    $PkgVer = $txtPkgVersion.Text.Trim()
    if (-not $PkgId -or -not $PkgVer) {
        Show-ThemedDialog -Title 'Missing Info' -Message 'Package Identifier and Version are required on the Create tab.' `
            -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' | Out-Null
        return
    }
    $AcctName = $StorageSel.Tag.Name
    $AcctInfo = $StorageSel.Tag
    $FileList = @()
    foreach ($Key in $Global:BatchDownloadedFiles.Keys) {
        $FilePath = $Global:BatchDownloadedFiles[$Key]
        if (Test-Path $FilePath) {
            $FileName = Split-Path $FilePath -Leaf
            $BlobPath = "$PkgId/$PkgVer/$FileName"
            $FileList += @{ Key = $Key; Path = $FilePath; BlobPath = $BlobPath; Name = $FileName }
        }
    }
    if ($FileList.Count -eq 0) {
        Show-Toast "No downloaded files found" -Type Warning
        return
    }
    # Confirm
    if ($chkConfirmUpload.IsChecked) {
        $Confirm = Show-ThemedDialog -Title 'Confirm Batch Upload' `
            -Message "Upload $($FileList.Count) installer(s) to Azure Storage?`n`nStorage: $AcctName`nContainer: $ContainerName`nPath: $PkgId/$PkgVer/" `
            -Icon ([string]([char]0xE898)) -IconColor 'ThemeAccentLight' `
            -Buttons @(
                @{ Text='Cancel'; IsAccent=$false; Result='Cancel' }
                @{ Text='Upload All'; IsAccent=$true; Result='Upload' }
            )
        if ($Confirm -ne 'Upload') { return }
    }
    $Global:BatchCancelled = $false
    $btnUploadAll.IsEnabled = $false
    $btnDownloadAll.IsEnabled = $false
    $btnCancelBatch.Visibility = 'Visible'
    $prgBatch.Visibility = 'Visible'
    $prgBatch.IsIndeterminate = $false
    $prgBatch.Maximum = $FileList.Count
    $prgBatch.Value = 0
    $lblBatchProgress.Text = "Uploading 0/$($FileList.Count)..."
    $lblStatus.Text = "Batch upload starting..."

    Start-BackgroundWork -Variables @{
        Files     = $FileList
        AccName   = $AcctName
        AccRG     = $AcctInfo.ResourceGroup
        CName     = $ContainerName
        SyncH     = $Global:SyncHash
    } -Work {
        $SyncH.BatchTotal = $Files.Count
        $SyncH.BatchDone  = 0
        $Ctx = New-AzStorageContext -StorageAccountName $AccName -UseConnectedAccount -ErrorAction Stop
        try { New-AzStorageContainer -Name $CName -Context $Ctx -Permission Off -ErrorAction Stop | Out-Null } catch { <# exists #> }
        $Results = @{}
        foreach ($F in $Files) {
            $SyncH.BatchCurrent = $F.Name
            try {
                Set-AzStorageBlobContent -File $F.Path -Container $CName -Blob $F.BlobPath -Context $Ctx -Force -ErrorAction Stop | Out-Null
                $BlobUrl = "https://$AccName.blob.core.windows.net/$CName/$($F.BlobPath)"
                $Results[$F.Key] = @{ Success = $true; BlobUrl = $BlobUrl; Key = $F.Key }
            } catch {
                $Results[$F.Key] = @{ Success = $false; Error = $_.Exception.Message; Key = $F.Key }
            }
            $SyncH.BatchDone = $SyncH.BatchDone + 1
        }
        return $Results
    } -OnComplete {
        param($Results, $Errors)
        $prgBatch.Visibility = 'Collapsed'
        $btnCancelBatch.Visibility = 'Collapsed'
        $btnDownloadAll.IsEnabled = $true
        $btnUploadAll.IsEnabled = $true
        $Global:SyncHash.BatchTotal = 0
        if ($Global:BatchCancelled) {
            $lblBatchProgress.Text = "Upload cancelled"
            $lblStatus.Text = "Batch upload cancelled"
            return
        }
        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{} }
        $SuccessCount = 0; $FailCount = 0
        foreach ($Key in $R.Keys) {
            $Entry = $R[$Key]
            if ($Entry.Success) {
                # Update installer URL to blob URL
                $IdxStr = $Key -replace '^inst_', ''
                $Idx = [int]$IdxStr
                if ($Idx -eq 0) {
                    $txtInstallerUrl0.Text = $Entry.BlobUrl
                } elseif ($Idx -gt 0 -and $Idx -le $Global:InstallerEntries.Count) {
                    $Global:InstallerEntries[$Idx - 1].InstallerUrl = $Entry.BlobUrl
                }
                $SuccessCount++
            } else {
                Write-DebugLog "Batch upload $Key failed: $($Entry.Error)" -Level 'ERROR'
                $FailCount++
            }
        }
        Update-InstallerSummaryList
        Update-BatchPanel
        Update-YamlPreview
        if ($FailCount -eq 0) {
            $lblBatchProgress.Text = "All $SuccessCount installers uploaded successfully"
            $lblStatus.Text = "Uploaded $SuccessCount installers to blob storage"
            Show-Toast "Uploaded $SuccessCount installer(s) to blob storage" -Type Success -DurationMs 5000
            Unlock-Achievement 'batch_upload'
            Save-PackageConfig -Force
        } else {
            $lblBatchProgress.Text = "$SuccessCount uploaded, $FailCount failed"
            $lblStatus.Text = "Batch upload: $SuccessCount OK, $FailCount failed"
            Show-Toast "$FailCount of $($SuccessCount+$FailCount) uploads failed" -Type Warning
        }
    }
})

$btnCancelBatch.Add_Click({
    $Global:BatchCancelled = $true
    $Job = $Global:BgJobs | Select-Object -Last 1
    if ($Job) {
        try { $Job.PS.Stop() } catch { <# already stopped #> }
        try { $Job.PS.Dispose() } catch { <# already disposed #> }
        try { $Job.Runspace.Dispose() } catch { <# already disposed #> }
        $Global:BgJobs.Remove($Job)
    }
    $prgBatch.Visibility = 'Collapsed'
    $btnCancelBatch.Visibility = 'Collapsed'
    $btnDownloadAll.IsEnabled = $true
    $btnUploadAll.IsEnabled = ($Global:BatchDownloadedFiles.Count -gt 0)
    $Global:SyncHash.BatchTotal = 0
    $lblBatchProgress.Text = "Operation cancelled"
    $lblStatus.Text = "Batch operation cancelled"
    Write-DebugLog "Batch operation cancelled by user" -Level 'WARN'
})

# ── REST Source Health Check ──────────────────────────────────────────────────
$btnTestRestSource = $Window.FindName("btnTestRestSource")
$btnTestRestSource.Add_Click({ Test-RestSourceConnection })

# ── Connect Command (winget source add) ──────────────────────────────────────
$btnConnectCmd = $Window.FindName("btnConnectCmd")
if ($btnConnectCmd) { $btnConnectCmd.Add_Click({ Show-RestSourceConnectCommand }) }

# ── Saved REST Sources ───────────────────────────────────────────────────────
$cmbSavedSources   = $Window.FindName("cmbSavedSources")
$btnSaveSource     = $Window.FindName("btnSaveSource")
$btnDeleteSource   = $Window.FindName("btnDeleteSource")

$cmbSavedSources.Add_SelectionChanged({
    $Sel = $cmbSavedSources.SelectedItem
    if ($Sel -and $Sel.Tag) {
        $txtRestSourceName.Text = $Sel.Tag
        # Sync to global combo
        for ($i = 0; $i -lt $cmbRestSource.Items.Count; $i++) {
            if ($cmbRestSource.Items[$i].Tag -eq $Sel.Tag) {
                $cmbRestSource.SelectedIndex = $i
                break
            }
        }
    }
})

$btnSaveSource.Add_Click({
    $Name = $txtRestSourceName.Text.Trim()
    if (-not $Name) {
        $lblStatus.Text = "Enter a source name first"
        return
    }
    # Check for duplicates
    $Existing = $cmbSavedSources.Items | Where-Object { $_.Tag -eq $Name }
    if ($Existing) {
        $lblStatus.Text = "Source '$Name' already saved"
        return
    }
    $Item = New-Object System.Windows.Controls.ComboBoxItem
    $Item.Content = $Name
    $Item.Tag = $Name
    $cmbSavedSources.Items.Add($Item) | Out-Null
    $cmbSavedSources.SelectedItem = $Item
    Save-UserPrefs
    $lblStatus.Text = "Source '$Name' saved"
})

$btnDeleteSource.Add_Click({
    $Sel = $cmbSavedSources.SelectedItem
    if (-not $Sel) {
        $lblStatus.Text = "Select a saved source to delete"
        return
    }
    $Name = $Sel.Tag
    $Confirm = Show-ThemedDialog -Title 'Remove Saved Source' `
        -Message "Remove '$Name' from saved sources?`n`nThis only removes it from the list — it does not delete Azure resources." `
        -Icon ([string]([char]0xE74D)) -IconColor 'ThemeWarning' `
        -Buttons @(
            @{ Text='Cancel'; IsAccent=$false; Result='Cancel' }
            @{ Text='Remove'; IsDanger=$true;  Result='Remove' }
        )
    if ($Confirm -ne 'Remove') { return }
    $cmbSavedSources.Items.Remove($Sel)
    if ($cmbSavedSources.Items.Count -gt 0) {
        $cmbSavedSources.SelectedIndex = 0
    }
    Save-UserPrefs
    $lblStatus.Text = "Source '$Name' removed"
})

# ── Global REST Source Selector (top bar) ────────────────────────────────────
$cmbRestSource.Add_SelectionChanged({
    $Sel = $cmbRestSource.SelectedItem
    if ($Sel -and $Sel.Tag) {
        # Sync back to settings textbox
        $txtRestSourceName.Text = $Sel.Tag
        # Update Manage tab label if visible
        $lblManageSourceName.Text = "REST Source: $($Sel.Tag)"
        # Update backup/restore combos
        $cmbBackupRestSource.Text = $Sel.Tag
        Populate-BackupRestoreCombos
    }
})

# ── Manage Tab Actions ───────────────────────────────────────────────────────
$btnRefreshManage.Add_Click({ Refresh-ManagePackageList -ForceRefresh })

$lstManagePackages.Add_SelectionChanged({
    $Sel = $lstManagePackages.SelectedItem
    if (-not $Sel -or -not $Sel.Tag) {
        $pnlManageDetails.Visibility = 'Collapsed'
        return
    }
    $Pkg = $Sel.Tag
    $lblManagePkgId.Text = $Pkg.Id
    $lblManageVersions.Text = "Published versions: $(($Pkg.Versions -join ', '))"
    $cmbManageVersion.Items.Clear()
    foreach ($V in $Pkg.Versions) {
        $cmbManageVersion.Items.Add($V) | Out-Null
    }
    if ($cmbManageVersion.Items.Count -gt 0) {
        $cmbManageVersion.SelectedIndex = 0
    }
    $chkDeleteBlob.IsChecked = $false
    $chkDeleteConfig.IsChecked = $false
    $pnlManageDetails.Visibility = 'Visible'
})

$btnRemoveVersion.Add_Click({
    $Sel = $lstManagePackages.SelectedItem
    if (-not $Sel -or -not $Sel.Tag) { return }
    $PkgId = $Sel.Tag.Id
    $SelVer = $cmbManageVersion.SelectedItem
    if (-not $SelVer) {
        Show-ThemedDialog -Title 'No Version' -Message 'Select a version to remove.' `
            -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' | Out-Null
        return
    }
    $Version = $SelVer.ToString()
    $Confirm = Show-ThemedDialog -Title 'Confirm Removal' `
        -Message "Remove version $Version of $PkgId from the REST source?`n`nThis action cannot be undone." `
        -Icon ([string]([char]0xE74D)) -IconColor 'ThemeError' `
        -Buttons @(
            @{ Text='Cancel'; IsAccent=$false; Result='Cancel' }
            @{ Text='Remove'; IsDanger=$true; Result='Remove' }
        )
    if ($Confirm -ne 'Remove') { return }
    Remove-PackageFromRestSource -PackageId $PkgId -Version $Version `
        -DeleteBlob $chkDeleteBlob.IsChecked -DeleteConfig $chkDeleteConfig.IsChecked
})

$btnRemovePackage.Add_Click({
    $Sel = $lstManagePackages.SelectedItem
    if (-not $Sel -or -not $Sel.Tag) { return }
    $PkgId = $Sel.Tag.Id
    $VerCount = $Sel.Tag.Versions.Count
    $Confirm = Show-ThemedDialog -Title 'Confirm Full Removal' `
        -Message "Remove ALL $VerCount version(s) of $PkgId from the REST source?`n`nThis will permanently delete the entire package. This action cannot be undone." `
        -Icon ([string]([char]0xE711)) -IconColor 'ThemeError' `
        -Buttons @(
            @{ Text='Cancel'; IsAccent=$false; Result='Cancel' }
            @{ Text='Remove All'; IsDanger=$true; Result='Remove' }
        )
    if ($Confirm -ne 'Remove') { return }
    Remove-PackageFromRestSource -PackageId $PkgId -Version '' `
        -DeleteBlob $chkDeleteBlob.IsChecked -DeleteConfig $chkDeleteConfig.IsChecked
})

# Storage account selection → list containers
$cmbUploadStorage.Add_SelectionChanged({
    $Sel = $cmbUploadStorage.SelectedItem
    if (-not $Sel -or -not $Sel.Tag) { return }
    $AcctInfo = $Sel.Tag
    # Track which account we're loading containers for to detect stale callbacks
    $Global:ContainerLoadForAccount = $AcctInfo.Name
    Write-DebugLog "Loading containers for: $($AcctInfo.Name)" -Level 'DEBUG'
    Start-BackgroundWork -Variables @{
        AccName = $AcctInfo.Name
        AccRG   = $AcctInfo.ResourceGroup
        ResolveOidFnSrc = $Script:ResolveOidFnSrc
    } -Context @{
        # Pin the loaded-account name at QUEUE TIME so the stale-check in OnComplete
        # doesn't read a clobbered script-scope $AcctInfo (which gets re-bound by
        # other SelectionChanged events fired between queue and complete).
        LoadedAcctName = $AcctInfo.Name
    } -Work {
        try {
            # Use OAuth context (Entra ID) instead of storage keys
            $SA  = Get-AzStorageAccount -Name $AccName -ResourceGroupName $AccRG -ErrorAction Stop
            $Ctx = New-AzStorageContext -StorageAccountName $AccName -UseConnectedAccount -ErrorAction Stop

            # Check Storage Blob Data Contributor role
            $RoleOk = $false
            $RoleMsg = ''
            try {
                # Recreate the helper inside this runspace (avoid cross-thread session state).
                Invoke-Expression $ResolveOidFnSrc
                $CurrentUser = (Get-AzContext).Account.Id
                $ObjectId = Resolve-CurrentUserObjectId -Email $CurrentUser
                if ($ObjectId) {
                    # IMPORTANT: Owner/Contributor are CONTROL-plane roles and do NOT grant
                    # data-plane blob access (uploads via -UseConnectedAccount). Only the
                    # 'Storage Blob Data *' roles work. We accept Owner/Contributor only as
                    # a hint that the user CAN self-assign the data role.
                    $AllRoles = @(Get-AzRoleAssignment -ObjectId $ObjectId -Scope $SA.Id -ErrorAction SilentlyContinue)
                    $DataRole = $AllRoles | Where-Object { $_.RoleDefinitionName -in @('Storage Blob Data Contributor','Storage Blob Data Owner') } | Select-Object -First 1
                    $CtrlRole = $AllRoles | Where-Object { $_.RoleDefinitionName -in @('Owner','Contributor','User Access Administrator') } | Select-Object -First 1
                    if ($DataRole) {
                        $RoleOk = $true
                        $RoleMsg = "Role: $($DataRole.RoleDefinitionName)"
                    } elseif ($CtrlRole) {
                        $RoleMsg = "Has $($CtrlRole.RoleDefinitionName) (control-plane only) but missing 'Storage Blob Data Contributor' (data-plane). Uploads will fail with 403 until assigned."
                    } else {
                        $RoleMsg = 'Missing Storage Blob Data Contributor role'
                    }
                } else {
                    # Could not resolve OID at all — surface as a warning so the prompt still fires
                    $RoleMsg = "Could not resolve directory ObjectId for '$CurrentUser' (guest user?). Cannot verify role — upload may fail with 403."
                }
            } catch { $RoleMsg = "Role check error: $($_.Exception.Message)" }

            $Containers = @(Get-AzStorageContainer -Context $Ctx -ErrorAction Stop | ForEach-Object { $_.Name })
            # Auto-create 'packages' container if none exist
            if ($Containers.Count -eq 0) {
                New-AzStorageContainer -Name 'packages' -Context $Ctx -Permission Blob -ErrorAction Stop | Out-Null
                $Containers = @('packages')
            }
            # Gather extended storage account info for the info panel
            $SubCtx = Get-AzContext -ErrorAction SilentlyContinue
            $SubName = if ($SubCtx) { $SubCtx.Subscription.Name } else { '' }
            $SubId   = if ($SubCtx) { $SubCtx.Subscription.Id   } else { '' }
            return @{
                Containers    = $Containers
                Error         = $null
                RoleOk        = $RoleOk
                RoleMsg       = $RoleMsg
                AcctId        = $SA.Id
                CreationTime  = if ($SA.CreationTime) { $SA.CreationTime.ToString('yyyy-MM-dd HH:mm') } else { '' }
                AccessTier    = [string]$SA.AccessTier
                HttpsOnly     = [string]$SA.EnableHttpsTrafficOnly
                MinTls        = [string]$SA.MinimumTlsVersion
                PublicAccess  = [string]$SA.AllowBlobPublicAccess
                PublicNetwork = [string]$SA.PublicNetworkAccess
                BlobEndpoint  = [string]$SA.PrimaryEndpoints.Blob
                SubName       = $SubName
                SubId         = $SubId
            }
        } catch {
            return @{ Containers = @(); Error = $_.Exception.Message; RoleOk = $false; RoleMsg = ''; AcctId = '' }
        }
    } -OnComplete {
        param($Results, $Errors, $Ctx)
        Write-DebugLog "[ContainerCB] OnComplete ENTERED (Results=$($Results.Count), Errors=$($Errors.Count))" -Level 'DEBUG'
        try {
        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ Containers = @() } }
        # Discard stale callback — storage account changed while containers were loading.
        # Use the QUEUE-TIME value pinned in $Ctx, not script-scope $AcctInfo (which
        # may have been re-bound by another SelectionChanged event in between).
        $CurrentAcct = if ($cmbUploadStorage.SelectedItem -and $cmbUploadStorage.SelectedItem.Tag) { $cmbUploadStorage.SelectedItem.Tag.Name } else { '' }
        $LoadedAcct = if ($Ctx -and $Ctx.LoadedAcctName) { $Ctx.LoadedAcctName } else { '' }
        if ($CurrentAcct -ne $LoadedAcct) {
            Write-DebugLog "[ContainerCB] Discarding stale container results for '$LoadedAcct' (current='$CurrentAcct')" -Level 'DEBUG'
            return
        }
        Write-DebugLog "[ContainerCB] cmbUploadContainer is null? $($null -eq $cmbUploadContainer)" -Level 'DEBUG'
        $cmbUploadContainer.Items.Clear()
        Write-DebugLog "[ContainerCB] Cleared items, adding $($R.Containers.Count) containers" -Level 'DEBUG'
        foreach ($C in $R.Containers) {
            $Item = New-Object System.Windows.Controls.ComboBoxItem
            $Item.Content = $C
            $cmbUploadContainer.Items.Add($Item) | Out-Null
        }
        if ($R.Containers.Count -gt 0) {
            # If a config load requested a specific container, select it; otherwise prefer 'packages'
            $TargetIdx = -1
            if ($Global:PendingContainerSelect) {
                for ($ci = 0; $ci -lt $cmbUploadContainer.Items.Count; $ci++) {
                    if ($cmbUploadContainer.Items[$ci].Content -eq $Global:PendingContainerSelect) { $TargetIdx = $ci; break }
                }
                $Global:PendingContainerSelect = $null
            }
            if ($TargetIdx -lt 0) {
                # Default: prefer container from settings
                $DefaultCont = $txtDefaultContainer.Text.Trim()
                if (-not $DefaultCont) { $DefaultCont = 'packages' }
                for ($ci = 0; $ci -lt $cmbUploadContainer.Items.Count; $ci++) {
                    if ($cmbUploadContainer.Items[$ci].Content -eq $DefaultCont) { $TargetIdx = $ci; break }
                }
            }
            $cmbUploadContainer.SelectedIndex = if ($TargetIdx -ge 0) { $TargetIdx } else { 0 }
        }
        # Report role prereq status
        if ($R.RoleOk) {
            Write-DebugLog "Storage prereq OK — $($R.RoleMsg)"
        } elseif ($R.RoleMsg) {
            Write-DebugLog "Storage prereq WARNING — $($R.RoleMsg)" -Level 'WARN'
            $Answer = Show-ThemedDialog -Title 'Storage Permission Required' `
                -Message "Your account is missing the Storage Blob Data Contributor role on this storage account.`n`nWithout it, uploads will fail with a 403 error.`n`nAssign the role now?" `
                -Icon ([string]([char]0xE72E)) -IconColor 'ThemeWarning' `
                -Buttons @(
                    @{ Text='Not Now'; IsAccent=$false; Result='No' }
                    @{ Text='Assign Role'; IsAccent=$true; Result='Yes' }
                )
            if ($Answer -eq 'Yes') {
                $SelAcctName = if ($cmbUploadStorage.SelectedItem -and $cmbUploadStorage.SelectedItem.Tag) { $cmbUploadStorage.SelectedItem.Tag.Name } else { '' }
                Assign-StorageBlobRole -StorageAccountId $R.AcctId -StorageAccountName $SelAcctName
            }
        }
        Write-DebugLog "Found $($R.Containers.Count) container(s)"

        # Update Storage Info panel with structured key-value rows
        Write-DebugLog "[ContainerCB] cmbUploadStorage is null? $($null -eq $cmbUploadStorage), SelectedItem is null? $($null -eq $cmbUploadStorage.SelectedItem)" -Level 'DEBUG'
        $AcctTag = $cmbUploadStorage.SelectedItem.Tag
        Write-DebugLog "[ContainerCB] AcctTag is null? $($null -eq $AcctTag)" -Level 'DEBUG'
        if ($AcctTag) {
            # Clear the storage info rows panel and rebuild with styled rows
            $pnlStorageInfoRows.Children.Clear()
            # Hide the placeholder label
            $lblStorageInfo.Visibility = 'Collapsed'

            # Helper to add a key-value row
            $AddInfoRow = {
                param([string]$Label, [string]$Value, [string]$Icon)
                $RowGrid = New-Object System.Windows.Controls.Grid
                $RowGrid.Margin = [System.Windows.Thickness]::new(0, 3, 0, 0)
                $Col0 = New-Object System.Windows.Controls.ColumnDefinition
                $Col0.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Auto)
                $Col1 = New-Object System.Windows.Controls.ColumnDefinition
                $Col1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
                $RowGrid.ColumnDefinitions.Add($Col0) | Out-Null
                $RowGrid.ColumnDefinitions.Add($Col1) | Out-Null

                # Label with icon
                $LblPanel = New-Object System.Windows.Controls.StackPanel
                $LblPanel.Orientation = 'Horizontal'
                $LblPanel.Margin = [System.Windows.Thickness]::new(0, 0, 8, 0)
                if ($Icon) {
                    $IconTb = New-Object System.Windows.Controls.TextBlock
                    $IconTb.Text = $Icon
                    $IconTb.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe Fluent Icons, Segoe MDL2 Assets")
                    $IconTb.FontSize = 9
                    $IconTb.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
                    $IconTb.VerticalAlignment = 'Center'
                    $IconTb.Margin = [System.Windows.Thickness]::new(0, 0, 4, 0)
                    $LblPanel.Children.Add($IconTb) | Out-Null
                }
                $LblTb = New-Object System.Windows.Controls.TextBlock
                $LblTb.Text = $Label
                $LblTb.FontSize = 9.5
                $LblTb.FontWeight = [System.Windows.FontWeights]::SemiBold
                $LblTb.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
                $LblTb.VerticalAlignment = 'Center'
                $LblPanel.Children.Add($LblTb) | Out-Null
                [System.Windows.Controls.Grid]::SetColumn($LblPanel, 0)
                $RowGrid.Children.Add($LblPanel) | Out-Null

                # Value
                $ValTb = New-Object System.Windows.Controls.TextBlock
                $ValTb.Text = if ($Value) { $Value } else { '—' }
                $ValTb.FontSize = 9.5
                $ValTb.TextWrapping = 'Wrap'
                $ValTb.TextTrimming = 'CharacterEllipsis'
                $ValTb.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextSecondary')
                $ValTb.VerticalAlignment = 'Center'
                [System.Windows.Controls.Grid]::SetColumn($ValTb, 1)
                $RowGrid.Children.Add($ValTb) | Out-Null

                $pnlStorageInfoRows.Children.Add($RowGrid) | Out-Null
            }

            # Section separator helper
            $AddSeparator = {
                $Sep = New-Object System.Windows.Controls.Border
                $Sep.Height = 1
                $Sep.Margin = [System.Windows.Thickness]::new(0, 6, 0, 3)
                $Sep.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeBorder')
                $pnlStorageInfoRows.Children.Add($Sep) | Out-Null
            }

            # General section
            & $AddInfoRow 'Name' $AcctTag.Name ([string][char]0xE753)
            & $AddInfoRow 'Resource Group' $AcctTag.ResourceGroup ([string][char]0xF168)
            & $AddInfoRow 'Region' $AcctTag.Location ([string][char]0xE774)
            if ($R.SubName) { & $AddInfoRow 'Subscription' $R.SubName ([string][char]0xE8D4) }

            & $AddSeparator

            # Configuration section
            & $AddInfoRow 'Kind' $AcctTag.Kind ([string][char]0xE950)
            & $AddInfoRow 'SKU' $AcctTag.Sku ([string][char]0xE945)
            if ($R.AccessTier) { & $AddInfoRow 'Access Tier' $R.AccessTier ([string][char]0xE74A) }
            if ($R.CreationTime) { & $AddInfoRow 'Created' $R.CreationTime ([string][char]0xE787) }

            & $AddSeparator

            # Security section
            if ($R.RoleMsg) { & $AddInfoRow 'Access' $R.RoleMsg ([string][char]0xE72E) }
            if ($R.HttpsOnly) { & $AddInfoRow 'HTTPS Only' $R.HttpsOnly ([string][char]0xE72E) }
            if ($R.MinTls) { & $AddInfoRow 'Min TLS' $R.MinTls ([string][char]0xE72E) }
            if ($R.PublicAccess) { & $AddInfoRow 'Blob Public Access' $R.PublicAccess ([string][char]0xE72E) }
            if ($R.PublicNetwork) { & $AddInfoRow 'Public Network' $R.PublicNetwork ([string][char]0xE774) }

            # ── Inline action: toggle public blob access (placed in Security section
            # for discoverability — no scrolling required to find it) ─────────────
            $IsPublic = ($R.PublicAccess -eq 'True')
            $TargetContainer = if ($cmbUploadContainer.SelectedItem) { [string]$cmbUploadContainer.SelectedItem.Content } else { 'packages' }
            $SelAcctTag = $cmbUploadStorage.SelectedItem.Tag

            $ActionCard = New-Object System.Windows.Controls.Border
            $ActionCard.Margin = [System.Windows.Thickness]::new(0, 8, 0, 4)
            $ActionCard.Padding = [System.Windows.Thickness]::new(8, 6, 8, 8)
            $ActionCard.CornerRadius = [System.Windows.CornerRadius]::new(4)
            $ActionCard.BorderThickness = [System.Windows.Thickness]::new(1)
            if ($IsPublic) {
                $ActionCard.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeWarning')
                $ActionCard.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeInputBg')
            } else {
                $ActionCard.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThemeBorder')
                $ActionCard.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeInputBg')
            }

            $ActionStack = New-Object System.Windows.Controls.StackPanel

            # Header row: icon + label + status pill
            $HeaderGrid = New-Object System.Windows.Controls.Grid
            $HeaderGrid.Margin = [System.Windows.Thickness]::new(0, 0, 0, 6)
            $HC0 = New-Object System.Windows.Controls.ColumnDefinition; $HC0.Width = 'Auto'
            $HC1 = New-Object System.Windows.Controls.ColumnDefinition; $HC1.Width = '*'
            $HeaderGrid.ColumnDefinitions.Add($HC0) | Out-Null
            $HeaderGrid.ColumnDefinitions.Add($HC1) | Out-Null

            $HdrLbl = New-Object System.Windows.Controls.TextBlock
            $HdrLbl.Text = 'PUBLIC ACCESS'
            $HdrLbl.FontSize = 9
            $HdrLbl.FontWeight = [System.Windows.FontWeights]::Bold
            $HdrLbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextDim')
            $HdrLbl.VerticalAlignment = 'Center'
            [System.Windows.Controls.Grid]::SetColumn($HdrLbl, 0)
            $HeaderGrid.Children.Add($HdrLbl) | Out-Null

            $StatusPill = New-Object System.Windows.Controls.Border
            $StatusPill.HorizontalAlignment = 'Right'
            $StatusPill.CornerRadius = [System.Windows.CornerRadius]::new(8)
            $StatusPill.Padding = [System.Windows.Thickness]::new(7, 1, 7, 1)
            if ($IsPublic) {
                $StatusPill.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeWarning')
            } else {
                $StatusPill.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ThemeGreenAccent')
            }
            $StatusTb = New-Object System.Windows.Controls.TextBlock
            $StatusTb.Text = if ($IsPublic) { 'PUBLIC' } else { 'PRIVATE' }
            $StatusTb.FontSize = 8.5
            $StatusTb.FontWeight = [System.Windows.FontWeights]::Bold
            $StatusTb.Foreground = [System.Windows.Media.Brushes]::White
            $StatusPill.Child = $StatusTb
            [System.Windows.Controls.Grid]::SetColumn($StatusPill, 1)
            $HeaderGrid.Children.Add($StatusPill) | Out-Null

            $ActionStack.Children.Add($HeaderGrid) | Out-Null

            # Action button
            $ActionBtn = New-Object System.Windows.Controls.Button
            $ActionBtn.Padding = [System.Windows.Thickness]::new(10, 5, 10, 5)
            $ActionBtn.HorizontalAlignment = 'Stretch'
            $ActionBtn.HorizontalContentAlignment = 'Center'
            $ActionBtn.FontSize = 10.5
            $ActionBtn.Cursor = [System.Windows.Input.Cursors]::Hand
            if ($IsPublic) {
                $ActionBtn.Content = "Disable public read on '$TargetContainer'"
                $ActionBtn.ToolTip = "Sets account flag AllowBlobPublicAccess=false and container ACL to Off. Anonymous downloads will stop working immediately."
            } else {
                $ActionBtn.Content = "Enable public read on '$TargetContainer'"
                $ActionBtn.ToolTip = "TEST MODE ONLY: Sets account flag AllowBlobPublicAccess=true and container ACL to Blob (anonymous read). Required for WinGet client installer downloads until Private Endpoints are in place."
            }
            try { $ActionBtn.Style = $Window.Resources['GhostButton'] } catch { }

            $CaptureAcct = $SelAcctTag.Name
            $CaptureRG   = $SelAcctTag.ResourceGroup
            $CaptureCnt  = $TargetContainer
            $CaptureMode = if ($IsPublic) { 'Disable' } else { 'Enable' }
            $ActionBtn.Add_Click({
                if ($CaptureMode -eq 'Enable') {
                    $Confirm = Show-ThemedDialog -Title 'Enable Public Read Access?' `
                        -Message "This will make all blobs in '$CaptureCnt' on '$CaptureAcct' downloadable by ANYONE on the internet (no auth required).`n`nUse only for testing / pre-production while Private Endpoints are not yet deployed.`n`nDisable again when VNet integration is complete." `
                        -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' `
                        -Buttons @(
                            @{ Text='Cancel'; IsAccent=$false; Result='Cancel' }
                            @{ Text='Enable Public Access'; IsAccent=$true; Result='OK' }
                        )
                } else {
                    $Confirm = Show-ThemedDialog -Title 'Disable Public Read Access?' `
                        -Message "This will set '$CaptureCnt' on '$CaptureAcct' back to private. Existing anonymous WinGet client downloads will fail immediately." `
                        -Icon ([string]([char]0xE72E)) -IconColor 'ThemeAccentLight' `
                        -Buttons @(
                            @{ Text='Cancel'; IsAccent=$false; Result='Cancel' }
                            @{ Text='Disable Public Access'; IsAccent=$true; Result='OK' }
                        )
                }
                if ($Confirm -eq 'OK') {
                    Set-StoragePublicAccess -StorageAccountName $CaptureAcct -ResourceGroupName $CaptureRG -ContainerName $CaptureCnt -Mode $CaptureMode
                }
            }.GetNewClosure())

            $ActionStack.Children.Add($ActionBtn) | Out-Null

            # Caption text (warning when public)
            $WarnTb = New-Object System.Windows.Controls.TextBlock
            $WarnTb.FontSize = 9
            $WarnTb.TextWrapping = 'Wrap'
            $WarnTb.Margin = [System.Windows.Thickness]::new(0, 4, 0, 0)
            if ($IsPublic) {
                $WarnTb.Text = "⚠ Anonymous downloads ENABLED on '$CaptureCnt'. Disable when Private Endpoints are live."
                $WarnTb.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeWarning')
            } else {
                $WarnTb.Text = "Private — WinGet clients require auth or PE. Enable for test/preview only."
                $WarnTb.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThemeTextMuted')
            }
            $ActionStack.Children.Add($WarnTb) | Out-Null

            $ActionCard.Child = $ActionStack
            $pnlStorageInfoRows.Children.Add($ActionCard) | Out-Null

            & $AddSeparator

            # Data section
            & $AddInfoRow 'Containers' "$($R.Containers.Count)" ([string][char]0xE838)
            if ($R.Containers.Count -gt 0) {
                $ContainerList = ($R.Containers | Sort-Object) -join ', '
                & $AddInfoRow '' $ContainerList $null
            }
            if ($R.BlobEndpoint) { & $AddInfoRow 'Blob Endpoint' $R.BlobEndpoint ([string][char]0xE71B) }

            Write-DebugLog "[ContainerCB] Storage info panel update DONE" -Level 'DEBUG'
        }
        Write-DebugLog "[ContainerCB] OnComplete EXIT (success)" -Level 'DEBUG'
        } catch {
            Write-DebugLog "[ContainerCB] CRASH: $($_.Exception.Message)" -Level 'ERROR'
            Write-DebugLog "[ContainerCB] ScriptStackTrace: $($_.ScriptStackTrace)" -Level 'ERROR'
        }
    }
})

# ── Blob Browser ──────────────────────────────────────────────────────────────
function Refresh-BlobBrowser {
    <# Lists blobs in the currently selected container and populates the blob browser #>
    $StorageSel = $cmbUploadStorage.SelectedItem
    $ContainerSel = $cmbUploadContainer.SelectedItem
    if (-not $StorageSel -or -not $StorageSel.Tag -or -not $ContainerSel) {
        $lstBlobBrowser.Items.Clear()
        $lblBlobCount.Text = "Select a container to browse blobs"
        $btnDeleteBlob.IsEnabled = $false
        return
    }
    $AcctName = $StorageSel.Tag.Name
    $AcctRG   = $StorageSel.Tag.ResourceGroup
    $ContName = $ContainerSel.Content.ToString()
    $lblBlobCount.Text = "Loading blobs..."
    $lstBlobBrowser.Items.Clear()

    Start-BackgroundWork -Variables @{
        AN = $AcctName; RG = $AcctRG; CN = $ContName
    } -Work {
        try {
            $Ctx = New-AzStorageContext -StorageAccountName $AN -UseConnectedAccount -ErrorAction Stop
            $Blobs = @(Get-AzStorageBlob -Container $CN -Context $Ctx -ErrorAction Stop |
                Select-Object Name, Length, LastModified, ContentType)
            return @{ Blobs = $Blobs; Error = $null }
        } catch {
            return @{ Blobs = @(); Error = $_.Exception.Message }
        }
    } -OnComplete {
        param($Results, $Errors)
        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ Blobs = @(); Error = 'No result' } }
        $lstBlobBrowser.Items.Clear()
        if ($R.Error) {
            $lblBlobCount.Text = "Error: $($R.Error)"
            Write-DebugLog "Blob browser error: $($R.Error)" -Level 'ERROR'
            return
        }
        if ($R.Blobs.Count -eq 0) {
            $lblBlobCount.Text = "Container is empty"
            return
        }
        foreach ($B in $R.Blobs) {
            $SizeKB = if ($B.Length) { [math]::Round($B.Length / 1KB, 1) } else { 0 }
            $SizeStr = if ($SizeKB -ge 1024) { "$([math]::Round($SizeKB / 1024, 1)) MB" } else { "$SizeKB KB" }
            $ModStr  = if ($B.LastModified) { $B.LastModified.ToString('yyyy-MM-dd HH:mm') } else { '' }
            $Item = New-StyledListViewItem -IconChar ([char]0xE8B7) -PrimaryText $B.Name `
                -SubtitleText "$SizeStr  |  $ModStr" -TagValue $B.Name
            $lstBlobBrowser.Items.Add($Item) | Out-Null
        }
        $lblBlobCount.Text = "$($R.Blobs.Count) blob(s) in container"
        $btnDeleteBlob.IsEnabled = $false
    }
}

# Refresh blobs when container selection changes
$cmbUploadContainer.Add_SelectionChanged({
    if ($cmbUploadContainer.SelectedItem) { Refresh-BlobBrowser }
})

# Enable delete button when a blob is selected
$lstBlobBrowser.Add_SelectionChanged({
    $btnDeleteBlob.IsEnabled = ($lstBlobBrowser.SelectedItem -ne $null)
})

$btnRefreshBlobs.Add_Click({ Refresh-BlobBrowser })

$btnDeleteBlob.Add_Click({
    $Sel = $lstBlobBrowser.SelectedItem
    if (-not $Sel -or -not $Sel.Tag) {
        Show-Toast "Select a blob to delete" -Type Warning
        return
    }
    $BlobName = $Sel.Tag
    $ContName = $cmbUploadContainer.SelectedItem.Content.ToString()
    $AcctName = $cmbUploadStorage.SelectedItem.Tag.Name
    $Confirm = Show-ThemedDialog -Title 'Delete Blob' `
        -Message "Delete blob '$BlobName' from container '$ContName'?`n`nThis cannot be undone." `
        -Icon ([string]([char]0xE74D)) -IconColor 'ThemeError' `
        -Buttons @(
            @{ Text='Cancel'; IsAccent=$false; Result='Cancel' }
            @{ Text='Delete'; IsAccent=$true; Result='Delete' }
        )
    if ($Confirm -ne 'Delete') { return }
    $AcctRG = $cmbUploadStorage.SelectedItem.Tag.ResourceGroup
    $lblBlobCount.Text = "Deleting..."
    Start-BackgroundWork -Variables @{
        AN = $AcctName; RG = $AcctRG; CN = $ContName; BN = $BlobName
    } -Work {
        try {
            $Ctx = New-AzStorageContext -StorageAccountName $AN -UseConnectedAccount -ErrorAction Stop
            Remove-AzStorageBlob -Container $CN -Blob $BN -Context $Ctx -Force -ErrorAction Stop
            return @{ Success = $true }
        } catch {
            return @{ Success = $false; Error = $_.Exception.Message }
        }
    } -OnComplete {
        param($Results, $Errors)
        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ Success = $false; Error = 'No result' } }
        if ($R.Success) {
            Show-Toast "Blob deleted: $BlobName" -Type Success
            Write-DebugLog "Deleted blob: $ContName/$BlobName" -Level 'INFO'
            Unlock-Achievement 'blob_cleanup'
            Refresh-BlobBrowser
        } else {
            Show-Toast "Delete failed: $($R.Error)" -Type Error
            Write-DebugLog "Blob delete failed: $($R.Error)" -Level 'ERROR'
            $lblBlobCount.Text = "Delete failed"
        }
    }
})

# ── Prereq UI ─────────────────────────────────────────────────────────────────
$btnCopyCommand.Add_Click({
    if ($txtPrereqCommand.Text) {
        [System.Windows.Clipboard]::SetText($txtPrereqCommand.Text)
        $lblStatus.Text = "Command copied to clipboard"
    }
})

$btnRetryPrereq.Add_Click({
    Write-DebugLog "Retrying prerequisite check..."
    $Results = Test-Prerequisites
    Show-PrereqResults -Results $Results
})

# ==============================================================================
# DEPLOY TAB — GUI wrapper for Deploy-WinGetSource.ps1 / Remove-WinGetSource.ps1
# ==============================================================================

$txtDeployName          = $Window.FindName("txtDeployName")
$txtDeployRG            = $Window.FindName("txtDeployRG")
$cmbDeployRegion        = $Window.FindName("cmbDeployRegion")
$cmbDeployTier          = $Window.FindName("cmbDeployTier")
$cmbDeployAuth          = $Window.FindName("cmbDeployAuth")
$txtDeployDisplayName   = $Window.FindName("txtDeployDisplayName")
$txtDeployManifestPath  = $Window.FindName("txtDeployManifestPath")
$btnDeployBrowseManifests = $Window.FindName("btnDeployBrowseManifests")
$chkDeploySkipAPIM      = $Window.FindName("chkDeploySkipAPIM")
$chkDeployRegisterSource = $Window.FindName("chkDeployRegisterSource")
$chkDeployCosmosZoneRedundant = $Window.FindName("chkDeployCosmosZoneRedundant")
$cmbDeployProfile       = $Window.FindName("cmbDeployProfile")
$btnSaveDeployProfile   = $Window.FindName("btnSaveDeployProfile")
$btnDeleteDeployProfile = $Window.FindName("btnDeleteDeployProfile")
$txtDeployPublisherName = $Window.FindName("txtDeployPublisherName")
$txtDeployPublisherEmail = $Window.FindName("txtDeployPublisherEmail")
$cmbDeployCosmosRegion  = $Window.FindName("cmbDeployCosmosRegion")
$chkDeployEnablePrivateEndpoints = $Window.FindName("chkDeployEnablePrivateEndpoints")
$pnlPrivateEndpointFields = $Window.FindName("pnlPrivateEndpointFields")
$txtDeployVNetName      = $Window.FindName("txtDeployVNetName")
$txtDeploySubnetName    = $Window.FindName("txtDeploySubnetName")
$txtDeployVNetRG        = $Window.FindName("txtDeployVNetRG")
$chkDeployRegisterDnsZones = $Window.FindName("chkDeployRegisterDnsZones")
$chkDeployWhatIf        = $Window.FindName("chkDeployWhatIf")
$chkRemoveSkipSourceRemoval = $Window.FindName("chkRemoveSkipSourceRemoval")
$chkRemoveSkipEntraCleanup = $Window.FindName("chkRemoveSkipEntraCleanup")
$chkRemoveSkipSoftDeletePurge = $Window.FindName("chkRemoveSkipSoftDeletePurge")
$chkRemovePurgeOnly     = $Window.FindName("chkRemovePurgeOnly")
$cmbRemoveOverrideFuncApp = $Window.FindName("cmbRemoveOverrideFuncApp")
$cmbRemoveOverrideKeyVault = $Window.FindName("cmbRemoveOverrideKeyVault")
$cmbRemoveOverrideStorage = $Window.FindName("cmbRemoveOverrideStorage")
$cmbRemoveOverrideAPIM  = $Window.FindName("cmbRemoveOverrideAPIM")
$cmbRemoveName          = $Window.FindName("cmbRemoveName")
$cmbRemoveRG            = $Window.FindName("cmbRemoveRG")
$cmbRemoveProfile       = $Window.FindName("cmbRemoveProfile")
$txtRemoveDisplayName   = $Window.FindName("txtRemoveDisplayName")
$btnRunDeploy           = $Window.FindName("btnRunDeploy")
$btnRunRemove           = $Window.FindName("btnRunRemove")
$btnDeployCancel        = $Window.FindName("btnDeployCancel")
$btnRemoveCancel        = $Window.FindName("btnRemoveCancel")
$prgDeploy              = $Window.FindName("prgDeploy")
$prgRemove              = $Window.FindName("prgRemove")
$lblDeployStatus        = $Window.FindName("lblDeployStatus")
$lblRemoveStatus        = $Window.FindName("lblRemoveStatus")
$lblDeployTargetSub     = $Window.FindName("lblDeployTargetSub")
$lblRemoveTargetSub     = $Window.FindName("lblRemoveTargetSub")

# Toggle Private Endpoint fields visibility
$chkDeployEnablePrivateEndpoints.Add_Checked({ $pnlPrivateEndpointFields.Visibility = 'Visible' })
$chkDeployEnablePrivateEndpoints.Add_Unchecked({ $pnlPrivateEndpointFields.Visibility = 'Collapsed' })

$Global:DeployConfigsPath = Join-Path $Global:Root "deploy_profiles.json"
$Global:DeployProcess     = $null

# ── Deploy Profile Save/Load ─────────────────────────────────────────────────
function Get-DeployProfiles {
    if (Test-Path $Global:DeployConfigsPath) {
        try { return (Get-Content $Global:DeployConfigsPath -Raw | ConvertFrom-Json) }
        catch { <# Corrupted profiles file — return empty list #> return @() }
    }
    return @()
}

function Save-DeployProfiles { param($Profiles)
    $Profiles | ConvertTo-Json -Depth 5 | Set-Content $Global:DeployConfigsPath -Force
}

function Refresh-DeployProfileCombo {
    $cmbDeployProfile.Items.Clear()
    $Profiles = Get-DeployProfiles
    foreach ($P in $Profiles) {
        $Item = New-Object System.Windows.Controls.ComboBoxItem
        $Item.Content = $P.ProfileName
        $Item.Tag = $P
        $cmbDeployProfile.Items.Add($Item) | Out-Null
    }
    # Also populate Remove tab profile combo
    if ($cmbRemoveProfile) {
        $cmbRemoveProfile.Items.Clear()
        foreach ($P in $Profiles) {
            $Item = New-Object System.Windows.Controls.ComboBoxItem
            $Item.Content = $P.ProfileName
            $Item.Tag = $P
            $cmbRemoveProfile.Items.Add($Item) | Out-Null
        }
    }
}

function Get-DeployParamsFromUI {
    return @{
        ProfileName     = $txtDeployName.Text.Trim()
        Name            = $txtDeployName.Text.Trim()
        ResourceGroup   = $txtDeployRG.Text.Trim()
        Region          = (Get-ComboBoxSelectedText $cmbDeployRegion)
        PerformanceTier = (Get-ComboBoxSelectedText $cmbDeployTier)
        Authentication  = (Get-ComboBoxSelectedText $cmbDeployAuth)
        DisplayName     = $txtDeployDisplayName.Text.Trim()
        PublisherName   = $txtDeployPublisherName.Text.Trim()
        PublisherEmail  = $txtDeployPublisherEmail.Text.Trim()
        CosmosDBRegion  = $(  $r = Get-ComboBoxSelectedText $cmbDeployCosmosRegion
                              if ($r -eq '(Same as Region)') { '' } else { $r }  )
        ManifestPath    = $txtDeployManifestPath.Text.Trim()
        SkipAPIM        = [bool]$chkDeploySkipAPIM.IsChecked
        RegisterSource  = [bool]$chkDeployRegisterSource.IsChecked
        CosmosZoneRedundant = [bool]$chkDeployCosmosZoneRedundant.IsChecked
        EnablePrivateEndpoints = [bool]$chkDeployEnablePrivateEndpoints.IsChecked
        VNetName        = $txtDeployVNetName.Text.Trim()
        SubnetName      = $txtDeploySubnetName.Text.Trim()
        VNetResourceGroup = $txtDeployVNetRG.Text.Trim()
        RegisterPrivateDnsZones = [bool]$chkDeployRegisterDnsZones.IsChecked
        WhatIf          = [bool]$chkDeployWhatIf.IsChecked
        BackupStorageAcct = $cmbBackupStorageAcct.Text.Trim()
    }
}

function Populate-DeployFromProfile { param($P)
    $txtDeployName.Text = if ($P.Name) { $P.Name } else { '' }
    $txtDeployRG.Text = if ($P.ResourceGroup) { $P.ResourceGroup } else { '' }
    if ($P.Region) { Set-ComboBoxByContent $cmbDeployRegion $P.Region }
    if ($P.PerformanceTier) { Set-ComboBoxByContent $cmbDeployTier $P.PerformanceTier }
    if ($P.Authentication) { Set-ComboBoxByContent $cmbDeployAuth $P.Authentication }
    $txtDeployDisplayName.Text = if ($P.DisplayName) { $P.DisplayName } else { '' }
    $txtDeployPublisherName.Text = if ($P.PublisherName) { $P.PublisherName } else { '' }
    $txtDeployPublisherEmail.Text = if ($P.PublisherEmail) { $P.PublisherEmail } else { '' }
    if ($P.CosmosDBRegion) { Set-ComboBoxByContent $cmbDeployCosmosRegion $P.CosmosDBRegion } else { $cmbDeployCosmosRegion.SelectedIndex = 0 }
    $txtDeployManifestPath.Text = if ($P.ManifestPath) { $P.ManifestPath } else { '' }
    $chkDeploySkipAPIM.IsChecked = [bool]$P.SkipAPIM
    $chkDeployRegisterSource.IsChecked = [bool]$P.RegisterSource
    $chkDeployCosmosZoneRedundant.IsChecked = [bool]$P.CosmosZoneRedundant
    $chkDeployEnablePrivateEndpoints.IsChecked = [bool]$P.EnablePrivateEndpoints
    $txtDeployVNetName.Text = if ($P.VNetName) { $P.VNetName } else { '' }
    $txtDeploySubnetName.Text = if ($P.SubnetName) { $P.SubnetName } else { '' }
    $txtDeployVNetRG.Text = if ($P.VNetResourceGroup) { $P.VNetResourceGroup } else { '' }
    $chkDeployRegisterDnsZones.IsChecked = [bool]$P.RegisterPrivateDnsZones
    if ($P.BackupStorageAcct) { $cmbBackupStorageAcct.Text = $P.BackupStorageAcct }
}

# Load profiles on startup
Refresh-DeployProfileCombo

$cmbDeployProfile.Add_SelectionChanged({
    $Sel = $cmbDeployProfile.SelectedItem
    if ($Sel -and $Sel.Tag) { Populate-DeployFromProfile $Sel.Tag }
})

# Remove tab: auto-fill overrides when source is selected
if ($cmbRemoveName) {
    $cmbRemoveName.Add_SelectionChanged({
        $Sel = $cmbRemoveName.SelectedItem
        if ($Sel -and $Sel -is [System.Windows.Controls.ComboBoxItem] -and $Sel.Tag) {
            $cmbRemoveRG.Text = $Sel.Tag.ResourceGroup
            # Auto-populate resource name overrides using standard naming convention
            $SourceName = $Sel.Content.ToString() -replace '[^a-zA-Z0-9-]', ''
            $cmbRemoveOverrideFuncApp.Text  = "func-$SourceName"
            $cmbRemoveOverrideKeyVault.Text = "kv-$SourceName"
            $cmbRemoveOverrideStorage.Text  = "st$($SourceName -replace '-', '')"
            $cmbRemoveOverrideAPIM.Text     = "apim-$SourceName"
        }
    })
}

# Remove tab: load profile into Remove fields
if ($cmbRemoveProfile) {
    $cmbRemoveProfile.Add_SelectionChanged({
        $Sel = $cmbRemoveProfile.SelectedItem
        if (-not $Sel -or -not $Sel.Tag) { return }
        $P = $Sel.Tag
        if ($cmbRemoveName) { $cmbRemoveName.Text = if ($P.Name) { $P.Name } else { '' } }
        if ($cmbRemoveRG) { $cmbRemoveRG.Text = if ($P.ResourceGroup) { $P.ResourceGroup } else { '' } }
        if ($txtRemoveDisplayName) { $txtRemoveDisplayName.Text = if ($P.DisplayName) { $P.DisplayName } else { '' } }
        # Auto-derive resource overrides from naming convention
        $CleanName = ($P.Name -replace '[^a-zA-Z0-9-]', '')
        if ($CleanName) {
            $cmbRemoveOverrideFuncApp.Text  = "func-$CleanName"
            $cmbRemoveOverrideKeyVault.Text = "kv-$CleanName"
            $cmbRemoveOverrideStorage.Text  = "st$($CleanName -replace '-', '')"
            $cmbRemoveOverrideAPIM.Text     = "apim-$CleanName"
        }
    })
}

$btnSaveDeployProfile.Add_Click({
    $Params = Get-DeployParamsFromUI
    if (-not $Params.Name) {
        Show-Toast "Enter a source name first" -Type Warning
        return
    }
    $Profiles = @(Get-DeployProfiles)
    $Existing = $Profiles | Where-Object { $_.ProfileName -eq $Params.ProfileName }
    if ($Existing) {
        $Confirm = Show-ThemedDialog -Title 'Overwrite Profile' `
            -Message "Profile '$($Params.ProfileName)' already exists. Overwrite?" `
            -Icon ([string]([char]0xE74C)) -IconColor 'ThemeWarning' `
            -Buttons @(
                @{ Text='Cancel'; IsAccent=$false; Result='Cancel' }
                @{ Text='Overwrite'; IsAccent=$true; Result='Overwrite' }
            )
        if ($Confirm -ne 'Overwrite') { return }
        # Update in-place
        $Profiles = @($Profiles | ForEach-Object {
            if ($_.ProfileName -eq $Params.ProfileName) { [PSCustomObject]$Params } else { $_ }
        })
    } else {
        $Profiles += [PSCustomObject]$Params
    }
    Save-DeployProfiles $Profiles
    Refresh-DeployProfileCombo
    # Re-select
    for ($i = 0; $i -lt $cmbDeployProfile.Items.Count; $i++) {
        if ($cmbDeployProfile.Items[$i].Content -eq $Params.ProfileName) {
            $cmbDeployProfile.SelectedIndex = $i; break
        }
    }
    Show-Toast "Profile '$($Params.ProfileName)' saved" -Type Success
})

$btnDeleteDeployProfile.Add_Click({
    $Sel = $cmbDeployProfile.SelectedItem
    if (-not $Sel) { Show-Toast "Select a profile to delete" -Type Warning; return }
    $Name = $Sel.Content
    $Confirm = Show-ThemedDialog -Title 'Delete Profile' `
        -Message "Delete deployment profile '$Name'?" `
        -Icon ([string]([char]0xE74D)) -IconColor 'ThemeWarning' `
        -Buttons @(
            @{ Text='Cancel'; IsAccent=$false; Result='Cancel' }
            @{ Text='Delete'; IsDanger=$true;  Result='Delete' }
        )
    if ($Confirm -ne 'Delete') { return }
    $Profiles = @(Get-DeployProfiles) | Where-Object { $_.ProfileName -ne $Name }
    Save-DeployProfiles @($Profiles)
    Refresh-DeployProfileCombo
    Show-Toast "Profile '$Name' deleted" -Type Info
})

$btnDeployBrowseManifests.Add_Click({
    $Dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $Dlg.Description = "Select folder containing WinGet YAML manifests"
    if ($Dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtDeployManifestPath.Text = $Dlg.SelectedPath
    }
})

# ── Script Execution via Runspace ────────────────────────────────────────────
function Start-DeployScript {
    param([string]$ScriptPath, [hashtable]$Arguments, [string]$Label)

    if ($Global:DeployProcess) {
        Show-Toast "A deployment operation is already running" -Type Warning
        return
    }

    $FullPath = Join-Path $Global:Root $ScriptPath
    if (-not (Test-Path $FullPath)) {
        Show-Toast "Script not found: $ScriptPath" -Type Error
        return
    }

    Write-DebugLog "Starting $Label ..." -Level 'INFO'
    Show-GlobalProgressIndeterminate -Text "$Label in progress..."
    $lblDeployStatus.Text = "$Label in progress..."
    $prgDeploy.Visibility = 'Visible'
    $btnDeployCancel.Visibility = 'Visible'
    $btnRunDeploy.IsEnabled = $false
    $btnRunRemove.IsEnabled = $false

    # Build argument string
    $ArgParts = @()
    foreach ($K in $Arguments.Keys) {
        $V = $Arguments[$K]
        if ($V -is [switch] -or $V -is [bool]) {
            if ($V) { $ArgParts += "-$K" }
        } elseif ($V -and $V.ToString().Trim()) {
            $Safe = $V.ToString() -replace '"', '\"'
            $ArgParts += "-$K `"$Safe`""
        }
    }
    $ArgString = $ArgParts -join ' '

    $Global:DeployProcess = [System.Diagnostics.Process]::new()
    $Global:DeployProcess.StartInfo = [System.Diagnostics.ProcessStartInfo]@{
        FileName               = 'pwsh.exe'
        Arguments              = "-NoProfile -ExecutionPolicy Bypass -File `"$FullPath`" $ArgString -Confirm:`$false"
        UseShellExecute        = $false
        RedirectStandardOutput = $true
        RedirectStandardError  = $true
        CreateNoWindow         = $true
        WorkingDirectory       = $Global:Root
    }

    # Capture output asynchronously via events — queue to UI timer
    # Uses compiled C# ProcessOutputHelper to avoid "no Runspace available" crash
    # on .NET thread pool threads (Process events fire outside the PS runspace).
    $Global:DeployOutputQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
    $Helper = [ProcessOutputHelper]::new($Global:DeployOutputQueue)

    $Global:DeployProcess.EnableRaisingEvents = $true
    $Global:DeployProcess.add_Exited(
        [System.Delegate]::CreateDelegate([System.EventHandler], $Helper, 'OnExited'))
    $Global:DeployProcess.add_OutputDataReceived(
        [System.Delegate]::CreateDelegate([System.Diagnostics.DataReceivedEventHandler], $Helper, 'OnOutput'))
    $Global:DeployProcess.add_ErrorDataReceived(
        [System.Delegate]::CreateDelegate([System.Diagnostics.DataReceivedEventHandler], $Helper, 'OnError'))

    $Global:DeployProcess.Start() | Out-Null
    $Global:DeployProcess.BeginOutputReadLine()
    $Global:DeployProcess.BeginErrorReadLine()

    # Show the output panel
    if (-not $Global:BottomPanelOpen) { Toggle-BottomPanel $true }

    Write-DebugLog "$Label started (PID: $($Global:DeployProcess.Id))" -Level 'INFO'
}

# Timer polls DeployOutputQueue (handled in the existing dispatcher timer)
$Global:DeployOutputQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

$btnRunDeploy.Add_Click({
    $Params = Get-DeployParamsFromUI
    if (-not $Params.Name) {
        Show-Toast "Source Name is required" -Type Warning
        return
    }
    $Confirm = Show-ThemedDialog -Title 'Deploy WinGet Source' `
        -Message "Deploy '$($Params.Name)' to $($Params.ResourceGroup) in $($Params.Region)?`nSubscription: $($lblDeployTargetSub.Text)`nTier: $($Params.PerformanceTier)`nAuth: $($Params.Authentication)" `
        -Icon ([string]([char]0xE72E)) -IconColor 'ThemeAccent' `
        -Buttons @(
            @{ Text='Cancel'; IsAccent=$false; Result='Cancel' }
            @{ Text='Deploy'; IsAccent=$true;  Result='Deploy' }
        )
    if ($Confirm -ne 'Deploy') { return }

    $Args = @{
        Name = $Params.Name
        ResourceGroup = $Params.ResourceGroup
        Region = $Params.Region
        PerformanceTier = $Params.PerformanceTier
        Authentication = $Params.Authentication
        SourceDisplayName = $Params.DisplayName
    }
    if ($Params.PublisherName) { $Args.PublisherName = $Params.PublisherName }
    if ($Params.PublisherEmail) { $Args.PublisherEmail = $Params.PublisherEmail }
    if ($Params.CosmosDBRegion) { $Args.CosmosDBRegion = $Params.CosmosDBRegion }
    if ($Params.ManifestPath) { $Args.ManifestPath = $Params.ManifestPath }
    if ($Params.SkipAPIM) { $Args.SkipAPIM = $true }
    if ($Params.RegisterSource) { $Args.RegisterSource = $true }
    if ($Params.CosmosZoneRedundant) { $Args.CosmosDBZoneRedundant = $true }
    if ($Params.EnablePrivateEndpoints) {
        $Args.EnablePrivateEndpoints = $true
        if ($Params.VNetName) { $Args.VNetName = $Params.VNetName }
        if ($Params.SubnetName) { $Args.SubnetName = $Params.SubnetName }
        if ($Params.VNetResourceGroup) { $Args.VNetResourceGroup = $Params.VNetResourceGroup }
        if ($Params.RegisterPrivateDnsZones) { $Args.RegisterPrivateDnsZones = $true }
    }
    if ($Params.WhatIf) { $Args.WhatIf = $true }

    # Pass current subscription to deploy script
    $SelSub = $cmbSubscription.SelectedItem
    if ($SelSub -and $SelSub.Tag) { $Args.SubscriptionId = $SelSub.Tag }

    Start-DeployScript -ScriptPath 'Deploy-WinGetSource.ps1' -Arguments $Args -Label $(if ($Params.WhatIf) { 'Deployment (WhatIf)' } else { 'Deployment' })
})

$btnRunRemove.Add_Click({
    $Name = $cmbRemoveName.Text.Trim()
    if (-not $Name) {
        Show-Toast "Source Name is required" -Type Warning
        return
    }
    $RG = $cmbRemoveRG.Text.Trim()
    if (-not $RG) {
        Show-Toast "Resource Group is required" -Type Warning
        return
    }
    $SubText = $lblRemoveTargetSub.Text
    $Confirm = Show-ThemedDialog -Title 'Remove WinGet Source' `
        -Message "This will PERMANENTLY DELETE all Azure resources for '$Name' in resource group '$RG'.`nSubscription: $SubText`n`nAffected: Function App, Cosmos DB, Storage, Key Vault, APIM, App Config, Entra ID registrations." `
        -Icon ([string]([char]0xE74D)) -IconColor 'ThemeError' `
        -InputPrompt "Type '$Name' to confirm:" `
        -InputMatch $Name `
        -Buttons @(
            @{ Text='Cancel'; IsAccent=$false; Result='Cancel' }
            @{ Text='Remove'; IsDanger=$true;  Result='Remove' }
        )
    if ($Confirm -ne 'Remove') { return }

    $Args = @{
        Name = $Name
        ResourceGroup = $RG
        SourceDisplayName = $txtRemoveDisplayName.Text.Trim()
    }
    if ([bool]$chkRemoveSkipSourceRemoval.IsChecked) { $Args.SkipSourceRemoval = $true }
    if ([bool]$chkRemoveSkipEntraCleanup.IsChecked) { $Args.SkipEntraCleanup = $true }
    if ([bool]$chkRemoveSkipSoftDeletePurge.IsChecked) { $Args.SkipSoftDeletePurge = $true }
    if ([bool]$chkRemovePurgeOnly.IsChecked) { $Args.PurgeOnly = $true }

    # Resource name overrides — pass as env vars
    $Env:WINGETMM_OVERRIDE_FUNCAPP  = $cmbRemoveOverrideFuncApp.Text.Trim()
    $Env:WINGETMM_OVERRIDE_KEYVAULT = $cmbRemoveOverrideKeyVault.Text.Trim()
    $Env:WINGETMM_OVERRIDE_STORAGE  = $cmbRemoveOverrideStorage.Text.Trim()
    $Env:WINGETMM_OVERRIDE_APIM     = $cmbRemoveOverrideAPIM.Text.Trim()

    # Pass current subscription name to remove script
    $SelSub = $cmbSubscription.SelectedItem
    if ($SelSub -and $SelSub.Content) { $Args.SubscriptionName = $SelSub.Content }

    # Track that this is a remove operation for status/cancel handling
    $Global:DeployContext = 'Remove'
    $lblRemoveStatus.Text = 'Removal in progress...'
    $prgRemove.Visibility = 'Visible'
    $btnRemoveCancel.Visibility = 'Visible'
    $btnRunRemove.IsEnabled = $false

    Start-DeployScript -ScriptPath 'Remove-WinGetSource.ps1' -Arguments $Args -Label 'Removal'
})

$btnRemoveCancel.Add_Click({
    if ($Global:DeployProcess -and -not $Global:DeployProcess.HasExited) {
        try { $Global:DeployProcess.Kill($true) } catch { <# Process may have already exited #> }
        Write-DebugLog "Remove operation cancelled by user" -Level 'WARN'
    }
    $Global:DeployProcess = $null
    $Global:DeployContext = $null
    $prgRemove.Visibility = 'Collapsed'
    $btnRemoveCancel.Visibility = 'Collapsed'
    $btnRunRemove.IsEnabled = $true
    $lblRemoveStatus.Text = 'Cancelled'
    Hide-GlobalProgress
})

$btnDeployCancel.Add_Click({
    if ($Global:DeployProcess -and -not $Global:DeployProcess.HasExited) {
        try { $Global:DeployProcess.Kill($true) } catch { <# Process may have already exited #> }
        Write-DebugLog "Deploy operation cancelled by user" -Level 'WARN'
    }
    $Global:DeployProcess = $null
    $Global:DeployContext = $null
    $prgDeploy.Visibility = 'Collapsed'
    $btnDeployCancel.Visibility = 'Collapsed'
    $btnRunDeploy.IsEnabled = $true
    $lblDeployStatus.Text = "Cancelled"
    Hide-GlobalProgress
})

# ==============================================================================
# DISCOVER EXISTING DEPLOYMENTS
# ==============================================================================

$btnDiscoverDeploys     = $Window.FindName("btnDiscoverDeploys")
$lstDiscoveredDeploys   = $Window.FindName("lstDiscoveredDeploys")
$lblDiscoverStatus      = $Window.FindName("lblDiscoverStatus")
$lblDiscoverEmpty       = $Window.FindName("lblDiscoverEmpty")
$prgDiscover            = $Window.FindName("prgDiscover")

$btnDiscoverDeploys.Add_Click({
    if (-not $Global:IsAuthenticated) {
        Show-Toast "Sign in to Azure first" -Type Warning
        return
    }
    $lblDiscoverStatus.Text = "Scanning subscription..."
    $prgDiscover.Visibility = 'Visible'
    Show-GlobalProgressIndeterminate -Text 'Scanning subscription for deployments...'
    $lstDiscoveredDeploys.Items.Clear()
    $lstDiscoveredDeploys.Visibility = 'Collapsed'
    $lblDiscoverEmpty.Visibility = 'Collapsed'
    Write-DebugLog "Discovering WinGet REST source deployments..."

    Start-BackgroundWork -Work {
        try {
            # Find Function Apps that match the WinGet REST source naming pattern
            $FuncApps = @(Get-AzFunctionApp -ErrorAction Stop | Where-Object {
                $_.Name -match '^func-' -or $_.Name -match 'winget' -or
                ($_.Tags -and ($_.Tags.ContainsKey('WinGetSource') -or $_.Tags.ContainsKey('winget-restsource')))
            } | ForEach-Object {
                [PSCustomObject]@{
                    Name          = $_.Name
                    ResourceGroup = $_.ResourceGroup
                    Location      = $_.Location
                    State         = $_.State
                    Kind          = 'FunctionApp'
                    DefaultHostName = $_.DefaultHostName
                }
            })

            # Also find storage accounts that likely belong to WinGet deployments
            $StorageAccts = @(Get-AzStorageAccount -ErrorAction Stop | Where-Object {
                $_.StorageAccountName -match 'winget|wgsrc' -or
                ($_.Tags -and ($_.Tags.ContainsKey('WinGetSource') -or $_.Tags.ContainsKey('winget-restsource')))
            } | ForEach-Object {
                [PSCustomObject]@{
                    Name          = $_.StorageAccountName
                    ResourceGroup = $_.ResourceGroupName
                    Location      = $_.PrimaryLocation
                    State         = $_.ProvisioningState
                    Kind          = 'StorageAccount'
                    DefaultHostName = "$($_.StorageAccountName).blob.core.windows.net"
                }
            })

            return @{ FuncApps = $FuncApps; StorageAccts = $StorageAccts; Error = $null }
        } catch {
            return @{ FuncApps = @(); StorageAccts = @(); Error = $_.Exception.Message }
        }
    } -OnComplete {
        param($Results, $Errors)
        $prgDiscover.Visibility = 'Collapsed'
        Hide-GlobalProgress
        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ FuncApps = @(); StorageAccts = @(); Error = 'No result' } }

        if ($R.Error) {
            Write-DebugLog "Discovery failed: $($R.Error)" -Level 'ERROR'
            $lblDiscoverStatus.Text = "Scan failed"
            $lblDiscoverEmpty.Text = "Error: $($R.Error)"
            $lblDiscoverEmpty.Visibility = 'Visible'
            return
        }

        $AllItems = @($R.FuncApps) + @($R.StorageAccts)
        if ($AllItems.Count -eq 0) {
            $lblDiscoverStatus.Text = "No deployments found"
            $lblDiscoverEmpty.Text = "No WinGet REST source deployments found in this subscription."
            $lblDiscoverEmpty.Visibility = 'Visible'
            return
        }

        $lstDiscoveredDeploys.Items.Clear()
        foreach ($D in $AllItems) {
            $IconChar = if ($D.Kind -eq 'FunctionApp') { [char]0xE943 } else { [char]0xE753 }
            $Subtitle = "$($D.Kind) | $($D.ResourceGroup) | $($D.Location) | $($D.State)"
            $Item = New-StyledListViewItem -IconChar $IconChar -PrimaryText $D.Name `
                -SubtitleText $Subtitle -TagValue $D
            $lstDiscoveredDeploys.Items.Add($Item) | Out-Null
        }
        $lstDiscoveredDeploys.Visibility = 'Visible'
        $FuncCount = @($R.FuncApps).Count
        $StoreCount = @($R.StorageAccts).Count
        $lblDiscoverStatus.Text = "Found $FuncCount Function App(s), $StoreCount Storage Account(s)"
        Write-DebugLog "Discovered $FuncCount Function App(s), $StoreCount Storage Account(s)"
    }
})

# Double-click discovered item → load into deploy form + auto-pair related resources
$lstDiscoveredDeploys.Add_MouseDoubleClick({
    $Sel = $lstDiscoveredDeploys.SelectedItem
    if (-not $Sel -or -not $Sel.Tag) { return }
    $D = $Sel.Tag

    # Collect all discovered items for auto-pairing
    $AllItems = @()
    foreach ($Item in $lstDiscoveredDeploys.Items) {
        if ($Item.Tag) { $AllItems += $Item.Tag }
    }

    if ($D.Kind -eq 'FunctionApp') {
        # Strip 'func-' prefix to get source name
        $SourceName = $D.Name -replace '^func-', ''
        $txtDeployName.Text = $SourceName
        $txtDeployRG.Text = $D.ResourceGroup
        Set-ComboBoxByContent $cmbDeployRegion $D.Location
        $cmbBackupRestSource.Text = $D.Name

        # Auto-pair: find matching storage account (st<name> or contains source name, same RG)
        $Paired = $AllItems | Where-Object {
            $_.Kind -eq 'StorageAccount' -and (
                $_.Name -eq "st$SourceName" -or
                $_.Name -match [regex]::Escape($SourceName) -or
                $_.ResourceGroup -eq $D.ResourceGroup
            )
        } | Select-Object -First 1
        if ($Paired) {
            $cmbBackupStorageAcct.Text = $Paired.Name
            Show-Toast "Loaded '$($D.Name)' + paired storage '$($Paired.Name)'" -Type Success
        } else {
            Show-Toast "Loaded '$($D.Name)' into deploy form" -Type Success
        }
    } elseif ($D.Kind -eq 'StorageAccount') {
        $cmbBackupStorageAcct.Text = $D.Name

        # Auto-pair: find matching Function App (same RG or name contains storage name pattern)
        $StoreName = $D.Name -replace '^st', ''
        $Paired = $AllItems | Where-Object {
            $_.Kind -eq 'FunctionApp' -and (
                $_.Name -eq "func-$StoreName" -or
                $_.Name -match [regex]::Escape($StoreName) -or
                $_.ResourceGroup -eq $D.ResourceGroup
            )
        } | Select-Object -First 1
        if ($Paired) {
            $SourceName = $Paired.Name -replace '^func-', ''
            $txtDeployName.Text = $SourceName
            $txtDeployRG.Text = $Paired.ResourceGroup
            Set-ComboBoxByContent $cmbDeployRegion $Paired.Location
            $cmbBackupRestSource.Text = $Paired.Name
            Show-Toast "Loaded '$($D.Name)' + paired function '$($Paired.Name)'" -Type Success
        } else {
            Show-Toast "Loaded '$($D.Name)' into backup storage field" -Type Success
        }
    }
    Write-DebugLog "Loaded discovered resource: $($D.Name) ($($D.Kind))"
})

# ==============================================================================
# BACKUP / RESTORE
# ==============================================================================

$chkBackupRestSource   = $Window.FindName("chkBackupRestSource")
$chkBackupStorage      = $Window.FindName("chkBackupStorage")
$cmbBackupRestSource   = $Window.FindName("cmbBackupRestSource")
$cmbBackupStorageAcct  = $Window.FindName("cmbBackupStorageAcct")
$txtBackupFolder       = $Window.FindName("txtBackupFolder")
$btnBrowseBackupFolder = $Window.FindName("btnBrowseBackupFolder")
$btnRunBackup          = $Window.FindName("btnRunBackup")
$lblBackupStatus       = $Window.FindName("lblBackupStatus")
$prgBackup             = $Window.FindName("prgBackup")

$txtRestoreFolder         = $Window.FindName("txtRestoreFolder")
$btnBrowseRestoreFolder   = $Window.FindName("btnBrowseRestoreFolder")
$cmbRestoreRestTarget     = $Window.FindName("cmbRestoreRestTarget")
$cmbRestoreStorageTarget  = $Window.FindName("cmbRestoreStorageTarget")
$btnRunRestore            = $Window.FindName("btnRunRestore")
$lblRestoreStatus         = $Window.FindName("lblRestoreStatus")
$prgRestore               = $Window.FindName("prgRestore")

# Set default backup folder
$txtBackupFolder.Text = Join-Path $Global:Root "backups"

function Populate-BackupRestoreCombos {
    <# Populates the Backup and Restore REST source + Storage Account ComboBoxes from discovered data #>
    # Preserve current text values
    $BkRest  = $cmbBackupRestSource.Text
    $BkStor  = $cmbBackupStorageAcct.Text
    $RsRest  = $cmbRestoreRestTarget.Text
    $RsStor  = $cmbRestoreStorageTarget.Text

    # ── REST source combos: mirror global combo items ──
    foreach ($Cmb in @($cmbBackupRestSource, $cmbRestoreRestTarget)) {
        $Cmb.Items.Clear()
        foreach ($Item in $cmbRestSource.Items) {
            if ($Item.Tag) {
                $CI = New-Object System.Windows.Controls.ComboBoxItem
                $CI.Content = $Item.Tag
                $CI.Tag = $Item.Tag
                $Cmb.Items.Add($CI) | Out-Null
            }
        }
    }

    # ── Storage account combos: use discovered storage accounts ──
    $StorageList = if ($Global:StorageAccounts) { @($Global:StorageAccounts) } else { @() }
    foreach ($Cmb in @($cmbBackupStorageAcct, $cmbRestoreStorageTarget)) {
        $Cmb.Items.Clear()
        foreach ($SA in $StorageList) {
            $CI = New-Object System.Windows.Controls.ComboBoxItem
            $CI.Content = "$($SA.Name)"
            $CI.Tag = $SA.Name
            $Cmb.Items.Add($CI) | Out-Null
        }
    }

    # Restore previous text values (editable combo preserves typed text)
    if ($BkRest)  { $cmbBackupRestSource.Text = $BkRest }
    if ($BkStor)  { $cmbBackupStorageAcct.Text = $BkStor }
    if ($RsRest)  { $cmbRestoreRestTarget.Text = $RsRest }
    if ($RsStor)  { $cmbRestoreStorageTarget.Text = $RsStor }
}

# Auto-fill from settings and existing data
$ActiveSource = Get-ActiveRestSource
if ($ActiveSource) {
    $cmbBackupRestSource.Text = $ActiveSource
}
# Auto-fill storage account from the currently selected one in the main app
if ($cmbUploadStorage.SelectedItem -and $cmbUploadStorage.SelectedItem.Tag) {
    $cmbBackupStorageAcct.Text = $cmbUploadStorage.SelectedItem.Tag.Name
}
# Sync storage selection changes to backup field
$cmbUploadStorage.Add_SelectionChanged({
    if ($cmbUploadStorage.SelectedItem -and $cmbUploadStorage.SelectedItem.Tag -and -not $cmbBackupStorageAcct.Text.Trim()) {
        $cmbBackupStorageAcct.Text = $cmbUploadStorage.SelectedItem.Tag.Name
    }
}.GetNewClosure())

# Browse buttons
$btnBrowseBackupFolder.Add_Click({
    $Dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $Dlg.Description = "Select backup destination folder"
    if ($Dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtBackupFolder.Text = $Dlg.SelectedPath
    }
})

$btnBrowseRestoreFolder.Add_Click({
    $Dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $Dlg.Description = "Select backup folder to restore from"
    if ($Dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtRestoreFolder.Text = $Dlg.SelectedPath
        # Try to read metadata and populate target fields
        $MetaFile = Join-Path $Dlg.SelectedPath "backup_metadata.json"
        if (Test-Path $MetaFile) {
            try {
                $Meta = Get-Content $MetaFile -Raw | ConvertFrom-Json
                if ($Meta.RestSourceName) { $cmbRestoreRestTarget.Text = $Meta.RestSourceName }
                if ($Meta.StorageAccountName) { $cmbRestoreStorageTarget.Text = $Meta.StorageAccountName }
                Write-DebugLog "Loaded backup metadata: REST=$($Meta.RestSourceName), Storage=$($Meta.StorageAccountName)"
            } catch {
                Write-DebugLog "Could not read backup metadata: $($_.Exception.Message)" -Level 'WARN'
            }
        }
    }
})

# ── Backup Handler ────────────────────────────────────────────────────────────
$btnRunBackup.Add_Click({
    $DoRest    = [bool]$chkBackupRestSource.IsChecked
    $DoStorage = [bool]$chkBackupStorage.IsChecked

    if (-not $DoRest -and -not $DoStorage) {
        Show-Toast "Select at least one backup target" -Type Warning
        return
    }

    $RestSource = $cmbBackupRestSource.Text.Trim()
    $StorageAcct = $cmbBackupStorageAcct.Text.Trim()
    $BackupBase = $txtBackupFolder.Text.Trim()

    if ($DoRest -and -not $RestSource) {
        Show-Toast "Enter the REST Source Function name" -Type Warning
        return
    }
    if ($DoStorage -and -not $StorageAcct) {
        Show-Toast "Enter the Storage Account name" -Type Warning
        return
    }
    if (-not $BackupBase) {
        Show-Toast "Select a backup folder" -Type Warning
        return
    }

    # Create timestamped subfolder
    $Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $Label = if ($RestSource) { $RestSource -replace '^func-','' } elseif ($StorageAcct) { $StorageAcct } else { 'backup' }
    $BackupDir = Join-Path $BackupBase "${Label}_${Timestamp}"

    $lblBackupStatus.Text = "Backing up..."
    $prgBackup.Visibility = 'Visible'
    $prgBackup.IsIndeterminate = $true
    $btnRunBackup.IsEnabled = $false
    Show-GlobalProgressIndeterminate -Text 'Backing up REST source...'
    Write-DebugLog "Starting backup to: $BackupDir"

    if (-not $Global:BottomPanelOpen) { Toggle-BottomPanel $true }

    Start-BackgroundWork -Variables @{
        DoRest     = $DoRest
        DoStorage  = $DoStorage
        RestSrc    = $RestSource
        StorAcct   = $StorageAcct
        BkDir      = $BackupDir
    } -Work {
        $Stats = @{ ManifestCount = 0; BlobCount = 0; Errors = @() }
        New-Item -Path $BkDir -ItemType Directory -Force | Out-Null

        # ── Backup REST Source manifests ──
        if ($DoRest) {
            $ManDir = Join-Path $BkDir "manifests"
            New-Item -Path $ManDir -ItemType Directory -Force | Out-Null
            try {
                Import-Module Microsoft.WinGet.RestSource -ErrorAction Stop
                $Pkgs = @(Find-WinGetManifest -FunctionName $RestSrc -ErrorAction Stop)
                foreach ($P in $Pkgs) {
                    $PkgId = if ($P.PSObject.Properties['PackageIdentifier']) { $P.PackageIdentifier }
                             elseif ($P.PSObject.Properties['PackageIdentidier']) { $P.PackageIdentidier }
                             else { continue }
                    try {
                        $M = Get-WinGetManifest -FunctionName $RestSrc -PackageIdentifier $PkgId -ErrorAction Stop
                        $PkgDir = Join-Path $ManDir $PkgId
                        New-Item -Path $PkgDir -ItemType Directory -Force | Out-Null
                        $M | ConvertTo-Json -Depth 10 | Set-Content (Join-Path $PkgDir "manifest.json") -Force -Encoding UTF8
                        $Stats.ManifestCount++
                    } catch {
                        $Stats.Errors += "Manifest ${PkgId}: $($_.Exception.Message)"
                    }
                }
            } catch {
                $Stats.Errors += "REST Source query: $($_.Exception.Message)"
            }
        }

        # ── Backup Storage Account blobs (packages container only) ──
        if ($DoStorage) {
            $StorDir = Join-Path $BkDir "storage"
            New-Item -Path $StorDir -ItemType Directory -Force | Out-Null
            try {
                $Ctx = New-AzStorageContext -StorageAccountName $StorAcct -UseConnectedAccount -ErrorAction Stop
                $Containers = @(Get-AzStorageContainer -Context $Ctx -ErrorAction Stop |
                    Where-Object { $_.Name -eq 'packages' })
                foreach ($C in $Containers) {
                    $ContDir = Join-Path $StorDir $C.Name
                    New-Item -Path $ContDir -ItemType Directory -Force | Out-Null
                    $Blobs = @(Get-AzStorageBlob -Container $C.Name -Context $Ctx -ErrorAction Stop)
                    foreach ($B in $Blobs) {
                        try {
                            $BlobLocalPath = Join-Path $ContDir ($B.Name -replace '/', [IO.Path]::DirectorySeparatorChar)
                            $BlobDir = Split-Path $BlobLocalPath -Parent
                            if (-not (Test-Path $BlobDir)) {
                                New-Item -Path $BlobDir -ItemType Directory -Force | Out-Null
                            }
                            Get-AzStorageBlobContent -Container $C.Name -Blob $B.Name -Destination $BlobLocalPath -Context $Ctx -Force -ErrorAction Stop | Out-Null
                            $Stats.BlobCount++
                        } catch {
                            $Stats.Errors += "Blob $($C.Name)/$($B.Name): $($_.Exception.Message)"
                        }
                    }
                }
            } catch {
                $Stats.Errors += "Storage query: $($_.Exception.Message)"
            }
        }

        # ── Write backup metadata ──
        $Meta = @{
            Timestamp          = (Get-Date).ToString('o')
            RestSourceName     = if ($DoRest) { $RestSrc } else { $null }
            StorageAccountName = if ($DoStorage) { $StorAcct } else { $null }
            ManifestCount      = $Stats.ManifestCount
            BlobCount          = $Stats.BlobCount
            ErrorCount         = $Stats.Errors.Count
        }
        $Meta | ConvertTo-Json -Depth 3 | Set-Content (Join-Path $BkDir "backup_metadata.json") -Force -Encoding UTF8

        $Stats.BackupDir = $BkDir
        return $Stats
    } -OnComplete {
        param($Results, $Errors)
        try {
        $prgBackup.Visibility = 'Collapsed'
        $prgBackup.IsIndeterminate = $false
        $btnRunBackup.IsEnabled = $true
        Hide-GlobalProgress

        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ ManifestCount = 0; BlobCount = 0; Errors = @('No result') } }

        $Msg = "Backup complete: $($R.ManifestCount) manifest(s), $($R.BlobCount) blob(s)"
        if ($R.Errors -and $R.Errors.Count -gt 0) {
            $Msg += " ($($R.Errors.Count) error(s))"
            foreach ($E in $R.Errors) { Write-DebugLog "Backup error: $E" -Level 'ERROR' }
            # Show detailed error dialog
            $ErrList = ($R.Errors | Select-Object -First 10) -join "`n"
            if ($R.Errors.Count -gt 10) { $ErrList += "`n... and $($R.Errors.Count - 10) more" }
            Show-ThemedDialog -Title 'Backup Completed with Errors' `
                -Message "$Msg`n`n$ErrList" `
                -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' | Out-Null
        }
        $lblBackupStatus.Text = $Msg
        Write-DebugLog $Msg -Level $(if ($R.Errors.Count -gt 0) { 'WARN' } else { 'SUCCESS' })
        if ($R.BackupDir) { Write-DebugLog "Backup saved to: $($R.BackupDir)" }
        Show-Toast $Msg -Type $(if ($R.Errors.Count -gt 0) { 'Warning' } else { 'Success' })
        if (-not $R.Errors -or $R.Errors.Count -eq 0) { Unlock-Achievement 'first_backup' }
        } catch {
            Write-DebugLog "Backup OnComplete CRASH: $($_.Exception.Message)" -Level 'ERROR'
            $prgBackup.Visibility = 'Collapsed'
            $btnRunBackup.IsEnabled = $true
            $lblBackupStatus.Text = "Backup completed — check Activity Log for details"
            Show-Toast "Backup completed with errors — see Activity Log" -Type Warning
        }
    }
})

# ── Restore Handler ───────────────────────────────────────────────────────────
$btnRunRestore.Add_Click({
    $RestoreDir = $txtRestoreFolder.Text.Trim()
    if (-not $RestoreDir -or -not (Test-Path $RestoreDir)) {
        Show-Toast "Select a valid backup folder" -Type Warning
        return
    }

    $RestTarget = $cmbRestoreRestTarget.Text.Trim()
    $StorTarget = $cmbRestoreStorageTarget.Text.Trim()

    $ManifestDir = Join-Path $RestoreDir "manifests"
    $StorageDir  = Join-Path $RestoreDir "storage"
    $HasManifests = Test-Path $ManifestDir
    $HasStorage   = Test-Path $StorageDir

    if (-not $HasManifests -and -not $HasStorage) {
        Show-Toast "No manifests/ or storage/ subfolder found in the backup" -Type Warning
        return
    }

    if ($HasManifests -and -not $RestTarget) {
        Show-Toast "Enter a target REST Source Function name" -Type Warning
        return
    }
    if ($HasStorage -and -not $StorTarget) {
        Show-Toast "Enter a target Storage Account name" -Type Warning
        return
    }

    # Count what we'll restore for confirmation
    $ManCount = if ($HasManifests) { @(Get-ChildItem $ManifestDir -Filter "manifest.json" -Recurse).Count } else { 0 }
    $BlobCount = if ($HasStorage) { @(Get-ChildItem $StorageDir -File -Recurse).Count } else { 0 }

    $Confirm = Show-ThemedDialog -Title 'Confirm Restore' `
        -Message "Restore from backup?`n`nManifests: $ManCount → $RestTarget`nBlobs: $BlobCount → $StorTarget`n`nExisting packages with the same ID/version will be overwritten." `
        -Icon ([string]([char]0xE777)) -IconColor 'ThemeAccent' `
        -Buttons @(
            @{ Text='Cancel'; IsAccent=$false; Result='Cancel' }
            @{ Text='Restore'; IsAccent=$true;  Result='Restore' }
        )
    if ($Confirm -ne 'Restore') { return }

    $lblRestoreStatus.Text = "Restoring..."
    $prgRestore.Visibility = 'Visible'
    $prgRestore.IsIndeterminate = $true
    $btnRunRestore.IsEnabled = $false
    Write-DebugLog "Starting restore from: $RestoreDir"

    if (-not $Global:BottomPanelOpen) { Toggle-BottomPanel $true }

    Start-BackgroundWork -Variables @{
        ManDir      = $ManifestDir
        StorDir     = $StorageDir
        RestTgt     = $RestTarget
        StorTgt     = $StorTarget
        HasMan      = $HasManifests
        HasStor     = $HasStorage
    } -Work {
        $Stats = @{ ManifestsRestored = 0; BlobsRestored = 0; Errors = @() }

        # ── Restore manifests to REST source ──
        if ($HasMan -and $RestTgt) {
            try {
                Import-Module Microsoft.WinGet.RestSource -ErrorAction Stop
                $ManifestFiles = @(Get-ChildItem $ManDir -Filter "manifest.json" -Recurse)
                foreach ($MFile in $ManifestFiles) {
                    try {
                        Add-WinGetManifest -FunctionName $RestTgt -Path $MFile.DirectoryName -ErrorAction Stop | Out-Null
                        $Stats.ManifestsRestored++
                    } catch {
                        $Stats.Errors += "Restore manifest ($($MFile.DirectoryName)): $($_.Exception.Message)"
                    }
                }
            } catch {
                $Stats.Errors += "REST module: $($_.Exception.Message)"
            }
        }

        # ── Restore blobs to storage account ──
        if ($HasStor -and $StorTgt) {
            try {
                $Ctx = New-AzStorageContext -StorageAccountName $StorTgt -UseConnectedAccount -ErrorAction Stop
                $Containers = @(Get-ChildItem $StorDir -Directory)
                foreach ($ContFolder in $Containers) {
                    $ContName = $ContFolder.Name
                    # Ensure container exists
                    try { New-AzStorageContainer -Name $ContName -Context $Ctx -Permission Off -ErrorAction Stop | Out-Null } catch { <# Container already exists — safe to ignore #> }
                    $Files = @(Get-ChildItem $ContFolder.FullName -File -Recurse)
                    foreach ($F in $Files) {
                        try {
                            $RelPath = $F.FullName.Substring($ContFolder.FullName.Length + 1) -replace '\\', '/'
                            Set-AzStorageBlobContent -File $F.FullName -Container $ContName -Blob $RelPath -Context $Ctx -Force -ErrorAction Stop | Out-Null
                            $Stats.BlobsRestored++
                        } catch {
                            $Stats.Errors += "Restore blob $ContName/$($F.Name): $($_.Exception.Message)"
                        }
                    }
                }
            } catch {
                $Stats.Errors += "Storage context: $($_.Exception.Message)"
            }
        }

        return $Stats
    } -OnComplete {
        param($Results, $Errors)
        try {
        $prgRestore.Visibility = 'Collapsed'
        $prgRestore.IsIndeterminate = $false
        $btnRunRestore.IsEnabled = $true
        Hide-GlobalProgress

        $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ ManifestsRestored = 0; BlobsRestored = 0; Errors = @('No result') } }

        $Msg = "Restore complete: $($R.ManifestsRestored) manifest(s), $($R.BlobsRestored) blob(s)"
        if ($R.Errors -and $R.Errors.Count -gt 0) {
            $Msg += " ($($R.Errors.Count) error(s))"
            foreach ($E in $R.Errors) { Write-DebugLog "Restore error: $E" -Level 'ERROR' }
            # Show detailed error dialog
            $ErrList = ($R.Errors | Select-Object -First 10) -join "`n"
            if ($R.Errors.Count -gt 10) { $ErrList += "`n... and $($R.Errors.Count - 10) more" }
            Show-ThemedDialog -Title 'Restore Completed with Errors' `
                -Message "$Msg`n`n$ErrList" `
                -Icon ([string]([char]0xE7BA)) -IconColor 'ThemeWarning' | Out-Null
        }
        $lblRestoreStatus.Text = $Msg
        Write-DebugLog $Msg -Level $(if ($R.Errors.Count -gt 0) { 'WARN' } else { 'SUCCESS' })
        Show-Toast $Msg -Type $(if ($R.Errors.Count -gt 0) { 'Warning' } else { 'Success' })
        if (-not $R.Errors -or $R.Errors.Count -eq 0) { Unlock-Achievement 'first_restore' }
        } catch {
            Write-DebugLog "Restore OnComplete CRASH: $($_.Exception.Message)" -Level 'ERROR'
            $prgRestore.Visibility = 'Collapsed'
            $btnRunRestore.IsEnabled = $true
            $lblRestoreStatus.Text = "Restore completed — check Activity Log for details"
            Show-Toast "Restore completed with errors — see Activity Log" -Type Warning
        }
    }
})
$btnClearLog.Add_Click({
    $Global:DebugSB.Clear()
    $Global:DebugLineCount = 0
    if ($paraLog) { $paraLog.Inlines.Clear() }
    # Clear minimap + heatmap
    $Global:MinimapDots.Clear()
    $Global:MinimapLineTotal = 0
    $Global:MinimapLastRenderedIdx = 0
    $Global:HeatmapData.Clear()
    $Global:HeatmapLastRenderedIdx = 0
    $cnvMM = $Window.FindName('cnvMinimap'); if ($cnvMM) { $cnvMM.Children.Clear() }
    $cnvHM = $Window.FindName('cnvHeatmap'); if ($cnvHM) { $cnvHM.Children.Clear() }
    Write-DebugLog "Activity log cleared"
})

# ── Minimap click-to-jump ────────────────────────────────────────────────────
$cnvMinimapCtrl = $Window.FindName('cnvMinimap')
if ($cnvMinimapCtrl) {
    $cnvMinimapCtrl.Add_MouseLeftButtonDown({
        param($sender, $e)
        $Pos = $e.GetPosition($sender)
        $H = $sender.ActualHeight
        if ($H -gt 0 -and $logScroller.ExtentHeight -gt 0) {
            $Ratio = $Pos.Y / $H
            $logScroller.ScrollToVerticalOffset($Ratio * $logScroller.ExtentHeight)
        }
    })
}

# ── Heatmap click-to-jump ────────────────────────────────────────────────────
$cnvHeatmapCtrl = $Window.FindName('cnvHeatmap')
if ($cnvHeatmapCtrl) {
    $cnvHeatmapCtrl.Add_MouseLeftButtonDown({
        param($sender, $e)
        $Pos = $e.GetPosition($sender)
        $W = $sender.ActualWidth
        if ($W -gt 0 -and $logScroller.ExtentHeight -gt 0) {
            $Ratio = $Pos.X / $W
            $logScroller.ScrollToVerticalOffset($Ratio * $logScroller.ExtentHeight)
        }
    })
}

# ── Scroll sync — update minimap/heatmap viewport on scroll ──────────────────────
$logScroller.Add_ScrollChanged({
    Render-Minimap
    Render-Heatmap
})

# ── New Config ────────────────────────────────────────────────────────────────
$btnNewConfig.Add_Click({
    # Confirm if form has data to avoid accidental loss
    $HasData = $txtPkgId.Text.Trim() -or $txtPkgName.Text.Trim() -or $txtPkgVersion.Text.Trim()
    if ($HasData) {
        $Confirm = Show-ThemedDialog `
            -Title 'New Manifest' `
            -Message "Clear the current form and start a new manifest?`n`nAny unsaved changes will be lost." `
            -Icon ([string]([char]0xE7C3)) `
            -Buttons @(
                @{ Text='Cancel'; IsAccent=$false; Result='No' }
                @{ Text='Clear Form'; IsAccent=$true; Result='Yes' }
            )
        if ($Confirm -ne 'Yes') { return }
    }
    # Clear all form fields
    $txtPkgId.Text = ""
    $txtPkgVersion.Text = ""
    $txtPkgName.Text = ""
    $txtPublisher.Text = ""
    $txtLicense.Text = ""
    $txtShortDesc.Text = ""
    $txtDescription.Text = ""
    $txtAuthor.Text = ""
    $txtMoniker.Text = ""
    $txtPackageUrl.Text = ""
    $txtPublisherUrl.Text = ""
    $txtLicenseUrl.Text = ""
    $txtSupportUrl.Text = ""
    $txtCopyright.Text = ""
    $txtPrivacyUrl.Text = ""
    $txtTags.Text = ""
    $txtReleaseNotes.Text = ""
    $txtReleaseNotesUrl.Text = ""
    $txtLocalFile0.Text = ""
    $lblHash0.Text = ""
    $txtInstallerUrl0.Text = ""
    $txtSilent0.Text = ""
    $txtSilentProgress0.Text = ""
    $txtProductCode0.Text = ""
    $txtCommands0.Text = ""
    $txtFileExt0.Text = ""
    $txtCustomSwitch0.Text = ""
    $txtInteractive0.Text = ""
    $txtLog0.Text = ""
    $txtRepair0.Text = ""
    $txtNestedPath0.Text = ""
    $txtPortableAlias0.Text = ""
    $chkModeInteractive.IsChecked = $true
    $chkModeSilent.IsChecked = $true
    $chkModeSilentProgress.IsChecked = $false
    $chkInstallLocationRequired.IsChecked = $false
    $chkPlatformDesktop.IsChecked = $true
    $chkPlatformUniversal.IsChecked = $false
    $txtMinOSVersion0.Text = ""
    $txtExpectedReturnCodes0.Text = ""
    $txtAFDisplayName0.Text = ""
    $txtAFPublisher0.Text = ""
    $txtAFDisplayVersion0.Text = ""
    $txtAFProductCode0.Text = ""
    $Global:InstallerEntries.Clear()
    Update-InstallerSummaryList
    $cmbInstallerPreset.SelectedIndex = 0
    $cmbNestedType0.SelectedIndex = 0
    $pnlNestedInstaller.Visibility = 'Collapsed'
    $Global:CurrentHash0 = ""
    $txtYamlPreview.Text = ""
    # Clear upload-related fields
    $txtBlobPath.Text = ""
    $txtUploadFile.Text = ""
    $lblUploadHash.Text = "Select a file to compute hash..."
    $lblUploadProgress.Text = ""
    Switch-Tab 'Create'
    $lblStatus.Text = "New manifest — fill in the fields"
    Write-DebugLog "New config form cleared"
})

$btnRefreshConfigs.Add_Click({ Refresh-ConfigList })
$btnRefreshStorage.Add_Click({
    if ($Global:IsAuthenticated) { Discover-StorageAccounts }
    else { $lblStatus.Text = "Sign in first to discover storage accounts" }
})

# ── New Storage Account ──────────────────────────────────────────────────────
$btnNewStorage.Add_Click({
    if (-not $Global:IsAuthenticated) {
        $lblStatus.Text = "Sign in first to create a storage account"
        return
    }
    Show-NewStoragePanel
})
$btnCloseNewStorage.Add_Click({ Hide-NewStoragePanel })
$btnCancelNewStorage.Add_Click({ Hide-NewStoragePanel })
$btnRefreshRGs.Add_Click({ Refresh-ResourceGroups })
$btnCreateStorage.Add_Click({ New-StorageAccountFromUI })

# Real-time storage name validation hint
$txtNewStorageName.Add_TextChanged({
    $Name = $txtNewStorageName.Text.Trim().ToLower()
    if (-not $Name) {
        $lblStorageNameHint.Text = "3-24 characters, lowercase letters and numbers only"
        $lblStorageNameHint.Foreground = $Window.Resources['ThemeTextDisabled']
    } elseif ($Name.Length -lt 3) {
        $lblStorageNameHint.Text = "Too short ($($Name.Length)/3 min)"
        $lblStorageNameHint.Foreground = $Window.Resources['ThemeWarning']
    } elseif ($Name -notmatch '^[a-z0-9]+$') {
        $lblStorageNameHint.Text = "Only lowercase letters and numbers allowed"
        $lblStorageNameHint.Foreground = $Window.Resources['ThemeError']
    } else {
        $lblStorageNameHint.Text = "$($Name.Length)/24 characters"
        $lblStorageNameHint.Foreground = $Window.Resources['ThemeSuccess']
    }
})

# ── Help Button ─────────────────────────────────────────────────────────────────
$btnHelp = $Window.FindName("btnHelp")
if ($btnHelp) {
    $btnHelp.Add_Click({
        $HelpText = @"
KEYBOARD SHORTCUTS
──────────────────────────────────────────────────
Ctrl+S          Save manifest to disk
Ctrl+Shift+S    Save as config
Ctrl+Shift+P    Publish to REST source
Ctrl+N          New manifest (clear form)
Ctrl+Shift+C    Copy YAML to clipboard
F5              Refresh YAML preview
Ctrl+1..7       Switch tabs (Create, Edit, Upload,
                Manage, Deploy, Remove, Backup)

TABS
──────────────────────────────────────────────────
Create      Build WinGet manifests with live
            YAML preview
Edit        Search and edit packages on
            connected REST source
Upload      Upload .yaml files to Azure
            Blob Storage
Manage      Browse packages, view versions,
            remove entries
Deploy      Deploy or update Azure REST source
            infrastructure
Remove      Tear down deployed REST source
            resources
Backup      Backup/restore REST source
            manifests + blobs

TIPS
──────────────────────────────────────────────────
• Cache refreshes automatically every 5 min
• Toggle Debug overlay for verbose logging
• Minimap (right) + Heatmap (top) show log
  activity at a glance
• Click either strip to jump to that position
"@
        # Build a custom dialog with monospace font for the help text
        $Palette = if ($Global:IsLightMode) { $Global:ThemeLight } else { $Global:ThemeDark }
        $Br = { param([string]$Key) [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($Palette[$Key])) }

        # Size relative to main window
        $DlgHeight = [Math]::Max(400, $Window.ActualHeight - 120)
        $DlgWidth  = 540

        $Dlg = New-Object System.Windows.Window
        $Dlg.Title = 'Help — WinGet Manifest Manager'
        $Dlg.Width = $DlgWidth; $Dlg.Height = $DlgHeight
        $Dlg.ResizeMode = 'NoResize'
        $Dlg.WindowStartupLocation = 'CenterOwner'
        $Dlg.Owner = $Window
        $Dlg.Background = (& $Br 'ThemeCardBg')
        $Dlg.WindowStyle = 'None'
        $Dlg.AllowsTransparency = $true

        $OuterBorder = New-Object System.Windows.Controls.Border
        $OuterBorder.Background = (& $Br 'ThemeCardBg')
        $OuterBorder.BorderBrush = (& $Br 'ThemeBorderElevated')
        $OuterBorder.BorderThickness = [System.Windows.Thickness]::new(1)
        $OuterBorder.CornerRadius = [System.Windows.CornerRadius]::new(12)
        $OuterBorder.Padding = [System.Windows.Thickness]::new(28, 24, 28, 20)
        $OuterBorder.Effect = [System.Windows.Media.Effects.DropShadowEffect]@{
            Color = [System.Windows.Media.Colors]::Black; Direction = 270; ShadowDepth = 8; BlurRadius = 28; Opacity = 0.45
        }

        # Use a DockPanel so ScrollViewer fills remaining space
        $Root = New-Object System.Windows.Controls.DockPanel
        $Root.LastChildFill = $true

        # Header
        $HeaderGrid = New-Object System.Windows.Controls.Grid
        $HeaderGrid.Margin = [System.Windows.Thickness]::new(0, 0, 0, 16)
        $Col1 = New-Object System.Windows.Controls.ColumnDefinition; $Col1.Width = [System.Windows.GridLength]::new(0, [System.Windows.GridUnitType]::Auto)
        $Col2 = New-Object System.Windows.Controls.ColumnDefinition; $Col2.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $HeaderGrid.ColumnDefinitions.Add($Col1) | Out-Null; $HeaderGrid.ColumnDefinitions.Add($Col2) | Out-Null

        $IconBadge = New-Object System.Windows.Controls.Border
        $IconBadge.Width = 36; $IconBadge.Height = 36; $IconBadge.CornerRadius = [System.Windows.CornerRadius]::new(8)
        $IconBadge.Background = (& $Br 'ThemeAccentDim'); $IconBadge.Margin = [System.Windows.Thickness]::new(0, 0, 14, 0)
        $IconTB = New-Object System.Windows.Controls.TextBlock
        $IconTB.Text = '?'; $IconTB.FontSize = 17; $IconTB.FontWeight = 'Bold'
        $IconTB.Foreground = (& $Br 'ThemeAccentLight'); $IconTB.HorizontalAlignment = 'Center'; $IconTB.VerticalAlignment = 'Center'
        $IconBadge.Child = $IconTB
        [System.Windows.Controls.Grid]::SetColumn($IconBadge, 0); $HeaderGrid.Children.Add($IconBadge) | Out-Null

        $TitleTB = New-Object System.Windows.Controls.TextBlock
        $TitleTB.Text = 'Help — WinGet Manifest Manager'; $TitleTB.FontSize = 15; $TitleTB.FontWeight = 'Bold'
        $TitleTB.Foreground = (& $Br 'ThemeTextPrimary'); $TitleTB.VerticalAlignment = 'Center'
        [System.Windows.Controls.Grid]::SetColumn($TitleTB, 1); $HeaderGrid.Children.Add($TitleTB) | Out-Null
        [System.Windows.Controls.DockPanel]::SetDock($HeaderGrid, [System.Windows.Controls.Dock]::Top)
        $Root.Children.Add($HeaderGrid) | Out-Null

        $Sep = New-Object System.Windows.Controls.Border; $Sep.Height = 1
        $Sep.Background = (& $Br 'ThemeBorder'); $Sep.Margin = [System.Windows.Thickness]::new(0, 0, 0, 14)
        [System.Windows.Controls.DockPanel]::SetDock($Sep, [System.Windows.Controls.Dock]::Top)
        $Root.Children.Add($Sep) | Out-Null

        # Close button (docked at bottom)
        $BtnPanel = New-Object System.Windows.Controls.StackPanel
        $BtnPanel.Orientation = 'Horizontal'; $BtnPanel.HorizontalAlignment = 'Right'
        $BtnPanel.Margin = [System.Windows.Thickness]::new(0, 14, 0, 0)
        $CloseBtnXaml = @"
<Button xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" MinWidth="80" Padding="16,8" FontSize="12" Cursor="Hand" Content="Close">
    <Button.Template>
        <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="$($Palette['ThemeCardAltBg'])" BorderBrush="$($Palette['ThemeBorder'])" BorderThickness="1" CornerRadius="6" Padding="{TemplateBinding Padding}">
                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter TargetName="bd" Property="Background" Value="$($Palette['ThemeHoverBg'])"/>
                </Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate>
    </Button.Template>
</Button>
"@
        $CloseBtn = [System.Windows.Markup.XamlReader]::Parse($CloseBtnXaml)
        $CloseBtn.Foreground = (& $Br 'ThemeTextBody')
        $CloseBtn.Add_Click({ $Dlg.Close() })
        $BtnPanel.Children.Add($CloseBtn) | Out-Null
        [System.Windows.Controls.DockPanel]::SetDock($BtnPanel, [System.Windows.Controls.Dock]::Bottom)
        $Root.Children.Add($BtnPanel) | Out-Null

        # Styled ScrollViewer with themed scrollbar (fills remaining space)
        $ThumbColor = $Palette['ThemeScrollThumb']
        $TrackColor = $Palette['ThemeScrollTrack']
        $AccentColor = $Palette['ThemeAccentLight']
        $PrimaryColor = $Palette['ThemeTextPrimary']
        $SvXaml = @"
<ScrollViewer xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
              xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
              VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled"
              Focusable="False">
    <ScrollViewer.Resources>
        <Style TargetType="ScrollBar">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Width" Value="8"/>
            <Setter Property="MinWidth" Value="8"/>
            <Setter Property="Opacity" Value="0.8"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollBar">
                        <Grid Background="$TrackColor">
                            <Track x:Name="PART_Track" IsDirectionReversed="True" Focusable="False">
                                <Track.DecreaseRepeatButton>
                                    <RepeatButton Command="ScrollBar.PageUpCommand" Opacity="0" Focusable="False"/>
                                </Track.DecreaseRepeatButton>
                                <Track.Thumb>
                                    <Thumb>
                                        <Thumb.Template>
                                            <ControlTemplate TargetType="Thumb">
                                                <Border x:Name="ThumbBorder" Background="$ThumbColor" CornerRadius="4" MinHeight="40" Margin="1,0"/>
                                                <ControlTemplate.Triggers>
                                                    <Trigger Property="IsMouseOver" Value="True">
                                                        <Setter TargetName="ThumbBorder" Property="Background" Value="$AccentColor"/>
                                                    </Trigger>
                                                    <Trigger Property="IsDragging" Value="True">
                                                        <Setter TargetName="ThumbBorder" Property="Background" Value="$PrimaryColor"/>
                                                    </Trigger>
                                                </ControlTemplate.Triggers>
                                            </ControlTemplate>
                                        </Thumb.Template>
                                    </Thumb>
                                </Track.Thumb>
                                <Track.IncreaseRepeatButton>
                                    <RepeatButton Command="ScrollBar.PageDownCommand" Opacity="0" Focusable="False"/>
                                </Track.IncreaseRepeatButton>
                            </Track>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Opacity" Value="1.0"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </ScrollViewer.Resources>
</ScrollViewer>
"@
        $SV = [System.Windows.Markup.XamlReader]::Parse($SvXaml)
        $MsgTB = New-Object System.Windows.Controls.TextBlock
        $MsgTB.Text = $HelpText; $MsgTB.FontSize = 11.5
        $MsgTB.FontFamily = [System.Windows.Media.FontFamily]::new('Cascadia Code, Cascadia Mono, Consolas, Courier New')
        $MsgTB.Foreground = (& $Br 'ThemeTextSecondary'); $MsgTB.TextWrapping = 'NoWrap'
        $MsgTB.LineHeight = 18
        $SV.Content = $MsgTB
        $Root.Children.Add($SV) | Out-Null

        $OuterBorder.Child = $Root; $Dlg.Content = $OuterBorder
        $Dlg.ShowDialog() | Out-Null
    })
}

# ── Keyboard Shortcuts ─────────────────────────────────────────────────────────
$Window.Add_PreviewKeyDown({
    param($sender, $e)
    $ctrl  = [System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control
    $shift = [System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Shift
    # Ctrl+S — Save manifest to disk
    if ($ctrl -and -not $shift -and $e.Key -eq 'S') {
        $e.Handled = $true
        Unlock-Achievement 'shortcut_user'
        Save-ManifestToDisk
    }
    # Ctrl+Shift+S — Save as config
    if ($ctrl -and $shift -and $e.Key -eq 'S') {
        $e.Handled = $true
        Unlock-Achievement 'shortcut_user'
        Save-PackageConfig
    }
    # Ctrl+Shift+P — Publish
    if ($ctrl -and $shift -and $e.Key -eq 'P') {
        $e.Handled = $true
        Unlock-Achievement 'shortcut_user'
        $btnPublish.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
    }
    # Ctrl+N — New manifest
    if ($ctrl -and -not $shift -and $e.Key -eq 'N') {
        $e.Handled = $true
        $btnNewConfig.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
    }
    # Ctrl+Shift+C — Copy YAML to clipboard
    if ($ctrl -and $shift -and $e.Key -eq 'C') {
        $e.Handled = $true
        $btnCopyYaml.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
    }
    # F5 — Refresh (context-dependent)
    if ($e.Key -eq 'F5') {
        $e.Handled = $true
        Update-YamlPreview
    }
    # Ctrl+1..7 — Switch tabs
    if ($ctrl -and -not $shift) {
        switch ($e.Key) {
            'D1' { $e.Handled = $true; Switch-Tab 'Create' }
            'D2' { $e.Handled = $true; Switch-Tab 'Edit' }
            'D3' { $e.Handled = $true; Switch-Tab 'Upload' }
            'D4' { $e.Handled = $true; Switch-Tab 'Manage' }
            'D5' { $e.Handled = $true; Switch-Tab 'Deploy' }
            'D6' { $e.Handled = $true; Switch-Tab 'Remove' }
            'D7' { $e.Handled = $true; Switch-Tab 'Backup' }
        }
    }
})

# ── Window Closing ────────────────────────────────────────────────────────────
$Window.Add_Closing({
    if ($chkAutoSave.IsChecked) { Save-UserPrefs }
})

# ==============================================================================
# SECTION 14: STARTUP SEQUENCE
# ==============================================================================

Write-DebugLog "=========================================" -Level 'DEBUG'
Write-DebugLog "$Global:AppTitle starting..." -Level 'DEBUG'
Write-DebugLog "=========================================" -Level 'DEBUG'

# 1. Check prerequisites
$PrereqResults = Test-Prerequisites
Show-PrereqResults -Results $PrereqResults

# 2. If prerequisites are met, try to validate cached Azure context
if ($Global:PrereqsPassed) {
    $CachedCtx = Get-AzContext -ErrorAction SilentlyContinue
    if ($CachedCtx -and $CachedCtx.Account) {
        $lblAuthStatus.Text = "Validating token for $($CachedCtx.Account.Id)..."
        Write-DebugLog "Startup: cached context found for $($CachedCtx.Account.Id), validating in background..."
        Start-BackgroundWork -Variables @{
            AccountId = $CachedCtx.Account.Id
            SubName   = $CachedCtx.Subscription.Name
            SubId     = $CachedCtx.Subscription.Id
        } -Work {
            $R = @{ TokenValid = $false; AccountId = $AccountId; SubName = $SubName; SubId = $SubId; Subs = @(); Error = '' }
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
            $R = if ($Results -and $Results.Count -gt 0) { $Results[0] } else { @{ TokenValid = $false; Error = 'No result' } }
            if ($R.TokenValid) {
                Set-AuthUIConnected -AccountId $R.AccountId -SubName $R.SubName -SubId $R.SubId -Subscriptions $R.Subs
                $lblStatus.Text = "Logged in as $($R.AccountId)"
                Discover-StorageAccounts
                Discover-RestSources
                Refresh-ResourceGroups
            } else {
                Set-AuthUIDisconnected -Reason "Token expired for $($R.AccountId)"
                $lblStatus.Text = "Cached token expired — click Sign In to re-authenticate"
            }
        }
    }
} else {
    $lblStatus.Text = "Missing required modules — see error panel"
}

# 3. Start the dispatcher timer
$Timer.Start()

# 3a. Clean up stale temp files from previous sessions
Clear-TempFiles

# 4. Refresh config list
Refresh-ConfigList

# 5. Load user preferences (applies saved theme, toggles, window position)
Load-UserPrefs

# 6. Default to dark mode if no prefs file exists
if (-not (Test-Path $Global:PrefsPath)) {
    $Global:IsLightMode = $false
    & $ApplyTheme $false
    if ($btnThemeToggle) { $btnThemeToggle.Content = [char]0x2600 }
    if ($chkDarkMode)    { $chkDarkMode.IsChecked = $true }
}

# 7. Generate initial YAML preview
Update-YamlPreview

# 8. Load gamification state
Load-Achievements
Load-PublishStreak
Render-Achievements
Update-StreakDisplay
$null = Update-QualityMeter

# 9. Trap unhandled WPF dispatcher exceptions for diagnostics
$Window.Dispatcher.Add_UnhandledException({
    param($sender, $e)
    $Ex = $e.Exception
    $Inner = if ($Ex.InnerException) { $Ex.InnerException } else { $Ex }
    Write-DebugLog "DISPATCHER UNHANDLED: $($Inner.GetType().FullName): $($Inner.Message)" -Level 'ERROR'
    Write-DebugLog "  .NET Stack: $($Inner.StackTrace)" -Level 'ERROR'
    if ($Inner -ne $Ex) {
        Write-DebugLog "  Outer: $($Ex.GetType().FullName): $($Ex.Message)" -Level 'ERROR'
    }
    # Let the app crash so the user sees it — but the log now has the source
    $e.Handled = $false
})

# Bring window to front on launch, then release Topmost so it doesn't
# permanently cover other apps.
$Window.Topmost = $true
$Window.Add_ContentRendered({
    $Window.Topmost = $false
    $Window.Activate()
})

# Show window (blocks until closed)
$Window.ShowDialog() | Out-Null

# 10. Cleanup on exit
$Timer.Stop()
for ($i = $Global:BgJobs.Count - 1; $i -ge 0; $i--) {
    try { $Global:BgJobs[$i].PS.Dispose() } catch { Write-DebugLog "Exit cleanup PS[$i]: $($_.Exception.Message)" -Level 'DEBUG' }
    try { $Global:BgJobs[$i].Runspace.Dispose() } catch { Write-DebugLog "Exit cleanup Runspace[$i]: $($_.Exception.Message)" -Level 'DEBUG' }
}
