#Requires -Version 5.1
<#
.SYNOPSIS
    Windows 365 Discovery — Automated Cloud PC environment inventory and assessment.
.DESCRIPTION
    Standalone discovery script that connects to Microsoft Graph (beta), enumerates
    Windows 365 Cloud PC resources (Cloud PCs, provisioning policies, user settings,
    Azure Network Connections, device/gallery images, service plans, audit events),
    runs a small set of automated checks against published Windows 365 best practices,
    and exports a portable JSON file for import into the Windows 365 Assessor GUI.

    All advanced/manual checks are intentionally left for the GUI. This script is a
    data collector + lightweight rule engine.
.PARAMETER OutputPath
    Path to save the discovery JSON file. Defaults to
    .\assessments\discovery_<timestamp>.json relative to the script.
.PARAMETER TenantId
    Optional Entra ID tenant ID. If omitted, uses the user's home tenant.
.PARAMETER SkipLogin
    Skip interactive login and use existing Microsoft Graph context.
.PARAMETER InactiveDays
    Threshold in days to flag Cloud PCs as inactive. Default: 30.
.PARAMETER ImageAgeWarnDays
    Threshold in days to warn on stale custom images. Default: 90.
.EXAMPLE
    .\Invoke-W365Discovery.ps1
    .\Invoke-W365Discovery.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    .\Invoke-W365Discovery.ps1 -OutputPath "C:\temp\w365_discovery.json"
.NOTES
    Author : Anton Romanyuk
    Version: 0.2.0
    Date   : 2026-07-18

    Required Graph scopes (core tier — requested unconditionally):
      CloudPC.Read.All                          Cloud PCs, provisioning/user policies, ANCs, images, reports
      DeviceManagementConfiguration.Read.All    Intune config/compliance/baseline profiles, Endpoint Analytics
      DeviceManagementManagedDevices.Read.All   Intune managed devices (Cloud PC Defender/compliance health)
      Directory.Read.All                        Tenant/account context

    Optional Graph scope (requested if the admin consents; checks that need it
    degrade to a Status 'Error' CheckResult naming the missing scope when absent):
      Policy.Read.All                           Conditional Access policies targeting Cloud PC sign-in apps

    API versions:
      /v1.0/deviceManagement/virtualEndpoint    cloudPCs, provisioningPolicies, userSettings (GA)
      /beta/deviceManagement/virtualEndpoint    onPremisesConnections, deviceImages, galleryImages,
                                                servicePlans, auditEvents, reports/* (reports = beta-only)
      /beta/deviceManagement                    Intune managedDevices/compliance/config/EA/update summary
      /v1.0/identity/conditionalAccess/policies Conditional Access (GA)

    Required PowerShell modules:
      Microsoft.Graph.Authentication (>= 2.0.0)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [switch]$SkipLogin,

    [Parameter(Mandatory = $false)]
    [int]$InactiveDays = 30,

    [Parameter(Mandatory = $false)]
    [int]$ImageAgeWarnDays = 90
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Strip OneDrive module path to prevent old module version conflicts
$env:PSModulePath = ($env:PSModulePath -split ';' |
    Where-Object { $_ -notlike '*OneDrive*' }) -join ';'

$ScriptRoot = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($ScriptRoot)) { $ScriptRoot = $PWD.Path }

$ScriptVersion = '0.2.0'
# Windows 365 GA surface (cloudPCs, provisioningPolicies, userSettings) migrated to /v1.0.
$GraphBaseV1   = 'https://graph.microsoft.com/v1.0/deviceManagement/virtualEndpoint'
# Beta retained for endpoints not yet GA / verified beta-only: onPremisesConnections, deviceImages,
# galleryImages, servicePlans, auditEvents, and the whole reports/* surface (cloudPcReports is beta-only).
$GraphBase     = 'https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint'
$ReportsBase   = "$GraphBase/reports"
$IntuneBase    = 'https://graph.microsoft.com/beta/deviceManagement'
$CaPolicyUri   = 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies'

# First-party application IDs used to sign in to / broker Cloud PC + AVD sessions. Conditional Access
# policies should target these to protect Cloud PC access.
#   Windows Cloud Login  (formerly "Microsoft Remote Desktop" / brokers Cloud PC + AVD SSO)
#   Azure Virtual Desktop (session host / gateway app, also brokers Cloud PC connections)
#   Windows 365 (the Windows 365 web portal / provisioning app)
# NOTE: Windows Cloud Login and Azure Virtual Desktop IDs are well-known Microsoft first-party IDs.
# The Windows 365 portal ID below is provided per the assessment spec and flagged for doc verification.
$AppIdWindowsCloudLogin  = '270efc09-cd0d-444b-a71f-39af4910ec45'
$AppIdAzureVirtualDesktop = '9cdead84-a844-4324-93f2-b2e6bb768d07'
$AppIdWindows365Portal   = '0af06dc6-e4b5-4f28-818e-e78e62d137a5'
$CloudPcSignInAppIds     = @($AppIdWindowsCloudLogin, $AppIdAzureVirtualDesktop, $AppIdWindows365Portal)

# ═══════════════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════════════

function Write-Status {
    param([string]$Message, [string]$Level = 'INFO')
    $ts = (Get-Date).ToString('HH:mm:ss')
    $Icon = switch ($Level) {
        'ERROR'   { 'X' }
        'WARN'    { '!' }
        'SUCCESS' { '+' }
        'CHECK'   { '>' }
        'SECTION' { '-' }
        default   { '.' }
    }
    $Color = switch ($Level) {
        'ERROR'   { 'Red' }
        'WARN'    { 'Yellow' }
        'SUCCESS' { 'Green' }
        'CHECK'   { 'Cyan' }
        'SECTION' { 'DarkCyan' }
        default   { 'Gray' }
    }
    if ($Level -eq 'SECTION') {
        Write-Host ""
        Write-Host "  --- " -NoNewline -ForegroundColor DarkCyan
        Write-Host $Message -NoNewline -ForegroundColor Cyan
        Write-Host (" " + ('-' * [math]::Max(1, 48 - $Message.Length))) -ForegroundColor DarkCyan
    } else {
        Write-Host "  " -NoNewline
        Write-Host $Icon -NoNewline -ForegroundColor $Color
        Write-Host " " -NoNewline
        Write-Host "[$ts] " -NoNewline -ForegroundColor DarkGray
        Write-Host $Message -ForegroundColor $(if ($Level -eq 'INFO') { 'White' } else { $Color })
    }
}

function Write-Metric {
    param([string]$Label, [int]$Value, [string]$Icon = '|')
    Write-Host "  $Icon  " -NoNewline -ForegroundColor DarkCyan
    Write-Host $Label.PadRight(28) -NoNewline -ForegroundColor Gray
    Write-Host $Value -ForegroundColor White
}

function New-CheckResult {
    param(
        [string]$Id,
        [string]$Category,
        [string]$Name,
        [string]$Description,
        [ValidateSet('Pass','Fail','Warning','N/A','Error')][string]$Status,
        [string]$Severity = 'Medium',
        [string]$Details = '',
        [string]$Recommendation = '',
        [string]$Reference = '',
        [object]$Evidence = $null
    )
    return [PSCustomObject]@{
        Id             = $Id
        Category       = $Category
        Name           = $Name
        Description    = $Description
        Status         = $Status
        Severity       = $Severity
        Details        = $Details
        Recommendation = $Recommendation
        Reference      = $Reference
        Evidence       = $Evidence
        Timestamp      = (Get-Date -Format 'o')
        Source         = 'Automated'
    }
}

function Invoke-GraphPaged {
    <#
    .SYNOPSIS
        Invokes a Graph GET and follows @odata.nextLink, returning all pages.
    #>
    param(
        [Parameter(Mandatory = $true)] [string]$Uri
    )
    $all = New-Object System.Collections.Generic.List[object]
    $next = $Uri
    while ($next) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop
        if ($null -ne $resp.value) {
            foreach ($v in $resp.value) { [void]$all.Add($v) }
        } elseif ($null -ne $resp) {
            [void]$all.Add($resp)
        }
        $next = $resp.'@odata.nextLink'
    }
    return $all.ToArray()
}

function Invoke-GraphReport {
    <#
    .SYNOPSIS
        POSTs to a virtualEndpoint/reports/<action> endpoint and returns the raw response.
        Report responses are columnar: { Schema:[...], Values:[[...]], TotalRowCount:n }.
    #>
    param(
        [Parameter(Mandatory = $true)] [string]$Action,
        [hashtable]$Body
    )
    if (-not $Body) { $Body = @{} }
    $uri = "$ReportsBase/$Action"
    return Invoke-MgGraphRequest -Method POST -Uri $uri -Body ($Body | ConvertTo-Json -Depth 6) -ContentType 'application/json' -ErrorAction Stop
}

function Get-ReportRowCount {
    <#
    .SYNOPSIS
        Best-effort row count from a columnar report response (handles TotalRowCount / Values).
    #>
    param([object]$Report)
    if ($null -eq $Report) { return -1 }
    if ($null -ne $Report.TotalRowCount) { return [int]$Report.TotalRowCount }
    if ($null -ne $Report.Values)        { return @($Report.Values).Count }
    return -1
}

# ═══════════════════════════════════════════════════════════════════════════
# BANNER
# ═══════════════════════════════════════════════════════════════════════════

Write-Host ""
Write-Host "   _    _ _____  __ _____    ___                                   " -ForegroundColor Cyan
Write-Host "  | |  | |___ /  / /| ____|  / _ \                                 " -ForegroundColor Cyan
Write-Host "  | |  | |  |_ \ / _ \  _|   | | | | ___  ___ ___  ___ ___  ___  _ __" -ForegroundColor Cyan
Write-Host "  | |/\| |__ _) | (_) | |___ | |_| |/ __|/ __/ _ \/ __/ __|/ _ \| '__|" -ForegroundColor Cyan
Write-Host "   \___/  |____/ \___/|_____| \__\_\\__ \__ \  __/__ \__ \  __/| |   " -ForegroundColor Cyan
Write-Host "                                  |___/___/\___|___/___/\___||_|   " -ForegroundColor Cyan
Write-Host ""
Write-Host "  v$ScriptVersion" -NoNewline -ForegroundColor DarkGray
Write-Host "  -  " -NoNewline -ForegroundColor DarkGray
Write-Host "CAF" -NoNewline -ForegroundColor Green
Write-Host "  -  " -NoNewline -ForegroundColor DarkGray
Write-Host "WAF" -NoNewline -ForegroundColor Blue
Write-Host "  -  " -NoNewline -ForegroundColor DarkGray
Write-Host "LZA" -NoNewline -ForegroundColor Yellow
Write-Host "  -  " -NoNewline -ForegroundColor DarkGray
Write-Host "SEC" -ForegroundColor Red
Write-Host ("  " + ('-' * 56)) -ForegroundColor DarkGray
Write-Host ""

# ═══════════════════════════════════════════════════════════════════════════
# PREREQUISITE CHECK
# ═══════════════════════════════════════════════════════════════════════════

Write-Status "Prerequisites" -Level 'SECTION'
$ReqMod = 'Microsoft.Graph.Authentication'
$Mod = Get-Module -ListAvailable -Name $ReqMod | Sort-Object Version -Descending | Select-Object -First 1
if (-not $Mod) {
    Write-Status "Required module not installed: $ReqMod" -Level 'ERROR'
    Write-Status "Install with: Install-Module $ReqMod -Scope CurrentUser" -Level 'INFO'
    exit 1
}
Write-Status "Module $ReqMod $($Mod.Version)" -Level 'SUCCESS'
Import-Module $ReqMod -ErrorAction Stop -WarningAction SilentlyContinue

# ═══════════════════════════════════════════════════════════════════════════
# AUTHENTICATION
# ═══════════════════════════════════════════════════════════════════════════

Write-Status "Authentication" -Level 'SECTION'
$Scopes = @(
    'CloudPC.Read.All'
    'DeviceManagementConfiguration.Read.All'
    'DeviceManagementManagedDevices.Read.All'
    'Directory.Read.All'
)
# Optional tier: requested at connect time but NOT enforced by the missing-scope gate below, so a
# tenant that declines it still runs (the Conditional Access checks degrade to Status 'Error').
$OptionalScopes = @(
    'Policy.Read.All'
)
$RequestScopes = @($Scopes + $OptionalScopes)

$Context = Get-MgContext -ErrorAction SilentlyContinue
if ($SkipLogin -and -not $Context) {
    Write-Status "SkipLogin specified but no existing Graph context found" -Level 'ERROR'
    exit 1
}

$NeedConnect = -not $Context
if ($Context) {
    $missing = @($Scopes | Where-Object { $Context.Scopes -notcontains $_ })
    if ($missing.Count -gt 0) {
        Write-Status "Existing context missing scopes: $($missing -join ', ')" -Level 'WARN'
        $NeedConnect = $true
    }
}

if ($NeedConnect -and -not $SkipLogin) {
    Write-Status "Connecting to Microsoft Graph..." -Level 'INFO'
    try {
        if ($TenantId) {
            Connect-MgGraph -Scopes $RequestScopes -TenantId $TenantId -NoWelcome -ErrorAction Stop | Out-Null
        } else {
            Connect-MgGraph -Scopes $RequestScopes -NoWelcome -ErrorAction Stop | Out-Null
        }
    } catch {
        Write-Status "Graph connection failed: $($_.Exception.Message)" -Level 'ERROR'
        exit 1
    }
    $Context = Get-MgContext
}
Write-Status "Tenant: $($Context.TenantId)" -Level 'SUCCESS'
Write-Status "Account: $($Context.Account)" -Level 'SUCCESS'

# ═══════════════════════════════════════════════════════════════════════════
# DISCOVERY
# ═══════════════════════════════════════════════════════════════════════════

$Discovery = [PSCustomObject]@{
    SchemaVersion = '1.0'
    ToolVersion   = $ScriptVersion
    Timestamp     = (Get-Date -Format 'o')
    AssessorId    = $Context.Account
    TenantId      = $Context.TenantId
    Inventory     = [PSCustomObject]@{
        CloudPCs                = @()
        ProvisioningPolicies    = @()
        UserSettings            = @()
        AzureNetworkConnections = @()
        DeviceImages            = @()
        GalleryImages           = @()
        ServicePlans            = @()
        AuditEvents             = @()
        ExternalPartnerSettings = @()
    }
    CheckResults  = @()
    Errors        = @()
}

$AllChecks = [System.Collections.ArrayList]::new()

function Add-DiscoveryError {
    <#
    .SYNOPSIS
        Records a collection failure to $Discovery.Errors AND emits a Status 'Error' CheckResult
        (A-3) so the failure is visible in the GUI import instead of being silently swallowed —
        a missing scope must never look like "no resources exist".
    #>
    param(
        [string]$Section,
        [string]$Message,
        [string]$Category = 'Discovery',
        [string]$Scope = 'CloudPC.Read.All'
    )
    Write-Status "$Section : $Message" -Level 'ERROR'
    $Discovery.Errors += "[$Section] $Message"
    $sectionId = ($Section -replace '[^A-Za-z0-9]', '').ToUpper()
    [void]$AllChecks.Add((New-CheckResult `
        -Id "W365-$sectionId-ERROR" -Category $Category `
        -Name "Collection failed: $Section" `
        -Description "The discovery collector could not enumerate the '$Section' resource set. Results for related checks are incomplete and must not be read as an all-clear." `
        -Status 'Error' -Severity 'High' `
        -Details "Graph call for '$Section' failed: $Message" `
        -Recommendation "Confirm the signed-in account holds '$Scope' (and admin consent is granted), that the tenant is licensed for Windows 365, and re-run discovery. A transient Graph error may also cause this." `
        -Reference 'https://learn.microsoft.com/en-us/graph/permissions-reference' `
        -Evidence @{ Section = $Section; Error = $Message; RequiredScope = $Scope }))
}

# ─── CLOUD PCs ────────────────────────────────────────────────────────────
Write-Status "Cloud PCs" -Level 'SECTION'
$CloudPCs = @()
try {
    $CloudPCs = @(Invoke-GraphPaged -Uri "$GraphBaseV1/cloudPCs")
    Write-Status "Found $($CloudPCs.Count) Cloud PC(s)" -Level 'SUCCESS'
    foreach ($cpc in $CloudPCs) {
        $Discovery.Inventory.CloudPCs += [PSCustomObject]@{
            Id                       = $cpc.id
            DisplayName              = $cpc.displayName
            Status                   = $cpc.status
            UserPrincipalName        = $cpc.userPrincipalName
            ImageDisplayName         = $cpc.imageDisplayName
            ProvisioningPolicyId     = $cpc.provisioningPolicyId
            ProvisioningPolicyName   = $cpc.provisioningPolicyName
            ProvisioningType         = $cpc.provisioningType
            ServicePlanName          = $cpc.servicePlanName
            ServicePlanId            = $cpc.servicePlanId
            ManagedDeviceId          = $cpc.managedDeviceId
            AadDeviceId              = $cpc.aadDeviceId
            OnPremisesConnectionName = $cpc.onPremisesConnectionName
            LastModifiedDateTime     = $cpc.lastModifiedDateTime
            LastLoginResult          = $cpc.lastLoginResult
            GracePeriodEndDateTime   = $cpc.gracePeriodEndDateTime
            DiskEncryptionState      = $cpc.diskEncryptionState
        }
    }
} catch {
    Add-DiscoveryError 'CloudPCs' $_.Exception.Message
}

# ─── PROVISIONING POLICIES ────────────────────────────────────────────────
Write-Status "Provisioning Policies" -Level 'SECTION'
$ProvPols = @()
try {
    $ProvPols = @(Invoke-GraphPaged -Uri "$GraphBaseV1/provisioningPolicies?`$expand=assignments")
    Write-Status "Found $($ProvPols.Count) provisioning policy/policies" -Level 'SUCCESS'
    foreach ($pp in $ProvPols) {
        $Discovery.Inventory.ProvisioningPolicies += [PSCustomObject]@{
            Id                           = $pp.id
            DisplayName                  = $pp.displayName
            Description                  = $pp.description
            DomainJoinConfigurations     = $pp.domainJoinConfigurations
            ImageId                      = $pp.imageId
            ImageDisplayName             = $pp.imageDisplayName
            ImageType                    = $pp.imageType
            EnableSingleSignOn           = $pp.enableSingleSignOn
            LocalAdminEnabled            = $pp.localAdminEnabled
            ProvisioningType             = $pp.provisioningType
            CloudPcGroupDisplayName      = $pp.cloudPcGroupDisplayName
            CloudPcNamingTemplate        = $pp.cloudPcNamingTemplate
            MicrosoftManagedDesktop      = $pp.microsoftManagedDesktop
            WindowsSetting               = $pp.windowsSetting
            AlternateResourceUrl         = $pp.alternateResourceUrl
            GracePeriodInHours           = $pp.gracePeriodInHours
            AutopatchEnabled             = $pp.autopatch.autopatchGroupId -ne $null
            AssignmentCount              = (@($pp.assignments)).Count
            Assignments                  = $pp.assignments
        }
    }
} catch {
    Add-DiscoveryError 'ProvisioningPolicies' $_.Exception.Message
}

# ─── USER SETTINGS ────────────────────────────────────────────────────────
Write-Status "User Settings Policies" -Level 'SECTION'
$UserSet = @()
try {
    $UserSet = @(Invoke-GraphPaged -Uri "$GraphBaseV1/userSettings?`$expand=assignments")
    Write-Status "Found $($UserSet.Count) user settings policy/policies" -Level 'SUCCESS'
    foreach ($us in $UserSet) {
        $Discovery.Inventory.UserSettings += [PSCustomObject]@{
            Id                              = $us.id
            DisplayName                     = $us.displayName
            LocalAdminEnabled               = $us.localAdminEnabled
            ResetEnabled                    = $us.resetEnabled
            RestorePointFrequencyInHours    = $us.restorePointSetting.frequencyInHours
            RestorePointUserRestoreEnabled  = $us.restorePointSetting.userRestoreEnabled
            CrossRegionDisasterRecoverySetting = $us.crossRegionDisasterRecoverySetting
            NotificationSetting             = $us.notificationSetting
            AssignmentCount                 = (@($us.assignments)).Count
            Assignments                     = $us.assignments
        }
    }
} catch {
    Add-DiscoveryError 'UserSettings' $_.Exception.Message
}

# ─── AZURE NETWORK CONNECTIONS ────────────────────────────────────────────
Write-Status "Azure Network Connections" -Level 'SECTION'
$ANCs = @()
try {
    $ANCs = @(Invoke-GraphPaged -Uri "$GraphBase/onPremisesConnections")
    Write-Status "Found $($ANCs.Count) network connection(s)" -Level 'SUCCESS'
    foreach ($anc in $ANCs) {
        $Discovery.Inventory.AzureNetworkConnections += [PSCustomObject]@{
            Id                           = $anc.id
            DisplayName                  = $anc.displayName
            Type                         = $anc.type
            ConnectionType               = $anc.connectionType
            HealthCheckStatus            = $anc.healthCheckStatus
            # Real schema: the detail object is 'healthCheckStatusDetail' (singular) carrying a
            # 'healthChecks' collection; accept the legacy plural key too, and surface the array.
            HealthCheckStatusDetail      = $anc.healthCheckStatusDetail
            HealthChecks                 = @(
                if ($anc.healthCheckStatusDetail -and $anc.healthCheckStatusDetail.healthChecks) {
                    $anc.healthCheckStatusDetail.healthChecks
                } elseif ($anc.healthCheckStatusDetails -and $anc.healthCheckStatusDetails.healthChecks) {
                    $anc.healthCheckStatusDetails.healthChecks
                } elseif ($anc.healthChecks) {
                    $anc.healthChecks
                }
            )
            InUse                        = $anc.inUse
            SubscriptionId               = $anc.subscriptionId
            SubscriptionName             = $anc.subscriptionName
            ResourceGroupId              = $anc.resourceGroupId
            VirtualNetworkId             = $anc.virtualNetworkId
            VirtualNetworkLocation       = $anc.virtualNetworkLocation
            SubnetId                     = $anc.subnetId
            AdDomainName                 = $anc.adDomainName
            AdDomainUsername             = $anc.adDomainUsername
            OrganizationalUnit           = $anc.organizationalUnit
            ConnectionStatus             = $anc.connectionStatus
        }
    }
} catch {
    Add-DiscoveryError 'AzureNetworkConnections' $_.Exception.Message
}

# ─── DEVICE IMAGES (CUSTOM) ───────────────────────────────────────────────
Write-Status "Custom Device Images" -Level 'SECTION'
$DevImgs = @()
try {
    $DevImgs = @(Invoke-GraphPaged -Uri "$GraphBase/deviceImages")
    Write-Status "Found $($DevImgs.Count) custom image(s)" -Level 'SUCCESS'
    foreach ($di in $DevImgs) {
        $Discovery.Inventory.DeviceImages += [PSCustomObject]@{
            Id                = $di.id
            DisplayName       = $di.displayName
            Version           = $di.version
            OsBuildNumber     = $di.osBuildNumber
            OperatingSystem   = $di.operatingSystem
            SourceImageResourceId = $di.sourceImageResourceId
            Status            = $di.status
            StatusDetails     = $di.statusDetails
            ErrorCode         = $di.errorCode
            LastModifiedDateTime = $di.lastModifiedDateTime
        }
    }
} catch {
    Add-DiscoveryError 'DeviceImages' $_.Exception.Message
}

# ─── GALLERY IMAGES ───────────────────────────────────────────────────────
Write-Status "Gallery Images" -Level 'SECTION'
$GalImgs = @()
try {
    $GalImgs = @(Invoke-GraphPaged -Uri "$GraphBase/galleryImages")
    Write-Status "Found $($GalImgs.Count) gallery image(s)" -Level 'SUCCESS'
    foreach ($gi in $GalImgs) {
        $Discovery.Inventory.GalleryImages += [PSCustomObject]@{
            Id                  = $gi.id
            DisplayName         = $gi.displayName
            PublisherName       = $gi.publisherName
            OfferName           = $gi.offerName
            SkuName             = $gi.skuName
            SizeInGB            = $gi.sizeInGB
            Status              = $gi.status
            StartDateTime       = $gi.startDateTime
            ExpirationDateTime  = $gi.expirationDateTime
            EndOfSupportDateTime = $gi.endOfSupportDateTime
        }
    }
} catch {
    Add-DiscoveryError 'GalleryImages' $_.Exception.Message
}

# ─── SERVICE PLANS ────────────────────────────────────────────────────────
Write-Status "Service Plans (SKUs)" -Level 'SECTION'
try {
    $Plans = @(Invoke-GraphPaged -Uri "$GraphBase/servicePlans")
    Write-Status "Found $($Plans.Count) service plan(s)" -Level 'SUCCESS'
    foreach ($sp in $Plans) {
        $Discovery.Inventory.ServicePlans += [PSCustomObject]@{
            Id           = $sp.id
            DisplayName  = $sp.displayName
            Type         = $sp.type
            VCpuCount    = $sp.vCpuCount
            RamInGB      = $sp.ramInGB
            StorageInGB  = $sp.storageInGB
            UserProfileInGB = $sp.userProfileInGB
        }
    }
} catch {
    Add-DiscoveryError 'ServicePlans' $_.Exception.Message
}

# ─── AUDIT EVENTS (last 30 days) ──────────────────────────────────────────
Write-Status "Audit Events (last 30 days)" -Level 'SECTION'
try {
    # A-2 (PR #1): emit a UTC 'Z' stamp — .ToString('o') produces a local-offset stamp whose '+'
    # breaks the Graph $filter for UTC-positive clients (BadRequest).
    $since = (Get-Date).ToUniversalTime().AddDays(-30).ToString('yyyy-MM-ddTHH:mm:ssZ')
    $auditUri = "$GraphBase/auditEvents?`$filter=activityDateTime ge $since&`$top=200"
    $audits = @(Invoke-GraphPaged -Uri $auditUri)
    Write-Status "Found $($audits.Count) audit event(s)" -Level 'SUCCESS'
    foreach ($ae in $audits) {
        $Discovery.Inventory.AuditEvents += [PSCustomObject]@{
            Id               = $ae.id
            DisplayName      = $ae.displayName
            ActivityType     = $ae.activityType
            ActivityResult   = $ae.activityResult
            ActivityDateTime = $ae.activityDateTime
            CategoryName     = $ae.category
            ActorUpn         = $ae.actor.userPrincipalName
            ActorAppName     = $ae.actor.applicationDisplayName
            ComponentName    = $ae.componentName
        }
    }
} catch {
    Add-DiscoveryError 'AuditEvents' $_.Exception.Message
}

# ═══════════════════════════════════════════════════════════════════════════
# AUTOMATED CHECKS
# ═══════════════════════════════════════════════════════════════════════════

Write-Status "Automated checks" -Level 'SECTION'

# A-1: single evaluation clock, hoisted above every date-delta check (IMG-007 gallery EOS loop and
# the COST-001 inactivity loop both run before the old assignment point and were silently failing).
$now = Get-Date

# Inventory snapshot (informational with state breakdown)
$StateGroups = $Discovery.Inventory.CloudPCs | Group-Object Status | Sort-Object Count -Descending
$StateSummary = if ($StateGroups) { ($StateGroups | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ', ' } else { 'none' }
$ProvAssigned = @($Discovery.Inventory.ProvisioningPolicies | Where-Object { $_.AssignmentCount -gt 0 }).Count
$AncHealthy = @($Discovery.Inventory.AzureNetworkConnections | Where-Object { $_.HealthCheckStatus -eq 'passed' }).Count
[void]$AllChecks.Add((New-CheckResult `
    -Id 'W365-INV-001' -Category 'Inventory & Topology' `
    -Name 'Cloud PC inventory' `
    -Description 'Snapshot of Cloud PC resources in the tenant. Use as a baseline for sizing and cost analysis.' `
    -Status 'Pass' -Severity 'Low' `
    -Details "Cloud PCs: $($CloudPCs.Count) ($StateSummary). Provisioning policies: $($ProvPols.Count) ($ProvAssigned assigned). User settings: $($UserSet.Count). ANCs: $($ANCs.Count) ($AncHealthy healthy). Custom images: $($DevImgs.Count). Gallery images: $($GalImgs.Count)." `
    -Recommendation 'Use the Provisioning Policies and Cloud PCs panels in the GUI to drill into details. Schedule monthly inventory reviews to catch orphaned licenses and stale resources early.' `
    -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/overview' `
    -Evidence @{ CloudPCs = $CloudPCs.Count; Policies = $ProvPols.Count; ANCs = $ANCs.Count; States = $StateSummary }))

# PROV-001: Provisioning policies must have at least one assignment
foreach ($pp in $Discovery.Inventory.ProvisioningPolicies) {
    $hasAssign = ($pp.AssignmentCount -gt 0)
    [void]$AllChecks.Add((New-CheckResult `
        -Id "W365-PROV-001-$($pp.Id)" -Category 'Provisioning Policies' `
        -Name "Policy assignment: $($pp.DisplayName)" `
        -Description 'A provisioning policy without at least one user-group assignment cannot provision Cloud PCs.' `
        -Status $(if ($hasAssign) { 'Pass' } else { 'Fail' }) `
        -Severity 'High' `
        -Details "Policy '$($pp.DisplayName)' has $($pp.AssignmentCount) assignment(s)." `
        -Recommendation 'Assign the provisioning policy to one or more Microsoft Entra security groups containing licensed users.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/provisioning' `
        -Evidence @{ PolicyName = $pp.DisplayName; PolicyId = $pp.Id; AssignmentCount = $pp.AssignmentCount }))
}

# IAM-001: SSO recommended on every provisioning policy.
# NOTE (reconciliation): this SSO check previously emitted W365-PROV-002-*. The PROV-002 id is
# reassigned by the fix spec to the Cloud PC naming-template check, and this check's catalog
# cross-map is IAM-001, so it now emits W365-IAM-001-* (importer arm: W365-IAM-001-* -> IAM-001).
foreach ($pp in $Discovery.Inventory.ProvisioningPolicies) {
    [void]$AllChecks.Add((New-CheckResult `
        -Id "W365-IAM-001-$($pp.Id)" -Category 'Identity & Access' `
        -Name "Single sign-on: $($pp.DisplayName)" `
        -Description 'Single sign-on (SSO) reduces sign-in prompts when connecting to a Cloud PC and is the Microsoft-recommended configuration.' `
        -Status $(if ($pp.EnableSingleSignOn) { 'Pass' } else { 'Warning' }) `
        -Severity 'Medium' `
        -Details "EnableSingleSignOn = $($pp.EnableSingleSignOn) on policy '$($pp.DisplayName)'. Without SSO, users see a second Windows password prompt after the Cloud PC connection establishes." `
        -Recommendation 'Edit the provisioning policy and enable Single Sign-On. Requires Entra ID Join (or Hybrid Join with cloud Kerberos trust) and pairs well with Windows Hello for Business for passwordless sign-in.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/set-up-tenants-passwordless-authentication' `
        -Evidence @{ PolicyName = $pp.DisplayName; SSO = $pp.EnableSingleSignOn }))
}

# PROV-003: Local admin enabled is a security signal — surface as Warning to force review
foreach ($pp in $Discovery.Inventory.ProvisioningPolicies) {
    if ($pp.LocalAdminEnabled) {
        [void]$AllChecks.Add((New-CheckResult `
            -Id "W365-PROV-003-$($pp.Id)" -Category 'Security & Compliance' `
            -Name "Local admin enabled: $($pp.DisplayName)" `
            -Description 'Granting users local admin on their Cloud PC widens the blast radius of malware and weakens compliance posture.' `
            -Status 'Warning' -Severity 'High' `
            -Details "Policy '$($pp.DisplayName)' has localAdminEnabled = true." `
            -Recommendation 'Disable local admin in the provisioning policy unless required by an approved exception. Use Endpoint Privilege Management or Intune just-in-time elevation instead.' `
            -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/create-provisioning-policy' `
            -Evidence @{ PolicyName = $pp.DisplayName; LocalAdminEnabled = $true }))
    }
}

# USER-001: User settings must have at least one assignment, otherwise restore points/local admin policy is not applied
foreach ($us in $Discovery.Inventory.UserSettings) {
    $hasAssign = ($us.AssignmentCount -gt 0)
    [void]$AllChecks.Add((New-CheckResult `
        -Id "W365-USER-001-$($us.Id)" -Category 'User Settings & Resilience' `
        -Name "User settings assignment: $($us.DisplayName)" `
        -Description 'Unassigned user settings policies do not apply restore-point frequency, local admin, or DR settings to any user.' `
        -Status $(if ($hasAssign) { 'Pass' } else { 'Warning' }) `
        -Severity 'Medium' `
        -Details "User settings '$($us.DisplayName)' has $($us.AssignmentCount) assignment(s)." `
        -Recommendation 'Assign every user settings policy to a security group, or delete unused policies.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/create-user-settings-policy' `
        -Evidence @{ Name = $us.DisplayName; AssignmentCount = $us.AssignmentCount }))
}

# USER-002: Cross-region DR enabled = resilience signal
foreach ($us in $Discovery.Inventory.UserSettings) {
    $crEnabled = $false
    if ($us.CrossRegionDisasterRecoverySetting) {
        $crEnabled = [bool]$us.CrossRegionDisasterRecoverySetting.disasterRecoveryType -and `
                     $us.CrossRegionDisasterRecoverySetting.disasterRecoveryType -ne 'notConfigured'
    }
    [void]$AllChecks.Add((New-CheckResult `
        -Id "W365-USER-002-$($us.Id)" -Category 'User Settings & Resilience' `
        -Name "Cross-region DR: $($us.DisplayName)" `
        -Description 'Cross-region disaster recovery enables a failover Cloud PC in a paired region.' `
        -Status $(if ($crEnabled) { 'Pass' } else { 'Warning' }) `
        -Severity 'Medium' `
        -Details "DR type: $($us.CrossRegionDisasterRecoverySetting.disasterRecoveryType)." `
        -Recommendation 'For business-critical personas, configure cross-region DR in the user settings policy.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/cross-region-disaster-recovery' `
        -Evidence @{ Name = $us.DisplayName; DR = $us.CrossRegionDisasterRecoverySetting }))
}

# USER-003: Restore point frequency on user settings policies (Microsoft default = 12h; >24h is risky)
foreach ($us in $Discovery.Inventory.UserSettings) {
    $freq = [int]($us.RestorePointFrequencyInHours | ForEach-Object { if ($_) { $_ } else { 0 } })
    if ($freq -le 0) { continue }
    $st = if ($freq -le 12) { 'Pass' } elseif ($freq -le 24) { 'Warning' } else { 'Fail' }
    [void]$AllChecks.Add((New-CheckResult `
        -Id "W365-USER-003-$($us.Id)" -Category 'User Settings & Resilience' `
        -Name "Restore point frequency: $($us.DisplayName)" `
        -Description 'Restore-point frequency determines the maximum data-loss window if a Cloud PC is rolled back. Microsoft default is every 12 hours.' `
        -Status $st -Severity 'Medium' `
        -Details "Restore point frequency = $freq hour(s) on '$($us.DisplayName)'." `
        -Recommendation 'For most personas, leave at 12 hours (default). For data-critical personas, consider 4 or 6 hours. Frequencies above 24 hours significantly widen the data-loss window.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/configure-restore-points-and-users' `
        -Evidence @{ Name = $us.DisplayName; FrequencyInHours = $freq }))
}

# USER-004: User self-service reset enabled signal
foreach ($us in $Discovery.Inventory.UserSettings) {
    [void]$AllChecks.Add((New-CheckResult `
        -Id "W365-USER-004-$($us.Id)" -Category 'User Settings & Resilience' `
        -Name "User self-service reset: $($us.DisplayName)" `
        -Description 'User self-service reset lets a user reprovision their own Cloud PC, reducing helpdesk tickets but also expanding the change blast radius if misused.' `
        -Status $(if ($us.ResetEnabled) { 'Pass' } else { 'Warning' }) `
        -Severity 'Low' `
        -Details "ResetEnabled = $($us.ResetEnabled) on '$($us.DisplayName)'." `
        -Recommendation 'Enable user self-service reset for low-risk personas to reduce ticket volume; keep it disabled for regulated/admin personas where reset must follow change control.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/create-user-settings-policy' `
        -Evidence @{ Name = $us.DisplayName; ResetEnabled = $us.ResetEnabled }))
}

# PROV-004: Windows Autopatch integration on provisioning policy
foreach ($pp in $Discovery.Inventory.ProvisioningPolicies) {
    [void]$AllChecks.Add((New-CheckResult `
        -Id "W365-PROV-004-$($pp.Id)" -Category 'Provisioning Policies' `
        -Name "Windows Autopatch: $($pp.DisplayName)" `
        -Description 'Windows Autopatch automates Windows quality, feature, driver, and Microsoft 365 Apps updates with built-in safeguards. Enabling it on a provisioning policy delegates patch management to the Microsoft service.' `
        -Status $(if ($pp.AutopatchEnabled) { 'Pass' } else { 'Warning' }) `
        -Severity 'Medium' `
        -Details "AutopatchEnabled = $($pp.AutopatchEnabled) on '$($pp.DisplayName)'." `
        -Recommendation 'Where update control can be delegated, enable Windows Autopatch on the provisioning policy and assign the Cloud PCs to a managed Autopatch group.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/windows-autopatch' `
        -Evidence @{ PolicyName = $pp.DisplayName; AutopatchEnabled = $pp.AutopatchEnabled }))
}

# PROV-005: Grace period configured.
# C-3 VERIFIED: gracePeriodInHours IS a real Int32 property on cloudPcProvisioningPolicy (v1.0 + beta) —
# "the number of hours to wait before reprovisioning/deprovisioning happens". The audit's suspicion
# (that grace is a fixed 7-day service constant only) is NOT correct; the policy-level value is real
# and evaluable, so the check is kept. 0 = deprovision immediately on license loss (no recovery window).
foreach ($pp in $Discovery.Inventory.ProvisioningPolicies) {
    $gp = [int]($pp.GracePeriodInHours | ForEach-Object { if ($_) { $_ } else { 0 } })
    $st = if ($gp -le 0) { 'Warning' } elseif ($gp -ge 1 -and $gp -le 168) { 'Pass' } else { 'Warning' }
    $details = "gracePeriodInHours = $gp on '$($pp.DisplayName)' (hours to wait before deprovisioning after license loss)."
    [void]$AllChecks.Add((New-CheckResult `
        -Id "W365-PROV-005-$($pp.Id)" -Category 'Provisioning Policies' `
        -Name "Grace period configuration: $($pp.DisplayName)" `
        -Description 'The grace period delays Cloud PC de-provisioning after license loss, giving the user a chance to recover data and admins a window to reassign.' `
        -Status $st -Severity 'Medium' `
        -Details $details `
        -Recommendation 'Use a grace period of at least 1 hour (recommended 7 days / 168 hours where data recovery is critical). A grace period of 0 deprovisions immediately on license loss with no recovery window.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/grace-period' `
        -Evidence @{ PolicyName = $pp.DisplayName; GracePeriodInHours = $gp }))
}

# NET-001: ANC health check status
foreach ($anc in $Discovery.Inventory.AzureNetworkConnections) {
    $h = "$($anc.HealthCheckStatus)"
    $status = switch -Wildcard ($h) {
        'passed'     { 'Pass' }
        'warning'    { 'Warning' }
        'failed'     { 'Fail' }
        'running'    { 'Warning' }
        ''           { 'Warning' }
        default      { 'Warning' }
    }
    # C-4: parse failed sub-checks from the real healthChecks[] collection. Each item is a
    # cloudPcOnPremisesConnectionHealthCheck with { displayName, status, errorType, recommendedAction }.
    $failedTests = @()
    if ($anc.HealthChecks) {
        $failedTests = @($anc.HealthChecks |
            Where-Object { "$($_.status)" -in @('failed','warning','error') } |
            ForEach-Object {
                $et = if ($_.errorType) { " ($($_.errorType))" } else { '' }
                "$($_.displayName)$et"
            })
    }
    $detailExtra = if ($failedTests) { " Failing sub-checks: $($failedTests -join ', ')." } else { '' }
    [void]$AllChecks.Add((New-CheckResult `
        -Id "W365-NET-001-$($anc.Id)" -Category 'Network (ANC)' `
        -Name "ANC health: $($anc.DisplayName)" `
        -Description 'Azure Network Connection health checks validate AD reachability, DNS resolution, DHCP, endpoint connectivity, NSG rules, IP availability, and Intune enrollment endpoints. A failing ANC blocks all new provisioning on that connection.' `
        -Status $status -Severity 'High' `
        -Details "ANC '$($anc.DisplayName)' (region: $($anc.VirtualNetworkLocation), type: $($anc.ConnectionType)) last health check status: $h.$detailExtra" `
        -Recommendation 'Open the Azure Network Connection in Intune > Devices > Windows 365 > Azure network connections and run health checks. Remediate by checking VNet routing, NSG outbound rules to required FQDNs (WindowsVirtualDesktop, Windows365 service tags), DNS resolution, and (for Hybrid Join) AD DS reachability and credentials.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/azure-network-connections' `
        -Evidence @{ Name = $anc.DisplayName; Status = $h; FailedTests = $failedTests; Region = $anc.VirtualNetworkLocation; ConnectionType = $anc.ConnectionType }))
}

# NET-009: ANC defined but not in use
foreach ($anc in $Discovery.Inventory.AzureNetworkConnections) {
    if ($anc.InUse -eq $false -or $anc.InUse -eq $null) {
        # Check if any provisioning policy actually references it
        $referenced = @($Discovery.Inventory.ProvisioningPolicies | Where-Object {
            $_.DomainJoinConfigurations -and ($_.DomainJoinConfigurations | Where-Object { $_.onPremisesConnectionId -eq $anc.Id }).Count -gt 0
        }).Count
        if ($referenced -eq 0) {
            [void]$AllChecks.Add((New-CheckResult `
                -Id "W365-NET-009-$($anc.Id)" -Category 'Network (ANC)' `
                -Name "Unused ANC: $($anc.DisplayName)" `
                -Description 'Azure Network Connections that are defined but not referenced by any provisioning policy add operational overhead and are easily forgotten when network changes occur.' `
                -Status 'Warning' -Severity 'Low' `
                -Details "ANC '$($anc.DisplayName)' (region: $($anc.VirtualNetworkLocation)) is not referenced by any provisioning policy." `
                -Recommendation 'Delete the ANC if no longer needed, or document its intended future use. Unused ANCs still have health checks running and consume tenant quota.' `
                -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/azure-network-connections' `
                -Evidence @{ Name = $anc.DisplayName; InUse = $anc.InUse; ReferencedBy = $referenced }))
        }
    }
}

# IMG-007: Gallery image at or near end-of-support
foreach ($gi in $Discovery.Inventory.GalleryImages) {
    if (-not $gi.EndOfSupportDateTime) { continue }
    try { $eosDays = ([datetime]$gi.EndOfSupportDateTime - $now).Days } catch { continue }
    if ($eosDays -le 0) {
        [void]$AllChecks.Add((New-CheckResult `
            -Id "W365-IMG-007-$($gi.Id)" -Category 'Images & App Delivery' `
            -Name "Gallery image past end-of-support: $($gi.DisplayName)" `
            -Description 'Cloud PCs provisioned from a gallery image past end-of-support no longer receive security patches and may not be eligible for support.' `
            -Status 'Fail' -Severity 'High' `
            -Details "Gallery image '$($gi.DisplayName)' reached end-of-support $([math]::Abs($eosDays)) day(s) ago ($($gi.EndOfSupportDateTime))." `
            -Recommendation 'Update provisioning policies to a supported gallery image SKU and reprovision affected Cloud PCs in a planned maintenance window.' `
            -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/gallery-image' `
            -Evidence @{ Name = $gi.DisplayName; EOS = $gi.EndOfSupportDateTime; DaysSince = [math]::Abs($eosDays) }))
    } elseif ($eosDays -le 90) {
        [void]$AllChecks.Add((New-CheckResult `
            -Id "W365-IMG-007-$($gi.Id)" -Category 'Images & App Delivery' `
            -Name "Gallery image near end-of-support: $($gi.DisplayName)" `
            -Description 'Gallery image is approaching end-of-support and should be planned for replacement.' `
            -Status 'Warning' -Severity 'Medium' `
            -Details "Gallery image '$($gi.DisplayName)' reaches end-of-support in $eosDays day(s) ($($gi.EndOfSupportDateTime))." `
            -Recommendation 'Identify a replacement supported gallery image SKU and schedule the migration of provisioning policies and existing Cloud PCs before the EOS date.' `
            -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/gallery-image' `
            -Evidence @{ Name = $gi.DisplayName; EOS = $gi.EndOfSupportDateTime; DaysUntil = $eosDays }))
    }
}

# MON-007: Audit log activity in the last 30 days (proxy for log retention/activity)
$auditCount = $Discovery.Inventory.AuditEvents.Count
[void]$AllChecks.Add((New-CheckResult `
    -Id 'W365-MON-007' -Category 'Monitoring & Diagnostics' `
    -Name 'Audit log activity (30d)' `
    -Description 'Windows 365 audit events flow through Microsoft Graph and have limited retention. Forwarding them to Sentinel or Log Analytics extends retention and enables alerting.' `
    -Status $(if ($auditCount -gt 0) { 'Pass' } else { 'Warning' }) `
    -Severity 'Medium' `
    -Details "$auditCount audit event(s) recorded in the last 30 days." `
    -Recommendation $(if ($auditCount -eq 0) { 'No audit activity detected — confirm admin operations are being captured. If absent, validate Microsoft Entra audit log integration and connector status.' } else { 'For long-term retention (>30d) and alerting on policy deletion / mass reprovision / DR activation, forward Windows 365 audit events to Microsoft Sentinel via Graph API connector or Diagnostic Settings (when available).' }) `
    -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/audit-logs' `
    -Evidence @{ EventsLast30Days = $auditCount }))

# IMG-001: Custom image age ($now hoisted to top of section — A-1)
foreach ($di in $Discovery.Inventory.DeviceImages) {
    if (-not $di.LastModifiedDateTime) { continue }
    try {
        $age = ($now - [datetime]$di.LastModifiedDateTime).Days
    } catch { $age = -1 }
    if ($age -lt 0) { continue }
    $status = if ($age -le $ImageAgeWarnDays) { 'Pass' } elseif ($age -le ($ImageAgeWarnDays * 2)) { 'Warning' } else { 'Fail' }
    [void]$AllChecks.Add((New-CheckResult `
        -Id "W365-IMG-001-$($di.Id)" -Category 'Images & App Delivery' `
        -Name "Image age: $($di.DisplayName)" `
        -Description 'Stale custom Cloud PC images mean newly provisioned PCs ship with outdated baseline software and security patches.' `
        -Status $status -Severity 'Medium' `
        -Details "Custom image '$($di.DisplayName)' last modified $age day(s) ago (threshold $ImageAgeWarnDays)." `
        -Recommendation 'Rebuild the custom image at least quarterly. Consider Azure Image Builder + a CI pipeline.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/device-images' `
        -Evidence @{ Name = $di.DisplayName; AgeDays = $age; LastModified = $di.LastModifiedDateTime }))
}

# CPC-001: Cloud PCs in any unhealthy state (C-2: evaluate the full unhealthy-state set, not just
# 'failed'). Hard-fail states are not usable; transient/degraded states are surfaced as Warning.
# 'inGracePeriod' keeps its own dedicated check (CPC-002); 'provisioned' is the healthy state.
$failStates = @('failed','notProvisioned')
$warnStates = @('provisionedWithWarnings','provisioning','deprovisioning','resizing','restoring','pendingProvision','movingRegion','unknown')
foreach ($cpc in $Discovery.Inventory.CloudPCs) {
    $s = "$($cpc.Status)"
    if ($s -in $failStates) {
        [void]$AllChecks.Add((New-CheckResult `
            -Id "W365-CPC-001-$($cpc.Id)" -Category 'Inventory & Topology' `
            -Name "Cloud PC in unhealthy state ($s): $($cpc.DisplayName)" `
            -Description 'Cloud PCs in failed or not-provisioned states are not usable and usually indicate a provisioning, networking, or licensing issue.' `
            -Status 'Fail' -Severity 'High' `
            -Details "Cloud PC '$($cpc.DisplayName)' for $($cpc.UserPrincipalName) is in status: $s." `
            -Recommendation 'Investigate via Intune > Devices > Cloud PCs. Check the provisioning policy, ANC health, and license assignment; reprovision or open a support ticket if persistent.' `
            -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/known-issues-provisioning' `
            -Evidence @{ Name = $cpc.DisplayName; UPN = $cpc.UserPrincipalName; Status = $s }))
    } elseif ($s -in $warnStates) {
        [void]$AllChecks.Add((New-CheckResult `
            -Id "W365-CPC-001-$($cpc.Id)" -Category 'Inventory & Topology' `
            -Name "Cloud PC in transient/degraded state ($s): $($cpc.DisplayName)" `
            -Description 'The Cloud PC is not in the steady "provisioned" state. Transient states are expected briefly; persistence indicates a stuck operation or a partially-successful provision.' `
            -Status 'Warning' -Severity 'Medium' `
            -Details "Cloud PC '$($cpc.DisplayName)' for $($cpc.UserPrincipalName) is in status: $s." `
            -Recommendation 'If the state persists beyond the expected operation window, review the Cloud PC action status and ANC health, then reprovision if stuck.' `
            -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/known-issues-provisioning' `
            -Evidence @{ Name = $cpc.DisplayName; UPN = $cpc.UserPrincipalName; Status = $s }))
    } elseif ($s -eq 'inGracePeriod') {
        [void]$AllChecks.Add((New-CheckResult `
            -Id "W365-CPC-002-$($cpc.Id)" -Category 'Inventory & Topology' `
            -Name "Cloud PC in grace period: $($cpc.DisplayName)" `
            -Description 'Cloud PCs in grace period will be deprovisioned at the end of the grace window.' `
            -Status 'Warning' -Severity 'Medium' `
            -Details "Cloud PC '$($cpc.DisplayName)' grace ends $($cpc.GracePeriodEndDateTime)." `
            -Recommendation 'Reassign a license, end the grace period to deprovision now, or restore the user assignment.' `
            -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/end-grace-period' `
            -Evidence @{ Name = $cpc.DisplayName; UPN = $cpc.UserPrincipalName; GraceEnd = $cpc.GracePeriodEndDateTime }))
    }
}

# COST-001 (was CPC-003): Cloud PCs likely inactive — license reclaim / downsize candidates.
# C-1: use the real inactivity signal, LastLoginResult (already collected), instead of
# LastModifiedDateTime (which bumps on any config change). Cloud PCs whose login data is null are
# counted and handled via the reports fallback (getInactiveCloudPcReport family) in the Reports section.
$CpcNoLoginCount = 0
foreach ($cpc in $Discovery.Inventory.CloudPCs) {
    if ("$($cpc.Status)" -notin @('provisioned','provisionedWithWarnings')) { continue }
    $lastLogin = $null
    if ($cpc.LastLoginResult) {
        foreach ($k in @('lastLoginDateTime','LastLoginDateTime','time','Time')) {
            if ($cpc.LastLoginResult.$k) { $lastLogin = $cpc.LastLoginResult.$k; break }
        }
    }
    if (-not $lastLogin) { $CpcNoLoginCount++; continue }
    try { $idle = ($now - [datetime]$lastLogin).Days } catch { continue }
    if ($idle -lt 0) { continue }
    if ($idle -gt $InactiveDays) {
        [void]$AllChecks.Add((New-CheckResult `
            -Id "W365-COST-001-$($cpc.Id)" -Category 'Cost & Optimization' `
            -Name "Inactive Cloud PC: $($cpc.DisplayName)" `
            -Description 'Provisioned Cloud PCs with no interactive sign-in for an extended period are candidates for license reclaim or SKU downsizing.' `
            -Status 'Warning' -Severity 'Low' `
            -Details "Cloud PC '$($cpc.DisplayName)' ($($cpc.UserPrincipalName)) last signed in $idle day(s) ago (threshold $InactiveDays). Signal: LastLoginResult." `
            -Recommendation 'Confirm inactivity in Endpoint Analytics / the inactive Cloud PC report, then reclaim the license or downsize the SKU.' `
            -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/report-cloud-pc-utilization' `
            -Evidence @{ Name = $cpc.DisplayName; UPN = $cpc.UserPrincipalName; IdleDays = $idle; LastLogin = "$lastLogin" }))
    }
}

# PROV-002 (was PROV-008): Cloud PC naming template present AND sane.
# F-automation: id reassigned from PROV-008 to PROV-002 per the fix spec (the old PROV-002 SSO check
# moved to IAM-001). "Sane" = present, <=15-char output budget, and carries a uniqueness token
# (%RAND% or %USERNAME%) so device names don't collide.
foreach ($pp in $Discovery.Inventory.ProvisioningPolicies) {
    $tpl = "$($pp.CloudPcNamingTemplate)"
    $hasTpl = -not [string]::IsNullOrWhiteSpace($tpl)
    $hasUniqueToken = $tpl -match '%RAND' -or $tpl -match '%USERNAME'
    # Static text length once tokens are stripped — the fixed portion must leave room within 15 chars.
    $staticLen = ($tpl -replace '%[A-Za-z]+(:\d+)?%', '').Length
    $sane = $hasTpl -and $hasUniqueToken -and $staticLen -le 15
    $st = if (-not $hasTpl) { 'Warning' } elseif ($sane) { 'Pass' } else { 'Warning' }
    $why = if (-not $hasTpl) { '(none)' }
           elseif (-not $hasUniqueToken) { "$tpl (no %RAND%/%USERNAME% uniqueness token — collision risk)" }
           elseif ($staticLen -gt 15) { "$tpl (static text exceeds the 15-char name budget)" }
           else { $tpl }
    [void]$AllChecks.Add((New-CheckResult `
        -Id "W365-PROV-002-$($pp.Id)" -Category 'Provisioning Policies' `
        -Name "Cloud PC naming template: $($pp.DisplayName)" `
        -Description 'A naming template (e.g. CPC-%USERNAME:5%-%RAND:5%) produces predictable, traceable, collision-free Cloud PC device names. Cloud PC names are capped at 15 characters and must include a uniqueness token.' `
        -Status $st -Severity 'Low' `
        -Details "Policy '$($pp.DisplayName)' template: $why." `
        -Recommendation 'Set a naming template that encodes user identity plus a uniqueness token and fits 15 characters (e.g. ''CPC-%USERNAME:5%-%RAND:5%'').' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/create-provisioning-policy' `
        -Evidence @{ PolicyName = $pp.DisplayName; Template = $pp.CloudPcNamingTemplate; HasUniquenessToken = $hasUniqueToken }))
}

# PROV-007 (was PROV-009): Provisioning policy windowsSetting locale/config present.
# F-automation: id reassigned from PROV-009 to PROV-007 per the fix spec.
foreach ($pp in $Discovery.Inventory.ProvisioningPolicies) {
    $ws = $pp.WindowsSetting
    $hasLocale = $ws -and -not [string]::IsNullOrWhiteSpace("$($ws.locale)")
    [void]$AllChecks.Add((New-CheckResult `
        -Id "W365-PROV-007-$($pp.Id)" -Category 'Provisioning Policies' `
        -Name "Windows setting (locale) configured: $($pp.DisplayName)" `
        -Description 'The provisioning policy windowsSetting (locale) ensures Cloud PCs ship with the correct OS language/region without first-login MUI churn.' `
        -Status $(if ($hasLocale) { 'Pass' } else { 'Warning' }) `
        -Severity 'Low' `
        -Details "Policy '$($pp.DisplayName)' windowsSetting locale: $(if ($hasLocale) { $ws.locale } else { '(not set — defaults to en-US)' })." `
        -Recommendation 'Configure the Windows setting on the provisioning policy with the locale appropriate to the user persona/region.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/create-provisioning-policy' `
        -Evidence @{ PolicyName = $pp.DisplayName; Locale = "$($ws.locale)" }))
}

# CPC-004: Provisioned Cloud PC with missing UPN (orphaned)
foreach ($cpc in $Discovery.Inventory.CloudPCs) {
    if ([string]::IsNullOrWhiteSpace($cpc.UserPrincipalName) -and $cpc.Status -in @('provisioned','provisionedWithWarnings','inGracePeriod')) {
        [void]$AllChecks.Add((New-CheckResult `
            -Id "W365-CPC-004-$($cpc.Id)" -Category 'Inventory & Topology' `
            -Name "Orphaned Cloud PC (no user assigned): $($cpc.DisplayName)" `
            -Description 'A provisioned Cloud PC without an assigned user is consuming a license but cannot be signed into. This usually indicates a stale assignment after an HR offboarding or a failed reassignment.' `
            -Status 'Fail' -Severity 'High' `
            -Details "Cloud PC '$($cpc.DisplayName)' (status: $($cpc.Status)) has no userPrincipalName." `
            -Recommendation 'Reassign the Cloud PC to an active user via the provisioning policy assignment, or end the grace period to deprovision and reclaim the license.' `
            -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/end-grace-period' `
            -Evidence @{ Name = $cpc.DisplayName; Status = $cpc.Status }))
    }
}

# SEC-010: Cloud PC disk encryption state
foreach ($cpc in $Discovery.Inventory.CloudPCs) {
    $enc = "$($cpc.DiskEncryptionState)"
    if ([string]::IsNullOrWhiteSpace($enc) -or $enc -eq 'notAvailable') { continue }
    if ($enc -ne 'encryptedUsingPlatformManagedKey' -and $enc -ne 'encryptedUsingCustomerManagedKey') {
        [void]$AllChecks.Add((New-CheckResult `
            -Id "W365-SEC-010-$($cpc.Id)" -Category 'Security & Compliance' `
            -Name "Disk encryption: $($cpc.DisplayName)" `
            -Description 'Azure managed disks for Cloud PCs are encrypted at rest by default with platform-managed keys. A reported state other than encrypted indicates a configuration drift or an in-progress operation.' `
            -Status 'Warning' -Severity 'High' `
            -Details "Cloud PC '$($cpc.DisplayName)' diskEncryptionState: $enc." `
            -Recommendation 'Validate the Cloud PC reaches an encrypted state. If the tenant requires customer-managed keys, configure a Disk Encryption Set referenced by the provisioning policy.' `
            -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/customer-managed-keys-overview' `
            -Evidence @{ Name = $cpc.DisplayName; State = $enc }))
    }
}

# IMG-008: Custom device image build status failed
foreach ($di in $Discovery.Inventory.DeviceImages) {
    if ($di.Status -eq 'failed' -or $di.ErrorCode) {
        [void]$AllChecks.Add((New-CheckResult `
            -Id "W365-IMG-008-$($di.Id)" -Category 'Images & App Delivery' `
            -Name "Failed custom image: $($di.DisplayName)" `
            -Description 'A custom Cloud PC image in failed state cannot be used by provisioning policies. Any policy referencing it will block all new provisioning.' `
            -Status 'Fail' -Severity 'High' `
            -Details "Custom image '$($di.DisplayName)' status: $($di.Status); errorCode: $($di.ErrorCode); details: $($di.StatusDetails)." `
            -Recommendation 'Investigate the image build failure (typically source image deletion, sysprep failure, or replication issue). Re-upload or rebuild the image and reassign provisioning policies.' `
            -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/device-images' `
            -Evidence @{ Name = $di.DisplayName; Status = $di.Status; ErrorCode = $di.ErrorCode }))
    }
}

# GOV-009: High-impact admin actions in audit events (last 30d)
$highImpactActivities = @('delete','reprovision','restore','endGracePeriod','setReviewStatus')
$auditHits = @($Discovery.Inventory.AuditEvents | Where-Object {
    $at = "$($_.ActivityType)"
    foreach ($k in $highImpactActivities) { if ($at -match $k) { return $true } }
    $false
})
if ($auditHits.Count -gt 0) {
    $top = $auditHits | Group-Object ActivityType | Sort-Object Count -Descending | Select-Object -First 5
    $summary = ($top | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ', '
    [void]$AllChecks.Add((New-CheckResult `
        -Id 'W365-GOV-009' -Category 'Governance & Operations' `
        -Name 'High-impact admin actions detected (30d)' `
        -Description 'Provisioning policy deletion, mass reprovisioning, restore, and grace-period termination are high-impact actions that warrant review against change records.' `
        -Status 'Warning' -Severity 'Medium' `
        -Details "$($auditHits.Count) high-impact event(s) in last 30 days. Top: $summary." `
        -Recommendation 'Cross-reference these audit events against approved change records. Forward Windows 365 audit events to Microsoft Sentinel for long-term retention and alerting.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/audit-logs' `
        -Evidence @{ EventCount = $auditHits.Count; TopActivities = $summary }))
}

# ─── INV-003: Edition / provisioning-type mix (no new calls) ───────────────
$typeGroups  = $Discovery.Inventory.CloudPCs | Group-Object ProvisioningType | Sort-Object Count -Descending
$typeSummary = if ($typeGroups) { ($typeGroups | ForEach-Object { "$(if ($_.Name) { $_.Name } else { '(unset)' })=$($_.Count)" }) -join ', ' } else { 'none' }
$flexCount   = @($Discovery.Inventory.CloudPCs | Where-Object { "$($_.ProvisioningType)" -eq 'sharedByUser' }).Count
$dedicated   = @($Discovery.Inventory.CloudPCs | Where-Object { "$($_.ProvisioningType)" -eq 'dedicated' }).Count
[void]$AllChecks.Add((New-CheckResult `
    -Id 'W365-INV-003-MIX' -Category 'Inventory & Topology' `
    -Name 'Edition / provisioning-type mix' `
    -Description 'The split of Cloud PCs across provisioning types — dedicated vs shared. provisioningType = sharedByUser is Windows 365 Flex (formerly Frontline) shared licensing, which changes concurrency, cost, and DR behaviour.' `
    -Status 'Pass' -Severity 'Low' `
    -Details "Provisioning-type mix across $($Discovery.Inventory.CloudPCs.Count) Cloud PC(s): $typeSummary. Dedicated: $dedicated. Flex/shared-by-user (sharedByUser): $flexCount." `
    -Recommendation 'Confirm the dedicated vs Flex split matches the licensing plan. Flex (sharedByUser) personas should be validated for concurrency limits and non-persistent data handling.' `
    -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/introduction-windows-365-flex' `
    -Evidence @{ Mix = $typeSummary; Dedicated = $dedicated; Flex = $flexCount; Total = $Discovery.Inventory.CloudPCs.Count }))

# ─── INV-004: Service plan (SKU) coverage (no new calls; uses the collected ServicePlans set) ──
$planCount  = $Discovery.Inventory.ServicePlans.Count
$plansInUse = @($Discovery.Inventory.CloudPCs | ForEach-Object { $_.ServicePlanId } | Where-Object { $_ } | Select-Object -Unique)
$skuSummary = if ($Discovery.Inventory.ServicePlans) { (@($Discovery.Inventory.ServicePlans | ForEach-Object { $_.DisplayName }) -join ', ') } else { 'none returned' }
[void]$AllChecks.Add((New-CheckResult `
    -Id 'W365-INV-004' -Category 'Inventory & Topology' `
    -Name 'Service plan (SKU) coverage' `
    -Description 'The set of Windows 365 service plans (SKUs) available to the tenant and how many are actually in use. A wide unused SKU spread can complicate sizing and cost governance.' `
    -Status $(if ($planCount -gt 0) { 'Pass' } else { 'Warning' }) `
    -Severity 'Medium' `
    -Details "$planCount service plan(s) available; $($plansInUse.Count) distinct SKU(s) in use across provisioned Cloud PCs. Available: $skuSummary." `
    -Recommendation 'Standardise on a small set of vetted SKUs per persona. Right-size using the Cloud PC recommendation/utilization reports before adding new SKUs.' `
    -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/planning-guide' `
    -Evidence @{ PlansAvailable = $planCount; SkusInUse = $plansInUse.Count }))

# ─── PROV-006: Domain-join type (Entra vs Hybrid) consistency; Entra preferred ─────────────
foreach ($pp in $Discovery.Inventory.ProvisioningPolicies) {
    $djTypes = @($pp.DomainJoinConfigurations | ForEach-Object { "$($_.domainJoinType)" } | Where-Object { $_ } | Select-Object -Unique)
    if ($djTypes.Count -eq 0) {
        $st = 'Warning'; $detail = 'no domainJoinConfigurations returned (unexpected — every policy must define a join type)'
    } elseif ($djTypes -contains 'hybridAzureADJoin') {
        $st = 'Warning'
        $detail = if ($djTypes.Count -gt 1) { "mixed join types: $($djTypes -join ', ') — Hybrid Entra Join present" } else { 'hybridAzureADJoin (Hybrid Entra Join)' }
    } else {
        $st = 'Pass'; $detail = ($djTypes -join ', ')
    }
    [void]$AllChecks.Add((New-CheckResult `
        -Id "W365-PROV-006-$($pp.Id)" -Category 'Identity & Access' `
        -Name "Domain-join type: $($pp.DisplayName)" `
        -Description 'Microsoft Entra join is the recommended, simplest join type for Cloud PCs and is required for SSO and passwordless. Hybrid Entra Join adds on-prem AD/line-of-sight dependencies and is only for legacy app constraints.' `
        -Status $st -Severity 'Medium' `
        -Details "Policy '$($pp.DisplayName)' domain-join type: $detail." `
        -Recommendation 'Prefer Microsoft Entra Join. Only use Hybrid Entra Join where an on-prem-bound dependency genuinely requires it, and keep join types consistent across the estate.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/identity-authentication' `
        -Evidence @{ PolicyName = $pp.DisplayName; JoinTypes = $djTypes }))
}

# ─── PROV-010: Cloud Apps / shared (Flex) posture per provisioning policy ───────────────────
foreach ($pp in $Discovery.Inventory.ProvisioningPolicies) {
    $pt = "$($pp.ProvisioningType)"
    if ($pt -notin @('sharedByUser','sharedByEntraGroup','shared')) { continue }
    [void]$AllChecks.Add((New-CheckResult `
        -Id "W365-PROV-010-$($pp.Id)" -Category 'Provisioning Policies' `
        -Name "Shared / Flex posture: $($pp.DisplayName)" `
        -Description 'Shared provisioning (sharedByUser = Windows 365 Flex, formerly Frontline) delivers non-persistent, concurrency-limited Cloud PCs. It suits task/shift and app-delivery (Cloud Apps) personas but requires explicit concurrency sizing and data-persistence handling.' `
        -Status 'Warning' -Severity 'Low' `
        -Details "Policy '$($pp.DisplayName)' provisioningType = $pt (shared/Flex). Validate concurrency limits, non-persistence, and whether an app-delivery (Cloud Apps) model fits the persona." `
        -Recommendation 'Size the Flex license concurrency for the assigned population, confirm profile/data persistence expectations, and consider Cloud Apps delivery for app-only personas.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/introduction-windows-365-flex' `
        -Evidence @{ PolicyName = $pp.DisplayName; ProvisioningType = $pt }))
}

# ═══════════════════════════════════════════════════════════════════════════
# REPORTS API  (POST /beta/deviceManagement/virtualEndpoint/reports/* — CloudPC.Read.All)
# Report actions are beta-only. Several spec-named actions are deprecated/renamed — current action
# names are used and each family degrades to a Status 'Error' CheckResult on failure.
# ═══════════════════════════════════════════════════════════════════════════
Write-Status "Reports API (Cloud PC)" -Level 'SECTION'

# COST-002 (right-sizing) + MON-008-REC (recommendations) + COST-001 report fallback.
# Spec named 'getCloudPcRecommendationReports'; current action is 'retrieveCloudPcRecommendationReports'.
try {
    $recReport = Invoke-GraphReport -Action 'retrieveCloudPcRecommendationReports' -Body @{ top = 25 }
    $recRows = Get-ReportRowCount $recReport
    Write-Status "Recommendation report rows: $recRows" -Level 'SUCCESS'
    [void]$AllChecks.Add((New-CheckResult `
        -Id 'W365-COST-002' -Category 'Cost & Optimization' `
        -Name 'Cloud PC right-sizing recommendations' `
        -Description 'The Cloud PC recommendation report surfaces under- and over-utilised Cloud PCs for SKU right-sizing and license reclaim.' `
        -Status $(if ($recRows -gt 0) { 'Warning' } else { 'Pass' }) `
        -Severity 'Medium' `
        -Details "Recommendation report returned $([math]::Max(0,$recRows)) row(s). Non-zero rows indicate right-sizing/reclaim opportunities to review." `
        -Recommendation 'Review the recommendation report in Intune > Reports > Cloud PC and act on down-size / reclaim candidates.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/report-cloud-pc-utilization' `
        -Evidence @{ Rows = $recRows }))
    [void]$AllChecks.Add((New-CheckResult `
        -Id 'W365-MON-008-REC' -Category 'Monitoring & Diagnostics' `
        -Name 'Cloud PC recommendation reporting available' `
        -Description 'Availability of the Cloud PC recommendation/usage reporting surface for ongoing optimisation monitoring.' `
        -Status 'Pass' -Severity 'Low' `
        -Details "Recommendation reporting reachable ($([math]::Max(0,$recRows)) row(s))." `
        -Recommendation 'Schedule periodic review of Cloud PC recommendation reports as part of monthly operations.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/report-cloud-pc-utilization' `
        -Evidence @{ Rows = $recRows }))
    if ($CpcNoLoginCount -gt 0) {
        [void]$AllChecks.Add((New-CheckResult `
            -Id 'W365-COST-001-REPORT' -Category 'Cost & Optimization' `
            -Name 'Inactive Cloud PC report fallback' `
            -Description 'Some Cloud PCs returned no LastLoginResult, so inactivity for those must be confirmed via the Cloud PC inactivity/recommendation report rather than the per-device login signal.' `
            -Status 'Warning' -Severity 'Low' `
            -Details "$CpcNoLoginCount Cloud PC(s) had no login data. Recommendation/inactivity report is reachable ($([math]::Max(0,$recRows)) row(s)) — use it to confirm inactivity for those devices." `
            -Recommendation 'Cross-check the report''s inactive/underused rows against the Cloud PCs lacking login telemetry before reclaiming licenses.' `
            -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/report-cloud-pc-utilization' `
            -Evidence @{ CloudPcsWithoutLoginData = $CpcNoLoginCount; ReportRows = $recRows }))
    }
} catch {
    Write-Status "Recommendation report unavailable: $($_.Exception.Message)" -Level 'WARN'
    foreach ($eid in @('W365-COST-002','W365-MON-008-REC')) {
        [void]$AllChecks.Add((New-CheckResult `
            -Id $eid -Category 'Cost & Optimization' `
            -Name 'Cloud PC recommendation report unavailable' `
            -Description 'The Cloud PC recommendation report (retrieveCloudPcRecommendationReports) could not be retrieved.' `
            -Status 'Error' -Severity 'Medium' `
            -Details "POST reports/retrieveCloudPcRecommendationReports failed: $($_.Exception.Message)" `
            -Recommendation 'Confirm the signed-in account holds CloudPC.Read.All. Note: this recommendation action was deprecated by Microsoft (2025) and may be unavailable in some tenants; review recommendations in the Intune portal instead.' `
            -Reference 'https://learn.microsoft.com/en-us/graph/api/resources/cloudpcreports?view=graph-rest-beta' `
            -Evidence @{ Action = 'retrieveCloudPcRecommendationReports'; RequiredScope = 'CloudPC.Read.All'; Error = $_.Exception.Message }))
    }
    if ($CpcNoLoginCount -gt 0) {
        [void]$AllChecks.Add((New-CheckResult `
            -Id 'W365-COST-001-REPORT' -Category 'Cost & Optimization' `
            -Name 'Inactive Cloud PC report fallback unavailable' `
            -Status 'Error' -Severity 'Low' `
            -Description 'Login data was missing for some Cloud PCs and the inactivity report fallback could not be retrieved.' `
            -Details "$CpcNoLoginCount Cloud PC(s) lacked login data and the report fallback failed: $($_.Exception.Message)" `
            -Recommendation 'Grant CloudPC.Read.All and re-run, or confirm inactivity manually in the Intune Cloud PC reports.' `
            -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/report-cloud-pc-utilization' `
            -Evidence @{ CloudPcsWithoutLoginData = $CpcNoLoginCount; RequiredScope = 'CloudPC.Read.All' }))
    }
}

# MON-002-Q (connection quality) + UX-002-CONN (connection round-trip UX).
# Spec named 'getConnectionQualityReports' (deprecated); current action is 'retrieveConnectionQualityReports'.
try {
    $cqReport = Invoke-GraphReport -Action 'retrieveConnectionQualityReports' -Body @{ top = 25 }
    $cqRows = Get-ReportRowCount $cqReport
    Write-Status "Connection quality report rows: $cqRows" -Level 'SUCCESS'
    [void]$AllChecks.Add((New-CheckResult `
        -Id 'W365-MON-002-Q' -Category 'Monitoring & Diagnostics' `
        -Name 'Connection quality reporting' `
        -Description 'The connection quality report exposes round-trip time, available bandwidth, and gateway region per Cloud PC connection — the primary telemetry for remoting experience.' `
        -Status $(if ($cqRows -ge 0) { 'Pass' } else { 'Warning' }) `
        -Severity 'Medium' `
        -Details "Connection quality report reachable ($([math]::Max(0,$cqRows)) row(s))." `
        -Recommendation 'Monitor connection quality trends; investigate personas/regions with high round-trip time or low bandwidth (RDP Shortpath, gateway region, ANC placement).' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/report-connection-quality' `
        -Evidence @{ Rows = $cqRows }))
    [void]$AllChecks.Add((New-CheckResult `
        -Id 'W365-UX-002-CONN' -Category 'User Experience' `
        -Name 'Connection quality / round-trip posture' `
        -Description 'End-user remoting experience is dominated by connection round-trip time and bandwidth; the connection quality report is the objective UX signal.' `
        -Status $(if ($cqRows -ge 0) { 'Pass' } else { 'Warning' }) `
        -Severity 'Medium' `
        -Details "Connection quality telemetry available ($([math]::Max(0,$cqRows)) row(s)) for UX round-trip analysis." `
        -Recommendation 'Baseline acceptable round-trip time per region and alert on regressions; enable RDP Shortpath and place ANCs close to users.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/report-connection-quality' `
        -Evidence @{ Rows = $cqRows }))
} catch {
    Write-Status "Connection quality report unavailable: $($_.Exception.Message)" -Level 'WARN'
    [void]$AllChecks.Add((New-CheckResult `
        -Id 'W365-MON-002-Q' -Category 'Monitoring & Diagnostics' `
        -Name 'Connection quality report unavailable' `
        -Status 'Error' -Severity 'Medium' `
        -Description 'The connection quality report (retrieveConnectionQualityReports) could not be retrieved.' `
        -Details "POST reports/retrieveConnectionQualityReports failed: $($_.Exception.Message)" `
        -Recommendation 'Confirm the signed-in account holds CloudPC.Read.All and re-run.' `
        -Reference 'https://learn.microsoft.com/en-us/graph/api/cloudpcreports-retrieveconnectionqualityreports?view=graph-rest-beta' `
        -Evidence @{ Action = 'retrieveConnectionQualityReports'; RequiredScope = 'CloudPC.Read.All'; Error = $_.Exception.Message }))
    [void]$AllChecks.Add((New-CheckResult `
        -Id 'W365-UX-002-CONN' -Category 'User Experience' `
        -Name 'Connection quality UX signal unavailable' `
        -Status 'Error' -Severity 'Medium' `
        -Description 'Connection quality telemetry for UX analysis could not be retrieved.' `
        -Details "POST reports/retrieveConnectionQualityReports failed: $($_.Exception.Message)" `
        -Recommendation 'Confirm the signed-in account holds CloudPC.Read.All and re-run.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/report-connection-quality' `
        -Evidence @{ Action = 'retrieveConnectionQualityReports'; RequiredScope = 'CloudPC.Read.All'; Error = $_.Exception.Message }))
}

# MON-010-R (resource performance). Spec named 'getResourcePerformanceReport' (does not exist);
# current action for Cloud PC performance is 'retrieveCloudPcTenantMetricsReport'.
try {
    $perfReport = Invoke-GraphReport -Action 'retrieveCloudPcTenantMetricsReport' -Body @{ top = 25 }
    $perfRows = Get-ReportRowCount $perfReport
    Write-Status "Resource performance report rows: $perfRows" -Level 'SUCCESS'
    [void]$AllChecks.Add((New-CheckResult `
        -Id 'W365-MON-010-R' -Category 'Monitoring & Diagnostics' `
        -Name 'Resource performance reporting' `
        -Description 'The Cloud PC tenant metrics / resource performance report exposes CPU, RAM, and disk performance signals used to detect under-provisioned SKUs and noisy-neighbour effects.' `
        -Status $(if ($perfRows -ge 0) { 'Pass' } else { 'Warning' }) `
        -Severity 'Medium' `
        -Details "Resource performance report reachable ($([math]::Max(0,$perfRows)) row(s))." `
        -Recommendation 'Review resource performance to right-size SKUs; sustained high CPU/RAM pressure indicates the persona needs a larger SKU.' `
        -Reference 'https://learn.microsoft.com/en-us/graph/api/cloudpcreports-retrievecloudpctenantmetricsreport?view=graph-rest-beta' `
        -Evidence @{ Rows = $perfRows }))
} catch {
    Write-Status "Resource performance report unavailable: $($_.Exception.Message)" -Level 'WARN'
    [void]$AllChecks.Add((New-CheckResult `
        -Id 'W365-MON-010-R' -Category 'Monitoring & Diagnostics' `
        -Name 'Resource performance report unavailable' `
        -Status 'Error' -Severity 'Medium' `
        -Description 'The Cloud PC resource performance report (retrieveCloudPcTenantMetricsReport) could not be retrieved.' `
        -Details "POST reports/retrieveCloudPcTenantMetricsReport failed: $($_.Exception.Message)" `
        -Recommendation 'Confirm the signed-in account holds CloudPC.Read.All and re-run.' `
        -Reference 'https://learn.microsoft.com/en-us/graph/api/cloudpcreports-retrievecloudpctenantmetricsreport?view=graph-rest-beta' `
        -Evidence @{ Action = 'retrieveCloudPcTenantMetricsReport'; RequiredScope = 'CloudPC.Read.All'; Error = $_.Exception.Message }))
}

# ═══════════════════════════════════════════════════════════════════════════
# INTUNE  (/beta/deviceManagement — DeviceManagementManagedDevices.Read.All +
#          DeviceManagementConfiguration.Read.All, both already consented but previously unused)
# ═══════════════════════════════════════════════════════════════════════════
Write-Status "Intune (Cloud PC managed devices & policy)" -Level 'SECTION'

# SEC-002-MDE: Cloud PC managed-device health (compliance / Defender posture).
try {
    $cpcMde = @(Invoke-GraphPaged -Uri "$IntuneBase/managedDevices?`$filter=contains(model,'Cloud PC')")
    $mdeTotal = $cpcMde.Count
    $mdeNonCompliant = @($cpcMde | Where-Object { "$($_.complianceState)" -notin @('compliant','') }).Count
    Write-Status "Cloud PC managed devices: $mdeTotal ($mdeNonCompliant non-compliant)" -Level 'SUCCESS'
    [void]$AllChecks.Add((New-CheckResult `
        -Id 'W365-SEC-002-MDE' -Category 'Security & Compliance' `
        -Name 'Cloud PC managed-device health' `
        -Description 'Cloud PCs enrolled in Intune report compliance and (via Defender) protection state. Non-compliant Cloud PCs may be blocked by Conditional Access and indicate missing baseline/AV controls.' `
        -Status $(if ($mdeTotal -eq 0) { 'Warning' } elseif ($mdeNonCompliant -gt 0) { 'Warning' } else { 'Pass' }) `
        -Severity 'High' `
        -Details "$mdeTotal Cloud PC managed device(s) found; $mdeNonCompliant non-compliant." `
        -Recommendation 'Investigate non-compliant Cloud PCs (Defender onboarding, compliance policy failures) and ensure Defender for Endpoint is deployed to Cloud PCs.' `
        -Reference 'https://learn.microsoft.com/en-us/mem/intune/protect/device-compliance-get-started' `
        -Evidence @{ CloudPcDevices = $mdeTotal; NonCompliant = $mdeNonCompliant }))
} catch {
    Write-Status "Cloud PC managed devices unavailable: $($_.Exception.Message)" -Level 'WARN'
    [void]$AllChecks.Add((New-CheckResult `
        -Id 'W365-SEC-002-MDE' -Category 'Security & Compliance' `
        -Name 'Cloud PC managed-device health unavailable' `
        -Status 'Error' -Severity 'High' `
        -Description 'Cloud PC managed devices could not be enumerated from Intune.' `
        -Details "GET managedDevices (filter model contains 'Cloud PC') failed: $($_.Exception.Message)" `
        -Recommendation 'Confirm the signed-in account holds DeviceManagementManagedDevices.Read.All and re-run.' `
        -Reference 'https://learn.microsoft.com/en-us/graph/api/resources/intune-devices-manageddevice?view=graph-rest-beta' `
        -Evidence @{ RequiredScope = 'DeviceManagementManagedDevices.Read.All'; Error = $_.Exception.Message }))
}

# SEC-004-COMP: device compliance policies exist and are assigned.
try {
    $compPols = @(Invoke-GraphPaged -Uri "$IntuneBase/deviceCompliancePolicies?`$expand=assignments")
    $compAssigned = @($compPols | Where-Object { @($_.assignments).Count -gt 0 }).Count
    Write-Status "Compliance policies: $($compPols.Count) ($compAssigned assigned)" -Level 'SUCCESS'
    [void]$AllChecks.Add((New-CheckResult `
        -Id 'W365-SEC-004-COMP' -Category 'Security & Compliance' `
        -Name 'Device compliance policies present & assigned' `
        -Description 'Intune device compliance policies gate Cloud PC access via Conditional Access. Without an assigned compliance policy, compliance-based CA cannot protect Cloud PCs.' `
        -Status $(if ($compPols.Count -eq 0) { 'Fail' } elseif ($compAssigned -eq 0) { 'Warning' } else { 'Pass' }) `
        -Severity 'High' `
        -Details "$($compPols.Count) compliance policy/policies; $compAssigned assigned." `
        -Recommendation 'Author and assign at least one compliance policy covering Cloud PCs, then require compliant device in Conditional Access.' `
        -Reference 'https://learn.microsoft.com/en-us/mem/intune/protect/device-compliance-get-started' `
        -Evidence @{ Policies = $compPols.Count; Assigned = $compAssigned }))
} catch {
    Write-Status "Compliance policies unavailable: $($_.Exception.Message)" -Level 'WARN'
    [void]$AllChecks.Add((New-CheckResult `
        -Id 'W365-SEC-004-COMP' -Category 'Security & Compliance' `
        -Name 'Device compliance policies unavailable' `
        -Status 'Error' -Severity 'High' `
        -Description 'Intune device compliance policies could not be enumerated.' `
        -Details "GET deviceCompliancePolicies failed: $($_.Exception.Message)" `
        -Recommendation 'Confirm the signed-in account holds DeviceManagementConfiguration.Read.All and re-run.' `
        -Reference 'https://learn.microsoft.com/en-us/graph/api/resources/intune-deviceconfig-devicecompliancepolicy?view=graph-rest-beta' `
        -Evidence @{ RequiredScope = 'DeviceManagementConfiguration.Read.All'; Error = $_.Exception.Message }))
}

# SEC-003-BASE: security baseline / configuration profiles referencing Windows 365 / Cloud PC.
try {
    $cfgPols = @(Invoke-GraphPaged -Uri "$IntuneBase/configurationPolicies?`$expand=assignments")
    $w365Cfg = @($cfgPols | Where-Object { "$($_.name) $($_.description)" -match '(?i)cloud pc|windows 365|w365' })
    Write-Status "Config profiles: $($cfgPols.Count) ($($w365Cfg.Count) reference Windows 365)" -Level 'SUCCESS'
    [void]$AllChecks.Add((New-CheckResult `
        -Id 'W365-SEC-003-BASE' -Category 'Security & Compliance' `
        -Name 'Security baseline / config profiles for Windows 365' `
        -Description 'A Cloud PC security baseline (current baseline is the Windows 365 24H1 baseline) or dedicated configuration profiles harden Cloud PCs beyond defaults.' `
        -Status $(if ($cfgPols.Count -eq 0) { 'Warning' } elseif ($w365Cfg.Count -gt 0) { 'Pass' } else { 'Warning' }) `
        -Severity 'Medium' `
        -Details "$($cfgPols.Count) configuration policy/policies; $($w365Cfg.Count) explicitly reference Windows 365 / Cloud PC by name or description." `
        -Recommendation 'Apply the Windows 365 Security Baseline (24H1) and/or Cloud PC-scoped configuration profiles, and confirm they are assigned to Cloud PC groups.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/security-baseline' `
        -Evidence @{ ConfigPolicies = $cfgPols.Count; Windows365Referencing = $w365Cfg.Count }))
} catch {
    Write-Status "Config profiles unavailable: $($_.Exception.Message)" -Level 'WARN'
    [void]$AllChecks.Add((New-CheckResult `
        -Id 'W365-SEC-003-BASE' -Category 'Security & Compliance' `
        -Name 'Security baseline / config profiles unavailable' `
        -Status 'Error' -Severity 'Medium' `
        -Description 'Intune configuration profiles could not be enumerated.' `
        -Details "GET configurationPolicies failed: $($_.Exception.Message)" `
        -Recommendation 'Confirm the signed-in account holds DeviceManagementConfiguration.Read.All and re-run.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/security-baseline' `
        -Evidence @{ RequiredScope = 'DeviceManagementConfiguration.Read.All'; Error = $_.Exception.Message }))
}

# MON-001-EA: Endpoint Analytics device scores availability.
try {
    $eaScores = @(Invoke-GraphPaged -Uri "$IntuneBase/userExperienceAnalyticsDeviceScores?`$top=50")
    Write-Status "Endpoint Analytics device scores: $($eaScores.Count)" -Level 'SUCCESS'
    [void]$AllChecks.Add((New-CheckResult `
        -Id 'W365-MON-001-EA' -Category 'Monitoring & Diagnostics' `
        -Name 'Endpoint Analytics device scores' `
        -Description 'Endpoint Analytics scores (startup, app reliability, work-from-anywhere) provide the objective end-user experience baseline for Cloud PCs.' `
        -Status $(if ($eaScores.Count -gt 0) { 'Pass' } else { 'Warning' }) `
        -Severity 'Medium' `
        -Details "Endpoint Analytics returned $($eaScores.Count) device score record(s)." `
        -Recommendation 'Enable Endpoint Analytics for Cloud PCs and track startup/reliability scores; low scores flag SKU or image problems.' `
        -Reference 'https://learn.microsoft.com/en-us/mem/analytics/overview' `
        -Evidence @{ DeviceScores = $eaScores.Count }))
} catch {
    Write-Status "Endpoint Analytics unavailable: $($_.Exception.Message)" -Level 'WARN'
    [void]$AllChecks.Add((New-CheckResult `
        -Id 'W365-MON-001-EA' -Category 'Monitoring & Diagnostics' `
        -Name 'Endpoint Analytics unavailable' `
        -Status 'Error' -Severity 'Medium' `
        -Description 'Endpoint Analytics device scores could not be retrieved.' `
        -Details "GET userExperienceAnalyticsDeviceScores failed: $($_.Exception.Message)" `
        -Recommendation 'Confirm the signed-in account holds DeviceManagementConfiguration.Read.All and that Endpoint Analytics is enabled.' `
        -Reference 'https://learn.microsoft.com/en-us/mem/analytics/overview' `
        -Evidence @{ RequiredScope = 'DeviceManagementConfiguration.Read.All'; Error = $_.Exception.Message }))
}

# MON-005-UPD: software update status summary.
try {
    $updSummary = Invoke-MgGraphRequest -Method GET -Uri "$IntuneBase/softwareUpdateStatusSummary" -ErrorAction Stop
    $updCompliant = [int]("$($updSummary.compliantDeviceCount)" -replace '[^0-9]', '')
    $updNonCompliant = [int]("$($updSummary.nonCompliantDeviceCount)" -replace '[^0-9]', '')
    Write-Status "Update status: $updCompliant compliant / $updNonCompliant non-compliant" -Level 'SUCCESS'
    [void]$AllChecks.Add((New-CheckResult `
        -Id 'W365-MON-005-UPD' -Category 'Monitoring & Diagnostics' `
        -Name 'Software update compliance summary' `
        -Description 'The tenant software-update status summary indicates how many managed devices (including Cloud PCs) are current on updates — a core patch-hygiene signal.' `
        -Status $(if ($updNonCompliant -gt 0) { 'Warning' } else { 'Pass' }) `
        -Severity 'Medium' `
        -Details "Update status summary: $updCompliant compliant, $updNonCompliant non-compliant device(s)." `
        -Recommendation 'Drive update non-compliance to zero via Windows Autopatch / update rings scoped to Cloud PCs.' `
        -Reference 'https://learn.microsoft.com/en-us/graph/api/resources/intune-softwareupdate-softwareupdatestatussummary?view=graph-rest-beta' `
        -Evidence @{ Compliant = $updCompliant; NonCompliant = $updNonCompliant }))
} catch {
    Write-Status "Update status summary unavailable: $($_.Exception.Message)" -Level 'WARN'
    [void]$AllChecks.Add((New-CheckResult `
        -Id 'W365-MON-005-UPD' -Category 'Monitoring & Diagnostics' `
        -Name 'Software update compliance summary unavailable' `
        -Status 'Error' -Severity 'Medium' `
        -Description 'The software update status summary could not be retrieved.' `
        -Details "GET softwareUpdateStatusSummary failed: $($_.Exception.Message)" `
        -Recommendation 'Confirm the signed-in account holds DeviceManagementConfiguration.Read.All and re-run.' `
        -Reference 'https://learn.microsoft.com/en-us/graph/api/resources/intune-softwareupdate-softwareupdatestatussummary?view=graph-rest-beta' `
        -Evidence @{ RequiredScope = 'DeviceManagementConfiguration.Read.All'; Error = $_.Exception.Message }))
}

# ═══════════════════════════════════════════════════════════════════════════
# CONDITIONAL ACCESS  (/v1.0/identity/conditionalAccess/policies — Policy.Read.All, OPTIONAL tier)
# The entire family degrades to Status 'Error' naming Policy.Read.All when the scope is not granted.
# ═══════════════════════════════════════════════════════════════════════════
Write-Status "Conditional Access (Cloud PC sign-in)" -Level 'SECTION'
try {
    $caPolicies = @(Invoke-GraphPaged -Uri $CaPolicyUri)
    $caEnabled  = @($caPolicies | Where-Object { "$($_.state)" -eq 'enabled' })
    Write-Status "Conditional Access policies: $($caPolicies.Count) ($($caEnabled.Count) enabled)" -Level 'SUCCESS'

    # Helper: does an enabled policy target a Cloud PC sign-in app (explicitly or via 'All')?
    $targetsCloudPc = {
        param($p)
        $inc = @($p.conditions.applications.includeApplications)
        if ($inc -contains 'All') { return $true }
        return (@($inc | Where-Object { $CloudPcSignInAppIds -contains $_ }).Count -gt 0)
    }
    $cpcTargeting = @($caEnabled | Where-Object { & $targetsCloudPc $_ })
    $wclTargeting = @($caEnabled | Where-Object {
        $inc = @($_.conditions.applications.includeApplications)
        ($inc -contains 'All') -or ($inc -contains $AppIdWindowsCloudLogin)
    })

    # IAM-003-CA: at least one enabled CA policy targets Cloud PC sign-in apps.
    [void]$AllChecks.Add((New-CheckResult `
        -Id 'W365-IAM-003-CA' -Category 'Identity & Access' `
        -Name 'Conditional Access targets Cloud PC sign-in' `
        -Description 'Conditional Access must target the apps used to sign in to Cloud PCs (Windows Cloud Login, Azure Virtual Desktop, Windows 365 portal), or all apps, to enforce access controls on Cloud PC sessions.' `
        -Status $(if ($cpcTargeting.Count -gt 0) { 'Pass' } else { 'Fail' }) `
        -Severity 'High' `
        -Details "$($cpcTargeting.Count) enabled CA policy/policies target Cloud PC sign-in apps (or All apps) out of $($caEnabled.Count) enabled." `
        -Recommendation 'Create/scope a Conditional Access policy to the Windows Cloud Login and Azure Virtual Desktop apps (or All apps) covering Cloud PC users.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/set-conditional-access-policies' `
        -Evidence @{ Targeting = $cpcTargeting.Count; EnabledPolicies = $caEnabled.Count; AppIds = $CloudPcSignInAppIds }))

    # IAM-004-MFA: targeting policies require MFA.
    $mfaTargeting = @($cpcTargeting | Where-Object {
        @($_.grantControls.builtInControls) -contains 'mfa'
    })
    [void]$AllChecks.Add((New-CheckResult `
        -Id 'W365-IAM-004-MFA' -Category 'Identity & Access' `
        -Name 'MFA enforced for Cloud PC sign-in' `
        -Description 'Cloud PC sign-in should require multi-factor authentication via a Conditional Access grant control.' `
        -Status $(if ($cpcTargeting.Count -eq 0) { 'Fail' } elseif ($mfaTargeting.Count -gt 0) { 'Pass' } else { 'Fail' }) `
        -Severity 'High' `
        -Details "$($mfaTargeting.Count) of $($cpcTargeting.Count) Cloud PC-targeting CA policy/policies require MFA." `
        -Recommendation 'Require multi-factor authentication (or a phishing-resistant authentication strength) in the Conditional Access policy covering Cloud PC sign-in apps.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/set-conditional-access-policies' `
        -Evidence @{ MfaPolicies = $mfaTargeting.Count; Targeting = $cpcTargeting.Count }))

    # W365-IAM-010 (NEW): Windows Cloud Login app coverage specifically.
    [void]$AllChecks.Add((New-CheckResult `
        -Id 'W365-IAM-010' -Category 'Identity & Access' `
        -Name 'Windows Cloud Login coverage' `
        -Description 'Windows Cloud Login (app 270efc09-cd0d-444b-a71f-39af4910ec45) is the identity app that brokers Cloud PC sign-in. Conditional Access should explicitly cover it (directly or via All apps).' `
        -Status $(if ($wclTargeting.Count -gt 0) { 'Pass' } else { 'Fail' }) `
        -Severity 'High' `
        -Details "$($wclTargeting.Count) enabled CA policy/policies cover the Windows Cloud Login app (directly or via All apps)." `
        -Recommendation 'Ensure a Conditional Access policy explicitly includes the Windows Cloud Login app so Cloud PC sign-in is always governed by CA.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/set-conditional-access-policies' `
        -Evidence @{ Covering = $wclTargeting.Count; WindowsCloudLoginAppId = $AppIdWindowsCloudLogin }))

    # W365-IAM-011 (NEW): token protection / sign-in frequency session controls on Cloud PC access.
    $tokenProt = @($cpcTargeting | Where-Object { $_.sessionControls.secureSignInSession.isEnabled -eq $true })
    $signInFreq = @($cpcTargeting | Where-Object {
        $sif = $_.sessionControls.signInFrequency
        $sif -and ("$($sif.frequencyInterval)" -eq 'everyTime' -or $sif.isEnabled -eq $true)
    })
    [void]$AllChecks.Add((New-CheckResult `
        -Id 'W365-IAM-011' -Category 'Identity & Access' `
        -Name 'Token protection / sign-in frequency for Cloud PC' `
        -Description 'Session controls — token protection (secureSignInSession) and sign-in frequency — reduce token-theft and session-persistence risk for Cloud PC access.' `
        -Status $(if ($cpcTargeting.Count -eq 0) { 'Warning' } elseif ($tokenProt.Count -gt 0 -or $signInFreq.Count -gt 0) { 'Pass' } else { 'Warning' }) `
        -Severity 'Medium' `
        -Details "Cloud PC-targeting CA policies with token protection: $($tokenProt.Count); with sign-in frequency configured: $($signInFreq.Count)." `
        -Recommendation 'Add token protection (sign-in session) and an appropriate sign-in frequency to the Conditional Access policy covering Cloud PC sign-in.' `
        -Reference 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/concept-token-protection' `
        -Evidence @{ TokenProtection = $tokenProt.Count; SignInFrequency = $signInFreq.Count; Targeting = $cpcTargeting.Count }))
} catch {
    Write-Status "Conditional Access unavailable (needs Policy.Read.All): $($_.Exception.Message)" -Level 'WARN'
    foreach ($ca in @(
        @{ Id = 'W365-IAM-003-CA'; Name = 'Conditional Access targeting'; Sev = 'High' },
        @{ Id = 'W365-IAM-004-MFA'; Name = 'MFA for Cloud PC sign-in'; Sev = 'High' },
        @{ Id = 'W365-IAM-010'; Name = 'Windows Cloud Login coverage'; Sev = 'High' },
        @{ Id = 'W365-IAM-011'; Name = 'Token protection / sign-in frequency'; Sev = 'Medium' }
    )) {
        [void]$AllChecks.Add((New-CheckResult `
            -Id $ca.Id -Category 'Identity & Access' `
            -Name "$($ca.Name) — not assessed" `
            -Status 'Error' -Severity $ca.Sev `
            -Description 'Conditional Access policies could not be read, so Cloud PC access-control posture was not assessed.' `
            -Details "GET conditionalAccess/policies failed: $($_.Exception.Message)" `
            -Recommendation 'Grant the optional Policy.Read.All scope (admin consent) and re-run to assess Conditional Access coverage for Cloud PC sign-in.' `
            -Reference 'https://learn.microsoft.com/en-us/graph/api/resources/conditionalaccesspolicy' `
            -Evidence @{ RequiredScope = 'Policy.Read.All'; Error = $_.Exception.Message }))
    }
}

# Finalise
$Discovery.CheckResults = $AllChecks.ToArray()

# ═══════════════════════════════════════════════════════════════════════════
# SUMMARY & EXPORT
# ═══════════════════════════════════════════════════════════════════════════

Write-Status "Summary" -Level 'SECTION'
Write-Metric 'Cloud PCs'                $Discovery.Inventory.CloudPCs.Count
Write-Metric 'Provisioning Policies'    $Discovery.Inventory.ProvisioningPolicies.Count
Write-Metric 'User Settings'            $Discovery.Inventory.UserSettings.Count
Write-Metric 'Network Connections'      $Discovery.Inventory.AzureNetworkConnections.Count
Write-Metric 'Custom Device Images'     $Discovery.Inventory.DeviceImages.Count
Write-Metric 'Gallery Images'           $Discovery.Inventory.GalleryImages.Count
Write-Metric 'Service Plans'            $Discovery.Inventory.ServicePlans.Count
Write-Metric 'Audit Events (30d)'       $Discovery.Inventory.AuditEvents.Count

$pass = @($AllChecks | Where-Object { $_.Status -eq 'Pass' }).Count
$warn = @($AllChecks | Where-Object { $_.Status -eq 'Warning' }).Count
$fail = @($AllChecks | Where-Object { $_.Status -eq 'Fail' }).Count
$err  = @($AllChecks | Where-Object { $_.Status -eq 'Error' }).Count
Write-Host ""
Write-Metric 'Checks: Pass'  $pass '+'
Write-Metric 'Checks: Warn'  $warn '!'
Write-Metric 'Checks: Fail'  $fail 'X'
Write-Metric 'Checks: Error' $err  'E'

# Resolve output path
if (-not $OutputPath) {
    $assessDir = Join-Path $ScriptRoot 'assessments'
    if (-not (Test-Path $assessDir)) { New-Item -ItemType Directory -Path $assessDir | Out-Null }
    $OutputPath = Join-Path $assessDir ("discovery_{0}.json" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
} else {
    $parent = Split-Path -Parent $OutputPath
    if ($parent -and -not (Test-Path $parent)) { New-Item -ItemType Directory -Path $parent | Out-Null }
}

try {
    $json = $Discovery | ConvertTo-Json -Depth 12
    [System.IO.File]::WriteAllText($OutputPath, $json, [System.Text.UTF8Encoding]::new($false))
    Write-Host ""
    Write-Status "Discovery written to: $OutputPath" -Level 'SUCCESS'
} catch {
    Write-Status "Failed to write discovery JSON: $($_.Exception.Message)" -Level 'ERROR'
    exit 1
}

if ($Discovery.Errors.Count -gt 0) {
    Write-Host ""
    Write-Status "Completed with $($Discovery.Errors.Count) error(s) — review the Errors section in the JSON." -Level 'WARN'
}

Write-Host ""
