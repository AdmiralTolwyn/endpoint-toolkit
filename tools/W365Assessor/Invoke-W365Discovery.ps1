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
    Version: 0.1.0
    Date   : 2026-04-23

    Required Graph scopes:
      CloudPC.Read.All
      DeviceManagementConfiguration.Read.All
      DeviceManagementManagedDevices.Read.All
      Directory.Read.All

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

$ScriptVersion = '0.1.0'
$GraphBase     = 'https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint'

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
            Connect-MgGraph -Scopes $Scopes -TenantId $TenantId -NoWelcome -ErrorAction Stop | Out-Null
        } else {
            Connect-MgGraph -Scopes $Scopes -NoWelcome -ErrorAction Stop | Out-Null
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
    param([string]$Section, [string]$Message)
    Write-Status "$Section : $Message" -Level 'ERROR'
    $Discovery.Errors += "[$Section] $Message"
}

# ─── CLOUD PCs ────────────────────────────────────────────────────────────
Write-Status "Cloud PCs" -Level 'SECTION'
$CloudPCs = @()
try {
    $CloudPCs = @(Invoke-GraphPaged -Uri "$GraphBase/cloudPCs")
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
    $ProvPols = @(Invoke-GraphPaged -Uri "$GraphBase/provisioningPolicies?`$expand=assignments")
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
    $UserSet = @(Invoke-GraphPaged -Uri "$GraphBase/userSettings?`$expand=assignments")
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
            HealthCheckStatusDetails     = $anc.healthCheckStatusDetails
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
    $since = (Get-Date).AddDays(-30).ToString('o')
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

# PROV-002: SSO recommended on every provisioning policy
foreach ($pp in $Discovery.Inventory.ProvisioningPolicies) {
    [void]$AllChecks.Add((New-CheckResult `
        -Id "W365-PROV-002-$($pp.Id)" -Category 'Identity & Access' `
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

# PROV-005: Grace period configured (default = 60min; 0 means immediate deprovision = data-loss risk)
foreach ($pp in $Discovery.Inventory.ProvisioningPolicies) {
    $gp = [int]($pp.GracePeriodInHours | ForEach-Object { if ($_) { $_ } else { 0 } })
    $st = if ($gp -le 0) { 'Warning' } elseif ($gp -ge 1 -and $gp -le 168) { 'Pass' } else { 'Warning' }
    $details = "GracePeriodInHours = $gp on '$($pp.DisplayName)'."
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
    # Parse failed sub-checks from health details for richer context
    $failedTests = @()
    if ($anc.HealthCheckStatusDetails -and $anc.HealthCheckStatusDetails.endDateTime) {
        # Newer schema: status details has lastHealthCheckDateTime + sub-checks array
        if ($anc.HealthCheckStatusDetails.failedHealthCheckItems) {
            $failedTests = @($anc.HealthCheckStatusDetails.failedHealthCheckItems | ForEach-Object { $_.displayName })
        }
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

# IMG-001: Custom image age
$now = Get-Date
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

# CPC-001: Cloud PCs in failed / grace period / not-provisioned states
$badStates = @('failed','provisionedWithWarnings','notProvisioned','inGracePeriod','provisioning','deprovisioning','resizing','restoring','pendingProvision','unknown')
foreach ($cpc in $Discovery.Inventory.CloudPCs) {
    if ($cpc.Status -in @('failed')) {
        [void]$AllChecks.Add((New-CheckResult `
            -Id "W365-CPC-001-$($cpc.Id)" -Category 'Inventory & Topology' `
            -Name "Cloud PC in Failed state: $($cpc.DisplayName)" `
            -Description 'Cloud PCs in Failed state are not usable and may indicate a provisioning or licensing issue.' `
            -Status 'Fail' -Severity 'High' `
            -Details "Cloud PC '$($cpc.DisplayName)' for $($cpc.UserPrincipalName) is in status: $($cpc.Status)." `
            -Recommendation 'Investigate via Intune > Devices > Cloud PCs. Reprovision or open a support ticket if persistent.' `
            -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/known-issues-provisioning' `
            -Evidence @{ Name = $cpc.DisplayName; UPN = $cpc.UserPrincipalName; Status = $cpc.Status }))
    } elseif ($cpc.Status -eq 'inGracePeriod') {
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

# CPC-003: Cloud PCs likely inactive (last login = unknown can mean never used)
foreach ($cpc in $Discovery.Inventory.CloudPCs) {
    $lastTime = $cpc.LastModifiedDateTime
    if (-not $lastTime) { continue }
    try { $idle = ($now - [datetime]$lastTime).Days } catch { continue }
    if ($idle -gt $InactiveDays -and $cpc.Status -in @('provisioned','provisionedWithWarnings')) {
        [void]$AllChecks.Add((New-CheckResult `
            -Id "W365-CPC-003-$($cpc.Id)" -Category 'Cost & Optimization' `
            -Name "Inactive Cloud PC: $($cpc.DisplayName)" `
            -Description 'Provisioned Cloud PCs that have not been modified in a long time may be candidates for license reclaim or downsizing.' `
            -Status 'Warning' -Severity 'Low' `
            -Details "Cloud PC '$($cpc.DisplayName)' has been quiet for $idle day(s) (threshold $InactiveDays)." `
            -Recommendation 'Use Endpoint analytics utilization reports to confirm inactivity, then reclaim the license or downsize the SKU.' `
            -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/report-cloud-pc-utilization' `
            -Evidence @{ Name = $cpc.DisplayName; UPN = $cpc.UserPrincipalName; IdleDays = $idle }))
    }
}

# PROV-008: Provisioning policy missing a Cloud PC naming template
foreach ($pp in $Discovery.Inventory.ProvisioningPolicies) {
    $hasTpl = -not [string]::IsNullOrWhiteSpace($pp.CloudPcNamingTemplate)
    [void]$AllChecks.Add((New-CheckResult `
        -Id "W365-PROV-008-$($pp.Id)" -Category 'Provisioning Policies' `
        -Name "Cloud PC naming template: $($pp.DisplayName)" `
        -Description 'A naming template (e.g. CPC-%USERNAME:5%-%RAND:5%) produces predictable, traceable Cloud PC device names. Without it, Cloud PCs receive auto-generated names that are hard to correlate to users during incidents.' `
        -Status $(if ($hasTpl) { 'Pass' } else { 'Warning' }) `
        -Severity 'Low' `
        -Details "Policy '$($pp.DisplayName)' template: $(if ($hasTpl) { $pp.CloudPcNamingTemplate } else { '(none)' })." `
        -Recommendation 'Edit the provisioning policy and set a Cloud PC naming template that encodes user identity and environment (e.g. ''CPC-%USERNAME:5%-%RAND:5%'').' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/create-provisioning-policy' `
        -Evidence @{ PolicyName = $pp.DisplayName; Template = $pp.CloudPcNamingTemplate }))
}

# PROV-009: Provisioning policy Windows language/region setting present
foreach ($pp in $Discovery.Inventory.ProvisioningPolicies) {
    $hasLang = $pp.WindowsSetting -and -not [string]::IsNullOrWhiteSpace($pp.WindowsSetting.locale)
    [void]$AllChecks.Add((New-CheckResult `
        -Id "W365-PROV-009-$($pp.Id)" -Category 'Provisioning Policies' `
        -Name "Windows locale set: $($pp.DisplayName)" `
        -Description 'Setting Windows locale on the provisioning policy ensures Cloud PCs ship with the correct OS language without first-login MUI churn.' `
        -Status $(if ($hasLang) { 'Pass' } else { 'Warning' }) `
        -Severity 'Low' `
        -Details "Policy '$($pp.DisplayName)' locale: $(if ($hasLang) { $pp.WindowsSetting.locale } else { '(default en-US)' })." `
        -Recommendation 'Configure the Windows setting on the provisioning policy with the appropriate locale for the user persona/region.' `
        -Reference 'https://learn.microsoft.com/en-us/windows-365/enterprise/create-provisioning-policy' `
        -Evidence @{ PolicyName = $pp.DisplayName; Locale = $pp.WindowsSetting.locale }))
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
