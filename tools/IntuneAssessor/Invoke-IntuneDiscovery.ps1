#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath = (Join-Path $PWD ('intune-discovery-{0}.json' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))),
    [switch]$IncludeRbac,
    [switch]$IncludeAudit,
    [switch]$IncludeConfiguration,
    [switch]$IncludeEntra,
    [switch]$IncludeRecoveryMetadata,
    [switch]$IncludeEnrollment,
    [switch]$IncludeApple,
    [switch]$IncludeMam,
    [switch]$IncludeMamLaunch,
    [switch]$IncludeRemoteHelp,
    [switch]$IncludeConnectors,
    [switch]$IncludeAppConfiguration,
    [switch]$IncludePlatformCompliance,
    [switch]$IncludeTunnel,
    [string[]]$EndpointEvidencePaths = @(),
    [string]$DefenderEvidencePath,
    [string[]]$AppControlPolicyPaths = @(),
    [datetime]$AuditSinceUtc = ([datetime]::UtcNow.AddDays(-7)),
    [switch]$UseExistingConnection,
    [securestring]$GraphAccessToken,
    [string]$Assessor,
    [string]$ScopeDescription,
    [ValidateRange(1, 87600)][int]$MaxCollectionAgeHours,
    [ValidateRange(1, 87600)][int]$MaxPolicyReportAgeHours,
    [ValidateRange(1, 87600)][int]$MaxDeviceSyncAgeDays,
    [switch]$LibraryOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'IntuneExpansion.ps1')
$script:IntuneGraphContracts = @(foreach ($Contract in (Get-Content -LiteralPath (Join-Path $PSScriptRoot 'GraphContracts.json') -Raw | ConvertFrom-Json)) { $Contract })

function Get-IntuneValue {
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return ,$Object[$Name] }
    } elseif ($null -ne $Object.PSObject.Properties[$Name]) {
        return ,$Object.$Name
    }
    return $Default
}

function Get-IntuneGraphContract {
    param([string]$Path, [string]$Module)
    foreach ($Contract in $script:IntuneGraphContracts) {
        $Pattern = '^/' + [regex]::Escape($Contract.version + '/' + $Contract.path).Replace('\{id}', '[^/]+') + '$'
        if ($Path -cmatch $Pattern -and (-not $Module -or $Module -ceq $Contract.module)) { return $Contract }
    }
    return $null
}

function Test-IntuneUri {
    param([string]$Uri, [switch]$AllowExpansion)
    $Parsed = $null
    if (-not [uri]::TryCreate($Uri, [UriKind]::Absolute, [ref]$Parsed)) { return $false }
    if ($Parsed.Scheme -ne 'https' -or $Parsed.Host -ne 'graph.microsoft.com' -or -not $Parsed.IsDefaultPort -or $Parsed.UserInfo -or $Parsed.Fragment) { return $false }
    if ($Parsed.AbsolutePath -match '%|\.\.') { return $false }
    $Contract = Get-IntuneGraphContract $Parsed.AbsolutePath
    if ($null -eq $Contract -or -not (Get-IntuneValue $Contract 'enabled' $true)) { return $false }
    if (-not ('System.Web.HttpUtility' -as [type])) { Add-Type -AssemblyName System.Web }
    $QueryParameters = [System.Web.HttpUtility]::ParseQueryString($Parsed.Query)
    foreach ($Key in $QueryParameters.AllKeys) {
        if ($null -eq $Key -or $Key -cnotin @('$select', '$filter', '$orderby', '$top', '$skip', '$skiptoken', '$count') -or $QueryParameters.GetValues($Key).Count -ne 1) { return $false }
    }
    if ($AllowExpansion) {
        if ($QueryParameters['$select'] -match '(?i)(credentials|password|key|token)') { return $false }
        if ($Parsed.AbsolutePath -ceq '/v1.0/deviceAppManagement/targetedManagedAppConfigurations') { return $QueryParameters['$select'] -ceq 'id,displayName,createdDateTime,lastModifiedDateTime,version,isAssigned,deployedAppCount' }
        if ($Parsed.AbsolutePath -ceq '/v1.0/deviceAppManagement/mobileAppConfigurations') { return $QueryParameters['$select'] -ceq 'id,displayName,createdDateTime,lastModifiedDateTime,version,targetedMobileApps' }
        if (Test-IntuneServicePath $Parsed.AbsolutePath) { return $true }
        if ($Parsed.AbsolutePath -eq '/v1.0/deviceAppManagement/vppTokens') {
            return $QueryParameters['$select'] -ceq 'id,expirationDateTime,state,lastSyncDateTime,lastSyncStatus,automaticallyUpdateApps'
        }
        if ($Parsed.AbsolutePath -match '^/beta/deviceManagement/(configurationPolicies|intents|templates|roleScopeTags|assignmentFilters|deviceEnrollmentConfigurations|windowsAutopilotDeploymentProfiles)$') { return $true }
        if ($Parsed.AbsolutePath -match '^/beta/deviceManagement/(configurationPolicies|intents|deviceEnrollmentConfigurations|windowsAutopilotDeploymentProfiles)/[^/]+/assignments$') { return $true }
        if ($Parsed.AbsolutePath -match '^/beta/deviceManagement/configurationPolicies/[^/]+/settings(/[^/]+/settingDefinitions)?$') { return $true }
        if ($Parsed.AbsolutePath -match '^/v1\.0/(policies/deviceRegistrationPolicy|identity/conditionalAccess/policies|directory/deviceLocalCredentials|informationProtection/bitlocker/recoveryKeys|deviceManagement/applePushNotificationCertificate|deviceManagement/deviceCompliancePolicies/[^/]+/scheduledActionsForRule(/[^/]+/scheduledActionConfigurations)?)$') { return $true }
    }
    return $Parsed.AbsolutePath -match '^/v1\.0/(deviceManagement/(managedDevices|deviceCompliancePolicies|deviceConfigurations|roleDefinitions|auditEvents)|deviceManagement/deviceCompliancePolicies/[^/]+/(assignments|deviceStatuses)|deviceManagement/deviceConfigurations/[^/]+/assignments|deviceManagement/roleDefinitions/[^/]+/roleAssignments|deviceAppManagement/mobileApps(/[^/]+/assignments)?)$'
}

function Get-IntuneFieldList {
    param([string]$Module)
    $Expanded = @(Get-IntuneExpansionFields $Module)
    if ($Expanded.Count) { return $Expanded }
    switch ($Module) {
        'ManagedDevices' { return @('id', 'deviceName', 'operatingSystem', 'osVersion', 'managementAgent', 'managedDeviceOwnerType', 'deviceEnrollmentType', 'azureADDeviceId', 'enrolledDateTime', 'lastSyncDateTime', 'complianceState', 'isEncrypted', 'model', 'manufacturer') }
        'CompliancePolicies' { return @('id', '@odata.type', 'displayName', 'createdDateTime', 'lastModifiedDateTime', 'version') }
        'DeviceConfigurations' { return @('id', '@odata.type', 'displayName', 'createdDateTime', 'lastModifiedDateTime', 'version') }
        'ComplianceStates' { return @('id', 'deviceDisplayName', 'status', 'lastReportedDateTime', 'complianceGracePeriodExpirationDateTime') }
        'Applications' { return @('id', '@odata.type', 'displayName', 'publisher', 'createdDateTime', 'lastModifiedDateTime', 'publishingState') }
        'RoleDefinitions' { return @('id', 'displayName', 'isBuiltIn', 'rolePermissions') }
        'RoleAssignments' { return @('id', 'displayName', 'resourceScopes') }
        'AuditEvents' { return @('id', 'displayName', 'activity', 'activityDateTime', 'activityType', 'activityResult', 'category', 'correlationId') }
        'ComplianceAssignments' { return @('id', 'target') }
        'ConfigurationAssignments' { return @('id', 'target') }
        default { return @('id', 'intent', 'target') }
    }
}

function ConvertTo-IntuneSafeRow {
    param($Row, [string]$Module, [string]$ParentId)
    if ($Module -in @('RawSettings', 'RawDefinitions')) {
        if (-not (Get-IntuneValue $Row 'id')) { throw 'Missing setting identity' }
        return $Row
    }
    $Safe = [ordered]@{}
    foreach ($Field in (Get-IntuneFieldList $Module)) {
        $Value = Get-IntuneValue $Row $Field
        if ($null -eq $Value) { continue }
        if ($Field -eq 'modules') {
            $Safe[$Field] = ConvertTo-IntuneEndpointModules $Value
        } elseif ($Field -eq 'mobileAppIdentifier') {
            $Identifier = [ordered]@{}
            foreach ($Key in @('@odata.type', 'packageId', 'bundleId')) {
                $Item = Get-IntuneValue $Value $Key
                if ($Item -is [string] -and $Item.Length -le 256) { $Identifier[$Key] = $Item }
            }
            $Safe[$Field] = $Identifier
        } elseif ($Field -in @('allowedDataStorageLocations', 'targetedMobileApps')) {
            $Safe[$Field] = @(foreach ($Item in $Value) { if ($Item -is [string] -and $Item.Length -le 128) { $Item } })
        } elseif ($Field -eq 'target') {
            $Target = [ordered]@{}
            $TargetFields = @('@odata.type', 'groupId')
            if ($Module -in @('ModernAssignments', 'LegacyAssignments', 'EnrollmentAssignments', 'AutopilotAssignments')) {
                $TargetFields += @('deviceAndAppManagementAssignmentFilterId', 'deviceAndAppManagementAssignmentFilterType')
            }
            foreach ($Key in $TargetFields) {
                $Item = Get-IntuneValue $Value $Key
                if ($null -ne $Item -and $Item -is [string]) { $Target[$Key] = $Item }
            }
            $Safe[$Field] = $Target
        } elseif ($Field -eq 'rolePermissions') {
            $Safe[$Field] = @(foreach ($Permission in $Value) {
                $Actions = Get-IntuneValue $Permission 'resourceActions' @()
                @{ resourceActions = @(foreach ($Action in $Actions) {
                    $Allowed = Get-IntuneValue $Action 'allowedResourceActions' @()
                    $Denied = Get-IntuneValue $Action 'notAllowedResourceActions' @()
                    @{ allowedResourceActions = @($Allowed | Where-Object { $_ -is [string] }); notAllowedResourceActions = @($Denied | Where-Object { $_ -is [string] }) }
                }) }
            })
        } elseif ($Field -in @('resourceScopes', 'roleScopeTagIds', 'scopeMembers', 'selectedMobileAppIds')) {
            $Safe[$Field] = @($Value | Where-Object { $_ -is [string] })
        } elseif ($Field -eq 'rules') {
            $Safe[$Field] = @(foreach ($Rule in $Value) {
                $ProjectedRule = [ordered]@{}
                foreach ($Key in @('@odata.type', 'ruleType', 'operationType', 'operator')) {
                    $Item = Get-IntuneValue $Rule $Key
                    if ($Item -is [string] -and $Item.Length -le 256) { $ProjectedRule[$Key] = $Item }
                }
                $ProjectedRule
            })
        } elseif ($Field -eq 'outOfBoxExperienceSetting') {
            $Safe[$Field] = [ordered]@{}
            foreach ($Key in @('userType', 'deviceUsageType', 'privacySettingsHidden', 'eulaHidden', 'keyboardSelectionPageSkipped', 'escapeLinkHidden')) {
                $Item = Get-IntuneValue $Value $Key
                if ($Item -is [string] -or $Item -is [bool]) { $Safe[$Field][$Key] = $Item }
            }
        } elseif ($Field -in @('templateReference', 'installExperience', 'returnCodes', 'localAdminPassword', 'azureADJoin', 'conditions', 'grantControls')) {
            $Safe[$Field] = ConvertTo-IntuneExpansionValue $Value
        } elseif ($Value -is [string] -or $Value -is [bool] -or $Value -is [int] -or $Value -is [long]) {
            if ($Value -is [string] -and $Value.Length -gt 16384) { throw 'Oversized field' }
            $Safe[$Field] = $Value
        } elseif ($Value -is [datetime] -or $Value -is [datetimeoffset]) {
            $Safe[$Field] = $Value.ToUniversalTime().ToString('o')
        }
    }
    if ($ParentId) { $Safe['parentId'] = $ParentId }
    if (-not (Get-IntuneValue $Safe 'id')) { throw 'Missing identity' }
    return $Safe
}

function Get-IntuneHttpPage {
    param([string]$Uri)
    if (-not (Test-IntuneUri $Uri -AllowExpansion)) { return @{ StatusCode = 0; ErrorCode = 'BlockedEndpoint'; Body = $null } }
    try {
        $Response = $script:IntuneHttpClient.GetAsync($Uri).GetAwaiter().GetResult()
        try {
            $Status = [int]$Response.StatusCode
            $RetryAfter = 0
            if ($Response.Headers.RetryAfter) {
                if ($Response.Headers.RetryAfter.Delta) { $RetryAfter = [math]::Ceiling($Response.Headers.RetryAfter.Delta.TotalSeconds) }
                elseif ($Response.Headers.RetryAfter.Date) { $RetryAfter = [math]::Max(0, [math]::Ceiling(($Response.Headers.RetryAfter.Date - [datetimeoffset]::UtcNow).TotalSeconds)) }
            }
            if ($Status -ne 200) { return @{ StatusCode = $Status; RetryAfter = $RetryAfter; Body = $null } }
            $BodyText = $Response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            if ($BodyText.Length -gt 16MB) { return @{ StatusCode = 0; ErrorCode = 'ResponseTooLarge'; Body = $null } }
            return @{ StatusCode = $Status; RetryAfter = $RetryAfter; Body = ($BodyText | ConvertFrom-Json) }
        } finally { $Response.Dispose() }
    } catch {
        return @{ StatusCode = 0; ErrorCode = 'TransportOrInvalidResponse'; Body = $null }
    }
}

function Get-IntuneCollection {
    param(
        [string]$Uri, [string]$Module, [string]$ParentId,
        [switch]$AllowExpansion, [switch]$Singleton,
        [scriptblock]$Request = { param($Address) Get-IntuneHttpPage $Address },
        [scriptblock]$Delay = { param($Seconds) [System.Threading.Tasks.Task]::Delay([timespan]::FromSeconds($Seconds)).GetAwaiter().GetResult() },
        [scriptblock]$Now = { [datetime]::UtcNow },
        [datetime]$Deadline = ([datetime]::UtcNow.AddMinutes(30)),
        [int]$MaxRows = 100000
    )
    $Rows = [System.Collections.Generic.List[object]]::new()
    $Visited = [System.Collections.Generic.HashSet[string]]::new()
    $Ids = [System.Collections.Generic.HashSet[string]]::new()
    $Pages = 0
    $Failure = $null
    $Contract = Get-IntuneGraphContract ([uri]$Uri).AbsolutePath $Module
    if ($null -eq $Contract -or -not (Get-IntuneValue $Contract 'enabled' $true)) {
        return @{ Rows = @(); Status = @{ State = 'Unsupported'; ApiVersion = ([uri]$Uri).AbsolutePath.Split('/')[1]; Endpoint = ([uri]$Uri).AbsolutePath; PagesRead = 0; RowsRead = 0; ErrorCode = 'UnverifiedContract'; Details = 'No enabled module/route contract; no request sent.'; CompletedParentIds = @(); ScopeComplete = $false } }
    }
    $Next = $Uri
    while ($Next) {
        if (-not (Test-IntuneUri $Next -AllowExpansion:$AllowExpansion) -or ([uri]$Next).AbsolutePath -cne ([uri]$Uri).AbsolutePath) { $Failure = 'BlockedEndpoint'; break }
        if (-not $Visited.Add($Next)) { $Failure = 'PaginationCycle'; break }
        if ($Pages -ge 10000 -or (& $Now) -ge $Deadline) { $Failure = 'CollectionBudget'; break }
        $Response = $null
        for ($Attempt = 0; $Attempt -lt 4; $Attempt++) {
            try { $Response = & $Request $Next } catch { $Response = @{ StatusCode = 0; ErrorCode = 'TransportError' } }
            $Status = [int](Get-IntuneValue $Response 'StatusCode' 0)
            if ($Status -notin @(429, 503, 504)) { break }
            $Seconds = [double](Get-IntuneValue $Response 'RetryAfter' 0)
            if ($Seconds -le 0) { $Seconds = [math]::Pow(2, $Attempt + 1) }
            if ($Attempt -eq 3 -or (& $Now).AddSeconds($Seconds) -ge $Deadline) { break }
            & $Delay $Seconds
        }
        if ($Status -ne 200) {
            $Failure = switch ($Status) { 401 { 'Unauthorized' }; 403 { 'Forbidden' }; 429 { 'Throttled' }; 0 { Get-IntuneValue $Response 'ErrorCode' 'TransportError' }; default { "HTTP$Status" } }
            break
        }
        $Body = Get-IntuneValue $Response 'Body'
        $Values = Get-IntuneValue $Body 'value'
        if ($Singleton -and $null -ne $Body -and $null -eq (Get-IntuneValue $Body 'error')) {
            if ($null -eq $Values) { $Values = @($Body) }
            elseif ($Values -isnot [array]) { $Values = @($Values) }
        }
        if ($null -eq $Body -or $null -ne (Get-IntuneValue $Body 'error') -or $Values -isnot [array]) { $Failure = 'InvalidResponse'; break }
        $Pages++
        foreach ($Row in $Values) {
            if ($Rows.Count -ge $MaxRows) { $Failure = 'RowLimit'; break }
            try { $Safe = ConvertTo-IntuneSafeRow $Row $Module $ParentId } catch { $Failure = 'InvalidRow'; break }
            $RowIdentity = [string](Get-IntuneValue $Safe 'id')
            if (-not $RowIdentity) { $Failure = 'InvalidIdentity'; break }
            if (-not $Ids.Add($RowIdentity)) { $Failure = 'DuplicateIdentity'; break }
            $Rows.Add($Safe)
        }
        if ($Failure) { break }
        $NextValue = Get-IntuneValue $Body '@odata.nextLink'
        if ($null -ne $NextValue -and ($NextValue -isnot [string] -or -not $NextValue)) { $Failure = 'InvalidNextLink'; break }
        $Next = $NextValue
    }
    return @{ Rows = @($Rows.ToArray()); Status = [ordered]@{
        State = $(if (-not $Failure) { 'Complete' } elseif ($Pages -gt 0) { 'Partial' } else { 'Error' })
        ApiVersion = ([uri]$Uri).AbsolutePath.Split('/')[1]; Endpoint = ([uri]$Uri).AbsolutePath; PagesRead = $Pages; RowsRead = $Rows.Count
        ScopeComplete = $false; ErrorCode = $Failure
        Details = $(if ($Failure) { 'Collection unavailable or incomplete; no absence inference.' } else { 'All requested pages returned; tenant-wide visibility is not established.' })
        CompletedParentIds = @()
    } }
}

function Invoke-IntuneDiscoveryCore {
    param([string]$SelectedTenant, [bool]$Rbac, [bool]$Audit, [datetime]$AuditSince,
    [bool]$Configuration, [bool]$Entra, [bool]$Recovery, [bool]$Enrollment, [bool]$Apple,
    [bool]$Mam, [bool]$RemoteHelp, [bool]$Connectors, [bool]$AppConfiguration, [bool]$PlatformCompliance,
    [bool]$Tunnel,
    [bool]$MamLaunch,
    [string[]]$EndpointPaths = @(),
    [string]$DefenderPath, [string[]]$AppControlPaths = @(),
        [scriptblock]$Request = { param($Address) Get-IntuneHttpPage $Address },
        [scriptblock]$Delay = { param($Seconds) [System.Threading.Tasks.Task]::Delay([timespan]::FromSeconds($Seconds)).GetAwaiter().GetResult() },
        [hashtable]$Requirements = @{})
    $Started = [datetime]::UtcNow
    $Deadline = $Started.AddMinutes(30)
    $Inventory = [ordered]@{}
    $States = [ordered]@{}
    $Modules = [ordered]@{
        ManagedDevices = '/deviceManagement/managedDevices'
        CompliancePolicies = '/deviceManagement/deviceCompliancePolicies'
        DeviceConfigurations = '/deviceManagement/deviceConfigurations'
        Applications = '/deviceAppManagement/mobileApps'
    }
    if ($Rbac) { $Modules['RoleDefinitions'] = '/deviceManagement/roleDefinitions' }
    if ($Audit) { $Modules['AuditEvents'] = '/deviceManagement/auditEvents?$filter=' + [uri]::EscapeDataString(('activityDateTime ge {0}' -f $AuditSince.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'))) }
    foreach ($Module in $Modules.Keys) {
        $Result = Get-IntuneCollection -Uri ('https://graph.microsoft.com/v1.0' + $Modules[$Module]) -Module $Module -Request $Request -Delay $Delay -Deadline $Deadline
        $Inventory[$Module] = @($Result.Rows)
        $States[$Module] = $Result.Status
    }
    $Children = @(
        @{ Module = 'ComplianceAssignments'; Parent = 'CompliancePolicies'; Path = 'deviceManagement/deviceCompliancePolicies'; Child = 'assignments' },
        @{ Module = 'ComplianceStates'; Parent = 'CompliancePolicies'; Path = 'deviceManagement/deviceCompliancePolicies'; Child = 'deviceStatuses' },
        @{ Module = 'ConfigurationAssignments'; Parent = 'DeviceConfigurations'; Path = 'deviceManagement/deviceConfigurations'; Child = 'assignments' },
        @{ Module = 'ApplicationAssignments'; Parent = 'Applications'; Path = 'deviceAppManagement/mobileApps'; Child = 'assignments' }
    )
    if ($Rbac) { $Children += @{ Module = 'RoleAssignments'; Parent = 'RoleDefinitions'; Path = 'deviceManagement/roleDefinitions'; Child = 'roleAssignments' } }
    foreach ($Child in $Children) {
        $Rows = [System.Collections.Generic.List[object]]::new()
        $CompletedParents = [System.Collections.Generic.List[string]]::new()
        $Pages = 0
        $Failure = $null
        if ($States[$Child.Parent].State -ne 'Complete') { $Failure = 'ParentCollectionIncomplete' }
        foreach ($Parent in $Inventory[$Child.Parent]) {
            $ParentId = [string]$Parent['id']
            $Uri = 'https://graph.microsoft.com/v1.0/{0}/{1}/{2}' -f $Child.Path, [uri]::EscapeDataString($ParentId), $Child.Child
            $Result = Get-IntuneCollection -Uri $Uri -Module $Child.Module -ParentId $ParentId -Request $Request -Delay $Delay -Deadline $Deadline -MaxRows ([math]::Max(0, 100000 - $Rows.Count))
            foreach ($Row in $Result.Rows) { $Rows.Add($Row) }
            $Pages += $Result.Status.PagesRead
            if ($Result.Status.State -eq 'Complete') { $CompletedParents.Add($ParentId) } else { $Failure = 'ChildCollectionIncomplete' }
        }
        $Inventory[$Child.Module] = @($Rows.ToArray())
        $States[$Child.Module] = [ordered]@{
            State = $(if ($Failure) { 'Partial' } else { 'Complete' }); ApiVersion = 'v1.0'; Endpoint = ('/' + $Child.Path + '/{id}/' + $Child.Child)
            PagesRead = [math]::Max(1, $Pages); RowsRead = $Rows.Count; ScopeComplete = $false; ErrorCode = $Failure
            Details = "Completed child requests for $($CompletedParents.Count) of $($Inventory[$Child.Parent].Count) visible parent objects."
            CompletedParentIds = @($CompletedParents.ToArray())
        }
    }
    Invoke-IntuneExpansion -Inventory $Inventory -States $States -Request $Request -Delay $Delay -Deadline $Deadline -Configuration $Configuration -Rbac $Rbac -Entra $Entra -Recovery $Recovery -Enrollment $Enrollment -Apple $Apple
    Invoke-IntuneServices -Inventory $Inventory -States $States -Request $Request -Delay $Delay -Deadline $Deadline -Mam $Mam -RemoteHelp $RemoteHelp -Connectors $Connectors -AppConfiguration $AppConfiguration -PlatformCompliance $PlatformCompliance -MamLaunch $MamLaunch
    if ($Tunnel) { Invoke-IntuneTunnel -Inventory $Inventory -States $States -Request $Request -Delay $Delay -Deadline $Deadline }
    else {
        foreach ($Module in @('TunnelSites', 'TunnelServers')) {
            $Inventory[$Module] = @()
            $States[$Module] = @{ State = 'NotRequested'; ApiVersion = ''; RowsRead = 0; PagesRead = 0; CompletedParentIds = @() }
        }
    }
    if ($DefenderPath) {
        if ((Get-Item -LiteralPath $DefenderPath).Length -gt 64MB) { throw 'Defender evidence size limit.' }
        $DefenderDocument = [IO.File]::ReadAllText((Resolve-Path -LiteralPath $DefenderPath).Path) | ConvertFrom-Json
        if ((Get-IntuneValue $DefenderDocument 'TenantId') -ne $SelectedTenant -or (Get-IntuneValue $DefenderDocument 'SchemaVersion') -ne '1.0') { throw 'Defender evidence tenant/schema mismatch.' }
        $DefenderResult = Get-IntuneValue $DefenderDocument 'Result'
        $Inventory['DefenderMachines'] = @(foreach ($Machine in (Get-IntuneValue $DefenderResult 'Rows' @())) {
            $Safe = ConvertTo-IntuneSafeRow $Machine DefenderMachines
            $Safe['collectedAtUtc'] = Get-IntuneValue $DefenderDocument 'CollectedAtUtc'
            $Safe
        })
        $States['DefenderMachines'] = @{ State = (Get-IntuneValue $DefenderResult 'State'); ApiVersion = 'mde-v1'; Endpoint = 'https://api.security.microsoft.com/api/machines'; PagesRead = (Get-IntuneValue $DefenderResult 'PagesRead' 0); RowsRead = $Inventory.DefenderMachines.Count; ScopeComplete = $false; ErrorCode = (Get-IntuneValue $DefenderResult 'ErrorCode'); Details = 'Separate authorized Defender sample; device-group visibility and retention apply.'; CompletedParentIds = @() }
    }
    if ($AppControlPaths.Count) {
        $Inventory['AppControlPolicies'] = @(foreach ($PolicyPath in $AppControlPaths) {
            if ((Get-Item -LiteralPath $PolicyPath).Length -gt 1MB) { throw 'App Control policy exceeds 1 MB.' }
            ConvertTo-IntuneAppControlMetadata ([IO.File]::ReadAllText((Resolve-Path -LiteralPath $PolicyPath).Path))
        })
        $States['AppControlPolicies'] = @{ State = 'Complete'; ApiVersion = 'xml-1.0'; Endpoint = 'Explicit App Control XML files'; PagesRead = 1; RowsRead = $Inventory.AppControlPolicies.Count; ScopeComplete = $false; ErrorCode = $null; Details = 'Explicit source-policy metadata only; not tenant or device coverage.'; CompletedParentIds = @() }
    }
    if ($EndpointPaths.Count) {
        . (Join-Path $PSScriptRoot 'IntuneEndpointTimestamps.ps1')
        $EndpointRows = @(foreach ($EvidencePath in $EndpointPaths) {
            if ((Get-Item -LiteralPath $EvidencePath).Length -gt 1MB) { throw 'Endpoint evidence exceeds 1 MB.' }
            $Local = ConvertFrom-IntuneEndpointJson ([IO.File]::ReadAllText((Resolve-Path -LiteralPath $EvidencePath).Path))
            if ((Get-IntuneValue $Local 'TenantId') -ne $SelectedTenant -or (Get-IntuneValue $Local 'SchemaVersion') -ne '1.0') { throw 'Endpoint evidence tenant/schema mismatch.' }
            $LocalRow = @{ id = (Get-IntuneValue $Local 'DeviceId'); collectedAtUtc = (Get-IntuneValue $Local 'CollectedAtUtc'); modules = (Get-IntuneValue $Local 'Modules') }
            ConvertTo-IntuneSafeRow -Row $LocalRow -Module EndpointEvidence
        })
        $Inventory['EndpointEvidence'] = $EndpointRows
        $States['EndpointEvidence'] = @{ State = 'Complete'; ApiVersion = 'local-1.0'; Endpoint = 'Imported endpoint evidence'; PagesRead = 1; RowsRead = $EndpointRows.Count; ScopeComplete = $false; ErrorCode = $null; Details = 'Explicit local samples only; not fleet coverage.'; CompletedParentIds = @() }
    }
    foreach ($Module in @('ModernConfigurationPolicies', 'EnrollmentPolicies', 'EndpointSecurityPolicies', 'UpdatePolicies', 'AppProtectionPolicies', 'AppConfigurationPolicies', 'AssignmentFilters', 'Connectors', 'TenantSettings', 'RoleDefinitions', 'RoleAssignments', 'AuditEvents', 'SecuritySettings', 'LegacyIntents', 'LegacyTemplates', 'ScopeTags', 'ModernAssignments', 'LegacyAssignments', 'AutopilotProfiles', 'AutopilotAssignments', 'EnrollmentAssignments', 'NoncomplianceRules', 'NoncomplianceActions', 'DeviceRegistrationPolicy', 'ConditionalAccessPolicies', 'LapsMetadata', 'BitLockerMetadata', 'ApplePushCertificate', 'VppTokens')) {
        if ($States.Contains($Module)) { continue }
        $Inventory[$Module] = @()
        $States[$Module] = @{ State = 'NotRequested'; ApiVersion = ''; Endpoint = ''; PagesRead = 0; RowsRead = 0; ScopeComplete = $false; ErrorCode = $null; Details = 'No selected, verified collector adapter. Manual review required.'; CompletedParentIds = @() }
    }
    $Observed = [datetime]::UtcNow.ToString('o')
    $Observations = [System.Collections.Generic.List[object]]::new()
    foreach ($Module in @('ManagedDevices', 'ComplianceStates')) {
        foreach ($Row in $Inventory[$Module]) {
            $ParentId = Get-IntuneValue $Row 'parentId'
            $Observations.Add([ordered]@{
                ObservationId = ($Module + ':' + $ParentId + ':' + $Row['id']); Module = $Module
                EntityType = $(if ($Module -eq 'ManagedDevices') { 'managedDevice' } else { 'deviceComplianceDeviceStatus' })
                EntityId = $Row['id']; ParentId = $ParentId; Platform = (Get-IntuneValue $Row 'operatingSystem')
                CollectedAtUtc = $Observed; ReportedAtUtc = $(if ($Module -eq 'ManagedDevices') { Get-IntuneValue $Row 'lastSyncDateTime' } else { Get-IntuneValue $Row 'lastReportedDateTime' })
                ApiVersion = 'v1.0'; EvidenceRefs = @('/Inventory/' + $Module); Data = $Row
            })
        }
    }
    $Completed = [datetime]::UtcNow.ToString('o')
    if ($Requirements.Count) { $Requirements['AssessmentAsOfUtc'] = $Completed; $Requirements['ConfirmedAtUtc'] = $Completed }
    return [ordered]@{
        SchemaVersion = '1.0'; PackId = 'intune'; CollectionId = [guid]::NewGuid().ToString()
        Collector = @{ Name = 'Invoke-IntuneDiscovery'; Version = '0.5.10' }; Tenant = @{ Id = $SelectedTenant; Cloud = 'Global' }
        StartedAtUtc = $Started.ToString('o'); CompletedAtUtc = $Completed
        Scope = @{ RequestedModules = @($States.Keys | Where-Object { $States[$_].State -ne 'NotRequested' }); RequestedPlatforms = @('All'); Visibility = 'Unknown'; ScopeEvidenceRefs = @() }
        CollectionStatus = $States; Inventory = $Inventory; Observations = @($Observations.ToArray()); AssessmentRequirements = $Requirements
    }
}

if ($LibraryOnly) { return }
$ParsedTenant = [guid]::Empty
if (-not [guid]::TryParse($TenantId, [ref]$ParsedTenant) -or $ParsedTenant -eq [guid]::Empty) { throw 'Supply an explicit nonempty -TenantId GUID.' }
if (Test-Path -LiteralPath $OutputPath) { throw 'Output already exists; choose a new output path.' }
$Scopes = @('DeviceManagementManagedDevices.Read.All', 'DeviceManagementConfiguration.Read.All')
if ($IncludeRbac) { $Scopes += 'DeviceManagementRBAC.Read.All' }
if ($IncludeAudit) { $Scopes += 'DeviceManagementApps.Read.All' }
if (($IncludeMam -or $IncludeMamLaunch -or $IncludeAppConfiguration) -and 'DeviceManagementApps.Read.All' -notin $Scopes) { $Scopes += 'DeviceManagementApps.Read.All' }
if ($IncludeEnrollment -or $IncludeApple -or $IncludeConnectors) { $Scopes += 'DeviceManagementServiceConfig.Read.All' }
if ($IncludeEntra) { $Scopes += @('Policy.Read.DeviceConfiguration', 'Policy.Read.All') }
if ($IncludeRecoveryMetadata) { $Scopes += @('DeviceLocalCredential.ReadBasic.All', 'BitlockerKey.ReadBasic.All') }
$OriginalContext = $null
$ChangedContext = $false
$script:IntuneHttpClient = $null
$Handler = $null
$PlainToken = $null
try {
    if ($null -eq $GraphAccessToken) {
        if (-not (Get-Module -ListAvailable Az.Accounts)) { throw 'Install Az.Accounts or supply a delegated GraphAccessToken as SecureString from an approved authentication flow. Nothing was installed automatically.' }
        Import-Module Az.Accounts
        $OriginalContext = Get-AzContext
        if (-not $UseExistingConnection) {
            Connect-AzAccount -Tenant $TenantId -Scope Process -SkipContextPopulation | Out-Null
            $ChangedContext = $true
        }
        $Context = Get-AzContext
        if ($null -eq $Context -or $Context.Tenant.Id -ne $TenantId -or $Context.Environment.Name -ne 'AzureCloud') { throw 'A connection to the selected commercial-cloud tenant is required.' }
        $TokenResult = Get-AzAccessToken -ResourceTypeName MSGraph -TenantId $TenantId
        if ($TokenResult.Token -is [securestring]) { $GraphAccessToken = $TokenResult.Token }
        else { $GraphAccessToken = ConvertTo-SecureString $TokenResult.Token -AsPlainText -Force }
    }
    $PlainToken = [System.Net.NetworkCredential]::new('', $GraphAccessToken).Password
    $TokenParts = $PlainToken.Split('.')
    if ($TokenParts.Count -ne 3) { throw 'Graph token metadata cannot be verified locally.' }
    $Payload = $TokenParts[1].Replace('-', '+').Replace('_', '/')
    $Payload = $Payload.PadRight($Payload.Length + ((4 - $Payload.Length % 4) % 4), '=')
    try { $Claims = [System.Text.Encoding]::UTF8.GetString([convert]::FromBase64String($Payload)) | ConvertFrom-Json }
    catch { throw 'Graph token metadata cannot be verified locally.' }
    if ((Get-IntuneValue $Claims 'tid') -ne $TenantId -or (Get-IntuneValue $Claims 'aud') -notin @('https://graph.microsoft.com', 'https://graph.microsoft.com/', '00000003-0000-0000-c000-000000000000')) { throw 'Graph token tenant or audience does not match the requested scope.' }
    $GrantedScopes = @(([string](Get-IntuneValue $Claims 'scp' '')).Split(' '))
    $Missing = @($Scopes | Where-Object { $_ -notin $GrantedScopes })
    if ($Missing.Count) { throw ('Required delegated read scopes are unavailable: ' + ($Missing -join ', ') + '. Use an administrator-approved reader application/token. The collector never grants consent.') }
    $Expires = [long](Get-IntuneValue $Claims 'exp' 0)
    if ($Expires -le [datetimeoffset]::UtcNow.ToUnixTimeSeconds()) { throw 'Graph token is expired.' }
    Add-Type -AssemblyName System.Net.Http
    $Handler = [System.Net.Http.HttpClientHandler]::new()
    $Handler.AllowAutoRedirect = $false
    $script:IntuneHttpClient = [System.Net.Http.HttpClient]::new($Handler)
    $script:IntuneHttpClient.Timeout = [timespan]::FromSeconds(120)
    $script:IntuneHttpClient.MaxResponseContentBufferSize = 16MB
    $script:IntuneHttpClient.DefaultRequestHeaders.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer', $PlainToken)
    $script:IntuneHttpClient.DefaultRequestHeaders.UserAgent.ParseAdd('IntuneAssessor/0.5.10')
    $PlainToken = $null
    $Requirements = @{}
    if ($Assessor -and $ScopeDescription) {
        $Requirements = @{ Assessor = $Assessor; ScopeDescription = $ScopeDescription; ScopeConfirmed = $true }
        foreach ($Name in @('MaxCollectionAgeHours', 'MaxPolicyReportAgeHours', 'MaxDeviceSyncAgeDays')) {
            if ($PSBoundParameters.ContainsKey($Name)) { $Requirements[$Name] = $PSBoundParameters[$Name] }
        }
    }
    $Export = Invoke-IntuneDiscoveryCore -SelectedTenant $TenantId -Rbac $IncludeRbac.IsPresent -Audit $IncludeAudit.IsPresent -AuditSince $AuditSinceUtc -Requirements $Requirements -Configuration $IncludeConfiguration.IsPresent -Entra $IncludeEntra.IsPresent -Recovery $IncludeRecoveryMetadata.IsPresent -Enrollment $IncludeEnrollment.IsPresent -Apple $IncludeApple.IsPresent -Mam $IncludeMam.IsPresent -MamLaunch $IncludeMamLaunch.IsPresent -RemoteHelp $IncludeRemoteHelp.IsPresent -Connectors $IncludeConnectors.IsPresent -AppConfiguration $IncludeAppConfiguration.IsPresent -PlatformCompliance $IncludePlatformCompliance.IsPresent -Tunnel $IncludeTunnel.IsPresent -EndpointPaths $EndpointEvidencePaths -DefenderPath $DefenderEvidencePath -AppControlPaths $AppControlPolicyPaths
    $Json = $Export | ConvertTo-Json -Depth 30
    $Bytes = [System.Text.UTF8Encoding]::new($false).GetBytes($Json)
    if ($Bytes.Length -gt 64MB) { throw 'Export exceeds the Assay 64 MB import limit. No output was written.' }
    $ResolvedOutput = [System.IO.Path]::GetFullPath($OutputPath)
    $Stream = [System.IO.File]::Open($ResolvedOutput, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    try { $Stream.Write($Bytes, 0, $Bytes.Length) } finally { $Stream.Dispose() }
    Write-Output ('Exported Intune observations to ' + $ResolvedOutput)
    foreach ($Module in $Export.CollectionStatus.Keys) { Write-Output ('{0}: {1}, {2} rows' -f $Module, $Export.CollectionStatus[$Module].State, $Export.CollectionStatus[$Module].RowsRead) }
} finally {
    $PlainToken = $null
    if ($null -ne $script:IntuneHttpClient) { $script:IntuneHttpClient.Dispose(); $script:IntuneHttpClient = $null }
    if ($null -ne $Handler) { $Handler.Dispose() }
    if ($ChangedContext -and $null -ne $OriginalContext) { Set-AzContext -Context $OriginalContext -Scope Process | Out-Null }
}