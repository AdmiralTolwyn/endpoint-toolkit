#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Invoke-IntuneDiscovery.ps1') -LibraryOnly
function Assert-Expansion([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
$App = ConvertTo-IntuneSafeRow -Module Applications -Row @{ id = 'app'; detectionRules = @(@{ scriptContent = 'SECRET' }); rules = @(@{ ruleType = 'detection'; operationType = 'notConfigured'; scriptContent = 'SECRET' }); roleScopeTagIds = @('unsupported') }
Assert-Expansion ($App.rules[0].ruleType -eq 'detection' -and -not (($App | ConvertTo-Json -Depth 10) -match 'SECRET|unsupported|detectionRules')) 'Documented Win32 rules not safely projected'
$Autopilot = ConvertTo-IntuneSafeRow -Module AutopilotProfiles -Row @{ id = 'profile'; outOfBoxExperienceSettings = @{ userType = 'administrator' }; outOfBoxExperienceSetting = @{ userType = 'standard'; privacySettingsHidden = $true } }
Assert-Expansion ($Autopilot.outOfBoxExperienceSetting.userType -eq 'standard' -and $Autopilot.outOfBoxExperienceSetting.privacySettingsHidden -and -not $Autopilot.Contains('outOfBoxExperienceSettings')) 'Deprecated Autopilot OOBE field used'
Assert-Expansion (-not (Test-IntuneUri 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies')) 'Beta became default'
Assert-Expansion (Test-IntuneUri 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies' -AllowExpansion) 'Opt-in beta blocked'
foreach ($Path in @('/beta/deviceManagement/roleScopeTags/id/settings', '/beta/deviceManagement/templates/id/assignments', '/beta/deviceManagement/intents/id/settings/id/settingDefinitions', '/v1.0/deviceManagement/managedDevices/id/roleAssignments', '/v1.0/deviceManagement/deviceConfigurations/id/deviceStatuses', '/v1.0/deviceManagement/auditEvents/id/assignments')) {
    Assert-Expansion (-not (Test-IntuneUri ('https://graph.microsoft.com' + $Path) -AllowExpansion)) 'Undocumented or unused route combination accepted'
}
foreach ($Uri in @('https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?%24expand=settings', 'https://graph.microsoft.com/v1.0/deviceAppManagement/vppTokens?$select=id,expirationDateTime,state,lastSyncDateTime,lastSyncStatus,automaticallyUpdateApps&%24select=token', 'https://graph.microsoft.com/v1.0/directory/deviceLocalCredentials?%24select=credentials')) {
    Assert-Expansion (-not (Test-IntuneUri $Uri -AllowExpansion)) 'Encoded/duplicate secret query accepted'
}
foreach ($Uri in @('https://graph.microsoft.com/v1.0/directory/deviceLocalCredentials/device-id', 'https://graph.microsoft.com/v1.0/informationProtection/bitlocker/recoveryKeys/key-id', 'https://graph.microsoft.com/v1.0/deviceAppManagement/vppTokens', 'https://graph.microsoft.com/v1.0/deviceAppManagement/vppTokens?$select=token', 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/policy/assign')) {
    Assert-Expansion (-not (Test-IntuneUri $Uri -AllowExpansion)) "Unsafe expansion endpoint: $Uri"
}
$Definition = @{ id = 'opaque'; baseUri = './Device/Vendor/MSFT/Policy/Config/Defender'; offsetUri = 'AllowCloudProtection'; version = '1'; options = @(@{ itemId = 'opaque_disabled_1'; optionValue = @{ value = 0 } }, @{ itemId = 'opaque_enabled_0'; optionValue = @{ value = 1 } }) }
$Facts = @(ConvertTo-IntuneSettingFacts @{ settingDefinitionId = 'opaque'; choiceSettingValue = @{ value = 'opaque_disabled_1' } } @($Definition) 'policy' 'setting')
Assert-Expansion ($Facts[0].value -eq 0 -and $Facts[0].resolution -eq 'Resolved') 'Choice suffix guessed instead of definition join'
$Facts = @(ConvertTo-IntuneSettingFacts @{ settingDefinitionId = 'password'; simpleSettingValue = @{ value = 'SECRET' } } @(@{ id = 'password'; baseUri = './Device/Vendor/MSFT'; offsetUri = 'Unreviewed/Password' }) 'policy' 'setting')
Assert-Expansion (-not (($Facts | ConvertTo-Json -Depth 20).Contains('SECRET'))) 'Unknown setting secret retained'
$Facts = @(ConvertTo-IntuneSettingFacts @{ settingDefinitionId = 'opaque'; choiceSettingValue = @{ value = 'opaque_enabled_0'; settingValueTemplateReference = @{ useTemplateDefault = $true } } } @($Definition) 'policy' 'setting')
Assert-Expansion ($Facts[0].resolution -eq 'UnresolvedTemplateDefault') 'Template default guessed'
$Request = {
    param($Uri)
    $Path = ([uri]$Uri).AbsolutePath
    $Rows = @()
    switch -Regex ($Path) {
        '/managedDevices$' { $Rows = @(@{ id = 'managed-sample'; deviceName = 'Synthetic device'; operatingSystem = 'Windows'; azureADDeviceId = '33333333-3333-4333-8333-333333333333'; managementAgent = 'mdm' }) }
        '/configurationPolicies$' { $Rows = @(@{ id = 'modern-1'; name = 'Renamed baseline'; platforms = 'windows10'; templateReference = @{ templateFamily = 'baseline'; templateId = 'template-1'; templateDisplayVersion = '25H2'; password = 'SECRET' }; roleScopeTagIds = @('tag-1'); isAssigned = $true }) }
        '/configurationPolicies/modern-1/settings$' { $Rows = @(@{ id = '0'; settingInstance = @{ settingDefinitionId = 'opaque'; choiceSettingValue = @{ value = 'opaque_enabled_0' } } }) }
        '/settingDefinitions$' { $Rows = @($Definition) }
        '/roleScopeTags$' { $Rows = @(@{ id = 'tag-1'; displayName = 'Example'; isBuiltIn = $false }) }
        '/deviceRegistrationPolicy$' { return @{ StatusCode = 200; Body = @{ id = 'deviceRegistrationPolicy'; localAdminPassword = @{ isEnabled = $true; password = 'SECRET' } } } }
        '/deviceLocalCredentials$' { $Rows = @(@{ id = '33333333-3333-4333-8333-333333333333'; lastBackupDateTime = '2026-09-17T08:00:00Z'; credentials = @(@{ passwordBase64 = 'SECRET' }) }) }
        '/recoveryKeys$' { $Rows = @(@{ id = 'key-metadata'; deviceId = '33333333-3333-4333-8333-333333333333'; key = 'SECRET'; volumeType = 'operatingSystemVolume' }) }
        '/applePushNotificationCertificate$' { return @{ StatusCode = 200; Body = @{ id = 'apns'; expirationDateTime = '2027-09-17T00:00:00Z'; certificate = 'SECRET' } } }
        '/vppTokens$' { $Rows = @(@{ id = 'vpp'; token = 'SECRET'; state = 'valid'; expirationDateTime = '2027-09-17T00:00:00Z'; lastSyncStatus = 'completed' }) }
    }
    return @{ StatusCode = 200; Body = (@{ value = $Rows } | ConvertTo-Json -Depth 30 | ConvertFrom-Json) }
}
$Document = Invoke-IntuneDiscoveryCore -SelectedTenant '22222222-2222-4222-8222-222222222222' -Configuration $true -Rbac $true -Entra $true -Recovery $true -Enrollment $true -Apple $true -Request $Request -Requirements @{ Assessor = 'Offline expansion fixture'; ScopeDescription = 'Synthetic configuration'; ScopeConfirmed = $true; MaxCollectionAgeHours = 24; ConfigurationReview = @{ ReferenceProfile = 'windows-25h2'; PolicyIds = @('modern-1'); DefenderPrimary = $true } }
Assert-Expansion ($Document.CollectionStatus.SecuritySettings.State -eq 'Complete') 'Settings collection incomplete'
Assert-Expansion ($Document.CollectionStatus.ModernAssignments.State -eq 'Unsupported' -and $Document.CollectionStatus.AutopilotAssignments.State -eq 'Unsupported') 'Quarantined assignment contracts looked complete'
$Quarantined = Get-IntuneCollection -Uri 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/policy/assignments' -Module ModernAssignments -AllowExpansion -Request { throw 'Must not request' }
Assert-Expansion ($Quarantined.Status.State -eq 'Unsupported' -and $Quarantined.Status.PagesRead -eq 0) 'Quarantined assignment request escaped guard'
Assert-Expansion ($Document.Inventory.SecuritySettings[0].value -eq 1) 'Decoded setting lost'
Assert-Expansion ($Document.CollectionStatus.SecuritySettings.CompletedParentIds[0] -eq 'modern-1') 'Per-policy settings completeness lost'
Assert-Expansion ($Document.Inventory.DeviceRegistrationPolicy[0].localAdminPassword.isEnabled -eq $true) 'Singleton lost'
$Json = $Document | ConvertTo-Json -Depth 30
Assert-Expansion (-not $Json.Contains('SECRET')) 'Export contains secret'
$Result = Get-IntuneCollection -Uri 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies' -Module ModernConfigurationPolicies -AllowExpansion -Request { @{ StatusCode = 200; Body = @{ value = @(); '@odata.nextLink' = 'https://graph.microsoft.com/beta/deviceManagement/intents' } } }
Assert-Expansion ($Result.Status.ErrorCode -eq 'BlockedEndpoint') 'Cross-resource paging accepted'
if ($env:ASSAY_INTUNE_EXPANSION_FIXTURE) { [IO.File]::WriteAllText($env:ASSAY_INTUNE_EXPANSION_FIXTURE, $Json, [Text.UTF8Encoding]::new($false)) }
Write-Output 'PASS: opt-in modules, metadata-only endpoints, definition choice mapping, unresolved defaults, child coverage and secret sanitization.'
. (Join-Path $PSScriptRoot 'Get-IntuneEndpointEvidence.ps1') -LibraryOnly
$Endpoint = New-IntuneEndpointEvidence -SelectedTenant 'test' -DeviceId 'test' -Read {
    param($Module)
    if ($Module -eq 'DefenderStatus') { return @{ AMRunningMode = 'Normal'; IsTamperProtected = $true; password = 'SECRET' } }
    throw 'Synthetic missing provider'
}
$SafeEndpoint = ConvertTo-IntuneSafeRow -Module EndpointEvidence -Row @{ id = 'test'; modules = $Endpoint.Modules; collectedAtUtc = $Endpoint.CollectedAtUtc }
Assert-Expansion ($SafeEndpoint.modules.DefenderStatus.Rows[0].IsTamperProtected -eq $true) 'Endpoint Boolean lost'
Assert-Expansion ($SafeEndpoint.modules.DeviceGuard.State -eq 'Error') 'Endpoint provider error hidden'
Assert-Expansion (-not (($SafeEndpoint | ConvertTo-Json -Depth 20).Contains('SECRET'))) 'Endpoint secret retained'
Write-Output 'PASS: injected endpoint provider states and field projection; no local device collection executed.'
$Admx = ConvertTo-IntuneAdmxMetadata '<enabled/><data id="UseTPMPIN" value="1"/><data id="unexpected" value="SECRET"/>'
Assert-Expansion ($Admx.enabled -and $Admx.data.UseTPMPIN -eq 1 -and -not ($Admx | ConvertTo-Json).Contains('SECRET')) 'ADMX metadata parsing failed'
$Rejected = $false
try { Read-IntunePolicyXml '<!DOCTYPE root [<!ENTITY secret SYSTEM "file:///C:/secret">]><root>&secret;</root>' } catch { $Rejected = $true }
Assert-Expansion $Rejected 'External XML entity accepted'
$Policy = ConvertTo-IntuneAppControlMetadata '<SiPolicy xmlns="urn:schemas-microsoft-com:sipolicy" PolicyType="Base Policy"><PolicyID>{11111111-1111-4111-8111-111111111111}</PolicyID><BasePolicyID>{11111111-1111-4111-8111-111111111111}</BasePolicyID><Rules><Rule><Option>Enabled:Audit Mode</Option></Rule><Rule><Option>Enabled:Managed Installer</Option></Rule></Rules><Unexpected>SECRET</Unexpected></SiPolicy>'
Assert-Expansion ($Policy.auditMode -and $Policy.managedInstaller -and -not ($Policy | ConvertTo-Json).Contains('SECRET')) 'App Control policy metadata failed'
Write-Output 'PASS: bounded ADMX/App Control XML metadata parsing and external-entity rejection.'
$Temporary = Join-Path ([IO.Path]::GetTempPath()) ('intune-companions-' + [guid]::NewGuid().ToString())
[IO.Directory]::CreateDirectory($Temporary) | Out-Null
try {
    $Time = [datetime]::UtcNow.ToString('o')
    $Tenant = '22222222-2222-4222-8222-222222222222'
    $DefenderPath = Join-Path $Temporary 'defender.json'
    $EndpointPath = Join-Path $Temporary 'endpoint.json'
    $PolicyPath = Join-Path $Temporary 'policy.xml'
    $DefenderDocument = @{ SchemaVersion = '1.0'; TenantId = $Tenant; CollectedAtUtc = $Time; Result = @{ State = 'Complete'; PagesRead = 1; Rows = @(@{ id = 'mde'; aadDeviceId = '33333333-3333-4333-8333-333333333333'; healthStatus = 'Active'; onboardingStatus = 'onboarded'; lastSeen = $Time; lastIpAddress = 'SECRET' }) } }
    [IO.File]::WriteAllText($DefenderPath, ($DefenderDocument | ConvertTo-Json -Depth 10))
    [IO.File]::WriteAllText($EndpointPath, (@{ SchemaVersion = '1.0'; TenantId = $Tenant; DeviceId = '33333333-3333-4333-8333-333333333333'; CollectedAtUtc = $Time; Modules = $Endpoint.Modules } | ConvertTo-Json -Depth 15))
    [IO.File]::WriteAllText($PolicyPath, '<SiPolicy xmlns="urn:schemas-microsoft-com:sipolicy" PolicyType="Base Policy"><PolicyID>{11111111-1111-4111-8111-111111111111}</PolicyID><BasePolicyID>{11111111-1111-4111-8111-111111111111}</BasePolicyID><Rules><Rule><Option>Enabled:Audit Mode</Option></Rule></Rules></SiPolicy>')
    $Imported = Invoke-IntuneDiscoveryCore -SelectedTenant $Tenant -Configuration $true -Rbac $true -Entra $true -Recovery $true -Enrollment $true -Apple $true -Request $Request -DefenderPath $DefenderPath -EndpointPaths @($EndpointPath) -AppControlPaths @($PolicyPath) -Requirements @{ Assessor = 'Offline expansion fixture'; ScopeDescription = 'Synthetic configuration and companion evidence'; ScopeConfirmed = $true; MaxCollectionAgeHours = 24; ConfigurationReview = @{ ReferenceProfile = 'windows-25h2'; PolicyIds = @('modern-1'); DeviceIds = @('managed-sample'); DefenderPrimary = $true; MaxDefenderReportAgeHours = 48 } }
    Assert-Expansion ($Imported.CollectionStatus.DefenderMachines.ApiVersion -eq 'mde-v1' -and $Imported.Inventory.DefenderMachines[0].lastSeen) 'Defender import lost version/time'
    Assert-Expansion ($Imported.CollectionStatus.AppControlPolicies.ApiVersion -eq 'xml-1.0' -and $Imported.Inventory.AppControlPolicies[0].auditMode) 'App Control import lost XML state'
    Assert-Expansion ($Imported.Inventory.EndpointEvidence.Count -eq 1) 'Endpoint companion import missing'
    $Export = $Imported | ConvertTo-Json -Depth 30
    Assert-Expansion (-not $Export.Contains('SECRET')) 'Companion import leaked extra fields'
    if ($env:ASSAY_INTUNE_EXPANSION_FIXTURE) { [IO.File]::WriteAllText($env:ASSAY_INTUNE_EXPANSION_FIXTURE, $Export, [Text.UTF8Encoding]::new($false)) }
    $DefenderDocument.TenantId = 'different-tenant'
    [IO.File]::WriteAllText($DefenderPath, ($DefenderDocument | ConvertTo-Json -Depth 10))
    $Rejected = $false
    try { Invoke-IntuneDiscoveryCore -SelectedTenant $Tenant -Request $Request -DefenderPath $DefenderPath | Out-Null } catch { $Rejected = $_.Exception.Message -match 'tenant/schema mismatch' }
    Assert-Expansion $Rejected 'Cross-tenant Defender file accepted'
} finally { [IO.Directory]::Delete($Temporary, $true) }
Write-Output 'PASS: companion imports preserve identity/time/source, exclude secrets and reject tenant mismatch.'