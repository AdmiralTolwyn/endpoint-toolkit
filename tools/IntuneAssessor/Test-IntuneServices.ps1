#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Invoke-IntuneDiscovery.ps1') -LibraryOnly
function Assert-Service([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
$Policy = ConvertTo-IntuneSafeRow -Module AppProtectionPolicies -Row ('{"id":"mam","@odata.type":"#microsoft.graph.androidManagedAppProtection","pinRequired":false,"minimumPinLength":6,"encryptAppData":true,"customSettings":{"token":"SECRET"},"description":"SECRET"}' | ConvertFrom-Json)
Assert-Service ($Policy.pinRequired -eq $false -and $Policy.minimumPinLength -eq 6 -and $Policy.encryptAppData) 'Typed MAM fields lost'
Assert-Service (-not ($Policy | ConvertTo-Json -Depth 10).Contains('SECRET')) 'Unreviewed MAM data leaked'
$Remote = ConvertTo-IntuneSafeRow -Module RemoteAssistanceSettings -Row @{ id = 'remote'; remoteAssistanceState = 'disabled'; allowSessionsToUnenrolledDevices = $false; blockChat = $true; secret = 'SECRET' }
Assert-Service ($Remote.remoteAssistanceState -eq 'disabled' -and $Remote.allowSessionsToUnenrolledDevices -eq $false) 'Remote Help false/state lost'
Assert-Service (-not ($Remote | ConvertTo-Json).Contains('SECRET')) 'Remote Help extra fields retained'
Write-Output 'PASS: typed MAM/Remote Help projections preserve false values and omit unreviewed payloads.'
$EpmDefinition = @{ id = 'device_vendor_msft_policy_elevationclientsettings_enableepm'; offsetUri = 'PrivilegeManagement/ElevationClientSettings/EnableEPM'; options = @(@{ itemId = 'opaque_option'; optionValue = @{ value = 1 } }) }
$EpmFacts = @(ConvertTo-IntuneSettingFacts @{ settingDefinitionId = $EpmDefinition.id; choiceSettingValue = @{ value = 'opaque_option' } } @($EpmDefinition) 'epm' '0')
Assert-Service ($EpmFacts[0].value -eq 1 -and $EpmFacts[0].cspUri -ceq 'PrivilegeManagement/ElevationClientSettings/EnableEPM') 'Exact EPM definition binding failed'
$EpmDefinition.offsetUri = 'Unknown/Secret'
$EpmFacts = @(ConvertTo-IntuneSettingFacts @{ settingDefinitionId = $EpmDefinition.id; simpleSettingValue = @{ value = 'SECRET' } } @($EpmDefinition) 'epm' '0')
Assert-Service (-not ($EpmFacts | ConvertTo-Json -Depth 10).Contains('SECRET')) 'EPM identity/path mismatch leaked data'
$EpmPaths = [ordered]@{ enableepm = 'EnableEPM'; defaultelevationresponse = 'DefaultElevationResponse'; senddata = 'SendData'; reportingscope = 'ReportingScope' }
foreach ($Suffix in $EpmPaths.Keys) {
	$EpmId = 'device_vendor_msft_policy_elevationclientsettings_' + $Suffix
	$EpmPath = 'PrivilegeManagement/ElevationClientSettings/' + $EpmPaths[$Suffix]
	foreach ($DecodeJson in @($false, $true)) {
		foreach ($Defect in @('none', 'string-base', 'array-base', 'number-base', 'array-offset', 'object-offset', 'array-id')) {
			$BoundDefinition = @{ id = $EpmId; offsetUri = $EpmPath }
			switch ($Defect) {
				'string-base' { $BoundDefinition.baseUri = '' }
				'array-base' { $BoundDefinition.baseUri = @('./Device/Vendor/MSFT') }
				'number-base' { $BoundDefinition.baseUri = 0 }
				'array-offset' { $BoundDefinition.offsetUri = @($EpmPath) }
				'object-offset' { $BoundDefinition.offsetUri = @{ path = 'PRIVATE_PATH_VALUE' } }
				'array-id' { $BoundDefinition.id = @($EpmId) }
			}
			if ($DecodeJson) { $BoundDefinition = $BoundDefinition | ConvertTo-Json -Depth 10 | ConvertFrom-Json }
			$BoundPath = Get-IntuneEpmPath $BoundDefinition
			$BoundFacts = @(ConvertTo-IntuneSettingFacts @{ settingDefinitionId = $EpmId; simpleSettingValue = @{ value = 1 } } @($BoundDefinition) 'epm' 'binding')
			if ($Defect -in @('none', 'string-base')) {
				Assert-Service ($BoundPath -ceq $EpmPath -and $BoundFacts[0].cspUri -ceq $EpmPath -and $BoundFacts[0].value -eq 1) 'Reviewed EPM mapping lost without a generic base URI'
			} else {
				Assert-Service ($null -eq $BoundPath -and $BoundFacts[0].resolution -ceq 'UnsupportedDefinition' -and -not $BoundFacts[0].Contains('cspUri') -and -not $BoundFacts[0].Contains('value')) 'Malformed EPM path or ID bypassed typed binding'
			}
			Assert-Service (-not (($BoundFacts | ConvertTo-Json -Depth 10) -match 'PRIVATE_PATH_VALUE')) 'Malformed EPM metadata leaked'
		}
	}
}
$script:ServicePaths = [Collections.Generic.List[string]]::new()
$Request = {
	param($Address)
	$script:ServicePaths.Add($Address)
	$Path = ([uri]$Address).AbsolutePath
	if ($Path -eq '/beta/deviceManagement/remoteAssistanceSettings') { return @{ StatusCode = 200; Body = @{ value = @{ id = 'remote'; remoteAssistanceState = 'enabled'; allowSessionsToUnenrolledDevices = $false; blockChat = $true } } } }
	$Rows = switch ($Path) {
		'/v1.0/deviceAppManagement/targetedManagedAppConfigurations' { @(@{ id = 'appconfig'; isAssigned = $false; deployedAppCount = 0; customSettings = @(@{ name = 'password'; value = 'SECRET' }) }) }
		'/v1.0/deviceAppManagement/mobileAppConfigurations' { @(@{ id = 'deviceconfig'; targetedMobileApps = @('application'); encodedSettingXml = 'SECRET' }) }
		'/beta/deviceManagement/deviceCompliancePolicies' { @(@{ id = 'owner'; '@odata.type' = '#microsoft.graph.androidDeviceOwnerCompliancePolicy'; storageRequireEncryption = $true; securityRequireIntuneAppIntegrity = $true }) }
		'/beta/deviceManagement/compliancePolicies' { @(@{ id = 'linux'; name = 'Linux compliance'; platforms = 'linux'; isAssigned = $true; settingCount = 2 }) }
		'/beta/deviceAppManagement/windowsManagedAppProtections' { @(@{ id = 'windows-mam'; '@odata.type' = '#microsoft.graph.windowsManagedAppProtection'; allowedOutboundClipboardSharingLevel = 'none' }) }
		'/v1.0/deviceManagement/mobileThreatDefenseConnectors' { @(@{ id = 'connector'; partnerState = 'enabled'; androidMobileApplicationManagementEnabled = $true; androidEnabled = $false; lastHeartbeatDateTime = [datetime]::UtcNow.ToString('o'); partnerUnresponsivenessThresholdInDays = 7 }) }
		'/v1.0/deviceAppManagement/managedAppPolicies' { @(@{ id = 'mam'; '@odata.type' = '#microsoft.graph.androidManagedAppProtection'; pinRequired = $true; minimumPinLength = 6; disableAppPinIfDevicePinIsSet = $false; isAssigned = $true; encryptAppData = $true; allowedDataStorageLocations = @('oneDriveForBusiness'); customSettings = 'SECRET' }) }
		'/v1.0/deviceAppManagement/androidManagedAppProtections/mam/assignments' { @(@{ id = 'assignment'; target = @{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = 'group' } }) }
		'/v1.0/deviceAppManagement/androidManagedAppProtections/mam/apps' { @(@{ id = 'app'; mobileAppIdentifier = @{ '@odata.type' = '#microsoft.graph.androidMobileAppIdentifier'; packageId = 'com.example.app'; token = 'SECRET' } }) }
		default { @() }
	}
	return @{ StatusCode = 200; Body = (@{ value = @($Rows) } | ConvertTo-Json -Depth 15 | ConvertFrom-Json) }
}
$Document = Invoke-IntuneDiscoveryCore -SelectedTenant '22222222-2222-4222-8222-222222222222' -Mam $true -RemoteHelp $true -Connectors $true -AppConfiguration $true -PlatformCompliance $true -Request $Request -Requirements @{ ScopeConfirmed = $true; Assessor = 'Offline service test'; ScopeDescription = 'Synthetic services'; MaxCollectionAgeHours = 24 }
Assert-Service ($Document.Inventory.PlatformCompliancePolicies[0].securityRequireIntuneAppIntegrity -and $Document.CollectionStatus.ModernCompliancePolicies.ApiVersion -eq 'beta') 'Platform compliance fields/version missing'
Assert-Service ($Document.Inventory.WindowsAppProtectionPolicies[0].allowedOutboundClipboardSharingLevel -eq 'none') 'Windows MAM conflated with mobile clipboard enum'
Assert-Service ($Document.Inventory.DeviceAppConfigurations[0].targetedMobileApps[0] -eq 'application' -and $Document.CollectionStatus.AppConfigurationAssignments.CompletedParentIds -contains 'appconfig') 'App configuration scope lost'
Assert-Service ($Document.Inventory.ThreatConnectors[0].androidMobileApplicationManagementEnabled -and $Document.Inventory.ThreatConnectors[0].androidEnabled -eq $false) 'MAM and MDM connector enablement conflated'
Assert-Service ($Document.Inventory.ManagedDevices.Count -eq 0 -and $Document.Inventory.AppProtectionPolicies.Count -eq 1) 'MAM depends on enrolled Windows devices'
Assert-Service ($Document.CollectionStatus.MamAssignments.CompletedParentIds -contains 'mam') 'MAM parent coverage missing'
Assert-Service ($Document.Inventory.MamApps[0].mobileAppIdentifier.packageId -eq 'com.example.app') 'MAM app identity lost'
Assert-Service ($Document.Inventory.RemoteAssistanceSettings[0].allowSessionsToUnenrolledDevices -eq $false) 'Remote Help singleton false lost'
Assert-Service (-not ($Document | ConvertTo-Json -Depth 30).Contains('SECRET')) 'Service export leaked unknown data'
Assert-Service (-not (Test-IntuneUri 'https://graph.microsoft.com/beta/deviceManagement/remoteAssistanceSettings')) 'Remote Help became default'
Assert-Service (-not (Test-IntuneUri 'https://graph.microsoft.com/v1.0/deviceAppManagement/managedAppRegistrations' -AllowExpansion)) 'User registration collection unexpectedly enabled'
Assert-Service (-not (Test-IntuneUri 'https://graph.microsoft.com/v1.0/deviceAppManagement/targetedManagedAppConfigurations' -AllowExpansion)) 'Unprojected app configuration accepted'
Assert-Service (-not (Test-IntuneUri 'https://graph.microsoft.com/v1.0/deviceAppManagement/mobileAppConfigurations?$select=encodedSettingXml' -AllowExpansion)) 'Raw app configuration accepted'
$Partial = Get-IntuneCollection -Uri 'https://graph.microsoft.com/v1.0/deviceAppManagement/managedAppPolicies' -Module AppProtectionPolicies -AllowExpansion -Request { @{ StatusCode = 403 } }
Assert-Service ($Partial.Status.State -eq 'Error') '403 inferred MAM absence'
$script:ServicePaths.Clear()
$Default = Invoke-IntuneDiscoveryCore -SelectedTenant '22222222-2222-4222-8222-222222222222' -Request $Request
Assert-Service ($Default.CollectionStatus.AppProtectionPolicies.State -eq 'NotRequested' -and $script:ServicePaths -notcontains 'https://graph.microsoft.com/beta/deviceManagement/remoteAssistanceSettings') 'Opt-in services requested by default'
if ($env:ASSAY_INTUNE_SERVICES_FIXTURE) { [IO.File]::WriteAllText($env:ASSAY_INTUNE_SERVICES_FIXTURE, ($Document | ConvertTo-Json -Depth 30), [Text.UTF8Encoding]::new($false)) }
Write-Output 'PASS: opt-in services, child completion, BYOD-independent MAM, singleton Remote Help and secret boundaries.'