. (Join-Path $PSScriptRoot 'IntunePolicyPayloads.ps1')
. (Join-Path $PSScriptRoot 'IntuneServices.ps1')
. (Join-Path $PSScriptRoot 'IntuneEpmRules.ps1')

function Get-IntuneExpansionFields {
    param([string]$Module)
    $ServiceFields = @(Get-IntuneServiceFields $Module)
    if ($ServiceFields.Count) { return $ServiceFields }
    $Common = @('id', '@odata.type', 'displayName', 'createdDateTime', 'lastModifiedDateTime', 'version', 'roleScopeTagIds')
    $Excluded = switch ($Module) {
        { $_ -in @('CompliancePolicies', 'DeviceConfigurations', 'Applications') } { @('roleScopeTagIds') }
        'ModernConfigurationPolicies' { @('displayName', 'version') }
        'LegacyIntents' { @('createdDateTime', 'version') }
        'LegacyTemplates' { @('createdDateTime', 'lastModifiedDateTime', 'version', 'roleScopeTagIds') }
        'AssignmentFilters' { @('version', 'roleScopeTagIds') }
        'ScopeTags' { @('createdDateTime', 'lastModifiedDateTime', 'version', 'roleScopeTagIds') }
        'AutopilotProfiles' { @('version') }
        'ConditionalAccessPolicies' { @('lastModifiedDateTime', 'version', 'roleScopeTagIds') }
        default { @() }
    }
    $Common = @($Common | Where-Object { $_ -notin $Excluded })
    if ($Module -eq 'ModernConfigurationPolicies') { $Common += 'name' }
    switch ($Module) {
        'EndpointEvidence' { return @('id', 'collectedAtUtc', 'modules') }
        'DefenderMachines' { return @('id', 'aadDeviceId', 'onboardingStatus', 'healthStatus', 'firstSeen', 'lastSeen', 'osPlatform', 'osBuild', 'version', 'collectedAtUtc') }
        'CompliancePolicies' { return $Common + @('bitLockerEnabled', 'storageRequireEncryption', 'secureBootEnabled', 'codeIntegrityEnabled', 'osMinimumVersion', 'osMaximumVersion', 'deviceThreatProtectionEnabled', 'deviceThreatProtectionRequiredSecurityLevel', 'passwordRequired', 'passwordMinimumLength', 'passcodeRequired', 'passcodeMinimumLength', 'passcodeBlockSimple', 'passwordBlockSimple', 'securityBlockJailbrokenDevices', 'systemIntegrityProtectionEnabled', 'firewallEnabled', 'firewallBlockAllIncoming', 'firewallEnableStealthMode', 'minAndroidSecurityPatchLevel', 'securityRequireVerifyApps', 'securityRequireSafetyNetAttestationBasicIntegrity', 'securityRequireSafetyNetAttestationCertifiedDevice', 'securityRequireCompanyPortalAppIntegrity') }
        'DeviceConfigurations' { return $Common + @('qualityUpdatesPaused', 'featureUpdatesPaused', 'qualityUpdatesPauseExpiryDateTime', 'featureUpdatesPauseExpiryDateTime', 'qualityUpdatesDeferralPeriodInDays', 'featureUpdatesDeferralPeriodInDays', 'deadlineForQualityUpdatesInDays', 'deadlineForFeatureUpdatesInDays', 'deadlineGracePeriodInDays', 'automaticUpdateMode') }
        'Applications' { return $Common + @('publisher', 'publishingState', 'applicableArchitectures', 'allowedArchitectures', 'minimumSupportedWindowsRelease', 'installExperience', 'rules', 'returnCodes') }
        'ModernConfigurationPolicies' { return $Common + @('platforms', 'technologies', 'isAssigned', 'templateReference', 'settingCount') }
        'LegacyIntents' { return $Common + @('templateId', 'isAssigned', 'isMigratingToConfigurationPolicy') }
        'LegacyTemplates' { return $Common + @('templateType', 'platformType', 'isDeprecated', 'publishedDateTime') }
        'ScopeTags' { return $Common + @('isBuiltIn') }
        'EnrollmentPolicies' { return $Common + @('priority', 'allowNonBlockingAppInstallation', 'allowLogCollectionOnInstallFailure', 'allowDeviceResetOnInstallFailure', 'allowDeviceUseOnInstallFailure', 'blockDeviceSetupRetryByUser', 'installProgressTimeoutInMinutes', 'selectedMobileAppIds', 'installQualityUpdates') }
        'AutopilotProfiles' { return $Common + @('outOfBoxExperienceSetting', 'deviceType') }
        'AssignmentFilters' { return $Common + @('platform', 'rule') }
        'NoncomplianceRules' { return @('id', 'ruleName') }
        'NoncomplianceActions' { return @('id', 'actionType', 'gracePeriodHours', 'notificationTemplateId') }
        'ApplePushCertificate' { return @('id', 'expirationDateTime', 'lastModifiedDateTime') }
        'VppTokens' { return @('id', 'expirationDateTime', 'state', 'lastSyncDateTime', 'lastSyncStatus', 'automaticallyUpdateApps') }
        'DeviceRegistrationPolicy' { return @('id', 'localAdminPassword', 'azureADJoin', 'userDeviceQuota', 'multiFactorAuthConfiguration') }
        'LapsMetadata' { return @('id', 'lastBackupDateTime', 'refreshDateTime') }
        'BitLockerMetadata' { return @('id', 'deviceId', 'createdDateTime', 'volumeType') }
        'ConditionalAccessPolicies' { return $Common + @('modifiedDateTime', 'state', 'conditions', 'grantControls') }
        'ModernAssignments' { return @('id', 'target', 'source', 'sourceId') }
        'LegacyAssignments' { return @('id', 'target') }
        'EnrollmentAssignments' { return @('id', 'target') }
        'AutopilotAssignments' { return @('id', 'target') }
        'RoleAssignments' { return @('id', 'displayName', 'resourceScopes') }
        default { return @() }
    }
}

function ConvertTo-IntuneExpansionValue {
    param($Value, [int]$Depth = 0)
    if ($Depth -gt 12) { throw 'Nested metadata limit' }
    if ($null -eq $Value) { return $null }
    if ($Value -is [string]) { if ($Value.Length -gt 16384) { throw 'Oversized metadata' }; return $Value }
    if ($Value -is [bool] -or $Value -is [int] -or $Value -is [long] -or $Value -is [double]) { return $Value }
    if ($Value -is [datetime] -or $Value -is [datetimeoffset]) { return $Value.ToUniversalTime().ToString('o') }
    if ($Value -is [array]) {
        if ($Value.Count -gt 10000) { throw 'Metadata array limit' }
        return ,@(foreach ($Item in $Value) { ConvertTo-IntuneExpansionValue $Item ($Depth + 1) })
    }
    $Allowed = @('@odata.type', 'templateId', 'templateFamily', 'templateDisplayName', 'templateDisplayVersion', 'runAsAccount', 'deviceRestartBehavior', 'returnCode', 'type', 'operator', 'detectionType', 'requirementType', 'isEnabled', 'enableGlobalAdmins', 'registeringUsers', 'allowedToJoin', 'isAdminConfigurable', 'userType', 'hidePrivacySettings', 'hideEULA', 'skipKeyboardSelectionPage', 'deviceUsageType', 'builtInControls', 'customAuthenticationFactors', 'termsOfUse', 'authenticationStrength', 'id', 'users', 'applications', 'platforms', 'locations', 'devices', 'clientAppTypes', 'signInRiskLevels', 'userRiskLevels', 'servicePrincipalRiskLevels', 'clientApplications', 'includeApplications', 'excludeApplications', 'includeUserActions', 'includeAuthenticationContextClassReferences', 'applicationFilter', 'deviceFilter', 'mode', 'rule', 'includeUsers', 'excludeUsers', 'includeGroups', 'excludeGroups', 'includeRoles', 'excludeRoles', 'includePlatforms', 'excludePlatforms', 'includeLocations', 'excludeLocations', 'includeGuestsOrExternalUsers', 'excludeGuestsOrExternalUsers', 'guestOrExternalUserTypes', 'externalTenants', 'membershipKind', 'members', 'groups')
    $Result = [ordered]@{}
    foreach ($Key in $Allowed) {
        $Item = Get-IntuneValue $Value $Key
        if ($null -ne $Item) { $Result[$Key] = ConvertTo-IntuneExpansionValue $Item ($Depth + 1) }
    }
    return $Result
}

function Test-IntuneSecurityPath {
    param([string]$Path)
    return $Path -cmatch '^((Policy/Config/Defender/(AllowBehaviorMonitoring|AllowCloudProtection|AllowRealtimeMonitoring|AllowScriptScanning|AllowIOAVProtection|AllowOnAccessProtection|PUAProtection|EnableNetworkProtection|CloudBlockLevel|CloudExtendedTimeout|SubmitSamplesConsent|AttackSurfaceReductionRules|AttackSurfaceReductionOnlyExclusions))|(LAPS/Policies/(BackupDirectory|AdministratorAccountName|PasswordAgeDays|PasswordLength|PasswordComplexity|PassphraseLength|PostAuthenticationResetDelay|PostAuthenticationActions|AutomaticAccountManagementEnabled))|(Policy/Config/(DeviceGuard/(EnableVirtualizationBasedSecurity|LsaCfgFlags|RequirePlatformSecurityFeatures)|LocalPoliciesSecurityOptions/(MicrosoftNetworkClient_DigitallySignCommunicationsAlways|MicrosoftNetworkServer_DigitallySignCommunicationsAlways|UserAccountControl_RunAllAdministratorsInAdminApprovalMode)|SmartScreen/(EnableSmartScreenInShell|PreventOverrideForFilesInShell)|LanmanWorkstation/EnableInsecureGuestLogons|MSSecurityGuide/(ConfigureSMBV1ClientDriver|ConfigureSMBV1Server)|InternetExplorer/DisableInternetExplorerLaunchViaCOM|LocalSecurityAuthority/ConfigureLsaProtectedProcess|VirtualizationBasedTechnology/HypervisorEnforcedCodeIntegrity|WindowsPowerShell/TurnOnPowerShellScriptBlockLogging))|(BitLocker/(RequireDeviceEncryption|AllowWarningForOtherDiskEncryption|AllowStandardUserEncryption|SystemDrivesRequireStartupAuthentication|ConfigureRecoveryPasswordRotation))|(Firewall/MdmStore/(DomainProfile|PrivateProfile|PublicProfile)/(EnableFirewall|DefaultInboundAction|DefaultOutboundAction|EnableLogDroppedPackets|EnableLogSuccessConnections|AllowLocalPolicyMerge|AllowLocalIpsecPolicyMerge)))$'
}

function ConvertTo-IntuneSettingFacts {
    param($Instance, $Definitions, [string]$PolicyId, [string]$SettingId, [int]$Depth = 0)
    if ($Depth -gt 12) { throw 'Setting depth limit' }
    $DefinitionId = [string](Get-IntuneValue $Instance 'settingDefinitionId')
    $Definition = @($Definitions | Where-Object { (Get-IntuneValue $_ 'id') -ceq $DefinitionId })
    $Fact = [ordered]@{ id = ($SettingId + ':' + $DefinitionId); parentId = $PolicyId; definitionId = $DefinitionId; resolution = 'UnsupportedDefinition' }
    if ($Definition.Count -eq 1) {
        $Path = (([string](Get-IntuneValue $Definition[0] 'baseUri')).TrimEnd('/') + '/' + ([string](Get-IntuneValue $Definition[0] 'offsetUri')).TrimStart('/')) -creplace '^\./(Device/)?Vendor/MSFT/', ''
        $EpmPath = Get-IntuneEpmPath $Definition[0]
        if ($EpmPath) { $Path = $EpmPath }
        if ($EpmPath -or (Test-IntuneSecurityPath $Path) -or (Test-IntuneExtendedPath $Path)) {
            $Fact['cspUri'] = $Path
            $Fact['definitionVersion'] = [string](Get-IntuneValue $Definition[0] 'version')
            $Choice = Get-IntuneValue $Instance 'choiceSettingValue'
            $Simple = Get-IntuneValue $Instance 'simpleSettingValue'
            $Fact['resolution'] = 'UnresolvedValue'
            $Selected = $Simple
            if ($null -ne $Choice) {
                $Options = @(foreach ($Option in (Get-IntuneValue $Definition[0] 'options' @())) {
                    if ((Get-IntuneValue $Option 'itemId') -ceq (Get-IntuneValue $Choice 'value')) { $Option }
                })
                if ($Options.Count -eq 1) { $Selected = Get-IntuneValue $Options[0] 'optionValue' }
            }
            $Template = Get-IntuneValue $Selected 'settingValueTemplateReference'
            $ChoiceTemplate = Get-IntuneValue $Choice 'settingValueTemplateReference'
            $Scalar = Get-IntuneValue $Selected 'value'
            $UnreviewedAdmx = $Path -cin @('Policy/Config/MSSecurityGuide/ConfigureSMBV1ClientDriver', 'Policy/Config/MSSecurityGuide/ConfigureSMBV1Server', 'Policy/Config/WindowsPowerShell/TurnOnPowerShellScriptBlockLogging')
            if ($Scalar -is [string] -and (Test-IntuneStructuredPath $Path)) {
                try { $Fact['admx'] = ConvertTo-IntuneAdmxMetadata $Scalar -CspPath $Path; $Fact['resolution'] = 'ResolvedAdmx' }
                catch { $Fact['resolution'] = 'UnresolvedAdmx' }
            }
            if ((Get-IntuneValue $Template 'useTemplateDefault' $false) -or (Get-IntuneValue $ChoiceTemplate 'useTemplateDefault' $false)) {
                $Fact['resolution'] = 'UnresolvedTemplateDefault'
            } elseif ($UnreviewedAdmx) {
                $Fact['resolution'] = 'UnresolvedAdmx'
            } elseif (-not (Test-IntuneStructuredPath $Path) -and ($Scalar -is [int] -or $Scalar -is [long] -or $Scalar -is [bool] -or ($Scalar -is [string] -and ((Test-IntuneExtendedPath $Path) -or $Path -in @('LAPS/Policies/AdministratorAccountName', 'Policy/Config/Defender/AttackSurfaceReductionRules', 'Policy/Config/Defender/AttackSurfaceReductionOnlyExclusions'))))) {
                if ($Scalar -is [string] -and $Scalar.Length -gt 16384) { throw 'Setting size limit' }
                $Fact['value'] = $Scalar
                $Fact['resolution'] = 'Resolved'
            }
            if ((Get-IntuneValue $Template 'useTemplateDefault' $false) -or (Get-IntuneValue $ChoiceTemplate 'useTemplateDefault' $false)) { $Fact.Remove('admx') }
        }
    }
    if ($DefinitionId) { $Fact }
    $Ordinal = 0
    foreach ($ValueName in @('choiceSettingValue', 'groupSettingCollectionValue', 'choiceSettingCollectionValue')) {
        foreach ($SettingValue in (Get-IntuneValue $Instance $ValueName @())) {
            foreach ($Child in (Get-IntuneValue $SettingValue 'children' @())) {
                $Ordinal++
                ConvertTo-IntuneSettingFacts $Child $Definitions $PolicyId ($SettingId + '.' + $Ordinal) ($Depth + 1)
            }
        }
    }
}

function ConvertTo-IntuneEndpointModules {
    param($Modules)
    $Fields = @{
        DefenderStatus = @('AMRunningMode', 'AMProductVersion', 'AMEngineVersion', 'AntivirusEnabled', 'RealTimeProtectionEnabled', 'BehaviorMonitorEnabled', 'AntivirusSignatureLastUpdated', 'IsTamperProtected', 'ControlledConfigurationState', 'TamperProtectionSource')
        DefenderPreferences = @('AttackSurfaceReductionRules_Ids', 'AttackSurfaceReductionRules_Actions', 'EnableNetworkProtection', 'PUAProtection', 'DisableRealtimeMonitoring', 'DisableBehaviorMonitoring', 'DisableScriptScanning', 'MAPSReporting', 'EnableControlledFolderAccess')
        FirewallProfiles = @('Name', 'Enabled', 'DefaultInboundAction', 'DefaultOutboundAction', 'LogAllowed', 'LogBlocked')
        BitLockerVolumes = @('MountPoint', 'VolumeType', 'VolumeStatus', 'ProtectionStatus', 'EncryptionPercentage')
        DeviceGuard = @('VirtualizationBasedSecurityStatus', 'SecurityServicesConfigured', 'SecurityServicesRunning')
    }
    $Safe = [ordered]@{}
    foreach ($Module in $Fields.Keys) {
        $InputModule = Get-IntuneValue $Modules $Module
        if ($null -eq $InputModule) { continue }
        $Safe[$Module] = @{ State = [string](Get-IntuneValue $InputModule 'State'); Rows = @(foreach ($InputRow in (Get-IntuneValue $InputModule 'Rows' @())) {
            $OutputRow = [ordered]@{}
            foreach ($Field in $Fields[$Module]) {
                $Value = Get-IntuneValue $InputRow $Field
                if ($null -ne $Value) { $OutputRow[$Field] = ConvertTo-IntuneExpansionValue $Value }
            }
            $OutputRow
        }) }
    }
    return $Safe
}

function Invoke-IntuneExpansion {
    param($Inventory, $States, [scriptblock]$Request, [scriptblock]$Delay, [datetime]$Deadline,
        [bool]$Configuration, [bool]$Rbac, [bool]$Entra, [bool]$Recovery, [bool]$Enrollment, [bool]$Apple)
    $Modules = [ordered]@{}
    if ($Configuration) {
        $Modules['ModernConfigurationPolicies'] = '/beta/deviceManagement/configurationPolicies'
        $Modules['LegacyIntents'] = '/beta/deviceManagement/intents'
        $Modules['LegacyTemplates'] = '/beta/deviceManagement/templates'
        $Modules['AssignmentFilters'] = '/beta/deviceManagement/assignmentFilters'
    }
    if ($Rbac -and $Configuration) { $Modules['ScopeTags'] = '/beta/deviceManagement/roleScopeTags' }
    if ($Enrollment) {
        $Modules['EnrollmentPolicies'] = '/beta/deviceManagement/deviceEnrollmentConfigurations'
        $Modules['AutopilotProfiles'] = '/beta/deviceManagement/windowsAutopilotDeploymentProfiles'
    }
    if ($Entra) {
        $Modules['DeviceRegistrationPolicy'] = '/v1.0/policies/deviceRegistrationPolicy'
        $Modules['ConditionalAccessPolicies'] = '/v1.0/identity/conditionalAccess/policies'
    }
    if ($Recovery) {
        $Modules['LapsMetadata'] = '/v1.0/directory/deviceLocalCredentials?$select=id,lastBackupDateTime,refreshDateTime'
        $Modules['BitLockerMetadata'] = '/v1.0/informationProtection/bitlocker/recoveryKeys'
    }
    if ($Apple) {
        $Modules['ApplePushCertificate'] = '/v1.0/deviceManagement/applePushNotificationCertificate?$select=id,expirationDateTime,lastModifiedDateTime'
        $Modules['VppTokens'] = '/v1.0/deviceAppManagement/vppTokens?$select=id,expirationDateTime,state,lastSyncDateTime,lastSyncStatus,automaticallyUpdateApps'
    }
    foreach ($Module in $Modules.Keys) {
        $Result = Get-IntuneCollection -Uri ('https://graph.microsoft.com' + $Modules[$Module]) -Module $Module -Request $Request -Delay $Delay -Deadline $Deadline -AllowExpansion -Singleton:($Module -in @('DeviceRegistrationPolicy', 'ApplePushCertificate'))
        $Inventory[$Module] = @($Result.Rows)
        $States[$Module] = $Result.Status
    }
    $Children = @()
    if ($Configuration) {
        $Children += @{ Module = 'ModernAssignments'; Parent = 'ModernConfigurationPolicies'; Path = '/beta/deviceManagement/configurationPolicies'; Child = 'assignments' }
        $Children += @{ Module = 'LegacyAssignments'; Parent = 'LegacyIntents'; Path = '/beta/deviceManagement/intents'; Child = 'assignments' }
        $Children += @{ Module = 'NoncomplianceRules'; Parent = 'CompliancePolicies'; Path = '/v1.0/deviceManagement/deviceCompliancePolicies'; Child = 'scheduledActionsForRule' }
    }
    if ($Enrollment) {
        $Children += @{ Module = 'EnrollmentAssignments'; Parent = 'EnrollmentPolicies'; Path = '/beta/deviceManagement/deviceEnrollmentConfigurations'; Child = 'assignments' }
        $Children += @{ Module = 'AutopilotAssignments'; Parent = 'AutopilotProfiles'; Path = '/beta/deviceManagement/windowsAutopilotDeploymentProfiles'; Child = 'assignments' }
    }
    if ($Configuration) { $Children += @{ Module = 'NoncomplianceActions'; Parent = 'NoncomplianceRules'; Path = '/v1.0/deviceManagement/deviceCompliancePolicies'; Child = 'scheduledActionConfigurations' } }
    foreach ($Child in $Children) {
        $Contracts = @($script:IntuneGraphContracts | Where-Object { $_.module -ceq $Child.Module })
        if ($Contracts.Count -ne 1 -or -not (Get-IntuneValue $Contracts[0] 'enabled' $true)) {
            $Inventory[$Child.Module] = @()
            $States[$Child.Module] = @{ State = 'Unsupported'; ApiVersion = $Child.Path.Split('/')[1]; Endpoint = $Child.Path; PagesRead = 0; RowsRead = 0; ScopeComplete = $false; ErrorCode = 'UnverifiedContract'; Details = 'Assignment GET contract quarantined pending exact documentation; no requests sent.'; CompletedParentIds = @() }
            continue
        }
        $Rows = [System.Collections.Generic.List[object]]::new()
        $Completed = [System.Collections.Generic.List[string]]::new()
        $Pages = 0
        $Failure = $States[$Child.Parent].State -ne 'Complete'
        foreach ($Parent in $Inventory[$Child.Parent]) {
            $ParentId = [string]$Parent['id']
            $ParentPath = [uri]::EscapeDataString($ParentId)
            $Identity = $ParentId
            if ($Child.Module -eq 'NoncomplianceActions') {
                $ParentPath = [uri]::EscapeDataString($Parent['parentId']) + '/scheduledActionsForRule/' + $ParentPath
                $Identity = $Parent['parentId'] + '/' + $ParentId
            }
            $Uri = 'https://graph.microsoft.com{0}/{1}/{2}' -f $Child.Path, $ParentPath, $Child.Child
            $Result = Get-IntuneCollection -Uri $Uri -Module $Child.Module -ParentId $Identity -Request $Request -Delay $Delay -Deadline $Deadline -AllowExpansion -MaxRows ([math]::Max(0, 100000 - $Rows.Count))
            foreach ($Row in $Result.Rows) { $Rows.Add($Row) }
            $Pages += $Result.Status.PagesRead
            if ($Result.Status.State -eq 'Complete') { $Completed.Add($Identity) } else { $Failure = $true }
        }
        $Inventory[$Child.Module] = @($Rows.ToArray())
        $States[$Child.Module] = @{ State = $(if ($Failure) { 'Partial' } else { 'Complete' }); ApiVersion = $Child.Path.Split('/')[1]; Endpoint = $Child.Path; PagesRead = [math]::Max(1, $Pages); RowsRead = $Rows.Count; ScopeComplete = $false; ErrorCode = $(if ($Failure) { 'ChildCollectionIncomplete' } else { $null }); Details = 'Visible parent requests only; effective targeting is not inferred.'; CompletedParentIds = @($Completed.ToArray()) }
    }
    if ($Configuration) {
        $Facts = [System.Collections.Generic.List[object]]::new()
        $EpmRules = [System.Collections.Generic.List[object]]::new()
        $Completed = [System.Collections.Generic.List[string]]::new()
        $Pages = 0
        $Failure = $States.ModernConfigurationPolicies.State -ne 'Complete'
        foreach ($Policy in $Inventory.ModernConfigurationPolicies) {
            $PolicyId = [string]$Policy['id']
            $Uri = 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/' + [uri]::EscapeDataString($PolicyId) + '/settings'
            $Settings = Get-IntuneCollection -Uri $Uri -Module RawSettings -Request $Request -Delay $Delay -Deadline $Deadline -AllowExpansion
            $Pages += $Settings.Status.PagesRead
            $PolicyComplete = $Settings.Status.State -eq 'Complete'
            $ExpectedCount = Get-IntuneValue $Policy 'settingCount'
            if ($null -ne $ExpectedCount -and $ExpectedCount -ne $Settings.Rows.Count) { $PolicyComplete = $false }
            foreach ($Setting in $Settings.Rows) {
                $SettingId = [string](Get-IntuneValue $Setting 'id')
                $Definitions = Get-IntuneCollection -Uri ($Uri + '/' + [uri]::EscapeDataString($SettingId) + '/settingDefinitions') -Module RawDefinitions -Request $Request -Delay $Delay -Deadline $Deadline -AllowExpansion
                $Pages += $Definitions.Status.PagesRead
                if ($Definitions.Status.State -ne 'Complete') { $PolicyComplete = $false }
                try {
                    if ((Get-IntuneValue (Get-IntuneValue $Policy 'templateReference') 'templateFamily') -ceq 'endpointSecurityEndpointPrivilegeManagement') {
                        foreach ($Rule in @(ConvertTo-IntuneEpmRules (Get-IntuneValue $Setting 'settingInstance') $Definitions.Rows $PolicyId $SettingId)) {
                            if ($EpmRules.Count -ge 100000) { throw 'EPM rule limit' }
                            $EpmRules.Add($Rule)
                        }
                    }
                    foreach ($Fact in @(ConvertTo-IntuneSettingFacts (Get-IntuneValue $Setting 'settingInstance') $Definitions.Rows $PolicyId $SettingId)) {
                        if ($Facts.Count -ge 100000) { throw 'Setting row limit' }
                        $Facts.Add($Fact)
                    }
                } catch { $PolicyComplete = $false }
            }
            if ($PolicyComplete) { $Completed.Add($PolicyId) } else { $Failure = $true }
        }
        $Inventory['SecuritySettings'] = @($Facts.ToArray())
        $Inventory['EpmRules'] = @($EpmRules.ToArray())
        $States['EpmRules'] = @{ State = $(if ($Failure) { 'Partial' } else { 'Complete' }); ApiVersion = 'beta'; PagesRead = [math]::Max(1, $Pages); RowsRead = $EpmRules.Count; CompletedParentIds = @($Completed.ToArray()); Details = 'Verified grouped EPM name/file/path/type subset; no certificates, arguments, hashes or scripts collected.' }
        $States['SecuritySettings'] = @{ State = $(if ($Failure) { 'Partial' } else { 'Complete' }); ApiVersion = 'beta'; Endpoint = '/deviceManagement/configurationPolicies/{id}/settings'; PagesRead = [math]::Max(1, $Pages); RowsRead = $Facts.Count; ScopeComplete = $false; ErrorCode = $(if ($Failure) { 'SettingsIncomplete' } else { $null }); Details = 'Only reviewed CSP values retained; unsupported definitions have no value. No effective-device claim.'; CompletedParentIds = @($Completed.ToArray()) }
    }
}