. (Join-Path $PSScriptRoot 'IntuneMamLaunch.ps1')

function Get-IntuneEpmPath {
    param($Definition)
    $Mappings = @{
        device_vendor_msft_policy_elevationclientsettings_enableepm = 'PrivilegeManagement/ElevationClientSettings/EnableEPM'
        device_vendor_msft_policy_elevationclientsettings_defaultelevationresponse = 'PrivilegeManagement/ElevationClientSettings/DefaultElevationResponse'
        device_vendor_msft_policy_elevationclientsettings_senddata = 'PrivilegeManagement/ElevationClientSettings/SendData'
        device_vendor_msft_policy_elevationclientsettings_reportingscope = 'PrivilegeManagement/ElevationClientSettings/ReportingScope'
    }
    $Identity = [string](Get-IntuneValue $Definition 'id')
    if ($Mappings.ContainsKey($Identity) -and (Get-IntuneValue $Definition 'offsetUri') -ceq $Mappings[$Identity]) { return $Mappings[$Identity] }
    return $null
}

function Get-IntuneServiceFields {
    param([string]$Module)
    $LaunchFields = @(Get-IntuneMamLaunchFields $Module)
    if ($LaunchFields.Count) { return $LaunchFields }
    switch ($Module) {
        'AppProtectionPolicies' { return ('id @odata.type displayName createdDateTime lastModifiedDateTime version isAssigned deployedAppCount pinRequired minimumPinLength maximumPinRetries simplePinBlocked pinCharacterSet disableAppPinIfDevicePinIsSet organizationalCredentialsRequired allowedInboundDataTransferSources allowedOutboundDataTransferDestinations allowedOutboundClipboardSharingLevel dataBackupBlocked saveAsBlocked allowedDataStorageLocations printBlocked contactSyncBlocked deviceComplianceRequired managedBrowserToOpenLinksRequired managedBrowser periodOfflineBeforeAccessCheck periodOfflineBeforeWipeIsEnforced minimumRequiredOsVersion minimumWarningOsVersion minimumRequiredAppVersion minimumRequiredSdkVersion minimumRequiredPatchVersion encryptAppData disableAppEncryptionIfDeviceEncryptionIsEnabled screenCaptureBlocked appDataEncryptionType' -split ' ') }
        'RemoteAssistanceSettings' { return @('id', '@odata.type', 'remoteAssistanceState', 'allowSessionsToUnenrolledDevices', 'blockChat') }
        'ThreatConnectors' { return ('id @odata.type lastHeartbeatDateTime partnerState androidMobileApplicationManagementEnabled iosMobileApplicationManagementEnabled androidEnabled iosEnabled windowsEnabled androidDeviceBlockedOnMissingPartnerData iosDeviceBlockedOnMissingPartnerData windowsDeviceBlockedOnMissingPartnerData partnerUnsupportedOsVersionBlocked partnerUnresponsivenessThresholdInDays allowPartnerToCollectIOSApplicationMetadata allowPartnerToCollectIOSPersonalApplicationMetadata microsoftDefenderForEndpointAttachEnabled' -split ' ') }
        'MamAssignments' { return @('id', '@odata.type', 'target') }
        'MamApps' { return @('id', '@odata.type', 'mobileAppIdentifier') }
        'ManagedAppConfigurations' { return ('id @odata.type displayName createdDateTime lastModifiedDateTime version isAssigned deployedAppCount' -split ' ') }
        'DeviceAppConfigurations' { return ('id @odata.type displayName createdDateTime lastModifiedDateTime version targetedMobileApps' -split ' ') }
        'AppConfigurationAssignments' { return @('id', '@odata.type', 'target') }
        'DeviceAppConfigurationAssignments' { return @('id', '@odata.type', 'target') }
        'AppConfigurationApps' { return @('id', '@odata.type', 'mobileAppIdentifier') }
        'WindowsAppProtectionPolicies' { return ('id @odata.type displayName version isAssigned deployedAppCount printBlocked allowedInboundDataTransferSources allowedOutboundClipboardSharingLevel allowedOutboundDataTransferDestinations maximumAllowedDeviceThreatLevel appActionIfUnableToAuthenticateUser mobileThreatDefenseRemediationAction minimumRequiredOsVersion minimumWarningOsVersion maximumRequiredOsVersion minimumRequiredAppVersion minimumRequiredSdkVersion periodOfflineBeforeWipeIsEnforced periodOfflineBeforeAccessCheck' -split ' ') }
        'PlatformCompliancePolicies' { return @('id', '@odata.type', 'displayName', 'version', 'passwordRequired', 'storageRequireEncryption', 'securityBlockJailbrokenDevices', 'securityRequireIntuneAppIntegrity', 'osMinimumVersion', 'osMaximumVersion', 'minAndroidSecurityPatchLevel', 'securityRequireSafetyNetAttestationBasicIntegrity', 'securityRequireSafetyNetAttestationCertifiedDevice', 'securityRequiredAndroidSafetyNetEvaluationType', 'requireNoPendingSystemUpdates') }
        'ModernCompliancePolicies' { return @('id', '@odata.type', 'name', 'platforms', 'technologies', 'settingCount', 'isAssigned') }
        'TunnelSites' { return @('id', '@odata.type', 'displayName', 'upgradeAutomatically', 'upgradeAvailable', 'upgradeWindowUtcOffsetInMinutes', 'upgradeWindowStartTime', 'upgradeWindowEndTime') }
        'TunnelServers' { return @('id', '@odata.type', 'displayName', 'tunnelServerHealthStatus', 'lastCheckinDateTime', 'agentImageDigest', 'serverImageDigest', 'deploymentMode') }
        default { return @() }
    }
}

function Test-IntuneServicePath {
    param([string]$Path)
    if ($Path -cin @('/beta/deviceAppManagement/androidManagedAppProtections', '/beta/deviceAppManagement/iosManagedAppProtections')) { return $true }
    if ($Path -cmatch '^/beta/deviceManagement/microsoftTunnelSites(/[^/]+/microsoftTunnelServers)?$') { return $true }
    if ($Path -cin @('/beta/deviceAppManagement/windowsManagedAppProtections', '/beta/deviceManagement/deviceCompliancePolicies', '/beta/deviceManagement/compliancePolicies')) { return $true }
    if ($Path -cmatch '^/v1\.0/deviceAppManagement/targetedManagedAppConfigurations(/[^/]+/(assignments|apps))?$' -or $Path -cmatch '^/v1\.0/deviceAppManagement/mobileAppConfigurations(/[^/]+/assignments)?$') { return $true }
    return $Path -ceq '/beta/deviceManagement/remoteAssistanceSettings' -or $Path -ceq '/v1.0/deviceManagement/mobileThreatDefenseConnectors' -or $Path -ceq '/v1.0/deviceAppManagement/managedAppPolicies' -or $Path -cmatch '^/v1\.0/deviceAppManagement/(androidManagedAppProtections|iosManagedAppProtections)/[^/]+/(assignments|apps)$'
}

function Invoke-IntuneTunnel {
    param($Inventory, $States, [scriptblock]$Request, [scriptblock]$Delay, [datetime]$Deadline)
    $Root = 'https://graph.microsoft.com/beta/deviceManagement/microsoftTunnelSites'
    $Sites = Get-IntuneCollection -Uri $Root -Module TunnelSites -Request $Request -Delay $Delay -Deadline $Deadline -AllowExpansion
    $Inventory['TunnelSites'] = @($Sites.Rows)
    $States['TunnelSites'] = $Sites.Status
    $Rows = [Collections.Generic.List[object]]::new()
    $Completed = [Collections.Generic.List[string]]::new()
    $Failure = $Sites.Status.State -ne 'Complete'
    $Pages = 0
    foreach ($Site in $Sites.Rows) {
        $Identity = [string](Get-IntuneValue $Site 'id')
        $Result = Get-IntuneCollection -Uri ($Root + '/' + [uri]::EscapeDataString($Identity) + '/microsoftTunnelServers') -Module TunnelServers -ParentId $Identity -Request $Request -Delay $Delay -Deadline $Deadline -AllowExpansion -MaxRows ([math]::Max(0, 100000 - $Rows.Count))
        foreach ($Row in $Result.Rows) { $Rows.Add($Row) }
        $Pages += $Result.Status.PagesRead
        if ($Result.Status.State -eq 'Complete') { $Completed.Add($Identity) } else { $Failure = $true }
    }
    $Inventory['TunnelServers'] = @($Rows.ToArray())
    $States['TunnelServers'] = @{ State = $(if ($Failure) { 'Partial' } else { 'Complete' }); ApiVersion = 'beta'; PagesRead = [math]::Max(1, $Pages); RowsRead = $Rows.Count; CompletedParentIds = @($Completed.ToArray()); Details = 'Visible Tunnel server metadata only; no probes, upgrades, log actions or configuration changes.' }
}

function Invoke-IntuneServices {
    param($Inventory, $States, [scriptblock]$Request, [scriptblock]$Delay, [datetime]$Deadline, [bool]$Mam, [bool]$RemoteHelp, [bool]$Connectors, [bool]$AppConfiguration, [bool]$PlatformCompliance, [bool]$MamLaunch)
    foreach ($Module in @('AppProtectionPolicies', 'RemoteAssistanceSettings', 'ThreatConnectors')) {
        $Enabled = switch ($Module) { 'AppProtectionPolicies' { $Mam }; 'RemoteAssistanceSettings' { $RemoteHelp }; 'ThreatConnectors' { $Connectors } }
        if (-not $Enabled) {
            $Inventory[$Module] = @()
            $States[$Module] = @{ State = 'NotRequested'; ApiVersion = ''; RowsRead = 0; PagesRead = 0; CompletedParentIds = @() }
            continue
        }
        $Path = switch ($Module) { 'AppProtectionPolicies' { '/v1.0/deviceAppManagement/managedAppPolicies' }; 'RemoteAssistanceSettings' { '/beta/deviceManagement/remoteAssistanceSettings' }; 'ThreatConnectors' { '/v1.0/deviceManagement/mobileThreatDefenseConnectors' } }
        $Result = Get-IntuneCollection -Uri ('https://graph.microsoft.com' + $Path) -Module $Module -Request $Request -Delay $Delay -Deadline $Deadline -AllowExpansion -Singleton:($Module -eq 'RemoteAssistanceSettings')
        $Inventory[$Module] = @($Result.Rows)
        $States[$Module] = $Result.Status
    }
    $Additional = [ordered]@{}
    if ($MamLaunch) {
        $Additional['MamAndroidLaunchPolicies'] = '/beta/deviceAppManagement/androidManagedAppProtections'
        $Additional['MamIosLaunchPolicies'] = '/beta/deviceAppManagement/iosManagedAppProtections'
    } else {
        foreach ($Module in @('MamAndroidLaunchPolicies', 'MamIosLaunchPolicies')) {
            $Inventory[$Module] = @()
            $States[$Module] = @{ State = 'NotRequested'; ApiVersion = ''; PagesRead = 0; RowsRead = 0; CompletedParentIds = @() }
        }
    }
    if ($Mam) { $Additional['WindowsAppProtectionPolicies'] = '/beta/deviceAppManagement/windowsManagedAppProtections' }
    if ($PlatformCompliance) {
        $Additional['PlatformCompliancePolicies'] = '/beta/deviceManagement/deviceCompliancePolicies'
        $Additional['ModernCompliancePolicies'] = '/beta/deviceManagement/compliancePolicies'
    }
    if ($AppConfiguration) {
        $Additional['ManagedAppConfigurations'] = '/v1.0/deviceAppManagement/targetedManagedAppConfigurations?$select=id,displayName,createdDateTime,lastModifiedDateTime,version,isAssigned,deployedAppCount'
        $Additional['DeviceAppConfigurations'] = '/v1.0/deviceAppManagement/mobileAppConfigurations?$select=id,displayName,createdDateTime,lastModifiedDateTime,version,targetedMobileApps'
    }
    foreach ($Module in $Additional.Keys) {
        $Result = Get-IntuneCollection -Uri ('https://graph.microsoft.com' + $Additional[$Module]) -Module $Module -Request $Request -Delay $Delay -Deadline $Deadline -AllowExpansion
        $Inventory[$Module] = @($Result.Rows)
        $States[$Module] = $Result.Status
    }
    if ($AppConfiguration) {
        foreach ($Child in @(
            @{ Module = 'AppConfigurationAssignments'; Parent = 'ManagedAppConfigurations'; Family = 'targetedManagedAppConfigurations'; Leaf = 'assignments' },
            @{ Module = 'AppConfigurationApps'; Parent = 'ManagedAppConfigurations'; Family = 'targetedManagedAppConfigurations'; Leaf = 'apps' },
            @{ Module = 'DeviceAppConfigurationAssignments'; Parent = 'DeviceAppConfigurations'; Family = 'mobileAppConfigurations'; Leaf = 'assignments' }
        )) {
            $Rows = [Collections.Generic.List[object]]::new()
            $Completed = [Collections.Generic.List[string]]::new()
            $Failure = $States[$Child.Parent].State -ne 'Complete'
            $Pages = 0
            foreach ($Policy in $Inventory[$Child.Parent]) {
                $Identity = [string](Get-IntuneValue $Policy 'id')
                $Uri = 'https://graph.microsoft.com/v1.0/deviceAppManagement/' + $Child.Family + '/' + [uri]::EscapeDataString($Identity) + '/' + $Child.Leaf
                $Result = Get-IntuneCollection -Uri $Uri -Module $Child.Module -ParentId $Identity -Request $Request -Delay $Delay -Deadline $Deadline -AllowExpansion -MaxRows ([math]::Max(0, 100000 - $Rows.Count))
                foreach ($Row in $Result.Rows) { $Rows.Add($Row) }
                $Pages += $Result.Status.PagesRead
                if ($Result.Status.State -eq 'Complete') { $Completed.Add($Identity) } else { $Failure = $true }
            }
            $Inventory[$Child.Module] = @($Rows.ToArray())
            $States[$Child.Module] = @{ State = $(if ($Failure) { 'Partial' } else { 'Complete' }); ApiVersion = 'v1.0'; PagesRead = [math]::Max(1, $Pages); RowsRead = $Rows.Count; CompletedParentIds = @($Completed.ToArray()); Details = 'App configuration metadata only; arbitrary key/value or XML payloads excluded.' }
        }
    }
    if (-not $Mam) { return }
    foreach ($Module in @('MamAssignments', 'MamApps')) {
        $Rows = [Collections.Generic.List[object]]::new()
        $Completed = [Collections.Generic.List[string]]::new()
        $Failure = $States.AppProtectionPolicies.State -ne 'Complete'
        $Pages = 0
        foreach ($Policy in $Inventory.AppProtectionPolicies) {
            $Family = switch -CaseSensitive (Get-IntuneValue $Policy '@odata.type') {
                '#microsoft.graph.androidManagedAppProtection' { 'androidManagedAppProtections' }
                '#microsoft.graph.iosManagedAppProtection' { 'iosManagedAppProtections' }
                default { $null }
            }
            if (-not $Family) { continue }
            $Identity = [string](Get-IntuneValue $Policy 'id')
            $Leaf = if ($Module -eq 'MamAssignments') { 'assignments' } else { 'apps' }
            $Uri = 'https://graph.microsoft.com/v1.0/deviceAppManagement/' + $Family + '/' + [uri]::EscapeDataString($Identity) + '/' + $Leaf
            $Result = Get-IntuneCollection -Uri $Uri -Module $Module -ParentId $Identity -Request $Request -Delay $Delay -Deadline $Deadline -AllowExpansion -MaxRows ([math]::Max(0, 100000 - $Rows.Count))
            foreach ($Row in $Result.Rows) { $Rows.Add($Row) }
            $Pages += $Result.Status.PagesRead
            if ($Result.Status.State -eq 'Complete') { $Completed.Add($Identity) } else { $Failure = $true }
        }
        $Inventory[$Module] = @($Rows.ToArray())
        $States[$Module] = @{ State = $(if ($Failure) { 'Partial' } else { 'Complete' }); ApiVersion = 'v1.0'; PagesRead = [math]::Max(1, $Pages); RowsRead = $Rows.Count; CompletedParentIds = @($Completed.ToArray()); Details = 'Visible iOS/Android app protection children only; no user registration or effective targeting inferred.' }
    }
}