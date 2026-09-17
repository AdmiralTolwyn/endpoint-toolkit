#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Invoke-IntuneDiscovery.ps1') -LibraryOnly
function Assert-MamLaunch([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
$script:MamLaunchRequests = [Collections.Generic.List[string]]::new()
$LaunchRequest = {
    param($Address)
    $script:MamLaunchRequests.Add($Address)
    $Rows = switch (([uri]$Address).AbsolutePath) {
        '/beta/deviceAppManagement/androidManagedAppProtections' { @(@{ id = 'android-policy'; '@odata.type' = '#microsoft.graph.androidManagedAppProtection'; displayName = 'Synthetic Android APP'; deviceComplianceRequired = $true; appActionIfDeviceComplianceRequired = 'block'; pinRequired = $true; disableAppPinIfDevicePinIsSet = $false; maximumPinRetries = 5; appActionIfMaximumPinRetriesExceeded = 'block'; periodOfflineBeforeAccessCheck = 'P7D'; periodOfflineBeforeWipeIsEnforced = 'P30D'; appActionIfUnableToAuthenticateUser = 'block'; requiredAndroidSafetyNetDeviceAttestationType = 'basicIntegrityAndDeviceCertification'; appActionIfAndroidSafetyNetDeviceAttestationFailed = 'block'; requiredAndroidSafetyNetAppsVerificationType = 'enabled'; appActionIfAndroidSafetyNetAppsVerificationFailed = 'block'; customSettings = @{ password = 'SECRET' }; allowedOutboundClipboardSharingExceptionLength = 0; notificationRestriction = 'blockOrganizationalData' }) }
        '/beta/deviceAppManagement/iosManagedAppProtections' { @(@{ id = 'ios-policy'; '@odata.type' = '#microsoft.graph.iosManagedAppProtection'; displayName = 'Synthetic iOS APP'; deviceComplianceRequired = $false; appActionIfDeviceComplianceRequired = 'wipe'; genmojiConfigurationState = 'unknownFutureValue'; exemptedAppProtocols = @(@{ value = 'SECRET' }) }) }
        default { @() }
    }
    @{ StatusCode = 200; Body = (@{ value = @($Rows) } | ConvertTo-Json -Depth 20 | ConvertFrom-Json) }
}
$Document = Invoke-IntuneDiscoveryCore -SelectedTenant '22222222-2222-4222-8222-222222222222' -MamLaunch $true -Request $LaunchRequest -Requirements @{ ScopeConfirmed = $true; ScopeDescription = 'Selected synthetic APP launch policies'; Assessor = 'Offline test'; MaxCollectionAgeHours = 24; ConfigurationReview = @{ MamReferenceLevel = 2; MamLaunchPolicyIds = @('android:android-policy', 'ios:ios-policy') } }
Assert-MamLaunch ($Document.Inventory.ManagedDevices.Count -eq 0) 'MAM launch depends on enrolled devices'
Assert-MamLaunch ($Document.CollectionStatus.MamAndroidLaunchPolicies.ApiVersion -eq 'beta') 'Launch API version lost'
Assert-MamLaunch ($Document.Inventory.MamIosLaunchPolicies[0].deviceComplianceRequired -eq $false) 'Explicit false lost'
Assert-MamLaunch ($Document.Inventory.MamIosLaunchPolicies[0].genmojiConfigurationState -eq 'unknownFutureValue') 'Future enum dropped instead of retained'
Assert-MamLaunch (-not ($Document | ConvertTo-Json -Depth 30).Contains('SECRET')) 'Unreviewed payload leaked'
Assert-MamLaunch ($Document.CollectionStatus.AppProtectionPolicies.State -eq 'NotRequested') 'Launch replaced v1 APP collection'
Assert-MamLaunch (-not (Test-IntuneUri 'https://graph.microsoft.com/beta/deviceAppManagement/androidManagedAppProtections')) 'Launch endpoint enabled by default'
Assert-MamLaunch (-not (Test-IntuneUri 'https://graph.microsoft.com/beta/deviceAppManagement/androidManagedAppProtections/id/assign' -AllowExpansion)) 'Action endpoint enabled'
$script:MamLaunchRequests.Clear()
$Default = Invoke-IntuneDiscoveryCore -SelectedTenant '22222222-2222-4222-8222-222222222222' -Request $LaunchRequest
Assert-MamLaunch ($Default.CollectionStatus.MamAndroidLaunchPolicies.State -eq 'NotRequested' -and @($script:MamLaunchRequests | Where-Object { $_ -match '/beta/deviceAppManagement/(android|ios)ManagedAppProtections' }).Count -eq 0) 'Default run collected launch policies'
$Failed = Get-IntuneCollection -Uri 'https://graph.microsoft.com/beta/deviceAppManagement/androidManagedAppProtections' -Module MamAndroidLaunchPolicies -AllowExpansion -Request { @{ StatusCode = 403 } }
Assert-MamLaunch ($Failed.Status.State -eq 'Error') 'Denied MAM read inferred absence'
if ($env:ASSAY_INTUNE_MAM_LAUNCH_FIXTURE) { [IO.File]::WriteAllText($env:ASSAY_INTUNE_MAM_LAUNCH_FIXTURE, ($Document | ConvertTo-Json -Depth 30), [Text.UTF8Encoding]::new($false)) }
Write-Output 'PASS: independent beta APP launch reads, type projection, explicit false/unknown values, privacy and opt-in boundaries.'