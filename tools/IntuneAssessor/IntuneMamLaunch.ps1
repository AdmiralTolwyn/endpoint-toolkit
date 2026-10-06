<#
.SYNOPSIS
    Field allowlists for the opt-in app-protection launch-condition modules.
.DESCRIPTION
    Dot-sourced by IntuneServices.ps1 for -IncludeMamLaunch.
.NOTES
    Author    : Anton Romanyuk
    Requires  : Windows PowerShell 5.1 or PowerShell 7
    Disclaimer: This script is provided "AS IS" with no warranties and confers no rights.
#>


<#
.SYNOPSIS
    Returns the exported fields for an app-protection launch module.
.PARAMETER Module
    MamAndroidLaunchPolicies or MamIosLaunchPolicies. Other modules return an empty list.
.OUTPUTS
    System.String[]
#>
function Get-IntuneMamLaunchFields {
    param([string]$Module)
    $Common = 'id @odata.type displayName version lastModifiedDateTime isAssigned targetedAppManagementLevels appGroupType deviceComplianceRequired appActionIfDeviceComplianceRequired pinRequired disableAppPinIfDevicePinIsSet maximumPinRetries appActionIfMaximumPinRetriesExceeded periodOfflineBeforeAccessCheck periodOfflineBeforeWipeIsEnforced appActionIfUnableToAuthenticateUser maximumAllowedDeviceThreatLevel mobileThreatDefenseRemediationAction mobileThreatDefensePartnerPriority pinRequiredInsteadOfBiometricTimeout allowedOutboundClipboardSharingExceptionLength allowedOutboundClipboardSharingLevel notificationRestriction' -split ' '
    $Common += 'periodOnlineBeforeAccessCheck'
    switch ($Module) {
        'MamAndroidLaunchPolicies' { return $Common + ('requiredAndroidSafetyNetDeviceAttestationType appActionIfAndroidSafetyNetDeviceAttestationFailed requiredAndroidSafetyNetAppsVerificationType appActionIfAndroidSafetyNetAppsVerificationFailed requiredAndroidSafetyNetEvaluationType appActionIfSamsungKnoxAttestationRequired biometricAuthenticationBlocked requireClass3Biometrics requirePinAfterBiometricChange deviceLockRequired appActionIfDeviceLockNotSet appActionIfDevicePasscodeComplexityLessThanLow appActionIfDevicePasscodeComplexityLessThanMedium appActionIfDevicePasscodeComplexityLessThanHigh' -split ' ') }
        'MamIosLaunchPolicies' { return $Common + ('fingerprintBlocked faceIdBlocked genmojiConfigurationState screenCaptureConfigurationState writingToolsConfigurationState allowWidgetContentSync filterOpenInToOnlyManagedApps disableProtectionOfManagedOutboundOpenInData' -split ' ') }
        default { return @() }
    }
}