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