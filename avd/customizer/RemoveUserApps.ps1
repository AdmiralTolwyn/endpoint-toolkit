<#
.SYNOPSIS
    Removes per-user-only inbox AppX packages that have NO matching provisioned
    package on the image, so Sysprep does not fail with the well-known
    "package was installed for a user, but not provisioned for all users" error.

.DESCRIPTION
    During AVD / Windows 365 image bake the SYSTEM account (and any locally signed-in
    user) can end up with AppX packages that exist only in their per-user store. If
    those packages are not also in Get-AppxProvisionedPackage, Sysprep /generalize
    aborts with 0x80073CF2 / "the package(s) cannot be found".

    This script:
      1. Snapshots Get-AppxProvisionedPackage as the allow-list (Name + Version).
      2. Walks Get-AppxPackage TWO times (to handle dependency ordering).
      3. Removes any user-context package whose SignatureKind is NOT 'System' AND
         whose (Name, Version) tuple is missing from the provisioned snapshot.

    Logic credit: Michael Niehaus's classic Sysprep cleanup pattern.

.NOTES
    File:    avd/customizer/RemoveUserApps.ps1
    Author:  Anton Romanyuk (logic based on Michael Niehaus)
    Version: 2.0.0
    Context: Azure Image Builder / Packer customizer. Runs as SYSTEM, late stage,
             immediately before AdminSysPrep + Sysprep.
    Requires: Windows 10/11 / Server, PowerShell 5.1+, admin.

.DISCLAIMER
    This script is provided "AS IS" with no warranties and confers no rights.
    It is not supported under any Microsoft standard support program or service.
    Use of this script is entirely at your own risk. The customer is solely
    responsible for testing and validating this script in their environment
    before deploying to production.

.EXAMPLE
    .\RemoveUserApps.ps1
#>

#Requires -RunAsAdministrator
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO'
    )
    $ts    = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
    $color = switch ($Level) { 'WARN' {'Yellow'} 'ERROR' {'Red'} 'SUCCESS' {'Green'} default {'Gray'} }
    Write-Host "[$ts] [$Level] [RemoveUserApps] $Message" -ForegroundColor $color
}

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
Write-Log "Starting RemoveUserApps (Sysprep prep) phase" -Level SUCCESS

try {
    Write-Log "Snapshotting provisioned packages (system-wide allow-list)"
    $provisioned = @(Get-AppxProvisionedPackage -Online)
    Write-Log "Found $($provisioned.Count) provisioned package(s)"

    $removedCount = 0
    for ($pass = 1; $pass -le 2; $pass++) {
        Write-Log "Removal pass $pass of 2"
        $userPackages = Get-AppxPackage | Where-Object { $_.SignatureKind -ne 'System' }

        foreach ($app in $userPackages) {
            $isProvisioned = $provisioned | Where-Object {
                $_.DisplayName -eq $app.Name -and $_.Version -eq $app.Version
            }

            if ($null -eq $isProvisioned) {
                Write-Log "Removing non-provisioned user app: $($app.Name) ($($app.Version))" -Level WARN
                try {
                    Remove-AppxPackage -Package $app.PackageFullName -ErrorAction Stop
                    $removedCount++
                    Write-Log "Removed $($app.Name)" -Level SUCCESS
                }
                catch {
                    Write-Log "Failed to remove $($app.Name): $($_.Exception.Message)" -Level WARN
                }
            }
        }
    }

    if ($removedCount -eq 0) {
        Write-Log "Image is clean - no non-provisioned user apps found" -Level SUCCESS
    } else {
        Write-Log "Cleanup complete - removed $removedCount package(s)" -Level SUCCESS
    }
}
catch {
    Write-Log "RemoveUserApps failed unexpectedly: $($_.Exception.Message)" -Level ERROR
    exit 1
}

$stopwatch.Stop()
Write-Log "RemoveUserApps completed in $($stopwatch.Elapsed)" -Level SUCCESS
exit 0
