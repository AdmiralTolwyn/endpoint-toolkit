<#
.SYNOPSIS
    Changes a disabled dmwappushservice to Automatic to address blocked MDM sync.

.DESCRIPTION
    Implements Microsoft's fix for Intune clients with dmwappushservice disabled.
    Rechecks the service before making a change and verifies the startup mode
    afterwards. Manual and Automatic are unchanged, including on repeated runs.
    Does not start/restart services, restart IME, change enrollment, clear policy
    caches or request sync. Success confirms startup configuration only.

.INPUTS
    None.

.OUTPUTS
    System.String. One status or error message for the management agent.

.EXAMPLE
    .\Remediate-IntuneMdmSyncService.ps1

    Changes Disabled to Automatic and reads back the result.

.NOTES
    Run in 64-bit Windows PowerShell 5.1 as SYSTEM or an elevated administrator.
    Exit 0: the change was verified or the service was already Auto/Manual.
    Exit 1: a query, change or verification failed; investigate the output.
    Pair with Detect-IntuneMdmSyncService.ps1. No parameters or dependencies.
    Investigate any policy or script that disables the service again. Confirm
    actual MDM sync separately after remediation.

.LINK
    https://learn.microsoft.com/en-us/troubleshoot/mem/intune/device-management/cannot-sync-windows-10-devices
#>
#Requires -Version 5.1

$ErrorActionPreference = 'Stop'
$serviceName = 'dmwappushservice'

try {
    $service = Get-CimInstance -ClassName Win32_Service -Filter "Name='$serviceName'" -OperationTimeoutSec 10 -ErrorAction Stop
    if ($null -eq $service) { throw "Service '$serviceName' was not found." }

    if ($service.StartMode -in @('Auto', 'Manual')) {
        Write-Output "$serviceName is not disabled (startup: $($service.StartMode)). No change needed."
        exit 0
    }

    if ($service.StartMode -ne 'Disabled') {
        throw "Unexpected startup mode '$($service.StartMode)' for '$serviceName'."
    }

    Set-Service -Name $serviceName -StartupType Automatic -ErrorAction Stop
    $service = Get-CimInstance -ClassName Win32_Service -Filter "Name='$serviceName'" -OperationTimeoutSec 10 -ErrorAction Stop
    if ($null -eq $service -or $service.StartMode -ne 'Auto') {
        throw "Startup mode for '$serviceName' could not be verified as Automatic."
    }

    Write-Output "$serviceName changed from Disabled to Automatic; verified."
    exit 0
}
catch {
    Write-Output "Remediation error: $($_.Exception.Message)"
    exit 1
}