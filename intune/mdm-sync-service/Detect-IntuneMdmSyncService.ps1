<#
.SYNOPSIS
    Detects a disabled dmwappushservice that can prevent Intune MDM sync.

.DESCRIPTION
    Read-only detection for the disabled-service issue documented by Microsoft.
    Flags Disabled only. Manual and Automatic are left alone; a stopped service
    is not necessarily disabled. Does not test enrollment or actual sync health.

.INPUTS
    None.

.OUTPUTS
    System.String. One status or error message for the management agent.

.EXAMPLE
    .\Detect-IntuneMdmSyncService.ps1

    Checks the service startup mode without changing the device.

.NOTES
    Run in 64-bit Windows PowerShell 5.1 as SYSTEM or an elevated administrator.
    Intune Remediations runs the paired remediation only for detection exit 1.
    Exit 0: the service is not disabled; this does not prove sync works.
    Exit 1: the service is disabled and needs the documented fix.
    Exit 2: the service is missing, unreadable or has an unexpected startup mode.
    Pair with Remediate-IntuneMdmSyncService.ps1. No parameters or dependencies.

.LINK
    https://learn.microsoft.com/en-us/troubleshoot/mem/intune/device-management/cannot-sync-windows-10-devices
#>
#Requires -Version 5.1

$ErrorActionPreference = 'Stop'
$serviceName = 'dmwappushservice'

try {
    $service = Get-CimInstance -ClassName Win32_Service -Filter "Name='$serviceName'" -OperationTimeoutSec 10 -ErrorAction Stop
    if ($null -eq $service) { throw "Service '$serviceName' was not found." }

    if ($service.StartMode -eq 'Disabled') {
        Write-Output "$serviceName is Disabled. Remediation is required."
        exit 1
    }

    if ($service.StartMode -notin @('Auto', 'Manual')) {
        throw "Unexpected startup mode '$($service.StartMode)' for '$serviceName'."
    }

    Write-Output "$serviceName is not disabled (startup: $($service.StartMode))."
    exit 0
}
catch {
    Write-Output "Detection error: $($_.Exception.Message)"
    exit 2
}