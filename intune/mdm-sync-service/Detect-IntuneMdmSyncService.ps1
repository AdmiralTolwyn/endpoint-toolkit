<#
.SYNOPSIS
    Detects a disabled dmwappushservice that can prevent Intune MDM sync.

.DESCRIPTION
    Detects the disabled-service issue documented by Microsoft without changing
    management configuration. Appends status messages to its own IME log file.
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
    Log: C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\Detect-IntuneMdmSyncService.log.
    Creates the folder if missing. Log-write failures warn without changing exit codes.

.LINK
    https://learn.microsoft.com/en-us/troubleshoot/mem/intune/device-management/cannot-sync-windows-10-devices
#>
#Requires -Version 5.1

param()

function Write-ClientSyncLog {
    <#
    .SYNOPSIS
        Appends a UTC-stamped message without changing the success output stream.
    .PARAMETER Message
        Text or a report object to append; objects use compressed JSON.
    .PARAMETER OutputMessage
        Also emits the original message on the success stream for service scripts.
    .NOTES
        A write failure warns once per run; it does not change the operation result.
    #>
    param([object]$Message, [switch]$OutputMessage)
    try {
        [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($script:ClientSyncLogPath))
        $text = $Message
        if ($Message -isnot [string]) { $text = $Message | ConvertTo-Json -Depth 10 -Compress }
        $line = '{0} [PID:{1}] {2}{3}' -f [DateTime]::UtcNow.ToString('o'), $PID, $text, [Environment]::NewLine
        [IO.File]::AppendAllText($script:ClientSyncLogPath, $line, (New-Object System.Text.UTF8Encoding($false)))
    }
    catch {
        if (-not $script:ClientSyncLogWarningWritten) {
            $script:ClientSyncLogWarningWritten = $true
            Write-Warning "Cannot write log '$script:ClientSyncLogPath': $($_.Exception.Message)" -WarningAction Continue
        }
    }
    if ($OutputMessage) { Write-Output $Message }
}

$ErrorActionPreference = 'Stop'
$serviceName = 'dmwappushservice'
$script:ClientSyncLogPath = Join-Path $env:ProgramData 'Microsoft\IntuneManagementExtension\Logs\Detect-IntuneMdmSyncService.log'
$script:ClientSyncLogWarningWritten = $false
Write-ClientSyncLog 'Starting disabled-service detection.'

try {
    $service = Get-CimInstance -ClassName Win32_Service -Filter "Name='$serviceName'" -OperationTimeoutSec 10 -ErrorAction Stop
    if ($null -eq $service) { throw "Service '$serviceName' was not found." }

    if ($service.StartMode -eq 'Disabled') {
        Write-ClientSyncLog "$serviceName is Disabled. Remediation is required." -OutputMessage
        Write-ClientSyncLog 'Finished detection; exit code=1.'
        exit 1
    }

    if ($service.StartMode -notin @('Auto', 'Manual')) {
        throw "Unexpected startup mode '$($service.StartMode)' for '$serviceName'."
    }

    Write-ClientSyncLog "$serviceName is not disabled (startup: $($service.StartMode))." -OutputMessage
    Write-ClientSyncLog 'Finished detection; exit code=0.'
    exit 0
}
catch {
    Write-ClientSyncLog "Detection error: $($_.Exception.Message)" -OutputMessage
    Write-ClientSyncLog 'Finished detection; exit code=2.'
    exit 2
}