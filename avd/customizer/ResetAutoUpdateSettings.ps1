<#
.SYNOPSIS
    Reverts the image-bake hardening applied by DisableAutoUpdates.ps1 / UpdateWinGet.ps1
    so a deployed AVD / Windows 365 host receives Windows Updates and Store updates again.

.DESCRIPTION
    Designed to run AFTER image capture (e.g. as a Run-Command on the deployed host or
    as the very last AIB step on golden images that should ship with updates ENABLED).

    Reverses three image-bake registry policies:
      * HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoUpdate
      * HKLM\SOFTWARE\Policies\Microsoft\WindowsStore\AutoDownload
      * HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager\Debug
        \ContentDeliveryAllowedOverride

    And re-enables the \Microsoft\Windows\WindowsUpdate\Scheduled Start task so WU
    background scans resume.

.NOTES
    File:    avd/customizer/ResetAutoUpdateSettings.ps1
    Author:  Anton Romanyuk
    Version: 2.0.0
    Context: Run AFTER image capture - either as the last AIB customizer step (when
             you want updates enabled in the captured image) or as a deployment-time
             Run-Command on hosts that need their hardening reverted.
    Requires: Windows 10/11 / Server, PowerShell 5.1+, admin.

.DISCLAIMER
    This script is provided "AS IS" with no warranties and confers no rights.
    It is not supported under any Microsoft standard support program or service.
    Use of this script is entirely at your own risk. The customer is solely
    responsible for testing and validating this script in their environment
    before deploying to production.

.EXAMPLE
    .\ResetAutoUpdateSettings.ps1
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
    Write-Host "[$ts] [$Level] [ResetAutoUpdateSettings] $Message" -ForegroundColor $color
}

function Remove-RegPolicy {
<#
.SYNOPSIS
    Idempotently removes a single registry value if present, logging WARN/INFO/SUCCESS appropriately.
.PARAMETER Path
    Full registry path.
.PARAMETER Name
    Value name to delete.
.PARAMETER Description
    Friendly label used only for log output.
#>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Description
    )
    try {
        if (-not (Test-Path -LiteralPath $Path)) {
            Write-Log "$Description : path '$Path' not present (nothing to revert)"
            return
        }
        $prop = Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction SilentlyContinue
        if (-not $prop) {
            Write-Log "$Description : value '$Name' not present (already reverted)"
            return
        }
        Remove-ItemProperty -LiteralPath $Path -Name $Name -Force
        Write-Log "$Description : removed '$Name' from '$Path'" -Level SUCCESS
    }
    catch {
        Write-Log "$Description : failed - $($_.Exception.Message)" -Level WARN
    }
}

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
Write-Log "Starting ResetAutoUpdateSettings phase" -Level SUCCESS

Remove-RegPolicy -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU' `
                 -Name 'NoAutoUpdate' `
                 -Description 'Re-enable Windows Auto Updates'

Remove-RegPolicy -Path 'HKLM:\SOFTWARE\Policies\Microsoft\WindowsStore' `
                 -Name 'AutoDownload' `
                 -Description 'Re-enable Microsoft Store auto-downloads'

Remove-RegPolicy -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager\Debug' `
                 -Name 'ContentDeliveryAllowedOverride' `
                 -Description 'Re-enable Content Delivery Manager pre-installs'

$task = '\Microsoft\Windows\WindowsUpdate\Scheduled Start'
try {
    Enable-ScheduledTask -TaskName $task -ErrorAction Stop | Out-Null
    Write-Log "Enabled scheduled task: $task" -Level SUCCESS
}
catch {
    Write-Log "Could not enable $task : $($_.Exception.Message)" -Level WARN
}

$stopwatch.Stop()
Write-Log "ResetAutoUpdateSettings completed in $($stopwatch.Elapsed)" -Level SUCCESS
exit 0
