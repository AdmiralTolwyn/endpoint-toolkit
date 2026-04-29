<#
.SYNOPSIS
    Disables Microsoft Store auto-updates and the Windows Update scheduled-start trigger
    so MSIX app attach packages and the base image are not silently updated mid-bake or
    after capture.

.DESCRIPTION
    Image-bake hardening for AVD / Windows 365 reference VMs:
      1. HKLM\Software\Policies\Microsoft\WindowsStore\AutoDownload = 2 (off)
         Prevents the Microsoft Store from auto-downloading updates that could rev MSIX
         app attach packages or stub apps after the image is captured.
      2. HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager\Debug
         \ContentDeliveryAllowedOverride = 2 (off)
         Disables the Content Delivery Manager pre-install path that re-downloads
         Candy Crush / Spotify-style consumer apps.
      3. Disables \Microsoft\Windows\WindowsUpdate\Scheduled Start so WU does not kick
         off mid-customizer.

    NOTE: The original legacy script also wrote to HKCU - that is meaningless when run
    from the AIB SYSTEM context (HKCU = SYSTEM's hive), so it has been removed.

.NOTES
    File:    avd/customizer/DisableAutoUpdates.ps1
    Author:  Anton Romanyuk
    Version: 2.0.0
    Context: Azure Image Builder / Packer customizer. Runs as SYSTEM.
    Requires: Windows 10/11 / Server, PowerShell 5.1+, admin.

.DISCLAIMER
    This script is provided "AS IS" with no warranties and confers no rights.
    It is not supported under any Microsoft standard support program or service.
    Use of this script is entirely at your own risk. The customer is solely
    responsible for testing and validating this script in their environment
    before deploying to production.

.EXAMPLE
    .\DisableAutoUpdates.ps1
#>

#Requires -RunAsAdministrator
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

function Write-Log {
<#
.SYNOPSIS
    Console-only logger; format [timestamp] [LEVEL] [DisableAutoUpdates] message.
.PARAMETER Message
    Free-form text.
.PARAMETER Level
    INFO | WARN | ERROR | SUCCESS. Default INFO.
#>
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO'
    )
    $ts    = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
    $color = switch ($Level) { 'WARN' {'Yellow'} 'ERROR' {'Red'} 'SUCCESS' {'Green'} default {'Gray'} }
    Write-Host "[$ts] [$Level] [DisableAutoUpdates] $Message" -ForegroundColor $color
}

function Set-RegDword {
<#
.SYNOPSIS
    Idempotently sets a DWORD registry value, creating the parent key tree if absent.
.PARAMETER Path
    Full registry path (e.g. HKLM:\SOFTWARE\...).
.PARAMETER Name
    Value name to write.
.PARAMETER Value
    DWORD value to set.
#>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][int]   $Value
    )
    try {
        if (-not (Test-Path -LiteralPath $Path)) {
            New-Item -Path $Path -Force | Out-Null
        }
        New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType DWord -Force | Out-Null
        Write-Log "Set $Path\$Name = $Value" -Level SUCCESS
    }
    catch {
        Write-Log "Failed to set $Path\$Name : $($_.Exception.Message)" -Level ERROR
        throw
    }
}

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
Write-Log "Starting DisableAutoUpdates customizer phase" -Level SUCCESS

try {
    Set-RegDword -Path 'HKLM:\Software\Policies\Microsoft\WindowsStore'                                       -Name 'AutoDownload'                  -Value 2
    Set-RegDword -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager\Debug'         -Name 'ContentDeliveryAllowedOverride' -Value 2

    $task = '\Microsoft\Windows\WindowsUpdate\Scheduled Start'
    try {
        Disable-ScheduledTask -TaskPath '\Microsoft\Windows\WindowsUpdate\' -TaskName 'Scheduled Start' -ErrorAction Stop | Out-Null
        Write-Log "Disabled scheduled task: $task" -Level SUCCESS
    }
    catch {
        Write-Log "Could not disable $task : $($_.Exception.Message)" -Level WARN
    }
}
catch {
    Write-Log "Aborting DisableAutoUpdates: $($_.Exception.Message)" -Level ERROR
    exit 1
}

$stopwatch.Stop()
Write-Log "DisableAutoUpdates completed in $($stopwatch.Elapsed)" -Level SUCCESS
exit 0
