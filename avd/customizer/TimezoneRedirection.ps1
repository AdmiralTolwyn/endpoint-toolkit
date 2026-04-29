<#
.SYNOPSIS
    Enables remote-desktop time-zone redirection on the AVD / Windows 365 image so the
    session host inherits the client's local time zone instead of the host VM's.

.DESCRIPTION
    Sets the per-machine policy
        HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services
        \fEnableTimeZoneRedirection = 1
    which is the documented switch for "Allow time zone redirection" in the
    Computer Configuration\Administrative Templates\Windows Components\Remote Desktop
    Services\Remote Desktop Session Host\Device and Resource Redirection GPO node.

    Reference:
      https://learn.microsoft.com/azure/virtual-desktop/configure-device-redirections
      https://learn.microsoft.com/windows-server/remote/remote-desktop-services/clients/rdp-files

.NOTES
    File:    avd/customizer/TimezoneRedirection.ps1
    Author:  Anton Romanyuk
    Version: 2.0.0
    Context: Azure Image Builder / Packer customizer. Runs as SYSTEM.
    Requires: Windows 10/11 multi-session or Server, PowerShell 5.1+, admin.

.DISCLAIMER
    This script is provided "AS IS" with no warranties and confers no rights.
    It is not supported under any Microsoft standard support program or service.
    Use of this script is entirely at your own risk. The customer is solely
    responsible for testing and validating this script in their environment
    before deploying to production.

.EXAMPLE
    .\TimezoneRedirection.ps1
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
    Write-Host "[$ts] [$Level] [TimezoneRedirection] $Message" -ForegroundColor $color
}

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
Write-Log "Starting TimezoneRedirection customizer phase" -Level SUCCESS

$Path  = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services'
$Name  = 'fEnableTimeZoneRedirection'
$Value = 1

try {
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }
    New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType DWord -Force | Out-Null
    Write-Log "Set $Path\$Name = $Value" -Level SUCCESS
}
catch {
    Write-Log "Failed to enable time zone redirection: $($_.Exception.Message)" -Level ERROR
    exit 1
}

$stopwatch.Stop()
Write-Log "TimezoneRedirection completed in $($stopwatch.Elapsed)" -Level SUCCESS
exit 0
