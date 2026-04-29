<#
.SYNOPSIS
    Patches the Azure VM agent's deprovisioning script so Sysprep is invoked in /mode:vm
    instead of the default /quiet /quit. Required for AVD / multi-session image bake.

.DESCRIPTION
    The Azure Linux Agent for Windows ships a default C:\DeprovisioningScript.ps1 that
    runs Sysprep with `/oobe /generalize /quiet /quit`. For AVD / Windows 365 image
    capture this needs to be `/oobe /generalize /quit /mode:vm` so the resulting image
    can be deployed to a virtual machine without re-arming the OOBE flow.

    This script does an in-place text replacement on the deprovisioning script on the
    capture VM. Designed to be wired into Azure Image Builder (AIB) / Packer as a
    PowerShell customizer step that runs late in the image bake, just before capture.

.NOTES
    File:    avd/customizer/AdminSysPrep.ps1
    Author:  Anton Romanyuk
    Version: 2.0.0
    Context: Azure Image Builder / Packer customizer. Runs as SYSTEM.
    Requires: Windows 10/11 multi-session or Windows Server, PowerShell 5.1+, admin.

.DISCLAIMER
    This script is provided "AS IS" with no warranties and confers no rights.
    It is not supported under any Microsoft standard support program or service.
    Use of this script is entirely at your own risk. The customer is solely
    responsible for testing and validating this script in their environment
    before deploying to production.

.EXAMPLE
    .\AdminSysPrep.ps1
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
    Write-Host "[$ts] [$Level] [AdminSysPrep] $Message" -ForegroundColor $color
}

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
Write-Log "Starting AdminSysPrep customizer phase" -Level SUCCESS

$DeprovisionScript = 'C:\DeprovisioningScript.ps1'
$Find    = 'Sysprep.exe /oobe /generalize /quiet /quit'
$Replace = 'Sysprep.exe /oobe /generalize /quit /mode:vm'

if (-not (Test-Path -LiteralPath $DeprovisionScript)) {
    Write-Log "Deprovisioning script not found at $DeprovisionScript - skipping (non-Azure VM?)" -Level WARN
    exit 0
}

try {
    $content = Get-Content -LiteralPath $DeprovisionScript -Raw
    if ($content -notlike "*$Find*") {
        Write-Log "Sysprep line already patched (or signature changed). No action taken." -Level WARN
        exit 0
    }
    ($content -replace [regex]::Escape($Find), $Replace) |
        Set-Content -LiteralPath $DeprovisionScript -Encoding UTF8
    Write-Log "Patched $DeprovisionScript -> /mode:vm" -Level SUCCESS
}
catch {
    Write-Log "Failed to patch deprovisioning script: $($_.Exception.Message)" -Level ERROR
    exit 1
}

$stopwatch.Stop()
Write-Log "AdminSysPrep completed in $($stopwatch.Elapsed)" -Level SUCCESS
exit 0
