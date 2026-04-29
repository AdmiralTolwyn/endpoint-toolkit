<#
.SYNOPSIS
    De-provisions one or more inbox AppX packages from the image so they are not
    installed for new users created after capture.

.DESCRIPTION
    For each entry in -AppxPackages this script:
      1. Removes the matching provisioned package via Remove-AppxProvisionedPackage
         (so future users do not get it).
      2. Removes any existing per-user installs via Remove-AppxPackage -AllUsers.
      3. Removes the per-user install for the current SYSTEM context (defensive;
         normally a no-op).
      4. Special-case: when 'Microsoft.MSPaint' is requested, also removes the
         Microsoft.Windows.MSPaint Windows Capability (the MS Paint legacy FOD).

    Match is wildcard (*Name*), so a single entry like 'Bing' will sweep
    Microsoft.BingNews, Microsoft.BingWeather, etc.

.PARAMETER AppxPackages
    One or more AppX package name fragments to remove. Each is matched as *Name*
    against PackageName (provisioned) and Name (installed).

.NOTES
    File:    avd/customizer/RemoveAppxPackages.ps1
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
    .\RemoveAppxPackages.ps1 -AppxPackages 'Microsoft.BingNews','Microsoft.BingWeather','Microsoft.MSPaint'
#>

#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string[]]$AppxPackages
)

$ErrorActionPreference = 'Stop'

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO'
    )
    $ts    = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
    $color = switch ($Level) { 'WARN' {'Yellow'} 'ERROR' {'Red'} 'SUCCESS' {'Green'} default {'Gray'} }
    Write-Host "[$ts] [$Level] [RemoveAppxPackages] $Message" -ForegroundColor $color
}

function Remove-ProvidedAppxPackage {
<#
.SYNOPSIS
    Removes a single inbox AppX package (provisioned + per-user installs) by wildcard match.
.DESCRIPTION
    Helper used by the main loop. Errors are swallowed per-step so removing a missing
    package or one that's already gone never aborts the whole customizer run.
.PARAMETER AppName
    Wildcard fragment matched against PackageName / Name as *AppName*.
#>
    param([Parameter(Mandatory)][string]$AppName)

    try {
        Write-Log "Removing provisioned package: *$AppName*"
        Get-AppxProvisionedPackage -Online |
            Where-Object { $_.PackageName -like ("*{0}*" -f $AppName) } |
            Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Out-Null

        Write-Log "Removing per-user (-AllUsers) installs: *$AppName*"
        Get-AppxPackage -AllUsers -Name ("*{0}*" -f $AppName) |
            Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue

        Write-Log "Removing current-context installs: *$AppName*"
        Get-AppxPackage -Name ("*{0}*" -f $AppName) |
            Remove-AppxPackage -ErrorAction SilentlyContinue | Out-Null

        if ($AppName -eq 'Microsoft.MSPaint') {
            Write-Log "Special-case: removing Microsoft.Windows.MSPaint Windows Capability"
            Get-WindowsCapability -Online -Name '*Microsoft.Windows.MSPaint*' |
                Remove-WindowsCapability -Online -ErrorAction SilentlyContinue | Out-Null
        }
    }
    catch {
        Write-Log "Failed to remove '$AppName': $($_.Exception.Message)" -Level WARN
    }
}

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
Write-Log "Starting RemoveAppxPackages customizer phase ($($AppxPackages.Count) target(s))" -Level SUCCESS

foreach ($app in $AppxPackages) {
    Remove-ProvidedAppxPackage -AppName $app
}

$stopwatch.Stop()
Write-Log "RemoveAppxPackages completed in $($stopwatch.Elapsed)" -Level SUCCESS
exit 0
