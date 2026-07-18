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

.PARAMETER ContinueOnError
    Do not exit non-zero when one or more package/capability removals fail. By
    default the script exits 1 if any individual removal fails, since silently
    leaving a target un-removed can surface later in the captured image. Pass
    this switch to fall back to best-effort behaviour (always exit 0).

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

.EXAMPLE
    # Best-effort: do not fail the build even if some removals fail
    .\RemoveAppxPackages.ps1 -AppxPackages 'Microsoft.BingNews' -ContinueOnError
#>

#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string[]]$AppxPackages,

    [switch]$ContinueOnError
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
    Helper used by the main loop. Each matched provisioned package / per-user install /
    capability is removed in its own try/catch so a single locked or already-gone item
    never aborts the rest of the sweep for this -AppName (per-item isolation). Failures
    are counted and surfaced to the caller via the return value.
.PARAMETER AppName
    Wildcard fragment matched against PackageName / Name as *AppName*.
.OUTPUTS
    $true if every matched package/capability was removed successfully (or nothing
    matched for this fragment); $false if one or more removals failed.
#>
    param([Parameter(Mandatory)][string]$AppName)

    $failed = 0

    try {
        Write-Log "Removing provisioned package: *$AppName*"
        $provisionedMatches = @(Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
            Where-Object { $_.PackageName -like ("*{0}*" -f $AppName) })
        foreach ($pkg in $provisionedMatches) {
            try {
                Remove-AppxProvisionedPackage -Online -PackageName $pkg.PackageName -ErrorAction Stop | Out-Null
            }
            catch {
                $failed++
                Write-Log "Failed to remove provisioned package '$($pkg.PackageName)': $($_.Exception.Message)" -Level WARN
            }
        }

        Write-Log "Removing per-user (-AllUsers) installs: *$AppName*"
        $allUsersMatches = @(Get-AppxPackage -AllUsers -Name ("*{0}*" -f $AppName) -ErrorAction SilentlyContinue)
        foreach ($pkg in $allUsersMatches) {
            try {
                Remove-AppxPackage -Package $pkg.PackageFullName -AllUsers -ErrorAction Stop
            }
            catch {
                $failed++
                Write-Log "Failed to remove per-user package '$($pkg.PackageFullName)': $($_.Exception.Message)" -Level WARN
            }
        }

        Write-Log "Removing current-context installs: *$AppName*"
        $currentMatches = @(Get-AppxPackage -Name ("*{0}*" -f $AppName) -ErrorAction SilentlyContinue)
        foreach ($pkg in $currentMatches) {
            try {
                Remove-AppxPackage -Package $pkg.PackageFullName -ErrorAction Stop | Out-Null
            }
            catch {
                $failed++
                Write-Log "Failed to remove current-context package '$($pkg.PackageFullName)': $($_.Exception.Message)" -Level WARN
            }
        }

        if ($AppName -eq 'Microsoft.MSPaint') {
            Write-Log "Special-case: removing Microsoft.Windows.MSPaint Windows Capability"
            $capMatches = @(Get-WindowsCapability -Online -Name '*Microsoft.Windows.MSPaint*' -ErrorAction SilentlyContinue)
            foreach ($cap in $capMatches) {
                try {
                    Remove-WindowsCapability -Online -Name $cap.Name -ErrorAction Stop | Out-Null
                }
                catch {
                    $failed++
                    Write-Log "Failed to remove capability '$($cap.Name)': $($_.Exception.Message)" -Level WARN
                }
            }
        }
    }
    catch {
        # Only the Get-* enumeration calls above can land here (individual removals
        # already have their own try/catch); still per-target isolated from the caller's
        # perspective - the loop moves on to the next -AppName regardless.
        $failed++
        Write-Log "Failed to enumerate/remove '$AppName': $($_.Exception.Message)" -Level WARN
    }

    return ($failed -eq 0)
}

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
Write-Log "Starting RemoveAppxPackages customizer phase ($($AppxPackages.Count) target(s))" -Level SUCCESS

$failedTargetCount = 0
foreach ($app in $AppxPackages) {
    if (-not (Remove-ProvidedAppxPackage -AppName $app)) {
        $failedTargetCount++
    }
}

$stopwatch.Stop()

if ($failedTargetCount -gt 0) {
    Write-Log "SUMMARY: $failedTargetCount of $($AppxPackages.Count) target(s) had one or more removal failures." -Level ERROR
    if ($ContinueOnError) {
        Write-Log "-ContinueOnError specified - exiting 0 despite failure(s). Completed in $($stopwatch.Elapsed)" -Level WARN
        exit 0
    }
    Write-Log "RemoveAppxPackages failed after $($stopwatch.Elapsed). Use -ContinueOnError to treat failures as best-effort." -Level ERROR
    exit 1
}

Write-Log "RemoveAppxPackages completed in $($stopwatch.Elapsed)" -Level SUCCESS
exit 0
