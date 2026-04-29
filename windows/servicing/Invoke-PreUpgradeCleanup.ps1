<#
.SYNOPSIS
    Frees disk space prior to a Windows feature update or after a reference image build.

.DESCRIPTION
    Performs a configurable cleanup pass to reclaim space on the system drive before
    a Windows feature update (or after sysprep / reference image creation):

      1. Free-space gate (skips work if SystemDrive has more than -MinFreeGB free,
         unless -Force is supplied).
      2. Optional purge of per-user %TEMP% folders older than -MaxAgeDays.
      3. Optional purge of %WinDir%\Temp older than -MaxAgeDays.
      4. Optional purge of the SoftwareDistribution download cache (stops/starts wuauserv).
      5. Configures all known cleanmgr.exe VolumeCaches (StateFlags<SageId>) and runs
         `cleanmgr.exe /sagerun:<SageId>`.
      6. Runs `dism.exe /Online /Cleanup-Image /StartComponentCleanup /ResetBase`
         to shrink WinSxS (irreversible — uninstalls superseded updates).
      7. Re-checks free space and aborts with 1602 if still below threshold.
      8. Refuses to proceed on battery-only power (exit 1603) unless -IgnoreBattery.

.PARAMETER UserTmp
    Clean per-user AppData\Local\Temp folders (items older than -MaxAgeDays).

.PARAMETER WindowsTmp
    Clean %WinDir%\Temp (items older than -MaxAgeDays).

.PARAMETER SoftwareDistribution
    Stop wuauserv, purge C:\Windows\SoftwareDistribution, restart wuauserv.

.PARAMETER Force
    Run cleanup regardless of current free space.

.PARAMETER IgnoreBattery
    Do not abort when running on battery power.

.PARAMETER MinFreeGB
    Free-space threshold (GB) on SystemDrive. Default: 20.

.PARAMETER MaxAgeDays
    Minimum age (days) for files in temp folders before they are deleted. Default: 7.

.PARAMETER SageId
    Cleanmgr StateFlags slot id (1..9999). Default: 5432.

.PARAMETER LogDir
    Directory for the log file. Default: $env:TEMP.

.NOTES
    File:    windows/servicing/Invoke-PreUpgradeCleanup.ps1
    Author:  Anton Romanyuk
    Version: 1.0.0
    Requires: PowerShell 5.1+, elevated session.

    Exit codes:
      0    - Success / nothing to do
      1602 - Free space still below -MinFreeGB after cleanup
      1603 - Running on battery power (use -IgnoreBattery to override)

.DISCLAIMER
    This script is provided "AS IS" with no warranties and confers no rights.
    It is not supported under any Microsoft standard support program or service.
    Use of this script is entirely at your own risk. The customer is solely
    responsible for testing and validating this script in their environment
    before deploying to production. `dism /ResetBase` PERMANENTLY removes the
    ability to uninstall previously installed Windows updates.

.EXAMPLE
    # Full cleanup before a feature update
    .\Invoke-PreUpgradeCleanup.ps1 -UserTmp -WindowsTmp -SoftwareDistribution

.EXAMPLE
    # Force-run on a reference VM regardless of free space
    .\Invoke-PreUpgradeCleanup.ps1 -UserTmp -WindowsTmp -SoftwareDistribution -Force -IgnoreBattery
#>

#Requires -RunAsAdministrator
[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$UserTmp,
    [switch]$WindowsTmp,
    [switch]$SoftwareDistribution,
    [switch]$Force,
    [switch]$IgnoreBattery,
    [ValidateRange(1, 1024)] [int]$MinFreeGB  = 20,
    [ValidateRange(0,  365)] [int]$MaxAgeDays = 7,
    [ValidateRange(1, 9999)] [int]$SageId     = 5432,
    [string]$LogDir = $env:TEMP
)

$ErrorActionPreference = 'Stop'

# -----------------------------------------------------------------------------
# LOGGING
# -----------------------------------------------------------------------------
$ScriptName = $MyInvocation.MyCommand.Name
$LogFile    = Join-Path $LogDir ("{0}_{1}.log" -f [IO.Path]::GetFileNameWithoutExtension($ScriptName), (Get-Date -Format 'yyyyMMdd_HHmmss'))

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO'
    )
    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -LiteralPath $LogFile -Value $line -ErrorAction SilentlyContinue
    $color = switch ($Level) {
        'WARN'    { 'Yellow' }
        'ERROR'   { 'Red' }
        'SUCCESS' { 'Green' }
        default   { 'Gray' }
    }
    Write-Host $line -ForegroundColor $color
}

# -----------------------------------------------------------------------------
# HELPERS
# -----------------------------------------------------------------------------
function Get-SystemDriveFreeGB {
    $drive = $env:SystemDrive.TrimEnd(':')
    $vol   = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='${drive}:'"
    [math]::Round($vol.FreeSpace / 1GB, 2)
}

function Test-OnBattery {
    # Returns $true only if at least one battery is actively discharging AND no AC is reported.
    $batteries = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
    if (-not $batteries) { return $false }   # Desktop / VM — no battery present
    foreach ($b in $batteries) {
        # BatteryStatus 1 = Discharging
        if ($b.BatteryStatus -eq 1) { return $true }
    }
    return $false
}

function Remove-OldItems {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][int]$OlderThanDays
    )
    if (-not (Test-Path -LiteralPath (Split-Path -Path $Path -Parent) -ErrorAction SilentlyContinue)) {
        Write-Log "Skip (parent not found): $Path" -Level WARN
        return
    }
    $cutoff = (Get-Date).AddDays(-$OlderThanDays)
    Write-Log "Cleaning '$Path' (items older than $OlderThanDays days, cutoff $cutoff)"
    try {
        Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.CreationTime -lt $cutoff } |
            ForEach-Object {
                if ($PSCmdlet.ShouldProcess($_.FullName, 'Remove')) {
                    Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
    }
    catch {
        Write-Log "Cleanup of '$Path' raised: $($_.Exception.Message)" -Level WARN
    }
}

# Known cleanmgr.exe VolumeCaches handlers. Keys are subkey names under
# HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches
$VolumeCaches = @(
    'Active Setup Temp Folders'
    'BranchCache'
    'Content Indexer Cleaner'
    'Delivery Optimization Files'
    'Device Driver Packages'
    'Downloaded Program Files'
    'Internet Cache Files'
    'Memory Dump Files'
    'Offline Pages Files'
    'Old ChkDsk Files'
    'Previous Installations'
    'Recycle Bin'
    'RetailDemo Offline Content'
    'Service Pack Cleanup'
    'Setup Log Files'
    'System error memory dump files'
    'System error minidump files'
    'Temporary Files'
    'Temporary Setup Files'
    'Temporary Sync Files'
    'Thumbnail Cache'
    'Update Cleanup'
    'Upgrade Discarded Files'
    'User file versions'
    'Windows Defender'
    'Windows Error Reporting Archive Files'
    'Windows Error Reporting Queue Files'
    'Windows Error Reporting System Archive Files'
    'Windows Error Reporting System Queue Files'
    'Windows Error Reporting Files'
    'Windows ESD installation files'
    'Windows Reset Log Files'
    'Windows Upgrade Log Files'
)

function Set-CleanMgrSageRun {
    param(
        [Parameter(Mandatory)][int]$SageId,
        [Parameter(Mandatory)][string[]]$Caches
    )
    $valueName = "StateFlags{0:D4}" -f $SageId
    $rootPath  = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches'
    Write-Log "Configuring cleanmgr handlers ($valueName)"
    foreach ($cache in $Caches) {
        $keyPath = Join-Path $rootPath $cache
        try {
            if (-not (Test-Path -LiteralPath $keyPath)) {
                # Skip — not all handlers exist on every Windows SKU
                continue
            }
            New-ItemProperty -Path $keyPath -Name $valueName -Value 2 -PropertyType DWord -Force | Out-Null
        }
        catch {
            Write-Log "Failed to set $valueName on '$cache': $($_.Exception.Message)" -Level WARN
        }
    }
}

function Invoke-Tool {
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [string[]]$ArgumentList = @()
    )
    Write-Log ("Running: {0} {1}" -f $FilePath, ($ArgumentList -join ' '))
    $proc = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -Wait -NoNewWindow -PassThru
    Write-Log "$FilePath exited with code $($proc.ExitCode)"
    return $proc.ExitCode
}

# -----------------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------------
Write-Log "=== $ScriptName starting ===" -Level SUCCESS
Write-Log "Log file: $LogFile"

$freeBefore = Get-SystemDriveFreeGB
Write-Log "Free space on $($env:SystemDrive): $freeBefore GB (threshold: $MinFreeGB GB)"

if ($freeBefore -gt $MinFreeGB -and -not $Force) {
    Write-Log "Free space above threshold and -Force not set. Skipping cleanup." -Level SUCCESS
}
else {
    if ($Force) { Write-Log "-Force specified — running cleanup unconditionally." -Level WARN }

    if ($UserTmp) {
        Remove-OldItems -Path 'C:\Users\*\AppData\Local\Temp\*' -OlderThanDays $MaxAgeDays
    }

    if ($WindowsTmp) {
        Remove-OldItems -Path (Join-Path $env:WinDir 'Temp\*') -OlderThanDays $MaxAgeDays
    }

    Set-CleanMgrSageRun -SageId $SageId -Caches $VolumeCaches
    [void](Invoke-Tool -FilePath 'cleanmgr.exe' -ArgumentList @("/sagerun:$SageId"))

    if ($SoftwareDistribution) {
        Write-Log "Purging SoftwareDistribution cache"
        [void](Invoke-Tool -FilePath 'net.exe' -ArgumentList @('stop','wuauserv'))
        try {
            $sdPath = Join-Path $env:WinDir 'SoftwareDistribution'
            if (Test-Path -LiteralPath $sdPath) {
                Get-ChildItem -LiteralPath $sdPath -Recurse -Force -ErrorAction SilentlyContinue |
                    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
        finally {
            [void](Invoke-Tool -FilePath 'net.exe' -ArgumentList @('start','wuauserv'))
        }
    }

    [void](Invoke-Tool -FilePath 'dism.exe' `
        -ArgumentList @('/Online','/Cleanup-Image','/StartComponentCleanup','/ResetBase'))
}

# -----------------------------------------------------------------------------
# POST-CHECKS
# -----------------------------------------------------------------------------
$freeAfter = Get-SystemDriveFreeGB
Write-Log "Free space after cleanup: $freeAfter GB (delta: $([math]::Round($freeAfter - $freeBefore, 2)) GB)" -Level SUCCESS

if ($freeAfter -lt $MinFreeGB) {
    Write-Log "Free space ($freeAfter GB) still below required $MinFreeGB GB. Exit 1602." -Level ERROR
    exit 1602
}

if (-not $IgnoreBattery -and (Test-OnBattery)) {
    Write-Log "Device is on battery power and discharging. Exit 1603." -Level ERROR
    exit 1603
}

Write-Log "=== $ScriptName completed ===" -Level SUCCESS
exit 0
