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

.PARAMETER IncludeUserTemp
    Clean per-user AppData\Local\Temp folders (items older than -MaxAgeDays).

.PARAMETER IncludeWindowsTemp
    Clean %WinDir%\Temp (items older than -MaxAgeDays).

.PARAMETER IncludeSoftwareDistribution
    Stop wuauserv, purge C:\Windows\SoftwareDistribution, restart wuauserv.
    NOT recommended for routine cleanup. This is a Windows Update RESET:
      - Forces a full WU metadata re-sync on the next scan (catalog re-download).
      - Discards any partially staged update payloads (next WU pass re-downloads
        multi-GB ESDs / cumulative updates).
      - Clears Settings -> Update history (cosmetic, but visible to users / audit).
      - On WSUS-managed devices, breaks reporting until the client re-handshakes.
    Use only when WU is broken (0x8024xxxx, stuck scans, corrupt BITS queue),
    when prepping a sysprep'd reference image, or when the disk is critically
    full and cleanmgr + DISM cleanup did not free enough space.

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

.PARAMETER ExcludeHandler
    One or more cleanmgr VolumeCaches handler names to skip (case-insensitive,
    exact match). Useful to opt out of an aggressive handler for a one-off run
    without editing the script (e.g. 'DownloadsFolder','Previous Installations').
    Names must match the subkeys under
    HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches.

.PARAMETER IncludeOnlyHandler
    Restrict cleanmgr to only the named handler(s) (case-insensitive, exact match).
    When supplied, every handler NOT listed here is skipped, even before -ExcludeHandler
    is applied. Use this when you want a surgical run (e.g. only 'Update Cleanup' +
    'Windows Error Reporting Files').

.PARAMETER LogDirectory
    Directory for the log file. Default: $env:TEMP.

.NOTES
    File:    windows/servicing/Invoke-PreUpgradeCleanup.ps1
    Author:  Anton Romanyuk
    Version: 1.2.0
    Requires: PowerShell 5.1+, elevated session.

    Changes:
      1.2.0 - Renamed switches to MS-idiomatic full-word names:
                -UserTmp              -> -IncludeUserTemp
                -WindowsTmp           -> -IncludeWindowsTemp
                -SoftwareDistribution -> -IncludeSoftwareDistribution
                -LogDir               -> -LogDirectory
              BREAKING CHANGE: callers using the old names must update.
      1.1.0 - Added -ExcludeHandler and -IncludeOnlyHandler parameters.
              Expanded $VolumeCaches with handlers seen on modern Windows 11/Server
              (D3D Shader Cache, Diagnostic Data Viewer database files, DownloadsFolder,
              Feedback Hub Archive log files, Language Pack).
      1.0.0 - Initial release. Refactor of legacy CleanupBeforeUpgrade.ps1.

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
    .\Invoke-PreUpgradeCleanup.ps1 -IncludeUserTemp -IncludeWindowsTemp -IncludeSoftwareDistribution

.EXAMPLE
    # Force-run on a reference VM regardless of free space
    .\Invoke-PreUpgradeCleanup.ps1 -IncludeUserTemp -IncludeWindowsTemp -IncludeSoftwareDistribution -Force -IgnoreBattery

.EXAMPLE
    # Skip the Downloads folder and the Previous Installations rollback
    .\Invoke-PreUpgradeCleanup.ps1 -IncludeUserTemp -IncludeWindowsTemp -ExcludeHandler 'DownloadsFolder','Previous Installations'

.EXAMPLE
    # Surgical run: only Update Cleanup + WER
    .\Invoke-PreUpgradeCleanup.ps1 -IncludeOnlyHandler 'Update Cleanup','Windows Error Reporting Files' -Force
#>

#Requires -RunAsAdministrator
[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$IncludeUserTemp,
    [switch]$IncludeWindowsTemp,
    [switch]$IncludeSoftwareDistribution,
    [switch]$Force,
    [switch]$IgnoreBattery,
    [ValidateRange(1, 1024)] [int]$MinFreeGB  = 20,
    [ValidateRange(0,  365)] [int]$MaxAgeDays = 7,
    [ValidateRange(1, 9999)] [int]$SageId     = 5432,
    [string[]]$ExcludeHandler     = @(),
    [string[]]$IncludeOnlyHandler = @(),
    [string]$LogDirectory = $env:TEMP
)

$ErrorActionPreference = 'Stop'

# -----------------------------------------------------------------------------
# LOGGING
# -----------------------------------------------------------------------------
$ScriptName = $MyInvocation.MyCommand.Name
$LogFile    = Join-Path $LogDirectory ("{0}_{1}.log" -f [IO.Path]::GetFileNameWithoutExtension($ScriptName), (Get-Date -Format 'yyyyMMdd_HHmmss'))

function Write-Log {
<#
.SYNOPSIS
    Writes a timestamped, level-tagged line to both the console and the log file.
.DESCRIPTION
    Uniform logger used by the rest of the script. Format on disk and on console:
        [yyyy-MM-dd HH:mm:ss] [LEVEL] message
    Console output is colour-coded by level. File writes use SilentlyContinue so a
    transient lock on the log file never aborts the cleanup pipeline.
.PARAMETER Message
    Free-form text to record.
.PARAMETER Level
    INFO | WARN | ERROR | SUCCESS. Default INFO.
#>
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
<#
.SYNOPSIS
    Returns the free space on $env:SystemDrive in GB, rounded to two decimals.
.DESCRIPTION
    Used both as the free-space gate before cleanup and for the before/after
    delta reported in the summary. Reads via CIM (Win32_LogicalDisk).
.OUTPUTS
    [double] free space in gigabytes.
#>
    $drive = $env:SystemDrive.TrimEnd(':')
    $vol   = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='${drive}:'"
    [math]::Round($vol.FreeSpace / 1GB, 2)
}

function Test-OnBattery {
<#
.SYNOPSIS
    Returns $true when the device is currently running on battery power.
.DESCRIPTION
    Used to bail out before kicking off a long, IO-heavy cleanup on a laptop
    that might lose power mid-run. Returns $false on devices with no battery
    present (desktops, VMs, servers). The bail-out can be overridden with
    -IgnoreBattery on the script.

    Match heuristic: at least one Win32_Battery instance reports BatteryStatus
    = 1 (Discharging).
.OUTPUTS
    [bool]
#>
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
<#
.SYNOPSIS
    Recursively deletes items under $Path whose CreationTime is older than $OlderThanDays.
.DESCRIPTION
    Resilient deletion helper used for %TEMP% and %WinDir%\Temp purges:
      * Skips when the parent folder does not exist (returns silently).
      * Uses CreationTime (not LastWriteTime) so freshly written files in old
        folders are still considered young.
      * Honours -WhatIf / -Confirm via $PSCmdlet.ShouldProcess.
      * Swallows individual file errors (file in use, ACL denial) so a single
        failure cannot abort the whole pre-upgrade workflow.
.PARAMETER Path
    Wildcard or fully-qualified path to clean (e.g. C:\Users\*\AppData\Local\Temp\*).
.PARAMETER OlderThanDays
    Minimum age (days) before an item is eligible for deletion.
#>
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
# Handlers not present on the running SKU are silently skipped at write-time.
$VolumeCaches = @(
    'Active Setup Temp Folders'
    'BranchCache'
    'Content Indexer Cleaner'
    'D3D Shader Cache'
    'Delivery Optimization Files'
    'Device Driver Packages'
    'Diagnostic Data Viewer database files'
    'Downloaded Program Files'
    'DownloadsFolder'
    'Feedback Hub Archive log files'
    'Internet Cache Files'
    'Language Pack'
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

function Resolve-EffectiveHandlers {
<#
.SYNOPSIS
    Filters the master cleanmgr VolumeCaches list using -IncludeOnlyHandler and -ExcludeHandler.
.DESCRIPTION
    Apply -IncludeOnlyHandler and -ExcludeHandler to the master list.
    Matching is case-insensitive; unknown names produce a warning so a typo
    is visible instead of silently selecting nothing.

    Order of operations:
      1. If -IncludeOnly is supplied, restrict the set to those names.
      2. Then drop anything listed in -Exclude.
.PARAMETER All
    Master handler list (typically $VolumeCaches).
.PARAMETER IncludeOnly
    When non-empty, only these handler names survive step 1.
.PARAMETER Exclude
    Handler names to drop after step 1.
.OUTPUTS
    [string[]] effective handler set (always returned as an array, never $null).
#>
    param(
        [Parameter(Mandatory)][string[]]$All,
        [string[]]$IncludeOnly = @(),
        [string[]]$Exclude     = @()
    )
    $set = [System.Collections.Generic.HashSet[string]]::new([string[]]$All, [System.StringComparer]::OrdinalIgnoreCase)

    if ($IncludeOnly.Count -gt 0) {
        foreach ($name in $IncludeOnly) {
            if (-not $set.Contains($name)) {
                Write-Log "-IncludeOnlyHandler '$name' is not in the known handler list (typo?)." -Level WARN
            }
        }
        $effective = $All | Where-Object { $IncludeOnly -contains $_ }
    }
    else {
        $effective = $All
    }

    if ($Exclude.Count -gt 0) {
        foreach ($name in $Exclude) {
            if (-not $set.Contains($name)) {
                Write-Log "-ExcludeHandler '$name' is not in the known handler list (typo?)." -Level WARN
            }
        }
        $effective = $effective | Where-Object { $Exclude -notcontains $_ }
    }

    return ,@($effective)
}

function Set-CleanMgrSageRun {
<#
.SYNOPSIS
    Configures cleanmgr.exe StateFlags<SageId> for the supplied handler list.
.DESCRIPTION
    Writes the DWORD value StateFlagsNNNN = 2 under each
    HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\<handler>
    subkey so that `cleanmgr.exe /sagerun:NNNN` will process exactly those
    handlers. Handler subkeys that do not exist on the running SKU are skipped
    silently (legitimate - not all handlers ship on every Windows version).
.PARAMETER SageId
    The numeric StateFlags slot id (1..9999) used in /sagerun.
.PARAMETER Caches
    Handler subkey names to enable.
#>
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
<#
.SYNOPSIS
    Runs an external executable synchronously and logs its exit code.
.DESCRIPTION
    Thin Start-Process wrapper used for cleanmgr.exe and dism.exe so the call
    sites read top-to-bottom and every invocation gets the same logging shape
    (full command line in, exit code out).
.PARAMETER FilePath
    Path or name of the executable.
.PARAMETER ArgumentList
    Arguments forwarded to the process.
.OUTPUTS
    [int] process exit code.
#>
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

    if ($IncludeUserTemp) {
        Remove-OldItems -Path 'C:\Users\*\AppData\Local\Temp\*' -OlderThanDays $MaxAgeDays
    }

    if ($IncludeWindowsTemp) {
        Remove-OldItems -Path (Join-Path $env:WinDir 'Temp\*') -OlderThanDays $MaxAgeDays
    }

    $effectiveHandlers = Resolve-EffectiveHandlers `
        -All         $VolumeCaches `
        -IncludeOnly $IncludeOnlyHandler `
        -Exclude     $ExcludeHandler

    if ($effectiveHandlers.Count -eq 0) {
        Write-Log "No cleanmgr handlers selected after applying -IncludeOnlyHandler/-ExcludeHandler. Skipping cleanmgr." -Level WARN
    }
    else {
        Write-Log ("Cleanmgr handlers selected: {0}/{1}" -f $effectiveHandlers.Count, $VolumeCaches.Count)
        Set-CleanMgrSageRun -SageId $SageId -Caches $effectiveHandlers
        [void](Invoke-Tool -FilePath 'cleanmgr.exe' -ArgumentList @("/sagerun:$SageId"))
    }

    if ($IncludeSoftwareDistribution) {
        Write-Log "Purging SoftwareDistribution cache (WU RESET: forces full metadata re-sync on next scan)" -Level WARN
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
