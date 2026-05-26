<#
.SYNOPSIS
    Reports EFI System Partition (ESP) size and free space as a single line of JSON for
    log-scrape ingestion (Grafana / Loki / Promtail / Telegraf exec / Azure Monitor).

.DESCRIPTION
    Locates every ESP on the device (GPT type {c12a7328-f81f-11d2-ba4b-00a0c93ec93b}),
    correlates it with Win32_Volume via the partition's volume-GUID access path
    (\\?\Volume{...}\), and emits SizeMB / FreeMB / UsedMB / FreePct per ESP.

    Designed for the May 2026 servicing failure (KB5089549, error 0x800f0922) where
    devices with <= 10 MB free on the ESP fail at ~35-36 % during the post-reboot
    phase and roll back. CBS.log on affected devices shows:
        "SpaceCheck: Insufficient free space"
        "ServicingBootFiles failed. Error = 0x70"
        "SpaceCheck: <value> used by third-party/OEM files outside of Microsoft boot directories"

    No mountvol / drive-letter assignment is performed (ESPs are typically letterless
    and read-only). Sizes come from CIM (Win32_Volume.Capacity / FreeSpace) so the
    script works on any modern Windows 10 / 11 / Server build without elevation
    quirks beyond what Get-Partition itself requires.

.PARAMETER CriticalFreeMB
    Free-space (MB) threshold below which exit code 2 (CRITICAL) is returned.
    Default: 15. KB5089549 fails at <= 10 MB; default is set above that so devices
    in the danger zone exit critical even if they're not quite at 10 MB yet.

.PARAMETER WarningFreeMB
    Free-space (MB) threshold below which exit code 1 (WARNING) is returned (but
    above -CriticalFreeMB). Default: 30.

.PARAMETER OutputPath
    Optional. When supplied, the JSON line is ALSO appended to this file (UTF-8, no
    BOM) so Promtail / Filebeat / agent file-scrape configs can tail it. The script
    still writes the JSON line to stdout.

.PARAMETER Pretty
    Emit indented JSON instead of single-line. Useful for ad-hoc inspection; do not
    enable for Loki / Promtail scrape jobs (they expect one JSON object per line).

.EXAMPLE
    # One-shot run (stdout only)
    .\Get-EspPartitionStatus.ps1

.EXAMPLE
    # Telegraf exec input + Loki tail: write each run as a line into a log file
    .\Get-EspPartitionStatus.ps1 -OutputPath 'C:\ProgramData\EspMonitor\esp_status.log'

.EXAMPLE
    # Tighter alert: warn at 50 MB, critical at 20 MB
    .\Get-EspPartitionStatus.ps1 -WarningFreeMB 50 -CriticalFreeMB 20

.NOTES
    File:     windows/servicing/EspPartitionStatus/Get-EspPartitionStatus.ps1
    Author:   Anton Romanyuk
    Version:  1.0.0
    Requires: PowerShell 5.1+, Get-Partition (Storage module), CIM access.

    Exit codes:
      0 - OK       (free space > WarningFreeMB on every ESP)
      1 - WARNING  (any ESP between Critical and Warning thresholds)
      2 - CRITICAL (any ESP at or below CriticalFreeMB)
      3 - ESP not found / partition table not GPT
      4 - Unexpected error (caught at the top level)

    Related workaround (per KB5089549 advisory):
      reg add "HKLM\SYSTEM\CurrentControlSet\Control\Bfsvc" /v EspPaddingPercent /t REG_DWORD /d 0 /f
    Restart and retry the update. This script does NOT apply the workaround --
    it only reports the symptom so the dashboard can flag affected devices.

.DISCLAIMER
    THIS SCRIPT IS PROVIDED "AS-IS" WITHOUT WARRANTY OF ANY KIND. Test against a
    representative device before deploying as a scheduled task / Intune script.
#>

#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateRange(1, 1024)]
    [int]    $CriticalFreeMB = 15,

    [ValidateRange(1, 1024)]
    [int]    $WarningFreeMB  = 30,

    [string] $OutputPath,

    [switch] $Pretty
)

$ErrorActionPreference = 'Stop'

$ESP_GPT_TYPE = '{c12a7328-f81f-11d2-ba4b-00a0c93ec93b}'

# --- exit-code constants ---
$EXIT_OK         = 0
$EXIT_WARN       = 1
$EXIT_CRITICAL   = 2
$EXIT_NOTFOUND   = 3
$EXIT_ERROR      = 4

function Convert-ToMB {
    param([Parameter(Mandatory)][AllowNull()] $Bytes)
    if ($null -eq $Bytes) { return $null }
    return [Math]::Round([int64]$Bytes / 1MB, 2)
}

function Get-EspVolumeInfo {
    $partitions = @(Get-Partition -ErrorAction Stop | Where-Object { $_.GptType -eq $ESP_GPT_TYPE })
    if (-not $partitions) { return @() }

    # Pull the full Win32_Volume list once -- ~5 ms on a normal box -- then match by
    # the partition's \\?\Volume{guid}\ access path, which Win32_Volume reports as
    # its DeviceID.
    $allVolumes = Get-CimInstance -ClassName Win32_Volume -ErrorAction Stop

    foreach ($p in $partitions) {
        $volPath = $p.AccessPaths | Where-Object { $_ -like '\\?\Volume*' } | Select-Object -First 1
        $vol     = $null
        if ($volPath) {
            $vol = $allVolumes | Where-Object { $_.DeviceID -eq $volPath } | Select-Object -First 1
        }

        $sizeBytes = if ($vol) { [int64]$vol.Capacity }  else { [int64]$p.Size }
        $freeBytes = if ($vol) { [int64]$vol.FreeSpace } else { $null }

        [pscustomobject]@{
            DiskNumber      = [int]$p.DiskNumber
            PartitionNumber = [int]$p.PartitionNumber
            VolumeGuidPath  = $volPath
            FileSystem      = if ($vol) { $vol.FileSystem } else { $null }
            Label           = if ($vol) { $vol.Label } else { $null }
            SizeMB          = Convert-ToMB $sizeBytes
            FreeMB          = Convert-ToMB $freeBytes
            UsedMB          = if ($null -ne $freeBytes) { Convert-ToMB ($sizeBytes - $freeBytes) } else { $null }
            FreePct         = if ($null -ne $freeBytes -and $sizeBytes -gt 0) {
                                  [Math]::Round(($freeBytes / $sizeBytes) * 100, 2)
                              } else { $null }
        }
    }
}

function Get-StatusLabel {
    param(
        [Parameter(Mandatory)][AllowNull()] $FreeMB,
        [Parameter(Mandatory)][int] $WarnMB,
        [Parameter(Mandatory)][int] $CritMB
    )
    if ($null -eq $FreeMB)        { return 'UNKNOWN' }
    if ($FreeMB -le $CritMB)      { return 'CRITICAL' }
    if ($FreeMB -le $WarnMB)      { return 'WARNING' }
    return 'OK'
}

try {
    $esps = Get-EspVolumeInfo

    $partitionPayload = foreach ($e in $esps) {
        $status = Get-StatusLabel -FreeMB $e.FreeMB -WarnMB $WarningFreeMB -CritMB $CriticalFreeMB
        [pscustomobject]@{
            DiskNumber      = $e.DiskNumber
            PartitionNumber = $e.PartitionNumber
            VolumeGuidPath  = $e.VolumeGuidPath
            FileSystem      = $e.FileSystem
            Label           = $e.Label
            SizeMB          = $e.SizeMB
            FreeMB          = $e.FreeMB
            UsedMB          = $e.UsedMB
            FreePct         = $e.FreePct
            Status          = $status
        }
    }

    # Overall status = worst-case across all ESPs (some rigs have a recovery+ESP combo).
    if (-not $esps) {
        $overall  = 'NOT_FOUND'
        $exitCode = $EXIT_NOTFOUND
    } else {
        $hasCrit = $partitionPayload | Where-Object { $_.Status -eq 'CRITICAL' }
        $hasWarn = $partitionPayload | Where-Object { $_.Status -eq 'WARNING' }
        if     ($hasCrit) { $overall = 'CRITICAL'; $exitCode = $EXIT_CRITICAL }
        elseif ($hasWarn) { $overall = 'WARNING';  $exitCode = $EXIT_WARN }
        else              { $overall = 'OK';       $exitCode = $EXIT_OK }
    }

    $payload = [pscustomobject]@{
        timestamp        = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        hostname         = $env:COMPUTERNAME
        os_build         = (Get-CimInstance -ClassName Win32_OperatingSystem).BuildNumber
        thresholds_mb    = [pscustomobject]@{ warning = $WarningFreeMB; critical = $CriticalFreeMB }
        esp_count        = $partitionPayload.Count
        overall_status   = $overall
        partitions       = $partitionPayload
    }

    $json = if ($Pretty) {
        $payload | ConvertTo-Json -Depth 6
    } else {
        # ConvertTo-Json emits a multi-line object even at -Compress for nested arrays;
        # collapse to a true single line so log scrapers (Loki, Promtail) treat it as one event.
        ($payload | ConvertTo-Json -Depth 6 -Compress)
    }

    Write-Output $json

    if ($OutputPath) {
        $dir = Split-Path -Path $OutputPath -Parent
        if ($dir -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
        # Append, UTF-8 (no BOM), Unix newline so Promtail/Filebeat don't see partial reads.
        [System.IO.File]::AppendAllText(
            $OutputPath,
            ($json + "`n"),
            (New-Object System.Text.UTF8Encoding $false)
        )
    }

    exit $exitCode
}
catch {
    $err = [pscustomobject]@{
        timestamp      = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        hostname       = $env:COMPUTERNAME
        overall_status = 'ERROR'
        error          = $_.Exception.Message
    }
    Write-Output ($err | ConvertTo-Json -Depth 4 -Compress)
    exit $EXIT_ERROR
}
