# Get-EspPartitionStatus.ps1

Reports EFI System Partition (ESP) size and free space as a single line of JSON for Grafana / Loki / Promtail / Telegraf ingest. Built to flag devices in the danger zone for [KB5089549](https://support.microsoft.com/) (May 2026 cumulative update) where `0x800f0922` failures roll back at ~35-36 % when the ESP has &lt;= 10 MB free.

## Background

After installing the May 2026 Windows security update **KB5089549**, devices with limited free space on the EFI System Partition can fail at the post-reboot phase with `0x800f0922` and the message *"Something didn't go as planned. Undoing changes."*. `C:\Windows\Logs\CBS\CBS.log` on affected devices shows:

```
SpaceCheck: Insufficient free space
ServicingBootFiles failed. Error = 0x70
SpaceCheck: <value> used by third-party/OEM files outside of Microsoft boot directories
```

The advisory workaround (apply at your own risk - back up the registry first) is:

```cmd
reg add "HKLM\SYSTEM\CurrentControlSet\Control\Bfsvc" /v EspPaddingPercent /t REG_DWORD /d 0 /f
```

This script does **not** apply the workaround - it only reports the symptom so the dashboard can flag affected devices.

## Behavior

1. Enumerates every partition with the EFI System Partition GPT type (`{c12a7328-f81f-11d2-ba4b-00a0c93ec93b}`).
2. Correlates each partition with `Win32_Volume` via the volume-GUID access path (`\\?\Volume{...}\`). No `mountvol` / drive-letter assignment is performed - ESPs are typically letterless and read-only.
3. Emits a single JSON object on stdout with `timestamp`, `hostname`, `os_build`, thresholds, `overall_status` (worst of all ESPs), and a `partitions[]` array carrying per-ESP `SizeMB` / `FreeMB` / `UsedMB` / `FreePct` / `Status`.
4. Optionally appends the same JSON line to `-OutputPath` (UTF-8, no BOM, Unix newline) for Promtail / Filebeat file-tail ingestion.

## Exit codes

| Code | Meaning |
|------|---------|
| `0`  | OK - free space &gt; `-WarningFreeMB` on every ESP |
| `1`  | WARNING - at least one ESP between `-CriticalFreeMB` and `-WarningFreeMB` |
| `2`  | CRITICAL - at least one ESP at or below `-CriticalFreeMB` |
| `3`  | ESP not found / partition table not GPT |
| `4`  | Unexpected error (caught at the top level; JSON envelope still emitted) |

Default thresholds are `-CriticalFreeMB 15` and `-WarningFreeMB 30`. KB5089549 fails at &lt;= 10 MB free, so the default critical threshold is intentionally set above that to alert before the update fires.

## Examples

### Full script (recommended)

```powershell
# One-shot run (stdout only)
.\Get-EspPartitionStatus.ps1

# Telegraf exec input + Loki tail: write each run as a line into a log file
.\Get-EspPartitionStatus.ps1 -OutputPath 'C:\ProgramData\EspMonitor\esp_status.log'

# Tighter alert: warn at 50 MB, critical at 20 MB
.\Get-EspPartitionStatus.ps1 -WarningFreeMB 50 -CriticalFreeMB 20

# Indented JSON for ad-hoc inspection (do NOT use for Loki scrape)
.\Get-EspPartitionStatus.ps1 -Pretty
```

Sample output (compact):

```json
{"timestamp":"2026-05-26T10:05:30Z","hostname":"WS-001","os_build":"26200","thresholds_mb":{"warning":30,"critical":15},"esp_count":1,"overall_status":"OK","partitions":[{"DiskNumber":0,"PartitionNumber":1,"VolumeGuidPath":"\\\\?\\Volume{60c2128d-9323-4889-9f05-850504f1fd47}\\","FileSystem":"FAT32","Label":"SYSTEM","SizeMB":256.0,"FreeMB":217.61,"UsedMB":38.39,"FreePct":85.0,"Status":"OK"}]}
```

### Minimal two-liner (no JSON, no exit-code envelope)

When the only goal is *"what size is the ESP and how much is free"* and the dashboard can scrape plain `Select-Object` output, the core query collapses to two lines:

```powershell
$esp = Get-Partition | Where-Object GptType -eq '{c12a7328-f81f-11d2-ba4b-00a0c93ec93b}' | Select-Object -First 1
Get-CimInstance Win32_Volume | Where-Object DeviceID -eq (($esp.AccessPaths | Where-Object { $_ -like '\\?\Volume*' }) | Select-Object -First 1) | Select-Object @{n='SizeMB';e={[math]::Round($_.Capacity/1MB,2)}}, @{n='FreeMB';e={[math]::Round($_.FreeSpace/1MB,2)}}, @{n='FreePct';e={[math]::Round($_.FreeSpace/$_.Capacity*100,2)}}
```

Sample output:

```
SizeMB FreeMB FreePct
------ ------ -------
   256 217.61      85
```

For a single-line `key=value` form (Telegraf `exec` / Promtail-friendly), still two lines:

```powershell
$esp = Get-Partition | Where-Object GptType -eq '{c12a7328-f81f-11d2-ba4b-00a0c93ec93b}' | Select-Object -First 1
Get-CimInstance Win32_Volume | Where-Object DeviceID -eq (($esp.AccessPaths | Where-Object { $_ -like '\\?\Volume*' }) | Select-Object -First 1) | ForEach-Object { "esp_size_mb={0} esp_free_mb={1} esp_free_pct={2}" -f [math]::Round($_.Capacity/1MB,2), [math]::Round($_.FreeSpace/1MB,2), [math]::Round($_.FreeSpace/$_.Capacity*100,2) }
```

What the two-liner gives up vs. the full script:

- **No multi-ESP handling.** Picks the first ESP. Rare (some Surface SKUs, BitLocker decrypt-on-write rigs) but worth knowing.
- **No exit codes for alerting.** Grafana has to threshold on the metric itself instead of the runner's exit code.
- **No `-OutputPath` tee** for Promtail file-tail; pipe the output yourself.
- **No `hostname` / `os_build` / `timestamp` envelope** - add them yourself if the dashboard needs labels.
- **No error handling** - if `Get-Partition` fails (very rare, e.g. on a non-GPT disk) the runner sees an exception instead of `overall_status=ERROR`.

If the dashboard contract is fixed and you want proper exit codes for Intune Proactive Remediation or alerting, keep the full script.

## Parameter quick reference

| Parameter         | Default | Notes |
|-------------------|---------|-------|
| `-CriticalFreeMB` | `15`    | Free MB at or below which exit code is `2`. KB5089549 fails at &lt;= 10 MB. |
| `-WarningFreeMB`  | `30`    | Free MB at or below (but above critical) which exit code is `1`. |
| `-OutputPath`     | -       | Optional file the JSON line is appended to (UTF-8 no BOM, `\n`). |
| `-Pretty`         | off     | Indented JSON. Do not enable for line-based scrape jobs. |

## Requirements

- PowerShell 5.1+ (no PS 7-only syntax).
- `Get-Partition` (Storage module, shipped with Windows 10/11/Server).
- CIM access for `Win32_Volume` and `Win32_OperatingSystem` (standard user is sufficient; no elevation needed).
