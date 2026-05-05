# Get-SetupTimeline.ps1

Reconstructs a phase-by-phase timeline of a Windows in-place upgrade by parsing
`setupact.log`, separating **active upgrade time** from **idle time** the
device spent shut down, hibernated, or asleep.

Built for the most common feature-update post-mortem question:

> *"My users say the upgrade took the whole afternoon — but how long did Setup
> actually need?"*

## Highlights

- **Compiled C# scanner** (StreamReader + precompiled regex) — multi-GB
  `setupact.log` files parse at ~30 MB/s.
- **Idle-gap detection** — gaps between consecutive log timestamps that exceed
  `-IdleGapSeconds` (default **600 s** / 10 min) are subtracted from each
  phase's active duration.
- **Recognises the canonical Microsoft layout** — `<root>\setupact.log` plus
  `<root>\UnattendGC\setupact.log`. Works on `C:\Windows\Panther` (post-upgrade
  success path), `C:\$WINDOWS.~BT\Sources\Panther` (in-flight / rolled-back),
  or any custom export that mirrors the same structure.
- **Auto-discovery** — run with no parameters on the just-upgraded machine and
  the script finds the logs itself.
- **Multiple output modes** — formatted table (default), `-AsObject`, `-Csv`
  (Excel / Power BI), `-TotalActiveMinutes` (single integer for CI / telemetry).
- **Compatible with Windows PowerShell 5.1 and PowerShell 7+.**

## Quick start

```powershell
# Run on the just-upgraded machine — no parameters needed:
.\Get-SetupTimeline.ps1 -Verbose

# Analyse a captured Panther archive:
.\Get-SetupTimeline.ps1 -LogPath C:\Cases\Device42\Panther

# Include the Downlevel phase (excluded by default):
.\Get-SetupTimeline.ps1 -LogPath C:\Windows\Panther -IncludeDownlevel

# Just the rounded minute count (CI / telemetry):
$mins = .\Get-SetupTimeline.ps1 -LogPath C:\Windows\Panther -TotalActiveMinutes
```

## Parameters

| Parameter | Default | Purpose |
|---|---|---|
| `-LogPath` | auto-discover | File or folder containing `setupact.log`. Folder may be a Panther root or its `UnattendGC` subfolder. If omitted, the script probes `C:\Windows\Panther` and then `C:\$WINDOWS.~BT\Sources\Panther`. |
| `-IncludeDownlevel` | off | Include the Downlevel phase. Excluded by default because the user is still productive while it runs in the background. |
| `-IdleGapSeconds` | `600` | Threshold above which a gap between consecutive log timestamps is treated as the device being off / asleep. Empirically 600 s avoids false positives from legitimate Setup pauses (driver install, dynamic update download, BCD writes). |
| `-AsObject` | off | Emit the timeline as `PSCustomObject`s (`Phase`, `Start`, `End`, `Duration`, `Idle`, `Wall`, `HResult`). |
| `-Csv` | off | Emit the timeline as a semicolon-separated CSV on stdout. |
| `-TotalActiveMinutes` | off | Suppress the table and emit only the rounded total active upgrade time in minutes. |

## Sample output — default

```text
  Setup Timeline
  Log:       C:\Windows\Panther\setupact.log
  Downlevel: excluded (user productive)
  Idle gap:  > 600s treated as off / standby / sleep

  Phase                     Start            End      Active        Idle  HRESULT
--------------------------------------------------------------------------------
  Pre-Finalize     04-28 11:42:14       16:54:41     35m 10s  5h 37m 17s  -
  Finalize               16:54:41       17:00:47      6m 06s           -  -
  Safe OS                17:00:47       17:21:05     20m 18s           -  0x00000000
  WinDeploy/OOBE         17:21:51       17:21:52          1s           -  -
  Pre First Boot         17:21:52       17:22:16         24s           -  -
  Pre SysPrep            17:22:16       17:22:17          1s           -  0x00000000
  Post SysPrep           17:24:01       17:24:01          0s           -  -
  Post First Boot        17:24:01       17:25:51      2m 50s           -  -
  Pre OOBE Boot          17:25:51       17:27:14      1m 23s           -  -
  Pre OOBE               17:27:14       17:27:42         28s           -  0x00000000
  Post OOBE              17:27:43       17:29:00      1m 17s           -  -
  Post OOBE Boot         17:29:00       17:29:00          0s           -  -
  End                    17:29:00       17:30:35      2m 35s           -  -
--------------------------------------------------------------------------------
  Active upgrade time : 1h 08m 33s      (69 min)
  Excluded idle time  : 5h 37m 17s
  Wall-clock span     : 6h 48m 21s      (2026-04-28 11:42 -> 2026-04-28 17:30)
```

In this real-world example the user's wall-clock perception of "almost 7
hours" is correct, but Setup itself only needed **69 minutes** — the
remaining **5 h 37 min** was a single idle window where the device was
suspended after the Pre-Finalize phase reached its first reboot.

## Sample output — `-IncludeDownlevel`

```text
  Setup Timeline
  Log:       C:\Windows\Panther\setupact.log
  Downlevel: INCLUDED
  Idle gap:  > 600s treated as off / standby / sleep

  Phase                     Start            End      Active        Idle  HRESULT
--------------------------------------------------------------------------------
  Downlevel        04-28 11:32:09       11:42:14     10m 05s           -  -
  Pre-Finalize           11:42:14       16:54:41     35m 10s  5h 37m 17s  -
  Finalize               16:54:41       17:00:47      6m 06s           -  -
  ...
--------------------------------------------------------------------------------
  Active upgrade time : 1h 18m 38s      (79 min)
  Excluded idle time  : 5h 37m 17s
  Wall-clock span     : 6h 58m 26s      (2026-04-28 11:32 -> 2026-04-28 17:30)
```

## Sample output — `-Csv`

```csv
"Phase";"Start";"End";"ActiveSec";"IdleSec";"WallSec";"HResult"
"Pre-Finalize";"2026-04-28 11:42:14";"2026-04-28 16:54:41";"2110";"16637";"18747";
"Finalize";"2026-04-28 16:54:41";"2026-04-28 17:00:47";"366";"0";"366";
"Safe OS";"2026-04-28 17:00:47";"2026-04-28 17:21:05";"1218";"0";"1218";"0x00000000"
"WinDeploy/OOBE";"2026-04-28 17:21:51";"2026-04-28 17:21:52";"1";"0";"1";
...
```

## Sample output — `-TotalActiveMinutes`

```text
69
```

## Sample output — `-Verbose`

```text
VERBOSE: Get-SetupTimeline v1.1.0
VERBOSE: PowerShell 5.1.26100.8115 on Microsoft Windows NT 10.0.26200.0
VERBOSE: Main log     : C:\Windows\Panther\setupact.log (710.3 MB)
VERBOSE: UnattendGC   : C:\Windows\Panther\UnattendGC\setupact.log (0.1 MB)
VERBOSE: Idle gap     : > 600 seconds treated as off / standby
VERBOSE: Downlevel    : excluded
VERBOSE: Scanned setupact.log in 23.9s (29.7 MB/s) - 20 phase markers, 6 idle gaps
VERBOSE: Scanned setupact.log in 0.0s - 2 markers, 7 idle gaps
```

## Why a C# scanner?

`setupact.log` is routinely **500 MB – 1 GB** with millions of lines, and
every line needs both a timestamp parse (for idle-gap detection) and a regex
match (for phase markers). A pure PowerShell loop using `Get-Content` /
`Select-String` is **50–100× slower** than the equivalent .NET code on the
same machine — a single run can take 20+ minutes on a large log.

To stay practical, the script compiles a small C# helper into the current
PowerShell session via `Add-Type` on first use. The helper:

- Opens the file with `FileStream` + `StreamReader` (64 KB buffer,
  `FileShare.ReadWrite` so it can read a log that Setup is still writing).
- Uses **`Regex.Compiled`** patterns — paid once at JIT time, then matched
  against every line at native speed.
- Returns a strongly-typed `ScanResult` (markers + idle gaps + last
  timestamp + line count) in a single pass.

End result: ~30 MB/s throughput, so a 700 MB log finishes in under 30 s. The
compiled type stays loaded for the lifetime of the PowerShell process, so
subsequent runs in the same session skip the compilation step entirely.

## How the timeline is built

1. **Scan** the main `setupact.log` once with a compiled C# `StreamReader` /
   `Regex.Compiled` pipeline, capturing:
   - `OPERATIONTRACK ExecuteOperations: Start execution phase <Name>` markers
   - `Execution phase [<Name>] exiting with HRESULT [0x...]` markers
   - **Every** timestamped line, to detect gaps between consecutive lines.
2. **Optionally scan** the OOBE log
   (`<root>\UnattendGC\setupact.log`) for `[windeploy.exe]` lines to add a
   synthetic `WinDeploy/OOBE` segment that the main log does not record.
3. **Sort** markers by `(Time, Order)`. The secondary key matters because
   `Sort-Object` is not stable on Windows PowerShell 5.1 and several markers
   can share a 1-second timestamp.
4. **Build segments** by walking markers in order; each `Start` opens a phase
   and the next `Exit` (or the next `Start` of a different phase) closes it.
   Same-phase repeats (e.g. Downlevel restart) are coalesced.
5. **Compute Active = Wall − overlapping idle gaps** for every segment.
6. **Render** the table, optionally suppressed by `-AsObject`, `-Csv`, or
   `-TotalActiveMinutes`.

## Phase reference

| Phase | Notes |
|---|---|
| `Downlevel` | Old OS still running. User productive. Excluded by default. |
| `Pre-Finalize` | Bulk of the offline work pre-staged online. Largest phase on most modern devices. |
| `Finalize` | Last online step before the SafeOS reboot. |
| `Safe OS` | WinRE phase. Driver injection, image apply. |
| `WinDeploy/OOBE` | Synthetic — bracketed by `windeploy.exe` log lines from `UnattendGC\setupact.log`. |
| `Pre First Boot` / `Post First Boot` | First-boot SetupPlatform passes. |
| `Pre SysPrep` / `Post SysPrep` | Sysprep specialize. |
| `Pre OOBE Boot` / `Pre OOBE` / `Post OOBE` / `Post OOBE Boot` | OOBE pipeline. |
| `End` | Final SetupPlatform pass; logs close out. |

## Limitations

- The synthetic `WinDeploy/OOBE` window is bounded by the first and last
  `[windeploy.exe]` lines in the UnattendGC log; OOBE phases that do not log
  through windeploy.exe will not extend it.
- Idle-gap detection is based purely on log-timestamp deltas. A device that
  stays powered on but where every Setup component happens to fall silent for
  more than `-IdleGapSeconds` will be misclassified as idle. The default of
  600 s is conservative enough that this has not been observed in practice.
- The script reports timestamps in the local time zone recorded by Setup; it
  does not convert across DST transitions inside the upgrade window.
