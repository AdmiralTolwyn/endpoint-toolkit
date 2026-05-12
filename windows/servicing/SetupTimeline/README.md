# Get-SetupTimeline.ps1

Reconstructs a phase-by-phase timeline of a Windows in-place upgrade by parsing
`setupact.log`, separating **active upgrade time** from **idle time** the
device spent shut down, hibernated, or asleep.

Built for the most common feature-update post-mortem question:

> *"My users say the upgrade took the whole afternoon — but how long did Setup
> actually need?"*

## Highlights

- **Online vs offline phase split** — by default only the **offline** phases
  (Safe OS, First Boot, OOBE…) count toward "active upgrade time". The
  **online** phases (Downlevel, Pre-Finalize, Finalize) run inside the
  source OS while the user can still work, so they are excluded by default
  but visible as `online` in the `Mode` column. Use `-IncludeDownlevel` to
  fold them into the total.
- **Compiled C# scanner** (StreamReader + precompiled regex) — multi-GB
  `setupact.log` files parse at ~30 MB/s.
- **Idle-gap detection** — gaps between consecutive log timestamps that exceed
  `-IdleGapSeconds` (default **600 s** / 10 min) are subtracted from each
  phase's active duration.
- **In-flight upgrade aware** — trailing phases without an exit marker are
  tagged `(running)` and closed at the log's last timestamp.
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

# Include the online phases (Downlevel, Pre-Finalize, Finalize - excluded by default):
.\Get-SetupTimeline.ps1 -LogPath C:\Windows\Panther -IncludeDownlevel

# Just the rounded minute count (CI / telemetry):
$mins = .\Get-SetupTimeline.ps1 -LogPath C:\Windows\Panther -TotalActiveMinutes
```

## Online vs offline phases

Windows in-place upgrade phases fall into two buckets:

| Mode | Phases | What's happening | User experience |
|------|--------|------------------|-----------------|
| **online**  | `Downlevel`, `Pre-Finalize`, `Finalize` | Setup runs inside the old / source OS — staging the image, copying files, applying drivers, writing the BCD. | Device is up, desktop available, **user can keep working**. |
| **offline** | `Safe OS`, `Pre First Boot`, `Pre SysPrep`, `Post SysPrep`, `Post First Boot`, `Pre OOBE Boot`, `Pre OOBE`, `Post OOBE`, `Post OOBE Boot`, `End` (and synthetic `WinDeploy/OOBE`) | Device has rebooted into WinRE / Safe OS or the first boot of the new OS. | Device is locked / OOBE screen — **user is locked out**. |

By default the script reports **only offline time** — the "how long was the
user locked out" number that IT typically wants to quote. The `Mode` column
still shows the classification for each row, and `-IncludeDownlevel` folds
the online phases back into the total.

## Parameters

| Parameter | Default | Purpose |
|---|---|---|
| `-LogPath` | auto-discover | File or folder containing `setupact.log`. Folder may be a Panther root or its `UnattendGC` subfolder. If omitted, the script probes `C:\Windows\Panther` and then `C:\$WINDOWS.~BT\Sources\Panther`. |
| `-IncludeDownlevel` | off | Include the **online** phases (Downlevel, Pre-Finalize, Finalize). Excluded by default because the user is still productive while they run in the source OS. Name kept for backward compat — the switch covers all three online phases, not just Downlevel. |
| `-IdleGapSeconds` | `600` | Threshold above which a gap between consecutive log timestamps is treated as the device being off / asleep. Empirically 600 s avoids false positives from legitimate Setup pauses (driver install, dynamic update download, BCD writes). |
| `-AsObject` | off | Emit the timeline as `PSCustomObject`s (`Phase`, `Mode`, `Start`, `End`, `Duration`, `Idle`, `Gap`, `Wall`, `HResult`). |
| `-Csv` | off | Emit the timeline as a semicolon-separated CSV on stdout. |
| `-TotalActiveMinutes` | off | Suppress the table and emit only the rounded total active upgrade time in minutes. |

## Sample output — default (offline phases only)

```text
  Setup Timeline
  Log:        C:\Windows\Panther\setupact.log
  Online:     excluded (Downlevel, Pre-Finalize, Finalize - user productive)
  Idle gap:   > 600s treated as off / standby / sleep

  Phase            Mode             Start            End      Active        Idle       Gap  HRESULT
----------------------------------------------------------------------------------------------------
  Safe OS          offline 04-13 12:54:29       13:11:40     17m 11s           -         -  0x00000000
  Pre First Boot   offline       13:12:18       13:12:36         18s           -       38s  -
  Pre SysPrep      offline       13:12:36       13:12:37          1s           -         -  0x00000000
  Post SysPrep     offline       13:15:08       13:15:08          0s           -    3m 31s  -
  Post First Boot  offline       13:15:08       13:16:40      2m 32s           -         -  -
  Pre OOBE Boot    offline       13:16:40       13:18:13      2m 33s           -         -  -
  Pre OOBE         offline       13:18:13       13:18:54         41s           -         -  0x00000000
  Post OOBE        offline       13:18:56       13:20:11      1m 15s           -        2s  -
  Post OOBE Boot   offline       13:20:11       13:20:11          0s           -         -  -
  End              offline       13:20:11       13:20:11          0s           -         -  -
----------------------------------------------------------------------------------------------------
  Active upgrade time : 23m 31s         (23 min)
  Excluded idle time  : 0s
  Inter-phase gaps    : 3m 11s          (reboots and phase handoffs)
  Wall-clock span     : 26m 42s         (2026-04-13 12:54 -> 2026-04-13 13:20)
```

In this example the wall clock from first reboot to End is ~27 min, of which
Setup was actually running offline work for **23 min** — that is the time the
user was locked out. The 3m 11s of inter-phase gap time is reboots and phase
handoffs (not user-visible work). The online portion (Downlevel /
Pre-Finalize / Finalize) is hidden by default and would add a further ~1h
04m of "upgrade in the background while you keep working" time.

## Sample output — `-IncludeDownlevel` (online + offline)

```text
  Setup Timeline
  Log:        C:\Windows\Panther\setupact.log
  Online:     INCLUDED (Downlevel, Pre-Finalize, Finalize)
  Idle gap:   > 600s treated as off / standby / sleep

  Phase            Mode             Start            End      Active        Idle       Gap  HRESULT
----------------------------------------------------------------------------------------------------
  Downlevel        online  04-13 11:53:56       12:05:46     12m 50s           -         -  -
  Pre-Finalize     online        12:05:46       12:49:32     44m 46s           -         -  -
  Finalize         online        12:49:32       12:54:29      5m 57s           -         -  -
  Safe OS          offline       12:54:29       13:11:40     17m 11s           -         -  0x00000000
  ...
----------------------------------------------------------------------------------------------------
  Active upgrade time : 1h 23m 04s      (83 min)
  Excluded idle time  : 0s
  Inter-phase gaps    : 3m 11s          (reboots and phase handoffs)
  Wall-clock span     : 1h 26m 15s      (2026-04-13 11:53 -> 2026-04-13 13:20)
```

## Sample output — `-Csv`

```csv
"Phase";"Mode";"Start";"End";"ActiveSec";"IdleSec";"GapSec";"WallSec";"HResult"
"Safe OS";"offline";"2026-04-13 12:54:29";"2026-04-13 13:11:40";"1031";"0";"0";"1031";"0x00000000"
"Pre First Boot";"offline";"2026-04-13 13:12:18";"2026-04-13 13:12:36";"18";"0";"38";"18";
"Pre SysPrep";"offline";"2026-04-13 13:12:36";"2026-04-13 13:12:37";"1";"0";"0";"1";"0x00000000"
...
```

## Sample output — `-TotalActiveMinutes`

```text
23
```

## Sample output — `-Verbose`

```text
VERBOSE: Get-SetupTimeline v1.2.0
VERBOSE: PowerShell 5.1.26100.8115 on Microsoft Windows NT 10.0.26200.0
VERBOSE: Main log     : C:\Windows\Panther\setupact.log (710.3 MB)
VERBOSE: UnattendGC   : C:\Windows\Panther\UnattendGC\setupact.log (0.1 MB)
VERBOSE: Idle gap     : > 600 seconds treated as off / standby
VERBOSE: Online phases: excluded (Downlevel, Pre-Finalize, Finalize)
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

| Phase | Mode | Notes |
|---|---|---|
| `Downlevel` | online | Old OS still running. Image staging starts. User productive. |
| `Pre-Finalize` | online | Bulk of the offline work pre-staged online — typically the largest phase. Image copy, driver staging. User productive. |
| `Finalize` | online | Last step before the SafeOS reboot. BCD write, final disk prep. User productive. |
| `Safe OS` | offline | WinRE phase after reboot. Driver injection, image apply. User locked out. |
| `WinDeploy/OOBE` | offline | Synthetic — bracketed by `windeploy.exe` log lines from `UnattendGC\setupact.log`. |
| `Pre First Boot` / `Post First Boot` | offline | First-boot SetupPlatform passes. |
| `Pre SysPrep` / `Post SysPrep` | offline | Sysprep specialize. |
| `Pre OOBE Boot` / `Pre OOBE` / `Post OOBE` / `Post OOBE Boot` | offline | OOBE pipeline. |
| `End` | offline | Terminal transition marker (collapsed to 0s; logs continue but Setup is done). |

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
