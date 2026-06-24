# Get-SetupTimeline.ps1

Reconstructs a phase-by-phase timeline of a Windows in-place upgrade by parsing
`setupact.log`, separating **active upgrade time** from **idle time** the
device spent shut down, hibernated, or asleep.

Built for the most common feature-update post-mortem question:

> *"My users say the upgrade took the whole afternoon — but how long did Setup
> actually need?"*

## Highlights

- **Online vs offline phase split** — by default only the **offline** phases
  (Finalize, Safe OS, First Boot, OOBE…) count toward "active upgrade time".
  The **online** phases (Pre-Downlevel lead-in, Downlevel, Pre-Finalize) run
  inside the source OS while the user can still work, so they are excluded by
  default but visible as `online` in the `Mode` column. Use `-IncludeDownlevel`
  to fold them into the total.
- **Whole-experience anchoring** — a synthetic **`Pre-Downlevel`** lead-in
  segment runs from the log's first timestamp to the first phase marker,
  capturing the SetupHost / compatibility-scan / dynamic-update download window
  that precedes formal phase tracking. It is tagged `online` and gated behind
  `-IncludeDownlevel`, so the default "lockout time" view is unchanged.
- **Dynamic Update download breakdown** (`-ShowDownloads`) — parses the DCAT
  transfer markers to answer "was it the download or the apply that took so
  long?". Reports file count, largest payload sizes (LCU / FODs / language
  packs), download span, instantaneous throughput (avg / peak / min) and how
  long the transfer ran below 2 Mbps — all in the same single pass.
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

# Include the online phases (Pre-Downlevel, Downlevel, Pre-Finalize - excluded by default):
.\Get-SetupTimeline.ps1 -LogPath C:\Windows\Panther -IncludeDownlevel

# Break down the Dynamic Update download (size, throughput, slow-link time):
.\Get-SetupTimeline.ps1 -LogPath C:\Windows\Panther -IncludeDownlevel -ShowDownloads

# Just the rounded minute count (CI / telemetry):
$mins = .\Get-SetupTimeline.ps1 -LogPath C:\Windows\Panther -TotalActiveMinutes
```

## Online vs offline phases

Windows in-place upgrade phases fall into two buckets:

| Mode | Phases | What's happening | User experience |
|------|--------|------------------|-----------------|
| **online**  | `Pre-Downlevel` (synthetic lead-in), `Downlevel`, `Pre-Finalize` | SetupHost initialization, compatibility scan, dynamic update / ESD download, then Setup staging the image, copying files, applying drivers — all inside the old / source OS. | Device is up, desktop available, **user can keep working**. |
| **offline** | `Finalize`, `Safe OS`, `Pre First Boot`, `Pre SysPrep`, `Post SysPrep`, `Post First Boot`, `Pre OOBE Boot`, `Pre OOBE`, `Post OOBE`, `Post OOBE Boot`, `End` (and synthetic `WinDeploy/OOBE`) | Finalize shows the full-screen "restarting" UI; everything after has rebooted into WinRE / Safe OS or the first boot of the new OS. | Desktop gone / OOBE screen — **user is locked out**. |

By default the script reports **only offline time** — the "how long was the
user locked out" number that IT typically wants to quote. The `Mode` column
still shows the classification for each row, and `-IncludeDownlevel` folds
the online phases back into the total.

## Parameters

| Parameter | Default | Purpose |
|---|---|---|
| `-LogPath` | auto-discover | File or folder containing `setupact.log`. Folder may be a Panther root or its `UnattendGC` subfolder. If omitted, the script probes `C:\Windows\Panther` and then `C:\$WINDOWS.~BT\Sources\Panther`. |
| `-IncludeDownlevel` | off | Include the **online** phases (Pre-Downlevel lead-in, Downlevel, Pre-Finalize). Excluded by default because the user is still productive while they run in the source OS. Name kept for backward compat — the switch covers all online phases, not just Downlevel. |
| `-IdleGapSeconds` | `600` | Threshold above which a gap between consecutive log timestamps is treated as the device being off / asleep. Empirically 600 s avoids false positives from legitimate Setup pauses (driver install, dynamic update download, BCD writes). |
| `-AsObject` | off | Emit the timeline as `PSCustomObject`s (`Phase`, `Mode`, `Start`, `End`, `Duration`, `Idle`, `Gap`, `Wall`, `HResult`). |
| `-Csv` | off | Emit the timeline as a semicolon-separated CSV on stdout. |
| `-ShowDownloads` | off | Append a Dynamic Update / download breakdown beneath the timeline (file count, largest payload sizes, download span, throughput, slow-link time). Parses DCAT transfer markers in the same pass. Applies to the default report only (ignored with `-AsObject` / `-Csv` / `-TotalActiveMinutes`). |
| `-TotalActiveMinutes` | off | Suppress the table and emit only the rounded total active upgrade time in minutes. |

## Sample output — default (offline phases only)

```text
  Setup Timeline
  Log:        C:\Windows\Panther\setupact.log
  Online:     excluded (Downlevel, Pre-Finalize - user productive)
  Idle gap:   > 600s treated as off / standby / sleep

  Phase            Mode             Start            End      Active        Idle       Gap  HRESULT
----------------------------------------------------------------------------------------------------
  Finalize         offline 04-13 12:49:32       12:54:29      5m 57s           -         -  -
  Safe OS          offline       12:54:29       13:11:40     17m 11s           -         -  0x00000000
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
  Active upgrade time : 27m 28s         (27 min)
  Excluded idle time  : 0s
  Inter-phase gaps    : 3m 11s          (reboots and phase handoffs)
  Wall-clock span     : 31m 39s         (2026-04-13 12:49 -> 2026-04-13 13:20)
```

In this example the wall clock from Finalize to End is ~32 min, of which
Setup did **27 min** of offline work — that is the time the user was locked
out (Finalize's full-screen "restarting" UI through the post-OOBE first
login). The 3m 11s of inter-phase gap time is reboots and phase handoffs
(not user-visible work). The online portion (Downlevel + Pre-Finalize) is
hidden by default and would add a further ~58 min of "upgrade in the
background while you keep working" time.

## Sample output — `-IncludeDownlevel` (online + offline)

```text
  Setup Timeline
  Log:        C:\Windows\Panther\setupact.log
  Online:     INCLUDED (Pre-Downlevel, Downlevel, Pre-Finalize)
  Idle gap:   > 600s treated as off / standby / sleep

  Phase            Mode             Start            End      Active        Idle       Gap  HRESULT
----------------------------------------------------------------------------------------------------
  Pre-Downlevel    online  04-13 09:41:40       11:53:56  1h 49m 17s     22m 59s         -  -
  Downlevel        online        11:53:56       12:05:46     12m 50s           -         -  -
  Pre-Finalize     online        12:05:46       12:49:32     44m 46s           -         -  -
  Finalize         offline       12:49:32       12:54:29      5m 57s           -         -  -
  Safe OS          offline       12:54:29       13:11:40     17m 11s           -         -  0x00000000
  ...
----------------------------------------------------------------------------------------------------
  Active upgrade time : 3h 12m 21s      (192 min)
  Excluded idle time  : 22m 59s
  Inter-phase gaps    : 3m 11s          (reboots and phase handoffs)
  Wall-clock span     : 3h 38m 31s      (2026-04-13 09:41 -> 2026-04-13 13:20)
```

The `Pre-Downlevel` row is the lead-in window — the log's first timestamp
through the first formal phase marker — where SetupHost runs the compatibility
scan and downloads the upgrade payload. It is `online` (the user keeps working)
and its own idle gaps are subtracted, so it reflects active prep work, not the
full wall-clock window.

## Sample output — `-ShowDownloads`

Appended beneath the timeline. Use it to tell a slow **download** apart from a
slow **apply** — the DCAT transfer markers carry the payload size and the
instantaneous byte rate, so a long `Pre-Downlevel` can be attributed precisely.

```text
  Dynamic Update / Downloads
----------------------------------------------------------------------------------------------------
  DU media version    : 10.0.26100.8037, 10.0.26200.8037
  Files transferred   : 54
  Measured payload    : 5.43 GB   (largest single 4.50 GB)
  Download span       : 09:41:55 -> 11:26:15   (1h 44m 20s wall, 1h 44m 20s active)
  Throughput (inst)   : avg 9.7 Mbps   peak 58.4 Mbps   min 21 Kbps
  Slow-link (<2 Mbps) : 15m 07s of the transfer ran below 2 Mbps

  Largest payloads                                           Size    Duration     Avg rate
  Windows11.0-KB5083769-x64.msu                           4.50 GB  1h 13m 59s     8.7 Mbps
  Windows11.0-KB5043080-x64.msu                          509.0 MB      4m 52s    14.6 Mbps
  Windows11.0-KB5083826-x64.cab                          120.4 MB      1m 30s    11.2 Mbps
  Microsoft-Windows-NetFx3-OnDemand-Package~31bf385...    68.0 MB         43s    13.3 Mbps
  Microsoft-Windows-Client-LanguagePack-Package~31b...    24.4 MB         12s    17.1 Mbps
```

In this example the 4.5 GB checkpoint cumulative update (`KB5083769`) alone took
**74 minutes** at ~8.7 Mbps — the single biggest contributor to the upgrade.
A low **peak** rate or a large **slow-link** figure points at the network /
Delivery Optimization / proxy path rather than at Setup itself. For the actual
transfer source (CDN vs. peer) and any throttling, pull the Delivery
Optimization log separately with `Get-DeliveryOptimizationLog`.

> The download breakdown is independent of `-IncludeDownlevel` — it still
> renders in the default (offline-only) view, since the download happens during
> the online lead-in regardless of which phases the table shows.

## Sample output — `-Csv`

```csv
"Phase";"Mode";"Start";"End";"ActiveSec";"IdleSec";"GapSec";"WallSec";"HResult"
"Finalize";"offline";"2026-04-13 12:49:32";"2026-04-13 12:54:29";"357";"0";"0";"357";
"Safe OS";"offline";"2026-04-13 12:54:29";"2026-04-13 13:11:40";"1031";"0";"0";"1031";"0x00000000"
"Pre First Boot";"offline";"2026-04-13 13:12:18";"2026-04-13 13:12:36";"18";"0";"38";"18";
...
```

## Sample output — `-TotalActiveMinutes`

```text
27
```

## Sample output — `-Verbose`

```text
VERBOSE: Get-SetupTimeline v1.3.0
VERBOSE: PowerShell 5.1.26100.8115 on Microsoft Windows NT 10.0.26200.0
VERBOSE: Main log     : C:\Windows\Panther\setupact.log (710.3 MB)
VERBOSE: UnattendGC   : C:\Windows\Panther\UnattendGC\setupact.log (0.1 MB)
VERBOSE: Idle gap     : > 600 seconds treated as off / standby
VERBOSE: Online phases: excluded (Pre-Downlevel, Downlevel, Pre-Finalize)
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
5. **Prepend the synthetic `Pre-Downlevel` lead-in** from the log's first
   timestamp to the first phase marker (only shown with `-IncludeDownlevel`).
6. **Compute Active = Wall − overlapping idle gaps** for every segment.
7. **Render** the table, optionally suppressed by `-AsObject`, `-Csv`, or
   `-TotalActiveMinutes`.

## Phase reference

| Phase | Mode | Notes |
|---|---|---|
| `Pre-Downlevel` | online | Synthetic lead-in — log's first timestamp to the first phase marker. SetupHost init, compatibility scan, dynamic update / ESD download. User productive. Shown only with `-IncludeDownlevel`. |
| `Downlevel` | online | Old OS still running. Image staging starts. User productive. |
| `Pre-Finalize` | online | Bulk of the offline work pre-staged online — typically the largest phase. Image copy, driver staging. User productive. |
| `Finalize` | offline | Last step before the SafeOS reboot. Setup shows its full-screen "restarting" UI — user is locked out of the desktop. |
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
