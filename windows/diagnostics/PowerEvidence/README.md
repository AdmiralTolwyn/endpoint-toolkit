# Get-PowerEvidence.ps1

**Version:** 1.0.0
**Author:** Anton Romanyuk

> **Disclaimer:** This script is provided "as-is" without warranty of any kind, express or implied. Use at your own risk. The author assumes no liability for any damage or data loss resulting from its use. Always test in a non-production environment before deployment.

Collects a portable evidence bundle for **power, standby, and screen-on behaviour** in one pass, writes a plain-text `summary.txt`, and packages everything as a single `.zip` the user can send back. It is the "gather everything a battery/sleep/wake investigation needs, in the right order, without missing a source" script — so you don't have to talk a user through half a dozen `powercfg` commands over the phone.

It runs **without administrator rights**. When elevated it additionally captures the SleepStudy and System Power reports, which require elevation.

## Problem

"The laptop drains overnight", "it won't sleep", "the screen won't turn off", and "it wakes by itself" are all the same class of ticket, and they all need the same handful of artifacts: available sleep states, the battery report, SleepStudy, wake history, and the power events from the System log. Collecting these by hand means running several `powercfg` variants and a `Get-WinEvent` query per machine, in the right order, remembering which ones need elevation — and then zipping the right folder to send back. This script does all of it and produces one file.

## What it collects

| # | Section | Source | Elevation |
|---|---------|--------|-----------|
| 1 | Available sleep states | `powercfg /a` → `powercfg-a.txt` | No |
| 2 | Battery report (HTML + XML) | `powercfg /batteryreport` → `batteryreport.html` / `.xml` | No |
| 3 | SleepStudy + System Power report | `powercfg /sleepstudy`, `powercfg /systempowerreport` → `sleepstudy.html` / `systempowerreport.html` | **Yes** |
| 4 | Wake diagnostics | `powercfg /lastwake`, `/waketimers`, `/requests` → `powercfg-wake.txt` | `/requests` only |
| 5 | Power events from the System log | Kernel-General / Kernel-Power / Kernel-Boot IDs 12, 13, 27, 41, 42, 107, 109 → `power-events.csv` / `.txt` | No |
| 6 | Fast Startup / Hibernate configuration | `HiberbootEnabled` + sleep-state availability → `summary.txt` | No |

All artifacts land in a timestamped folder (`PowerEvidence_<host>_<stamp>`), `summary.txt` is added, and the folder is compressed to `<folder>.zip` unless `-NoZip` is supplied.

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `OutputRoot` | `string` | `%USERPROFILE%\Desktop` | Parent folder for the evidence bundle. |
| `Days` | `int` | `7` | History window (1–60) for the battery/sleep reports and the System log query. |
| `NoZip` | `switch` | Off | Keep the raw folder only; skip creating the `.zip` archive. |
| `NoLaunch` | `switch` | Off | Do not open the output folder in Explorer when finished. |

## Bundle contents

| File | What it captures |
|---|---|
| `summary.txt` | Run header, per-section status, the last 15 power events, and the Fast Startup / Hibernate / Modern Standby findings. Read this first. |
| `powercfg-a.txt` | Available sleep states (S0/S1/S3/S4, Modern Standby, Hibernate) and why any are unavailable. |
| `batteryreport.html` / `.xml` | Battery capacity, usage, and recent charge/discharge history. |
| `sleepstudy.html` | Modern Standby drain breakdown per session (elevated only). |
| `systempowerreport.html` | Connected-standby / system power summary (elevated only). |
| `powercfg-wake.txt` | Last wake source, armed wake timers, and active power requests (`/requests` elevated only). |
| `power-events.csv` / `.txt` | Kernel power/boot events over the window, decoded to plain-English meanings incl. boot type (cold / Fast Startup / resume). |

The script makes **no changes to the system** — it only writes into its own output folder, then zips it.

## Usage

For the complete data set (SleepStudy, System Power, active power requests), run **as administrator**. Unelevated runs still collect everything else and exit with code `3` (partial).

### Basic — collect 7 days to the Desktop and zip

```powershell
.\Get-PowerEvidence.ps1
# -> creates ...\Desktop\PowerEvidence_<host>_<stamp>.zip; send that back
```

### Wider window, custom location

```powershell
.\Get-PowerEvidence.ps1 -Days 14 -OutputRoot C:\Temp
```

### Automation — raw folder only, no Explorer

```powershell
.\Get-PowerEvidence.ps1 -NoZip -NoLaunch
```

## Interpreting the output

- **Event 41 (Kernel-Power)** — the machine rebooted without a clean shutdown (power loss, hard hang, or bugcheck). Correlate with the surrounding 12/13/109 events.
- **Event 42 → 107 pairs** — sleep entry followed by resume; the gap is how long it actually slept. Frequent short pairs point to something waking it (check `powercfg-wake.txt`).
- **Event 27 `BootType=2`** — a resume from hibernation/Fast Startup rather than a cold boot; useful when "it never really shuts down" is suspected.
- **`Modern Standby (S0) available`** — an S0 machine never enters S3; overnight drain is expected behaviour unless `sleepstudy.html` shows a specific offender.
- **`Fast Startup enabled`** — explains why a "shutdown" behaves like a resume (event 27 `BootType=1`).

## Exit codes

| Code | Meaning |
|---|---|
| `0` | OK — all sections collected (run was elevated). |
| `3` | PARTIAL — ran unelevated; SleepStudy / System Power / `/requests` were skipped. |
| `4` | ERROR — unexpected failure caught at top level. |

## Requirements

- PowerShell 5.1+
- Elevation only for the SleepStudy, System Power, and `/requests` sections; every other section works unelevated.

## References

- [Modern Standby](https://learn.microsoft.com/windows-hardware/design/device-experiences/modern-standby)
- [System power states](https://learn.microsoft.com/windows/win32/power/system-power-states)
