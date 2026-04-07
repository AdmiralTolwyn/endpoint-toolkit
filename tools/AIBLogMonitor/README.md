# AIB Log Monitor

**Real-time Packer log streaming and build analytics for Azure Image Builder.**

AIB Log Monitor is a PowerShell/WPF desktop application that discovers active AIB builds, streams Packer logs from Azure Container Instances in real time, and provides build health grading, phase progression tracking, cost estimation, and error clustering — all from a single pane of glass.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![WPF](https://img.shields.io/badge/GUI-WPF-blueviolet)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)

---

## Why AIB Log Monitor?

Azure Image Builder runs Packer inside ephemeral container instances that are difficult to observe. Build logs are only available after completion (or failure), and diagnosing mid-build issues requires manual `az container logs` polling. AIB Log Monitor solves this by:

- **Streaming live**: Delta-polling ACI container logs at configurable intervals with full backfill from blob storage
- **Classifying everything**: 11 regex-based color-coding rules highlight errors, warnings, package installs, phase transitions, and script execution
- **Grading builds**: A–F health scores based on error/warning ratios with ETA estimation and running cost tracking
- **Clustering errors**: Duplicate detection and known-issue auto-tagging with fix suggestions for 10 common failure patterns

---

## Features

### Live Log Streaming
- Automatic build discovery by scanning AIB staging resource groups
- Delta-polling of ACI container logs at configurable intervals (1–300s)
- Full Packer log backfill from Azure Blob Storage (`packerlogs/customization.log`)
- Automatic subscription switching and session persistence

### Log Analysis & Visualization
- 11 regex-based color-coding rules (errors, warnings, packages, phases, scripts, etc.)
- Build health grading (A–F) based on error/warning ratios
- ETA estimation with progress bar (% complete, minutes remaining)
- Running cost estimate based on detected VM SKU
- Phase progression tracker (11 customization phases)
- Flame chart showing phase duration bars
- Error clustering with duplicate detection
- Known issue auto-tagger with fix suggestions
- Script duration profiler (execution time per customizer script)

### Search & Export
- Full-text search with match highlighting
- Save logs to disk with timestamp naming
- Copy logs to clipboard
- Diff against previous builds
- Auto-save completed/failed logs

### UI
- Dark/Light theme with 40+ color keys and system theme detection
- Collapsible left sidebar (Builds, Info, Settings) and right sidebar (Analytics: Stats, Profiler, History)
- Debug overlay with ring buffer (500 lines, timestamped)
- Build status ticker strip (Running, Failed, OK counts)
- Minimap scroll indicator and error heatmap timeline

### Achievements
15 gamification badges tracking build monitoring milestones (First Build, Speed Demon, Zero Errors, Centurion, and more).

---

## Prerequisites

| Component | Version | Purpose |
|---|---|---|
| PowerShell | 5.1+ | Windows PowerShell or PowerShell 7+ |
| Az.Accounts | 2.0.0+ | Entra ID authentication |
| Az.Resources | 6.0.0+ | Resource group discovery |
| Az.ContainerInstance | 3.0.0+ | Container log streaming |
| Az.Storage | 5.0.0+ | Blob storage access for log backfill |

---

## Quick Start

```batch
Launch_AIBLogMonitor.bat
```

The launcher auto-detects PowerShell version and unblocks files. Once the UI opens:

1. **Sign In** — Click the sign-in button for Entra ID authentication
2. **Select Build** — Choose an active build from the left sidebar
3. **Stream** — Logs appear in the center panel with real-time color coding
4. **Analyze** — View health score, phase progress, and errors in the right sidebar
5. **Export** — Save the log via the Save button or Ctrl+S

---

## Files

| File | Description |
|---|---|
| `AIBLogMonitor.ps1` | Main application (~4,900 lines) |
| `AIBLogMonitor_UI.xaml` | WPF UI definition |
| `Launch_AIBLogMonitor.bat` | Smart launcher (detects pwsh vs powershell) |
| `user_prefs.json` | Persisted user preferences |
| `achievements.json` | Unlocked achievement badges |
| `build_history.json` | Historical build records with phase durations |
| `saved_logs/` | Exported/archived log files |
