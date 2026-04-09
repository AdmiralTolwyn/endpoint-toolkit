# BaselinePilot — Windows Client Security Baseline Assessment Tool

BaselinePilot is a two-component security baseline assessment tool for Windows 11 clients. It combines Microsoft Security Baselines (Intune MDM + GPO) with ODA-style operational health checks in a rich WPF dashboard.

## Architecture

```
Customer Machine                        Assessor Workstation
┌─────────────────────────┐             ┌──────────────────────────────┐
│ Invoke-BaselineCollection│  JSON file  │ BaselinePilot.ps1 (WPF GUI) │
│ .ps1                     │ ──────────► │ + BaselinePilot_UI.xaml      │
│ (headless, admin, no     │  transfer   │ + checks.json (263 checks)  │
│  external modules)       │             │ + csp_metadata.json         │
└─────────────────────────┘             └──────────────────────────────┘
```

- **Collection** runs on the target machine with local admin rights — no modules, no internet, outputs a single JSON file
- **Assessment** runs on the assessor's workstation — WPF GUI with dashboard, findings, report export

## Quick Start

### 1. Collect Data (on target machine)

```powershell
# Run as Administrator
.\Invoke-BaselineCollection.ps1

# Quick run (skip event log collection, ~30s)
.\Invoke-BaselineCollection.ps1 -SkipEventCollection

# Summary-only events (counts + top-N, not individual events)
.\Invoke-BaselineCollection.ps1 -EventSummaryOnly

# Silent (for automation)
.\Invoke-BaselineCollection.ps1 -Quiet -OutputPath C:\Reports\baseline.json
```

Output: `<hostname>_baseline_<timestamp>.json`

### 2. Assess (on your workstation)

```powershell
# Double-click or run:
.\Launch_BaselinePilot.bat
```

Import the JSON file in the GUI → Dashboard populates with scores, findings, and remediation guidance.

## Data Collection Areas (22)

| # | Area | Method |
|---|------|--------|
| 1 | System Information | CIM/WMI |
| 2 | Join Type Detection | `dsregcmd /status` |
| 3 | Applied Policies | `gpresult /scope computer` |
| 4 | MDM Enrollment | Registry (Enrollments + PolicyManager) |
| 5 | Security Policy Export | `secedit /export` |
| 6 | Audit Policy | `auditpol /get /category:*` |
| 7 | Registry Baselines | ~300 registry keys (Intune + GPO paths) |
| 8 | Defender Configuration | `Get-MpPreference` + `Get-MpComputerStatus` |
| 9 | Firewall Profiles | `Get-NetFirewallProfile` |
| 10 | Services | `Get-Service` (36 baseline-relevant services) |
| 11 | BitLocker Status | `Get-BitLockerVolume` |
| 12 | Credential Guard / VBS | WMI `Win32_DeviceGuard` + Registry |
| 13 | Windows Update History | `Get-HotFix` |
| 14 | Driver Inventory | `Win32_PnPSignedDriver` |
| 15 | Startup Performance | Diagnostics-Performance Event 100 |
| 16 | Scheduled Tasks | `Get-ScheduledTask` |
| 17 | SMB Configuration | `Get-SmbServer/ClientConfiguration` |
| 18 | TLS Configuration | SCHANNEL registry (SSL 2.0–TLS 1.3) |
| 19 | PowerShell Configuration | Script block logging, transcription, CLM |
| 20 | WinRM Configuration | Registry + `winrm get` |
| 21 | Event Log Metadata | Log sizes, retention, record counts |
| 22 | Security Event Collection | 13 query groups across Security/System/Application logs |

### Collector Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-OutputPath` | `.\<host>_baseline_<ts>.json` | Output file path |
| `-LookbackDays` | 30 | Event log query lookback window |
| `-MaxEventsPerQuery` | 10000 | Cap per event query group |
| `-SkipEventCollection` | `$false` | Skip Area 22 entirely (~30s total) |
| `-EventSummaryOnly` | `$false` | Counts + top-N stats only |
| `-Quiet` | `$false` | Suppress console output |

### Join Type Awareness

The collector auto-detects the device join type and adjusts behavior:

| Join Type | gpresult | MDM PolicyManager | Registry Paths |
|-----------|----------|-------------------|----------------|
| Entra ID (Intune) | Skipped (no DC) | Full scan | CSP paths |
| Domain-joined (GPO) | With 90s timeout | Skipped | Policy paths |
| Hybrid (both) | With 90s timeout | Full scan | Both |
| Workgroup | Skipped | Skipped | Local policy |

## GUI Tabs

| Tab | Purpose |
|-----|---------|
| **Dashboard** | Overall score, category cards, system info, maturity dimensions |
| **Baseline** | Per-check comparison: expected vs actual, grouped by category |
| **Findings** | Filtered view (Fail + Warning) with sort/filter by severity, category, effort |
| **Report** | Executive summary preview with RTF clipboard copy, HTML/CSV export |
| **Settings** | Theme, preferences, assessor name, baseline version |

## Check Categories (263 checks)

| Category | Prefix | Count | Scope |
|----------|--------|-------|-------|
| Security Configuration | SEC | 87 | Security options, services, user rights, SmartScreen, RDP, WinRM, PowerShell |
| Defender & Endpoint Security | DEF | 42 | Defender settings, ASR rules, VBS, Credential Guard |
| Authentication & Credentials | AUTH | 32 | Password policy, lockout, Kerberos, LAPS, NTLM |
| Monitoring & Audit | MON | 50 | Audit policy, event log sizing, audit gap detection |
| Network Security | NET | 35 | Firewall, SMB, TLS/SCHANNEL, LLMNR/NetBIOS |
| Data Protection | DATA | 24 | BitLocker, encryption, privacy, removable media |
| User Account Control | UAC | 11 | UAC settings, elevation prompts, admin approval |
| Operations & Health | OPS | 20 | Updates, drivers, services, tasks |

## Scoring Model

**Weighted Risk Score** (0–100):

```
Score = Σ(points × weight) / Σ(weight)
```

| Status | Points | Severity | Weight |
|--------|--------|----------|--------|
| Pass | 100 | Critical | 5× |
| Warning | 50 | High | 4× |
| Fail | 0 | Medium | 3× |
| Deferred | 0 | Low | 2× |
| Accepted Risk | — | — | Excluded |
| N/A | — | — | Excluded |

**Baseline Compliance %**: Flat pass/total ratio across all assessed checks.

## Governance Actions

Each failing check supports one of four governance states:

| Action | Icon | Effect |
|--------|------|--------|
| **Remediate** | ✓ green | Marked for remediation (counted toward projected score) |
| **Accept Risk** | shield amber | Excluded from scoring with mandatory justification |
| **N/A** | ○ gray | Not applicable to this environment |
| **Defer** | clock blue | Acknowledged but deferred — still counts as Fail |

## Executive Summary & RTF Export

The Report tab includes an **Executive Summary** generator with one-click clipboard copy:

- **Plain text**: Structured 8-section summary (Device Info, Scores, Category Breakdown, Key Passes, Failures by severity, Quick Wins, Governance Overrides, Methodology)
- **Rich RTF**: Professional formatted report with color-coded tables, severity badges, category score cards — pastes directly into Word, Outlook, or OneNote with full formatting
- **Dual clipboard**: `DataObject` carries both RTF and UnicodeText — rich apps get formatted output, plain editors get clean text

## Check Origins

Each check is tagged with its origin for traceability:

| Badge | Color | Source |
|-------|-------|--------|
| SCT | Blue | Microsoft Security Compliance Toolkit (GPO baselines) |
| INTUNE | Teal | Microsoft Intune Security Baseline |
| OPS | Gray | Operational health checks (ODA-inspired) |

## Prerequisites

### Collection Script
- PowerShell 5.1+ (ships with Windows 10/11)
- Local administrator rights
- No external modules

### BaselinePilot GUI
- PowerShell 5.1+ with WPF (PresentationFramework)
- .NET Framework 4.7.2+ (ships with Windows 10 1803+)
- No external modules

## File Structure

```
BaselineAssessor/
├── BaselinePilot.ps1              # WPF GUI application (~4500 lines)
├── BaselinePilot_UI.xaml          # WPF XAML layout
├── Invoke-BaselineCollection.ps1  # Headless data collector (22 areas)
├── checks.json                    # 263 check definitions
├── csp_metadata.json              # CSP metadata (descriptions, allowed values)
├── admx_metadata.json             # ADMX policy metadata
├── Launch_BaselinePilot.bat       # Batch launcher
├── PLAN.md                        # Detailed design document
├── README.md                      # This file
├── assessments/                   # Saved assessment JSON files
├── reports/                       # Generated HTML reports
└── templates/                     # Report templates
```

## Design Document

See [PLAN.md](PLAN.md) for the full design document covering architecture, check schema, scoring model, enrichment sources, event collection design, and implementation phases.
