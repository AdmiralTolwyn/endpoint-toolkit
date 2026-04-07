# AVD Assessor

**Assess Azure Virtual Desktop environments against CAF, Well-Architected Framework, and Landing Zone Accelerator best practices.**

AVD Assessor is a PowerShell/WPF desktop application that combines automated Azure subscription discovery with a workshop-friendly manual checklist — 145 checks across 10 categories — to produce scored readiness reports with maturity dimensions, category breakdowns, and exportable HTML/CSV/JSON deliverables.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![WPF](https://img.shields.io/badge/GUI-WPF-blueviolet)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)

---

## Why AVD Assessor?

Running an AVD readiness review typically involves spreadsheets, tribal knowledge, and hours of manual Azure portal checks. Findings are inconsistent between assessors, scoring is subjective, and reports are created from scratch each time. AVD Assessor solves this by:

- **Standardizing the framework**: 145 checks derived from Microsoft CAF, WAF, and LZA guidance, each with severity, weight, effort estimate, and documentation reference
- **Automating what can be automated**: A standalone discovery script scans Azure subscriptions and evaluates 90+ checks automatically — networking rules, VM configurations, scaling plans, storage security, monitoring, and more
- **Supporting the workshop**: A WPF GUI for interactive walkthroughs where the assessor and customer review checks together, add notes, and set statuses in real time
- **Scoring objectively**: Weighted category scores, an overall score, and a six-dimension maturity model (Initial → Optimized) provide a clear picture of readiness
- **Producing deliverables**: One-click export to dark/light HTML reports, CSV data dumps, or JSON snapshots for programmatic consumption

---

## Features

### Assessment Framework
- **145 checks** across 10 categories: Identity & Access, Networking, Session Hosts, FSLogix & Profiles, Security, Monitoring, BCDR, Governance & Cost, Application Delivery, Landing Zone
- Each check carries severity (Critical/High/Medium/Low), weight (1–5), effort estimate (Quick Win/Some Effort/Major Effort), and a Microsoft documentation URL
- Checks are defined in `checks.json` — extensible without code changes
- Check types: ~60 automated (discovery-backed) + ~80 manual (workshop review)

### Automated Discovery
- Standalone `Invoke-AvdDiscovery.ps1` script runs against one or more Azure subscriptions
- Evaluates 90+ checks automatically via Azure PowerShell modules
- Discovers host pools, session hosts (VM metadata, boot diagnostics, disk encryption, agent versions), application groups, workspaces, scaling plans, VNets, NSGs, storage accounts, Key Vaults, policies, alerts, quotas, budgets, reserved instances, orphaned resources
- Outputs a structured JSON that can be imported into the GUI for hybrid assessment
- Supports multi-subscription scanning, custom output paths, and `‑SkipLogin` for existing Az contexts

### Scoring & Maturity Model
- **Status scoring**: Pass = 100%, Warning = 50%, Fail = 0%, N/A and Not Assessed are excluded
- **Category score**: Weighted average of assessed checks within each category
- **Overall score**: Weighted average across all categories
- **Six maturity dimensions**: Security & Identity, Operations & Hosts, Networking, Resiliency & BCDR, Profiles & Storage, Monitoring
- **Five maturity levels**: Initial (0–34), Developing (35–54), Defined (55–74), Managed (75–89), Optimized (90–100)
- **Composite maturity**: Weighted score across all dimensions with a single maturity label

### Dashboard
- Real-time overall score and maturity level with color-coded indicator
- Per-category score bars with progress fill and percentage labels
- Score cards for each category showing pass/fail/warning count breakdown
- Six-dimension maturity radar (text-based) with individual dimension levels
- Assessment progress indicator (assessed vs. total checks)

### Report Generation
- **HTML export**: Full styled report with category sections, check tables, scoring summary, and maturity overview — supports dark and light themes, print-friendly
- **CSV export**: Flat data dump of all checks with status, notes, severity, category, and evidence
- **JSON export**: Complete assessment state including scores, maturity, timestamps, and discovery context
- **Report preview**: Live preview tab before exporting

### Assessment Management
- **Auto-save**: Rolling backups every 60 seconds (configurable) when assessment is dirty
- **Manual save**: Named saves using customer name, stored in `assessments/` folder
- **Rolling backups**: Up to 10 versioned backups in `_backups/` with automatic cleanup
- **Import**: Load saved assessments or discovery JSONs via Import Discovery / Assessment
- **Dirty tracking**: Visual indicator when unsaved changes exist, with confirmation on close/overwrite
- **Crash recovery**: Resume from latest auto-save backup

### Findings View
- Filtered view showing only checks with Fail, Warning, or Error status
- Grouped by category with severity badges
- Inline notes and recommendation text
- Quick navigation to assessment tab for status changes

### Theme System
- Dark and light themes with 40+ color keys (backgrounds, accents, borders, text, status indicators)
- Theme toggle button in the title bar (sun/moon icon)
- Animation toggle for reduced motion preference
- Theme persisted in user preferences

### Activity Log
- Collapsible bottom panel with timestamped log entries
- Color-coded by level: INFO, DEBUG, WARN, ERROR, SUCCESS
- Disk log file at `$env:TEMP\AvdAssessor_debug.log`
- DEBUG-level messages visible when debug overlay is enabled

### Achievements (20)
Gamification layer that rewards consistent usage patterns:

| Achievement | Trigger |
|---|---|
| First Steps | Complete first assessment |
| Repeat Auditor | Complete 5 assessments |
| Assessment Pro | Complete 10 assessments |
| Explorer | Import first discovery file |
| Full Sweep | Assess all checks in a single run |
| Flawless | Score 100% pass in any category |
| Well Managed | Reach Managed maturity level (75%+) |
| Peak Performance | Reach Optimized maturity (90%+) |
| Reporter | Export first HTML report |
| Data Wrangler | Export a CSV report |
| Night Owl | Run assessment between 00:00–05:00 |
| Early Bird | Run assessment between 05:00–07:00 |
| Weekend Warrior | Save assessment on a weekend |
| Chameleon | Toggle theme for the first time |
| Note Taker | Add notes to 5 or more checks |
| Scope Master | Exclude a check from scoring |
| Zero Critical | No critical-severity failures after 10+ assessments |
| Half Way | Pass 50 checks in a single assessment |
| Century Club | Pass 100 checks in a single assessment |
| Speed Demon | Complete assessment in under 5 minutes |

---

## Quick Start

### Option 1: GUI Tool (Workshop Mode)

Double-click **`Launch_AvdAssessor.bat`** or run:

```powershell
.\Launch_AvdAssessor.bat
```

The launcher auto-detects PowerShell 7 and falls back to Windows PowerShell 5.1.

### Option 2: Discovery Script (Automated Scan)

```powershell
# Interactive login — scans current subscription
.\Invoke-AvdDiscovery.ps1

# Specific subscription
.\Invoke-AvdDiscovery.ps1 -SubscriptionId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

# Multiple subscriptions, custom output
.\Invoke-AvdDiscovery.ps1 -SubscriptionId @("sub1", "sub2") -OutputPath "C:\temp\discovery.json"

# Reuse existing Az context (no login prompt)
.\Invoke-AvdDiscovery.ps1 -SkipLogin
```

Then import the discovery JSON into the GUI via **Import Discovery / Assessment**.

---

## Prerequisites

### GUI Tool

| Requirement | Details |
|---|---|
| OS | Windows 10/11 or Windows Server 2016+ |
| PowerShell | 5.1 (built-in) or 7+ |
| .NET | WPF support (included with Windows) |
| Azure modules | Not required for manual-only assessments |

### Discovery Script

| Requirement | Details |
|---|---|
| PowerShell | 5.1 or 7+ |
| Az.Accounts | Authentication and subscription management |
| Az.Compute | VM metadata, disk, and NIC details |
| Az.DesktopVirtualization | Host pools, session hosts, app groups, workspaces, scaling plans |
| Az.Storage | Storage account security and configuration checks |
| Az.KeyVault | Key Vault existence and private endpoint checks |
| Az.Monitor | Diagnostic settings, alerts, and monitoring checks |
| Azure RBAC | **Reader** role on target subscription(s) |

```powershell
Install-Module Az.Accounts, Az.Compute, Az.DesktopVirtualization, Az.Storage, Az.KeyVault, Az.Monitor -Scope CurrentUser
```

---

## Assessment Categories

| Category | Checks | ID Prefix | Sources | Key Areas |
|---|---|---|---|---|
| Session Hosts | 20 | SH | WAF, CAF | VM sizing, images, agents, boot diag, disk encryption, power states |
| Networking | 14 | NET | WAF, LZA | DNS, NSG rules, UDR, NAT Gateway, peering, subnet capacity |
| Monitoring | 14 | MON | WAF, CAF | Diagnostic settings, alerts, Log Analytics, Connection Monitor |
| Security | 12 | SEC | SEC, WAF | MFA, conditional access, endpoint protection, TLS, clipboard policy |
| Identity & Access | 9 | IAM | AVD, WAF | RBAC, Entra ID Join, SSO, service principals, admin isolation |
| Application Delivery | 9 | APP | WAF | MSIX, RemoteApp, App Attach, app layering, updates |
| Governance & Cost | 8 | GOV | CAF, WAF | Scaling plans, cost tagging, budgets, reserved instances, quotas |
| FSLogix & Profiles | 7 | PROF | WAF, FSL | Profile containers, VHD locations, exclusions, Azure Files SMB |
| BCDR | 7 | BCDR | WAF | Multi-region, backup, scaling plan schedules, disaster recovery |
| Landing Zone | 5 | LZ | LZA | Resource organization, naming, tagging, policy, hub-spoke |

### Check Metadata

Each check definition in `checks.json` includes:

| Field | Description |
|---|---|
| `id` | Unique identifier (e.g. `NET-003`) |
| `category` | Assessment category |
| `name` | Short human-readable label |
| `description` | What the check evaluates |
| `severity` | Critical, High, Medium, or Low |
| `weight` | Scoring weight (1–5) |
| `type` | `Auto` (discovery-backed) or `Manual` (workshop review) |
| `effort` | Remediation effort: Quick Win, Some Effort, or Major Effort |
| `recommendation` | Remediation guidance |
| `reference` | URL to Microsoft documentation |

### Severity Distribution

| Severity | Approximate Count | Weight Range |
|---|---|---|
| Critical | ~4 | 5 |
| High | ~50 | 4–5 |
| Medium | ~70 | 2–4 |
| Low | ~16 | 1–2 |

---

## Scoring Methodology

### Status Scoring

| Status | Score | Description |
|---|---|---|
| Pass | 100% | Check fully satisfied |
| Warning | 50% | Partially met or acceptable risk documented |
| Fail | 0% | Not met — remediation recommended |
| N/A | Excluded | Not applicable to this environment |
| Not Assessed | Excluded | Not yet evaluated |

### Calculation

**Category Score** = Σ (check_score × check_weight) / Σ (check_weight) for assessed checks in that category.

**Overall Score** = Σ (category_score × category_weight) / Σ (category_weight) across all categories with assessed checks.

### Maturity Dimensions

Checks are mapped to six maturity dimensions by their ID prefix:

| Dimension | Prefixes | Focus |
|---|---|---|
| Security & Identity | SEC-, IAM- | Authentication, authorization, endpoint protection |
| Operations & Hosts | OPS-, GOV-, SH- | VM management, scaling, cost governance |
| Networking | NET- | Connectivity, segmentation, DNS, NSG |
| Resiliency & BCDR | BCDR- | Disaster recovery, backup, multi-region |
| Profiles & Storage | PROF- | FSLogix, Azure Files, storage configuration |
| Monitoring | MON- | Diagnostics, alerts, Log Analytics |

### Maturity Levels

| Level | Score Range | Description |
|---|---|---|
| Initial | 0–34 | Ad-hoc processes, major gaps |
| Developing | 35–54 | Some practices in place, inconsistent |
| Defined | 55–74 | Standardized processes, most areas covered |
| Managed | 75–89 | Measured and controlled, proactive management |
| Optimized | 90–100 | Continuous improvement, industry-leading practices |

---

## Workflow

### Typical Workshop Flow

```
Pre-workshop         Workshop                Assessment            Deliverable
┌─────────────┐     ┌──────────────────┐    ┌──────────────────┐  ┌──────────────┐
│ Customer or  │     │ Launch GUI       │    │ Walk through     │  │ Export HTML   │
│ assessor     │────▶│ Import discovery │───▶│ manual checks    │─▶│ report and    │
│ runs         │     │ Review automated │    │ with customer    │  │ share         │
│ discovery    │     │ findings         │    │ Add notes per    │  │              │
│ script       │     │                  │    │ check            │  │              │
└─────────────┘     └──────────────────┘    └──────────────────┘  └──────────────┘
```

1. **Pre-workshop**: Run `Invoke-AvdDiscovery.ps1` against the customer's subscription — can be executed by the customer with Reader access
2. **Workshop**: Launch the GUI, import the discovery JSON, review auto-evaluated findings on the Dashboard
3. **Assessment**: Navigate to the Assessment tab, walk through manual checks with the customer, set status (Pass/Fail/Warning/N/A) and add notes
4. **Findings**: Switch to Findings tab to review all failures and warnings grouped by category
5. **Report**: Preview the HTML report, then export to HTML, CSV, or JSON and share with the customer

### Offline Assessment

The GUI works fully offline for manual-only assessments — no Azure connection required. All 145 checks can be evaluated manually based on customer documentation and interview.

### Hybrid Mode

Import discovery results to pre-populate automated check statuses, then overlay manual assessment on top. The GUI tracks which checks came from discovery (Auto) vs. manual review.

---

## UI Tabs

### Dashboard
Overall score, maturity level, per-category breakdown bars, dimension scores, and assessment progress. Refreshes in real time as checks are evaluated.

### Assessment
Full checklist of 145 checks grouped by category. Each check row shows severity badge, status dropdown, notes field, weight, and reference link. Supports filtering by category, status, severity, and text search.

### Findings
Filtered view of Fail, Warning, and Error checks only. Grouped by category with severity indicators and inline recommendations. 

### Report
Live HTML report preview with export buttons for HTML (dark/light theme), CSV, and JSON formats. The HTML report includes scoring summary, maturity dimensions, per-category sections with check tables, and assessment metadata.

### Settings
Auto-save interval, backup management, cache purge, debug overlay toggle, animation toggle, and reset to defaults.

---

## Discovery Script Details

`Invoke-AvdDiscovery.ps1` (v0.2.0) runs as a standalone script that scans Azure subscriptions and produces a structured JSON file.

### Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-SubscriptionId` | String or String[] | Current Az context | One or more subscription IDs to scan |
| `-OutputPath` | String | `assessments/discovery_<timestamp>.json` | Path for output JSON |
| `-SkipLogin` | Switch | `$false` | Skip interactive login, use existing Az context |

### Discovery Sections

The script scans these resource types and generates automated check results:

| Section | Resources Discovered | Checks Generated |
|---|---|---|
| Host Pools | Pool config, drain mode, load balancing, tags | Governance, BCDR |
| Session Hosts | VMs, boot diag, disks, NICs, agents, power state, image age | Session Host checks |
| Application Groups | Type, assignments, host pool reference | Identity & Access |
| Workspaces | Name, location, app group references | Configuration |
| Scaling Plans | Schedules, host pool assignments, time zones | Governance, BCDR |
| Virtual Networks | DNS, subnets, NSGs, UDRs, NAT Gateways, peerings | Networking |
| NSG Rules | RDP exposure, AVD outbound, SSH access | Security, Networking |
| Storage Accounts | Private endpoints, HTTPS, TLS, SMB, Kerberos auth | Profiles, Security |
| Key Vaults | Existence, private endpoints | Security |
| Monitoring | Diagnostic settings, alert rules, Network Watcher | Monitoring |
| Governance | Policies, budgets, quotas, reserved instances, tags | Governance, Cost |
| Orphans | Unattached disks, disconnected NICs | Operations |

### Output Format

```json
{
  "Metadata": {
    "Timestamp": "2026-03-27T14:30:00Z",
    "ScriptVersion": "0.2.0",
    "Subscriptions": ["sub-id-1"],
    "Duration": "00:02:45"
  },
  "Inventory": {
    "HostPools": [...],
    "SessionHosts": [...],
    "ApplicationGroups": [...],
    "Workspaces": [...],
    "ScalingPlans": [...],
    "VirtualNetworks": [...],
    "StorageAccounts": [...],
    "KeyVaults": [...]
  },
  "Checks": [
    {
      "Id": "NET-003",
      "Category": "Networking",
      "Name": "NSG Blocks Inbound RDP",
      "Status": "Fail",
      "Severity": "Critical",
      "Details": "NSG 'my-nsg' allows inbound RDP from Any",
      "Recommendation": "Remove or restrict inbound RDP rules.",
      "Reference": "https://learn.microsoft.com/...",
      "Source": "Automated",
      "Timestamp": "2026-03-27T14:30:15Z"
    }
  ],
  "Maturity": {
    "Dimensions": { "Security": 72, "Operations": 81, ... },
    "CompositeScore": 68,
    "MaturityLevel": "Defined"
  },
  "Errors": []
}
```

---

## File Structure

```
AvdAssessor/
├── AvdAssessor.ps1              # Main GUI application (41 functions, ~4900 lines)
├── AvdAssessor_UI.xaml          # WPF layout, styles, and resource dictionaries
├── Invoke-AvdDiscovery.ps1     # Standalone discovery script (90+ automated checks)
├── Launch_AvdAssessor.bat      # Launcher (auto-detects PS7, falls back to PS5.1)
├── checks.json                 # 145 check definitions with metadata
├── assessments/                # Saved assessment and discovery JSON files
│   └── discovery_*.json        # Discovery scan outputs
├── reports/                    # Exported HTML/CSV reports
├── _backups/                   # Rolling auto-save backups (max 10)
└── README.md                   # This file
```

---

## Architecture

### Code Organization

| Section | Purpose | Key Functions |
|---|---|---|
| Pre-load & Module Path | Version, module path sanitization | — |
| Theme Palettes | 40+ color keys for dark/light modes | — |
| Thread Sync & Background Work | Runspace pool with OnComplete callbacks | `Start-BackgroundWork` |
| XAML & Element Binding | WPF window and control references | — |
| Logging | Multi-destination debug logging | `Write-DebugLog` |
| Theme Engine | Dynamic brush/color application | — |
| Toast Notifications | Themed popup messages with auto-dismiss | `Show-Toast` |
| Themed Dialog | Modal confirm/cancel with icon support | `Show-ThemedDialog` |
| Tab Switching | Animated tab transitions with fade | `Switch-Tab`, `Invoke-TabFade` |
| Assessment Data Model | Check definitions, scoring, maturity | `Get-CategoryScore`, `Get-OverallScore`, `Get-DimensionScore`, `Get-MaturityLevel`, `Get-CompositeMaturityScore` |
| Assessment Logic | Reset, import, sync, dirty tracking | `Reset-Assessment`, `Import-DiscoveryJson`, `Sync-CheckDefinitions` |
| UI Rendering | Dashboard, checklist, findings views | `Update-Dashboard`, `Render-AssessmentChecks`, `Render-Findings`, `Update-Progress` |
| HTML Report Generation | Styled HTML with maturity + scores | `Build-HtmlReport`, `Export-HtmlReport` |
| Export | CSV and JSON export | `Export-CsvReport`, `Export-JsonAssessment` |
| Save/Load Assessment | Auto-save + manual save/load cycle | `AutoSave-Assessment`, `Save-Assessment`, `Load-Assessment` |
| Saved Assessments List | Assessment browser with refresh | `Refresh-AssessmentList` |
| User Preferences | Theme, window position, settings | `Save-UserPrefs`, `Load-UserPrefs` |
| Achievements | 20 achievements with badge UI | `Unlock-Achievement`, `Check-AssessmentAchievements` |

### Data Flow

```
checks.json ──────────────────────────────┐
                                          ▼
                                   ┌──────────────┐
Invoke-AvdDiscovery.ps1 ─────────▶│  Import      │
  (Azure subscription scan)       │  Discovery   │
  ┌─ Host Pools                   │  JSON        │
  ├─ Session Hosts (VM details)   └──────┬───────┘
  ├─ App Groups & Workspaces             │
  ├─ Scaling Plans                       ▼
  ├─ VNets, NSGs, Storage       ┌────────────────┐
  ├─ Key Vaults, Monitoring     │  Sync Check    │
  └─ Policies, Budgets, Quotas  │  Definitions   │──▶ 145 checks loaded
                                └────────┬───────┘
                                         │
                     ┌───────────────────┼───────────────────┐
                     ▼                   ▼                   ▼
              ┌──────────┐      ┌──────────────┐    ┌──────────────┐
              │ Dashboard │      │  Assessment  │    │   Findings   │
              │ Scores &  │      │  Checklist   │    │   Filtered   │
              │ Maturity  │◀────▶│  Status +    │───▶│   Fail/Warn  │
              │ Bars      │      │  Notes       │    │   View       │
              └──────────┘      └──────┬───────┘    └──────────────┘
                                       │
                     ┌─────────────────┼─────────────────┐
                     ▼                 ▼                  ▼
              ┌──────────┐     ┌─────────────┐    ┌──────────────┐
              │ Auto-Save │     │ HTML Report │    │  CSV / JSON  │
              │ _backups/ │     │ Preview &   │    │  Export      │
              │ (60s)     │     │ Export      │    │              │
              └──────────┘     └─────────────┘    └──────────────┘
```

### Background Work Pattern

Long-running operations (discovery import, report generation) run in background runspaces to keep the WPF UI responsive:

1. `Start-BackgroundWork` launches a PowerShell runspace with the task scriptblock
2. A DispatcherTimer polls the runspace every 50ms for completion
3. On completion, the `OnComplete` callback runs on the UI thread with access to all WPF controls
4. Exceptions are caught and surfaced via `Show-Toast` with ERROR level

---

## Troubleshooting

### Discovery Script Fails to Connect

```
Error: No Azure context found
```

Run `Connect-AzAccount` manually, then retry with `-SkipLogin`. Ensure the correct subscription is selected with `Set-AzContext -SubscriptionId <id>`.

### Missing Azure Modules

```
Error: The term 'Get-AzWvdHostPool' is not recognized
```

Install the required modules:

```powershell
Install-Module Az.DesktopVirtualization -Scope CurrentUser -Force
```

### Discovery Import Shows No Automated Checks

Ensure the discovery JSON was generated by `Invoke-AvdDiscovery.ps1` v0.2.0+. Older formats may not include the `Checks` array. Re-run the discovery script.

### GUI Appears Blank or Frozen

WPF rendering requires a desktop session. If running over RDP with restricted GPU settings, ensure `Hardware Graphics Adapter` is not disabled. On slow connections, disable animations in Settings.

### Auto-Save Not Working

Check that the `_backups/` directory exists and is writable. The auto-save timer requires an active dirty flag — if no checks have been modified since last save, no backup is created.

### Scores Don't Match Expected

- Only **assessed** checks (Pass/Fail/Warning) are included in scoring
- N/A and Not Assessed checks are excluded from both numerator and denominator
- Weights amplify the impact of higher-weighted checks — a single weight-5 Critical failure can significantly lower a category score

---

## Logging

| Destination | Details |
|---|---|
| PowerShell console | All messages via `Write-Host` (DarkGray) |
| Activity Log panel | Color-coded entries in the collapsible bottom panel |
| Disk log | `$env:TEMP\AvdAssessor_debug.log` |
| Debug overlay | Enable in Settings to show DEBUG-level messages |

Log format: `[HH:mm:ss.fff] [LEVEL] Message`

Levels: INFO, DEBUG, WARN, ERROR, SUCCESS

---

## Author

**Anton Romanyuk**

## License

Internal use.
