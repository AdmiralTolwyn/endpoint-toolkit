# AVD Assessor

**Assess Azure Virtual Desktop environments against CAF, Well-Architected Framework, and Landing Zone Accelerator best practices.**

AVD Assessor is a PowerShell/WPF desktop application that combines automated Azure subscription discovery with a workshop-friendly manual checklist — 183 checks across 11 categories — to produce scored readiness reports with maturity dimensions, category breakdowns, and exportable HTML/CSV/JSON deliverables.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![WPF](https://img.shields.io/badge/GUI-WPF-blueviolet)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)

---

## Why AVD Assessor?

Running an AVD readiness review typically involves spreadsheets, tribal knowledge, and hours of manual Azure portal checks. Findings are inconsistent between assessors, scoring is subjective, and reports are created from scratch each time. AVD Assessor solves this by:

- **Standardizing the framework**: 183 checks derived from Microsoft CAF, WAF, and LZA guidance, each with severity, weight, effort estimate, and documentation reference
- **Automating what can be automated**: A standalone discovery script scans Azure subscriptions and evaluates 121 checks automatically — networking rules, VM configurations, scaling plans, storage security, monitoring, Conditional Access / MFA (via Microsoft Graph), AVD Insights latency and telemetry (via Log Analytics KQL), Intune baselines / Credential Guard / VBS (via Intune Graph), optional in-guest FSLogix inspection, and more
- **Supporting the workshop**: A WPF GUI for interactive walkthroughs where the assessor and customer review checks together, add notes, and set statuses in real time
- **Scoring objectively**: Weighted category scores, an overall score, and a six-dimension maturity model (Initial → Optimized) provide a clear picture of readiness
- **Producing deliverables**: One-click export to dark/light HTML reports, CSV data dumps, or JSON snapshots for programmatic consumption

---

## Features

### Assessment Framework
- **183 checks** across 11 categories: Identity & Access, Networking, Session Hosts, FSLogix & Profiles, Security, Monitoring, BCDR, Governance & Cost, Application Delivery, Operations, Landing Zone
- Each check carries severity (Critical/High/Medium/Low), weight (1–5), effort estimate (Quick Win/Some Effort/Major Effort), and a Microsoft documentation URL
- Checks are defined in `checks.json` — extensible without code changes
- Check types: 121 automated (discovery-backed) + 62 manual (workshop review)

### Automated Discovery
- Standalone `Invoke-AvdDiscovery.ps1` script runs against one or more Azure subscriptions
- Evaluates 121 checks automatically via Azure PowerShell modules (including Conditional Access / MFA / passwordless via Microsoft Graph, AVD Insights KQL via Log Analytics, and Intune configuration via Intune Graph)
- Discovers host pools, session hosts (VM metadata, boot diagnostics, disk encryption, agent versions), application groups, workspaces, scaling plans, VNets, NSGs, storage accounts, Key Vaults, policies, alerts, quotas, budgets, reserved instances, orphaned resources
- Outputs a structured JSON that can be imported into the GUI for hybrid assessment
- Supports multi-subscription scanning, custom output paths, and `‑SkipLogin` for existing Az contexts

### Scoring & Maturity Model
- **Status scoring**: Pass = 100%, Warning = 50%, Fail = 0%. N/A is excluded from both the numerator and denominator. Not Assessed scores 0% but remains in the denominator — partial completion shows honest scores. Excluded checks are removed entirely from scoring
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

# Include opt-in in-guest FSLogix inspection (Run Command on up to 3 running hosts per pool)
.\Invoke-AvdDiscovery.ps1 -IncludeGuestChecks
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
| Az.DesktopVirtualization | Host pools, session hosts, app groups, workspaces, scaling plans |
| Az.Resources | Resource groups, tags, locks, generic resource queries |
| Az.Compute | VM metadata, disk, and NIC details |
| Az.Network | NSGs, VNets, peerings, NAT Gateways, route tables |
| Az.PrivateDns | Private DNS zone linkage checks |
| Az.Monitor | Diagnostic settings, alerts, and monitoring checks |
| Az.Storage | Storage account security and configuration checks |
| Az.KeyVault | Key Vault existence and private endpoint checks |
| Az.Security | Defender for Cloud, secure score, JIT, regulatory compliance, Defender for Servers plan (TVM) checks |
| Az.OperationalInsights | **Optional** — Log Analytics KQL checks (AVD Insights latency, Perf/Event collection, storage IOPS, profile load times). When absent, those six checks report `Error` instead of blocking the run |
| Azure RBAC | **Reader** role on target subscription(s). For the optional Log Analytics KQL checks, **Log Analytics Reader** on the workspace(s) receiving AVD diagnostics. For `-IncludeGuestChecks`, the **Microsoft.Compute/virtualMachines/runCommand/action** permission (e.g. Virtual Machine Contributor) on the session-host VMs |
| Microsoft Graph (Identity) | **Policy.Read.All** (Conditional Access / MFA / token-protection checks) and, optionally, **AuditLog.Read.All** (passwordless registration). The Graph token is acquired from the existing Az login — no extra module is required. Identity checks degrade to `Error` (not a crash) when these permissions are absent |
| Microsoft Graph (Intune) | **Optional** — `DeviceManagementConfiguration.Read.All` and `DeviceManagementManagedDevices.Read.All` for the Intune checks (security baselines, compliance/drift, patch posture, application control, Credential Guard, VBS/HVCI, OneDrive KFM, FSLogix AV exclusions). Each Intune check degrades to `Error` listing the missing scope when absent |

```powershell
# Required modules
Install-Module Az.Accounts, Az.DesktopVirtualization, Az.Resources, Az.Compute, Az.Network, Az.PrivateDns, Az.Monitor, Az.Storage, Az.KeyVault, Az.Security -Scope CurrentUser

# Optional module for the Log Analytics KQL checks
Install-Module Az.OperationalInsights -Scope CurrentUser
```

> **Two optional permission tiers.** Beyond the base `Reader` + `Policy.Read.All`, Phase 4 adds two opt-in signal sources that never hard-fail the run:
>
> - **Log Analytics Reader** on the workspace(s) that receive AVD diagnostics — powers the AVD Insights KQL checks (NET-008 latency, MON-002 Insights data flow, MON-008 Perf, MON-009 Events, MON-010 storage IOPS, PROF-010 profile load times). The workspace IDs are reused from the diagnostic settings already discovered. Requires the optional `Az.OperationalInsights` module; when the module or data is missing, these checks report `Error`/`Warning` rather than blocking.
> - **Intune Graph scopes** (`DeviceManagementConfiguration.Read.All`, `DeviceManagementManagedDevices.Read.All`) on the same Az-login token — powers SH-028 (baselines), SH-014 (drift), SH-005 (patch), SEC-001 (application control), SEC-003 (Credential Guard), SEC-004 (VBS/HVCI), PROF-007 (OneDrive KFM), and PROF-019 (FSLogix AV exclusions). Each degrades to a per-check `Error` naming the missing scope.

> **Graph permissions**: the discovery script calls Microsoft Graph (Conditional Access policies, authentication method registration) using a token from your Az login. Grant the signed-in account `Policy.Read.All` for the MFA / Conditional Access / Windows Cloud Login / token-protection checks (IAM-002/003/010/011), and `AuditLog.Read.All` (plus an Entra ID P1/P2 license) for the passwordless-registration check (IAM-009). Without them, those identity checks report `Error` rather than failing the run.

> **`-IncludeGuestChecks` (opt-in, skipped by default)**: adds in-guest FSLogix inspection. When set, the script runs a single consolidated PowerShell script via `Invoke-AzVMRunCommand` against **up to 3 representative running session hosts per host pool** (a sampling cap to keep runtime bounded) to read `HKLM\SOFTWARE\FSLogix\Profiles` / `ODFC` (Enabled, VHDLocations, VolumeType, FlipFlopProfileDirectoryName, DeleteLocalProfileWhenVHDShouldApply, SizeInMBs, CCDLocations, ODFCEnabled) and the FSLogix agent version. This converts PROF-001 (installed), PROF-008 (Cloud Cache), PROF-009 (ODFC), PROF-012 (version), PROF-013 (VHDX), PROF-014 (FlipFlop), and PROF-015 (DeleteLocalProfileWhenVHDShouldApply) to automated. It is opt-in because it needs running VMs and the `runCommand` action, and adds runtime; when the switch is absent these seven checks honestly report `N/A` ("guest checks not enabled") rather than nothing.

---

## Assessment Categories

| Category | Checks | ID Prefix | Sources | Key Areas |
|---|---|---|---|---|
| FSLogix & Profiles | 27 | PROF | WAF, FSL | Profile containers, VHD locations, exclusions, Azure Files SMB, Kerberos AES readiness |
| Session Hosts | 29 | SH | WAF, CAF | VM sizing, images, agents, OS end-of-support, GPU config, region proximity, disk encryption, power states |
| Security | 24 | SEC | SEC, WAF | MFA, conditional access, endpoint protection, secure score, JIT, security baseline, TLS, clipboard policy |
| Networking | 21 | NET | WAF, LZA | DNS, NSG rules, UDR, NAT Gateway, private link, peering, subnet capacity |
| Governance & Cost | 17 | GOV | CAF, WAF | Scaling plans, cost tagging, budgets, resource locks, reserved instances, quotas |
| Monitoring | 16 | MON | WAF, CAF | Diagnostic settings, alerts, service health, SIEM/Sentinel, scaling-plan diagnostics, Log Analytics |
| BCDR | 12 | BCDR | WAF | Multi-region, image replication, DR capacity reservation, backup, disaster recovery |
| Identity & Access | 12 | IAM | AVD, WAF | Conditional Access / MFA, passwordless, token protection, Entra Kerberos, RBAC, SSO, Entra ID Join |
| Landing Zone | 10 | LZ | LZA | Resource organization, naming, tagging, policy, hub-spoke |
| Application Delivery | 10 | APP | WAF | App Attach, RemoteApp, Teams SlimCore migration, app layering, updates |
| Operations | 5 | OPS | CAF, WAF | Operational readiness, run-books, Windows App client migration, day-2 processes |

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

| Severity | Count | Weight Range |
|---|---|---|
| Critical | 4 | 5 |
| High | 58 | 4 |
| Medium | 92 | 3 |
| Low | 23 | 2 |

---

## Scoring Methodology

### Status Scoring

| Status | Score | Denominator | Description |
|---|---|---|---|
| Pass | 100% | Included | Check fully satisfied |
| Warning | 50% | Included | Partially met or acceptable risk documented |
| Fail | 0% | Included | Not met — remediation recommended |
| N/A | Excluded | Excluded | Not applicable to this environment — removed from both numerator and denominator |
| Not Assessed | 0% | Included | Not yet evaluated — scores 0 but stays in the denominator so partial completion shows honest, unrounded scores |
| Excluded | Excluded | Excluded | Check explicitly excluded from scoring; removed entirely |

### Calculation

**Category Score** = Σ (check_score × check_weight) / Σ (check_weight) for all non-N/A, non-Excluded checks in that category (Not Assessed checks contribute 0 to the numerator but their weight still counts in the denominator).

**Overall Score** = Σ (category_score × category_weight) / Σ (category_weight) across all categories with in-scope checks.

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

The GUI works fully offline for manual-only assessments — no Azure connection required. All 183 checks can be evaluated manually based on customer documentation and interview.

### Hybrid Mode

Import discovery results to pre-populate automated check statuses, then overlay manual assessment on top. The GUI tracks which checks came from discovery (Auto) vs. manual review.

---

## UI Tabs

### Dashboard
Overall score, maturity level, per-category breakdown bars, dimension scores, and assessment progress. Refreshes in real time as checks are evaluated.

### Assessment
Full checklist of 183 checks grouped by category. Each check row shows severity badge, status dropdown, notes field, weight, and reference link. Supports filtering by category, status, severity, and text search.

### Findings
Filtered view of Fail, Warning, and Error checks only. Grouped by category with severity indicators and inline recommendations. 

### Report
Live HTML report preview with export buttons for HTML (dark/light theme), CSV, and JSON formats. The HTML report includes scoring summary, maturity dimensions, per-category sections with check tables, and assessment metadata.

### Settings
Auto-save interval, backup management, cache purge, debug overlay toggle, animation toggle, and reset to defaults.

---

## Discovery Script Details

`Invoke-AvdDiscovery.ps1` (v0.6.0) runs as a standalone script that scans Azure subscriptions and produces a structured JSON file.

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
    "ScriptVersion": "0.6.0",
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
├── AvdAssessor.ps1              # Main GUI application (~6k lines)
├── AvdAssessor_UI.xaml          # WPF layout, styles, and resource dictionaries
├── Invoke-AvdDiscovery.ps1     # Standalone discovery script (121 automated checks)
├── Launch_AvdAssessor.bat      # Launcher (auto-detects PS7, falls back to PS5.1)
├── checks.json                 # 183 check definitions with metadata
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
  └─ Policies, Budgets, Quotas  │  Definitions   │──▶ 183 checks loaded
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

Ensure the discovery JSON was generated by `Invoke-AvdDiscovery.ps1` v0.6.0+. Older formats may not include the `Checks` array. Re-run the discovery script.

### GUI Appears Blank or Frozen

WPF rendering requires a desktop session. If running over RDP with restricted GPU settings, ensure `Hardware Graphics Adapter` is not disabled. On slow connections, disable animations in Settings.

### Auto-Save Not Working

Check that the `_backups/` directory exists and is writable. The auto-save timer requires an active dirty flag — if no checks have been modified since last save, no backup is created.

### Scores Don't Match Expected

- Pass/Fail/Warning checks are included in scoring at 100%/0%/50% respectively
- N/A checks are excluded from both numerator and denominator
- Not Assessed checks score 0% but **remain in the denominator** — an incomplete assessment will show a lower score than a fully-assessed one, by design
- Excluded checks are removed entirely from scoring
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

## Changelog

**0.6.0** — Phase 4 automation tier: 22 Manual→Auto conversions across three optional signal sources (catalog v1.3, now 121 Auto / 62 Manual). **Tier 1 — Log Analytics KQL** (optional `Az.OperationalInsights`, reusing workspace IDs from diagnostic settings): NET-008 latency (`MON-LATENCY-*`), MON-002 AVD Insights data flow (`MON-INSIGHTS-*`), MON-008 performance counters (`MON-PERF-*`), MON-009 event logs (`MON-EVENTS-*`), MON-010 storage IOPS (`MON-STORIOPS-*`), PROF-010 profile load times (`PROF-LOADTIME-*`); checks report `Error` when the module is missing and `Warning` when no data is flowing. **Tier 2 — Intune Graph** (reusing the Az-login Graph token; needs `DeviceManagementConfiguration.Read.All` / `DeviceManagementManagedDevices.Read.All`): SH-028 baselines (`SH-BASELINE`), SH-014 drift (`SH-DRIFT`), SH-005 patch (`SH-PATCH`), SEC-001 application control (`SEC-APPCTRL`), SEC-003 Credential Guard (`SEC-CREDGUARD`), SEC-004 VBS/HVCI (`SEC-VBS`), SEC-007 TVM via Defender for Servers plan (`SEC-TVM-*`), PROF-007 OneDrive KFM (`PROF-KFM`), PROF-019 FSLogix AV exclusions (`PROF-AVEXCL`); each degrades to a per-check `Error` naming the missing scope. **Tier 3 — opt-in `-IncludeGuestChecks`**: `Invoke-AzVMRunCommand` on up to 3 running hosts per pool reads FSLogix registry + agent version, converting PROF-001/008/009/012/013/014/015 (`PROF-INSTALLED/CCACHE/ODFC/VER/VHDX/FLIPFLOP/DELLOCAL-*`); these report `N/A` when the switch is absent. Every new external call is wrapped in try/catch → `Error` emit (never crashes).

**0.5.0** — Graph identity collection, 16 Manual→Auto conversions, 6 new checks. Added a Microsoft Graph section to discovery (Conditional Access / MFA / Windows Cloud Login / token protection / passwordless registration; degrades to `Error` when Graph permissions are missing), converted 16 ARM/Defender-discoverable checks from Manual to Auto (IAM-002/003/009/010, BCDR-006/011, GOV-007, SH-020, APP-002/004, MON-003/012, NET-005, SEC-005/013/021), implemented the orphaned SH-018 (region proximity), reclassified NET-011 to Manual (not ARM-discoverable), and added new Auto checks IAM-011 (Windows Cloud Login CA coverage), IAM-012 (Entra Kerberos), SH-029 (OS end-of-support), SH-030 (GPU config), SH-031 (personal-pool autoscale), and MON-017 (scaling-plan diagnostics). Also corrected 15 per-subscription mapping arms whose singleton IDs (suffixed with the sub short-id by the A-1 fix) no longer matched their exact-string arm patterns.

**July 2026** — Applied catalog and docs fixes from a full-surface audit (see `AUDIT.md`): reclassified several checks that can't be discovered via ARM to Manual, refreshed guidance for Windows Cloud Login, Teams SlimCore, App Attach, Session Host Update GA, RDP Shortpath, FSLogix/Kerberos AES readiness, corrected scoring semantics documentation, fixed dead reference URLs, normalized metadata, and added four new Manual checks (APP-010, OPS-006, PROF-026, PROF-027).

## Author

**Anton Romanyuk**

## License

Internal use.
