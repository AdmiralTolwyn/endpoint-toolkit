# Windows 365 Assessor

**Assess Windows 365 (Cloud PC) Enterprise & Flex tenants against Microsoft CAF, Well-Architected Framework, Landing Zone Accelerator, and Security best practices.**

Windows 365 Assessor is a PowerShell/WPF desktop application paired with a Microsoft Graph data collector. It combines automated tenant discovery with a workshop-friendly manual checklist — **132 checks across 12 categories** — to produce scored readiness reports with a six-dimension W365 maturity radar, category breakdowns, and exportable HTML / CSV / JSON deliverables.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![WPF](https://img.shields.io/badge/GUI-WPF-blueviolet)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)
![Graph](https://img.shields.io/badge/Microsoft%20Graph-v1.0%20%2B%20beta-orange)

---

## Why Windows 365 Assessor?

A Windows 365 readiness review typically involves several Microsoft Learn tabs, a spreadsheet of provisioning policies, and tribal knowledge about Flex (formerly Frontline), Reserve, DR Plus, ANC health, and Conditional Access scoping. Findings are inconsistent between assessors, scoring is subjective, and reports are written from scratch each time.

Windows 365 Assessor solves this by:

- **Standardizing the framework** — 132 checks derived from Microsoft Learn (W365), Well-Architected Framework, Cloud Adoption Framework, Landing Zone Accelerator, and dedicated security guidance. Every check carries severity, weight, effort, applicability per W365 edition, evidence-based best practice, business impact, concrete remediation, and a Microsoft Learn URL.
- **Automating what Graph can answer** — the discovery script consents to read-only Cloud PC scopes and evaluates **41 checks automatically** (inventory, edition mix, service-plan coverage, ANC health, image freshness, Cloud PC operational state, SSO, Conditional Access and MFA targeting, local admin, cross-region DR, restore-point frequency, Autopatch, grace period, domain-join consistency, Windows settings, Cloud Apps, audit activity, disk encryption, naming, locale, orphaned PCs, failed images, Defender/compliance/baseline posture, Endpoint Analytics, connection quality, update compliance, recommendations, resource performance, sign-in performance, right-sizing, inactive Cloud PCs, and high-impact admin actions in the last 30 days).
- **Supporting the workshop** — the GUI is built for live customer walkthroughs: the assessor and customer review remaining checks together, set status, capture notes, and watch the score and maturity radar update in real time.
- **Scoring objectively** — weighted category scores, an overall score, and a six-dimension W365 maturity model (Initial → Optimized) provide a defensible picture of readiness.
- **Producing deliverables** — one-click export to dark/light HTML reports (with a deduplicated Sources Appendix), CSV data dumps, or JSON snapshots for programmatic consumption.

---

## What's covered

The catalog is built directly from the official **Windows 365 Enterprise** Microsoft Learn documentation and includes preview features (e.g. Windows Cloud Keyboard Input Protection, Cloud PC Monitoring, AI-enabled Cloud PC, Cloud PC Configurations, Windows App Settings, Remote Connection Experience, Power Platform connector).

It explicitly covers the modern W365 surface that older assessment templates miss:

- **Flex (formerly Frontline)** — Cloud Apps (RemoteApp mode), User Experience Sync, concurrency buffer sizing, bulk reprovision cadence, dedicated patching strategy
- **Resilience** — Cross-Region DR, Disaster Recovery Plus, Windows 365 Reserve, maintenance windows, Move Cloud PC, share restore points to a storage account
- **Security** — Customer Lockbox, Purview Customer Key, Cloud Keyboard Input Protection, Restrict Office 365 access to Cloud PCs, Place Cloud PC under review (forensic hold), Screen Capture Protection, idle session time limits, ANC domain credential lifecycle
- **Identity** — SSO, Entra ID Join vs Hybrid Join, Conditional Access scoping (including Windows Cloud Login coverage), MFA, token protection & sign-in frequency, Windows Hello for Business, RBAC, external identities (B2B), passwordless credential strategy
- **Network** — ANC health, FQDN allowlist, RDP Shortpath, RDP Multipath, Azure Firewall / NVA egress posture, hub-spoke, NSG, SSL inspection exclusions
- **Image & app delivery** — custom vs gallery, freshness, OS lifecycle (incl. ESU), Teams VDI, Webex / Zoom optimization, Convert to Gen 2, App Assure, nested virtualization, **App delivery via Intune** (no MSIX app attach — that is AVD-only)
- **Operations & governance** — ConfigMgr co-management, Cloud PC Recommendations engine review, Resource Performance report, Copilot in Intune for Windows 365, Cloud PC Configurations, Power Platform connector, Enrollment Status Page, Autopilot device preparation for Flex
- **End-user experience** — Windows App (incl. Connection Center posture), Windows 365 Switch, Windows 365 Boot, Windows 365 Link device posture, multi-monitor, GPU SKUs, AI-enabled Cloud PC

It also pairs with **[BaselinePilot](../BaselineAssessor/)** via dedicated check `W365-SEC-018` to validate that the security baseline declared in Intune is actually applied at the Cloud PC OS layer.

> **Windows 365 ≠ AVD.** This assessor deliberately does not include MSIX app attach, host pools, session hosts, FSLogix, or Azure compute scaling — those concepts do not apply to the Microsoft-managed Cloud PC service. For those, use [AVD Assessor](../AvdAssessor/).

---

## Features

### Assessment Framework
- **132 checks** across 12 categories: Inventory & Topology, Provisioning Policies, User Settings & Resilience, Identity & Access, Security & Compliance, Network (ANC), Images & App Delivery, Monitoring & Diagnostics, Governance & Operations, Cost & Optimization, End-User Experience, Landing Zone
- Every check is fully enriched: `id`, `category`, `name`, `description`, `severity`, `weight`, `type`, `effort`, `origin`, `bestPractice`, `impact`, `remediation`, `appliesTo`, `reference`, `additionalReferences`
- Catalog defined in `checks.json` — extensible without code changes
- **41 automated** checks emitted by the discovery script + **91 manual** checks for the workshop conversation

### Automated Discovery
- Standalone `Invoke-W365Discovery.ps1` script runs against the signed-in tenant via Microsoft Graph (v1.0, plus beta where required)
- Consents to read-only scopes (`CloudPC.Read.All`, `DeviceManagementConfiguration.Read.All`, `DeviceManagementManagedDevices.Read.All`, `Directory.Read.All`; optionally `Policy.Read.All` for the identity checks)
- Discovers Cloud PCs, provisioning policies, user settings, Azure Network Connections, custom + gallery images, service plans, Intune security/compliance posture, Cloud PC reports, Conditional Access targeting, and the last 30 days of audit events
- Evaluates 41 checks automatically and emits a structured JSON that can be imported into the GUI for hybrid assessment
- Supports `‑SkipLogin` for an existing Graph context, plus tunable `-InactiveDays` and `-ImageAgeWarnDays` thresholds

### Scoring & Maturity Model
- **Status scoring** — Pass = 100%, Warning = 50%, Fail = 0%; N/A is excluded entirely, while Not Assessed counts as 0 in the denominator (an unassessed check drags the score down until it is evaluated)
- **Category score** — weighted average of assessed checks within each category
- **Overall score** — weighted average across all categories
- **Six W365 maturity dimensions** — Identity & Security, Provisioning & Lifecycle, Networking (ANC), Resilience & User Settings, Operations & Governance, Experience & Delivery
- **Five maturity levels** — Initial (0–34), Developing (35–54), Defined (55–74), Managed (75–89), Optimized (90–100)
- **Composite maturity** — weighted score across all dimensions with a single maturity label

### Dashboard
- Real-time overall score and maturity level with color-coded indicator
- Per-category score bars with progress fill and percentage labels
- Score cards for each category showing pass/fail/warning count breakdown
- Six-dimension W365 maturity radar with individual dimension levels
- Inventory tiles: Cloud PCs, Provisioning Policies, User Settings, ANCs, Custom Images
- Assessment progress indicator (assessed vs. total checks)

### Report Generation
- **HTML export** — single self-contained file with score hero, category progress, six-dimension maturity radar, effort breakdown, per-check `BestPractice` callouts (always visible), `Impact` and `Remediation` exposed for Fail / Warning, `AppliesTo` SKU badges per check (Enterprise / Flex / Business), per-category narrative summaries driven by discovered inventory, twelve W365 Field Notes blocks with inline Microsoft Learn citations, and a deduplicated Sources Appendix
- **CSV export** — flat data dump of all checks with status, notes, severity, category, and evidence
- **JSON export** — complete assessment state including scores, maturity, timestamps, and discovery context
- **Report preview** — live preview tab before exporting
- Dark and light themes via CSS variables, print-friendly

### Assessment Management
- **Auto-save** — rolling backups every 60 seconds (configurable) when the assessment is dirty
- **Manual save** — named saves using customer name, stored in `assessments/`
- **Rolling backups** — up to 10 versioned backups in `_backups/` with automatic cleanup
- **Import** — load saved assessments or discovery JSONs via Import Discovery / Assessment
- **Dirty tracking** — visual indicator when unsaved changes exist, with confirmation on close/overwrite
- **Crash recovery** — resume from latest auto-save backup

### Findings View
- Filtered view showing only checks with Fail, Warning, or Error status
- Grouped by category with severity badges
- Inline notes and recommendation text
- Quick navigation back to the Assessment tab for status changes

### Theme System
- Dark and light themes with 40+ color keys
- Theme toggle button in the title bar (sun/moon icon)
- Animation toggle for reduced-motion preference
- Theme persisted in user preferences

### Activity Log
- Collapsible bottom panel with timestamped log entries
- Color-coded by level: INFO, DEBUG, WARN, ERROR, SUCCESS
- Disk log file at `$env:TEMP\W365Assessor_debug.log`
- DEBUG-level messages visible when debug overlay is enabled

---

## Quick Start

### Option 1: GUI Tool (Workshop Mode)

Double-click **`Launch_W365Assessor.bat`** or run:

```powershell
.\Launch_W365Assessor.bat
```

The launcher auto-detects PowerShell 7 (`pwsh`) and falls back to Windows PowerShell 5.1 (`powershell.exe`), runs in STA mode (required for WPF), and `Unblock-File`s the tool's own `.ps1`/`.xaml`/`.bat` files so MOTW-marked downloads load cleanly.

### Option 2: Discovery Script (Automated Scan)

```powershell
# Interactive sign-in to Microsoft Graph — current tenant
.\Invoke-W365Discovery.ps1

# Reuse an existing Graph context (no sign-in prompt)
.\Invoke-W365Discovery.ps1 -SkipLogin

# Custom output location
.\Invoke-W365Discovery.ps1 -OutputPath "C:\temp\w365_discovery.json"

# Tune evaluation thresholds:
#   -InactiveDays      — days without sign-in before a Cloud PC is flagged inactive (COST checks)
#   -ImageAgeWarnDays  — custom image age (days) that triggers a freshness warning (IMG checks)
.\Invoke-W365Discovery.ps1 -InactiveDays 60 -ImageAgeWarnDays 120
```

Then in the GUI: **File → Load discovery → select the JSON**. Auto-checks populate; the rest is a workshop-style manual review with the customer.

---

## Prerequisites

### GUI Tool

| Requirement | Details |
|---|---|
| OS | Windows 10/11 or Windows Server 2016+ |
| PowerShell | 5.1 (built-in) or 7+ |
| .NET | WPF support (included with Windows) |
| Graph modules | Not required for manual-only assessments |

### Discovery Script

| Requirement | Details |
|---|---|
| PowerShell | 5.1 or 7+ |
| `Microsoft.Graph.Authentication` | Sign-in and Graph request handling |
| Graph permissions | `CloudPC.Read.All`, `DeviceManagementConfiguration.Read.All`, `DeviceManagementManagedDevices.Read.All`, `Directory.Read.All` |
| Optional Graph permission | `Policy.Read.All` — required for the Conditional Access / MFA identity checks (`W365-IAM-003`, `W365-IAM-004`, `W365-IAM-010`, `W365-IAM-011`); without it these checks degrade to **Error** rather than blocking the rest of discovery |
| Tenant role | Cloud PC Reader, Intune Reader, or Global Reader |

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

> Cloud PC endpoints are called on the **Graph v1.0** profile (`/v1.0/deviceManagement/virtualEndpoint/*`) now that cloudPCs, provisioningPolicies, and userSettings are GA; **beta** is still used where required (e.g. `auditEvents`).

---

## Assessment Categories

| Category | Checks | ID Prefix | Sources | Key Areas |
|---|---|---|---|---|
| Security & Compliance | 18 | SEC | SEC, W365 | Local admin, Defender, baseline, redirections, watermark, Customer Lockbox, Customer Key, Cloud Keyboard Input Protection, restrict O365, forensic hold, idle session, screen capture, Windows App / Connection Center posture, BaselinePilot validation |
| Provisioning Policies | 15 | PROV | W365 | Assignment, naming, Flex, Autopatch, grace period, ESP, Autopilot Flex, Cloud Apps, User Experience Sync, concurrency buffer, Flex updates |
| Governance & Operations | 13 | GOV | CAF, W365 | Change control, OneDrive KFM, break-glass, runbook, tagging, BCP strategy, support tier, DR testing, audit, Copilot in Intune, Cloud PC Configurations, Power Platform connector, ConfigMgr co-management |
| Images & App Delivery | 13 | IMG | W365 | Custom vs gallery, freshness, app delivery via Intune, Teams VDI, OS lifecycle, language packs, EOS, failed builds, Convert to Gen 2, Webex/Zoom, nested virtualization, App Assure, ESU |
| Network (ANC) | 12 | NET | W365, WAF | ANC health, FQDN allowlist, DNS, hub-spoke, NSG, RDP hardening, Shortpath, Multipath, SSL inspection exclusions, unused ANCs, Azure Firewall / NVA, ANC credential lifecycle |
| User Settings & Resilience | 10 | USER | W365 | Assignment, Cross-Region DR, restore points, self-service reset, bulk reprovision Flex, Reserve, DR Plus, Move, maintenance windows, share restore points |
| Monitoring & Diagnostics | 10 | MON | W365, CAF | Endpoint Analytics, connection quality, audit monitoring, proactive remediations, update compliance, alerts, audit activity, Recommendations engine, Cloud PC Monitoring (preview), resource performance |
| Identity & Access | 11 | IAM | W365, SEC | SSO, Entra Join vs Hybrid, CA (incl. Windows Cloud Login), MFA, token protection & sign-in frequency, WHfB, RBAC, session lock, external identities, passwordless |
| End-User Experience | 9 | UX | W365 | Windows App, sign-in performance, W365 Switch, multi-monitor, GPU, AI-enabled (preview), Windows App Settings (preview), Remote Connection Experience (preview), W365 Boot |
| Landing Zone | 8 | LZ | LZA, CAF | App LZ, mgmt group, Azure Policy, ExpressRoute/VPN, Microsoft-hosted network, partner connectors, Government tenant, migration plan |
| Inventory & Topology | 8 | INV / CPC | W365 | Inventory snapshot, sizing, edition, service plan, region, operational state, orphaned Cloud PCs, Windows 365 Link posture |
| Cost & Optimization | 5 | COST | W365, CAF | Inactive Cloud PCs, right-sizing, Flex math, license offboarding, egress |

### Check Metadata

Each check definition in `checks.json` includes:

| Field | Description |
|---|---|
| `id` | Unique identifier (e.g. `W365-NET-007`) |
| `category` | Assessment category |
| `name` | Short human-readable label |
| `description` | What the check evaluates |
| `severity` | Critical, High, Medium, or Low |
| `weight` | Scoring weight (1–5) |
| `type` | `Auto` (discovery-backed) or `Manual` (workshop review) |
| `origin` | `W365`, `WAF`, `CAF`, `LZA`, or `SEC` |
| `effort` | Quick Win, Some Effort, or Major Effort |
| `bestPractice` | Evidence-based rationale (always shown in the report) |
| `impact` | What goes wrong if the check fails |
| `remediation` | Concrete steps to fix (shown for Fail / Warning) |
| `appliesTo` | Applicable W365 editions: Enterprise, Flex, Business (dedicated-only caveats are noted in check descriptions) |
| `reference` | Primary Microsoft Learn URL |
| `additionalReferences` | Additional supporting URLs |

### Severity Distribution

| Severity | Count |
|---|---|
| Critical | 3 |
| High | 39 |
| Medium | 59 |
| Low | 31 |

### Origin Distribution

| Origin | Count |
|---|---|
| W365 (Microsoft Learn — Windows 365) | 71 |
| SEC (security guidance) | 31 |
| WAF (Well-Architected Framework) | 18 |
| CAF (Cloud Adoption Framework) | 7 |
| LZA (Landing Zone Accelerator) | 5 |

> **Retired IDs:** `W365-SEC-005` and `W365-USER-005` are intentional gaps in the ID sequence — those checks were retired from the catalog and their IDs are not reused.

---

## Automated Checks

The discovery script emits **41 automated checks**. Per-resource checks generate one result per object (e.g. one row per provisioning policy) and roll up to the catalog ID shown.

| Discovery emits | Catalog ID | Signal |
|---|---|---|
| `W365-INV-001` | `W365-INV-001` | Tenant inventory snapshot with per-state Cloud PC breakdown |
| `W365-PROV-001-<id>` | `W365-PROV-001` | Each provisioning policy has at least one assignment |
| `W365-PROV-002-<id>` | `W365-IAM-001` | Single sign-on enabled per provisioning policy |
| `W365-PROV-003-<id>` | `W365-SEC-001` | Local admin enabled (warned for review) |
| `W365-PROV-004-<id>` | `W365-PROV-004` | Windows Autopatch integration |
| `W365-PROV-005-<id>` | `W365-PROV-005` | Grace period configuration (1–168h = Pass) |
| `W365-PROV-008-<id>` | `W365-PROV-008` | Provisioning policy missing a Cloud PC naming template |
| `W365-PROV-009-<id>` | `W365-PROV-009` | Provisioning policy Windows language/region setting present |
| `W365-USER-001-<id>` | `W365-USER-001` | User settings policy assigned |
| `W365-USER-002-<id>` | `W365-USER-002` | Cross-region DR configured |
| `W365-USER-003-<id>` | `W365-USER-003` | Restore-point frequency (≤12h Pass, ≤24h Warn, >24h Fail) |
| `W365-USER-004-<id>` | `W365-USER-004` | User self-service reset enabled |
| `W365-NET-001-<id>` | `W365-NET-001` | ANC health check status with failed sub-test names |
| `W365-NET-009-<id>` | `W365-NET-009` | ANC defined but not referenced by any provisioning policy |
| `W365-IMG-001-<id>` | `W365-IMG-001` | Custom image age (Pass / Warn / Fail thresholds) |
| `W365-IMG-007-<id>` | `W365-IMG-007` | Gallery image at or near end-of-support |
| `W365-IMG-008-<id>` | `W365-IMG-008` | Custom device image build status failed |
| `W365-CPC-001-<id>` / `-002-<id>` | `W365-CPC-001` | Cloud PCs in failed state / grace period |
| `W365-CPC-003-<id>` | `W365-COST-001` | Inactive Cloud PC license-reclaim candidate |
| `W365-CPC-004-<id>` | `W365-CPC-004` | Provisioned Cloud PC with missing UPN (orphaned) |
| `W365-SEC-010-<id>` | `W365-SEC-010` | Cloud PC disk encryption state |
| `W365-MON-007` | `W365-MON-007` | Audit log activity (proxy for forwarding / retention) |
| `W365-GOV-009` | `W365-GOV-009` | High-impact admin actions in audit events (last 30d) |

Discovery **0.2.0** adds the following automated evaluations, emitted directly under their catalog IDs:

| Catalog ID | Signal |
|---|---|
| `W365-INV-003` | Edition mix — dedicated vs. shared (Flex) `provisioningType` across policies |
| `W365-INV-004` | Service plan (SKU) coverage from `servicePlans` |
| `W365-PROV-002` | Cloud PC naming template / convention per provisioning policy |
| `W365-PROV-006` | Domain join configuration type consistency across policies |
| `W365-PROV-007` | Windows settings (language/region) present in provisioning policy |
| `W365-PROV-010` | Flex Cloud Apps (RemoteApp mode) configuration on shared policies |
| `W365-IAM-003` | Conditional Access policies target the Windows 365 apps (`Policy.Read.All`) |
| `W365-IAM-004` | MFA enforced via Conditional Access for Cloud PC sign-in (`Policy.Read.All`) |
| `W365-IAM-010` | CA coverage includes the Windows Cloud Login app (`Policy.Read.All`) |
| `W365-IAM-011` | Token protection and sign-in frequency session controls (`Policy.Read.All`) |
| `W365-COST-002` | Right-sizing candidates from Cloud PC recommendation reports |
| `W365-MON-001` | Endpoint Analytics enrollment state |
| `W365-MON-002` | Connection quality report signals |
| `W365-MON-005` | Windows update compliance posture |
| `W365-MON-008` | Cloud PC recommendations engine output |
| `W365-MON-010` | Resource performance report (CPU / RAM utilization) |
| `W365-UX-002` | Sign-in / startup performance signals |
| `W365-SEC-002` | Defender for Endpoint / managed device health |
| `W365-SEC-003` | Security baseline profile deployed to Cloud PC groups |
| `W365-SEC-004` | Device compliance policy targeting Cloud PCs |

> The four `Policy.Read.All`-backed identity checks report **Error** (not Fail) when the optional scope is not granted.

The remaining **89 checks are Manual** — designed to drive a workshop conversation with the customer and capture decisions in the assessment record.

### Discovery Inventory Buckets

| Inventory bucket | Graph endpoint (v1.0 unless noted) |
|---|---|
| Cloud PCs | `cloudPCs` |
| Provisioning policies (with assignments) | `provisioningPolicies?$expand=assignments` |
| User settings policies (with assignments) | `userSettings?$expand=assignments` |
| Azure Network Connections | `onPremisesConnections` |
| Custom device images | `deviceImages` |
| Gallery images | `galleryImages` |
| Service plans (SKUs) | `servicePlans` |
| Cloud PC reports (inactive, recommendations, connection quality, resource performance) | `reports/*` |
| Audit events (last 30 days) | `auditEvents?$filter=activityDateTime ge ...` (beta) |

---

## Scoring Methodology

### Status Scoring

| Status | Score | Description |
|---|---|---|
| Pass | 100% | Check fully satisfied |
| Warning | 50% | Partially met or acceptable risk documented |
| Fail | 0% | Not met — remediation recommended |
| N/A | Excluded | Not applicable to this environment — removed from the denominator |
| Not Assessed | 0% (in denominator) | Not yet evaluated — counts as 0 until assessed, so the score rises as the assessment progresses |

### Calculation

**Category Score** = Σ (check_score × check_weight) / Σ (check_weight) for all non-N/A checks in that category (Not Assessed contributes 0 to the numerator but its weight stays in the denominator).

**Overall Score** = Σ (category_score × category_weight) / Σ (category_weight) across all categories with assessed checks.

### W365 Maturity Dimensions

Checks are mapped to six W365-specific maturity dimensions by ID prefix. All 132 catalog checks map to exactly one dimension.

| Dimension | Prefixes | Focus |
|---|---|---|
| Identity & Security | `W365-IAM-`, `W365-SEC-` | Authentication, authorization, endpoint protection, data protection |
| Provisioning & Lifecycle | `W365-PROV-`, `W365-INV-`, `W365-CPC-` | Inventory, provisioning policies, Cloud PC operational state |
| Networking (ANC) | `W365-NET-` | ANC health, FQDN, DNS, hub-spoke, RDP optimization, egress |
| Resilience & User Settings | `W365-USER-` | DR, restore points, Reserve, maintenance windows, Flex reprovision |
| Operations & Governance | `W365-GOV-`, `W365-MON-`, `W365-COST-` | Change control, audit, monitoring, recommendations, cost |
| Experience & Delivery | `W365-IMG-`, `W365-UX-`, `W365-LZ-` | Images, app delivery, end-user experience, landing zone, partner connectors |

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
│ Customer or │     │ Launch GUI       │    │ Walk through     │  │ Export HTML  │
│ assessor    │────▶│ Import discovery │───▶│ manual checks    │─▶│ report and   │
│ runs        │     │ Review automated │    │ with customer    │  │ share        │
│ discovery   │     │ findings         │    │ Add notes per    │  │              │
│ script      │     │                  │    │ check            │  │              │
└─────────────┘     └──────────────────┘    └──────────────────┘  └──────────────┘
```

1. **Pre-workshop** — run `Invoke-W365Discovery.ps1` against the customer's tenant. Can be executed by the customer with `Cloud PC Reader` / `Intune Reader` rights.
2. **Workshop** — launch the GUI, import the discovery JSON, review auto-evaluated findings on the Dashboard.
3. **Assessment** — navigate to the Assessment tab, walk through manual checks with the customer, set status (Pass / Fail / Warning / N/A) and add notes.
4. **OS-layer validation (recommended)** — sign in to one Cloud PC per persona, run **[BaselinePilot](../BaselineAssessor/)** to validate the security baseline at the OS layer (catalog check `W365-SEC-018`), and attach the BaselinePilot HTML report to the evidence pack.
5. **Deliverable** — export an HTML report, share with the customer, and archive the JSON snapshot for the next review.

### Hybrid (recommended)
- Use the discovery script for the 41 auto-checks and to populate inventory tiles.
- Use the GUI for the remaining 89 manual checks, captured live in the workshop.
- Pair with BaselinePilot for OS-layer baseline conformance.

### Manual-only
- Skip the discovery script. Open the GUI, walk through every check with the customer.
- Useful when Graph access is not yet available or you only want the framework.

### Audit-ready
- Save the assessment as JSON, export HTML for the customer, archive both alongside the BaselinePilot report.
- Re-import the JSON months later to track progress against the same baseline.

---

## File Layout

```
W365Assessor/
├── W365Assessor.ps1            # WPF GUI app (entry point)
├── W365Assessor_UI.xaml        # WPF layout (dashboard, tiles, tabs)
├── Invoke-W365Discovery.ps1    # Microsoft Graph data collector
├── checks.json                 # Canonical 132-check catalog
├── Launch_W365Assessor.bat     # Convenience launcher (PS7 / 5.1 fallback + Unblock-File)
├── README.md                   # This file
├── assessments/                # (created on first run) discovery + saved assessments
├── reports/                    # (created on first run) HTML / CSV / JSON exports
└── _backups/                   # (created on first run) rolling autosave snapshots
```

---

## Versioning

| Component | Version |
|---|---|
| `Invoke-W365Discovery.ps1` | 0.2.0 |
| `checks.json` | 1.1 (schema), 132 checks |
| `W365Assessor.ps1` | 0.2.0 |

### Changelog

- **0.2.0 (2026-07-18)** — full audit fix pass; see [AUDIT.md](AUDIT.md) for the complete defect list and rationale. Highlights: Frontline → Flex product rename across the catalog; 18 checks converted from Manual to Auto (and 2 superseded Auto checks retired to Manual) plus 4 new checks (`W365-IAM-010`, `W365-IAM-011`, `W365-SEC-019`, `W365-INV-006`) for a total of 132 checks / 41 Auto; 60 dead Microsoft Learn reference URLs replaced with verified current slugs; guidance drift fixes (RD client retirement, GPU GA, DR Plus / CRDR / Boot / Switch GA, 24H1 baseline); discovery migrated to Graph v1.0 with optional `Policy.Read.All`; scoring semantics corrected (N/A excluded, Not Assessed 0-in-denominator).
- **0.1.0** — initial release (128 checks / 23 Auto).

---

## Roadmap / Out of Scope

These were considered and deferred:

- **License vs assignment gap report** — needs `Directory.Read.All` user enumeration cross-referenced against W365 service plan IDs.
- **ANC `runHealthChecks` action** — discovery reads the cached health status; we do not trigger a fresh health check (write scope, longer runtime).
- **Cost data** — Windows 365 license cost lives in the M365 admin / billing centre, not Graph. Utilization comes from Endpoint Analytics. Two data sources to stitch — left for manual inspection for now.
- **Hybrid Entra Join AD-side enumeration** — surfaced as manual checks; we do not enumerate AD DS objects.

---

## Related Tools in the Endpoint Toolkit

- **[AvdAssessor](../AvdAssessor/)** — sister tool for Azure Virtual Desktop. Shares the framework and report engine.
- **[BaselinePilot](../BaselineAssessor/)** — Windows 11 client security baseline assessment. Use inside a Cloud PC to validate `W365-SEC-018` (Cloud PC OS baseline conformance).
- **[PolicyPilot](../PolicyPilot/)** — CSP/ADMX policy authoring with full enrichment.
- **[ADMXPolicyComparer](../ADMXPolicyComparer/)** — diff GPOs and policies across baselines.

---

## License

See repository [LICENSE](../../LICENSE).
