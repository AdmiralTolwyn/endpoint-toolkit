# BaselinePilot — Windows Client Security Baseline Assessment Tool

BaselinePilot is a two-component security baseline assessment tool for Windows 11 clients. It combines Microsoft Security Baselines (Intune MDM + GPO) with ODA-style operational health checks in a rich WPF dashboard.

> The project folder is `tools/BaselineAssessor/` (matching this repo's tool-directory convention); the product itself is named **BaselinePilot** — the two names are not a typo.

**Versions**: App `0.2.0` · Collector `1.3.0` · Catalog (`checks.json`) `1.1`. See [`AUDIT.md`](AUDIT.md) for the July 2026 audit and fix-pass history behind the current versions.

### Assay Cloud PC Applicability (1.3.0)

```powershell
.\Invoke-BaselineCollection.ps1 -AssessmentProfile Windows365CloudPc -OutputPath .\cloud-pc-baseline.json
```

Use only for independently identified Windows 365 Cloud PCs. This adds an explicit
operator declaration (`assessmentContext`: schema 1.0, windows365-cloud-pc,
profile version 2026-09-18, selectionSource Operator); it does not detect devices,
skip collection or change configuration. Default `Generic` omits the declaration.

The updated **Assay** importer excludes eight guest BitLocker controls and UAC-011
Administrator Protection, based on Microsoft's [security overview](https://learn.microsoft.com/en-us/windows-365/enterprise/security-guidelines).
Removable-drive encryption DATA-013 stays unassessed for scope review. This is a
bounded exclusion overlay, not a complete Cloud PC security baseline. Other targets
are unchanged; service-side encryption and recovery protection are not evaluated.
Malformed/unsupported declarations or conflicting/missing client metadata stay
unassessed. Declaration provenance is not authenticated device identity.

The legacy WPF evaluator does **not** implement the profile. Existing exports and
saved assessments are not automatically migrated. Assay retains profile evidence
in results/reports, and mixed-machine Pass plus unknown no longer yields Pass.
`Test-BaselineApplicability.ps1` exercises actual production parameter/export blocks
without running the administrator collector. Set `ASSAY_BASELINE_APPLICABILITY_FIXTURE`
to a local output path to write synthetic production JSON for native tests.

Collector 1.2.1 adds a read-only [Get-TlsCipherSuite](https://learn.microsoft.com/en-us/powershell/module/tls/get-tlsciphersuite?view=windowsserver2025-ps)
inventory under `tlsConfig.cipherSuites`, with provider, names and explicit
Complete/Partial/Error state. Assay uses it for NET-028 instead of incorrectly
testing legacy protocol registry switches. NULL suites require application review,
not an inference of unencrypted traffic; see the [Microsoft cipher-suite reference](https://learn.microsoft.com/en-us/windows/win32/secauthn/tls-cipher-suites-in-windows-11-v22h2).
No TLS settings are changed. The legacy GUI evaluator is unchanged. Offline tests
mock the production provider block and do not query this machine's TLS settings.

### Collector Evidence Review (18 September 2026)

- Firewall profiles now use [ActiveStore](https://learn.microsoft.com/en-us/powershell/module/netsecurity/get-netfirewallprofile?view=windowsserver2025-ps), not the default PersistentStore. Unknown/NotConfigured values remain null, not false. Provider failures have explicit collection-failure markers.
- The fields `RealTimeProtectionEnabled`, `BehaviorMonitoringEnabled` and `IoavProtectionEnabled` now represent [Get-MpComputerStatus](https://learn.microsoft.com/en-us/powershell/module/defender/get-mpcomputerstatus?view=windowsserver2025-ps) observations. Preference-derived intent is retained separately in `*Configured` fields. Missing status is not inferred from policy intent.
- `TamperProtectionSource` retains the source property rather than duplicating `IsTamperProtected`. The [controlled-configuration reference](https://learn.microsoft.com/en-us/defender-endpoint/secure-controlled-configuration) documents their different meanings and preview limits.
- Unavailable [Win32_DeviceGuard](https://learn.microsoft.com/en-us/windows/security/hardware-security/enable-virtualization-based-protection-of-code-integrity#use-win32_deviceguard-wmi-class) runtime evidence remains null, even if registry configuration exists. Known absent services remain false; configured and running states are distinct.

Run `Test-BaselineCollector.ps1` for offline tests of the actual collector area
bodies with mocked providers. It does not run the administrator-only collector
or query the development device. Recollect with 1.1.2 for these fixes: older
exports cannot recover provider provenance that was not recorded.

The collector does not load `csp_metadata.json` or `admx_metadata.json`.
BaselinePilot loads CSP metadata for UI/remediation enrichment, not as baseline
targets; no ADMX loader is used in that app. Assay independently embeds its
reviewed source/unified catalog and evaluates the exported JSON. The metadata
files do not automatically change Assay checks, defaults or recommendations.

Collector 1.1.3 additionally records `powershellConfig.legacyEngine` using the
[Microsoft PowerShell team's documented feature queries](https://devblogs.microsoft.com/powershell/windows-powershell-2-0-deprecation/):
`Get-WindowsOptionalFeature -Online -FeatureName MicrosoftWindowsPowerShellV2`
on clients, or `Get-WindowsFeature -Name PowerShell-V2` on servers. No features
are changed. Unknown platform, unavailable provider, absent result and query
failure remain unsupported/error evidence, never inferred removal. DISM can
write its own diagnostic log. Assay's updated SEC-065 consumes this evidence;
the legacy BaselinePilot check catalog/evaluator is not updated by this change.

`Test-BaselineCollector.ps1 -AssayCatalogPath <assay>/rust/catalogs/source/baseline.json`
also checks every automatic registry binding against the actual collector read
list with mocked named/bulk reads. This proves declared path coverage, not real
provider success, safe policy precedence or complete coverage of nonregistry checks.

### Embedded Speculation Control (1.2.0)

The collector contains Microsoft's `Get-SpeculationControlSettings` function
from SpeculationControl 1.0.19, unchanged apart from line-ending normalization,
with the upstream MIT license and attribution. Only the collector script needs
to be deployed; no module installation or companion file is required.

```powershell
.\Invoke-BaselineCollection.ps1 -IncludeSpeculationControl -SkipEventCollection -OutputPath .\baseline.json
```

This opt-in calls the embedded function with `-Quiet`. It reads Windows native
mitigation information and CIM processor/OS information, using `Add-Type` for
the native query. It does not download, install a module, change execution policy,
write mitigation registry values, update firmware or remediate. Normal collector
administrator requirements and local diagnostic/output side effects still apply.
Restricted language mode, platform/provider incompatibility and query failures
produce explicit Error evidence, not a guessed protection status. No opt-in
means NotRequested. The optional section has its own state; the legacy 22-area
progress counters are unchanged.

The `speculationControl` section records schema/version, UTC capture time,
embedded source commit/hash, 39 allowlisted Boolean/null fields and five status
strings. Unsupported types make the section Partial; absent/conditional fields
remain absent. Arbitrary module properties and error text are not exported.
Source hashes identify the reviewed code, not an authenticated device attestation.

Assay replaces SEC-057's old registry-only comparison with 13 family results:
BTI, KVA shadow, SSBD, L1TF OS, MDS, three MMIO families, branch confusion, GDS,
SRSO, divide-by-zero and RFDS. Reported applicable enablement passes; reported
unaffected/immune hardware is not a failure; observed disabled mitigations warn
for review. Unknown reporting/applicability never becomes a clean result. BHB
flags are observations only until their applicability/reporting contract is
reviewed. PCID/retpoline optimizations are not independent security requirements.
Pass is scoped to these evaluated families, not all silicon vulnerabilities or
firmware freshness. A guest does not certify the host or L1TF VMM protection.
The legacy BaselinePilot catalog/evaluator is not changed by this addition.

Sources: [client guidance](https://support.microsoft.com/help/4073119),
[server guidance](https://support.microsoft.com/help/4072698),
[output interpretation](https://support.microsoft.com/help/4074629),
[pinned source](https://github.com/microsoft/SpeculationControl/blob/f4d2a2d4f32e93279703d50283b80672e3d3a2c3/SpeculationControl.psm1).
The implementation source is authoritative for exact field names and polarity;
some older KB prose/examples use inconsistent names or inverse wording. Do not
copy a single registry override value across clients, servers and CPU families.

`Test-SpeculationEvidence.ps1` parses the collector without executing its main
body, checks the embedded function SHA256 and license, and tests the wrapper
with synthetic results. Optional `-UpstreamPath` compares a separately obtained
pinned source file; no source download or detector execution occurs in tests.
The normalized function SHA256 is
`6ACA20A3EAD9E45CC9E6043223502B09DBFFB915A8D0EBF44700107985C87E21`.
The original Authenticode block was not copied: it would not sign the combined
collector. Organizations may sign the complete collector through their normal
deployment process; do not bypass execution policy for this feature.

## Architecture

```
Customer Machine                        Assessor Workstation
┌─────────────────────────┐             ┌──────────────────────────────┐
│ Invoke-BaselineCollection│  JSON file  │ BaselinePilot.ps1 (WPF GUI) │
│ .ps1                     │ ──────────► │ + BaselinePilot_UI.xaml      │
│ (headless, admin, no     │  transfer   │ + checks.json (312 checks)  │
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
| 3 | Applied Policies | `gpresult /scope computer` (opt-in via `-IncludeGpoData`) |
| 4 | MDM Enrollment | Registry (Enrollments + PolicyManager) |
| 5 | Security Policy Export | `secedit /export` |
| 6 | Audit Policy | `auditpol /get /category:*` |
| 7 | Registry Baselines | ~300 registry keys (Intune + GPO paths) |
| 8 | Defender Configuration | `Get-MpPreference` + `Get-MpComputerStatus` |
| 9 | Firewall Profiles | `Get-NetFirewallProfile` |
| 10 | Services | `Get-Service` (37 baseline-relevant services) |
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
| `-MaxEventsPerQuery` | 2000 | Cap per event query group (raise for deeper forensic pulls; large values can produce multi-MB JSON on busy hosts) |
| `-SkipEventCollection` | `$false` | Skip Area 22 entirely (~30s total) |
| `-EventSummaryOnly` | `$false` | Counts + top-N stats only |
| `-IncludeGpoData` | `$false` | Opt-in to Area 3 (`gpresult /scope computer`) — the most expensive/fragile collection step; skipped by default |
| `-IncludeSpeculationControl` | `$false` | Query the embedded Microsoft 1.0.19 detector; export explicit mitigation state and provenance without external modules or remediation |
| `-AssessmentProfile` | `Generic` | `Windows365CloudPc` records an operator-selected versioned applicability declaration for the updated Assay importer; no automatic detection |
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

## Check Categories (312 checks)

| Category | Prefix | Count | Scope |
|----------|--------|-------|-------|
| Security Configuration | SEC | 87 | Security options, services, user rights, SmartScreen, RDP, WinRM, PowerShell |
| Monitoring & Audit | MON | 50 | Audit policy, event log sizing, audit gap detection |
| Defender & Endpoint Security | DEF | 38 | Defender settings, ASR rules, VBS, Credential Guard |
| Network Security | NET | 35 | Firewall, SMB, TLS/SCHANNEL, LLMNR/NetBIOS |
| Authentication & Credentials | AUTH | 33 | Password policy, lockout, Kerberos, LAPS, NTLM |
| Operations & Health | OPS | 31 | Updates, drivers, services, tasks |
| Data Protection | DATA | 22 | BitLocker, encryption, privacy, removable media |
| User Account Control | UAC | 11 | UAC settings, elevation prompts, admin approval |
| Performance & Stability | PERF | 5 | Boot performance, reliability events |

Some numeric IDs within a category are non-contiguous — 10 IDs from earlier catalog revisions were retired rather than reused, to avoid breaking saved-assessment compatibility; the July 2026 fix pass documented these gaps explicitly in `_metadata.retiredIds` (`checks.json`) instead of leaving them unexplained.

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
| Not Assessed | 0 | — | — |
| Accepted Risk | — | — | Excluded from both |
| N/A | — | — | Excluded from both |

- **Not Assessed** (e.g. a collection section failed or a value could not be resolved) scores 0 points but **stays in the Risk Score denominator** — it is not the same as a Fail, but it is not silently dropped either. It **is excluded** from the Baseline Compliance % denominator.
- **Accepted Risk** and **N/A** are excluded from *both* the weighted Risk Score and Baseline Compliance % — this matches the documented governance intent (a risk that's been formally accepted, or a check that doesn't apply, should not drag down either score).

**Baseline Compliance %**: Flat pass/total ratio across all assessed checks (Not Assessed, Accepted Risk, and N/A excluded from the denominator).

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
├── checks.json                    # 312 check definitions
├── csp_metadata.json              # CSP metadata (descriptions, allowed values)
├── admx_metadata.json             # ADMX policy metadata
├── Launch_BaselinePilot.bat       # Batch launcher — starts the GUI
├── Run_Collection.bat             # Batch launcher — elevation-gated collector run
├── Test-BaselinePilot.ps1         # Offline validation tests (collection/catalog/evaluation)
├── README.md                      # This file
├── AUDIT.md                       # July 2026 audit + fix-pass changelog
├── assessments/                   # Saved assessment JSON files (created on first save, not shipped)
├── reports/                       # Generated HTML reports (created on first export, not shipped)
└── templates/                     # Report templates
```

## Changelog

- **2026-07 fix pass** — App `0.2.0` / Collector `1.1.0` / Catalog `1.1`. Resolved the correctness, drift, and duplicate-severity findings from the July 2026 audit (collector key mismatches, stale baseline values, Not-Assessed scoring, new 2026 checks). Full findings and fix status: [`AUDIT.md`](AUDIT.md).
