# PolicyPilot — Group Policy & MDM Documentation Tool

PolicyPilot is a WPF-based PowerShell tool that scans, documents, and compares Group Policy and Intune MDM settings across your environment. It replaces slow `Get-GPO`/`Get-GPOReport` calls with pure LDAP and SYSVOL reads for fast cross-domain scanning, enriches raw registry settings with human-readable ADMX/CSP policy names, and generates professional HTML reports.

## Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│                     PolicyPilot.ps1 (WPF GUI)                       │
│                     PolicyPilot_UI.xaml (XAML layout)                │
├─────────────┬───────────────┬──────────────┬────────────────────────┤
│  AD Scan    │  Local RSoP   │  Intune Scan │  Combined (all three)  │
│  LDAP/SYSVOL│  gpresult /x  │  Graph API   │                        │
├─────────────┴───────────────┴──────────────┴────────────────────────┤
│  admx_metadata.json (2,027 policies)  │  csp_metadata.json (~3,000) │
│  Built by Build-AdmxDatabase.ps1      │  Built by Build-CspDatabase │
└───────────────────────────────────────┴─────────────────────────────┘
```

## Quick Start

### 1. Launch the GUI

```powershell
# Double-click or run:
.\Launch_PolicyPilot.bat
```

The batch launcher auto-detects PowerShell 7 (`pwsh`), falls back to Windows PowerShell 5.1, unblocks downloaded files, and starts the app maximized with the console hidden.

### 2. Run a Scan

Select a scan mode in the sidebar, configure domain/DC/OU scope if needed, and click **Scan**.

| Mode | Source | Requirements |
|------|--------|-------------|
| **Local** | `gpresult /x` on this machine | None (built-in) |
| **AD** | LDAP + SYSVOL across the domain | Network access to a DC |
| **Intune** | Microsoft Graph API | `Microsoft.Graph.DeviceManagement` module |
| **Combined** | Merges Local + Intune results | Both of the above |

### 3. Generate a Headless Report (CLI)

```powershell
# Generate an AD scan report without the GUI
.\PolicyPilot.ps1 -Headless -ReportType AD -OutputPath .\report.html

# Scan local machine policies
.\PolicyPilot.ps1 -Headless -ReportType Local
```

## Scan Modes

### AD Mode (LDAP + SYSVOL)

All scanning uses native .NET — no `Get-GPO`, `Get-GPOReport`, or `Get-GPPermission` cmdlets in the hot path:

| Operation | Method | Performance |
|-----------|--------|-------------|
| GPO enumeration | LDAP `DirectorySearcher` with paged results | ~2s for 500 GPOs |
| Settings extraction | SYSVOL `registry.pol` binary parser (PReg format) | ~100ms per GPO |
| Security filtering | LDAP ACL read (`nTSecurityDescriptor`) | ~50ms per GPO |
| Link discovery | Batch LDAP `gPLink` attribute query | ~2s for entire domain |
| DC location | `FindDomainController()` — pinned for all calls | One-time ~1s |

*Cross-domain scans over VPN that previously took hours now complete in 3–5 minutes.*

**OU-scoped scanning**: Enter an OU distinguished name (or click **Detect My OU**) to scan only GPOs linked to that OU tree instead of the entire domain.

**Caching**: Per-GPO XML is cached to `$env:TEMP\PolicyPilot_GPOCache\<domain>\` with a 4-hour TTL and `WhenChanged` invalidation. Scan snapshots are saved via `Export-Clixml` for instant restore on next launch.

### Local Mode

Runs `gpresult /scope:computer /x` to get the Resultant Set of Policy (RSoP) applied to this machine. No modules or domain connectivity required.

### Intune Mode

Authenticates via `Connect-MgGraph` and queries device configuration profiles, compliance policies, and administrative templates. Supports WAM-based auth with subscription selection.

### Combined Mode

Runs Local + Intune scans and merges results into a single view with source tagging.

## GUI Tabs

| Tab | Purpose |
|-----|---------|
| **Dashboard** | Summary cards — GPO count, setting count, conflict count, link coverage, category breakdown |
| **GPO List** | All GPOs with status, link count, settings, security filtering. Click to expand details |
| **Settings** | Full setting inventory with scope, category, GPO source, registry key, value data |
| **Conflicts** | Duplicate settings across GPOs with precedence resolution and winner highlighting |
| **Intune Apps** | Intune-discovered applications (Win32, LOB, Store) |
| **Report** | Live HTML preview with export controls |
| **IME Logs** | Intune Management Extension log viewer with live tail, search, minimap, and heatmap |
| **GPO Logs** | Group Policy event/debug log viewer with live tail and `gpupdate` trigger |
| **MDM Sync** | MDM/OMA-DM sync log viewer with live tail |
| **Tools** | Registry quick-jump, status code lookup, Base64 decoder, ETW SyncML tracer, Autopilot hash decoder, Wi-Fi/VPN profile viewer, MDM Node Cache explorer, MMP-C sync trigger |
| **Settings** | Theme (dark/light/high-contrast), preferences, scan scope, DC/OU overrides |

## Export Formats

| Format | Description |
|--------|-------------|
| **HTML** | Rich report with dark/light toggle, sortable tables, conflict highlights, executive summary |
| **CSV** (settings) | All settings with GPO source, scope, category, registry key, value |
| **CSV** (conflicts) | Conflicting settings with precedence winner and all competing GPOs |
| **REG** | Windows `.reg` file for selected registry-based settings |
| **PS1** | PowerShell remediation script with `Get-GPO` + `Set-GPRegistryValue` calls |
| **JSON** | Baseline template for snapshot comparison |

## Policy Enrichment

### ADMX Metadata (`admx_metadata.json`)

Maps raw registry keys to human-readable policy names, categories, descriptions, and allowed values. Ships with **2,027 policies** from 46 ADMX files covering Windows, Edge, Office, Defender, and security baselines.

**Rebuild from your environment:**

```powershell
# Parse security-relevant ADMX subset (~50 files)
.\Build-AdmxDatabase.ps1

# Parse ALL ADMX files in PolicyDefinitions (~241+)
.\Build-AdmxDatabase.ps1 -IncludeAll

# Custom ADMX path (e.g., central store)
.\Build-AdmxDatabase.ps1 -AdmxPath '\\domain\SYSVOL\domain\Policies\PolicyDefinitions'
```

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-AdmxPath` | `$env:SystemRoot\PolicyDefinitions` | Directory containing `.admx` files |
| `-Language` | `en-US` | ADML language subfolder |
| `-OutputPath` | `admx_metadata.json` | Output JSON file path |
| `-IncludeAll` | `$false` | Parse all ADMX files vs. security-focused subset |

The builder also picks up SecGuide and MSS-Legacy ADMX templates from a local `templates/` folder if present (from the Microsoft Security Compliance Toolkit).

### CSP Metadata (`csp_metadata.json`)

Maps MDM policy CSP paths to friendly names, descriptions, defaults, allowed values, scope, editions, and minimum Windows version. Includes GP↔CSP cross-references. Ships with **~3,000 settings** across 100+ CSP areas.

**Rebuild from Microsoft docs:**

```powershell
# Scrape all CSP areas from learn.microsoft.com (~5-10 min)
.\Build-CspDatabase.ps1

# Specific areas only
.\Build-CspDatabase.ps1 -Areas 'Update', 'Defender', 'Browser'
```

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-Areas` | All (~100+ areas) | Specific CSP policy areas to scrape |
| `-OutputPath` | `csp_metadata.json` | Output JSON file path |

Recommended refresh interval: every 3–6 months (as Microsoft updates CSP documentation).

## Advanced Features

### Conflict Detection

Automatically identifies settings configured in multiple GPOs. Shows:
- All competing GPOs and their values
- The **winner** (highest precedence based on link order, OU hierarchy, and enforcement)
- Enforced vs. inherited status

### Snapshot Comparison

Save scan results as snapshots and compare them over time to detect:
- New or removed GPOs
- Changed settings (value drift)
- Link changes

### Impact Simulation

"What-if" tool: select a GPO and simulate its removal to see which settings would change or lose enforcement.

### Log Viewers

Three built-in log viewers with:
- **Live tail** with configurable refresh
- **Regex search** with match navigation
- **Minimap** and **heatmap** overlays for error/warning density
- **Severity classification** with color-coded badges
- **Content filtering** by category

| Viewer | Log Source |
|--------|-----------|
| IME | `%ProgramData%\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log` |
| GPO | Group Policy operational event log + debug log + GPP trace logs |
| MDM | `Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider` event log |

### Tools Tab

| Tool | Description |
|------|-------------|
| **Registry Quick-Jump** | Open `regedit` at PolicyManager, Enrollments, IME, DeclaredConfig, or custom paths |
| **Status Code Lookup** | Decode SyncML, OMA-DM, HRESULT, and Win32 error codes |
| **Base64 Decoder** | Decode Base64 strings with auto JSON pretty-print |
| **ETW SyncML Tracer** | Capture live OMA-DM SyncML traffic via ETW tracing |
| **Autopilot Hash Decoder** | Decode hardware hash TLV blobs (SMBIOS, disk serial, MAC, TPM EK, UUID) |
| **Wi-Fi/VPN Profiles** | View deployed network profiles |
| **MDM Node Cache** | Browse the `Provisioning\NodeCache` registry tree |
| **MMP-C Sync** | Trigger co-management (MMP-C) sync |

## Prerequisites

| Component | Required | For |
|-----------|----------|-----|
| PowerShell 5.1+ | Yes | Core runtime |
| .NET Framework 4.7.2+ | Yes | WPF (ships with Windows 10 1803+) |
| RSAT GroupPolicy module | No* | AD scan mode (auto-prompts to install) |
| `Microsoft.Graph.DeviceManagement` | No* | Intune scan mode |
| Network access to DC / SYSVOL | No* | AD scan mode |
| Local administrator | No* | Local RSoP scan (`gpresult`) |

*Required only for the corresponding scan mode.

## File Structure

```
PolicyPilot/
├── PolicyPilot.ps1            # Main WPF application (~12,000 lines)
├── PolicyPilot_UI.xaml        # WPF XAML layout (~3,000 lines)
├── Launch_PolicyPilot.bat     # Batch launcher (auto-detects PS7/PS5.1)
├── Build-AdmxDatabase.ps1    # ADMX → JSON metadata builder
├── Build-CspDatabase.ps1     # CSP → JSON metadata scraper
├── admx_metadata.json         # Pre-built ADMX database (2,027 policies, 2.4 MB)
├── csp_metadata.json          # Pre-built CSP database (~3,000 settings, 1.5 MB)
├── README.md                  # This file
├── reports/                   # Generated HTML/CSV reports (auto-created)
└── snapshots/                 # Scan snapshots for comparison (auto-created)
```

## Themes

Three built-in themes with full WPF brush remapping:

| Theme | Description |
|-------|-------------|
| **Dark** (default) | Dark background with accent colors |
| **Light** | Light background for projectors/presentations |
| **High Contrast** | Accessibility-focused with maximum contrast |

HTML reports include a standalone dark/light toggle independent of the app theme.

## Parameters

```powershell
.\PolicyPilot.ps1 [-Headless] [-ReportType <Local|AD|Intune|Combined>] [-OutputPath <path>]
```

| Parameter | Type | Description |
|-----------|------|-------------|
| `-Headless` | Switch | Run without the GUI — scan and generate an HTML report, then exit |
| `-ReportType` | String | Scan mode: `Local`, `AD`, `Intune`, or `Combined` |
| `-OutputPath` | String | Output path for the HTML report (default: `reports/` with timestamp) |
