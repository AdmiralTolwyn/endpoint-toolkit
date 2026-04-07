# Azure Change Tracker

A modern WPF desktop application that snapshots Azure subscription resources as JSON, diffs snapshots to detect additions, removals, and modifications, classifies changes by severity, and exports polished reports. Built entirely in PowerShell with a dark/light themed UI, timeline view, resource inventory tree, and background-threaded Azure operations.

![Version](https://img.shields.io/badge/version-0.1.0--alpha-blue)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![Platform](https://img.shields.io/badge/platform-Windows%2010%2B-0078D4?logo=windows)
![License](https://img.shields.io/badge/license-MIT-green)

---

## Why This Tool?

Azure subscriptions change constantly — resources are provisioned, reconfigured, scaled, tagged, and deleted. Without a change tracking mechanism outside of Azure Activity Log (which has a 90-day window and doesn't show property-level diffs), it's difficult to answer:

- *"What changed in our subscription between Tuesday and today?"*
- *"Did anyone modify the NSG rules or resize the VMs?"*
- *"How does our current state compare to the baseline we approved?"*

**Azure Change Tracker** solves this by capturing daily resource snapshots and producing precise, property-level diffs with severity classification:

- **One-click snapshots** via Azure Resource Graph — captures every resource in seconds
- **Three-pass diff engine** — detects added, removed, and modified resources with property-level detail
- **Severity classification** — Critical (RBAC, NSGs, KeyVault policies), High (SKU, location, VNet peerings), Medium (tags, new resources), Low (everything else)
- **Volatile change filtering** — automatically identifies and optionally hides metadata-only changes (timestamps, instanceView, modification tracking) that appear as false-positive diffs
- **Canonical hashing** — sorted-key JSON with normalised timestamps ensures stable comparison regardless of API response ordering
- **Beautiful exports** — self-contained HTML reports with dark/light theme, search, and drill-down; CSV for spreadsheet analysis; raw JSON for automation
- **Named baselines** — save a "known good" snapshot to compare against at any time

---

## Features

### Snapshot Engine

- **Azure Resource Graph powered** — queries all resources in the current subscription via `Search-AzGraph` with automatic pagination (1000 resources per page)
- **SHA-256 state hashing** — each resource gets a hash computed from its stable properties (volatile metadata like `instanceView`, `lastModifiedTime`, `statusTimestamp`, `changedTime` are stripped before hashing)
- **Background execution** — snapshots run in a separate STA runspace so the UI stays responsive
- **Daily file organisation** — snapshots are saved as `{date}.json` under `snapshots/sub_{subscriptionId}/`
- **Rich metadata** — each snapshot records: snapshot ID, subscription info, capture timestamp, captured-by account, resource/RG counts, and tool version
- **Configurable exclusions** — skip resource types you don't want to track (e.g. `Microsoft.Advisor/recommendations`) via the Settings tab
- **Snapshot overwrite protection** — prompts before overwriting today's existing snapshot

### Diff Engine

- **Three-pass comparison**:
  - **Pass 1 (Added)**: Resources in the newer snapshot not present in the older one
  - **Pass 2 (Removed)**: Resources in the older snapshot not present in the newer one
  - **Pass 3 (Modified)**: Resources present in both where the canonical hash differs — runs recursive property-level diff
- **Canonical JSON serialisation** — objects are serialised with alphabetically sorted keys at every nesting level; ISO 8601 timestamps are normalised to 3-digit millisecond `Z` format to prevent false positives from Azure returning varying sub-second precision
- **Recursive property diff** — `Compare-PropertyDiff` walks nested object trees and produces dot-notation paths (e.g. `properties.storageProfile.imageReference.sku`) with old/new values
- **Volatile change detection** — changes that only affect metadata properties (instance views, timestamps, modification tracking) are flagged as `volatileOnly` and can be hidden with the "Hide volatile" toggle
- **Severity classification** — each change is classified as Critical, High, Medium, or Low based on the resource type and changed properties

### Severity Rules

| Severity | Triggers |
|----------|----------|
| **Critical** | All resource removals; changes to `Microsoft.Authorization/*`, `Microsoft.Network/networkSecurityGroups*`, `Microsoft.KeyVault/vaults/accessPolicies`, `Microsoft.ManagedIdentity/*` |
| **High** | Changes to `sku.name`, `sku.tier`, `location`, public IP addresses, VNet peerings |
| **Medium** | Tag changes; newly added resources |
| **Low** | All other modifications |

### User Interface

| Tab | Description |
|-----|-------------|
| **Dashboard** | Stat cards (Total Resources, Changes, Added, Removed), severity distribution, recent changes summary |
| **Changes** | Colour-coded change cards with left accent bars, change type icons, resource name/type/group metadata, severity indicators, staggered fade-in animations. Includes resource type checkboxes, change type filter, text search, and volatile change toggle |
| **Timeline** | Vertical connector-dot timeline of all saved diffs. Each entry shows date range, total change count, mini stat pills (added/removed/modified), and expands to show individual resource changes grouped by type |
| **Resources** | Hierarchical TreeView of all resources in the latest snapshot, grouped by resource type with icons and count badges. Filterable by resource type, resource group, and text search |
| **Settings** | Dark/light theme toggle, auto-snapshot on launch, retention cleanup (configurable days), resource type exclusion list, achievement badges |

### Additional UI Components

| Component | Description |
|-----------|-------------|
| **Auth Bar** | Connection status dot, account label, subscription dropdown, Sign In / Sign Out buttons |
| **Icon Rail** | Vertical icon navigation with active tab indicators |
| **Collapsible Sidebar** | Context-sensitive panels per tab (snapshot list, diff controls, filters, achievements) |
| **Splash Overlay** | 3-step startup sequence with animated step dots (prerequisites → Azure auth → workspace init) |
| **Toolbar** | Snapshot Now, Export Report, Save Baseline action buttons |
| **Debug Console** | Collapsible bottom panel with colour-coded RichTextBox activity log, ring-buffer cleanup |
| **Toast Notifications** | Auto-dismissing stacked notifications (Success/Error/Warning/Info) with themed colours |
| **Themed Dialogs** | Modal dialogs that inherit dark/light theme from the main window |
| **Shimmer Progress** | Indeterminate progress animation during async operations |

### Theme Engine

- **Dark mode** — deep dark palette (`#111113` background, `#0078D4` accent) with semitransparent borders
- **Light mode** — clean light palette (`#F5F5F5` background) with solid borders
- **Dot grid background** — tiled ellipse pattern that repaints on theme change
- **Gradient glow** — subtle accent gradient overlay with variable opacity per theme
- **50+ theme keys** — comprehensive palette covering backgrounds, text hierarchy, borders, status colours, and scroll track/thumb

### Export

#### HTML Report
- **Self-contained** — single HTML file with embedded CSS and JavaScript
- **Dark/light toggle** — built-in theme switcher with localStorage persistence
- **Sticky toolbar** — subscription badge, date range display, global search, expand/collapse all, print button
- **Summary stat cards** — Total Changes, Added, Removed, Modified, Unchanged with accent colour bars
- **Proportional progress bar** — visual distribution of changes across the resource base
- **Collapsible sections** — per change type (Removed → Added → Modified), sub-grouped by resource group
- **Property diff tables** — for modified resources, shows Property, Previous Value, Current Value columns
- **Global search** — instant filtering across resource name, type, and group with live count updates
- **Print CSS** — optimised stylesheet for paper output
- **Responsive** — adapts to narrow viewports
- **XSS-safe** — all user-supplied values are HTML-encoded

#### CSV Export
- 8 columns: ResourceName, ResourceType, ResourceGroup, ChangeType, Severity, Property, OldValue, NewValue
- Modified resources emit one row per changed property
- UTF-8 encoded with proper quote escaping

#### JSON Export
- Raw diff object with full metadata and change details
- Suitable for automation pipelines and programmatic consumption

### Named Baselines

- Save any snapshot as a named baseline (e.g. "Pre-Migration", "Security Audit Q1")
- Baselines appear in the comparison dropdowns alongside regular dated snapshots
- Stored in `baselines/` with sanitised filenames
- Compare current state against an approved baseline at any time

### Gamification

- **15 achievements** with Segoe Fluent Icons badges:

  | Icon | Achievement | Unlock Criteria |
  |------|------------|-----------------|
  | 🔑 | First Login | Sign in to Azure for the first time |
  | 📷 | First Snapshot | Take your first resource snapshot |
  | 🔄 | First Compare | Run your first snapshot comparison |
  | 📤 | First Export | Export a change report |
  | 🏷️ | First Baseline | Save a baseline snapshot |
  | 📦 | Collector | Take 5 snapshots |
  | 📚 | Archivist | Take 10 snapshots |
  | 🎨 | Painter | Toggle between dark and light themes |
  | ✅ | All Clear | Run a comparison with zero changes |
  | 💯 | Resource Rich | Snapshot 100+ resources |
  | 🏢 | Enterprise Scale | Snapshot 500+ resources |
  | 🚨 | Watchdog | Detect a critical severity change |
  | 🔀 | Multi-Tenant | Switch between 2+ subscriptions |
  | ☰ | Space Maker | Toggle the sidebar |
  | 🧹 | Clean Slate | Clear the debug log |

- Earn timestamps are persisted in `achievements.json` (ISO 8601 format)
- Toast notification + badge re-render on each unlock

### Authentication

- **Interactive sign-in** via `Connect-AzAccount` (runs in background runspace)
- **Cached token validation** — checks for an existing `AzContext` on startup and validates it with a lightweight API call before prompting for re-auth
- **WAM (Web Account Manager) disabled** for compatibility — uses browser-based auth
- **Multi-subscription** — all enabled subscriptions in the current tenant are listed in the dropdown
- **Subscription switching** — sets `Az context` in background, reloads snapshots + resources automatically
- **Tenant-scoped** — subscription list is filtered to the authenticated tenant

### Data Retention

- **Configurable retention** — enable/disable cleanup and set retention days (default: 90) in Settings
- **Auto-cleanup** — runs on startup; deletes snapshot and diff files older than the configured period
- **Manual snapshot overwrite** — prompted when today's snapshot already exists

---

## Prerequisites

| Requirement | Minimum | Purpose |
|-------------|---------|---------|
| **OS** | Windows 10 or later | WPF requires Windows desktop |
| **PowerShell** | 5.1+ (or PowerShell 7+) | Script runtime |
| **Az.Accounts** | 2.7.5+ | Azure authentication (`Connect-AzAccount`) |
| **Az.ResourceGraph** | 0.13.0+ | Fast resource queries (`Search-AzGraph`) |
| **Az.Resources** | 6.0.0+ | Fallback resource queries, RBAC |

### Installing Prerequisites

```powershell
# Install all required Az modules
Install-Module -Name 'Az.Accounts', 'Az.ResourceGraph', 'Az.Resources' -Scope CurrentUser -Force -AllowClobber
```

The tool checks prerequisites on startup (Step 1 of the splash sequence) and shows the exact install command if any module is missing.

---

## Quick Start

### Option 1: Double-click the launcher

Run **`Launch_AzChangeTracker.bat`** — it auto-detects PowerShell 7 (preferred) vs Windows PowerShell, unblocks script files, and launches the app with the console window hidden.

### Option 2: Run from PowerShell

```powershell
# PowerShell 7 (recommended)
pwsh -NoProfile -ExecutionPolicy Bypass -File .\AzChangeTracker.ps1

# Windows PowerShell 5.1
powershell.exe -NoProfile -ExecutionPolicy Bypass -Sta -File .\AzChangeTracker.ps1
```

### First Run

1. **Wait for the splash screen** — the tool checks prerequisites, validates any cached Azure token, and initialises the workspace
2. **Sign in** — click **Sign In** in the auth bar to authenticate with Azure (browser-based)
3. **Select a subscription** — choose from the dropdown that populates after sign-in
4. **Take a snapshot** — click **Snapshot Now** in the toolbar to capture all resources
5. **Wait a day (or take another snapshot)** — snapshots are dated, so you need at least two to compare
6. **Compare** — in the Changes sidebar, select two snapshots and click **Compare**
7. **Explore results** — browse the Changes tab (filter pills, search, volatile toggle), Timeline tab, and Resources tab
8. **Export** — click **Export Report** for HTML, CSV, or JSON output
9. **Save a baseline** — click **Save Baseline** to preserve the current state for future comparison

---

## File Structure

```
AzChangeTracker/
├── AzChangeTracker.ps1          # Main script (~4000 lines)
├── AzChangeTracker_UI.xaml      # WPF layout definition
├── Launch_AzChangeTracker.bat   # One-click launcher (auto-detects PS7 vs PS5.1)
├── achievements.json            # Unlocked achievements with timestamps (auto-generated)
├── user_prefs.json              # Theme, settings, default subscription (auto-generated)
├── debug.log                    # Runtime debug log (auto-generated, in %TEMP%)
├── baselines/                   # Named baseline snapshots (auto-generated)
│   └── {name}.json
├── snapshots/                   # Daily resource snapshots (auto-generated)
│   └── sub_{subscriptionId}/
│       ├── 2026-03-25.json
│       └── 2026-03-26.json
├── diffs/                       # Saved comparison results (auto-generated)
│   └── sub_{subscriptionId}/
│       └── 2026-03-25_vs_2026-03-26.json
└── reports/                     # Default export directory (auto-generated)
    └── change-report_2026-03-26.html
```

---

## Architecture

The application is a single PowerShell script (~4000 lines) with an external XAML layout, organised into 21 numbered sections:

```
┌──────────────────────────────────────────────────────────────┐
│  Section 1: Pre-load & Initialisation                        │
│    Assembly loading, DPI awareness, storage dirs, constants  │
├──────────────────────────────────────────────────────────────┤
│  Section 2: Prerequisite Module Check                        │
│    Test-Prerequisites, Get-RemediationCommand                │
├──────────────────────────────────────────────────────────────┤
│  Section 3: Thread Synchronisation Bridge                    │
│    Start-BackgroundWork → STA runspace + queue poll          │
├──────────────────────────────────────────────────────────────┤
│  Sections 4–5: XAML Load & Element References                │
│    Window.FindName() for ~80 named elements                  │
├──────────────────────────────────────────────────────────────┤
│  Sections 6–7: Global State & Write-DebugLog                 │
│    Runtime variables, 3-destination logging (console/disk/UI)│
├──────────────────────────────────────────────────────────────┤
│  Sections 8–11: Theme, Toast, Dialog, Shimmer, Splash        │
│    Dark/light palette, toast stack, themed modal, animations │
├──────────────────────────────────────────────────────────────┤
│  Sections 12–13: Navigation & Preferences                    │
│    Tab switching, sidebar sync, JSON persistence             │
├──────────────────────────────────────────────────────────────┤
│  Section 14: Snapshot Engine                                 │
│    Search-AzGraph → SHA-256 hashing → JSON save              │
├──────────────────────────────────────────────────────────────┤
│  Section 15: Diff Engine                                     │
│    Three-pass compare → canonical JSON → severity classify   │
├──────────────────────────────────────────────────────────────┤
│  Section 16: UI Update Helpers                               │
│    Snapshot list, dashboard cards, timeline, changes list,   │
│    resource type filters, resource tree, filter combos       │
├──────────────────────────────────────────────────────────────┤
│  Section 17: Export Engine                                   │
│    HTML (self-contained), CSV, JSON export formats           │
├──────────────────────────────────────────────────────────────┤
│  Sections 17b–18: Auth UI & Event Wiring                     │
│    Login/logout, subscription switch, all button handlers    │
├──────────────────────────────────────────────────────────────┤
│  Sections 19–19b: Filters & Achievement System               │
│    Combo init, 15 achievement definitions, badge rendering   │
├──────────────────────────────────────────────────────────────┤
│  Section 20: DispatcherTimer                                 │
│    50ms poll loop: background job completion → UI callback   │
├──────────────────────────────────────────────────────────────┤
│  Section 21: Startup                                         │
│    3-step splash: prerequisites → auth → workspace init      │
└──────────────────────────────────────────────────────────────┘
```

### Key Technical Decisions

| Decision | Rationale |
|----------|-----------|
| **STA runspaces** for background work | WPF requires STA; separate runspaces avoid freezing the UI during Azure API calls |
| **OneDrive PSModulePath stripping** | Prevents old Az.Accounts versions (2.x) cached in OneDrive from loading before the system-installed v5.x, avoiding assembly binding conflicts |
| **Canonical JSON with sorted keys** | Azure Resource Graph doesn't guarantee property ordering; sorted-key serialisation ensures the same resource always produces the same hash |
| **Timestamp normalisation** to `yyyy-MM-ddTHH:mm:ss.fffZ` | Azure returns varying sub-second precision (3–7 digits) across API calls, causing false diffs. Normalising to 3 digits eliminates this |
| **Volatile property exclusion** | Fields like `instanceView`, `lastModifiedTime`, `statusTimestamp` change on every API call regardless of actual resource state changes |
| **Two-timer toast pattern** | `.GetNewClosure()` in PS 5.1 has scoping pitfalls; using separate dismiss and remove timers avoids the nested-closure issue |
| **`$Global:` scope for runtime state** | Background work OnComplete callbacks run on the UI thread but can't see `$Script:` variables; `$Global:` makes them accessible |

### Data Flow

```
  Azure Subscription
        │
        ▼
  Search-AzGraph (paginated, background runspace)
        │
        ▼
  Raw Resources → Strip Volatile → SHA-256 Hash → Snapshot JSON
        │                                              │
        ▼                                              ▼
  snapshots/sub_{id}/{date}.json              baselines/{name}.json
        │              │
        ▼              ▼
  Compare-Snapshots (3-pass: Added/Removed/Modified)
        │
        ▼
  Diff Object → Severity Classification → UI Cards + Timeline
        │
        ▼
  Export: HTML Report │ CSV │ JSON
```

---

## Data Persistence

| File | Purpose | Lifecycle |
|------|---------|-----------|
| `snapshots/sub_{id}/{date}.json` | Daily resource inventory | Retention cleanup (configurable days) |
| `baselines/{name}.json` | Named "known good" snapshots | Manual management |
| `diffs/sub_{id}/{from}_vs_{to}.json` | Saved comparison results for Timeline | Retention cleanup |
| `user_prefs.json` | Theme, auto-snapshot, retention, exclusions, subscription | Saved on close and setting change |
| `achievements.json` | Achievement unlock timestamps | Permanent, merged on save |
| `%TEMP%\AzChangeTracker_debug.log` | Debug log | Auto-rotates at 2 MB |
| `reports/` | Default export directory | User-managed |

---

## Use Cases

- **Daily Change Monitoring** — take a snapshot each morning (or enable auto-snapshot) and compare against yesterday to see what changed overnight
- **Change Approval Workflow** — save a baseline before a planned change window, take a new snapshot after, and compare to verify only expected changes occurred
- **Security Auditing** — detect critical changes to NSG rules, RBAC role assignments, KeyVault access policies, and managed identities
- **Cost Governance** — track SKU changes, newly added resources, and resource removals across subscriptions
- **Migration Tracking** — compare resource state before and after a migration to validate completeness
- **Incident Investigation** — use the Timeline view to trace when specific resources were added, removed, or modified
- **Compliance Reporting** — export HTML reports to attach to compliance artifacts or share with management

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Window doesn't appear | Ensure STA mode: `powershell.exe -Sta -File .\AzChangeTracker.ps1`. The bat launcher handles this automatically |
| "Missing modules" toast | Run `Install-Module -Name 'Az.Accounts','Az.ResourceGraph','Az.Resources' -Scope CurrentUser -Force` |
| Sign-in hangs | The browser auth window may be behind other windows. Check your taskbar |
| Empty snapshot (0 resources) | Verify the subscription has resources and your account has Reader access |
| Volatile-only changes everywhere | Toggle "Hide volatile changes" ON in the Changes sidebar. These are metadata-only diffs (timestamps, instanceView) |
| Old Az module version conflict | The tool strips OneDrive paths from PSModulePath. If issues persist, update Az modules: `Update-Module Az.Accounts, Az.ResourceGraph, Az.Resources` |
| Snapshots directory growing | Enable retention cleanup in Settings (default: 90 days) |
| Window opens off-screen | Delete `user_prefs.json` to reset |

---

## Contributing

Contributions are welcome! Please open an issue to discuss your idea before submitting a pull request.

### Development Notes

- All WPF elements are bound by name using `$Window.FindName()` — check Section 5 for the full element list
- Theme colours are applied as `DynamicResource` brushes and repainted via the `$ApplyTheme` scriptblock
- Background operations **must** use `Start-BackgroundWork` — never call Azure APIs on the UI thread
- The DispatcherTimer in Section 20 is the bridge between background runspaces and the WPF dispatcher
- Every function has a `.SYNOPSIS` / `.DESCRIPTION` doc header — use `Get-Help` for details

---

## Author

**Anton Romanyuk**

---

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.
