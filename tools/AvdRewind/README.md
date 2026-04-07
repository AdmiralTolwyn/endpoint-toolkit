# AVD Rewind

**Backup, visualize, and selectively restore Azure Virtual Desktop environments.**

AVD Rewind is a PowerShell/WPF desktop application that captures the full AVD control-plane topology — host pools, application groups, workspaces, session hosts, scaling plans, RBAC assignments, and diagnostic settings — visualizes it as an interactive tree, and enables selective or full restore for accidental deletion recovery or cross-region migration.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![WPF](https://img.shields.io/badge/GUI-WPF-blueviolet)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)

---

## Why AVD Rewind?

Azure Virtual Desktop doesn't have a native "undo" button. When a host pool, application group, or workspace is accidentally deleted — or when a misconfigured update changes critical properties — the only recovery path is manual ARM template redeploys or PowerShell scripting against documentation from memory. AVD Rewind solves this by:

- **Capturing everything**: Periodic topology snapshots store the complete control-plane state as JSON with SHA256 integrity verification
- **Showing what changed**: The diff engine compares any backup against the current live state, classifying every object as Added, Removed, Modified, or Unchanged
- **Restoring selectively**: A dependency-ordered restore manifest lets you pick exactly which objects to recreate or revert, with dry-run validation before committing
- **Supporting cross-region migration**: Resource group remapping and location override let you restore an environment into a different region or subscription

---

## Features

### Topology Discovery
- Enumerates workspaces, host pools, application groups (Desktop/RemoteApp), applications, RBAC role assignments, session hosts (with VM metadata), scaling plans, and diagnostic settings
- Batched RBAC and VM caching per resource group to minimize Azure API calls
- Progress reporting during discovery with per-host-pool status updates

### Interactive Topology Tree
- Hierarchical TreeView: Host Pools → App Groups → Applications/RBAC, Session Hosts (with status dots), Scaling Plans; Workspaces → App Group References
- Rich node headers with Segoe Fluent icons, type badges (Pooled, Desktop, RemoteApp), location hints, and color-coded status indicators
- Live text search with recursive filtering and ancestor expansion
- Resource group dropdown filter

### Backup Engine
- JSON serialization at depth 20 with schema version, tool version, timestamp, and optional labels
- SHA256 integrity sidecar files for corruption detection
- Automatic backup on discovery (configurable)
- Configurable retention with cleanup (deletes backups older than N days)
- Verified/unverified badges in the backup list

### Diff Engine
- Flattens topologies to resource-ID-keyed maps for accurate comparison
- Tracks Added, Removed, Modified, and Unchanged states
- Ignores volatile properties (LastHeartBeat, Status, UpdateState) to reduce false positives
- Per-property change summaries with human-readable before/after values (e.g., tag count changes, app group reference diffs)
- Session host count change detection per host pool
- Drift banner on the Dashboard when changes are detected

### Restore Engine
- Dependency-ordered manifest: ResourceGroup → HostPool → AppGroup → Application → Workspace → RBAC → ScalingPlan → DiagnosticSetting → SessionHost
- Selective restore via tri-state checkboxes (parent toggles all children)
- Dry-run validation: checks all required fields without making API calls
- Cross-region support: location override for different-region restores
- Resource group remapping: editable source→target RG textbox grid
- Auto-create missing resource groups before restore
- Confirmation dialog with typed "RESTORE" safety gate
- Export restore plans as JSON for auditing or CI/CD pipelines

### Session Host Re-Registration
- Generates 4-hour registration tokens via New-AzWvdRegistrationInfo
- Pushes re-registration scripts to VMs via Invoke-AzVMRunCommand
- VM pattern discovery: finds session host VMs matching naming conventions beyond backup data

### Theme System
- Dark mode (default) and Light mode with 40+ color keys each
- Instant theme switching with procedural background regeneration
- All UI elements use dynamic resource references for consistent theming

### Achievements
15 gamification achievements to encourage thorough usage:

| Achievement | Icon | Criteria |
|---|---|---|
| Explorer | 🔍 | First AVD discovery |
| Archivist | 💾 | First backup saved |
| Detective | 🔎 | First topology comparison |
| Cautious One | 🧪 | First dry-run restore |
| Time Traveler | ⏭ | First actual restore |
| Flawless | ✨ | 5+ items restored with zero failures |
| Speed Demon | ⚡ | Discovery completed in under 10 seconds |
| Hoarder | 📦 | 10+ backups accumulated |
| Night Owl | 🦉 | Used AVD Rewind after midnight |
| Indecisive | 🎨 | Toggled theme 5+ times in one session |
| Multi-Tenant | 🌐 | Switched between 3+ subscriptions |
| Whale | 🐋 | Discovered 10+ host pools in one subscription |
| Clean Slate | 🧹 | Retention cleanup removed old backups |
| Exporter | 📤 | Exported topology or restore plan to file |
| Dedicated | ⭐ | Opened AVD Rewind 10+ times |

---

## Quick Start

### Prerequisites

| Component | Minimum | Purpose |
|---|---|---|
| Windows 10/11 or Windows Server 2016+ | — | WPF host |
| PowerShell | 5.1+ (7+ recommended) | Runtime |
| Az.Accounts | 2.7.5 | Authentication, subscription context |
| Az.DesktopVirtualization | 4.0.0 | AVD control-plane cmdlets |
| Az.Resources | 6.0.0 | RBAC role assignments, resource groups |
| Az.Compute | 5.0.0 | VM metadata, Invoke-AzVMRunCommand |
| Az.Network | 5.0.0 | NIC/subnet/NSG mapping |
| Az.Monitor | 4.0.0 | Diagnostic settings backup/restore |

Install all modules at once:

```powershell
Install-Module Az.Accounts -MinimumVersion 2.7.5 -Force
Install-Module Az.DesktopVirtualization -MinimumVersion 4.0.0 -Force
Install-Module Az.Resources -MinimumVersion 6.0.0 -Force
Install-Module Az.Compute -MinimumVersion 5.0.0 -Force
Install-Module Az.Network -MinimumVersion 5.0.0 -Force
Install-Module Az.Monitor -MinimumVersion 4.0.0 -Force
```

### Launch

**Option 1** — Double-click `Launch_AvdRewind.bat` (auto-detects PowerShell 7, falls back to Windows PowerShell)

**Option 2** — From a PowerShell console:
```powershell
cd path\to\AvdRewind
.\AvdRewind.ps1
```

### First Run

1. The splash screen validates prerequisites (green dots = installed, red = missing)
2. Click **Login** in the auth bar to authenticate with Azure
3. Select a subscription from the dropdown
4. The Dashboard tab shows environment stats
5. Click **Refresh Topology** to run your first discovery
6. Click **Backup Now** to save your first snapshot
7. Make a change in Azure, then click **Compare with Current** in the Backups tab to see drift

---

## UI Tabs

### Dashboard
Five stat cards (Host Pools, App Groups, Workspaces, Session Hosts, Scaling Plans), drift detection banner, and quick-action buttons for backup, restore, and compare operations. Sidebar shows host pool count, session host count, and last backup timestamp.

### Topology
Interactive TreeView showing the full AVD hierarchy with rich headers, type badges, and status indicators. Sidebar provides resource group filtering and live text search with recursive node visibility.

### Backups
Chronological list of all saved backups with SHA256 verification badges (✅/⚠️), resource counts, file sizes, and subscription names. Select a backup and click **Compare with Current** to open the diff overlay showing Added/Removed/Modified/Unchanged counts and details.

### Restore
Full restore workflow:
1. Select a backup source from the dropdown
2. Choose same-region or different-region restore (with target region selector)
3. Review and edit resource group mappings (source → target)
4. Enable auto-create resource groups and/or VM discovery if needed
5. Review the restore manifest tree with per-item action badges (Create/Update) and property-level diffs
6. Run **Dry Run** to validate all entries
7. Execute **Restore** (requires typing "RESTORE" to confirm)
8. **Export Plan** saves the manifest as JSON for auditing

### Settings
Theme toggle (dark/light), auto-backup on launch, backup retention (enable + configurable days), prerequisite module grid with status badges and copy-to-clipboard install command, achievement badge grid with unlock progress, and reset-to-defaults button.

---

## Backup File Structure

```
backups/
├── backup_<subId>_20260326_193458.json        # Topology snapshot
├── backup_<subId>_20260326_193458.json.sha256  # SHA256 integrity sidecar
├── backup_<subId>_20260327_005456.json
└── backup_<subId>_20260327_005456.json.sha256
```

Each backup JSON contains:

```json
{
  "schemaVersion": "1.0",
  "toolVersion": "0.1.0-alpha",
  "timestamp": "2026-03-27T12:34:56.789Z",
  "subscriptionId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "subscriptionName": "Contoso-Prod",
  "label": "Pre-migration snapshot",
  "topology": {
    "schemaVersion": "1.0",
    "timestamp": "...",
    "workspaces": [ { "resourceId", "name", "location", "applicationGroupReferences", ... } ],
    "hostPools": [
      {
        "name", "location", "hostPoolType", "loadBalancerType", "maxSessionLimit", ...,
        "appGroups": [ { "applicationGroupType", "applications", "rbacAssignments", ... } ],
        "sessionHosts": [ { "name", "status", "vmMetadata", ... } ],
        "scalingPlans": [ { "schedules", "hostPoolReferences", ... } ],
        "diagnosticSettings": [ { "workspaceId", "logs", ... } ]
      }
    ],
    "resourceGroups": [ "rg-avd-prod-weu", ... ]
  }
}
```

---

## Diff Engine — Change Detection

The diff engine flattens both topologies into resource-ID-keyed maps and compares per-property:

| Category | Meaning |
|---|---|
| **Added** | Exists in current but not in baseline |
| **Removed** | Exists in baseline but not in current (restorable) |
| **Modified** | Key properties differ (volatile fields excluded) |
| **Unchanged** | Identical on all compared properties |

**Volatile properties** (always ignored): `LastHeartBeat`, `Status`, `UpdateState`, `Timestamp`, `provisioningState`

**Child collections** compared separately: `SessionHosts`, `AppGroups`, `ScalingPlans`, `DiagnosticSettings`, `Applications`, `RbacAssignments`

---

## Restore Workflow — Dependency Order

The restore engine respects the Azure resource dependency graph:

```
1. Resource Groups      (must exist first)
2. Host Pools           (containers for everything)
3. Application Groups   (reference host pools)
4. Applications         (belong to RemoteApp groups)
5. Workspaces           (reference app groups)
6. RBAC Assignments     (scope to app groups/workspaces)
7. Scaling Plans        (reference host pools)
8. Diagnostic Settings  (scope to host pools)
9. Session Hosts        (re-registration via VM run commands)
```

---

## Architecture

```
AvdRewind/
├── AvdRewind.ps1           # Main application (~4400 lines, 25 sections, 50+ functions)
├── AvdRewind_UI.xaml        # WPF layout with dark/light theme resources
├── Launch_AvdRewind.bat     # Auto-detect pwsh/powershell launcher
├── user_prefs.json          # Theme, retention, subscription, achievements
├── backups/                 # Versioned topology snapshots + SHA256 sidecars
└── reports/                 # Exported topology reports (JSON/CSV)
```

### Code Organization (25 Sections)

| # | Section | Purpose |
|---|---|---|
| 1 | Pre-load & Initialization | Assemblies, DPI, directories, constants |
| 2 | Prerequisite Module Check | Az module validation and status grid |
| 3 | Thread Synchronization Bridge | Background runspace launcher with queued completion |
| 4 | XAML GUI Load | Parse XAML, strip designer attributes, build Window |
| 5 | Element References | 100+ FindName() bindings |
| 6 | Global State | Auth, topology cache, UI flags, logging buffers |
| 7 | Write-DebugLog | 3-destination logging (console, disk, UI) |
| 8 | Theme System | Dark/Light palettes, procedural backgrounds |
| 9 | Toast Notifications | Auto-dismissing overlays with fade animations |
| 10 | Show-ThemedDialog | Modal dialogs with optional input validation |
| 11 | Shimmer & Progress | Determinate/indeterminate progress bar |
| 11b | Splash Step Helpers | Startup splash with 3-step indicators |
| 12 | Tab & Navigation | 5-tab switcher with icon rail and sidebar panels |
| 13 | User Preferences | JSON-based persistence and retention cleanup |
| 14 | AVD Discovery Engine | Full topology enumeration with batched caching |
| 15 | Backup Engine | JSON + SHA256 serialization and loading |
| 16 | Diff Engine | Property-level topology comparison |
| 17 | Restore Engine | Dependency-ordered manifest execution |
| 18 | Session Host Registration | Token generation and VM re-registration |
| 18b | RG Mapping & VM Discovery | Cross-environment remapping and pattern matching |
| 19 | UI Update Helpers | Dashboard cards, topology tree, backup list, restore tree |
| 20 | Export Engine | JSON/CSV topology export |
| 20b | Auth UI Helpers | Azure login, subscription dropdown, token management |
| 21 | Event Wiring | 200+ WPF event handlers |
| 22 | Filter & Search | Live topology search with recursive filtering |
| 23 | Achievement System | 15 achievements with badge grid and unlock toasts |
| 24 | Dispatcher Timer | 50ms polling for background job completion |
| 25 | Startup Sequence | Splash, prerequisites, auto-login, Window.ShowDialog |

---

## Data Flow

```
Azure Subscription
       │
       ▼
  Get-AvdTopology  ──────────────────►  $Global:CurrentTopology
  (background runspace)                        │
       │                                       ▼
       ▼                              Update-DashboardCards
  Save-AvdBackup                      Update-TopologyTree
       │                                       │
       ▼                                       ▼
  backups/*.json  ◄───  Get-BackupList  ◄── Update-BackupList
       │
       ▼
  Compare-AvdTopology ($Baseline, $Current)
       │
       ▼
  Build-RestoreManifest ──► Apply-RgMappings ──► Invoke-FullRestore
                                                       │
                                                       ▼
                                              Azure ARM API calls
                                              (New/Update-AzWvd*)
```

---

## Technical Decisions

| Decision | Rationale |
|---|---|
| Single .ps1 monolith | Zero-dependency deployment — copy folder and run |
| STA runspaces for background work | WPF requires STA; separate runspaces prevent UI freezes |
| PSModulePath cleanup | Strips OneDrive user-profile module paths to prevent old Az.Accounts versions from loading |
| SHA256 sidecar files | Detects backup corruption without external tooling |
| Batched RBAC/VM caching | Reduces Azure API calls from O(n) per app-group to O(1) per resource-group |
| Volatile property exclusion | Prevents false-positive diffs from transient session host state |
| Typed confirmation dialog | "RESTORE" text match prevents accidental production changes |
| Resource group remapping | Enables cross-environment cloning without modifying backup data |

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| Missing module error on startup | Run the install commands shown in the Settings tab prerequisite grid |
| "XAML file not found" | Ensure `AvdRewind_UI.xaml` is in the same folder as `AvdRewind.ps1` |
| Discovery returns 0 resources | Verify you have Reader/Contributor role on the subscription |
| Diff shows unexpected changes | Check if changes are in volatile properties (try re-running discovery) |
| Backup verification fails (⚠️) | The backup JSON was modified outside AVD Rewind — re-take the backup |
| Restore fails with "Missing ResourceGroupName" | Enable **Auto-Create Resource Groups** in the Restore sidebar |
| Session host re-registration fails | Ensure the VM is running and you have Contributor access on the VM RG |
| Theme doesn't apply to dialogs | Dialogs inherit Window resources — restart the app if resources are stale |

---

## Logging

AVD Rewind logs to three destinations simultaneously:

1. **PowerShell Console** — All messages via Write-Host (DarkGray)
2. **Disk File** — `%TEMP%\AvdRewind_debug.log` (auto-rotates at 2 MB, keeps `.old` backup)
3. **Activity Log Panel** — Color-coded RichTextBox in the right sidebar (ring buffer, 500 lines)

Log levels: `INFO` (default), `DEBUG` (verbose, hidden unless debug overlay enabled), `WARN` (amber), `ERROR` (red), `SUCCESS` (green).

---

## Author

**Anton Romanyuk**

## License

See repository root for license information.
