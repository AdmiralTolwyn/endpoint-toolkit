# ADMX Policy Comparer

A modern WPF desktop application for comparing Microsoft Group Policy ADMX/ADML administrative templates across Windows versions. Built entirely in PowerShell with a polished dark/light UI, word-level diff highlighting, rich registry metadata extraction, and self-contained HTML/CSV export.

![Version](https://img.shields.io/badge/version-0.1.0--alpha-blue)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![Platform](https://img.shields.io/badge/platform-Windows%2010%2B-0078D4?logo=windows)
![.NET](https://img.shields.io/badge/.NET%20Framework-4.7.2%2B-purple)
![License](https://img.shields.io/badge/license-MIT-green)

---

## Why This Tool?

Every time Microsoft releases a new Windows feature update, the Group Policy Administrative Templates (ADMX/ADML) change — policies are added, removed, modified, or reorganised. Tracking those changes manually across hundreds of XML files is tedious and error-prone.

**ADMX Policy Comparer** automates this entirely:

- **Point it at two template versions** and get a full colour-coded diff in seconds
- **See exactly which policies changed** — down to the word level, with strikethrough/underline highlighting
- **Get the registry impact** — every change is enriched with the registry key, value name, type, scope, and enabled/disabled values from the ADMX definitions
- **Export shareable reports** — self-contained HTML reports with search, filtering, and dark/light theme toggle, or CSV for further analysis
- **Track your work** — comparison history, daily streaks, and 24 unlockable achievements

---

## Features

### Comparison Engine

- **File-by-file ADML diffing** — detects Added, Removed, and Modified strings across all `.adml` files
- **Entire file detection** — flags files that were wholly added or removed between versions
- **Trivial change filter** — automatically skips whitespace-only changes that have no policy impact
- **Version-aware sorting** — correctly orders Windows versions from 1507 through 26H2
- **Word-level diff (LCS)** — uses a Longest Common Subsequence algorithm on whitespace-split tokens to produce precise inline diffs with deleted/added run highlighting
- **Performance guard** — falls back to plain diff for texts exceeding 500 words to avoid UI stalls

### ADMX/ADML Parsing

- **ADML string extraction** — parses `<stringTable>` entries from localised ADML files
- **ADMX policy metadata** — for each policy, extracts:
  - Parent category
  - Registry key and value name
  - Policy class/scope (Machine, User, or Both)
  - Value type (DWORD, REG_SZ, REG_MULTI_SZ, enum, composite)
  - Enabled/Disabled values
  - Elements tree: enumerations with choices, decimal ranges, boolean, text, list, and multiText types
- **Policy grouping** — derives a policy group from the ADMX `displayName` string ID and maps it back to ADML results for logical grouping in the DataGrid and HTML export

### User Interface

| Area | Description |
|------|-------------|
| **Dashboard** | Animated stat cards (Comparisons, Changes, Files, Streak), recent comparisons list, quick-action buttons |
| **Compare** | Version picker with auto-scanned combo boxes, comparison progress bar with shimmer animation, real-time log output |
| **Results** | DataGrid with filter pills (All / Added / Removed / Modified / File Added / File Removed), live search, collapsible file sidebar, policy grouping toggle, breakdown chart |
| **Settings** | Theme toggle, animation toggle, debug overlay, base path / export path configuration, achievement badges with unlock progress bar, reset button |
| **Detail Pane** | Resizable split pane showing side-by-side old/new values with word-level diff runs, registry info bar with key path, value name, scope badge, and value type details, copy buttons for Old value / New value / String ID |
| **Debug Console** | Collapsible bottom panel with colour-coded activity log (RichTextBox), vertical minimap, horizontal heatmap, line counter, clear/resize/hide controls |
| **Title Bar** | Custom chrome with help button (keyboard shortcuts), theme toggle, minimize/maximize/close |
| **Toast Notifications** | Slide-in notifications for success, warning, error, and info events with accent-coloured sidebar and icon |

### Theme Engine

- **Dark mode** — deep dark palette (`#111113` background, `#0078D4` accent) optimised for extended use
- **Light mode** — clean light palette (`#F5F5F5` background) with the same accent colour
- **Animated transitions** — smooth opacity fade-out/fade-in when toggling themes
- **Consistent colour-coding** — Added (green), Removed (red/orange), Modified (amber) across both themes
- **Log level colours** — INFO, SUCCESS, WARN, ERROR, DEBUG, STEP, and SYSTEM levels each have theme-aware colours

### Export

#### HTML Report
- **Self-contained** — single HTML file with embedded CSS, JavaScript, and all data
- **Dark/light toggle** — built-in theme switcher in the report header
- **Collapsible sidebar** — file list with change counts
- **Filter pills** — click to filter by change type (All / Added / Removed / Modified)
- **Live search** — instant filter across all rows
- **Expand/collapse** — per-file collapsible sections with policy sub-grouping
- **Inline diff highlighting** — word-level `<span>` highlighting with `.diff-del` (strikethrough red) and `.diff-add` (underline green)
- **Registry metadata column** — value type and scope badges per row
- **Badge summary** — total counts by change type in the header
- **Keyboard shortcut** — `Ctrl+F` focuses search in the HTML report

#### CSV Export
- 9 columns: ChangeType, FileName, Category, StringId, RegistryKey, ValueName, Scope, OldValue, NewValue
- Respects the currently active filter — exports only visible rows
- UTF-8 encoded with proper quote escaping

### Gamification

- **24 achievements** with emoji badges, unlock timestamps, and confetti celebrations:

  | Badge | Achievement | Unlock Criteria |
  |-------|------------|-----------------|
  | 🚀 | First Steps | Complete your first comparison |
  | 🌊 | Deep Diver | Complete 5 comparisons |
  | 🧙 | Policy Guru | Complete 10 comparisons |
  | 🎖️ | Veteran Admin | Complete 25 comparisons |
  | 🏆 | Centurion | Complete 100 comparisons |
  | 🦅 | Eagle Eye | Find 100+ changes in one comparison |
  | 🪡 | Needle Finder | Find exactly 1 change in a comparison |
  | ✨ | Clean Slate | Compare with zero differences |
  | ⚡ | Speed Demon | Complete comparison in under 5 seconds |
  | 🦉 | Night Owl | Compare between midnight and 5 AM |
  | 🐦 | Early Bird | Compare between 5 AM and 7 AM |
  | 🏖️ | Weekend Warrior | Compare on a weekend |
  | 📤 | Export Master | Export your first report |
  | 📚 | Batch Reporter | Export 5 reports |
  | 🔍 | Filter Pro | Use every filter type |
  | 🎨 | Theme Switcher | Toggle between dark and light mode |
  | 🗂️ | Explorer | Browse to a custom policy folder |
  | 📖 | Bookworm | Find 50+ modified strings in one run |
  | 🧹 | Spring Cleaning | Find 50+ removed strings in one run |
  | 🖌️ | Fresh Paint | Find 50+ added strings in one run |
  | 🔥 | On a Roll | 3-day comparison streak |
  | 🔥 | On Fire | 7-day comparison streak |
  | 💪 | Unstoppable | 30-day comparison streak |
  | 👑 | Completionist | Unlock 20 achievements |

- **Daily streak tracker** — consecutive-day counter with total comparison tally
- **Comparison history** — 50-entry ring buffer with timestamps, version info, and change counts; results persisted as JSON for reload
- **Progress bar** — visual unlock progress in the Settings tab

### Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `F1` | Show keyboard shortcuts help |
| `F5` | Run comparison |
| `Ctrl+E` | Export HTML report |
| `Ctrl+F` | Focus search box (switches to Results tab) |
| `Ctrl+B` | Toggle sidebar |
| `` Ctrl+` `` | Toggle debug console |
| `Ctrl+C` | Copy selected row (when not in text field) |
| `1` – `4` | Switch tabs: Dashboard, Compare, Results, Settings |

---

## Prerequisites

| Requirement | Minimum |
|-------------|---------|
| **OS** | Windows 10 or later |
| **PowerShell** | 5.1+ (ships with Windows 10/11) |
| **.NET Framework** | 4.7.2+ |
| **ADMX Templates** | Two or more versions of Microsoft Administrative Templates downloaded and extracted |

### Obtaining ADMX Templates

Download the Administrative Templates (.admx) for the Windows versions you want to compare from the Microsoft Download Center:

- [Windows 10 (22H2)](https://www.microsoft.com/en-us/download/details.aspx?id=104042)
- [Windows 11 (23H2)](https://www.microsoft.com/en-us/download/details.aspx?id=105667)
- [Windows 11 (24H2)](https://www.microsoft.com/en-us/download/details.aspx?id=106255)
- [Windows 11 (25H2)](https://www.microsoft.com/en-us/download/details.aspx?id=106384)

Extract each into its own folder so the structure looks like:

```
C:\ADMXTemplates\
├── Windows 11 Sep 2024 Update (24H2)\
│   └── PolicyDefinitions\
│       ├── en-US\
│       │   ├── ActiveXInstallService.adml
│       │   ├── AppxPackageManager.adml
│       │   └── ...
│       ├── ActiveXInstallService.admx
│       └── ...
├── Windows 11 Sep 2025 Update (25H2)\
│   └── PolicyDefinitions\
│       ├── en-US\
│       │   └── ...
│       └── ...
└── ...
```

The tool auto-discovers any subfolder that contains `PolicyDefinitions\en-US`.

---

## Quick Start

### Option 1: Double-click the launcher

Run `Launch_ADMXPolicyComparer.bat` — it starts PowerShell with the correct flags (`-STA`, `-NoProfile`, `-ExecutionPolicy Bypass`).

### Option 2: Run from PowerShell

```powershell
# Must run in STA mode for WPF
powershell.exe -Sta -NoProfile -ExecutionPolicy Bypass -File .\ADMXPolicyComparer.ps1
```

### First Run

1. **Set the base path** — in the Compare tab or Settings, browse to the folder containing your extracted ADMX template versions
2. **Scan versions** — click **Scan Versions** to discover available template sets
3. **Pick two versions** — select an Older and Newer version from the dropdowns
4. **Run comparison** — click **Compare** (or press `F5`)
5. **Explore results** — use filter pills, search, file sidebar, and policy grouping to drill in
6. **Click a row** — the detail pane shows side-by-side word-level diffs with registry metadata
7. **Export** — click **Export HTML** (`Ctrl+E`) or **Export CSV** for a shareable report

---

## File Structure

```
ADMXPolicyComparer/
├── ADMXPolicyComparer.ps1       # Main script (~4200 lines)
├── ADMXPolicyComparer_UI.xaml   # WPF layout definition
├── Launch_ADMXPolicyComparer.bat # One-click launcher
├── achievements.json            # Unlocked achievements with timestamps (auto-generated)
├── comparison_history.json      # Recent comparisons ring buffer (auto-generated)
├── streak.json                  # Daily streak counter (auto-generated)
├── user_prefs.json              # Window geometry, theme, paths (auto-generated)
├── debug.log                    # Debug log (rotates at 2 MB, auto-generated)
└── results/                     # Cached comparison results as JSON (auto-generated)
    ├── {uuid-1}.json
    └── {uuid-2}.json
```

---

## Architecture

The application is structured as a single-file PowerShell script with an external XAML layout:

```
┌─────────────────────────────────────────────────────────────┐
│                     Pre-flight & Globals                     │
│         Assembly loading, DPI awareness, theme palettes      │
├─────────────────────────────────────────────────────────────┤
│                    XAML Loading & UI Wiring                   │
│        Window.FindName() for ~90 named elements              │
├─────────────────────────────────────────────────────────────┤
│                       Theme Engine                           │
│     Set-Theme applies hex palette to WPF DynamicResources    │
├─────────────────────────────────────────────────────────────┤
│                     Debug Console                            │
│  Write-DebugLog → Console + RichTextBox + disk + minimap     │
├─────────────────────────────────────────────────────────────┤
│                  Persistence Layer                            │
│    Preferences · History · Streak · Achievements (JSON)      │
├─────────────────────────────────────────────────────────────┤
│                     Animations                               │
│  Toast · Confetti · CountUp · Breathe · Fade · Tab transitions│
├─────────────────────────────────────────────────────────────┤
│                  Comparison Engine                            │
│  Version scan → ADML parse → ADMX metadata → file-by-file   │
│  diff → trivial filter → result collection → achievements    │
├─────────────────────────────────────────────────────────────┤
│                 Filtering & Search                            │
│  CollectionView filter · policy grouping · file sidebar      │
├─────────────────────────────────────────────────────────────┤
│                     Export                                    │
│         HTML (self-contained) · CSV (filtered)               │
├─────────────────────────────────────────────────────────────┤
│                   Results UI                                  │
│  DataGrid · detail pane · word-level LCS diff · registry bar │
│  breakdown chart · empty state · status bar                  │
├─────────────────────────────────────────────────────────────┤
│                 Window Lifecycle                              │
│       Load · Close · Keyboard shortcuts · Micro-interactions │
└─────────────────────────────────────────────────────────────┘
```

### Key Components

| Component | Functions | Purpose |
|-----------|-----------|---------|
| **Version Scanner** | `Scan-Versions` | Enumerates template directories, sorts by Windows version code |
| **ADML Parser** | `Get-AdmlStrings` | Extracts string ID → localised text from ADML XML |
| **ADMX Parser** | `Get-AdmxCategories` | Extracts registry keys, value types, scope, enabled/disabled values, element trees from ADMX XML |
| **Comparison** | `Run-Comparison` | Orchestrates the full diff pipeline: enumerate files → parse → compare → enrich → store |
| **Word Diff** | `Get-WordDiffRuns` | LCS-based word-level diff producing tagged run arrays |
| **HTML Diff** | `Get-HtmlDiffSpans` | Wraps word diff in `<span class="diff-del/add">` for HTML export |
| **Filter System** | `Apply-Filters` | CollectionView predicate combining type filter + text search + file filter |
| **Detail Pane** | `Update-DetailPane` | Renders side-by-side RichTextBox with coloured diff runs and registry info bar |
| **Breakdown Chart** | `Update-BreakdownChart` | Proportional Grid bar chart showing Added/Removed/Modified percentages |
| **Dashboard** | `Update-Dashboard` | Refreshes stat cards, recent comparisons list, export history combo |
| **Theme** | `Set-Theme`, `Invoke-ThemeTransition` | Applies colour palette with animated fade transition |
| **Achievements** | `Check-ComparisonAchievements`, `Unlock-Achievement` | Evaluates 24 unlock criteria, persists with timestamps, triggers toast + confetti |

---

## Data Persistence

All state is stored as JSON files alongside the script — no database, no external services, fully portable:

| File | Purpose | Retention |
|------|---------|-----------|
| `user_prefs.json` | Window position/size, theme, animation toggle, base/export paths | Saved on close |
| `comparison_history.json` | Last 50 comparisons with metadata (versions, counts, timestamp, UUID) | Ring buffer |
| `results/{id}.json` | Full comparison results per history entry | Pruned when history entry deleted |
| `achievements.json` | Achievement unlock timestamps (ISO 8601) | Permanent |
| `streak.json` | Current streak length, last comparison date, total count | Updated per comparison |
| `debug.log` | Timestamped debug log | Rotates at 2 MB (keeps `.prev` backup) |

---

## Use Cases

- **Windows Update Impact Analysis** — compare ADMX templates before and after a feature update to identify new, changed, or removed Group Policy settings
- **GPO Migration Planning** — when upgrading from Windows 10 to Windows 11, identify which policies have changed registry paths, value types, or scope
- **Security Baseline Auditing** — verify which policies were added or removed between template releases to keep security baselines current
- **Change Documentation** — generate HTML reports to attach to change requests or share with your team
- **Policy Inventory** — use the registry metadata enrichment to understand exactly where each policy writes in the registry, what value type it uses, and what enabled/disabled values it sets

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Window doesn't appear | Ensure you're running in STA mode: `powershell.exe -Sta -File .\ADMXPolicyComparer.ps1` |
| "Base path not found" | Set the base path to the parent folder that contains your version subfolders |
| No versions detected | Each version folder must contain `PolicyDefinitions\en-US\` with `.adml` files |
| Comparison is slow | Large template sets (1000+ files) take longer; the progress bar shows real-time status |
| Export button greyed out | Run a comparison first — export requires results in memory |
| Window opens off-screen | Delete `user_prefs.json` to reset saved window position |
| Debug log growing large | The log auto-rotates at 2 MB; you can also clear it from the debug console |

---

## Contributing

Contributions are welcome! Please open an issue to discuss your idea before submitting a pull request.

### Development Notes

- The application is a single PowerShell script (`ADMXPolicyComparer.ps1`) with an external XAML layout file
- All WPF elements are wired by name using `$Window.FindName()` and stored in a `$ui` hashtable
- Theme colours are applied as `DynamicResource` brushes so they update globally when toggled
- The debug console (`Ctrl+``) is invaluable during development — it shows timestamped logs at millisecond precision with log-level colour coding

---

## Author

**Anton Romanyuk**

---

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.
