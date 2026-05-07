# Device Decommissioner

A WPF-based PowerShell GUI that removes a single device from **Active
Directory**, **Entra ID**, **Microsoft Intune**, **Windows Autopilot**, and
**Configuration Manager (SCCM)** in one guided workflow.

The destructive path is intentionally narrow: one device per run, look-up
first, per-system review, type-the-name confirmation, then removal. A
separate read-only **Inventory check** view supports multi-device discovery
without any deletion.

> Visual style mirrors the broader toolkit (WMM, AIB Pipeline Creator):
> custom WindowChrome title bar with EXPERIMENTAL badge, dark/light themes
> with dotted-grid background, icon rail + collapsible sidebar, gradient
> accent buttons, color-coded log panel, toasts, structured modals.

---

## Table of contents

1. [Quick start](#quick-start)
2. [Prerequisites](#prerequisites)
3. [Workflow](#workflow)
4. [Pattern matching & wildcards](#pattern-matching--wildcards)
5. [Cancel a running scan](#cancel-a-running-scan)
6. [Dry-run mode](#dry-run-mode)
7. [Inventory check](#inventory-check)
8. [Decommission history](#decommission-history)
9. [Achievements](#achievements)
10. [Re-check after decommission](#re-check-after-decommission)
11. [Settings](#settings)
12. [Stored credentials (DPAPI)](#stored-credentials-dpapi)
13. [Files written next to the script](#files-written-next-to-the-script)
14. [Output panel](#output-panel)
15. [Keyboard shortcuts](#keyboard-shortcuts)
16. [What the destructive cmdlets actually do](#what-the-destructive-cmdlets-actually-do)
17. [Security notes](#security-notes)
18. [Troubleshooting](#troubleshooting)
19. [Architecture](#architecture)
20. [Limitations & non-goals](#limitations--non-goals)

---

## Quick start

```cmd
:: From this folder
Launch_DeviceDecommissioner.bat
```

Or directly:

```powershell
powershell.exe -ExecutionPolicy Bypass -STA -File .\DeviceDecommissioner.ps1
```

The window opens centered. On first launch every credential indicator is
amber ("Not configured") and the **Decommission** button is disabled until a
lookup completes.

---

## Prerequisites

| System    | PowerShell module                                                          | Auth model                                                       |
| --------- | -------------------------------------------------------------------------- | ---------------------------------------------------------------- |
| AD        | `ActiveDirectory` (RSAT)                                                   | Stored DPAPI credentials *or* the current Windows user           |
| Entra ID  | `Microsoft.Graph.Identity.DirectoryManagement`                             | Interactive `Connect-MgGraph` (browser/WAM, UI thread)           |
| Intune    | `Microsoft.Graph.DeviceManagement`                                         | Same Graph context as Entra                                      |
| Autopilot | `Microsoft.Graph.DeviceManagement.Enrollment`                              | Same Graph context; needs `DeviceManagementServiceConfig.ReadWrite.All` |
| SCCM      | `ConfigurationManager.psd1` (Endpoint Manager Console / SMS_ADMIN_UI_PATH) | Stored DPAPI credentials *or* the current Windows user           |

PowerShell 5.1 or 7 on Windows. The launcher prefers `pwsh` if available
and falls back to `powershell.exe`. The script must run **STA** for WPF —
the launcher handles this; if you start it by hand, include `-STA`.

### Prerequisite banner

On launch the tool checks for missing modules and shows a yellow banner at
the top of the content area listing what's absent. The banner has two
actions:

* **Install Microsoft.Graph** — one-click `Install-Module -Scope CurrentUser`
  for the three Graph modules (no admin required). Visible only when at
  least one Graph module is missing.
* **How do I fix the rest?** — opens a modal with copy-pasteable PowerShell
  snippets (Graph SDK install, `Add-WindowsCapability` for RSAT, SCCM
  ConsoleSetup path). Useful when the tool runs without admin or when
  PSGallery is blocked.

If a module is missing, the corresponding card shows **Module missing** at
lookup time and that system is silently excluded from decommission. The
other systems still work.

---

## Workflow

```
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────────┐
│  1. Look up     │ -> │  2. Review 5    │ -> │  3. Confirm by      │
│  device by      │    │  discovery      │    │  typing device name │
│  hostname / id  │    │  cards          │    │  (or Run dry-run)   │
└─────────────────┘    └─────────────────┘    └─────────────────────┘
```

### 1. Look up

* Type a hostname (sAMAccountName / Intune deviceName / Entra displayName /
  SCCM ResourceName / Autopilot SerialNumber), an Entra device ObjectId
  GUID, **or** a wildcard pattern (see
  [Pattern matching](#pattern-matching--wildcards)). Press **Enter** or
  click **Look up**.
* Three background runspaces start in parallel — one for AD, one shared by
  Entra+Intune+Autopilot (one Graph context for all three), one for SCCM.
  The status pulse turns amber, the global progress bar appears in its own
  row, each card transitions Idle → **Searching…** → final state.
* A red **Cancel** button replaces Look up while a scan is running (see
  [Cancel a running scan](#cancel-a-running-scan)).
* Card states: **Found** (green), **Not found** (gray), **Error** (red,
  details on the card), **Module missing** (amber), **Skipped** (gray —
  unchecked at lookup time), **Cancelled** (gray).
* Lookup auto-times-out after **30 seconds**; any still-pending card flips
  to Error and the UI lock releases.

### 2. Review

Each card has a checkbox — uncheck any system you want to exclude from
decommission. The status bar summarises results
(`LAPTOP-ABC - 3 found, 1 not found, 1 skipped`). The action hint below the
cards reflects the current selection.

### 3. Confirm

Click **Decommission selected** — a modal opens with:

* **Pre-flight system cards** — one per found system, each showing a system
  pill (AD/Entra/Intune/Autopilot/SCCM), device name, and the relevant
  fields (OS, enabled state, last logon/sign-in/sync, owner, DN/ObjectId/
  serial, etc.). Dates render with relative time:
  `2026-02-03 05:07 (3 days ago)`.
* **Safety warning cards** (yellow) when applicable:
  * BitLocker recovery key(s) escrowed in AD or Entra that will be lost.
  * LAPS password currently stored in AD that will be lost.
  * Recently-active device (last activity within the configurable
    threshold, default 7 days).
* **Type-name prompt** — `Type the device name (HOSTNAME) to confirm:`
  above the input. Case-insensitive.

Toggle **Dry run** first to validate without destructive calls — same modal,
same warnings, but a single **Run dry-run** button (no type-name friction).

After every decommission (real or dry-run) an entry is appended to
`decommission-history.json`.

---

## Pattern matching & wildcards

`*` and `?` wildcards are accepted in the device name input.

| Pattern       | Example      | What happens                                                        |
| ------------- | ------------ | ------------------------------------------------------------------- |
| Exact         | `LAPTOP-ABC` | `-eq` filters everywhere                                            |
| Trailing `*`  | `LAPTOP-*`   | AD: `-like`; Entra/Intune: Graph `startswith()` (server-side, fast) |
| Other `*`/`?` | `*-ABC?`     | AD: `-like`; Entra/Intune: client-side filter over first 200 results |
| GUID          | `00000000-…` | Exact Entra `deviceId eq` lookup                                    |
| SCCM          | `LAPTOP-*`   | `Get-CMDevice -Name` natively accepts `*`/`?`                       |

Trailing-wildcard searches are the fastest. Mid-string or leading wildcards
fall back to a bounded client-side filter capped at 200 results.

---

## Cancel a running scan

While a lookup is in flight the **Look up** button is replaced by a red
**Cancel**. Clicking it:

1. Bumps the internal generation counter so any in-flight runspace result is
   silently dropped on arrival.
2. Calls `$ps.Stop()` on tracked runspaces (best-effort graceful shutdown).
3. Marks any still-searching cards as **Cancelled**.
4. Hides the progress bar and toasts "Lookup cancelled".

You can start a new lookup immediately.

---

## Dry-run mode

Toggle the **Dry run** switch in the action row to validate end-to-end
without calling any destructive cmdlet.

When enabled:

* Decommission button label becomes **Dry-run selected** (validation icon).
* Confirm modal is single-button **Run dry-run** (no type-name).
* Each per-system step still:
  * imports its required module,
  * verifies the cached Graph context has the required scope (no silent
    re-prompt — sign in via the toolbar before running),
  * **for AD/SCCM**: explicitly re-fetches the target object
    (`Get-ADObject` / `Get-CMDevice`) to confirm it still exists,
  * **for Graph systems**: validates the cached lookup result still has a
    usable identifier.
* On success the card flips to **Would remove** (light-accent dot) instead
  of **Removed**, with the message `Would call Remove-… -…`.
* All errors (auth / scope / RBAC / module / cred / network) surface
  exactly as they would in a real run.

Use dry-run after credential changes, after RBAC changes, or whenever you
want to verify the tool is wired up correctly before deleting anything.

---

## Inventory check

A separate **read-only** view (checklist icon in the rail) for fleet
inventory. Paste up to **100 hostnames** (one per line), click **Run
inventory check**, and get a matrix showing which directories each device
exists in. No deletions.

```
┌──────────────┬───────┬───────┬────────┬───────────┬──────┐
│ Device       │ AD    │ Entra │ Intune │ Autopilot │ SCCM │
├──────────────┼───────┼───────┼────────┼───────────┼──────┤
│ LAPTOP-ABC   │ Found │ Found │ Found  │ Not found │ -    │
│ DESKTOP-X1   │ Error │ Found │ -      │ -         │ -    │
│ TABLET-Z5    │ -     │ -     │ Found  │ Found     │ -    │
└──────────────┴───────┴───────┴────────┴───────────┴──────┘
```

| Cell value | Meaning                                                  |
| ---------- | -------------------------------------------------------- |
| `Found`    | Device exists in this system                             |
| `Not found`| Lookup completed; no match                               |
| `Error`    | Module missing or query failed                           |
| `Sign in`  | Graph session not established (sign in to Entra first)   |
| `N/A`      | Required PowerShell module not installed on this machine |
| `-`        | Skipped or no result                                     |

**Copy results** copies the full table as tab-separated text — paste into
Excel for further analysis.

---

## Decommission history

A persistent audit trail of every decommission and dry-run, viewable in-app
(clock icon in the rail). The grid shows:

| Column                      | Notes                                                          |
| --------------------------- | -------------------------------------------------------------- |
| When                        | Timestamp (yyyy-MM-dd HH:mm:ss, local time)                    |
| Device                      | Hostname or ObjectId targeted                                  |
| Operator                    | `DOMAIN\username` of who ran it                                |
| Mode                        | `Real` (accent pill) or `Dry-run` (muted pill)                 |
| AD/Entra/Intune/Autopilot/SCCM | ✓ success (green) / ✗ failure (red) / `-` not targeted     |

Type in the **filter** box to narrow by device, operator, or date. Toggle
**Hide dry-runs** to focus on real removals. Click a row +
**Look up selected** (or double-click) to re-run discovery for that device.

The file `decommission-history.json` is the source of truth. It's
append-only and never auto-rotated. Export to CSV from
**Settings → General → Export audit trail as CSV** for management reporting.

---

## Achievements

The trophy icon in the rail opens a view with **30 unlockable badges**
tracking usage milestones. Cosmetic — there to make the tool a bit more fun.
Categories:

* **First-time** — first lookup, first dry-run, first real decommission,
  first sign-in, first inventory check.
* **Volume** — 5 / 10 / 25 / 50 / 100 real decommissions; 10 dry-runs.
* **Coverage** — AD specialist (10 AD removals), Cloud Native (10 Entra),
  Intune Tamer, Autopilot Ace, SCCM Cleaner, Full Spectrum (all 5 systems
  in one run).
* **Time-based** — Night Owl, Early Bird, Weekend Warrior, Speed Demon
  (under 10 s).
* **Safety** — Heeded the Warning (cancelled with active warnings), LAPS-
  Aware, BitLocker-Aware, Pen & Paper (10 correctly typed names in a row),
  Dry-run Devotee.
* **Tooling** — Chameleon (theme toggle), Wildcard, Quick Reflexes
  (cancelled a lookup), Reporter (CSV export), **Completionist**
  (everything else unlocked).

Unlocks trigger a 5-second toast and a 60-particle confetti burst across
the window. State persists in `achievements.json` (gitignored).

---

## Re-check after decommission

After a successful **real** decommission, the tool waits 3 seconds then
automatically re-runs the lookup against the same device. Cards that were
`Removed` should flip to `Not found`, confirming the directories actually
dropped the object. Acts as a safety net for silent failures or eventual-
consistency delays in Graph / AD replication.

---

## Settings

Open via the gear icon in the title bar or the rail. Settings opens as an
**in-app tabbed view** (replaces the main content; press **Back** or **Esc**
to return).

| Tab        | Fields                                                                                                       |
| ---------- | ------------------------------------------------------------------------------------------------------------ |
| General    | Read-only paths (settings, recent devices, log, audit trail). Recently-active threshold. **Export audit CSV** button. |
| AD         | Server, Search base, Enabled-by-default checkbox.                                                             |
| Entra      | Tenant ID, Enabled-by-default checkbox.                                                                       |
| Intune     | Intune Enabled checkbox + **Autopilot Enabled** checkbox (shares Graph context).                              |
| SCCM       | Site server, Site code (validated `^[A-Z0-9]{1,3}$`), Enabled checkbox.                                       |
| Appearance | Theme (Light / Dark radio buttons).                                                                           |

Every field has inline helper text. Settings persist to `user_settings.json`.

---

## Stored credentials (DPAPI)

AD and SCCM operations can run as either:

* **The interactive Windows user** (leave credentials unconfigured), or
* **A dedicated service account** (configured per system via the **Set AD
  creds** / **Set SCCM creds** buttons in the credentials card).

The credentials card shows a green dot + username when configured, amber +
"Not configured" when not.

To set / replace a credential:

1. Click **Set AD creds** (or **Set SCCM creds**).
2. Enter the username (e.g. `CONTOSO\svc_decommission`) → **Next**.
3. Enter the password → **Save**.

Stored encrypted to `user_creds.dat` next to the script using **DPAPI per-
user** (`ConvertFrom-SecureString` with no key parameter).

> **Only the same Windows user, on the same machine, can decrypt the file.**
> Copying `user_creds.dat` elsewhere makes it unreadable. By design.

There is no shared / exportable credential mode. If you need that, use a
proper secret store (Key Vault, CredMan) and adapt the script.

---

## Files written next to the script

| File                                  | Purpose                                                         | Gitignored |
| ------------------------------------- | --------------------------------------------------------------- | ---------- |
| `user_settings.json`                  | Plain-text settings                                             | ✓          |
| `recent_devices.json`                 | Last 20 looked-up hostnames (sidebar shortcuts)                 | ✓          |
| `user_creds.dat`                      | DPAPI-encrypted AD + SCCM credentials                           | ✓          |
| `decommission-history.json`           | Append-only audit trail                                         | ✓          |
| `achievements.json`                   | Unlocked achievements with timestamps                           | ✓          |
| `%TEMP%\DeviceDecommissioner_debug.log` | Rolling log file (rotates at 2 MB, keeps `.prev`)             | n/a        |

Files live next to the script so the tool is fully portable per user — copy
the folder elsewhere and the tool moves with you (credentials won't decrypt
on the new machine, but everything else works).

---

## Output panel

A bottom log panel shows everything the tool does. Verbose **DEBUG**-level
logging is on by default — there is no toggle to disable it (alpha tool,
visibility wins).

* **Collapse** (down-chevron in the OUTPUT header) hides the panel; a
  compact **OUTPUT** button appears in the status bar to restore it.
  Visibility persists in settings; defaults to **collapsed** on first run.
* **Clear log** empties the panel and ring buffer (500 lines).
* **Copy log** copies the on-screen log to clipboard.

Color coding (theme-aware):

| Level   | Dark mode    | Light mode   |
| ------- | ------------ | ------------ |
| ERROR   | bright red   | dark red     |
| WARN    | orange       | dark orange  |
| SUCCESS | bright green | dark green   |
| INFO    | gray         | dark gray    |
| DEBUG   | dim gray     | mid gray     |

---

## Keyboard shortcuts

| Key       | Context           | Action                       |
| --------- | ----------------- | ---------------------------- |
| `Enter`   | Device input box  | Run lookup                   |
| `Enter`   | Modal text input  | Confirm                      |
| `Esc`     | Settings view     | Return to main view          |
| `Esc`     | Modal dialog      | Close / cancel               |
| `Tab`     | Anywhere          | Standard WPF focus traversal |

---

## What the destructive cmdlets actually do

| Target    | Cmdlet                                                    | Notes                                                                                                                        |
| --------- | --------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------- |
| AD        | `Remove-ADObject -Recursive`                              | Recursive — also clears child registration leaves like `msDS-DeviceRegistration`. "Object not found" is treated as success. |
| Entra     | `Remove-MgDevice`                                         | Removes the directory device object. Hybrid-joined devices may be re-synced from on-prem if you didn't also remove AD.       |
| Intune    | `Remove-MgDeviceManagementManagedDevice`                  | Deletes the managed-device record only — does NOT wipe or retire. Use the Intune portal for wipe.                           |
| Autopilot | `Remove-MgDeviceManagementWindowsAutopilotDeviceIdentity` | Removes the Autopilot enrollment identity (hardware hash, serial, group tag). Device won't auto-enroll on next reset.        |
| SCCM      | `Remove-CMDevice -Force`                                  | Drops the device record (and policy assignments) from the site. The client itself is not uninstalled.                        |

---

## Security notes

* **DPAPI per-user is binding.** `user_creds.dat` only opens for the same
  Windows user, on the same machine that wrote it. By design.
* The tool **never** logs passwords. Only usernames, hostnames, IDs, and
  operation results.
* Decommission is destructive and irreversible — confirmation requires
  typing the device name. Esc cancels safely.
* AD removal is **recursive** — read the table above and check the search
  base in settings if you don't want a wide search radius.
* The AD `-Filter` clause escapes embedded single quotes so a hostname with
  `'` in it cannot break out of the predicate.
* SCCM site code is regex-validated to `^[A-Z0-9]{1,3}$` on save.
* Graph operations require an existing signed-in `MgContext` — the tool
  will not interactively prompt from background runspaces. If a required
  scope is missing, it errors with **"Sign out and back in to acquire it"**
  rather than silently re-prompting.

---

## Troubleshooting

**"Module missing" on a card.** Install the required module (see
[Prerequisites](#prerequisites)). Re-run the lookup — no app restart needed.

**Lookup hangs, then fails after 30 s.** That's the safety timeout. The log
panel (verbose by default) lists the system that never returned. Usually
network / VPN / DNS or a hung Microsoft.Graph cmdlet.

**Graph sign-in pops a browser every lookup.** You're not signed in or the
cached token expired. Click **Sign in to Entra** on the toolbar — it runs
on the UI thread (browser/WAM dialog needs that) but the dispatcher pumps
the wait cursor + status text first so the UI shows feedback.

**Insufficient-scope error during decommission for a Graph system.** Your
cached Graph context lacks the scope. Sign out and back in via the toolbar
button — the tool requests all required scopes (Device.ReadWrite.All,
DeviceManagementManagedDevices.ReadWrite.All,
DeviceManagementServiceConfig.ReadWrite.All) on every fresh sign-in.

**SCCM card always errors.** Open Settings → SCCM and confirm the **Site
server** FQDN and **Site code** are filled in. The Configuration Manager
console must be installed on the machine (the script picks up
`$env:SMS_ADMIN_UI_PATH` first, then falls back to
`${env:ProgramFiles(x86)}\Microsoft Endpoint Manager\AdminConsole\bin\ConfigurationManager.psd1`).

**Click Clear during a lookup, then start a new one — does it race?** No.
A generation counter is bumped; any background callback from the previous
lookup is silently dropped. Verbose log shows `stale callback dropped`.

**The window won't close.** If a lookup or decommission is in flight you'll
get a "still in progress, close anyway?" prompt. Choose Yes to force.

---

## Architecture

* **WPF + WindowChrome** custom title bar. Dark/light themes are hashtables;
  `ApplyTheme` allocates fresh `[SolidColorBrush]::new()` per resource key
  (in-place mutation crashes on frozen brushes). Theme-aware resources
  (`DotColorBrush`, `GlowColorBrush`) drive the dotted-grid background and
  glow tint.
* **PowerShell 5.1 / 7 STA**. Background work uses `RunspaceFactory` plus a
  50 ms `DispatcherTimer` polling `IAsyncResult.IsCompleted` and marshalling
  results back to the UI thread. No `Start-Job` (too slow).
* **Three runspaces per lookup**: AD (its own), Entra+Intune+Autopilot
  (shared — single Graph context), SCCM (its own). Concurrent jobs tracked
  in `$Global:BgJobs`.
* **`Connect-MgGraph` runs on the UI thread** (it needs to show a
  browser/WAM dialog). The wait cursor + "Signing in to Entra..." status
  render before the blocking call via a dispatcher pump.
* **Background runspaces never call `Connect-MgGraph`** — the cached
  `MgContext` from the UI-thread sign-in is process-scoped so background
  threads can use it via `Get-MgContext`. If the context is missing, the
  runspace errors out cleanly with "Sign in via the toolbar first" rather
  than hanging on an interactive prompt that has no parent window.
* **Cancellation** via `$Global:LookupGen` counter. Every background
  callback receives `-Gen $myGen` (closure-captured); stale callbacks from
  cleared/replaced/cancelled lookups are silently dropped. The Cancel
  button additionally calls `$ps.Stop()` on tracked runspaces.
* **Wildcard support**: `*`/`?` switches AD from `-eq` to `-like`,
  Entra/Intune from `eq` to `startswith()` (trailing `*` only — server-side)
  or a bounded client-side filter. SCCM natively accepts `*`.
* **Modal dialogs** are a single shared overlay (`pnlModalOverlay`) reused
  for both confirm and input flows. Optional structured panels render
  device cards, warning callouts, and copyable command snippets.
* **Audit cache**: `$Global:AuditHistoryCache` avoids re-parsing
  `decommission-history.json` on Save → Achievements → History view.
  Invalidated on `Save-AuditEntry` and on explicit Refresh.
* **Card-element cache**: `$Script:CardElements` built once at startup so
  `Set-CardStatus` (called several times per lookup) skips per-call
  `Get-Variable` lookups.
* **Per-monitor DPI awareness** via
  `SetProcessDpiAwarenessContext(-4)` at startup (Win 10 1703+, graceful
  fallback on older OS).
* **Tests**: 40 Pester 5 specs covering AD filter quote-escape, SCCM
  site-code regex, generation counter staleness, Get-DefaultSettings
  defaults, safety warnings (BitLocker / LAPS / recently-active),
  audit-entry structure, Autopilot edge cases, history symbol conversion,
  AllSystems constant, achievement defs (count / keys / unique IDs).
* **PSScriptAnalyzer**: 0 errors, 0 warnings with the bundled
  `PSScriptAnalyzerSettings.psd1` (suppresses intentional patterns —
  `PSAvoidGlobalVars` for WPF DispatcherTimer scope,
  `PSAvoidUsingEmptyCatchBlock` for fire-and-forget cleanup,
  `PSUseApprovedVerbs` for internal GUI helpers).

The PS1 file is monolithic on purpose — single-file portability beats
modularity at this scale.

---

## Limitations & non-goals

* **One device per run for destructive operations.** Bulk delete is not
  supported by design — use the read-only [Inventory check](#inventory-check)
  for fleet checks, then run individual decommissions.
* **No Intune wipe/retire.** Delete of the managed-device record only.
* **No co-management awareness.** If a device is co-managed and policy
  authority lives in Intune for some workloads, removing one side may not
  fully decommission the device for those workloads.
* **One tenant per launch — no in-app tenant picker.** The Entra Tenant ID
  in Settings determines which tenant the Sign in to Entra button connects
  to (blank = home tenant of your account). To work across multiple
  tenants, copy the entire `DeviceDecommissioner` folder per tenant and
  launch each one independently. Each copy keeps its own settings, creds,
  audit, recents, and achievements files — naturally isolated.
* **No undo.** Removed = removed.
