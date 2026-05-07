# Device Decommissioner

A WPF-based PowerShell GUI tool that removes a single device from **Active
Directory**, **Entra ID**, **Microsoft Intune**, **Windows Autopilot**, and
**Configuration Manager (SCCM)** in one guided workflow.

The tool's primary path is intentionally narrow: one device per run, look-up
first, then explicit per-system confirmation, then removal. A separate
read-only **Inventory check** view is available for fleet inventory checks.

> Visual design language mirrors the broader toolkit (WinGet Manifest Manager,
> AIB Pipeline Creator, etc.) — custom WindowChrome title bar with
> EXPERIMENTAL badge, dark/light themes with dotted-grid background, icon rail
> + collapsible sidebar, gradient accent buttons, color-coded log panel,
> toasts, modal confirmations.

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
13. [Recent devices sidebar](#recent-devices-sidebar)
14. [Files written next to the script](#files-written-next-to-the-script)
15. [Output panel](#output-panel)
16. [UI map](#ui-map)
17. [Keyboard shortcuts](#keyboard-shortcuts)
18. [What the destructive cmdlets actually do](#what-the-destructive-cmdlets-actually-do)
19. [Security notes — read first](#security-notes--read-first)
20. [Troubleshooting](#troubleshooting)
21. [Architecture / how it works](#architecture--how-it-works)
22. [Limitations & non-goals](#limitations--non-goals)

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

The window opens centered. On first launch every credential indicator is amber
("Not configured") and the **Decommission** button is disabled until a lookup
completes.

---

## Prerequisites

| System    | PowerShell module                                                                | Auth model                                                       |
| --------- | -------------------------------------------------------------------------------- | ---------------------------------------------------------------- |
| AD        | `ActiveDirectory` (RSAT)                                                         | Stored DPAPI credentials *or* the current Windows user           |
| Entra ID  | `Microsoft.Graph.Identity.DirectoryManagement`                                   | Interactive `Connect-MgGraph` (WAM / browser sign-in)            |
| Intune    | `Microsoft.Graph.DeviceManagement`                                               | Interactive `Connect-MgGraph` (WAM / browser sign-in)            |
| Autopilot | `Microsoft.Graph.DeviceManagement.Enrollment`                                    | Same Graph context as Intune; needs `DeviceManagementServiceConfig.ReadWrite.All` |
| SCCM      | `ConfigurationManager.psd1` (Endpoint Manager Console / SMS_ADMIN_UI_PATH)       | Stored DPAPI credentials *or* the current Windows user           |

PowerShell 5.1+ on Windows. The script must run **STA** for WPF — the launcher
handles this; if you launch by hand, include `-STA`.

Install the Graph modules once:

```powershell
Install-Module `
  Microsoft.Graph.Identity.DirectoryManagement, `
  Microsoft.Graph.DeviceManagement, `
  Microsoft.Graph.DeviceManagement.Enrollment `
  -Scope CurrentUser
```

### Prerequisite banner

On launch the tool automatically checks for missing modules and shows a
**yellow prerequisite banner** at the top of the content area listing what's
absent. The banner has two actions:

* **Install Microsoft.Graph** — runs `Install-Module … -Scope CurrentUser`
  in a background runspace (no UI freeze, no admin needed) and re-checks
  on completion. Visible only when at least one Graph module is missing.
* **How do I fix the rest?** — opens a modal with **copy-pasteable
  PowerShell commands** for every prerequisite (Graph SDK, RSAT via
  `Add-WindowsCapability`, SCCM ConsoleSetup path). Each snippet sits in
  a monospace box with a copy button. Useful when the tool is running
  without admin rights or in environments where PSGallery is blocked.

If a module is missing, the corresponding discovery card shows **Module
missing** at lookup time and that system is silently excluded from
decommission. The other systems still work.

---

## Workflow

```
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────────┐
│  1. Look up     │ -> │  2. Review 4    │ -> │  3. Confirm by      │
│  device by      │    │  discovery      │    │  typing device name │
│  hostname / id  │    │  cards          │    │  (or Run dry-run)   │
└─────────────────┘    └─────────────────┘    └─────────────────────┘
```

### 1. Look up

* Type a device hostname (sAMAccountName / Intune deviceName / Entra
  displayName / SCCM ResourceName), an Entra device ObjectId GUID **or** a
  wildcard pattern (see [Pattern matching](#pattern-matching--wildcards)),
  then press **Enter** or click **Look up**.
* Four background runspaces (one per system) start in parallel. The status
  pulse turns amber, the global progress bar appears in its own row (never
  overlapping content), and each card transitions Idle → **Searching…** →
  final state.
* A **Cancel** button replaces Look up while a scan is running — click it to
  abort immediately (see [Cancel a running scan](#cancel-a-running-scan)).
* Results card states:
  * **Found** (green) — match returned, decommission target armed.
  * **Not found** (gray) — no match.
  * **Error** (red) — module/connectivity error; details visible on the card.
  * **Module missing** (amber) — required PowerShell module isn't installed.
  * **Skipped** (gray) — the system's checkbox was unchecked at lookup time.
* Lookup auto-times-out after **30 seconds**; any still-pending card flips to
  Error and the UI lock releases.

### 2. Review

* Each discovery card has a **checkbox** — uncheck any system you want to
  exclude from decommission.
* The status bar summarises results (e.g. `LAPTOP-ABC123 - 2 found, 1 not
  found, 1 skipped`).
* The action hint below the cards reflects current selection
  (`Ready to remove 'X' from selected systems` / `No matching device targets
  selected.`).

### 3. Confirm

* Click **Decommission selected** — a modal opens with:
  * **Pre-flight system cards** — one card per found system, each showing
    a system pill (AD/Entra/Intune/Autopilot/SCCM), device name, and the
    relevant fields (OS, enabled status, last logon/sign-in/sync, owner,
    DN/ID/serial, etc.). Dates render with relative time, e.g. `2026-02-03
    05:07 (3 days ago)`.
  * **Safety warning cards** (if any), shown as yellow alert boxes:
    * BitLocker recovery key(s) escrowed in AD or Entra that will be lost.
    * LAPS password currently stored in AD that will be lost.
    * Recently-active device — last logon/sign-in/sync within the
      configurable threshold (default 7 days).
  * **Type-name prompt** — `Type the device name (HOSTNAME) to confirm:`
    above the input box. Confirmation is case-insensitive.
* Or toggle **Dry run** first to validate without destructive calls. The
  dry-run modal reuses the same system cards + warnings but skips the
  type-name friction (single **Run dry-run** button).
* After every decommission (including dry-runs) an entry is appended to
  `decommission-history.json` — timestamp, operator, machine, device name,
  per-system result, and dry-run flag (see [Decommission history](#decommission-history)).

---

## Pattern matching & wildcards

You can use `*` and `?` wildcards in the device name input.

| Pattern       | Example input  | What happens                                                               |
| ------------- | -------------- | -------------------------------------------------------------------------- |
| Exact match   | `LAPTOP-ABC`   | `Name -eq` / `displayName eq` / `deviceName eq` / `Get-CMDevice -Name`    |
| Trailing `*`  | `LAPTOP-*`     | AD: `-like`; Entra/Intune: Graph `startswith()` filter (server-side)       |
| Other `*`/`?` | `*-ABC?`       | AD: `-like`; Entra/Intune: client-side filter over first 200 results       |
| GUID          | `00000000-…`   | Exact Entra `deviceId eq` look-up                                          |
| SCCM          | `LAPTOP-*`     | `Get-CMDevice -Name` natively accepts `*` and `?`                          |

Trailing-wildcard searches (`PREFIX-*`) are the fastest since they use
server-side `startswith` on Graph. Mid-string or leading wildcards fall back
to a bounded client-side filter and are capped at 200 results.

---

## Cancel a running scan

While a lookup is in flight the **Look up** button is replaced by a red
**Cancel** button. Clicking it:

1. Bumps the internal generation counter so any in-flight background runspace
   result is silently dropped.
2. Sends `$ps.Stop()` to tracked runspaces (best-effort graceful shutdown).
3. Marks any still-searching cards as **Cancelled**.
4. Hides the progress bar and toasts "Lookup cancelled".

You can start a new lookup immediately after cancelling.

---

## Dry-run mode

Toggle the **Dry run** switch (pill-style toggle) in the action row to
validate end-to-end without calling any destructive cmdlet.

When enabled:

* The Decommission button label changes to **"Dry-run selected"** with a
  validation icon.
* Confirmation modal is a single-button **"Run dry-run"** prompt — no
  type-the-device-name friction since nothing is destructive.
* Each per-system step still:
  * imports its required module,
  * connects to / validates the Graph context (and re-requests scopes if
    they're missing — this is also when interactive sign-in happens for the
    first time),
  * verifies the lookup result still has a usable identifier,
  * **for AD/SCCM** explicitly re-fetches the target object (Get-ADObject /
    Get-CMDevice) to confirm it still exists.
* On success the card flips to **Would remove** (light-accent dot) instead of
  **Removed**, and the message reads `Would call Remove-… -…`.
* All errors (auth / scope / RBAC / connectivity / module / cred / network)
  surface exactly as they would in a real run.

Use dry-run after credential changes, after RBAC changes, or any time you
want to verify the tool is wired up correctly before actually deleting
anything.

---

## Inventory check

A separate **read-only** view (checklist icon in the left rail) for fleet
inventory checks. Paste up to **100 hostnames** (one per line), click **Run
inventory check**, and get a matrix showing which directories each device
exists in. No deletions — this is purely a discovery tool.

```
┌──────────────┬───────┬───────┬────────┬───────────┬──────┐
│ Device       │ AD    │ Entra │ Intune │ Autopilot │ SCCM │
├──────────────┼───────┼───────┼────────┼───────────┼──────┤
│ LAPTOP-ABC   │ Found │ Found │ Found  │ Not found │ -    │
│ DESKTOP-X1   │ Error │ Found │ -      │ -         │ -    │
│ TABLET-Z5    │ -     │ -     │ Found  │ Found     │ -    │
└──────────────┴───────┴───────┴────────┴───────────┴──────┘
```

Cell values:

| Value      | Meaning                                                   |
| ---------- | --------------------------------------------------------- |
| `Found`    | Device exists in this system                              |
| `Not found`| Lookup completed; no match                                |
| `Error`    | Module missing or query failed                            |
| `Sign in`  | Graph session not established (sign in to Entra first)    |
| `N/A`      | Required PowerShell module not installed on this machine  |
| `-`        | Skipped or no result                                      |

Click **Copy results** to copy the full table as tab-separated text — paste
into Excel for further analysis.

---

## Decommission history

A persistent audit trail of every decommission and dry-run, viewable in-app
(clock icon in the rail). The grid shows:

| Column    | Notes                                                              |
| --------- | ------------------------------------------------------------------ |
| When      | Timestamp (yyyy-MM-dd HH:mm:ss, local time)                        |
| Device    | Hostname or ObjectId that was the target                           |
| Operator  | `DOMAIN\username` of who ran it                                    |
| Mode      | `Real` or `Dry-run`                                                |
| AD/Entra/Intune/Autopilot/SCCM | ✓ success, ✗ failure, `-` not targeted        |

**Filter** by typing in the search box (matches device, operator, or
timestamp). Toggle **Hide dry-runs** to focus on real removals only. Click a
row + **Look up selected** (or double-click) to re-run the discovery for that
device.

The file `decommission-history.json` is the source of truth. It's append-only
and never auto-rotated. Export to CSV from **Settings → General → Export
audit trail as CSV** for management reporting.

---

## Achievements

The trophy icon in the rail opens a view with **30 unlockable badges** that
track usage milestones. They're cosmetic — purely there to make the tool a
little more fun to live with. Categories include:

* **First-time milestones** — first lookup, first dry-run, first real
  decommission, first sign-in, first inventory check.
* **Volume tiers** — 5, 10, 25, 50, 100 real decommissions.
* **Coverage breadth** — AD specialist (10 AD removals), Cloud Native (10
  Entra), Intune Tamer, Autopilot Ace, SCCM Cleaner, Full Spectrum (all 5
  systems hit successfully in one run).
* **Time-based** — Night Owl, Early Bird, Weekend Warrior, Speed Demon
  (under 10s).
* **Safety / hygiene** — Dry-run Devotee, LAPS-Aware, BitLocker-Aware,
  Heeded the Warning (cancelled with active warnings), Pen & Paper (10
  correctly typed names in a row).
* **Tooling** — Chameleon (theme toggle), Wildcard, Quick Reflexes
  (cancelled a lookup), Reporter (CSV export), Completionist (everything
  else unlocked).

Unlocks trigger a 5-second toast, a confetti burst across the window, and a
60% opacity → full opacity flip on the badge in the achievements view.
State persists in `achievements.json` (per-user, gitignored). Locked badges
display as `?` with reduced opacity until unlocked.

---

## Re-check after decommission

After a successful **real** decommission (not dry-runs), the tool waits 3
seconds and automatically re-runs the lookup against the same device. Cards
that were `Removed` should flip to `Not found`, confirming the directories
have actually dropped the object.

This is a safety net in case a removal silently fails or hits eventual-
consistency delays in Graph or AD replication. If a card still shows `Found`
after the re-check, something didn't take effect.

---

## Settings

Click the **gear** icon in the title bar or the **Settings** icon in the left
rail. Settings open as an **in-app tabbed view** (replaces the main content;
press **Back** or **Esc** to return).

### Tabs

| Tab        | Fields & purpose                                                                                          |
| ---------- | --------------------------------------------------------------------------------------------------------- |
| General    | Read-only paths (settings, recent devices, log, audit trail). Recently-active threshold. **Export audit trail as CSV** button. |
| AD         | Server, Search base, Enabled checkbox. Helper text per field.                                              |
| Entra      | Tenant ID, Enabled checkbox.                                                                               |
| Intune     | Intune Enabled checkbox + **Autopilot Enabled** checkbox (shares Graph context with Intune).               |
| SCCM       | Site server, Site code (validated `^[A-Z0-9]{1,3}$`), Enabled checkbox.                                   |
| Appearance | Theme (Light / Dark radio buttons).                                                                        |

Every field has inline helper text explaining its purpose and format.

The "Enabled" flags only affect default checkbox state of the five discovery
cards on launch — you can override per-lookup at any time.

Settings are saved to `user_settings.json` next to the script.

---

## Stored credentials (DPAPI)

AD and SCCM operations can run as either:

* **The interactive Windows user** — leave credentials unconfigured.
* **A dedicated service account** — configured per system via the **Set AD
  creds** / **Set SCCM creds** buttons in the credentials card.

The credentials card shows a green dot + username when configured, amber +
"Not configured" when not.

To set / replace a credential:

1. Click **Set AD creds** (or **Set SCCM creds**).
2. Enter the username (e.g. `CONTOSO\svc_decommission`) → **Next**.
3. Enter the password → **Save**.

Stored encrypted to `user_creds.dat` next to the script using **DPAPI per-user**
(`ConvertFrom-SecureString` with no key parameter).

> **Only the same Windows user, on the same machine, can decrypt the file.**
> Copying `user_creds.dat` elsewhere makes it unreadable. This is by design.

There is no "shared" or "exportable" credential mode. If you need that, use a
proper secret store (Key Vault, CredMan, etc.) and adapt the script.

---

## Recent devices sidebar

The left **icon rail** includes a **History** button (clock icon) that toggles
a 260 px sidebar listing the last 20 looked-up device names. Click a name to
re-run the lookup instantly. The sidebar state (visible/collapsed) is persisted
in settings — both sidebar and output panel default to **collapsed on first
run** for a clean initial experience.

Recent devices are stored in `recent_devices.json` (separate from settings, so
settings stay clean and version-controllable). Legacy entries from
`user_settings.json` are auto-migrated on first launch.

---

## Files written next to the script

| File                                            | Purpose                                                 |
| ----------------------------------------------- | ------------------------------------------------------- |
| `user_settings.json`                            | Plain-text settings — see [Settings](#settings).        |
| `recent_devices.json`                           | JSON array of the last 20 looked-up device hostnames.   |
| `user_creds.dat`                                | DPAPI-encrypted JSON with AD + SCCM credentials.        |
| `decommission-history.json`                     | Append-only audit trail — see [Audit trail](#audit-trail). |
| `achievements.json`                             | Unlocked achievements with timestamps. Cosmetic only.       |
| `%TEMP%\DeviceDecommissioner_debug.log`         | Rolling log file (rotates at 2 MB, keeps `.prev`).      |

These files live in the script directory so the tool is fully portable per
user — copy the folder elsewhere and the tool moves with you (credentials
won't decrypt on the new machine, but everything else works).

---

## Audit trail

Every decommission (including dry-runs) appends a JSON object to
`decommission-history.json`. This file is **never auto-rotated** — it's your
permanent record. Each entry contains:

| Field       | Example                                |
| ----------- | -------------------------------------- |
| `Timestamp` | `2026-05-06T14:23:01.1234567+02:00`   |
| `Operator`  | `CONTOSO\alice`                        |
| `Machine`   | `WS-ALICE01`                           |
| `Device`    | `LAPTOP-ABC123`                        |
| `DryRun`    | `true` / `false`                       |
| `Targets`   | `["AD","Entra","Intune"]`              |
| `Results`   | `{ "AD": { "Success": true, "Message": "…" }, … }` |

The audit file path is shown in Settings → General.

---

## Safety warnings

The confirmation modal surfaces warnings when it detects conditions you should
know about before deleting a device:

### BitLocker recovery keys

If the look-up found **BitLocker recovery key objects** escrowed under the AD
computer object (`msFVE-RecoveryInformation` children) or in Entra ID, the
modal warns that those keys will be permanently lost.

### LAPS passwords

If the AD computer object currently holds a **LAPS password** (`ms-Mcs-AdmPwd`
or `msLAPS-Password`), the modal warns it will be lost.

### Recently-active device

If any system's last-activity timestamp (AD `LastLogonDate`, Entra
`ApproximateLastSignInDateTime`, Intune `LastSyncDateTime`) is within the
configurable threshold the modal warns the device may still be in use. The
threshold is set in Settings → General → **Recently-active threshold (days)**
(default 7).

> All warnings are **soft** — they do not block the operation, just ensure
> the operator sees them before typing the device name to confirm.

---

## Output panel

Every action is logged to three places at once:

1. **In-app output panel** at the bottom of the window (`RichTextBox`, ring
   buffer of 500 lines, color-coded by level).
2. **PowerShell console** the script was launched from (if visible).
3. **Disk file** at `%TEMP%\DeviceDecommissioner_debug.log` — full unfiltered
   stream including DEBUG entries; auto-rotates at 2 MB to `.prev`.

### Collapse / restore

The output panel can be **collapsed** via the down-chevron button in the
OUTPUT header. When collapsed the panel and grid splitter disappear entirely;
a compact **OUTPUT** button appears in the status bar to restore it. The
visibility state is saved in settings and defaults to **collapsed** on first
launch.

### Verbose mode

Verbose DEBUG-level logging is **on by default** — every lookup launch,
runspace completion, scope check, modal show/hide, and per-system result
is written to the panel and to the disk log. There is no toggle to turn
it off; in an alpha tool, more visibility into what's happening is more
useful than a clean panel.

### Header buttons

* **Clear log** — empties the panel and ring buffer.
* **Copy log** — copies the entire on-screen log to clipboard.
* **Collapse** (down-chevron) — hides the output panel.

### Color coding (theme-aware)

| Level   | Dark mode    | Light mode   |
| ------- | ------------ | ------------ |
| ERROR   | bright red   | dark red     |
| WARN    | orange       | dark orange  |
| SUCCESS | bright green | dark green   |
| INFO    | gray         | dark gray    |
| DEBUG   | dim gray     | mid gray     |

---

## UI map

```
┌─ TitleBar ──────────────────────────────────────────────────────────────┐
│ ■ Device Decommissioner  v0.1.0-alpha  EXPERIMENTAL    ❓ ⚙ ☀ ─ □ ✕     │
├─── Auth Toolbar ────────────────────────────────────────────────────────┤
│ ● Not signed in    🏢 Tenant: …                  [Sign in to Entra]    │
├─── Progress (own row — never overlays content) ─────────────────────────┤
│ ═══ Searching for 'LAPTOP-*' across configured directories... ═══      │
├─┬───────┬───────────────────────────────────────────────────────────────┤
│ │ ☰     │ ┌ RECENT DEVICES ─────── 🗑 ┐                                  │
│ │ 🖥     │ │ cpc                       │  ┌───────────────────────────┐  │
│ │ ⚙     │ │ LAPTOP-ABC                │  │ Decommission a device     │  │
│ │ ▦     │ │                           │  │                           │  │
│ │ 🕐    │ │                           │  │ ⚠ Prereq banner           │  │
│ │       │ └───────────────────────────┘  │ 🔐 Entra sign-in card      │  │
│ │       │                                │ [hostname / *wild*]        │  │
│ │       │                                │ [Look up] [Cancel]         │  │
│ │       │                                │                           │  │
│ │       │                                │ AD  Entra  Intune  Auto  │  │
│ │       │                                │             pilot   SCCM │  │
│ │       │                                │ CREDENTIALS                │  │
│ │       │                                │ Action row + Dry run       │  │
│ │  ❓   │                                 └───────────────────────────┘  │
├─┴───────┴───────────────────────────────────── ⇕ splitter ──────────────┤
│ 📄 OUTPUT                                              🗑 ⎘ ▼           │
│  [12:34:56.789] [INFO] AD: found 'LAPTOP-ABC123'                        │
├─────────────────────────────────────────────────────────────────────────┤
│ ● Ready                  [▲ OUTPUT]                       v0.1.0-alpha  │
└─────────────────────────────────────────────────────────────────────────┘
```

**Rail icons** (left, top to bottom):

| Icon | Action                                                          |
| ---- | --------------------------------------------------------------- |
| ☰    | Toggle sidebar (Recent devices)                                 |
| ⌂    | **Home** — dismiss any open view, return to decommission flow   |
| ▦    | Open Inventory check view (read-only multi-device lookup)       |
| 🕐   | Open Decommission History view                                   |
| 🏆   | Open Achievements view                                          |
| ⚙    | Open Settings view                                              |
| ❔   | Help (sits at the bottom of the rail)                            |

---

## Keyboard shortcuts

| Key                | Context                | Action                               |
| ------------------ | ---------------------- | ------------------------------------ |
| `Enter`            | Device input box       | Run lookup                           |
| `Enter`            | Modal text input       | Confirm                              |
| `Esc`              | Settings view          | Return to main view                  |
| `Esc`              | Modal dialog           | Close / cancel                       |
| `Tab` / `Shift+Tab`| Anywhere               | Standard WPF focus traversal         |

---

## What the destructive cmdlets actually do

| Target    | Cmdlet                                                              | Notes                                                                                                                                                          |
| --------- | ------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| AD        | `Remove-ADObject -Recursive`                                        | Recursive on purpose — also clears child registration leaves like `msDS-DeviceRegistration` etc. "Object not found" is treated as success (idempotent).        |
| Entra     | `Remove-MgDevice`                                                   | Removes the directory device object. Hybrid-joined devices may be re-synced from on-prem if you didn't also remove the AD object.                              |
| Intune    | `Remove-MgDeviceManagementManagedDevice`                            | Sends a delete on the managed-device record only — it does NOT wipe or retire the device. (If you want wipe / retire, use the Intune portal or a custom flow.) |
| Autopilot | `Remove-MgDeviceManagementWindowsAutopilotDeviceIdentity`           | Removes the Autopilot enrollment identity so the device won't auto-enroll on next reset. Hardware hash, serial, and group tag are gone.                        |
| SCCM      | `Remove-CMDevice -Force`                                            | Drops the device record (and policy assignments) from the site. The client itself is not uninstalled.                                                          |

---

## Security notes — read first

* **DPAPI per-user is binding.** `user_creds.dat` only opens for the same
  Windows user, on the same machine that wrote it. There is no shared/portable
  mode by design.
* The tool **never** logs passwords. The activity log only ever sees usernames,
  hostnames, IDs, and operation results.
* Decommission is destructive and irreversible — confirmation requires typing
  the device name. Esc/click-outside cancels safely.
* AD removal is **recursive** — read the table above and check the search
  base in settings if you don't want a wide search radius.
* The AD `-Filter` clause escapes embedded single quotes so a hostname with
  `'` in it cannot break out of the predicate.
* SCCM site code is regex-validated to `^[A-Z0-9]{1,3}$` on save.
* Graph scopes are requested **only if** the cached `MgContext` doesn't
  already have them. If you're prompted for new scopes mid-session, that's
  why.

---

## Troubleshooting

**"Module missing" on a card**
Install the module on this machine. See [Prerequisites](#prerequisites). The
tool checks via `Get-Module -ListAvailable` inside each background runspace —
no restart needed once installed.

**SCCM card always errors out**
Open Settings → SCCM and confirm the **Site server** FQDN and **Site code**
are filled in. The Configuration Manager console must be installed on the
machine running the tool (the script picks up `$env:SMS_ADMIN_UI_PATH` first,
then falls back to `${env:ProgramFiles(x86)}\Microsoft Endpoint
Manager\AdminConsole\bin\ConfigurationManager.psd1`).

**Graph sign-in pops a browser every lookup**
You're not signed in to Graph in this PowerShell session, or the cached token
expired. Click the **Sign in to Entra** card in the content area — sign-in
runs in a background runspace so the UI stays responsive. Once
`Connect-MgGraph` succeeds the context is cached for the session.

**Insufficient-scope error during decommission for one of the Graph systems**
Means your Graph context has scopes for the system you queried but not for
the one you're decommissioning. The tool now detects this and re-runs
`Connect-MgGraph` with the missing scope — accept the elevation prompt.

**Lookup hangs, then fails after 30s**
That's the safety timeout. Check the log panel (already verbose by default)
for which system never returned. Usually network / VPN / DNS or the
Microsoft.Graph cmdlet hung waiting for a response.

**Click "Clear" while a lookup is in flight, then start a new one — does it
race?**
No. There is a generation counter; any background callback from the previous
lookup is dropped. You'll see a `stale callback dropped` line in DEBUG mode.

**The window won't close**
If a lookup or decommission is in flight you'll get a "still in progress,
close anyway?" prompt. Choose Yes to force.

---

## Architecture / how it works

* **WPF + WindowChrome** for the custom title bar. Dark/light themes are
  hashtables; `ApplyTheme` allocates fresh `[SolidColorBrush]::new()` per
  resource key (in-place `.Color =` mutation crashes on frozen brushes).
  Theme-aware resources (`DotColorBrush`, `GlowColorBrush`) drive the dotted
  grid background and radial glow tint.
* **Icon rail + collapsible sidebar** (hamburger-toggled, 260 px) with
  RECENT DEVICES list. State persisted in settings alongside output-panel
  visibility.
* **In-app tabbed settings** — the settings view swaps visibility with the
  main `ScrollViewer` (no popup/overlay). Six tabs, hashtable-driven
  `Set-SettingsTab` that mutates `BorderBrush`/`Foreground`/`FontWeight`.
* **PowerShell 5.1 STA**. Background work uses `System.Management.Automation
  .Runspaces.RunspaceFactory` plus a 50 ms `DispatcherTimer` that polls
  `IAsyncResult.IsCompleted` and marshals results back onto the UI thread —
  no `Start-Job` spin-up overhead. Entra sign-in (`Connect-MgGraph`) also
  runs in a background runspace so the UI stays responsive.
* **Cancellation** via a `$Global:LookupGen` counter. Every background
  callback receives `-Gen $myGen` (closure-captured); stale callbacks from a
  cleared/replaced/cancelled lookup are silently ignored. The Cancel button
  additionally calls `$ps.Stop()` on tracked runspaces.
* **Wildcard support** — `*`/`?` in the query switches AD from `-eq` to
  `-like`, Entra/Intune from `eq` to `startswith()` (trailing `*`) or
  bounded client-side filter (other patterns). SCCM natively accepts `*`.
* **Progress bar** lives in its own grid row (height toggled between 0 and
  22 px in PowerShell) so it never overlaps content.
* **Modal dialogs** are a single shared overlay (`pnlModalOverlay`) reused
  for both `Show-ModalConfirm` and `Show-ModalInput`. Esc dismisses;
  callback exceptions are caught + Hide-Modal so the overlay can't get stuck.
* **Toasts** are a top-right `StackPanel` of borders with timer-driven fade
  in/out.
* **Graph scope check** uses `$ctx.Scopes -contains $needScope`; if missing,
  re-runs `Connect-MgGraph -Scopes @($needScope) -NoWelcome` to add it.
* **Prerequisite detection** at startup checks all four module families and
  surfaces a banner with optional one-click Graph install.

The PS1 file is monolithic on purpose — single-file portability beats
modularity at this scale.

---

## Limitations & non-goals

* **One device per run for destructive operations.** Bulk delete is not
  supported by design — use the read-only [Inventory check](#inventory-check)
  view for fleet checks, then run individual decommissions for each.
* **No Intune wipe/retire** — this is a delete of the managed-device record
  only. If you want wipe-then-delete, fork the script and call
  `Invoke-MgWipeDeviceManagementManagedDevice` first.
* **No co-management awareness**. If a device is co-managed and policy
  authority lives in Intune for some workloads, removing one side may not
  fully decommission the device for those workloads.
* **One tenant per launch — no in-app tenant picker.** The Entra Tenant ID
  in Settings determines which tenant the Sign in to Entra button connects
  to (blank = the home tenant of your account). To work across multiple
  tenants, copy the entire `DeviceDecommissioner` folder into a separate
  directory per tenant and launch each one independently. Each copy keeps
  its own `user_settings.json`, `user_creds.dat`, `decommission-history.json`,
  `recent_devices.json`, and `achievements.json` — naturally isolated and
  attributed correctly in the audit trail. Switching tenants inside one
  launch would require `Disconnect-MgGraph` + reconnect (re-prompt) on
  every switch and is not worth the safety risk of mixing tenants in the
  same audit file.
* **No undo.** Removed = removed.
