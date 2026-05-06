# WinGet Manifest Manager

> **PowerShell + WPF desktop app** for creating, editing, publishing, deploying, and managing WinGet package manifests on a private Azure-backed REST source.

![Platform](https://img.shields.io/badge/platform-Windows_10%2F11-blue)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)

---

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [User Interface Guide](#user-interface-guide)
  - [Title Bar & Authentication](#title-bar--authentication)
  - [Global REST Source Selector](#global-rest-source-selector)
  - [Left Navigation Rail](#left-navigation-rail)
  - [Activity Log](#activity-log)
- [Create Tab](#create-tab)
  - [Installer Presets](#installer-presets)
  - [PSADT Integration](#psadt-integration)
  - [Community Import](#community-import)
  - [Publishing](#publishing)
- [Edit Tab](#edit-tab)
- [Upload Tab](#upload-tab)
- [Manage Tab](#manage-tab)
- [Deploy Source Tab (BETA)](#deploy-source-tab-beta)
  - [Deploying a New REST Source](#deploying-a-new-rest-source)
  - [Discovering Existing Deployments](#discovering-existing-deployments)
  - [Removing a REST Source](#removing-a-rest-source)
  - [Deploy Profiles](#deploy-profiles)
- [Backup & Restore Tab (BETA)](#backup--restore-tab-beta)
  - [Backup](#backup)
  - [Restore](#restore)
- [Manifest Format Reference](#manifest-format-reference)
- [Package Configuration Files](#package-configuration-files)
- [Settings & Preferences](#settings--preferences)
- [Theme System](#theme-system)
- [Keyboard Shortcuts](#keyboard-shortcuts)
- [Gamification](#gamification)
- [Architecture](#architecture)
- [Troubleshooting](#troubleshooting)
- [FAQ](#faq)

---

## Overview

WinGet Manifest Manager is a desktop GUI for enterprise teams that operate a **private WinGet REST source** on Azure. It replaces manual YAML editing with a guided workflow spanning the full package lifecycle:

| Capability | Description |
|---|---|
| **Create** | Wizard-driven manifest creation with field validation and installer presets |
| **Edit** | Load, modify, and republish existing YAML manifests |
| **Upload** | Push installer binaries to Azure Blob Storage with auto-computed SHA256 |
| **Manage** | List, search, version-compare, and delete published packages |
| **Deploy Source** | One-click deployment of a new private REST source to Azure |
| **Backup / Restore** | Full backup of REST source manifests and storage blobs with timestamped snapshots |
| **Multi-Source** | Global REST source selector — switch between environments without visiting Settings |
| **PSADT** | First-class PSADT import (compiled .exe or PowerShell wrapper) |
| **Gamification** | Publish streaks, achievements, confetti, and quality meter |
| **Themes** | Dark and light mode with Windows 11 design language |

---

## Prerequisites

| Requirement | Version | Notes |
|---|---|---|
| **Windows** | 10 / 11 | WPF requires Windows desktop |
| **PowerShell** | 5.1+ | Ships with Windows; PowerShell 7+ also supported |
| **Az.Accounts** | 2.7.5+ | Azure authentication |
| **Az.Resources** | 6.0.0+ | Resource group and resource discovery |
| **Az.Storage** | 5.0.0+ | Blob Storage operations |
| **Microsoft.WinGet.RestSource** | 1.10.0+ | Manifest publishing (`Add-WinGetManifest`, `Find-WinGetManifest`, etc.) |

### Installing Prerequisites

```powershell
# Install all required modules in one command
Install-Module -Name Az.Accounts, Az.Resources, Az.Storage, Microsoft.WinGet.RestSource `
    -Scope CurrentUser -Force -AllowClobber
```

If you only plan to **deploy** new REST sources (Deploy Source tab), the deploy script requires PowerShell 7.4+.

---

## Installation

1. Clone or download the repository.
2. Navigate to `avd/scripts/WinGetManifestManager/`.
3. Launch via the batch file:

```cmd
Launch_WinGetManifestManager.bat
```

Or launch directly from PowerShell:

```powershell
powershell.exe -ExecutionPolicy Bypass -File WinGetManifestManager.ps1
```

No installation or admin rights required — the app runs from any directory.

---

## Quick Start

1. **Launch** the app via `Launch_WinGetManifestManager.bat`.
2. **Sign in** — click the **Sign In** button in the title bar. A browser window opens for Entra ID authentication.
3. **Select subscription** — use the subscription dropdown that appears after sign-in.
4. **Select REST source** — the global REST source combo auto-discovers Function Apps. Pick your target.
5. **Create a manifest** — fill in package identity, metadata, and installer details on the **Create** tab.
6. **Publish** — click **Publish to REST Source** (or press `Ctrl+Shift+P`).

---

## User Interface Guide

The application uses a three-panel layout:

```
┌────────────────────────────────────────────────────────────┐
│  Title Bar: Auth Status | Subscription | REST Source | Sign In/Out  │
├────────┬───────────────────────────────────┬───────────────┤
│  Left  │       Center Content Area         │               │
│  Rail  │  (Create/Edit/Upload/Manage/      │               │
│        │   Deploy/Backup tabs)             │               │
│        │                                   │               │
├────────┴───────────────────────────────────┴───────────────┤
│  Bottom: Activity Log                                      │
└────────────────────────────────────────────────────────────┘
```

### Title Bar & Authentication

The title bar contains:

- **Auth indicator** (colored dot) — green when connected, red when disconnected
- **Subscription dropdown** — lists all Azure subscriptions; switching triggers resource rediscovery
- **REST Source dropdown** — global source selector (see below)
- **Sign In / Sign Out** buttons
- **Theme toggle** — switch between dark and light mode

**Authentication flow:**

1. Click **Sign In** → browser-based Entra ID login opens.
2. On success: subscription and REST source dropdowns populate automatically.
3. Storage accounts, Function Apps, and resource groups are discovered in background.
4. All tabs become fully functional.

**Required Azure permissions:**

- **Reader** on the subscription (to list resources)
- **Storage Blob Data Contributor** on storage accounts (for blob upload/download)
- **Contributor** on the resource group (for Deploy tab operations)

### Global REST Source Selector

The REST source dropdown in the title bar is the **single source of truth** for which Function App all tabs communicate with:

- **Auto-discovers** Function Apps after sign-in (matches `func-*` prefix, `winget` substring, or `WinGetSource`/`winget-restsource` tags)
- **Merges** discovered sources with saved sources from Settings
- **Syncs** any selection change back to the Settings textbox, Manage tab label, and Backup/Restore fields
- Supports **editable text** — type a custom Function App name if not discovered

Changing the active source instantly affects Create (publish), Edit (publish), Manage (query/delete), and Backup (backup source).

### Left Navigation Rail

The left sidebar has three collapsible sections:

| Section | Purpose |
|---|---|
| **Configurations** | Saved package configs (JSON). Click to load into the Create tab form. |
| **Storage** | Discovered Azure Storage Accounts. Click to select for uploads. |
| **Settings** | Theme, debug logging, default container, manifest path, REST source name, saved sources list. |

### Activity Log

The bottom panel shows timestamped operations with color-coded log levels:

| Level | Color | When used |
|---|---|---|
| INFO | Gray | Standard operations |
| SUCCESS | Green | Successful publish, upload, backup |
| WARN | Orange | Non-fatal issues (e.g., partial backup) |
| ERROR | Red | Failures |
| DEBUG | Blue | Verbose detail (enable via Settings → Debug Overlay) |

Click **Clear** to reset the log. The log auto-scrolls and uses a ring buffer (max 500 lines).

---

## Create Tab

The Create tab is a form-driven wizard for building WinGet manifests. Fields are organized into collapsible sections:

**Identity** — Package Identifier (`Publisher.Package`), Version, Manifest schema version

**Metadata** — Name, Publisher, License, Short Description (max 256 chars), Description (multi-line), Author, Moniker, Tags, Copyright, Release Notes, and URLs (Package, Publisher, License, Support, Privacy)

**Installer** — Type, Architecture, Scope, local file path (with SHA256 auto-compute), Installer URL, silent/interactive/custom/log/repair switches, MinOSVersion, Expected Return Codes

**Installer Advanced** — Platform, Upgrade Behavior, Elevation Requirement, Package Family Name, Product Code, File Extensions, Commands, Nested Installer support

**Dependencies** — Package Dependencies, Windows Features, Windows Libraries, External Dependencies

**Uninstall** — Uninstall Command, Silent Uninstall switch

### Installer Presets

Click **Apply Preset** to auto-fill common installer switches:

| Preset | Silent Switch | Silent with Progress | Log Switch |
|---|---|---|---|
| **MSI** | `/qn /norestart` | `/qb /norestart` | `/l*v "%TEMP%\install.log"` |
| **Inno Setup** | `/VERYSILENT /NORESTART /SP-` | `/SILENT /NORESTART` | `/LOG="%TEMP%\install.log"` |
| **NSIS** | `/S` | `/S` | — |
| **Burn** | `/quiet /norestart` | `/passive /norestart` | `/log "%TEMP%\install.log"` |
| **MSIX** | (auto-silent) | (auto-silent) | — |
| **Portable (ZIP)** | — | — | — |
| **PSADT (.exe)** | `/S` | `/S` | — |
| **PSADT (.ps1)** | `-DeploymentType Install -DeployMode Silent` | `-DeploymentType Install` | — |

### PSADT Integration

Two approaches for packaging PSADT toolkits:

**Approach 1 — Compiled Deploy-Application.exe** — Click **Import PSADT**, select the `.exe` file. The app extracts metadata, computes SHA256, and auto-fills installer type as `exe` with appropriate silent switches.

**Approach 2 — PowerShell Script Wrapper** — Select the PSADT folder containing `Deploy-Application.ps1`. The app reads the script for version and metadata.

Expected PSADT structure:

```
psadt_toolkit/
├── Deploy-Application.ps1
├── Deploy-Application.exe    ← compiled version
├── AppDeployToolkit/
└── Files/
    └── [your-installer.exe]
```

### Community Import

The Create tab includes a **Discover** button that fetches package metadata from the public winget.run community repository. This lets you clone public manifests as a starting point for private packages.

### Publishing

Two ways to publish:

1. **Publish to REST Source** (button or `Ctrl+Shift+P`) — pushes the manifest directly via `Add-WinGetManifest`.
2. **Save to Disk** (`Ctrl+S`) — writes the three-file manifest split (version, installer, locale YAML) to the local `manifests/` folder for manual review.

Publishing requires sign-in and a valid REST source selection. The app validates all required fields before publishing and shows a confirmation dialog on success.

---

## Edit Tab

Load an existing manifest for modification:

- **Load from Disk** — Open a YAML file from the file system.
- **Load from Config** — Open a saved JSON package config.
- Raw YAML editor with syntax-aware display.
- **Publish** — Push edited manifest to the active REST source.
- **Save** — Write changes back to disk.

---

## Upload Tab

Upload installer binaries to Azure Blob Storage:

1. Click **Browse** to select a file.
2. SHA256 hash is computed automatically in the background.
3. Select a **Storage Account** and **Container** (defaults to `packages`).
4. Set the **Blob Path** (auto-generated from filename).
5. Click **Upload** — the blob URL is auto-copied to clipboard and set as the Installer URL on the Create tab.

The **Copy Hash** button copies the SHA256 to clipboard for pasting into the manifest.

---

## Manage Tab

View and manage all packages published to the active REST source:

- **Refresh** — queries `Find-WinGetManifest` to list all packages with version counts.
- **Search** — filter the package list by ID or name.
- **Version Details** — click a package to see all published versions.
- **New Version** — clone the selected package into the Create tab with a bumped version number.
- **Diff Versions** — side-by-side comparison of two versions of the same package.
- **Remove Version** — delete a specific version (with optional blob + config cleanup).
- **Remove Package** — delete all versions (with typed confirmation guard).

The tab header shows the active REST source name: `REST Source: func-corpwinget`.

---

## Deploy Source Tab (BETA)

Deploy or remove a complete WinGet REST source infrastructure on Azure — without touching the Azure Portal or writing ARM templates.

### Deploying a New REST Source

Fill in the deployment form:

| Field | Default | Description |
|---|---|---|
| Source Name | — | Alphanumeric, 3-24 chars. Used as base for all resource names. |
| Resource Group | `rg-winget-prod-001` | Created automatically if it doesn't exist. |
| Region | `westeurope` | Azure region for all resources. |
| Performance Tier | `Developer` | `Developer` / `Basic` / `Enhanced` / `BasicV2` / `StandardV2` — affects APIM, Cosmos DB and App Service SKUs. V2 SKUs use the modern stv2 platform with faster create / deletion and shorter cold start. |
| Publisher Name | (your display name) | Embedded in REST source metadata. |
| Publisher Email | (your Azure email) | Contact email for the source. |
| Authentication | `None` | `None` or `MicrosoftEntraId` — secures the REST API with Entra ID auth. |
| Cosmos DB Region | (same as primary) | Optional override to deploy Cosmos DB in a different region. |
| Manifest Path | (optional) | Folder of manifests to seed after deployment. |

**Advanced options:**
- Skip API Management (removes APIM post-deployment to save cost)
- Register source with local `winget` client
- Enable Cosmos DB zone redundancy
- Enable Private Endpoints (requires VNet/Subnet configuration)

Click **Deploy** to launch `Deploy-WinGetSource.ps1` in a real-time output window. The script creates:
- Azure Function App (REST API host)
- Cosmos DB account (manifest storage)
- Storage Account (configuration)
- App Service Plan
- Key Vault (secrets)
- App Configuration
- API Management (optional)

### Discovering Existing Deployments

Click **Scan Subscription** to discover existing WinGet REST source deployments. The scanner finds:

- **Function Apps** matching `func-*` prefix, `winget` substring, or `WinGetSource`/`winget-restsource` Azure tags
- **Storage Accounts** matching `winget`/`wgsrc` patterns or tagged accordingly

Double-click a discovered resource to load it into the deploy form. The app auto-pairs related resources (e.g., `func-corpwinget` + `stcorpwinget` in the same resource group).

### Removing a REST Source

Click **Remove** to tear down a deployment. The removal script:

1. Unregisters the local `winget source` (unless skipped)
2. Cleans up Entra ID app registrations
3. Purges soft-deleted resources (Key Vault, APIM, App Service)
4. Deletes the resource group

Options: Skip source removal, skip Entra cleanup, skip soft-delete purge, or purge-only mode.

### Deploy Profiles

Save deployment configurations as named profiles for repeatable deployments. Click **Save Profile** after filling the form, and **Load Profile** to restore settings.

---

## Backup & Restore Tab (BETA)

Full backup and restore of your WinGet REST source environment.

### Backup

Select what to back up:

- **REST Source Manifests** — downloads all package manifests via `Get-WinGetManifest` API
- **Storage Account Blobs** — downloads the `packages` container contents

Both the REST source and Storage Account fields are **dropdown selectors** populated from discovered resources. You can also type a custom name.

Set a **Backup Folder** (defaults to `<app>/backups/`). A timestamped subfolder is created automatically:

```
backups/
└── corpwinget_20260313_143022/
    ├── manifests/
    │   ├── MyOrg.7-Zip/
    │   │   └── manifest.json
    │   └── MyOrg.VSCode/
    │       └── manifest.json
    ├── storage/
    │   └── packages/
    │       └── [blob files]
    └── backup_metadata.json
```

The `backup_metadata.json` records the source names, timestamp, and counts for easy restore.

### Restore

1. **Browse** to a backup folder (the app reads `backup_metadata.json` and pre-fills target fields).
2. Select **Target REST Source** and **Target Storage Account** — defaults to the original source names but can be changed for cross-environment migration.
3. Click **Restore Now** — manifests are republished via `Add-WinGetManifest` and blobs re-uploaded.

A confirmation dialog shows the count of manifests and blobs before proceeding.

---

## Manifest Format Reference

The app generates WinGet-compliant **three-file manifest split**:

```
manifests/
└── M/
    └── MyOrg.AppName/
        └── 1.0.0/
            ├── MyOrg.AppName.yaml              # Version manifest
            ├── MyOrg.AppName.installer.yaml     # Installer manifest
            └── MyOrg.AppName.locale.en-US.yaml  # Default locale manifest
```

### Version Manifest

```yaml
PackageIdentifier: MyOrg.AppName
PackageVersion: 1.0.0
DefaultLocale: en-US
ManifestType: version
ManifestVersion: 1.9.0
```

### Installer Manifest

```yaml
PackageIdentifier: MyOrg.AppName
PackageVersion: 1.0.0
InstallerType: exe
Installers:
  - Architecture: x64
    InstallerUrl: https://stcorpwinget.blob.core.windows.net/packages/AppName-1.0.0.exe
    InstallerSha256: ABC123...
    InstallerSwitches:
      Silent: /S
      SilentWithProgress: /S
Scope: machine
ManifestType: installer
ManifestVersion: 1.9.0
```

### DefaultLocale Manifest

```yaml
PackageIdentifier: MyOrg.AppName
PackageVersion: 1.0.0
PackageLocale: en-US
Publisher: My Organization
PackageName: App Name
License: Proprietary
ShortDescription: A short description of the app
ManifestType: defaultLocale
ManifestVersion: 1.9.0
```

Multi-line descriptions use YAML block scalar syntax (`|`) automatically.

---

## Package Configuration Files

Package configs are JSON files that store all form state for a package. They live in the `package_configs/` folder and enable:

- **Save** — persist the entire Create tab form to disk.
- **Load** — restore all fields from a saved config.
- **Quick access** — browse configs in the left rail under "Configurations".

Example schema:

```json
{
  "PackageIdentifier": "MyOrg.7-Zip",
  "PackageVersion": "25.01",
  "PackageName": "7-Zip",
  "Publisher": "My Organization",
  "License": "LGPL-2.1",
  "ShortDescription": "File archiver with a high compression ratio",
  "InstallerType": "exe",
  "Architecture": "x64",
  "InstallerUrl": "https://stcorpwinget.blob.core.windows.net/packages/7-Zip-25.01.exe",
  "InstallerSha256": "...",
  "SilentSwitch": "/S",
  "Scope": "machine"
}
```

---

## Settings & Preferences

Settings are accessible from the left rail. User preferences are persisted to `user_prefs.json` in the app directory.

| Setting | Default | Description |
|---|---|---|
| Dark / Light Mode | Dark | Theme toggle (also in title bar) |
| Debug Overlay | Off | Show DEBUG-level entries in Activity Log |
| Auto-Save Configs | On | Save package configs automatically on changes |
| Confirm Upload | On | Show confirmation dialog before blob uploads |
| Default Container | `packages` | Default blob container for uploads |
| Manifest Version | `1.9.0` | WinGet manifest schema version |
| REST Source Name | — | Default Function App name (synced with global selector) |
| Saved Sources | — | Named list of REST source endpoints for quick switching |
| Source Discovery Pattern | `^func-` | Regex filter for Function App names during auto-discovery (always also matches `winget`) |
| Default Subscription | — | Auto-select this subscription by name after sign-in (dropdown populated from discovered subs) |

### Build Traceability

The **About** section in Settings shows the app version and the git commit hash for the current build. Build metadata is stored in `build.json`:

```json
{
    "commit": "f30652c",
    "buildDate": "2026-03-13"
}
```

Before distributing the tool to machines without git, run `Update-BuildInfo.ps1` to stamp the current commit hash:

```powershell
.\Update-BuildInfo.ps1
# build.json updated  →  commit abc1234  |  2026-03-13
```

---

## Theme System

The app supports dark and light modes with 22 dynamic brush resources.

The center content area features a subtle **dot grid background** with a blue accent gradient glow — inspired by the Bagel Commander design language. The dot pattern uses a tiled `DrawingBrush` and automatically adapts between light and dark modes.

**Dark Mode (default):**

| Element | Color |
|---|---|
| Window background | `#111113` |
| Card background | `#1E1E1E` |
| Text primary | `#FFFFFF` |
| Accent | `#0078D4` |

**Light Mode:**

| Element | Color |
|---|---|
| Window background | `#F5F5F5` |
| Card background | `#FFFFFF` |
| Text primary | `#111111` |
| Accent | `#0078D4` |

Toggle via the sun/moon button in the title bar or the checkbox in Settings.

---

## Keyboard Shortcuts

| Shortcut | Action |
|---|---|
| `Ctrl+S` | Save manifest to disk |
| `Ctrl+Shift+P` | Publish manifest to REST source |

---

## Gamification

Publishing packages unlocks engagement features:

### Publish Streak

Consecutive days of publishing are tracked. The title bar shows your current streak as a fire icon.

### Achievements

| Achievement | Trigger |
|---|---|
| First Publish | Publish your first manifest |
| Speed Demon | Publish within 60 seconds of launching |
| Streak Master | Reach a 7-day publish streak |
| PSADT Master | Publish a PSADT-packaged app |
| Bulk Publisher | Publish 10+ packages total |
| Version Bumper | Publish a version update of an existing package |
| Multi-Arch | Publish for multiple architectures |
| Night Owl | Publish between midnight and 5 AM |

Achievements are persisted in `achievements.json` and display with toast notifications and confetti animations. The achievements panel in the sidebar is **collapsible** via the chevron toggle — state is persisted across sessions.

### Quality Meter

A visual gauge on the Create tab shows manifest completeness. Filling optional fields (description, tags, release notes, URLs) increases the quality score.

---

## Architecture

```
WinGetManifestManager.ps1     Main application logic (~7,800 lines)
WinGetManifestManager_UI.xaml  WPF layout definition (~2,950 lines)
Deploy-WinGetSource.ps1        REST source deployment script
Remove-WinGetSource.ps1        REST source teardown script
Launch_WinGetManifestManager.bat  Launcher (auto-detects pwsh/powershell)
```

### Threading Model

The WPF window runs on the STA (Single-Threaded Apartment) UI thread. All Azure API calls, file hashing, blob uploads, and module imports run in **background runspaces** via the `Start-BackgroundWork` helper. A `DispatcherTimer` (50ms interval) polls completed jobs and marshals results back to the UI thread.

### File Structure

```
WinGetManifestManager/
├── WinGetManifestManager.ps1        # Main app
├── WinGetManifestManager_UI.xaml    # UI layout
├── Deploy-WinGetSource.ps1          # Deploy script
├── Remove-WinGetSource.ps1          # Remove script
├── Launch_WinGetManifestManager.bat # Launcher
├── README.md                        # This file
├── user_prefs.json                  # Persisted settings
├── publish_streak.json              # Streak tracking
├── achievements.json                # Unlocked achievements
├── manifests/                       # Locally saved YAML manifests
├── package_configs/                 # Saved JSON package definitions
├── psadt_packages/                  # PSADT template packages
├── backups/                         # Local backup snapshots
└── docs/                            # Additional documentation
```

---

## Troubleshooting

### "REST Source Not Set" dialog on publish

Set the REST source Function App name in one of these places:
- **Global dropdown** in the title bar (recommended)
- **Settings** → REST Source Name field in the left rail

### Sign-in opens browser but nothing happens

The app disables Windows Account Manager (WAM) and uses interactive browser auth. Ensure your default browser can reach `login.microsoftonline.com`. If you have a cached expired token, sign out and sign in again.

### Module not found errors

Install missing modules:

```powershell
Install-Module Microsoft.WinGet.RestSource -Scope CurrentUser -Force
Install-Module Az.Accounts, Az.Resources, Az.Storage -Scope CurrentUser -Force
```

### Upload fails with "Forbidden" or "AuthorizationFailure"

Ensure your account has **Storage Blob Data Contributor** role on the target storage account. Key-based access is not used — the app authenticates via Entra ID (`-UseConnectedAccount`).

The app's role check now distinguishes **data-plane** roles (`Storage Blob Data *`) from **control-plane** roles (Owner, Contributor, User Access Administrator). Owner/Contributor alone do **not** grant blob access — only the data-plane variant does. When the GUI detects this mismatch on the Upload tab, it shows a one-click "Assign Role" prompt and then polls every 15 s for RBAC propagation (up to 5 min).

### `winget install` fails with `Conflict (409)` / `Download request status is not success`

The blob exists but anonymous (unauthenticated) GET is blocked. WinGet client downloads installers without auth headers, so two layers must allow public read:

1. **Storage account** — `AllowBlobPublicAccess = $true` (master kill-switch; if `$false`, container ACL is ignored)
2. **Container ACL** — `Blob` (anonymous read on blobs only) or `Container` (also lists)

The Upload tab → Storage Info panel now exposes a one-click **Enable / Disable public read** button at the bottom that handles both layers in the correct order. Use this for test/preview environments only — disable it once VNet integration + Private Endpoints are in place. Manual equivalent:

```powershell
$rg = '<rg>'; $sa = '<storage-account>'
Set-AzStorageAccount -ResourceGroupName $rg -Name $sa -AllowBlobPublicAccess $true
$ctx = (Get-AzStorageAccount -ResourceGroupName $rg -Name $sa).Context
Set-AzStorageContainerAcl -Name 'packages' -Permission Blob -Context $ctx
```

If 409 persists after both layers are enabled, also check `PublicNetworkAccess = 'Enabled'` and the storage firewall (Networking → "Enabled from all networks" or your client IP allowed).

### Role-propagation poll runs at full speed (no 15 s gap between attempts)

Fixed in `d4db8a7`. `Start-Sleep -Seconds $Script:RBAC_POLL_SEC` inside `Poll-RolePropagation`'s background runspace was resolving to `$null` because background runspaces don't inherit `$Script:` scope. Result: the poll spun roughly once per second and threw `BgJob[0]: ERROR: Cannot validate argument on parameter 'Seconds'` repeatedly. The poll seconds value is now passed via the runspace `-Variables` hashtable with a 15 s fallback if missing.

### `Get-Date` throws `Cannot find a positional parameter` repeatedly in DispatcherTimer

Fixed in `306ce9b`. The WPF DispatcherTimer tick (50 ms) was calling `Get-Date` for elapsed time. Under certain WPF + PowerShell-7 thread states the cmdlet resolution intermittently fails on the dispatcher thread. Replaced with `[datetime]::Now` (no cmdlet lookup, pure CLR) — also faster.

### GUI freezes / Sign Out button stops responding after a background job runs

Fixed by commits `be59a99` (DispatcherTimer hardening — wrapped all WPF property writes in try/catch + outer catch block) and `8cb23b3` (cross-thread scriptblock fix — never pass a `[ScriptBlock]` via runspace `-Variables` and invoke with `&`; use a function-source string + `Invoke-Expression` per runspace instead). Symptom was `SessionStateUnauthorizedAccessException: Global scope cannot be removed` — a single uncaught exception in the dispatcher tick handler was killing the message loop.

### Container browser shows "Discarding stale container results for ''"

Fixed in `69cf656`. The container-load OnComplete callback compared the loaded account name against the closure-captured `$AcctInfo`, which had been re-bound by a subsequent `cmbUploadStorage_SelectionChanged` event. Result: legitimate results were discarded as "stale" and the role-warning prompt never appeared. Now the account name is **pinned at queue time** via the `-Context @{ LoadedAcctName = ... }` hashtable and read from the OnComplete's `$Ctx` parameter.

### Guest user (B2B) sign-in succeeds but role assignment shows "user not found"

Fixed in `e96ebcc`. `Get-AzADUser -UserPrincipalName <home-tenant-email>` returns `$null` against a customer's tenant for guest accounts (the UPN in the customer tenant is mangled — `email_home.com#EXT#@customer.onmicrosoft.com`). The new `Resolve-CurrentUserObjectId` function tries 4 strategies in order: `Get-AzADUser -SignedIn`, Microsoft Graph `/me`, `-UserPrincipalName`, then `-Mail`. Defined as a **here-string** (`$Script:ResolveOidFnSrc`) so each background runspace can re-create it via `Invoke-Expression` without cross-thread scriptblock corruption.

### Deploy script fails

The deploy script requires **PowerShell 7.4+** (`pwsh`). If you launched the app with Windows PowerShell 5.1, the deploy script will spawn a separate `pwsh` process. Ensure PowerShell 7 is installed.

### Backup shows 0 manifests

Verify the REST source Function App name is correct and the source has published packages. The backup uses `Find-WinGetManifest` → `Get-WinGetManifest` which requires the `Microsoft.WinGet.RestSource` module.

### Soft-deleted resources block redeployment

Use the **Remove** function with default settings — it automatically purges soft-deleted Key Vaults, APIM instances, and App Services that block name reuse.

### Deploy fails with `Standardv2` (or any V2 SKU) not in ValidateSet

The wrapper accepts the V2 SKU names but the **upstream** `Microsoft.WinGet.RestSource` module on PSGallery still has the legacy ValidateSet (`Developer,Basic,Standard,Premium,Consumption`). If a stale upstream copy on disk wins the version race, the inner `New-WinGetSource` rejects `BasicV2`/`StandardV2` 30+ minutes into the deploy.

The script now (a) enumerates **every** copy of `Microsoft.WinGet.RestSource` on disk and warns when more than one is present, and (b) probes the loaded module's ValidateSet **before** any ARM work via `Test-UpstreamSkuSupport`. If the requested SKU is not supported, the deploy aborts immediately with three resolution paths: install the patched fork at a higher version, hot-patch the installed copy's `Library\New-WinGetSource.ps1`, or pick a supported tier and resize APIM after deploy with `Update-AzApiManagement -Sku StandardV2`.

### Deploy fails with `Cannot bind parameter 'Sku'. Cannot convert value 'BasicV2' to type PsApiManagementSku`

The bundled fork's `New-ARMObjects.ps1` historically called `New-AzApiManagement -Sku $sku` **before** delegating to the ARM template. The Az.ApiManagement cmdlet binds `-Sku` to a strongly-typed `[PsApiManagementSku]` enum that only knows the legacy SKUs (`Developer`/`Basic`/`Standard`/`Premium`/`Consumption`) — `BasicV2`/`StandardV2` are rejected at parameter binding before any REST call is made.

The bundled fork now detects SKU values matching `/v2$/i`, skips the cmdlet path entirely, and lets `New-AzResourceGroupDeployment` create APIM via the ARM template (`sku.name` is a free-form string and accepts both V2 SKUs). The keyvault-access policy that the cmdlet path used to set is moved to a post-ARM-deploy block that fetches the freshly-created APIM identity. Legacy SKUs are unaffected.

### Deploy fails with `Provided list of regions contains duplicate regions. Please remove duplicates from <name>, <name>` (Cosmos DB)

`New-ARMParameterObjects` ships `cosmosdb.json` with two `locations[]` entries (failoverPriority 0 and 1) both defaulting to the primary region. When `-CosmosDBRegion` is supplied the wrapper renames *both* to the same region display name and ARM rejects the deploy.

The wrapper now adds a deduplication pass after the existing region/zone patches: it keeps the first occurrence of each `locationName` (case- and whitespace-insensitive) and renumbers `failoverPriority` sequentially from 0. Single-region deployments end up with a single locations entry; multi-region overrides retain ordering.

### Deploy aborts with `APIM SKU MISMATCH — IN-PLACE UPGRADE NOT SUPPORTED`

The Phase 3b resource audit detected a reused APIM whose SKU does not match the SKU implied by `-PerformanceTier`, and the change is a classic↔V2 jump (e.g. existing `Developer` and requested `BasicV2`). Azure does not support in-place migration between classic and V2 APIM SKUs and would otherwise burn ~30 min in ARM retries before giving up.

Resolution — delete the existing APIM first, then re-run:

```powershell
Remove-AzApiManagement -ResourceGroupName '<rg>' -Name '<apim-name>'
# wait ~10-15 min for delete to complete, then re-run Deploy-WinGetSource.ps1
```

Alternatively re-run with `-PerformanceTier` matching the existing SKU. Same-family upgrades (e.g. `Basic` → `Standard`) still proceed with a WARN.

### Deploy fails with `NameUnavailable` on App Configuration / Key Vault / APIM

All three resource types use Azure soft-delete with multi-day retention windows that reserve the name even after the resource group is deleted. The Phase 3b pre-flight audit auto-purges soft-deleted Key Vaults (recover-into-RG when policy blocks purge), APIMs and App Configuration stores before each deploy. If you saw `NameUnavailable` in earlier ARM-deploy logs, just re-run — the audit will resolve it on the next attempt.

Manual purge if needed:

```powershell
# App Configuration
$tok = (Get-AzAccessToken -ResourceUrl 'https://management.azure.com/').Token
$h   = @{ Authorization = "Bearer $tok" }
$sub = (Get-AzContext).Subscription.Id
Invoke-RestMethod -Uri "https://management.azure.com/subscriptions/$sub/providers/Microsoft.AppConfiguration/locations/<region>/deletedConfigurationStores/<name>/purge?api-version=2023-03-01" -Method POST -Headers $h

# APIM
Invoke-RestMethod -Uri "https://management.azure.com/subscriptions/$sub/providers/Microsoft.ApiManagement/locations/<region>/deletedservices/<name>?api-version=2022-08-01" -Method DELETE -Headers $h

# Key Vault
Remove-AzKeyVault -VaultName '<name>' -Location '<region>' -InRemovedState -Force
```

### Deploy aborts at Phase 1 with `Cannot find path '...Library\WinGet.RestSource.PowershellSupport\Microsoft.Winget.PowershellSupport.dll'`

The upstream fork repo's `Tools/PowershellModule/src/` only contains source `.ps1` files — the compiled C# helper (`Microsoft.WinGet.PowershellSupport.dll` + `runtimes/win-{x64,x86,arm64}/native/WinGetUtil.dll`) is built separately. Earlier bundled copies of the fork shipped without those binaries and crashed at `Add-Type` during module import.

The bundled fork now ships the full `Library\WinGet.RestSource.PowershellSupport\` folder (8 files, ≈12 MB) copied from the PSGallery v1.10.0 install. The startup sanity log includes a `Bundled support : <path> (exists=...)` line so missing native deps are obvious in the first ten lines of any deploy log.

### Function App returns "isolated worker" error after deploy

The .NET-isolated Functions host fails to start when `FUNCTIONS_WORKER_RUNTIME` ≠ `dotnet-isolated` or when `FUNCTIONS_EXTENSION_VERSION` ≠ `~4`. Upstream's zip upload (via `Publish-AzWebApp -ArchivePath`) does not always set these and silently fails on Flex Consumption / isolated-worker plans.

The script now runs `Assert-FunctionAppHealthy` after every deploy: it checks the two app settings, patches them via `Update-AzFunctionAppSetting` if wrong, restarts the host, and pings `/api/information` for up to 60 s. If the host is still unhealthy, it falls back to a OneDeploy re-publish (`Publish-FunctionZipOneDeploy`) which uses Kudu's `/api/publish?type=zip` REST endpoint with bearer auth — works on Consumption, Premium, Flex, and isolated-worker plans alike.

### `Microsoft.WinGet.RestSource` module update lands in the wrong folder

`Install-Module` / `Install-PSResource` install per edition: PowerShell 7 writes to `$HOME\Documents\PowerShell\Modules\`, Windows PowerShell 5.1 writes to `$HOME\Documents\WindowsPowerShell\Modules\`. If both paths contain copies, PowerShell loads the highest version regardless of edition, which is how a stale upstream copy can beat a freshly-installed fork. The deploy script now lists every copy it finds and `Remove-Module -Force`s any in-process instance before importing, so an explicit `Install-PSResource -Reinstall` actually takes effect on the next session.

---

## Changelog

### 2026-05-06 (revision 2) — GUI hardening + Connect / Public-Access UX

| Area | Change | Commit / File(s) |
|------|--------|-----------------|
| Connect command (new GUI button) | Added a **Connect** button next to **Test** in Settings. Resolves the WinGet REST source URL in a background runspace: looks up the Function App, then queries any APIM in the same RG via ARM REST (`api-version=2023-05-01-preview`) for `properties.gatewayUrl`. Shows themed dialog with the full `winget source add ...` command, `Invoke-RestMethod .../information` test snippet, and an APIM subscription-key hint. **Copy command** button → clipboard + toast. | `a7f7bea` · `Show-RestSourceConnectCommand` (new), `WinGetManifestManager_UI.xaml` (btnConnectCmd) |
| Public-access toggle (new GUI control) | Added a one-click **Enable / Disable public read** button at the bottom of the Storage Info panel. Sets `AllowBlobPublicAccess` at the account level **and** the container ACL (`Blob` / `Off`) in the correct order. Confirmation dialog spells out the security tradeoff. Auto-refreshes the storage info panel on completion. Designed for test/preview environments only — to be flipped off once Private Endpoints land. | `c0d34e8` · `Set-StoragePublicAccess` (new), inline action panel in `cmbUploadStorage` OnComplete |
| Container-load stale-check fix | OnComplete now reads the queued account name from `$Ctx.LoadedAcctName` (passed via `Start-BackgroundWork -Context @{...}`) instead of the closure-captured `$AcctInfo`, which could be re-bound by a subsequent `SelectionChanged` event. Without this, legitimate container results were silently discarded as "stale" and the role-prompt path never fired. | `69cf656` · `cmbUploadStorage.Add_SelectionChanged` |
| Cross-thread scriptblock corruption | `Resolve-CurrentUserObjectId` is now stored as a here-string (`$Script:ResolveOidFnSrc`) and recreated inside each background runspace via `Invoke-Expression`. Previously it was passed as a `[ScriptBlock]` via `-Variables` and invoked with `&`, which carries the original session state across threads and ultimately corrupts the main thread with `SessionStateUnauthorizedAccessException: Global scope cannot be removed`. | `8cb23b3` · `Resolve-CurrentUserObjectId`, `Start-BackgroundWork` callers |
| DispatcherTimer hardening | Wrapped all WPF property writes (`prgUpload.Value`, `lblUploadProgress.Text`, etc.) in `try/catch` and added the missing outer `catch` block to the timer tick. A single unhandled exception in the 50 ms tick was killing the dispatcher message loop, leaving the GUI frozen with no visible error. | `be59a99` · `BgJob` DispatcherTimer tick handler |
| `Get-Date` → `[datetime]::Now` | Replaced `Get-Date` with `[datetime]::Now` in the DispatcherTimer tick and BgJob queueing. Cmdlet resolution intermittently fails on the WPF dispatcher thread under certain PS-7 thread states; the CLR static is immune and faster. | `306ce9b` · DispatcherTimer tick, `Start-BackgroundWork` |
| Role-propagation poll Sleep | Fixed `BgJob[0]: ERROR: Cannot validate argument on parameter 'Seconds'` — `Start-Sleep -Seconds $Script:RBAC_POLL_SEC` resolved to `$null` inside the background runspace (no script scope inheritance). Poll seconds value now passed via `-Variables` with a 15 s fallback. | `d4db8a7` · `Start-RolePropagationCheck` |
| Data-plane vs control-plane role check | The Upload-tab role check now only treats `Storage Blob Data *` roles (Reader/Contributor/Owner) as upload-OK. Owner / Contributor / User Access Administrator are control-plane only and do **not** grant blob data access — Azure Storage will still 403 the request. Without this, the GUI silently skipped the role-prompt path on accounts that had only control-plane roles assigned. | `1a6190a` · storage-prereq check inside container-load |
| Guest user (B2B) OID resolution | `Resolve-CurrentUserObjectId` now tries 4 strategies in order: `Get-AzADUser -SignedIn`, Microsoft Graph `/me`, `-UserPrincipalName`, `-Mail`. Fixes guest-user role assignment in customer tenants where the UPN is mangled (`email_home.com#EXT#@customer.onmicrosoft.com`) and `-UserPrincipalName <home-email>` returns `$null`. | `e96ebcc` · `Resolve-CurrentUserObjectId` (new) |
| ARM-deploy retry visibility | Per-attempt banners with `▶`/`✓`/`✗` glyphs and elapsed time. Periodic APIM delete-poll progress every 60 s (was silent for 14+ min). Generic failure summary instead of stack trace. Retry reason printed before next attempt. | `acc3b38` · `Step-DeployRestSource` retry loop |
| Resilient `Write-Log` + redesigned summary table | Replaced `Add-Content` with `[System.IO.File]::AppendAllText` (3-attempt retry, 100 ms backoff). Without this, a transient "Stream was not readable" aborted the entire deploy. Summary table widened to 76 chars with `$fmt` scriptblock pattern + ellipsis truncation so long URLs no longer push the right border out of alignment. | `551ddb5` · `Write-Log`, `Write-Summary` |
| KV access-policy race pre-grant | Pre-grants the APIM identity → Key Vault `secrets get` permission between ARM-deploy retries when the previous attempt failed with `does not have secrets get permission on key vault`. Collapses 2 attempts to 1. | `92a9617` · ARM-deploy retry loop |
| Cleaner ARM deploy output | Suppressed cmdlet noise via `-ErrorAction SilentlyContinue` + `-ErrorVariable`, single-line failure summary via regex extraction (`[Code] Status Message`), friendlier recovery messages. V2 SKU warnings demoted to Verbose in `New-ARMObjects.ps1`. | `b584b09` · `Step-DeployRestSource`, `New-ARMObjects.ps1` |

### 2026-05-06 — Bundled fork hardening (post-customer-test fixes)

| Area | Change | Function(s) / File(s) |
|------|--------|-----------------------|
| Native helper bundled | Added `Library\WinGet.RestSource.PowershellSupport\` (compiled C# helper + win-{x64,x86,arm64} native `WinGetUtil.dll`, 8 files / ≈12 MB) so the bundled fork imports cleanly on machines without a PSGallery install of `Microsoft.WinGet.RestSource`. Without these, the fork's `.psm1` failed at `Add-Type` and aborted Phase 1. | `Modules\Microsoft.WinGet.RestSource\Library\WinGet.RestSource.PowershellSupport\` |
| Sanity log | Added a 3rd existence check (`Bundled support : <path> (exists=...)`) to the startup sanity log so a missing native helper is visible in the first ten lines of any deploy log. | `Deploy-WinGetSource.ps1` (top-level startup banner) |
| V2 SKU cmdlet binding | Bundled fork's `New-ARMObjects.ps1` now detects SKU values matching `/v2$/i`, skips the `New-AzApiManagement` cmdlet path (whose `[PsApiManagementSku]` enum doesn't accept V2 SKUs), and lets `New-AzResourceGroupDeployment` create APIM via the ARM template. A post-deploy block sets the keyvault access policy using the freshly-created APIM identity. Legacy SKUs are unaffected. | `Modules\Microsoft.WinGet.RestSource\Library\New-ARMObjects.ps1` |
| V2 KV access policy (pre-deploy) | When `Get-AzApiManagement` returns `$null` for a V2 APIM (cmdlet enum can't deserialize V2 SKU), the bundled fork now falls back to `Invoke-AzRestMethod GET .../service/<name>?api-version=2023-05-01-preview`, reads `identity.principalId` from the JSON, and applies `Set-AzKeyVaultAccessPolicy` **before** ARM resolves named-value KV references. Without this, ARM failed with `does not have secrets get permission on key vault` (Code:ValidationError). | `New-ARMObjects.ps1` (pre-deploy ApiManagement block) |
| V2 APIM URL retrieval | All `Get-AzApiManagement` call sites now fall back to ARM REST when the cmdlet throws `Error mapping types. String -> PsApiManagementSku` (V2 SKU). Reads `properties.gatewayUrl` from the REST response, so the printed `winget source add ...` connection command shows the real APIM gateway instead of falling back to the raw function URL. | `Step-DeployRestSource` (default + custom paths) |
| APIM SKU mismatch on reuse | Phase 3b audit now compares the existing APIM SKU against the SKU implied by `-PerformanceTier` and aborts immediately with a multi-line error block + `Remove-AzApiManagement` remediation when a classic↔V2 cross-tier upgrade is requested. Azure rejects these with `Failed to connect to Management endpoint Port 3443 ... for the <SKU> service` and would otherwise burn ~30 min in retries. Same-family changes (e.g. Basic → Standard) still proceed with a WARN. | `Step-AuditExistingResources` |
| APIM auto-recovery between retries | If an ARM-deploy attempt leaves APIM in `Failed`/`Cancelled`/`Canceled` provisioning state, the deploy retry loop now DELETEs the broken service via REST (works for V2 SKUs too), polls every 30 s up to 15 min for a 404, then DELETEs the soft-deleted shadow under `/providers/Microsoft.ApiManagement/locations/<region>/deletedservices/<name>` so the next attempt can reuse the name. Without this, a transient `ActivationFailed` on attempt 1 cascaded into 4 `ServiceInFailedProvisioningState` failures on the same broken service. Inter-attempt sleep drops 15 s → 5 s when recovery succeeded. | `Step-DeployRestSource` (retry loop) |
| AppConfig soft-delete purge | New 2c block in the Phase 3b audit (parallel to existing KV/APIM checks): lists soft-deleted stores via `GET .../providers/Microsoft.AppConfiguration/deletedConfigurationStores?api-version=2023-03-01`, matches by name, then `POST .../locations/<loc>/deletedConfigurationStores/<name>/purge`. Polls every 10 s up to 120 s. Without this, ARM `appconfig` deployment failed with `NameUnavailable` for up to 7 days after a teardown. Reads `properties.location` (not top-level) from the listing payload — with regex fallback against the resource id. | `Step-AuditExistingResources` (2c) |
| Cosmos `locations[]` dedupe | Added a deduplication pass after the existing `-CosmosDBRegion` / `-CosmosDBZoneRedundant` patches: keep the first occurrence of each `locationName` (case- and whitespace-insensitive) and renumber `failoverPriority` from 0. Fixes ARM `BadRequest: Provided list of regions contains duplicate regions`. | `Step-DeployRestSource` (Cosmos patch block) |
| Phase 2 auth probe | Extended the success regex on the `Get-AzResourceGroup` ARM-reachability probe to also match the newer Az wording `Provided resource group does not exist` and `ResourceNotFound` (was: `NotFound\|ResourceGroupNotFound\|could not be found`). Without this the probe misclassified the 404 as a fatal multi-tenant auth error and aborted Phase 2 before any work began. | `Step-ConnectAzure` |
| Healthy-host signal | `Assert-FunctionAppHealthy` now treats HTTP **401 / 403** as proof the host is up (auth-protected endpoint). Warmup raised 60 s → 180 s and probe timeout 15 s → 30 s to accommodate cold isolated-worker starts. | `Assert-FunctionAppHealthy` |
| Console hygiene | Suppressed the `Invoke-RestMethod` response body that leaked to host after `Purging soft-deleted APIM ...` — the DELETE on `deletedservices` returns the deleted-service object and was being printed as raw object output. | `Step-DeployRestSource` (APIM purge block) |
| RG delete reliability (Remove) | `Remove-WinGetSource.ps1` no longer rethrows on `Receive-Job` errors. Transient `An error occurred while sending the request` from the Az SDK is now captured, the script sleeps 30 s and re-checks `Get-AzResourceGroup` (delete usually succeeded server-side despite the client error), and re-issues `Remove-AzResourceGroup` up to 3 attempts. Final failure message includes a copy-paste fallback command. | `Remove-WinGetSource.ps1` (Step 3) |

### 2026-05-05 (revision 2) — Bundled fork module + live deploy progress

| Area | Change | Function(s) |
|------|--------|-------------|
| Bundled fork | Patched `Microsoft.WinGet.RestSource` module now ships under `Modules\Microsoft.WinGet.RestSource\` and is imported by **absolute path** before any name resolution can happen. PSGallery shadowing is no longer possible. The hot-patch path remains as a fallback. | `Step-ValidatePrerequisites` (1d) |
| Bundled ARM templates | Patched ARM templates (incl. `azurefunction.json`) ship under `Modules\Microsoft.WinGet.RestSource\Data\ARMTemplates\`. The Function App is now born with `FUNCTIONS_WORKER_RUNTIME=dotnet-isolated` and **without** the legacy `FUNCTIONS_INPROC_NET8_ENABLED=1` flag, so the *"in-proc dotnet migration"* deprecation banner never appears in the portal. The fork's `New-WinGetSource` resolves `$PSScriptRoot\..\Data\ARMTemplates` to this bundled folder automatically. | n/a (template-level) |
| APIM SKU hot-patch | Reflection-based widening of `New-WinGetSource` / `New-ARMObjects` ValidateSets via the `validValues` backing field; also patches `allowedValues` arrays in any bundled ARM template (with `.orig` backup). Only runs on the PSGallery-fallback path. | `Test-UpstreamSkuSupport` |
| Live progress | Background poller now emits its first heartbeat after 5 s (was 30 s), shows elapsed / ETA / % of estimated total, and drills into the active ARM **deployment operation** so the long APIM creation surfaces *which sub-resource* ARM is currently working on. Both deploy paths pass `-EstimateMinutes 35` for fresh APIM or `10` when reusing one. | `Watch-DeploymentProgress` |
| Banner / summary | Banner box recomputed with explicit width math (em dash no longer breaks alignment); summary box rebuilt with truncating padding helper so long URLs don't push the right border out of line. | `Write-Banner`, `Write-Summary` |
| Multi-tenant auth | Phase 2 probes ARM with the cached context; if a multi-tenant / expired-credential error is detected it auto-runs `Connect-AzAccount -TenantId <extracted>` and continues. Outer fatal handler now emits a clean *AZURE AUTHENTICATION REQUIRED* block with copy-paste-ready `Connect-AzAccount` / `Set-AzContext` commands instead of a stack trace dump. | `Step-ConnectAzure`, outer `catch` |

### 2026-05-05 — Customer-feedback patch (Deploy-WinGetSource.ps1)

| Area | Change | Function(s) |
|------|--------|-------------|
| Module hygiene | Enumerate **all** on-disk copies of `Microsoft.WinGet.RestSource` and warn when more than one is present; force-remove any in-process instance before re-import; accept `-MinimumVersion` and use `Install-PSResource -Reinstall` on fallback. | `Assert-ModuleAvailable` |
| APIM SKU validation | Reflect on the loaded `New-WinGetSource` cmdlet, extract its `ImplementationPerformance` ValidateSet, and abort **before** ARM work if the requested `-PerformanceTier` (e.g. `StandardV2`) is not in the upstream set. Provides three remediation paths in the error. | `Test-UpstreamSkuSupport` (new) — wired into `Step-ValidatePrerequisites` after module import |
| Function zip upload | New OneDeploy-based publisher using Kudu's `/api/publish?type=zip` REST endpoint with bearer auth from the current Az context. Works on Consumption, Premium, Flex, and `.NET`-isolated worker plans. | `Publish-FunctionZipOneDeploy` (new) |
| Post-deploy health | After every deploy: assert `FUNCTIONS_WORKER_RUNTIME=dotnet-isolated` and `FUNCTIONS_EXTENSION_VERSION=~4`, patch + restart if wrong, then poll `/api/information` for up to 60 s. If still unhealthy, automatically retry the zip upload via OneDeploy. | `Assert-FunctionAppHealthy` (new) — wired into `Step-DeployRestSource` |

#### Why each change exists (customer report)

1. *"the update of powershell module fails, its not putting it into a correct folder"* → addressed by `Assert-ModuleAvailable` enumerating both edition paths and reporting the winning copy.
2. *"creation of apim fails with StandardV2 sku, sor probably that part is not patched?"* → addressed by `Test-UpstreamSkuSupport` failing fast at prerequisite time instead of 30 minutes into ARM deployment.
3. *"upload of updated WinGet.RestSource.Functions.zip failed as well … still got error message with isolated worker"* → addressed by `Publish-FunctionZipOneDeploy` (works on Flex/isolated plans) and `Assert-FunctionAppHealthy` (auto-fixes runtime app-settings and re-publishes via OneDeploy if the host is unhealthy).

---

## FAQ

**Q: Can I manage multiple REST sources?**
A: Yes. The global REST source dropdown in the title bar lets you switch between sources. All tabs automatically use the selected source. Add sources via Settings → Saved Sources, or let the app discover them from your subscription.

**Q: Does the app work without Azure sign-in?**
A: Partially. You can create and save manifests locally, edit YAML files, and browse package configs. Publishing, uploading, managing, deploying, and backup/restore all require Azure authentication.

**Q: What manifest schema versions are supported?**
A: 1.4.0 through 1.9.0. The default is 1.9.0. Change it in Settings → Manifest Version.

**Q: Can I restore a backup to a different environment?**
A: Yes. When restoring, change the Target REST Source and Target Storage Account fields to point to the new environment. This is useful for dev → prod promotion or disaster recovery.

**Q: Is the app portable?**
A: Yes. Copy the entire `WinGetManifestManager/` folder to any Windows machine with the prerequisites installed. User preferences and configs are stored within the app directory.

**Q: How does the app discover Function Apps?**
A: It calls `Get-AzFunctionApp` and filters by: name starts with `func-`, name contains `winget`, or has Azure tags `WinGetSource` or `winget-restsource`.

**Q: What happens if I delete a package from Manage?**
A: The app calls `Remove-WinGetManifest` on the REST source. Optionally, it also deletes the corresponding blob from storage and the local JSON config. A typed confirmation is required for full package deletion.
