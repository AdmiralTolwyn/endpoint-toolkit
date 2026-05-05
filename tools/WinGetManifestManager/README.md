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
| Performance Tier | `Developer` | `Developer` / `Basic` / `Enhanced` — affects Cosmos DB and App Service SKUs. |
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

### Deploy script fails

The deploy script requires **PowerShell 7.4+** (`pwsh`). If you launched the app with Windows PowerShell 5.1, the deploy script will spawn a separate `pwsh` process. Ensure PowerShell 7 is installed.

### Backup shows 0 manifests

Verify the REST source Function App name is correct and the source has published packages. The backup uses `Find-WinGetManifest` → `Get-WinGetManifest` which requires the `Microsoft.WinGet.RestSource` module.

### Soft-deleted resources block redeployment

Use the **Remove** function with default settings — it automatically purges soft-deleted Key Vaults, APIM instances, and App Services that block name reuse.

### Deploy fails with `Standardv2` (or any V2 SKU) not in ValidateSet

The wrapper accepts the V2 SKU names but the **upstream** `Microsoft.WinGet.RestSource` module on PSGallery still has the legacy ValidateSet (`Developer,Basic,Standard,Premium,Consumption`). If a stale upstream copy on disk wins the version race, the inner `New-WinGetSource` rejects `BasicV2`/`StandardV2` 30+ minutes into the deploy.

The script now (a) enumerates **every** copy of `Microsoft.WinGet.RestSource` on disk and warns when more than one is present, and (b) probes the loaded module's ValidateSet **before** any ARM work via `Test-UpstreamSkuSupport`. If the requested SKU is not supported, the deploy aborts immediately with three resolution paths: install the patched fork at a higher version, hot-patch the installed copy's `Library\New-WinGetSource.ps1`, or pick a supported tier and resize APIM after deploy with `Update-AzApiManagement -Sku StandardV2`.

### Function App returns "isolated worker" error after deploy

The .NET-isolated Functions host fails to start when `FUNCTIONS_WORKER_RUNTIME` ≠ `dotnet-isolated` or when `FUNCTIONS_EXTENSION_VERSION` ≠ `~4`. Upstream's zip upload (via `Publish-AzWebApp -ArchivePath`) does not always set these and silently fails on Flex Consumption / isolated-worker plans.

The script now runs `Assert-FunctionAppHealthy` after every deploy: it checks the two app settings, patches them via `Update-AzFunctionAppSetting` if wrong, restarts the host, and pings `/api/information` for up to 60 s. If the host is still unhealthy, it falls back to a OneDeploy re-publish (`Publish-FunctionZipOneDeploy`) which uses Kudu's `/api/publish?type=zip` REST endpoint with bearer auth — works on Consumption, Premium, Flex, and isolated-worker plans alike.

### `Microsoft.WinGet.RestSource` module update lands in the wrong folder

`Install-Module` / `Install-PSResource` install per edition: PowerShell 7 writes to `$HOME\Documents\PowerShell\Modules\`, Windows PowerShell 5.1 writes to `$HOME\Documents\WindowsPowerShell\Modules\`. If both paths contain copies, PowerShell loads the highest version regardless of edition, which is how a stale upstream copy can beat a freshly-installed fork. The deploy script now lists every copy it finds and `Remove-Module -Force`s any in-process instance before importing, so an explicit `Install-PSResource -Reinstall` actually takes effect on the next session.

---

## Changelog

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
