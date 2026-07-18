# AVD / Windows 365 Image Builder Customizer Scripts

A curated set of PowerShell customizers for Azure Image Builder (AIB) /
Packer / Run-Command pipelines that bake AVD and Windows 365 reference
images.

All scripts in this folder follow the same conventions:

- `#Requires -RunAsAdministrator` (every script expects the AIB SYSTEM context, except `InstallProvisionedAppxPackage.ps1` which is designed for ADMIN user context).
- `[CmdletBinding()]` + comment-based help (`Get-Help .\<script>.ps1 -Full` works).
- `$ErrorActionPreference = 'Stop'` at the top.
- Uniform `Write-Log` helper writing `[utc-timestamp] [LEVEL] [ScriptName] message`
  to the AIB log via Write-Host (AIB only captures stdout).
- Non-zero `exit` on hard failure so AIB / Packer flags the customizer step red.

## Script index

| Script | Purpose | Stage |
|---|---|---|
| [AdminSysPrep.ps1](AdminSysPrep.ps1) | Patches `C:\DeprovisioningScript.ps1` so Sysprep runs with `/quit /mode:vm` | Late, just before image capture |
| [DisableAutoUpdates.ps1](DisableAutoUpdates.ps1) | Blocks Store auto-downloads + Content Delivery Manager + WU Scheduled Start during bake | Early bake |
| [InstallLanguagePacks.ps1](InstallLanguagePacks.ps1) | Installs one or more Windows Display Languages with retry + LanguageComponentsInstaller race-fix | Mid bake |
| [InstallProvisionedAppxPackage.ps1](InstallProvisionedAppxPackage.ps1) | Side-loads a single LOB UWP / MSIX bundle as a provisioned package via `Add-AppxProvisionedPackage -Online`. Auto-discovers license + `Dependencies\`, snapshots and restores `AllowAllTrustedApps` (with `-KeepSideloadingEnabled` for chained installs), maps HRESULT 0x80073D02 (pending reboot) to success | Mid bake (admin context) |
| [RemoveAppxPackages.ps1](RemoveAppxPackages.ps1) | De-provisions inbox AppX packages by wildcard name (`*Bing*`, `Microsoft.MSPaint`, …). Exits non-zero if any individual package/capability removal fails, unless `-ContinueOnError` is passed to fall back to best-effort (always exit 0) | Mid bake |
| [RemoveUserApps.ps1](RemoveUserApps.ps1) | Removes per-user AppX packages with no matching provisioned package — fixes Sysprep 0x80073CF2 | Late bake (immediately before AdminSysPrep) |
| [ResetAutoUpdateSettings.ps1](ResetAutoUpdateSettings.ps1) | Reverts the bake-time hardening (Windows Update + Store + CDM) | Post-deploy Run-Command only — **no longer part of the bake chain** |
| [TimezoneRedirection.ps1](TimezoneRedirection.ps1) | Sets `fEnableTimeZoneRedirection = 1` for RDS / AVD time-zone follow | Anywhere |
| [UpdateWinGet.ps1](UpdateWinGet.ps1) | Hardens, downloads + provisions WinGet, registers `-CustomSources`, optionally installs `-AppIds` with `--scope machine` (per-app source override supported) | Mid–late bake |
| [WindowsOptimization.ps1](WindowsOptimization.ps1) | Hardened wrapper around the Virtual Desktop Optimization Tool (VDOT). Ships VDOT JSON in-repo under [ConfigurationFiles/](ConfigurationFiles/) so no internet egress is required at bake time. Resilient access-denied handling, file logger, `-ConfigBasePath` (override / air-gapped path), `-LogDirectory`, `-ContinueOnError` | Late bake |

## Recommended pipeline order

This is the actual chain run by both [`img-build-custom-image.yml`](../pipelines/img-build-custom-image.yml)
and [`img-build-bicep-only.yml`](../pipelines/img-build-bicep-only.yml) — the
two pipelines are kept identical here by design (see
[pipelines/README.md](../pipelines/README.md#image-build-approaches)). Every
step is exit-code-checked; a non-zero exit from any customizer fails the build.

1. `DisableAutoUpdates.ps1`
2. `TimezoneRedirection.ps1`
3. `InstallLanguagePacks.ps1` (optional — commented out by default in both pipelines)
4. `RemoveAppxPackages.ps1` (de-bloat)
5. `UpdateWinGet.ps1` (provision WinGet + install `-AppIds`)
6. `WindowsOptimization.ps1 -Optimizations <selective list>` (VDOT pass)
7. `UpdateWinGet.ps1 -SkipApps -SkipUserRegistration` (provision-only re-run — re-registers WinGet after the VDOT pass, which can strip its AppX provisioning association)
8. `RemoveUserApps.ps1` (Sysprep prep — must run immediately before AdminSysPrep)
9. `AdminSysPrep.ps1` (must run **LAST**, immediately before Sysprep/capture)
10. *Sysprep / capture step (handled by AIB)*
11. *(optional, post-deploy only)* `ResetAutoUpdateSettings.ps1` on deployed hosts that need updates re-enabled — this is **no longer** run as an AIB bake step in either pipeline.

`InstallProvisionedAppxPackage.ps1` is not wired into either pipeline by
default — add it after `RemoveAppxPackages.ps1` / before `WindowsOptimization.ps1`
for per-LOB UWP / MSIX sideloading if your image needs it.

## Companions outside this folder

- [`avd/scripts/Get-StubAppPayloads.ps1`](../scripts/Get-StubAppPayloads.ps1) — pre-stage Microsoft Store stub-app payloads on a workstation.
- [`avd/scripts/Install-AppxPayloads.ps1`](../scripts/Install-AppxPayloads.ps1) — side-load the staged payloads during bake (fixes stub-app provisioning).
- [`windows/servicing/Invoke-PreUpgradeCleanup.ps1`](../../windows/servicing/Invoke-PreUpgradeCleanup.ps1) — disk-space cleanup before / after image bake.

## Conventions

- **Author:** Anton Romanyuk (with attribution where logic was adapted, e.g.
  Michael Niehaus's Sysprep cleanup pattern in `RemoveUserApps.ps1`).
- **Logging destination:** Write-Host (AIB / Packer log capture).
- **Exit codes:** `0` on success, `1` on hard failure.
- **Idempotency:** every script is safe to re-run; missing keys / packages are
  treated as already-clean and logged at INFO.

## Bundled VDOT configuration

[`ConfigurationFiles/`](ConfigurationFiles/) ships the seven Virtual Desktop
Optimization Tool JSON files used by `WindowsOptimization.ps1`:

| File | Purpose |
|---|---|
| `ScheduledTasks.json` | VDI-hostile scheduled tasks to disable. |
| `DefaultUserSettings.json` | Registry tweaks applied to the Default user hive (`C:\Users\Default\NTUSER.DAT`). |
| `Autologgers.Json` | Trace autologgers to disable. |
| `Services.json` | VDI-hostile services to set to `Disabled`. |
| `LanManWorkstation.json` | LanManWorkstation network tunings. |
| `PolicyRegSettings.json` | Local Group Policy registry settings (LGPO equivalents). |
| `EdgeSettings.json` | Microsoft Edge policy registry settings. |

Source: [`The-Virtual-Desktop-Team/Virtual-Desktop-Optimization-Tool`](https://github.com/The-Virtual-Desktop-Team/Virtual-Desktop-Optimization-Tool)
(`main/2009/ConfigurationFiles`). The JSON is bundled in-repo so AIB / Packer
bakes do **not** require runtime egress to `raw.githubusercontent.com`
(eliminates proxy / air-gap / supply-chain failure modes and makes runs
deterministic).

To refresh from upstream:

```powershell
$dir  = Join-Path $PSScriptRoot 'ConfigurationFiles'
$base = 'https://raw.githubusercontent.com/The-Virtual-Desktop-Team/Virtual-Desktop-Optimization-Tool/main/2009/ConfigurationFiles'
'ScheduledTasks.json','DefaultUserSettings.json','Autologgers.Json','Services.json',
'LanManWorkstation.json','PolicyRegSettings.json','EdgeSettings.json' |
    ForEach-Object { Invoke-WebRequest "$base/$_" -OutFile (Join-Path $dir $_) -UseBasicParsing }
```

Override the bundled folder with `-ConfigBasePath <path>` (e.g. an air-gapped
build share). Pass `-ConfigBasePath ''` to opt into the legacy GitHub
download path (NOT recommended).
