# AVD / Windows 365 Image Builder Customizer Scripts

A curated set of PowerShell customizers for Azure Image Builder (AIB) /
Packer / Run-Command pipelines that bake AVD and Windows 365 reference
images.

All scripts in this folder follow the same conventions:

- `#Requires -RunAsAdministrator` (every script expects the AIB SYSTEM context).
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
| [RemoveAppxPackages.ps1](RemoveAppxPackages.ps1) | De-provisions inbox AppX packages by wildcard name (`*Bing*`, `Microsoft.MSPaint`, …) | Mid bake |
| [RemoveUserApps.ps1](RemoveUserApps.ps1) | Removes per-user AppX packages with no matching provisioned package — fixes Sysprep 0x80073CF2 | Late bake (just before AdminSysPrep) |
| [ResetAutoUpdateSettings.ps1](ResetAutoUpdateSettings.ps1) | Reverts the bake-time hardening (Windows Update + Store + CDM) | Last AIB step OR post-deploy Run-Command |
| [TimezoneRedirection.ps1](TimezoneRedirection.ps1) | Sets `fEnableTimeZoneRedirection = 1` for RDS / AVD time-zone follow | Anywhere |
| [UpdateWinGet.ps1](UpdateWinGet.ps1) | Hardens, downloads + provisions WinGet, registers `-CustomSources`, optionally installs `-AppIds` with `--scope machine` (per-app source override supported) | Mid–late bake |
| [WindowsOptimization.ps1](WindowsOptimization.ps1) | Hardened wrapper around the Virtual Desktop Optimization Tool (VDOT). Ships VDOT JSON in-repo under [ConfigurationFiles/](ConfigurationFiles/) so no internet egress is required at bake time. Resilient access-denied handling, file logger, `-ConfigBasePath` (override / air-gapped path), `-LogDirectory`, `-ContinueOnError` | Late bake |

## Recommended pipeline order

1. `DisableAutoUpdates.ps1`
2. `TimezoneRedirection.ps1`
3. `InstallLanguagePacks.ps1` (if needed)
4. `RemoveAppxPackages.ps1` (de-bloat)
5. `UpdateWinGet.ps1` (provision WinGet + apps)
6. `WindowsOptimization.ps1 -Optimizations All` (or selective)
7. `RemoveUserApps.ps1` (Sysprep prep — must run AFTER any user-context installs)
8. `AdminSysPrep.ps1`
9. *Sysprep / capture step (handled by AIB)*
10. *(optional)* `ResetAutoUpdateSettings.ps1` on deployed hosts that need updates re-enabled.

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
