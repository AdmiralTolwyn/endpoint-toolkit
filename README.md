# Endpoint Toolkit

A collection of scripts, templates, and tools for managing Windows endpoints at scale — covering Azure Virtual Desktop image builds, session host lifecycle, and day-to-day operational tasks.

## Repository Structure

```
avd/
├── bicep/          # Bicep templates for AVD session host deployment
│   ├── modules/    # Reusable modules (session hosts, image templates)
│   └── main-*.bicep
├── customizer/     # AIB / Packer customizer scripts (image-bake)
│   └── ConfigurationFiles/  # Bundled VDOT JSON (no runtime download required)
├── pipelines/      # Azure DevOps YAML pipelines
└── scripts/        # PowerShell scripts used by pipelines

devops/
└── aib-task-v2/           # Azure Image Builder DevOps task (v2)

intune/
├── bitlocker/        # BitLocker detection & remediation scripts for Intune
└── mdm-enrollment/   # Repair expired Intune MDM device cert (omadmclient high-CPU)

macos/
└── servicing/      # Developer-storage cleanup and reclaim helpers for macOS

tools/              # Standalone PowerShell/WPF utilities

windows/
├── applications/   # Generic MSI uninstaller by name pattern / publisher / GUID
├── configuration/  # Modern Standby power plan, Windows 11 startup-app delay revert
├── diagnostics/    # Read-only endpoint diagnostics — Location policy state, Defender/EDR coexistence
├── dot3svc/        # Wired AutoConfig (dot3svc) migration reset
├── migration/      # Hybrid Join → Entra-only in-place migration (EntraCutover); stale/cross-tenant MDM enrollment detection & purge (StaleMdmEnrollment)
├── print/          # Windows Protected Print (WPP) readiness — flag third-party v3/v4 drivers
├── rdp/            # Per-user RDP file signing (no admin required)
├── security/       # Hardware speculation mitigations, Secure Boot CA 2023 remediation + fleet inventory
├── servicing/      # Pre-upgrade disk-space cleanup, WinRE partition resize, ESP free-space reporter, in-place upgrade timeline
└── w365/           # Windows 365 Cloud PC utilities (disk resize, keyboard layout)
```

## Tools

| Tool | Description |
|------|-------------|
| [ADMXPolicyComparer](tools/ADMXPolicyComparer/) | Compare ADMX policy baselines across Windows versions |
| [AIBLogMonitor](tools/AIBLogMonitor/) | Azure Image Builder log monitor |
| [AvdAssessor](tools/AvdAssessor/) | AVD environment assessment |
| [AvdRewind](tools/AvdRewind/) | AVD session host rollback |
| [AzChangeTracker](tools/AzChangeTracker/) | Azure resource change tracking |
| [BaselineAssessor](tools/BaselineAssessor/) | Windows security baseline assessment (263 checks) |
| [DeviceDecommissioner](tools/DeviceDecommissioner/) | Remove a device from AD, Entra ID, Intune, Autopilot, and SCCM in one guided workflow — pre-flight cards, BitLocker/LAPS warnings, dry-run, audit trail |
| [DumpPilot](tools/DumpPilot/) | Crash-dump triage — drives `cdb.exe` to extract 40+ structured fact categories (exception, stacks, registers, module vendors/symbols, VAS/heap/handle stats, disassembly at fault, NTSTATUS decode) into deterministic JSON, then pattern-DB matching, an HTML report and an escalation-grade LLM prompt. User-mode minidump/full and kernel dumps; WPF GUI (STA) plus single-dump and batch CLI; optional ProcMon correlation. Read-only post-mortem |
| [PolicyPilot](tools/PolicyPilot/) | Group Policy & MDM documentation — scans AD/Local/Intune, conflict detection, ADMX/CSP enrichment |
| [W365Assessor](tools/W365Assessor/) | Windows 365 (Cloud PC) Enterprise & Frontline tenant assessment — 128 checks, 23 automated via Microsoft Graph |
| [WinGetManifestManager](tools/WinGetManifestManager/) | WinGet package manifest manager for private repos |

## Scripts

| Area | Description |
|------|-------------|
| [avd/customizer/](avd/customizer/) | AIB / Packer image-bake customizers — AdminSysPrep, DisableAutoUpdates, InstallLanguagePacks, RemoveAppxPackages, RemoveUserApps, ResetAutoUpdateSettings, TimezoneRedirection, UpdateWinGet, WindowsOptimization (VDOT wrapper, JSON bundled in-repo) |
| [avd/scripts/](avd/scripts/) | AVD pipeline helpers — host-pool drain, deployment telemetry, FSLogix repair, Get-StubAppPayloads / Install-AppxPayloads, hybrid activator, Remove-AvdHosts |
| [avd/pipelines/](avd/pipelines/) | Azure DevOps YAML pipelines for AVD activation, host-pool updates, image bakes |
| [avd/bicep/](avd/bicep/) | Bicep templates for AVD session-host deployment (Entra ID + AD-joined variants) |
| [intune/bitlocker/](intune/bitlocker/) | Intune Proactive Remediation pair — ensure BitLocker recovery key escrow to Entra ID; MBAM client uninstall |
| [intune/mdm-enrollment/](intune/mdm-enrollment/) | `Repair-IntuneMdmCert.ps1` — audit (read-only) or repair hosts whose expired Intune MDM device cert wedges `omadmclient.exe` at high CPU. Repair tears down the enrollment + re-enrolls via device credential. Built for cloned AVD fleets that expire together |
| [macos/servicing/](macos/servicing/) | `macos_dev_cleanup.sh` — semi-interactive developer-storage cleanup (Xcode, VS Code/Cursor/Windsurf, .NET, Gradle, Android, Flutter, JetBrains, Homebrew, Docker, Time Machine) |
| [windows/applications/](windows/applications/UninstallMsiProduct/README.md) | `Uninstall-MsiProduct.ps1` — generic MSI uninstaller by DisplayName / Publisher / Version / ProductCode wildcards. Registry-driven (no `Win32_Product` side effects); built for vendor agents whose GUID changes per release (e.g. Quest / KACE Agent) |
| [windows/configuration/](windows/configuration/README.md) | `Set-ModernStandbyPowerPlan.ps1` — duplicate Balanced into a named plan for S0 Low Power Idle devices and set power-button, lid, screen and sleep-idle behaviour on AC and DC; idempotent, detects and updates its own plan. `Set-StartupAppsDelay.ps1` — revert the Windows 11 wait-for-idle startup delay (`Explorer\Serialize\WaitForIdleState`) on devices where Defender, OneDrive, ESP or an EDR agent never lets the machine reach idle. Both ship verbatim as Intune platform scripts with no runtime parameters; the startup script auto-adapts to context, patching HKCU when non-elevated and the Default User hive plus every loaded `HKU\<SID>` when SYSTEM. Not policy-backed — no ADMX or Settings Catalog equivalent exists |
| [windows/diagnostics/](windows/diagnostics/) | [`LocationPolicyState/Get-LocationPolicyState.ps1`](windows/diagnostics/LocationPolicyState/README.md) — report the effective Windows Location policy state and every author that can force/lock the toggle. [`MdeCoexistenceState/Get-MdeCoexistenceState.ps1`](windows/diagnostics/MdeCoexistenceState/README.md) — effective Defender AV / Defender for Endpoint state, detection of third-party AV/EDR sharing the endpoint (minifilters by altitude band, services, Security Center), sensor health from the SENSE log, and an automated exclusion-hygiene review that catches `%USERPROFILE%`-style rules that silently match nothing under LocalSystem. Read-only, JSON output, Intune exit codes |
| [windows/dot3svc/](windows/dot3svc/) | Reset 802.1X / wired-AutoConfig profiles after migration |
| [windows/migration/](windows/migration/EntraCutover/README.md) | `EntraCutover` — **experimental, not supported by Microsoft.** In-place Hybrid Join → Entra-only join migration (no reinstall). Resumable 5-phase state machine (Assess/Prepare/Teardown/Join/Finalize), Intune enrollment + stale-GPO cleanup, fresh-profile + OneDrive KFM, BitLocker re-escrow to the new device object, break-glass admin + `djoin` offline-rejoin rollback. CLI, CMTrace logging |
| [windows/migration/StaleMdmEnrollment/](windows/migration/StaleMdmEnrollment/README.md) | `Get-StaleMdmEnrollment.ps1` — read-only detector for devices left cross-tenant or orphaned after a tenant migration. Evaluates **both** MDM channels independently (`MS DM Server` and the declared-configuration / MMP-C channel), because a device whose Intune enrollment was retired away still counts as enrolled off the surviving one — so auto-enrollment and `deviceenroller.exe` exit silently with no events at all. Per-channel liveness scoring separates a live enrollment from teardown residue; reads `LinkedEnrollment` and `MmpcEnrollmentFlag`. JSON output for fleet collection. `Remove-StaleMdmEnrollment.ps1` — **unsupported**, dry-run by default: purges one enrollment GUID (tasks, eight registry roots, client cert) with full backup and a guard that refuses a healthy enrollment. The supported remediation is a device reset; the README grades every claim as measured / Microsoft-documented / community-only |
| [windows/print/](windows/print/) | `Get-PrintDriverWppReadiness.ps1` — flag machines with third-party v3/v4 print drivers (not yet Windows Protected Print ready) ahead of WPP enforcement. Intune Proactive Remediation detection script (exit 0/1) + standalone CSV/JSON fleet inventory; maps drivers to printers actually using them. Read-only |
| [windows/rdp/](windows/rdp/) | Sign `.rdp` files in user context (no admin required) |
| [windows/security/](windows/security/) | Hardware speculation mitigations. [**Secure Boot UEFI CA 2023**](windows/security/SecureBootRemediation/README.md) — two script families in one folder. *Drive the update*: Intune PR pair, standalone triage script, Ivanti detect — these write `AvailableUpdates` and trigger the Secure-Boot-Update task. *Measure the fleet*: `Get-SecureBootCertInventory.ps1` read-only JSON inventory (servicing state, `AvailableUpdates` bitmask, KEK/db certificate presence, derived Microsoft-only vs. third-party trust configuration, TPM-WMI events, confidence bucket, platform class), `Repair-SecureBootReportingPrereqs.ps1` which fixes only what blocks *reporting* and never writes `AvailableUpdates`, and `Split-SecureBootPopulation.ps1` which segments the Intune export so telemetry gaps stop masquerading as certificate failures |
| [windows/servicing/](windows/servicing/) | `Invoke-PreUpgradeCleanup.ps1` — reclaim disk space via cleanmgr + DISM before a feature update or after image bake. `Resize-RecoveryPartition.ps1` — resize the WinRE recovery partition (KB5034441 / CVE-2024-20666 remediation). `Get-EspPartitionStatus.ps1` — EFI System Partition size/free reporter as JSON for Grafana/Loki/Telegraf (KB5089549 / 0x800f0922 monitoring). [`Get-SetupTimeline.ps1`](windows/servicing/SetupTimeline/README.md) — parse `setupact.log` into a phase-by-phase in-place upgrade timeline, subtracting idle gaps where the device was off or asleep so the reported figure is real lockout time rather than wall clock; splits online phases (user still productive, excluded by default) from offline; optional Dynamic Update download breakdown with throughput and slow-link time; table, CSV, object or single-integer output for CI. Read-only |
| [windows/w365/](windows/w365/) | Windows 365 Cloud PC utilities — disk resize, keyboard layout configuration |

## Getting Started

Most pipeline files use `<YOURVALUE>` placeholders — search for `<YOUR` and replace with your environment-specific values before use.

## Requirements

- PowerShell 5.1+
- Azure CLI / Az PowerShell modules (for AVD scripts and pipelines)
- Windows 11 (for WPF-based tools)

## License

MIT
