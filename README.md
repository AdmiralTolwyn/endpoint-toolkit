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
└── aib-task-v1-patched/   # Patched Azure Image Builder DevOps task (v2)

intune/
└── bitlocker/      # BitLocker detection & remediation scripts for Intune

tools/              # Standalone PowerShell/WPF utilities

windows/
├── dot3svc/        # Wired AutoConfig (dot3svc) migration reset
├── rdp/            # Per-user RDP file signing (no admin required)
├── security/       # Hardware speculation mitigations, Secure Boot remediation
├── servicing/      # Pre-upgrade disk-space cleanup (cleanmgr + DISM)
└── w365/           # Windows 365 Cloud PC disk resize automation
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
| [windows/dot3svc/](windows/dot3svc/) | Reset 802.1X / wired-AutoConfig profiles after migration |
| [windows/rdp/](windows/rdp/) | Sign `.rdp` files in user context (no admin required) |
| [windows/security/](windows/security/) | Hardware speculation mitigations + Secure Boot UEFI CA 2023 remediation (Intune PR pair) |
| [windows/servicing/](windows/servicing/) | `Invoke-PreUpgradeCleanup.ps1` — reclaim disk space via cleanmgr + DISM before a feature update or after image bake |
| [windows/w365/](windows/w365/) | Windows 365 Cloud PC disk resize automation |

## Getting Started

Most pipeline files use `<YOURVALUE>` placeholders — search for `<YOUR` and replace with your environment-specific values before use.

## Requirements

- PowerShell 5.1+
- Azure CLI / Az PowerShell modules (for AVD scripts and pipelines)
- Windows 11 (for WPF-based tools)

## License

MIT
