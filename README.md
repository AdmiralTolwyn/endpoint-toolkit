# Endpoint Toolkit

A collection of scripts, templates, and tools for managing Windows endpoints at scale — covering Azure Virtual Desktop image builds, session host lifecycle, and day-to-day operational tasks.

## Repository Structure

```
avd/
├── bicep/          # Bicep templates for AVD session host deployment
│   ├── modules/    # Reusable modules (session hosts, image templates)
│   └── main-*.bicep
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

## Getting Started

Most pipeline files use `<YOURVALUE>` placeholders — search for `<YOUR` and replace with your environment-specific values before use.

## Requirements

- PowerShell 5.1+
- Azure CLI / Az PowerShell modules (for AVD scripts and pipelines)
- Windows 11 (for WPF-based tools)

## License

MIT
