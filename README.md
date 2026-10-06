# Endpoint Toolkit

Scripts, templates and tools for Windows endpoint operations, Azure Virtual Desktop,
Windows 365, Intune and selected macOS maintenance tasks.

## Assessment Collectors

Each guide includes prerequisites, examples, a complete argument table and evidence limitations.

| Tool | Guide |
| --- | --- |
| Windows Security Baseline | [BaselineAssessor](tools/BaselineAssessor/README.md#arguments) |
| Azure Virtual Desktop | [AvdAssessor](tools/AvdAssessor/README.md#arguments) |
| Windows 365 | [W365Assessor](tools/W365Assessor/README.md#arguments) |
| Microsoft Intune and companion evidence | [IntuneAssessor](tools/IntuneAssessor/README.md#arguments) |

## Other Tools

| Tool | Purpose |
| --- | --- |
| [ADMXPolicyComparer](tools/ADMXPolicyComparer/) | Compare ADMX policy baselines |
| [AIBLogMonitor](tools/AIBLogMonitor/) | Monitor Azure Image Builder logs |
| [AvdRewind](tools/AvdRewind/) | Roll back AVD session hosts |
| [AzChangeTracker](tools/AzChangeTracker/) | Track Azure resource changes |
| [DeviceDecommissioner](tools/DeviceDecommissioner/) | Guided device removal across management services |
| [PolicyPilot](tools/PolicyPilot/) | Document Group Policy and MDM configuration |
| [WinGetManifestManager](tools/WinGetManifestManager/) | Manage private package manifests |

## Scripts and Templates

| Area | Purpose |
| --- | --- |
| [avd/bicep](avd/bicep/) | Session-host deployment templates |
| [avd/customizer](avd/customizer/) | AIB/Packer image customization and bundled VDOT configuration |
| [avd/pipelines](avd/pipelines/), [avd/scripts](avd/scripts/) | Image builds and session-host lifecycle |
| [devops/aib-task-v2](devops/aib-task-v2/) | Azure Image Builder pipeline task |
| [intune/bitlocker](intune/bitlocker/) | Key-escrow detection/remediation |
| [intune/client-health](intune/client-health/README.md) | Read-only ConfigMgr baseline discovery of Intune client faults |
| [intune/mdm-enrollment](intune/mdm-enrollment/) | Expired enrollment-certificate audit and opt-in repair |
| [intune/mdm-sync-service](intune/mdm-sync-service/README.md) | Local diagnostics, service repair and opt-in sync |
| [intune/onedrive-photos](intune/onedrive-photos/) | Shortcut-only detection/remediation |
| [macos/servicing](macos/servicing/) | Developer-storage cleanup |
| [windows/applications](windows/applications/UninstallMsiProduct/README.md) | Registry-based MSI uninstallation |
| [windows/configuration](windows/configuration/README.md) | Startup delay, power plans and [processor boost](windows/configuration/ProcessorBoost/README.md) |
| [windows/diagnostics](windows/diagnostics/) | Location, Defender coexistence, Delivery Optimization, power and audio diagnostics |
| [NTLM usage](windows/diagnostics/NtlmUsageDetection/README.md) | Read-only NTLM evidence detector for Ivanti |
| [windows/dot3svc](windows/dot3svc/) | Wired AutoConfig migration reset |
| [EntraCutover](windows/migration/EntraCutover/README.md) | Experimental Hybrid-to-Entra migration; not supported by Microsoft |
| [windows/print](windows/print/) | Windows Protected Print readiness |
| [windows/rdp](windows/rdp/) | Per-user RDP file signing |
| [windows/security](windows/security/) | Speculation mitigation and Secure Boot servicing |
| [windows/servicing](windows/servicing/) | Upgrade cleanup, recovery-partition resize and ESP reporting |
| [windows/w365](windows/w365/) | Cloud PC disk and keyboard utilities |

## Before Running

Read the tool-specific README and `Get-Help <script> -Full`; privileges, modules and supported
PowerShell versions vary. WPF tools require Windows. Pipeline `<YOURVALUE>` placeholders must
be replaced. Review repair/deletion actions in a lab and use preview modes where supported.
Assessment evidence may contain confidential configuration and identifiers; follow the relevant
collector's data-handling guidance.

[MIT license](LICENSE). Scripts are provided "AS IS" with no warranties and confer no rights.