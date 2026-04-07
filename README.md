# Endpoint Toolkit

Production-ready PowerShell tools for Windows endpoint security, compliance assessment, and operational management across Azure Virtual Desktop, Intune, and hybrid-joined environments.

## What's Inside

| Tool | Description | Status |
|------|-------------|--------|
| **BaselineAssessor** | Security baseline collector + assessor with 323 checks mapped to Microsoft SCT, CIS, and Intune baselines. Collects 22 data areas, evaluates GPO + CSP/MDM policies, generates compliance reports. | Active |
| **Reset-Dot3SvcMigration** | Wired 802.1x profile remediation for post-upgrade scenarios. Handles symlinks, Windows.old recovery, iterative service resets. | Active |
| *More tools coming* | WinGet manifest management, image builder automation, AVD operations | Planned |

## Quick Start

### BaselineAssessor

**Collect** (run as admin on target machine):

```powershell
.\Invoke-BaselineCollection.ps1
```

**Assess** (run on any machine with the JSON):

```powershell
.\BaselinePilot.ps1
# Import the collection JSON via the UI
```

### Reset-Dot3SvcMigration

```powershell
.\Reset-Dot3SvcMigration.ps1 -RepairPolicies
```

## Requirements

- PowerShell 5.1+ (Windows built-in)
- Local administrator for data collection
- No external modules required

## Supported Environments

- **Azure Virtual Desktop** (session hosts, personal desktops)
- **Windows 365 Cloud PC**
- **Intune-managed devices** (Entra ID joined)
- **Hybrid Azure AD joined** (GPO + MDM co-management)
- **Domain-joined** (traditional AD DS with GPO)
- **Windows 11 22H2+** (25H2 recommended)

## Architecture

### BaselineAssessor

```
Invoke-BaselineCollection.ps1    Headless collector (22 areas, ~45s)
        |
        v
  hostname_baseline_<ts>.json    Machine security snapshot
        |
        v
BaselinePilot.ps1                WPF assessor UI (323 checks)
        |
    checks.json                  Check definitions + baselines
    csp_metadata.json            CSP-to-registry mappings
    admx_metadata.json           ADMX policy definitions
```

**Dual-path evaluation**: Checks resolve against both GPO registry paths (`SOFTWARE\Policies\`) and CSP/MDM paths (`SOFTWARE\Microsoft\`) with `MDMWinsOverGP` precedence for hybrid devices.

## License

MIT
