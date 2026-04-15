# Hardware Speculation Vulnerability Detection

## Overview

Detects CPU hardware vulnerabilities (MDS and SSB) using low-level `ntdll.dll` queries against the Windows speculation control information interface. Designed for use as an **Ivanti Custom Vulnerability Detection** script.

Based on the official [Microsoft SpeculationControl PowerShell module](https://www.powershellgallery.com/packages/SpeculationControl).

## Script

| File | Purpose |
|------|---------|
| `Get-HardwareSpeculationStatus.ps1` | Queries CPU speculation control flags and reports vulnerability status |

## Vulnerabilities Detected

| Advisory | Name | Description |
|----------|------|-------------|
| [ADV190013](https://msrc.microsoft.com/update-guide/vulnerability/ADV190013) | MDS (Microarchitectural Data Sampling) | Checks `MdsHardwareProtected` flag; AMD/ARM architectures treated as protected |
| [ADV180012](https://msrc.microsoft.com/update-guide/vulnerability/ADV180012) | SSB (Speculative Store Bypass) | Checks `SsbdRequired` flag to determine if SSBD mitigation is needed |

## Usage

```powershell
# Check both MDS and SSB (default)
.\Get-HardwareSpeculationStatus.ps1

# Check MDS only
# Edit the execution section to use: Get-HardwareSpeculationStatus -MdsOnly

# Check SSB only
# Edit the execution section to use: Get-HardwareSpeculationStatus -SsbOnly
```

## Output Format

The script outputs Ivanti-compatible detection fields:

- `detected` — `true` if vulnerable, `false` if protected or query failed
- `reason` — Human-readable explanation
- `expected` — Target-state string
- `found` — Actual-state string

## How It Works

1. Loads `NtQuerySystemInformation` from `ntdll.dll`
2. Queries system speculation control information (class 201)
3. Reads hardware protection flags from the returned bitmask
4. Cross-references CPU manufacturer/architecture for vendor-level protections (AMD, ARM)
5. Reports vulnerability status per advisory

## Requirements

- Windows 10/11 or Windows Server 2016+
- PowerShell 5.1+
- Administrator privileges (for `ntdll` system information query)
