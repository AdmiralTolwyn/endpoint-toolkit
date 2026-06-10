# Get-PrintDriverWppReadiness

**Version:** 1.0
**Author:** Anton Romanyuk

> **Disclaimer:** This script is provided "as-is" without warranty of any kind, express or implied. Use at your own risk. The author assumes no liability for any damage or data loss resulting from its use. Always test in a non-production environment before deployment.

Flags machines that still have **third-party v3/v4 print drivers** — i.e. devices that are **not yet Windows Protected Print (WPP) ready** — so a fleet can be screened proactively before WPP is enforced (e.g. ahead of a 27H2 rollout).

## Problem

[Windows Protected Print Mode (WPP)](https://learn.microsoft.com/en-us/windows-hardware/drivers/print/windows-protected-print-mode) only allows printing through the inbox **Microsoft IPP Class Driver** and a small subset of inbox Microsoft print drivers. Any **third-party v3** (classic kernel/user-mode) or **v4** print driver is incompatible and will stop working once WPP is enforced.

Microsoft has begun **end-of-servicing for legacy v3/v4 third-party print drivers**, and WPP is moving toward becoming the default. Before you can flip that switch fleet-wide, you need to know which machines still carry blocking drivers — and ideally which of those drivers are actually bound to a printer (real risk) versus just present (orphaned).

There is no built-in report for this, so this script provides one.

## Solution

The script enumerates installed print drivers via `Get-PrinterDriver`, then:

1. **Classifies** each driver as **WPP-safe** (Microsoft inbox) or **blocking** (third-party v3/v4)
2. **Reads `MajorVersion`** to label each blocking driver `v3` or `v4`
3. **Maps printers → drivers** (via `Get-Printer`) so the report shows which blocking drivers are actually *in use*
4. **Reports the on-box WPP policy state** (enforced via local state or Group Policy / MDM)
5. **Emits a single-line JSON summary** and an exit code for Intune Proactive Remediation

It is **read-only** — it never installs, removes, or changes any driver or policy.

## Two ways to run it

### 1. Intune Proactive Remediation — detection script

Paste the script into the **Detection script** slot (run in 64-bit PowerShell, **system context**). It exits:

| Exit code | Meaning |
|---|---|
| `0` | WPP-ready — no third-party v3/v4 drivers found |
| `1` | **Not** WPP-ready — one or more blocking drivers found |

The one-line JSON it writes to STDOUT surfaces in the pre-remediation output column, so you get fleet-wide reporting without a separate collection mechanism:

```json
{"Computer":"PC01","WppReady":false,"WppEnforced":false,"Total":12,"Blocking":3,"BlockingV3":1,"BlockingV4":2,"BlockingInUse":1,"Drivers":["v4:HP Photosmart 7520 series Class Driver","v3:Canon GX6000 series"]}
```

> No companion *remediation* script is shipped — removing print drivers is destructive and printer-specific. Use the detection signal to target machines, then remediate with a scoped, reviewed process.

### 2. Standalone fleet inventory

```powershell
# Full per-driver report (safe + blocking) to CSV
.\Get-PrintDriverWppReadiness.ps1 -CsvPath C:\Temp\wpp-drivers.csv -IncludeSafeDrivers

# Structured result (drivers + mapped printers + WPP state) to JSON
.\Get-PrintDriverWppReadiness.ps1 -JsonPath C:\Temp\wpp.json

# Fan out across a host list via remoting and collect the JSON summaries
Invoke-Command -ComputerName (Get-Content .\hosts.txt) -FilePath .\Get-PrintDriverWppReadiness.ps1
```

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `CsvPath` | `string` | — | Write the per-driver classification to this CSV path |
| `JsonPath` | `string` | — | Write the full structured result (drivers, mapped printers, WPP state) to this JSON path |
| `IncludeSafeDrivers` | `switch` | Off | Include WPP-safe (Microsoft inbox) drivers in CSV/JSON output. By default only blocking third-party v3/v4 drivers are written |
| `Quiet` | `switch` | Off | Suppress the human-readable console summary. The single-line JSON summary is still emitted to STDOUT |

## Sample console output

```
  WPP readiness — PC01
  --------------------------------------------------
  WPP currently enforced : False (NotConfigured)
  Total print drivers    : 12
  Result                 : NOT WPP-READY - 3 blocking driver(s)
    v3 (classic)         : 1
    v4                   : 2
    actively in use      : 1

    [v3] Canon GX6000 series  (Canon)  - in use by: Canon GX6000 series
    [v4] HP Photosmart 7520 series Class Driver  (HP)  - not bound to a printer
    [v4] Brother Laser Leg Type1 Class Driver  (Brother)  - not bound to a printer
```

## How classification works

A driver is treated as **WPP-safe** when either:

- its `Manufacturer` is `Microsoft`, **or**
- its `Name` matches the inbox allowlist in `$Script:WppSafeDriverNames` (e.g. *Microsoft IPP Class Driver*, *Microsoft Print To PDF*, *Microsoft XPS Document Writer*, *Remote Desktop Easy Print*), or the generic `Microsoft … Class Driver / Document Writer / Print To PDF` pattern.

Everything else is **blocking**, and labeled `v3` / `v4` from `MajorVersion`.

### Caveats — read before mass action

- **The inbox subset is not authoritatively published.** This script flags the real risk signal (third-party drivers) but a third-party driver named `… Class Driver` is still flagged for review (see the HP/Brother examples above). Spot-check the CSV from a pilot ring and, if needed, extend `$Script:WppSafeDriverNames` with drivers you've verified as genuinely IPP/Mopria-based.
- **Presence ≠ usage.** The `InUse` / `UsedByPrinters` columns show which blocking drivers are actually bound to a printer — triage those first.
- **`v3` vs `v4`** is read from `MajorVersion`. Some inbox drivers report `0`; those are evaluated by manufacturer/allowlist only.

## Requirements

- **PowerShell 5.1+** — Windows PowerShell or PowerShell 7+
- **PrintManagement module** — `Get-PrinterDriver` / `Get-Printer` (present on Windows client/server by default)
- **Print Spooler** running — if the Spooler is off or the module is missing, the script reports the error in JSON and exits `1`
- No elevation required for the read-only inventory; run in **system context** under Intune

## File Structure

```
windows/print/
├── Get-PrintDriverWppReadiness.ps1   # Main script (detection + report)
└── README.md                         # This file
```
