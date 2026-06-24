# Hardware Speculation Vulnerability Detection

## Overview

Detects CPU hardware vulnerabilities (MDS and SSB) using low-level `ntdll.dll` queries against the Windows speculation control information interface. Designed for use as an **Ivanti Custom Vulnerability Detection** script.

Based on the official [Microsoft SpeculationControl PowerShell module](https://www.powershellgallery.com/packages/SpeculationControl).

## Script

| File | Purpose |
|------|---------|
| `Get-HardwareSpeculationStatus.ps1` | Detection only. Queries CPU speculation control flags and reports vulnerability status |
| `Invoke-HardwareSpeculationRemediation.ps1` | Detection **and** remediation. Re-detects state, reports per-QID compliance, then (with `-ForceRemediation`) writes the `FeatureSettingsOverride`/`Mask` (and Hyper-V) registry values that clear Qualys QID 91462 (SSB) and QID 91537 (MDS/TAA) |

## Vulnerabilities Detected

| Advisory | Qualys QID | Name | Notes |
|----------|-----------|------|-------|
| [ADV180012](https://msrc.microsoft.com/update-guide/vulnerability/ADV180012) | 91462 | SSB (Speculative Store Bypass) | Spectre Variant 4 / CVE-2018-3639. All CPUs. |
| [ADV190013](https://msrc.microsoft.com/update-guide/vulnerability/ADV190013) | 91537 | MDS / TAA (Microarchitectural Data Sampling, Intel TSX) | Intel-specific; server-focused (clients enabled by default). |

> The detection-only script (`Get-HardwareSpeculationStatus.ps1`) reports the **runtime hardware** state via `ntdll` (`MdsHardwareProtected`, `SsbdRequired`; AMD/ARM treated as protected). The Qualys QIDs above check the **registry configuration** instead — see [Scanner behavior & QID quirks](#scanner-behavior--qid-quirks).

## Scanner behavior & QID quirks

These two Qualys QIDs are **registry-configuration checks**, not runtime hardware checks. This causes two non-obvious issues that the remediation script is specifically designed to handle:

1. **Flagged even when the CPU is protected.** Qualys never reads the live `ntdll` `SsbdRequired`/`MdsHardwareProtected` flags. A device can be technically protected (or report "mitigation not required") at runtime yet still be flagged purely because `FeatureSettingsOverride` / `FeatureSettingsOverrideMask` are missing or set to a non-accepted value. The remediation therefore gates on **registry-config compliance**, not the runtime flag.

2. **The two QIDs accept conflicting value sets.** QID 91462 (SSB) accepts a *broad* set of values, but QID 91537 (MDS) does an **exact match** on `72` or `8264`. The **only** values that clear **both** are `72` (`0x48`) and `8264` (`0x2048`):

   | `FeatureSettingsOverride` | QID 91462 (SSB) | QID 91537 (MDS) |
   |---|:---:|:---:|
   | `8` (`0x8`) | ✅ | ❌ |
   | **`72`** (`0x48`) | ✅ | ✅ |
   | **`8264`** (`0x2048`) | ✅ | ✅ |
   | `8388616` (`0x800008`, +BHI) | ✅ | ❌ |
   | `8396872` (`0x802048`, +BHI) | ✅ | ❌ |

   Because QID 91537 requires an *exact* value, it is **mutually exclusive** with preserving extra mitigation bits set by other advisories (e.g. BHI `0x800000` for CVE-2022-0001). You cannot satisfy QID 91537 *and* keep BHI in the same value. The script defaults to **additive (do-no-harm)** and offers `-ForceExactValue` for the trade-off — see [Additive vs exact](#additive-vs-exact-avoiding-regression).

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

## Remediation

`Invoke-HardwareSpeculationRemediation.ps1` automates the registry mitigation. It is
**detection-only by default** and only writes the registry when `-ForceRemediation` is
supplied, so it is safe to run as a compliance detect script.

```powershell
# Detection only - reports state, makes NO changes
.\Invoke-HardwareSpeculationRemediation.ps1

# Apply the mitigation (writes registry; reboot required)
.\Invoke-HardwareSpeculationRemediation.ps1 -ForceRemediation

# Apply the HT-disabled value (8264) for full L1TF/MDS coverage (disables Hyper-Threading)
.\Invoke-HardwareSpeculationRemediation.ps1 -ForceRemediation -DisableHyperThreading

# Force EXACTLY 72/8264 (drops any other mitigation bits - see "Additive vs exact" below)
.\Invoke-HardwareSpeculationRemediation.ps1 -ForceRemediation -ForceExactValue
```

The script:

- Writes the single `FeatureSettingsOverride` value that satisfies **both** scanner checks (SSB QID 91462 and MDS QID 91537): `72` (`0x48`) by default, or `8264` (`0x2048`) with `-DisableHyperThreading`, plus `FeatureSettingsOverrideMask = 3`.
- Both values enable the full documented mitigation set (MDS, TAA, Spectre, Meltdown, SSBD, L1TF) per KB4073119, with the Spectre V2 (`0x1`) / Meltdown (`0x2`) *disable* bits left clear so those mitigations stay enabled.
- On **Hyper-V hosts** (Intel), also sets `MinVmVersionForCpuBasedMitigations = "1.0"` under `...\CurrentVersion\Virtualization` (required by QID 91537).
- Is idempotent — skips the write when the registry already satisfies the applicable checks.

> **Why 72/8264 and not the minimal SSBD value 8?** Because `8` clears the SSB QID but **not** the MDS QID. Only `72`/`8264` satisfy both — see [Scanner behavior & QID quirks](#scanner-behavior--qid-quirks).

### Additive vs exact (avoiding regression)

`FeatureSettingsOverride` is a single additive bitmask — each bit is a separate mitigation toggle. To avoid regressing mitigations set by other advisories, the script is **additive (do-no-harm) by default**:

- It **preserves** every existing higher-order bit (e.g. BHI `0x800000` for CVE-2022-0001, MMIO bits), **forces** the inverted-logic Spectre V2 (`0x1`) / Meltdown (`0x2`) *disable* bits **clear**, then **sets** SSBD (`0x8`) + MDS/TAA/L1TF (`0x40`) (+ HT-disable `0x2000` with `-DisableHyperThreading`).
- On a device with the keys **missing** (the common scan finding), additive of `0` produces **exactly** `72` (or `8264`) — so it clears both QIDs with **zero** regression.
- On a device that already carries extra bits (e.g. BHI), the additive result is a **superset** (e.g. `0x800048`) — stronger, never weaker. **Caveat:** QID 91537 wants an *exact* `72`/`8264`, so the MDS QID may stay flagged in that case. The script warns when this happens.

`-ForceExactValue` overwrites to **exactly** `72`/`8264` + Mask `3` to force the QID 91537 exact-match to clear. This **drops** any extra mitigation bits (e.g. BHI/MMIO), so it is **opt-in only** and the script warns loudly before doing so. There is no single value that both preserves BHI **and** satisfies QID 91537's exact-match — they are mutually exclusive.

### SSB (ADV180012) - Registry Mitigation (Qualys QID 91462)

Enables Speculative Store Bypass Disable (SSBD). A reboot is required after applying. Manual equivalent:

```cmd
reg add "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management" /v FeatureSettingsOverride /t REG_DWORD /d 72 /f
reg add "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management" /v FeatureSettingsOverrideMask /t REG_DWORD /d 3 /f
```

> **Scanner note (Qualys QID 91462):** Qualys flags ADV180012 purely on the **registry config** — it checks that `FeatureSettingsOverride` exists with an accepted value **and** `FeatureSettingsOverrideMask = 3`. It does **not** read the live CPU/`ntdll` `SsbdRequired` flag, so a device can report "not required" at runtime yet still be flagged on the missing keys. Accepted values (all contain SSBD bit `0x8`, no `0x1`/`0x2` disable bits, Mask `3`): `8`, `72`, `8264`, `8388616`, `8388680`, `8396872`.

#### FeatureSettingsOverride Bitmask

| Bit | Hex | When SET (1) | When CLEAR (0) |
|-----|-----|--------------|----------------|
| 0 | `0x1` | **Disable** Spectre V2 (CVE-2017-5715) | Enable Spectre V2 |
| 1 | `0x2` | **Disable** Meltdown (CVE-2017-5754) | Enable Meltdown |
| 3 | `0x8` | **Enable** SSBD / SSB (CVE-2018-3639) | Disable SSB mitigation |
| 6 | `0x40` | **Enable** MDS / TAA / L1TF mitigation override | (off) |
| 13 | `0x2000` | **Disable** Hyper-Threading (full L1TF/MDS) | Hyper-Threading enabled |

> **Note:** Bits 0-1 use inverted logic (set = disable). Bits 3 and 6 use normal logic (set = enable). `FeatureSettingsOverrideMask = 3` tells the OS to respect the Spectre/Meltdown override bits.

**Common combined values:**

| Value | Hex | Effect |
|-------|-----|--------|
| `8` | `0x8` | SSBD only (clears QID 91462, **not** 91537) |
| `72` | `0x48` | Full mitigation set, Hyper-Threading **kept** (clears both QIDs) |
| `8264` | `0x2048` | Full mitigation set, Hyper-Threading **disabled** (clears both QIDs) |
| `3` | `0x3` | Spectre V2 + Meltdown **disabled** (do not use) |

### MDS / TAA (ADV190013) - Registry Mitigation (Qualys QID 91537)

The **hardware** MDS mitigation is enabled automatically once you apply (1) the latest cumulative
Windows update and (2) the CPU microcode update (BIOS/firmware or Windows Update), then reboot.

However, **Qualys QID 91537 is a registry-config check** (Intel, server-focused). It flags the device
unless the registry is set to one of the documented combined values:

```cmd
:: Without disabling Hyper-Threading
reg add "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management" /v FeatureSettingsOverride /t REG_DWORD /d 72 /f
reg add "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management" /v FeatureSettingsOverrideMask /t REG_DWORD /d 3 /f

:: With Hyper-Threading disabled (complete L1TF/MDS)
reg add "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management" /v FeatureSettingsOverride /t REG_DWORD /d 8264 /f
reg add "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management" /v FeatureSettingsOverrideMask /t REG_DWORD /d 3 /f
```

On **Hyper-V hosts**, also add:

```cmd
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Virtualization" /v MinVmVersionForCpuBasedMitigations /t REG_SZ /d "1.0" /f
```

> **Scanner note (Qualys QID 91537):** Exact match — `FeatureSettingsOverride` must be **`72` or `8264`** with `FeatureSettingsOverrideMask = 3`. This QID is Intel-specific (Intel TSX/TAA) and primarily targets servers; clients have the patches enabled by default. On Hyper-V hosts, fully shut down all VMs before rebooting so the firmware mitigation is applied to the host first.

## References

- [KB4073119 - Windows client guidance for IT Pros to protect against silicon-based microarchitectural and speculative execution side-channel vulnerabilities](https://support.microsoft.com/en-us/topic/kb4073119-windows-client-guidance-for-it-pros-to-protect-against-silicon-based-microarchitectural-and-speculative-execution-side-channel-vulnerabilities-35820a8a-ae13-1299-88cc-357f104f5b11)
- [ADV180012 - Microsoft Guidance for Speculative Store Bypass](https://msrc.microsoft.com/update-guide/vulnerability/ADV180012) (Qualys QID 91462)
- [ADV190013 - Microsoft Guidance to mitigate Microarchitectural Data Sampling vulnerabilities](https://msrc.microsoft.com/update-guide/vulnerability/ADV190013) (Qualys QID 91537)
- [Microsoft SpeculationControl PowerShell module](https://www.powershellgallery.com/packages/SpeculationControl)
