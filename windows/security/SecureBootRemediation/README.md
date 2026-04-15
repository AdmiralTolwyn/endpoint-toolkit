# Secure Boot UEFI CA 2023 Remediation

## Overview

Unified detection and remediation script for the **Secure Boot UEFI CA 2023** certificate deployment. Combines comprehensive state analysis with smart remediation logic and real-time monitoring.

Reference: [KB5016061 - Secure Boot DB and DBX variable update events](https://support.microsoft.com/en-us/topic/37e47cf8-608b-4a87-8175-bdead630eb69)

## Script

| File | Purpose |
|------|---------|
| `SecureBootCertRemediation.ps1` | Detects Secure Boot update state, applies remediation, and monitors progress |

## What It Does

### 1. Detection
- Queries all Secure Boot Playbook registry keys (`AvailableUpdates`, `AvailableUpdatesPolicy`, `HighConfidenceOptOut`, `MicrosoftUpdateManagedOptIn`, Servicing keys)
- Full event sweep from `System` log (`Microsoft-Windows-TPM-WMI` provider):
  - **Good events**: 1034, 1036, 1037, 1042-1045, 1799, 1800, 1808
  - **Warning/Bad events**: 1032, 1033, 1795-1798, 1801, 1802, 1803
  - **Note**: Event 1801 is a status/assessment event that fires when the update has NOT yet completed or when issues are detected - it is not a success indicator
- Event 1799 dual-log check (System + TPM-WMI/Operational)
- Reports confidence level, BucketId, error codes, and structured debug output

### 2. Remediation (Smart Logic)
- **Initial Run**: If `AvailableUpdates` is `0`, sets `0x5944` and triggers the scheduled task
- **Post-Reboot**: If `AvailableUpdates` is `0x4100`, triggers the task to finalize to `0x4000`
- Aborts if blocking issues (known firmware issues, missing KEK) are detected

### 3. Monitoring
- Loops for 30 seconds tracking `AvailableUpdates` registry changes in real-time

## Usage

```powershell
# Standard detection + smart remediation
.\SecureBootCertRemediation.ps1

# Force remediation even if state appears valid
.\SecureBootCertRemediation.ps1 -ForceRemediation
```

## AvailableUpdates State Machine

| Value | Meaning | Action |
|-------|---------|--------|
| `0x0` | Not started | Script sets `0x5944` and triggers task |
| `0x5944` | All updates queued | Wait for task processing |
| `0x4100` | Boot Manager staged, pending reboot | Reboot, then script triggers finalization |
| `0x4104` | KEK update pending | Reboot required |
| `0x4000` | Complete (conditional on 2011 CA) | No action needed |

## Event ID Reference

> Based on [KB5016061 - Secure Boot DB and DBX variable update events](https://support.microsoft.com/en-us/topic/37e47cf8-608b-4a87-8175-bdead630eb69)

All events are logged under the `Microsoft-Windows-TPM-WMI` provider in the **System** log. Event 1799 is additionally checked in `Microsoft-Windows-TPM-WMI/Operational`.

### Generic Secure Boot Events

These events apply to all devices and describe outcomes of DB, DBX, KEK, and boot manager updates.

| Event ID | Type | Meaning | Script Usage |
|----------|------|---------|--------------|
| 1032 | Error | Secure Boot variable update failed (generic failure) | Sets `Blocking` flag |
| 1033 | Error | Secure Boot policy validation failed | Sets `Blocking` flag |
| 1034 | Success | Secure Boot DB variable updated successfully | Counted as good event |
| 1036 | Success | Secure Boot DBX (revocation list) variable updated successfully | Counted as good event |
| 1037 | Success | Key Exchange Key (KEK) updated successfully | Counted as good event |
| 1042 | Success | Secure Boot content update staged for next boot | Counted as good event |
| 1043 | Success | Secure Boot DBX entry applied | Counted as good event |
| 1044 | Success | Secure Boot DB entry removed | Counted as good event |
| 1045 | Success | Secure Boot DBX entry removed | Counted as good event |
| 1796 | Error | Error code logged during Secure Boot update processing | Extracts error code for diagnostics |
| 1797 | Error | Secure Boot update validation failed (content or signature) | Sets `Blocking` flag |
| 1798 | Error | Secure Boot update staging failed | Sets `Blocking` flag |
| 1799 | Success | Boot manager updated to version signed with UEFI CA 2023 | Dual-log check (System + TPM-WMI/Operational); sets `BootloaderSwapped` |
| 1800 | Info | Reboot required to complete Secure Boot update | Sets `RebootPending` flag; not a blocker |

### Device-Specific Events

These events include telemetry fields specific to the device: `DeviceAttributes`, `BucketID`, `BucketConfidenceLevel`, and `UpdateType`.

| Event ID | Type | Meaning | Script Usage |
|----------|------|---------|--------------|
| 1795 | Error | Firmware rejected the Secure Boot update | Extracts error code for diagnostics |
| 1801 | Warning | Device assessment report - fires when update has NOT completed or issues detected | Extracts `BucketConfidenceLevel`, `BucketId`, `UpdateType`, `DeviceAttributes`, status summary; classified as warning (not a success indicator) |
| 1802 | Warning | Update skipped due to known firmware issue | Extracts `SkipReason` (e.g., `KI_xxxx`); blocks remediation |
| 1803 | Error | Missing KEK - OEM must supply PK-signed KEK update | Sets `MissingKEK` flag; blocks remediation |
| 1808 | Success | Secure Boot certificate update completed successfully | Sets `Success` flag; overrides status summary |

### BucketConfidenceLevel Values (Event 1801)

The `BucketConfidenceLevel` field in Event 1801 indicates how confidently the device can accept the update.

For a list of known High Confidence bucket hashes, see [microsoft/secureboot_objects - HighConfidenceBuckets](https://github.com/microsoft/secureboot_objects/tree/main/HighConfidenceBuckets).

| Confidence Level | Description |
|------------------|-------------|
| **High Confidence** | Device has demonstrated through observed data that it can successfully update firmware using the new Secure Boot certificates. `UpdateType` = `0x5944`. |
| **Temporarily Paused** | Device is affected by a known issue. Updates are paused while Microsoft and partners work toward a resolution (may require firmware update). Check Event 1802 for details. |
| **Not Supported - Known Limitation** | Device does not support the automated Secure Boot certificate update path due to hardware or firmware limitations. No automatic resolution available. |
| **Under Observation - More Data Needed** | Device is not blocked, but there is not yet enough data to classify it as high confidence. Updates may be deferred. |
| **No Data Observed - Action Required** | Microsoft has not observed this device in Secure Boot update data. Automatic certificate updates cannot be evaluated. Administrator action is required. See [aka.ms/SecureBootStatus](https://aka.ms/SecureBootStatus). |

## Key Registry Paths

- `HKLM\SYSTEM\CurrentControlSet\Control\SecureBoot\AvailableUpdates`
- `HKLM\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing\UEFICA2023Status`
- `HKLM\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing\UEFICA2023Error`
- `HKLM\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing\WindowsUEFICA2023Capable`

> **Known Issue:** `WindowsUEFICA2023Capable` is incorrectly reported as `0` on **Windows Server 2019** regardless of actual capability. Do not rely on this value alone for Server 2019 compliance decisions.

## Compliance Logic

Per PG recommendation, the script determines compliance using:

```
Compliant = (Event 1808 present) AND (UEFICA2023Status = "Updated")
```

Event 1808 confirms the Secure Boot certificate update completed successfully. This event is now also reliably generated on **Windows Server 2025**.

## Output

Produces a structured detection report with color-coded fields and an Ivanti-compatible detection summary:

- `Detected` — `true` (non-compliant) or `false` (compliant)
- `Reason` — Human-readable status explanation
- `Expected` / `Found` — State comparison strings

## Requirements

- Windows 10/11 or Windows Server 2016+
- PowerShell 5.1+
- Administrator privileges (registry writes, scheduled task execution)
- Scheduled task `\Microsoft\Windows\PI\Secure-Boot-Update` must exist (installed via Windows Update)
