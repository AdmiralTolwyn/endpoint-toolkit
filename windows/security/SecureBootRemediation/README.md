# Secure Boot UEFI CA 2023 Remediation

## Overview

Detection and remediation tooling for the **Secure Boot UEFI CA 2023** certificate deployment. Three flavors are included so the same logic can be wired into Intune Proactive Remediations, Ivanti baselines, or used standalone for manual triage.

References:
- [KB5016061 - Secure Boot DB and DBX variable update events](https://support.microsoft.com/en-us/topic/37e47cf8-608b-4a87-8175-bdead630eb69) (event-id semantics)
- [KB5072718 - Sample Secure Boot Inventory Data Collection script](https://support.microsoft.com/en-us/topic/d02971d2-d4b5-42c9-b58a-8527f0ffa30b) (signal coverage)
- [KB5084567 - Sample Secure Boot E2E Automation Guide](https://support.microsoft.com/en-us/topic/f850b329-9a6e-40d1-823a-0925c965b8a0) (`AvailableUpdatesPolicy` guidance, GPO/MDM ownership)

## Scripts

| File | Purpose | Modifies System? | Compliance Gate |
|------|---------|------------------|-----------------|
| [Detect_SecureBootUEFICA2023.ps1](Detect_SecureBootUEFICA2023.ps1) | Intune Proactive Remediation **detect** script | Never | **Strict**: `SecureBoot enabled AND Status=Updated AND Error=0 AND AvailableUpdates=0x4000 AND Event 1808` |
| [Remediate_SecureBootUEFICA2023.ps1](Remediate_SecureBootUEFICA2023.ps1) | Intune Proactive Remediation **remediate** script (two-pronged) | Yes (registry + scheduled task) | n/a (remediation) |
| [SecureBootCertRemediation.ps1](SecureBootCertRemediation.ps1) | Full detection + smart remediation + real-time monitoring (manual/triage) | Only with `-ForceRemediation` | **Strict**: `1808 AND Status=Updated AND Error=0 AND BootloaderSwapped (1799)` |
| [SecureBootCertDetection-Ivanti.ps1](SecureBootCertDetection-Ivanti.ps1) | Pure detection for Ivanti Custom Definition (Status / Reason / Expected / Found contract) | Never | **Legacy**: `Status=Updated AND Error=0` |

All scripts share the same TPM-WMI event-id classification and diagnostic signals; only the compliance gate, output contract, and remediation behavior differ.

### Why so many scripts?

- **`Detect_SecureBootUEFICA2023.ps1` / `Remediate_SecureBootUEFICA2023.ps1`** — the Intune PR pair (v1.1). Logs to the Intune Management Extension log directory in CMTrace format with verbose DEBUG-level instrumentation. The detect script enforces a strict 5-condition gate **plus** hard-blocker short-circuits and reboot-pending suppression to prevent PR retry loops on devices that cannot be remediated from the OS. The remediate script uses two-pronged logic based on the current `AvailableUpdates` bitmask, and aborts cleanly when the Secure-Boot-Update scheduled task is missing/disabled (see below).
- **`SecureBootCertRemediation.ps1`** — the standalone deployment / triage script. Detection-only by default; only writes registry / starts the scheduled task when explicitly invoked with `-ForceRemediation`. Strict gate avoids false-positive "compliant" verdicts caused by stale Event 1808 surviving NVRAM / BIOS resets.
- **`SecureBootCertDetection-Ivanti.ps1`** — the legacy Ivanti baseline detect script. The compliance verdict is intentionally unchanged to avoid baseline / ticket churn; the enhanced diagnostics (1808 / 1799 / 1801 / 1802 / 1803 / latest good / latest bad / FW errors) are surfaced in the `found =` line for triage only.

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

### Intune Proactive Remediation pair

Upload the two scripts to a single PR package:

| Setting | Value |
|---------|-------|
| Detection script | [Detect_SecureBootUEFICA2023.ps1](Detect_SecureBootUEFICA2023.ps1) |
| Remediation script | [Remediate_SecureBootUEFICA2023.ps1](Remediate_SecureBootUEFICA2023.ps1) |
| Run this script using the logged-on credentials | **No** |
| Enforce script signature check | **No** |
| Run script in 64-bit PowerShell | **Yes** |

**Detect** exits `1` (non-compliant) unless **all** of the following are true:

1. `Confirm-SecureBootUEFI` returns `$true`
2. `UEFICA2023Status` = `Updated`
3. `UEFICA2023Error` = `0`
4. `AvailableUpdates` = `0x4000` (terminal complete state)
5. Event ID 1808 from `Microsoft-Windows-TPM-WMI` present in System log

Before the gate is evaluated, the detect script applies several **hard-blocker short-circuits** that return exit `0` ("not actionable" / "pending reboot") instead of `1`. This prevents Intune from re-running the PR forever on devices that cannot be remediated from the OS:

| Condition | Detect exit | STDOUT marker |
|-----------|-------------|---------------|
| `HighConfidenceOptOut = 1` | `0` | `NOT-APPLICABLE: HighConfidenceOptOut=1 (admin-managed exclusion).` |
| Event 1803 present (missing PK-signed KEK) | `0` | `NOT-APPLICABLE: Event 1803 (missing KEK - OEM responsibility).` |
| Event 1802 present (known firmware issue) | `0` | `NOT-APPLICABLE: Event 1802 (known firmware issue [KI_<n>]).` |
| Event 1800 present, Event 1808 not yet present | `0` | `PENDING-REBOOT: Update staged, awaiting reboot. Gate would fail on: ...` |

In addition, the detect script emits **operational warnings** (logged but do **not** change the verdict) when:

- The `\Microsoft\Windows\PI\Secure-Boot-Update` scheduled task is missing or in a non-`Ready`/`Running` state — remediation will be ineffective.
- `CanAttemptUpdateAfter` (REG_BINARY/REG_QWORD `FILETIME` under `Servicing\DeviceAttributes`) is in the future — firmware throttle active; triggering the task is a no-op until it elapses.
- `AvailableUpdatesPolicy` disagrees with `AvailableUpdates` — GPO or Intune is the source of truth; direct registry writes may be reverted on the next policy refresh.

Diagnostic context surfaced in every detect run (logged at INFO level): `UEFICA2023ErrorEvent`, `MicrosoftUpdateManagedOptIn`, per-ID event counts (`1808`/`1803`/`1802`/`1801`/`1800`/`1795`/`1796`), parsed `BucketId` / `BucketConfidenceLevel` / `SkipReason`, and captured error codes from the latest `1795` / `1796`.

**Remediate** also short-circuits to exit `0` on the same hard-blocker conditions (`HighConfidenceOptOut=1`, Event 1803, Event 1802, Event 1800 without 1808) before touching the registry or the task. When none of those apply, it branches on the current `AvailableUpdates` value:

| Current value | Branch | Action |
|---------------|--------|--------|
| missing or `0` | **A. Initial arm** | Set `AvailableUpdates = 0x5944`, then trigger `\Microsoft\Windows\PI\Secure-Boot-Update` |
| non-zero, not `0x4000` (e.g. `0x5944`, `0x4100`, `0x4104`) | **B. Resume** | Trigger the scheduled task **without** modifying the registry (preserve in-flight state) |
| `0x4000` | **C. No-op** | Already at the terminal complete state |

The remediate script also performs a pre-flight check on the `Secure-Boot-Update` scheduled task **before** writing the registry. If the task is missing or in a non-`Ready`/`Running` state it exits `1` with a clear `FAIL:` message instead of arming the device only to discover the task cannot run.

> **About `AvailableUpdatesPolicy`** — this value is reserved for Group Policy and Intune. Per Microsoft guidance it must **only be written by GPO/MDM**, never by a remediation script. Both scripts read `AvailableUpdatesPolicy` for diagnostic context (and warn when it disagrees with `AvailableUpdates`) but **never write it**. The remediate script writes only `AvailableUpdates` (the volatile, OS-consumed value).

Both scripts log in CMTrace format to:

```text
%ProgramData%\Microsoft\IntuneManagementExtension\Logs\PR_SecureBootUEFICA2023.log
```

Logging is highly verbose: every function entry/exit, registry read/write, event-log query, and decision branch is recorded with severity tags (`[DEBUG]`, `[INFO]`, `[WARN]`, `[ERROR]`, `[SUCCESS]`, `[STEP]`). Logs auto-rotate at 250 KB and the archive copy is NTFS-compressed.

### `SecureBootCertRemediation.ps1` (detection-only by default)

```powershell
# Detection-only: prints structured report, never writes registry or starts the task
.\SecureBootCertRemediation.ps1

# Active remediation: applies registry value + triggers Secure-Boot-Update scheduled task
.\SecureBootCertRemediation.ps1 -ForceRemediation
```

> Without `-ForceRemediation` the script is read-only. This is the safe default for scheduled scans, Intune Proactive Remediation *detect* scripts, and one-off triage.

### `SecureBootCertDetection-Ivanti.ps1` (pure detection)

```powershell
.\SecureBootCertDetection-Ivanti.ps1
```

Emits exactly four lines on stdout (Ivanti contract):

```text
detected = true|false
reason   = <single sentence>
expected = Status: Updated | Error: 0
found    = Status: <s> | Error: <e> | Confidence: <c> | Capable: <cap> | Event1808: <bool> | BootloaderSwapped: <bool> | LatestGood: <id> | ...
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

> Authoritative source: [KB5016061 - Secure Boot DB and DBX variable update events](https://support.microsoft.com/en-us/topic/secure-boot-db-and-dbx-variable-update-events-37e47cf8-608b-4a87-8175-bdead630eb69)

All events are logged under the `Microsoft-Windows-TPM-WMI` provider in the **System** log. Event 1799 is additionally checked in `Microsoft-Windows-TPM-WMI/Operational` by the standalone scripts.

In every event message below, `<event type>` (where applicable) resolves to one of:
`DB`, `DBX`, `SBAT`, `Policy Update (SKU)`, `Windows UEFI CA 2023 (DB)`,
`Option ROM CA 2023 (DB)`, `3P UEFI CA 2023 (DB)`, `KEK 2023`, `DBX SVN`, or
`Revoke UEFI CA 2011 (DBX)`.

### Generic Secure Boot events (1032 - 1800)

These events apply to all devices and describe the outcome of DB, DBX, KEK, SBAT, and boot manager updates.

#### Success / informational

| Event ID | Level | What it means |
|----------|-------|---------------|
| **1034** | Information | Standard DBX revocations were applied to firmware (`Secure Boot Dbx update applied successfully`). Confirms a DBX (Forbidden Signatures Database) update was committed. |
| **1036** | Information | A DB variable update was applied (`Secure Boot Db update applied successfully`). Used when adding trusted certificates to the Allowed Signatures Database. |
| **1037** | Information | The Microsoft Windows Production PCA 2011 certificate was added to the DBX (`Secure Boot Dbx update to revoke Microsoft Windows Production PCA 2011 is applied successfully`). After this event boot applications signed by the 2011 PCA are no longer trusted -- this includes recovery media, PXE boot apps, and any third-party boot loader signed with that certificate. |
| **1043** | Information | The Microsoft Corporation **KEK CA 2023** certificate was added to the KEK variable. Required so the device can keep receiving DB/DBX updates after the existing **KEK CA 2011** expires in 2026. |
| **1044** | Information | The **Microsoft Option ROM UEFI CA 2023** certificate was added to the DB variable. Required for Option ROM continuity past the 2011 UEFI CA expiry. |
| **1045** | Information | The **Microsoft UEFI CA 2023** certificate was added to the DB variable. This is the new third-party / general UEFI signing root that replaces the **Microsoft UEFI CA 2011** in 2026. |
| **1799** | Information | **Boot Manager swap.** A new boot manager signed by the **Windows UEFI CA 2023** has been installed on the EFI partition. This is the strongest evidence that the OS half of the rollout is complete -- the device is now actually *using* the new bootmgr, not just trusting it. |
| **1800** | Warning | **Reboot required.** Applying the Secure Boot update in the current boot cycle would conflict with a recent change (boot manager update, VBS-related variable update, etc.). A restart clears the condition; not a blocker. |

#### Errors / blockers

| Event ID | Level | What it means | Resolution |
|----------|-------|---------------|------------|
| **1032** | Error | Update was skipped because the BitLocker configuration would force the device into recovery if the update were applied. | Suspend BitLocker for 2 reboots: `Manage-bde -Protectors -Disable %systemdrive% -RebootCount 2`, restart twice, then re-enable: `Manage-bde -Protectors -enable %systemdrive%`. |
| **1033** | Error | A potentially revoked / vulnerable boot manager was found on the EFI partition (event data includes `BootMgr` = path to the file). The DBX update is deferred each boot until the vulnerable module is replaced. | Update the offending bootloader (usually a third-party OS, recovery agent, or hypervisor loader). Microsoft re-evaluates on every boot. |
| **1796** | Error | A non-specific error occurred during `<event type>` update; raw firmware error code is included. Windows will retry on the next restart. | Use the error code as a diagnostic hint. Often clears on its own after a reboot or a firmware update. |
| **1797** | Error | DBX update to revoke the Windows Production PCA 2011 was intentionally failed because the **Windows UEFI CA 2023** is **not yet present in DB**. Doing so would leave the device unable to verify Microsoft-signed boot apps. | Wait for / force the DB update (Events 1045 / 1043) to complete first. |
| **1798** | Error | DBX update to revoke the Windows Production PCA 2011 was intentionally failed because the **default boot manager is not yet signed with the Windows UEFI CA 2023**. | Wait for the boot manager swap (Event 1799) to occur, then the DBX update can proceed. |

### Device-specific events (1795, 1801 - 1808)

These events carry per-device telemetry fields used by Microsoft to bucket devices and decide whether the high-confidence path can be enabled. Every event below includes:

| Field | Meaning |
|-------|---------|
| `DeviceAttributes` | Characteristics of the device (OEM, model, firmware ver, etc.). Inputs to the bucket hash. |
| `BucketId` | Stable hash identifying a group of equivalent devices. Changes when device attributes change (e.g. after a firmware update). |
| `BucketConfidenceLevel` | Microsoft's assessment of how confidently this bucket can accept the update. See [BucketConfidenceLevel values](#bucketconfidencelevel-values) below. |
| `UpdateType` | `0` or `22852` (`0x5944`). `0x5944` = High Confidence update path. |

#### Success / informational

| Event ID | Level | What it means |
|----------|-------|---------------|
| **1808** | Information | **Device is fully updated.** All required new Secure Boot certificates have been applied to firmware **and** the boot manager has been replaced with the version signed by **Windows UEFI CA 2023**. Per Microsoft (Apr 2026 KB note): the `BucketConfidenceLevel` reflects data coverage for similar devices and does *not* indicate further action is required on this device. |

#### Errors / blockers

| Event ID | Level | What it means | Resolution |
|----------|-------|---------------|------------|
| **1795** | Error | **Firmware rejected** the DB / DBX / KEK update with the included `<firmware error code>`. Windows will retry on the next restart. | Contact the OEM for a firmware update. The error code helps the OEM identify the failure mode. |
| **1801** | Error | **Status / assessment event.** New Secure Boot certificates have been published to the device but have *not* yet been applied to firmware. Includes the full bucket telemetry so administrators can correlate which devices still need updating. Note: although the KB classifies this as Error level, in practice it can also fire under "Under Observation" or "Temporarily Paused" conditions, so the standalone scripts treat it as a *warning* signal rather than a hard blocker. | Review [aka.ms/SecureBootStatus](https://aka.ms/SecureBootStatus) and ensure the device has applied the latest cumulative update; reboot. |
| **1802** | Error | **Update intentionally blocked** because the device matches a known-issue (KI) firmware/hardware condition that would cause failure or damage. The `SkipReason` field carries a `KI_<id>` value. | Look up the KI ID at [aka.ms/SecureBootKnownIssues](https://go.microsoft.com/fwlink/?linkid=2339472). Usually requires an OEM firmware fix. |
| **1803** | Error | **Missing PK-signed KEK.** The KEK can only be updated when the new KEK is signed by the device's Platform Key (PK). No PK-signed KEK for this device's PK was found in the cumulative update, so the KEK update cannot proceed. | Contact the OEM and ask for the PK-signed Microsoft KEK 2023 to be supplied to Microsoft for inclusion in a future cumulative update. |

> **Why no Event 1042?** The standalone `SecureBootCertRemediation.ps1` script's `$GoodEventIDs` list contains `1042` for historical reasons (early-preview behavior). It is not documented in the current KB5016061 and is harmless if absent -- the script does not require it for any compliance decision.

### BucketConfidenceLevel values

The `BucketConfidenceLevel` field appears on Events **1795**, **1801**, **1802**, **1803**, and **1808**. It tells you why the device is (or is not) eligible for the automated update path.

For a list of known High Confidence bucket hashes, see [microsoft/secureboot_objects - HighConfidenceBuckets](https://github.com/microsoft/secureboot_objects/tree/main/HighConfidenceBuckets).

| Confidence Level | Meaning | Typical operator action |
|------------------|---------|--------------------------|
| **High Confidence** | Devices in this bucket have demonstrated through observed Microsoft data that they can successfully update firmware using the new Secure Boot certificates. `UpdateType` = `0x5944`. | None -- the rollout will proceed automatically. |
| **Temporarily Paused** | Bucket is affected by a known issue. Updates are paused while Microsoft and partners work toward a supported resolution. Often requires a firmware update. | Look for a paired Event 1802 (`SkipReason: KI_xxxx`) for the specific KI ID. Apply the OEM firmware fix when available. |
| **Not Supported - Known Limitation** | Bucket cannot use the automated path because of a permanent hardware/firmware limitation. | Manual remediation only. The device may need OEM-provided tooling or hardware replacement. |
| **Under Observation - More Data Needed** | Not blocked, but Microsoft does not yet have enough telemetry to classify the bucket as high confidence. Updates may be deferred. | No action -- the bucket will reclassify automatically as more data arrives. Devices can still install the update manually if needed. |
| **No Data Observed - Action Required** | Microsoft has not seen this device in Secure Boot update telemetry. The automated path cannot be evaluated. | Follow [aka.ms/SecureBootStatus](https://aka.ms/SecureBootStatus). Usually requires manual deployment via the registry trigger pattern in `Remediate_SecureBootUEFICA2023.ps1`. |

## Key Registry Paths

- `HKLM\SYSTEM\CurrentControlSet\Control\SecureBoot\AvailableUpdates`
- `HKLM\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing\UEFICA2023Status`
- `HKLM\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing\UEFICA2023Error`
- `HKLM\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing\WindowsUEFICA2023Capable`

> **Known Issue:** `WindowsUEFICA2023Capable` is incorrectly reported as `0` on **Windows Server 2019** regardless of actual capability. Do not rely on this value alone for Server 2019 compliance decisions.

## Compliance Logic

The scripts deliberately use different compliance gates depending on their target platform.

### `Detect_SecureBootUEFICA2023.ps1` -- Intune PR strict gate

```
Compliant = (Confirm-SecureBootUEFI = $true)
        AND (UEFICA2023Status = "Updated")
        AND (UEFICA2023Error  = 0)
        AND (AvailableUpdates = 0x4000)
        AND (Event 1808 present)
```

Intune-friendly gate: relies on the deterministic `AvailableUpdates = 0x4000` terminal state plus the historical Event 1808. Does **not** require Event 1799 because the bootloader-swap event can age out of the System log on long-lived devices.

### `SecureBootCertRemediation.ps1` -- strict gate (v2.2+)

```
Compliant = (Event 1808 present)
        AND (UEFICA2023Status = "Updated")
        AND (UEFICA2023Error  = 0)
        AND (Event 1799 present -- BootloaderSwapped)
```

All four signals are required. The `Error=0` and `BootloaderSwapped` requirements were added in v2.2 to defend against stale Event 1808 entries surviving NVRAM / BIOS resets, which previously caused false-positive "system is fully updated" verdicts on regressed devices.

### `SecureBootCertDetection-Ivanti.ps1` -- legacy gate

```
Compliant = (UEFICA2023Status = "Updated") AND (UEFICA2023Error = 0)
```

Intentionally preserved to avoid churning existing Ivanti baselines and tickets. Event 1808 / 1799 are reported in the `found =` line as informational signals only and do **not** influence the verdict.

> Event 1808 is now also reliably generated on **Windows Server 2025**.

## Output

### `Detect_SecureBootUEFICA2023.ps1`

Single-line STDOUT summary suitable for the Intune PR detection column. Possible outputs:

```text
COMPLIANT: Secure Boot UEFI CA 2023 update fully applied.
NON-COMPLIANT: UEFICA2023Status missing | AvailableUpdates=0x5944 | Event 1808 missing
NOT-APPLICABLE: HighConfidenceOptOut=1 (admin-managed exclusion).
NOT-APPLICABLE: Event 1803 (missing KEK - OEM responsibility).
NOT-APPLICABLE: Event 1802 (known firmware issue [KI_12345]).
PENDING-REBOOT: Update staged, awaiting reboot. Gate would fail on: AvailableUpdates=0x4100 | Event 1808 missing
PREREQ: Not running in 64-bit PowerShell.
```

| Exit code | Meaning | Triggers remediation? |
|-----------|---------|----------------------|
| `0` | `COMPLIANT`, `NOT-APPLICABLE`, or `PENDING-REBOOT` | No |
| `1` | `NON-COMPLIANT` or `PREREQ` failure | Yes |

Full per-step trace is written to the CMTrace log file.

### `Remediate_SecureBootUEFICA2023.ps1`

Single-line STDOUT summary describing the action taken. Possible outputs:

```text
REMEDIATED: AvailableUpdates pre=0x0 post=0x5944. Reboot required to finalize.
NO-OP: Already at compliant terminal state (0x4000).
NOT-APPLICABLE: HighConfidenceOptOut=1 (admin-managed exclusion).
NOT-APPLICABLE: Event 1803 (missing KEK - OEM responsibility).
NOT-APPLICABLE: Event 1802 (known firmware issue [KI_12345]).
PENDING-REBOOT: Update staged, awaiting reboot.
ABORT: Secure Boot disabled.
FAIL: Scheduled task '\Microsoft\Windows\PI\Secure-Boot-Update' is missing.
FAIL: Scheduled task is 'Disabled' (not Ready). Re-enable before remediation.
FAIL: Could not start Secure-Boot-Update scheduled task.
FAIL: Registry write error - <message>
```

| Exit code | Meaning |
|-----------|---------|
| `0` | `REMEDIATED`, `NO-OP`, `NOT-APPLICABLE`, or `PENDING-REBOOT` |
| `1` | `PREREQ`, `ABORT`, or `FAIL` (registry / scheduled task / Secure Boot disabled) |

Full per-step trace is written to the same CMTrace log file as the detect script.

### `SecureBootCertRemediation.ps1`

Produces a structured, color-coded detection report (registry values, event sweep, AvailableUpdates decode) followed by an Ivanti-compatible detection summary block:

- `Detected` -- `true` (non-compliant) or `false` (compliant)
- `Reason`   -- human-readable status explanation
- `Expected` / `Found` -- state comparison strings

When `-ForceRemediation` is supplied, the script additionally logs the registry write, scheduled-task trigger, and a 30-second monitoring loop tracking `AvailableUpdates` changes.

### `SecureBootCertDetection-Ivanti.ps1`

Four `Write-Host` lines, no banner, no color, no extra output -- safe to consume verbatim from an Ivanti Custom Definition or any detect channel that parses `key = value` pairs.

## Requirements

- Windows 10/11 or Windows Server 2016+
- PowerShell 5.1+
- Administrator privileges (registry writes, scheduled task execution)
- Scheduled task `\Microsoft\Windows\PI\Secure-Boot-Update` must exist (installed via Windows Update)
