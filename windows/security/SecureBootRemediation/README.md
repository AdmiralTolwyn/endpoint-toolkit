# Secure Boot UEFI CA 2023 Remediation

## Overview

Tooling for the **Secure Boot UEFI CA 2023** certificate deployment, in two families that answer different questions.

- **Drive the update** — get a device onto the new certificates. Intune Proactive Remediation pair, a standalone triage script, and an Ivanti detect script. These write `AvailableUpdates` and trigger the Secure-Boot-Update scheduled task.
- **Measure the fleet** — find out what the estate is actually doing, including the devices that cannot *report* at all. Read-only inventory, a narrow reporting-prerequisite repair, and offline segmentation. These **never** write `AvailableUpdates`.
- **Screen for BitLocker exposure** — find the devices where driving the update would end in a recovery prompt, because BitLocker is sealed to the legacy PCR profile instead of PCR 7. Read-only.

The second family exists because at fleet scale a large share of devices land in the status report's **Unknown** bucket, which is a telemetry/reporting gap rather than a certificate failure. No amount of certificate remediation moves them, and pointing the first family at them just burns Proactive Remediation cycles.

### Which script do I want?

| I want to... | Use |
|--------------|-----|
| Push devices onto the 2023 certificates via Intune | [Detect_SecureBootUEFICA2023.ps1](Detect_SecureBootUEFICA2023.ps1) + [Remediate_SecureBootUEFICA2023.ps1](Remediate_SecureBootUEFICA2023.ps1) |
| Triage one device by hand, or push it manually | [SecureBootCertRemediation.ps1](SecureBootCertRemediation.ps1) |
| Feed an Ivanti baseline | [SecureBootCertDetection-Ivanti.ps1](SecureBootCertDetection-Ivanti.ps1) |
| Find out why devices show **Unknown** / never report | [Get-SecureBootCertInventory.ps1](Get-SecureBootCertInventory.ps1) |
| Fix the reporting blockers (without touching certificates) | [Repair-SecureBootReportingPrereqs.ps1](Repair-SecureBootReportingPrereqs.ps1) |
| Split "Not up to date" into who needs firmware vs. who is a false positive | [Split-SecureBootPopulation.ps1](Split-SecureBootPopulation.ps1) |
| Find devices that will drop into BitLocker recovery when the certificate update lands | [BitLockerPcrDetection-Ivanti.ps1](BitLockerPcrDetection-Ivanti.ps1) |

> **The two remediation scripts have opposite philosophies — do not treat them as interchangeable.** `Remediate_SecureBootUEFICA2023.ps1` arms the certificate update and deliberately exits `1` even when it takes no action, so the Intune dashboard never says `Fixed` for a device that is still non-compliant. `Repair-SecureBootReportingPrereqs.ps1` only repairs *reporting* prerequisites and exits `0` when it succeeds, because for that package "success" means the device can now report. Deploy them as separate Intune packages.

References:
- [KB5016061 - Secure Boot DB and DBX variable update events](https://support.microsoft.com/en-us/topic/37e47cf8-608b-4a87-8175-bdead630eb69) (event-id semantics)
- [KB5072718 - Sample Secure Boot Inventory Data Collection script](https://support.microsoft.com/en-us/topic/d02971d2-d4b5-42c9-b58a-8527f0ffa30b) (signal coverage)
- [KB5084567 - Sample Secure Boot E2E Automation Guide](https://support.microsoft.com/en-us/topic/f850b329-9a6e-40d1-823a-0925c965b8a0) (`AvailableUpdatesPolicy` guidance, GPO/MDM ownership)
- [Secure Boot playbook for certificates expiring in 2026](https://aka.ms/SecureBootPlaybook) (rollout strategy)
- [KB5085046 - Secure Boot troubleshooting guide](https://support.microsoft.com/en-US/servicing/os/secure-boot/2026/03/secure-boot-troubleshooting-guide) (`AvailableUpdates` bit order and expected progression)
- [KB5080921 - Monitoring Secure Boot certificate status with Intune Remediations](https://support.microsoft.com/en-US/servicing/os/secure-boot/2026/02/monitoring-secure-boot-certificate-status-with-microsoft-intune-remediations) (the reporting path used by the inventory script)
- [KB5085790 - Known issues and resolutions](https://support.microsoft.com/en-US/servicing/os/secure-boot/2026/03/known-issues-and-resolutions-for-secure-boot-certificates-updates) (hypervisor and OEM issues)
- [KB5080931 - Secure Boot certificate updates for Azure Virtual Desktop](https://support.microsoft.com/en-US/servicing/os/secure-boot/2026/02/secure-boot-certificate-updates-for-azure-virtual-desktop)
- [Secure Boot status report in Windows Autopatch](https://learn.microsoft.com/en-us/windows/deployment/windows-autopatch/monitor/secure-boot-status-report) (the report whose Unknown bucket this tooling explains)

## Scripts

### Drive the update

| File | Purpose | Modifies System? | Compliance Gate |
|------|---------|------------------|-----------------|
| [Detect_SecureBootUEFICA2023.ps1](Detect_SecureBootUEFICA2023.ps1) | Intune Proactive Remediation **detect** script | Never | **Strict**: `SecureBoot enabled AND Status=Updated AND Error=0 AND AvailableUpdates=0x4000 AND Event 1808` |
| [Remediate_SecureBootUEFICA2023.ps1](Remediate_SecureBootUEFICA2023.ps1) | Intune Proactive Remediation **remediate** script (two-pronged) | Yes (registry + scheduled task) | n/a (remediation) |
| [SecureBootCertRemediation.ps1](SecureBootCertRemediation.ps1) | Full detection + smart remediation + real-time monitoring (manual/triage) | Only with `-ForceRemediation` | **Strict**: `1808 AND Status=Updated AND Error=0 AND BootloaderSwapped (1799)` |
| [SecureBootCertDetection-Ivanti.ps1](SecureBootCertDetection-Ivanti.ps1) | Pure detection for Ivanti Custom Definition (Status / Reason / Expected / Found contract) | Never | **Legacy**: `Status=Updated AND Error=0` |

These four share the same TPM-WMI event-id classification and diagnostic signals; only the compliance gate, output contract, and remediation behavior differ.

### Measure the fleet

| File | Purpose | Modifies System? |
|------|---------|------------------|
| [Get-SecureBootCertInventory.ps1](Get-SecureBootCertInventory.ps1) | Intune Remediations **detection** script. Emits one compact JSON line covering every documented signal. | Never |
| [Repair-SecureBootReportingPrereqs.ps1](Repair-SecureBootReportingPrereqs.ps1) | Intune Remediations **remediation** script. Fixes only the three conditions that stop a device *reporting*. | Yes, narrowly (scheduled task, OneSettings, telemetry) |
| [Split-SecureBootPopulation.ps1](Split-SecureBootPopulation.ps1) | Offline. Segments the Intune CSV export into actionable populations. | Never (runs on your workstation) |

None of these three write `AvailableUpdates`, `MicrosoftUpdateManagedOptIn` or `HighConfidenceOptOut`, and none touch DBX, the boot manager, boot order or BitLocker. See [Fleet inventory and reporting prerequisites](#fleet-inventory-and-reporting-prerequisites).

### Screen for BitLocker exposure

| File | Purpose | Modifies System? | Compliance Gate |
|------|---------|------------------|-----------------|
| [BitLockerPcrDetection-Ivanti.ps1](BitLockerPcrDetection-Ivanti.ps1) | Reads the PCR validation profile of the OS volume's TPM protector. Same Status / Reason / Expected / Found contract as the Ivanti certificate detect script. | Never | Profile is exactly `7,11` or `4,7,11` |

The profile is read from `Win32_EncryptableVolume.GetKeyProtectorPlatformValidationProfile()` where that works, and from `manage-bde` where it does not. On many TPM 2.0 / Secure Boot integrity validation devices the WMI method returns `E_INVALIDARG` (`0x80070057`) even for a valid TPM protector, so the fallback is the normal path rather than an exception. The `manage-bde` parser anchors on the literal token `PCR` and then on a line of comma-separated integers — both survive localisation — and `found =` reports `Source: WMI` or `Source: manage-bde` so you can see which answered. Requires elevation either way.

The finding this screens for is the legacy profile **`0,2,4,11`**, which Windows falls back to when PCR 7 cannot be bound (`PCR7 Configuration = Binding Not Possible`). Devices in that state have been observed dropping into BitLocker recovery on the reboot that finalises a Secure Boot servicing update. Run this **before** pointing the *Drive the update* family at a population, and suspend BitLocker across the servicing reboot on anything it flags.

### Why so many scripts?

- **`Detect_SecureBootUEFICA2023.ps1` / `Remediate_SecureBootUEFICA2023.ps1`** — the Intune PR pair (v1.2). Logs to the Intune Management Extension log directory in CMTrace format with verbose DEBUG-level instrumentation. **Honest-reporting design** (symmetric across detect and remediate): both scripts exit `1` whenever the device is not actually running the new CA — even when the OS cannot help — so the Intune dashboard never marks a non-compliant device as `Compliant` (detect) or `Remediation successful` / `Fixed` (remediate). The remediate script still **takes no action** (no registry writes, no scheduled-task triggers) on hard-blocker conditions, so the PR cycle stays essentially free on devices that cannot be helped. Disambiguating STDOUT markers (`NON-COMPLIANT-NOT-ACTIONABLE` / `NON-COMPLIANT-PENDING-REBOOT` / `NON-COMPLIANT`) let dashboards filter work-to-do vs. operator-escalation cohorts. The remediate script also pre-flights the Secure-Boot-Update scheduled task and aborts cleanly when missing/disabled (see below).
- **`SecureBootCertRemediation.ps1`** — the standalone deployment / triage script. Detection-only by default; only writes registry / starts the scheduled task when explicitly invoked with `-ForceRemediation`. Strict gate avoids false-positive "compliant" verdicts caused by stale Event 1808 surviving NVRAM / BIOS resets.
- **`SecureBootCertDetection-Ivanti.ps1`** — the legacy Ivanti baseline detect script. The compliance verdict is intentionally unchanged to avoid baseline / ticket churn; the enhanced diagnostics (1808 / 1799 / 1801 / 1802 / 1803 / latest good / latest bad / FW errors) are surfaced in the `found =` line for triage only.
- **`BitLockerPcrDetection-Ivanti.ps1`** — answers a different question from every other script here. The rest ask *is this device on the 2023 certificates?*; this one asks *is it safe to put it there?* It never looks at `AvailableUpdates` or the TPM-WMI events, and a device can be perfectly compliant on the certificate gate while failing this one.

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

Before the gate is evaluated, both scripts apply several **hard-blocker reporting markers**. The device IS non-compliant for security posture purposes (it is not actually running the new CA), so both detect and remediate exit `1` — but the disambiguating STDOUT prefix lets the dashboard segment cohorts. The remediate script still takes **no action** on these paths (no registry writes, no task triggers), so the PR cycle stays essentially free on devices that cannot be helped:

| Condition | Detect exit | Remediate exit | STDOUT marker |
|-----------|-------------|----------------|---------------|
| `HighConfidenceOptOut = 1` | `1` | `1` (no side effects) | `NON-COMPLIANT-NOT-ACTIONABLE: HighConfidenceOptOut=1 (admin-managed exclusion).` |
| Event 1803 present (missing PK-signed KEK) | `1` | `1` (no side effects) | `NON-COMPLIANT-NOT-ACTIONABLE: Event 1803 (missing KEK - OEM responsibility).` |
| Event 1802 present (known firmware issue) | `1` | `1` (no side effects) | `NON-COMPLIANT-NOT-ACTIONABLE: Event 1802 (known firmware issue [KI_<n>]).` |
| Event 1800 present, Event 1808 not yet present | `1` | `1` (no side effects) | `NON-COMPLIANT-PENDING-REBOOT: Update staged, awaiting reboot.` |

**Reporting taxonomy** — dashboards can filter on the STDOUT marker prefix:

| Marker | Cohort | Operator action |
|--------|--------|-----------------|
| `COMPLIANT` | Done | None |
| `NON-COMPLIANT` | Active rollout cohort | Remediation will run automatically |
| `NON-COMPLIANT-PENDING-REBOOT` | Waiting on user/maintenance reboot | Trigger reboot policy / nudge users |
| `NON-COMPLIANT-NOT-ACTIONABLE` | Cannot be helped from the OS | Escalate to OEM, schedule manual workstreams, or document as accepted risk |
| `PREREQ` | Script ran in wrong context (32-bit PS) | Fix Intune PR settings (`Run script in 64-bit PowerShell = Yes`) |

In addition, the detect script emits **operational warnings** (logged but do **not** change the verdict) when:

- The `\Microsoft\Windows\PI\Secure-Boot-Update` scheduled task is missing or in a non-`Ready`/`Running` state — remediation will be ineffective.
- `CanAttemptUpdateAfter` (REG_BINARY/REG_QWORD `FILETIME` under `Servicing\DeviceAttributes`) is in the future — firmware throttle active; triggering the task is a no-op until it elapses.
- `AvailableUpdatesPolicy` disagrees with `AvailableUpdates` — GPO or Intune is the source of truth; direct registry writes may be reverted on the next policy refresh.

Diagnostic context surfaced in every detect run (logged at INFO level): `UEFICA2023ErrorEvent`, `MicrosoftUpdateManagedOptIn`, per-ID event counts (`1808`/`1803`/`1802`/`1801`/`1800`/`1795`/`1796`), parsed `BucketId` / `BucketConfidenceLevel` / `SkipReason`, and captured error codes from the latest `1795` / `1796`.

**Remediate** short-circuits with **no side effects** on the same hard-blocker conditions (`HighConfidenceOptOut=1`, Event 1803, Event 1802, Event 1800 without 1808): it exits `1` with a `NON-COMPLIANT-NOT-ACTIONABLE:` / `NON-COMPLIANT-PENDING-REBOOT:` STDOUT marker before touching the registry or the task. Exit `1` ensures Intune does **not** falsely report the device as `Remediation successful` / `Fixed` — the device is genuinely still non-compliant, we just chose not to make a system change because the OS cannot help. When none of those apply, it branches on the current `AvailableUpdates` value:

| Current value | Branch | Action |
|---------------|--------|--------|
| missing or `0` | **A. Initial arm** | Set `AvailableUpdates = 0x5944`, then trigger `\Microsoft\Windows\PI\Secure-Boot-Update` |
| non-zero, not `0x4000` (e.g. `0x5944`, `0x4100`, `0x4104`) | **B. Resume** | Trigger the scheduled task **without** modifying the registry (preserve in-flight state) |
| `0x4000` | **C. No-op** | Already at the terminal complete state |

The remediate script also performs a pre-flight check on the `Secure-Boot-Update` scheduled task **before** writing the registry. If the task is missing or in a non-`Ready`/`Running` state it exits `1` with a clear `FAIL:` message instead of arming the device only to discover the task cannot run.

> **Optional BitLocker suspend (Branch A only, v1.3+)** — top-of-file flag `$Script:SuspendBitLockerOnArm` (default `$false` for backwards compatibility). When set to `$true`, the script suspends BitLocker on the system drive for **1 reboot** *before* writing `AvailableUpdates = 0x5944`. This is a narrow workaround targeting HP KI devices whose firmware rolls the new certs back on the very next boot — 1 reboot covers the arm → firmware-apply transition without leaving BDE suspended any longer than necessary. Only Branch A triggers the suspend — branches B and C preserve in-flight state and don't move PCR 7 in a recovery-triggering way. If the suspend fails, the script aborts with `FAIL:` rather than arming an unprotected device. Intune PR cannot pass parameters, so flip the flag in the script before uploading to your tenant. Leave at `$false` if you already handle BDE suspend via a separate Intune Configuration Profile or pre-flight script.

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

### `BitLockerPcrDetection-Ivanti.ps1` (pure detection)

```powershell
# OS volume (default)
.\BitLockerPcrDetection-Ivanti.ps1

# A specific volume
.\BitLockerPcrDetection-Ivanti.ps1 -MountPoint D:
```

Must run elevated. Emits the same four-line contract:

```text
detected = true|false
reason   = <single sentence>
expected = Profile: 7,11 or 4,7,11
found    = Profile: <p> | Source: <WMI|manage-bde> | Protector: <t> | Protection: <s> | Conversion: <c> | SecureBoot: <b> | Event24604: <n> | ...
```

## AvailableUpdates State Machine

| Value | Meaning | Action |
|-------|---------|--------|
| `0x0` | Not started | Script sets `0x5944` and triggers task |
| `0x5944` | All updates queued | Wait for task processing |
| `0x4100` | Boot Manager staged, pending reboot | Reboot, then script triggers finalization |
| `0x4104` | KEK update pending | Reboot required |
| `0x4000` | Complete (conditional on 2011 CA) | No action needed |

### Processing order

The task processes the bits in a fixed order and **does not advance until the current step succeeds**:

| Order | Bit | Action | Success event | `AvailableUpdates` after |
|:-----:|-----|--------|:-------------:|--------------------------|
| 1 | `0x0040` | Windows UEFI CA 2023 -> db | 1036 | `0x5944` -> `0x5904` |
| 2 | `0x0800` | Microsoft Option ROM UEFI CA 2023 -> db | 1044 | `0x5904` -> `0x5104` |
| 3 | `0x1000` | Microsoft UEFI CA 2023 -> db | 1045 | `0x5104` -> `0x4104` |
| 4 | `0x0004` | Microsoft Corporation KEK 2K CA 2023 -> KEK | 1043 | `0x4104` -> `0x4100` |
| 5 | `0x0100` | 2023-signed boot manager | 1799 | `0x4100` -> `0x4000` |

Two consequences worth internalising, because both are commonly inverted:

- **The 2023 KEK is not a prerequisite for the DB updates.** DB updates are authorised by the *existing* 2011 KEK. A device sitting at `0x4104` with Event 1803 has already applied all three DB certificates and is blocked only on the OEM/hypervisor supplying a PK-signed KEK. If step 4 cannot be processed, the task still applies the boot manager at step 5.
- **Option ROM failures do not resolve themselves once KEK lands.** Option ROM is step 2, KEK is step 4. A device missing the Option ROM certificate is stalled *earlier* and will never reach step 4 on its own.

### Trust configuration changes what "missing" means

The `0x4000` modifier applies the Option ROM and Microsoft UEFI CA 2023 certificates **only if `Microsoft Corporation UEFI CA 2011` is already in db**. On a Microsoft-only-trust device those two certificates are *not applicable*, and Microsoft's own report shows the device **Up to date** with them absent.

A naive "are all five 2023 certificates present?" check therefore produces false positives on every Microsoft-only-trust device. `Get-SecureBootCertInventory.ps1` derives a `trust` value (`MSOnly` / `MSAnd3P` / `Unknown`) from db contents so the segmentation can account for it.

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

### `BitLockerPcrDetection-Ivanti.ps1` -- PCR profile gate

```
Compliant = (PCR validation profile of the OS volume TPM protector) IN { 7,11 ; 4,7,11 }
```

The measured profile is sorted and de-duplicated before comparison, so PCR order does not matter. Everything else is a finding, including devices with no TPM protector and devices where BitLocker is off — there is no profile to evaluate, so the script cannot vouch for them. `0,2,4,11` gets its own `reason =` string because it is the specific PCR 7 fallback this screen exists to catch.

A device that returns `Profile: N/A | Source: None` alongside a `ReadError` is **not** evidence of a bad profile — it means neither source answered. Treat those as a collection failure, not a finding.

## Output

### `Detect_SecureBootUEFICA2023.ps1`

Single-line STDOUT summary suitable for the Intune PR detection column. Possible outputs:

```text
COMPLIANT: Secure Boot UEFI CA 2023 update fully applied.
NON-COMPLIANT: UEFICA2023Status missing | AvailableUpdates=0x5944 | Event 1808 missing
NON-COMPLIANT-NOT-ACTIONABLE: HighConfidenceOptOut=1 (admin-managed exclusion).
NON-COMPLIANT-NOT-ACTIONABLE: Event 1803 (missing KEK - OEM responsibility).
NON-COMPLIANT-NOT-ACTIONABLE: Event 1802 (known firmware issue [KI_12345]).
NON-COMPLIANT-PENDING-REBOOT: Update staged, awaiting reboot. Gate failures: AvailableUpdates=0x4100 | Event 1808 missing
PREREQ: Not running in 64-bit PowerShell.
```

| Exit code | Meaning | Triggers remediation? |
|-----------|---------|----------------------|
| `0` | `COMPLIANT` only | No |
| `1` | Any `NON-COMPLIANT*` flavour or `PREREQ` failure | Yes (remediate also exits `1` with no side effects on `*-NOT-ACTIONABLE` / `*-PENDING-REBOOT`, so Intune does not falsely mark the device as `Fixed`) |

Full per-step trace is written to the CMTrace log file.

### `Remediate_SecureBootUEFICA2023.ps1`

Single-line STDOUT summary describing the action taken. Possible outputs:

```text
REMEDIATED: AvailableUpdates pre=0x0 post=0x5944. Reboot required to finalize.
NO-OP: Already at compliant terminal state (0x4000).
NON-COMPLIANT-NOT-ACTIONABLE: HighConfidenceOptOut=1 (admin-managed exclusion).
NON-COMPLIANT-NOT-ACTIONABLE: Event 1803 (missing KEK - OEM responsibility).
NON-COMPLIANT-NOT-ACTIONABLE: Event 1802 (known firmware issue [KI_12345]).
NON-COMPLIANT-PENDING-REBOOT: Update staged, awaiting reboot.
ABORT: Secure Boot disabled.
FAIL: Scheduled task '\Microsoft\Windows\PI\Secure-Boot-Update' is missing.
FAIL: Scheduled task is 'Disabled' (not Ready). Re-enable before remediation.
FAIL: Could not start Secure-Boot-Update scheduled task.
FAIL: Registry write error - <message>
```

| Exit code | Meaning | Intune PR dashboard |
|-----------|---------|----------------------|
| `0` | `REMEDIATED` or `NO-OP` (action taken or already compliant) | `Remediation successful` / `Fixed` |
| `1` | `NON-COMPLIANT-NOT-ACTIONABLE`, `NON-COMPLIANT-PENDING-REBOOT`, `PREREQ`, `ABORT`, or `FAIL` | Stays in the work queue (not falsely marked `Fixed`) |

Full per-step trace is written to the same CMTrace log file as the detect script.

### `SecureBootCertRemediation.ps1`

Produces a structured, color-coded detection report (registry values, event sweep, AvailableUpdates decode) followed by an Ivanti-compatible detection summary block:

- `Detected` -- `true` (non-compliant) or `false` (compliant)
- `Reason`   -- human-readable status explanation
- `Expected` / `Found` -- state comparison strings

When `-ForceRemediation` is supplied, the script additionally logs the registry write, scheduled-task trigger, and a 30-second monitoring loop tracking `AvailableUpdates` changes.

### `SecureBootCertDetection-Ivanti.ps1`

Four `Write-Host` lines, no banner, no color, no extra output -- safe to consume verbatim from an Ivanti Custom Definition or any detect channel that parses `key = value` pairs.

### `BitLockerPcrDetection-Ivanti.ps1`

The same four-line contract. A device on the legacy profile looks like this:

```
detected = true
reason = Legacy PCR profile 0,2,4,11 in use -- PCR 7 could not be bound. Secure Boot servicing on this device risks a recovery prompt.
expected = Profile: 7,11 or 4,7,11
found = Profile: 0,2,4,11 | Source: manage-bde | Protector: TPM | Protection: On | Conversion: FullyEncrypted | SecureBoot: True
```

`found =` also carries counts of BitLocker-Driver events **24604** (`the boot configuration options did not match expected values`) and **24636** (`bootmgr failed to obtain the volume master key from the TPM`) when either is present, so devices that have *already* been hit are visible in the detect output rather than only in a ticket. Both counts are informational and do not affect the verdict.

Diagnostic log: `C:\Windows\Temp\BitLockerPcrDetection-Ivanti.log` (append-only, never written to stdout).

## Fleet inventory and reporting prerequisites

### Why the status report is not enough

The Windows Autopatch **Secure Boot status report** (Intune admin center -> Reports -> Windows Autopatch -> Windows quality updates -> Reports tab) depends on Secure Boot diagnostic data reaching Microsoft. A device drops to **Unknown** or **Not applicable** when:

- `AllowTelemetry` is below **Required (1)**, or a proxy/firewall blocks diagnostic data
- the tenant has not enabled the **Data Processor Service for Windows (DPSW)**
- the `DisableOneSettingsDownloads` policy is enabled
- the device has been **inactive for more than 28 days**
- fewer than ~12 hours have passed since the update and restart

Autopatch also supports only **personal persistent** VMs on Azure Virtual Desktop, so multi-session, pooled non-persistent and RemoteApp session hosts never appear in it at all.

The Intune Remediations path has none of those dependencies: the script runs locally as SYSTEM and reports through the Intune script channel.

### `Get-SecureBootCertInventory.ps1`

Deploy as an Intune Remediations package **with no remediation script attached**, or paired with `Repair-SecureBootReportingPrereqs.ps1`.

| Setting | Value |
|---------|-------|
| Run this script using the logged-on credentials | **No** |
| Enforce script signature check | **No** |
| Run script in 64-bit PowerShell | **Yes** |

Collected signals:

```text
SecureBoot         AvailableUpdates, AvailableUpdatesPolicy, HighConfidenceOptOut,
                   MicrosoftUpdateManagedOptIn
SecureBoot\Servicing
                   UEFICA2023Status (NoValue when the key is absent), UEFICA2023Error,
                   UEFICA2023ErrorEvent, WindowsUEFICA2023Capable, BucketHash, ConfidenceLevel
UEFI variables     KEK 2023 / KEK 2011; db: Windows UEFI CA 2023, Option ROM UEFI CA 2023,
                   Microsoft UEFI CA 2023, Microsoft Corporation UEFI CA 2011,
                   Windows Production PCA 2011  -> derived trust configuration
Scheduled task     \Microsoft\Windows\PI\Secure-Boot-Update  state + last run
Policy             AllowTelemetry, DisableOneSettingsDownloads
Events             counts of 1795/1801/1802/1803/1032, latest event id + time, 1802 SkipReason
Platform           manufacturer, model, firmware version, OS build,
                   VMware / HyperV / Azure / OtherVM / Physical
```

Output is one JSON line with deliberately short keys, because **Intune truncates detection output at 2048 characters**. Check the length before adding fields.

Exit `0` when Secure Boot is disabled (out of scope) or `UEFICA2023Status = Updated`; exit `1` otherwise.

### `Repair-SecureBootReportingPrereqs.ps1`

Does exactly three things:

1. Re-enables `\Microsoft\Windows\PI\Secure-Boot-Update` if it is `Disabled`. It is deliberately **not started** — the task runs at startup and every 12 hours on its own.
2. Sets `DisableOneSettingsDownloads` to `0` if it is `1`.
3. Raises `AllowTelemetry` to `1` — **gated off by default** via `$SetTelemetry = $false` at the top of the file, because raising it is a data-sharing decision rather than a technical one. Until it is flipped, the script logs `NEEDS-APPROVAL (not changed)`.

If those policy values arrive from Group Policy or a Policy CSP, **fix them at source** — a local write regresses at the next policy refresh.

### `Split-SecureBootPopulation.ps1`

```powershell
# Intune > Devices > Remediations > <package> > Monitor > Device status
# Add the "Pre-remediation detection output" column before exporting.
.\Split-SecureBootPopulation.ps1 -Path .\DeviceStatus.csv -OutputCsv .\segmented.csv
```

Accepts the CSV export, a folder, or a file of one JSON object per line. First match wins:

| Segment | Meaning | Action |
|---------|---------|--------|
| `SecureBootOff` | Secure Boot disabled | Out of scope. Do **not** toggle Secure Boot to "fix" this — toggling can erase already-applied certificates |
| `ReportingBlocked` | Task not Ready, telemetry below Required, or OneSettings blocked | **This is the Unknown bucket.** Run the reporting-prereq remediation |
| `Updated` | `UEFICA2023Status = Updated` | None |
| `OptionRomNotApplicable` | Microsoft-only trust, with Windows UEFI CA 2023 and KEK 2023 both present | None — report false positive |
| `VirtualVMware` | VMware guest, not updated | Hypervisor-side. Microsoft lists no VMware entry in KB5085790; track Broadcom KB 423893 |
| `VirtualOtherHV` | Hyper-V / Azure / other guest | Hyper-V KEK 1795 fixed Mar 2026 (Apr 2026 for Server 2025) and needs the fix on **host and guest**. Azure Trusted Launch 1795 on KEK is an open known issue with no customer action |
| `FirmwareBlocked` | Event 1795/1802/1803/1032, or confidence `Temporarily Paused` / `Not Supported` | OEM firmware update, or document as an accepted exception |
| `PendingRestart` | `InProgress` with `AvailableUpdates = 0x4100` | Only the boot manager step remains; lands on the next restart |
| `NotTargeted` | Servicing key absent or `NotStarted`, prerequisites healthy | Nothing wrong — not yet targeted or not yet in a high-confidence bucket |
| `NeedsInvestigation` | Anything else | Manual |

`ReportingBlocked` is evaluated **before** status, so telemetry problems separate cleanly from certificate problems.

### Notes and caveats

- **`UEFICA2023Status` lives under the `Servicing` subkey.** Any inherited detection command that reads it directly from `...\Control\SecureBoot` returns nothing on every device, and any assessment built on that is void.
- **Certificate matching uses ISO-8859-1 (codepage 28591), not ASCII.** ASCII folds bytes above 127 to `?` and corrupts the scan. On builds that support `Get-SecureBootUEFI -Decoded`, prefer that.
- **Expiry is not a boot failure.** Devices that pass the deadline without the 2023 certificates still start and still take Windows updates; they stop receiving *early-boot* security fixes. Post-deadline remediation still works.
- **Watch the DBX.** This rollout adds certificates and does not revoke — the `0x5944` bitmask contains no revoke bit. But the revocation machinery is live (Event 1037 revokes Windows Production PCA 2011 into DBX). Once PCA 2011 is in DBX, **PXE boot applications and recovery media signed with it stop being trusted**. Plan boot-media re-signing before that happens.
- **Registry keys require the 11 Nov 2025 or later Windows update** on a supported build.
- **Intune Remediations requires** Windows 10/11 Enterprise E3/E5, Education A3/A5, or F3.

## Requirements

- Windows 10/11 or Windows Server 2016+
- PowerShell 5.1+
- Administrator privileges (registry writes, scheduled task execution)
- Scheduled task `\Microsoft\Windows\PI\Secure-Boot-Update` must exist (installed via Windows Update)
