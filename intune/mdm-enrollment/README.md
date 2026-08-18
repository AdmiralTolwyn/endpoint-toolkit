# Intune MDM Device-Cert Repair

Audit (and optionally repair) Windows hosts whose **Intune MDM device certificate** has expired, which wedges `omadmclient.exe` in a CPU-spinning cert-selection self-test (`CertificateManager::GetSslClientCertWithTipTest`).

Built for **AVD / cloned fleets**: when many hosts are cloned from one template, their MDM device certs share a `NotAfter` and expire together. Auto-renewal can silently fail on the clones, the cert lapses, and `omadmclient.exe` then loops at high CPU instead of failing clean — often several instances stacking up per host.

## Script

| Script | Modes | Purpose |
|--------|-------|---------|
| [Repair-IntuneMdmCert.ps1](Repair-IntuneMdmCert.ps1) | `Audit` (default, read-only) / `Repair` (destructive, gated) | Audit reports the MDM device cert's expiry, the renewal scheduled-task result, enrollment-endpoint reachability, MMP-C co-enrollment state, and `omadmclient.exe` pressure, then emits a per-host verdict. Repair tears down the MDM enrollment artifacts (both the Intune `MS DM Server` and the MMP-C `Microsoft Device Management` channels), removes the expired cert, re-enrolls headlessly, and polls the DeviceManagement event log to confirm the new enrollment. |

Local host only by design — fan out across a fleet with your own orchestration (Invoke-Command, a scheduled task, an AVD run command) and collect the emitted object.

## How It Works

```
┌─────────────────────────────────────────────────────────────────┐
│ AUDIT  (read-only, default)                                     │
└─────────────────────────────────────────────────────────────────┘

  [1] MDM enrollment(s)      HKLM:\SOFTWARE\Microsoft\Enrollments
                             primary = ProviderID 'MS DM Server' (Intune)
                             + MMP-C 'Microsoft Device Management' + MmpcEnrollmentFlag
  [2] MDM device cert        Cert:\LocalMachine\My, by Issuer regex
                             NotAfter -> DaysToExpiry / IsExpired / soon
  [3] Renewal task           EnterpriseMgmt\<GUID> task + LastTaskResult
  [4] Endpoint reachability  TCP-443 connect (authoritative) + HTTP GET enrich
  [5] omadmclient pressure   process count + accumulated CPU
       │
       └─► Verdict: NotEnrolled | CertMissing | CertExpired |
                    CertExpiringSoon | Healthy
           Fixable = CertExpired or CertMissing

┌─────────────────────────────────────────────────────────────────┐
│ REPAIR  (destructive, gated — only when Fixable)                │
└─────────────────────────────────────────────────────────────────┘

  Gate ─► skip if endpoint unreachable (unless -Force)
        ─► High-impact confirm prompt (unless -Confirm:$false)

  1. Back up enrollment registry roots (reg export) + cert manifest
     to %ProgramData%\MdmCertRepair\Backup_<timestamp>
  2. Stop omadmclient.exe
  3. Remove enrollment artifacts for each real enrollment GUID
     (MS DM Server + MMP-C 'Microsoft Device Management'; the
      WMI_Bridge_SCCM_Server co-mgmt bridge is left alone):
       • EnterpriseMgmt\<GUID> AND EnterpriseMgmtNonCritical\<GUID>
         scheduled tasks + task folders
       • 8 registry roots (Enrollments, Enrollments\Status,
         EnterpriseResourceManager\Tracked, PolicyManager\AdmxInstalled,
         PolicyManager\Providers, OMADM\Accounts/Logger/Sessions)
     3b. Clear MmpcEnrollmentFlag (blocks re-enroll if non-zero)
  4. Remove the expired Intune MDM device cert(s)
  5. Re-enroll: deviceenroller.exe /c /AutoEnrollMDM            (EnrollMode User, default)
              or /c /AutoEnrollMDMUsingAADDeviceCredential      (EnrollMode Device)
  6. Poll DeviceManagement event 75 (success) / 76 (failure) + cert
     store for up to -WaitSeconds to confirm the new enrollment
```

## Verdicts

| Verdict | Fixable | Meaning |
|---------|:-------:|---------|
| `Healthy` | no | Cert present and unexpired. No action. |
| `CertExpiringSoon` | no | Cert expires within `-ExpiryWarningDays` (default 30). Verify auto-renewal works before it lapses. |
| `CertExpired` | **yes** | Cert past `NotAfter`. Can't renew past expiry — re-enroll to mint a fresh one. |
| `CertMissing` | **yes** | Enrolled, but no cert from the expected issuer. Re-enroll to mint one. |
| `NotEnrolled` | no | No MDM enrollment found. Nothing to repair. |

## Usage

```powershell
# Read-only audit of the local host
.\Repair-IntuneMdmCert.ps1

# Audit and export the result object (fleet reporting)
.\Repair-IntuneMdmCert.ps1 | Export-Csv .\mdm-cert-audit.csv -NoTypeInformation

# Dry-run the repair (shows exactly what WOULD happen, changes nothing)
.\Repair-IntuneMdmCert.ps1 -Mode Repair -WhatIf

# Repair a user-driven Entra-joined host (default; run in the user's session)
.\Repair-IntuneMdmCert.ps1 -Mode Repair

# Repair a co-managed / AVD multi-session host (device credential)
.\Repair-IntuneMdmCert.ps1 -Mode Repair -EnrollMode Device

# Fully unattended (skip the confirm prompt; -Force skips the reachability gate)
.\Repair-IntuneMdmCert.ps1 -Mode Repair -EnrollMode Device -Force -Confirm:$false
```

| Parameter | Default | Notes |
|-----------|---------|-------|
| `-Mode` | `Audit` | `Audit` (read-only) or `Repair` (destructive). |
| `-ExpiryWarningDays` | `30` | Days-to-expiry threshold for `CertExpiringSoon`. |
| `-IssuerMatch` | `Microsoft Intune.*MDM Device CA` | Regex matched against the cert Issuer. Covers production and Beta-tenant issuers; excludes the unrelated `...Device Management` cert. |
| `-EnrollMode` | `User` | `User` -> `/c /AutoEnrollMDM` (user-driven Entra-joined; **must** run in the logged-on user's session). `Device` -> `/c /AutoEnrollMDMUsingAADDeviceCredential` (co-managed / Autopilot device-prep / AVD multi-session). |
| `-WaitSeconds` | `180` | How long Repair waits, polling DeviceManagement event 75/76 + the cert store, to confirm the new enrollment. `0` = fire-and-forget. |
| `-Force` | off | Proceed even when the endpoint-reachability probe says unreachable. Does **not** suppress the confirm prompt. |
| `-Confirm:$false` | — | Suppress the built-in High-impact confirmation prompt (for unattended runs). |

Run **elevated** for Repair.

## Notes & Caveats

- **Pick the right `-EnrollMode`.** A normal user-driven Entra-joined device (the expired enrollment carries a UPN) re-enrolls via the **user** credential (`/c /AutoEnrollMDM`, default) and **must** run in the logged-on user's session — as SYSTEM it can't get a user token and fails with an AAD `0xCAA8xxxx` error. Co-managed / Autopilot device-prep / AVD multi-session hosts use `-EnrollMode Device`. The script decodes these HRESULTs (`0xCAA82EE2` = AAD token WinHTTP timeout, `0x8018000A` = already-enrolled, …) into the output so you don't hand-decode.
- **Dual (MMP-C) enrollment blocks re-enroll.** Co-managed / hybrid boxes are often enrolled twice: `MS DM Server` (Intune classic) **and** `Microsoft Device Management` (MMP-C / declared config). Tearing down only the first leaves the second, and re-enroll then fails with `0x8018000A`. Repair tears down **both** real channels and clears `MmpcEnrollmentFlag`; it leaves the `WMI_Bridge_SCCM_Server` co-management bridge alone.
- **Confirm success from the event log, not just exit codes.** `deviceenroller.exe` returns immediately (exit 0 only means the trigger was accepted). Repair polls DeviceManagement event **75** (Auto MDM Enroll succeeded) / **76** (failed) and the reappearing cert for up to `-WaitSeconds`; events 77-80 (retry / `DMGetAadDeviceToken` "Access is denied") are transient and not treated as a verdict.
- **Re-enroll connectivity is the real gate.** Re-enrollment mints a new cert off the device's Entra identity. If an SSL-inspecting proxy sits in the path for Intune/Entra in **device** context, the new cert hits the same renewal failure and expires again. Fix the bypass before fixing the cert, or you'll re-hit this at the next expiry cycle. The TCP-443 reachability check confirms the endpoint answers, but it does **not** prove SSL inspection is absent.
- **Image-side fix.** Don't bake an already-enrolled state (with its soon-to-expire cert) into the gold image. Enroll **post-clone** and confirm auto-renewal works on the clones, or the fleet expires together again.
- **Expiry-only scope.** The audit verdict is driven purely by cert expiry. A date-valid cert with a dead/orphaned private key (a possible clone artifact) is reported `Healthy` and skipped.

## Source

The enrollment-teardown registry/task list and the headless re-enroll switch follow call4cloud (Rudy Ooms), [Troubleshooting Intune MDM Device enrollment errors](https://call4cloud.nl/intune-device-enrollment-errors-mdm-enrollment/) (§5.5 teardown, §7 device-credential re-enroll). This script discovers the enrollment GUID from the `MS DM Server` registry key and only ever iterates concrete GUIDs, avoiding the empty-`$EnrollmentID` "delete all tasks" bug present in older copies of that script.

## Requirements

- PowerShell 5.1+
- Elevation (for Repair, and for the registry/scheduled-task reads in Audit)
- An Entra-joined or Hybrid Entra-joined, Intune-enrolled Windows host
