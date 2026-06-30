# Intune MDM Device-Cert Repair

Audit (and optionally repair) Windows hosts whose **Intune MDM device certificate** has expired, which wedges `omadmclient.exe` in a CPU-spinning cert-selection self-test (`CertificateManager::GetSslClientCertWithTipTest`).

Built for **AVD / cloned fleets**: when many hosts are cloned from one template, their MDM device certs share a `NotAfter` and expire together. Auto-renewal can silently fail on the clones, the cert lapses, and `omadmclient.exe` then loops at high CPU instead of failing clean — often several instances stacking up per host.

## Script

| Script | Modes | Purpose |
|--------|-------|---------|
| [Repair-IntuneMdmCert.ps1](Repair-IntuneMdmCert.ps1) | `Audit` (default, read-only) / `Repair` (destructive, gated) | Audit reports the MDM device cert's expiry, the renewal scheduled-task result, enrollment-endpoint reachability, and `omadmclient.exe` pressure, then emits a per-host verdict. Repair tears down the MDM enrollment artifacts, removes the expired cert, and re-enrolls headlessly via the device credential. |

Local host only by design — fan out across a fleet with your own orchestration (Invoke-Command, a scheduled task, an AVD run command) and collect the emitted object.

## How It Works

```
┌─────────────────────────────────────────────────────────────────┐
│ AUDIT  (read-only, default)                                     │
└─────────────────────────────────────────────────────────────────┘

  [1] MDM enrollment(s)      HKLM:\SOFTWARE\Microsoft\Enrollments
                             primary = ProviderID 'MS DM Server' (Intune)
  [2] MDM device cert        Cert:\LocalMachine\My, by Issuer regex
                             NotAfter -> DaysToExpiry / IsExpired / soon
  [3] Renewal task           EnterpriseMgmt\<GUID> task + LastTaskResult
  [4] Endpoint reachability  HTTPS HEAD to the enrollment DiscoveryUrl
  [5] omadmclient pressure   process count + accumulated CPU
       │
       └─► Verdict: NotEnrolled | CertMissing | CertExpired |
                    CertExpiringSoon | Healthy
           Fixable = CertExpired or CertMissing

┌─────────────────────────────────────────────────────────────────┐
│ REPAIR  (destructive, gated — only when Fixable)                │
└─────────────────────────────────────────────────────────────────┘

  Gate ─► skip if endpoint unreachable (unless -Force)
        ─► confirm prompt (unless -Force)

  1. Back up enrollment registry roots (reg export) + cert manifest
     to %ProgramData%\MdmCertRepair\Backup_<timestamp>
  2. Stop omadmclient.exe
  3. Remove enrollment artifacts for the Intune enrollment GUID:
       • EnterpriseMgmt\<GUID> scheduled tasks + task folder
       • 8 registry roots (Enrollments, Enrollments\Status,
         EnterpriseResourceManager\Tracked, PolicyManager\AdmxInstalled,
         PolicyManager\Providers, OMADM\Accounts/Logger/Sessions)
  4. Remove the expired Intune MDM device cert(s)
  5. Re-enroll: deviceenroller.exe /c /AutoEnrollMDMUsingAADDeviceCredential
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

# Repair the local host (still gated on endpoint reachability)
.\Repair-IntuneMdmCert.ps1 -Mode Repair -Force
```

| Parameter | Default | Notes |
|-----------|---------|-------|
| `-Mode` | `Audit` | `Audit` (read-only) or `Repair` (destructive). |
| `-ExpiryWarningDays` | `30` | Days-to-expiry threshold for `CertExpiringSoon`. |
| `-IssuerMatch` | `Microsoft Intune.*MDM Device CA` | Regex matched against the cert Issuer. Covers production and Beta-tenant issuers; excludes the unrelated `...Device Management` cert. |
| `-Force` | off | Skip the confirmation prompt **and** the endpoint-reachability gate. |

Run **elevated** for Repair.

## Notes & Caveats

- **Re-enroll connectivity is the real gate.** Re-enrollment mints a new cert off the device's Entra identity. If an SSL-inspecting proxy sits in the path for Intune/Entra in **device** context, the new cert hits the same renewal failure and expires again. Fix the bypass before fixing the cert, or you'll re-hit this at the next expiry cycle. The reachability check confirms the endpoint answers, but it does **not** prove SSL inspection is absent.
- **Device-credential re-enroll** (`/c /AutoEnrollMDMUsingAADDeviceCredential`) is the headless/no-user path, supported for co-management and AVD multi-session host pools. It does **not** require a logged-in user.
- **Image-side fix.** Don't bake an already-enrolled state (with its soon-to-expire cert) into the gold image. Enroll **post-clone** and confirm auto-renewal works on the clones, or the fleet expires together again.
- **Expiry-only scope.** The audit verdict is driven purely by cert expiry. A date-valid cert with a dead/orphaned private key (a possible clone artifact) is reported `Healthy` and skipped.

## Source

The enrollment-teardown registry/task list and the headless re-enroll switch follow call4cloud (Rudy Ooms), [Troubleshooting Intune MDM Device enrollment errors](https://call4cloud.nl/intune-device-enrollment-errors-mdm-enrollment/) (§5.5 teardown, §7 device-credential re-enroll). This script discovers the enrollment GUID from the `MS DM Server` registry key and only ever iterates concrete GUIDs, avoiding the empty-`$EnrollmentID` "delete all tasks" bug present in older copies of that script.

## Requirements

- PowerShell 5.1+
- Elevation (for Repair, and for the registry/scheduled-task reads in Audit)
- An Entra-joined or Hybrid Entra-joined, Intune-enrolled Windows host
