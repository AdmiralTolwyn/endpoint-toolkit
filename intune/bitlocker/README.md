# BitLocker Key Escrow & MBAM Cleanup

Intune Proactive Remediation scripts to migrate BitLocker recovery keys from on-premises **MBAM** to **Entra ID** and (optionally) remove the legacy MBAM agent once escrow is confirmed.

Designed for hybrid fleets in the middle of a co-management / cloud-attach migration, where MBAM is still installed but recovery keys should now live in Entra ID.

## Scripts

| Script | Role | Purpose |
|--------|------|---------|
| [Detect_BitlockerBackupToAAD.ps1](Detect_BitlockerBackupToAAD.ps1) | Detection | Flags non-compliant if any fixed drive is not protected OR has no Entra ID escrow evidence (Event 845 or registry marker) |
| [Remediate_BitlockerBackupToAAD.ps1](Remediate_BitlockerBackupToAAD.ps1) | Remediation | Enables protection where needed, backs up every `RecoveryPassword` protector to Entra ID, waits for Event 845, then optionally uninstalls MBAM |
| [Detect-MbamUninstall.ps1](Detect-MbamUninstall.ps1) | Detection | Flags non-compliant only if MBAM is still present AND the OS drive already has confirmed Entra ID escrow (prevents premature removal) |
| [Remediate-MbamUninstall.ps1](Remediate-MbamUninstall.ps1) | Remediation | Re-confirms OS-drive escrow with a fresh backup + Event 845 wait, then stops/disables/uninstalls MBAM via `msiexec /x` |

The two pairs are independent and can be deployed together or separately:

- **Pair 1 (escrow)** — safe to roll out broadly; idempotent and read-mostly until a gap is found.
- **Pair 2 (MBAM cleanup)** — standalone cleanup wave for devices where escrow is already confirmed but the MBAM agent lingered.

## How It Works

```
┌─────────────────────────────────────────────────────────────────┐
│ Pair 1 — Backup BitLocker keys to Entra ID                      │
└─────────────────────────────────────────────────────────────────┘

  Detect ─► for each fixed, fully-encrypted drive:
              • Protection ON?
              • Event 845 (BitLocker-API) seen since cut-off date?
              • Registry marker HKLM:\SOFTWARE\ZF\BitLocker\Drive_X_*?
            ── any miss ──► exit 1 (non-compliant)

  Remediate ─► Phase 1: add TPM + RecoveryPassword protectors, turn on
              Phase 2: BackupToAAD-BitLockerKeyProtector for every key,
                       wait up to 30 min for Event 845, set registry marker
              Phase 3: if all drives confirmed AND OS drive marked,
                       uninstall MBAM (msiexec) — disable service on failure

┌─────────────────────────────────────────────────────────────────┐
│ Pair 2 — MBAM client cleanup                                    │
└─────────────────────────────────────────────────────────────────┘

  Detect ─► MBAMAgent service present  OR  MSI ProductCode installed?
            └─ AND OS drive has Event 845 / registry marker?
               └─ both true ──► exit 1 (non-compliant)

  Remediate ─► Fresh backup of all RecoveryPassword protectors on C:
              Wait up to 5 min for Event 845
              Stop + disable MBAMAgent
              msiexec /x {AEC5BCA3-A2C5-46D7-9873-7698E6D3CAA4} /qn
              Registry cleanup (package tracking key)
              Failure path: re-enable backup, disable service as fallback
```

## Compliance Signals

Both detection scripts treat a drive as "backed up" if **either** condition is true:

| Signal | Source | Notes |
|--------|--------|-------|
| Event ID **845** | `Microsoft-Windows-BitLocker-API` operational log | Filtered to `volume {drive} was backed up successfully to your Azure AD.` since 2022-01-01 |
| Registry marker | `HKLM:\SOFTWARE\ZF\BitLocker\Drive_{X}_BitLockerBackupToAAD = True` | Written by the remediation script after a confirmed backup; survives log rollover |

The registry marker is the durable source of truth — Event 845 disappears once the BitLocker-API log rolls over.

> **Note:** The registry root path (`HKLM:\SOFTWARE\ZF\BitLocker`) and package-tracking key (`HKLM:\SOFTWARE\ZF\SW-Distribution\Packages\ZF10001850`) reflect the original deployment context. Update these constants at the top of each script to match your own naming convention before publishing.

## MBAM Detection

MBAM presence is detected via:

1. **Service** — `MBAMAgent` Windows service exists
2. **MSI ProductCode** — `{AEC5BCA3-A2C5-46D7-9873-7698E6D3CAA4}` under either:
   - `HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall`
   - `HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall`

If your MBAM client version uses a different ProductCode, update `$Script:MbamProductGuid` in both Pair 2 scripts.

## Intune Proactive Remediation Settings

All four scripts assume:

| Setting | Value |
|---------|-------|
| Run this script using the logged-on credentials | **No** (SYSTEM context) |
| Enforce script signature check | **No** |
| Run script in 64-bit PowerShell | **Yes** |

Schedule the escrow pair daily; schedule the MBAM cleanup pair weekly (or on-demand) once you're confident the bulk of the fleet has escrowed.

## Logs

CMTrace-compatible logs are written to:

```
%ProgramData%\Microsoft\IntuneManagementExtension\Logs\PR_BackupBDE2EntraID.log
%ProgramData%\Microsoft\IntuneManagementExtension\Logs\PR_MBAMClientCleanup.log
```

Each log auto-rotates at **250 KB** (archived copy is NTFS-compressed via `compact /c`).

## Requirements

- Windows 10 / 11, joined to Entra ID (or Hybrid-joined)
- BitLocker-capable hardware with TPM
- PowerShell 5.1 (64-bit) — runs as SYSTEM under Intune Management Extension
- Devices must be able to reach Entra ID to escrow keys (`BackupToAAD-BitLockerKeyProtector`)

## Rollout Recommendation

1. **Pilot ring** — deploy Pair 1 only; verify Event 845 / registry markers and recovery-key visibility in Entra ID (Devices → device → BitLocker keys).
2. **Broad ring** — keep Pair 1 running fleet-wide as a guardrail; it remains idempotent.
3. **Cleanup wave** — once a target population shows confirmed escrow, deploy Pair 2 to remove the MBAM agent. The detection script self-gates on escrow confirmation, so it's safe to scope broadly.

## Disclaimer

Provided **AS-IS** without warranty of any kind. Test in a pilot ring before broad deployment. Removing the MBAM agent is irreversible without re-installation; verify recovery keys are visible in Entra ID before scaling Pair 2.
