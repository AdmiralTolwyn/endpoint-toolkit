# EntraCutover

> ## WARNING: EXPERIMENTAL - NOT SUPPORTED BY MICROSOFT
> This is community-built tooling that performs an operation Microsoft does **not**
> support in place. It is **experimental** (v0.1.0, not yet piloted on hardware).
> Do not run it in production without validating on a throwaway test ring first,
> and understand that Microsoft support may decline cases that touch device join
> state and tell you to reset the device.

In-place migration of a Windows device from **Hybrid Entra Join** (on-prem AD
joined + Intune enrolled) to **Entra-only join** — no OS reinstall, no Autopilot
reprovision, no wipe. CLI-only, CMTrace logging, resumable across reboots.

> **Unsupported by Microsoft.** Microsoft's documented path is wipe + Autopilot
> ("Windows Autopilot for existing devices"). This tool performs the same
> in-place technique that the commercial device-cutover products (ForensiT,
> Quest On Demand Migration, BitTitan/PowerSyncPro) use. Full rationale, source
> research, and the risk register are kept as an internal record (not published here).

---

## What it does

Five phases driven by a persisted state machine
(`%ProgramData%\EntraCutover\state.json`). After each reboot a SYSTEM startup
scheduled task re-invokes the tool with `-Mode Resume` and it continues from the
first incomplete step.

| Phase | Reversible? | What happens |
|-------|-------------|--------------|
| **Assess** | read-only | Preflight audit. Verdict `Ready` or a list of blockers. Safe to run anytime, on any host. |
| **Prepare** | yes | Stage tool copy + ppkg; **break-glass local admin** (password shown once); `djoin` offline-join rollback blob; registry backups; disable `Automatic-Device-Join` + set `BlockAADWorkplaceJoin`; suspend BitLocker; lock-screen notice; register resume task. |
| **Teardown** | **point of no return** | Purge Intune enrollment (reuses the proven `Repair-IntuneMdmCert` teardown + `MmpcEnrollmentFlag` fix); `dsregcmd /leave`; **unjoin the domain**; reboot. |
| **Join** | forward-only | Apply bulk-token provisioning package (`dism /Add-ProvisioningPackage`); verify Entra join; trigger + confirm MDM enrollment; reboot. Automatic **rollback-to-domain** from the blob after 2 failed attempts. |
| **Finalize** | forward-only | **Escrow BitLocker recovery key to the NEW device object** (hard-fail gate); resume BitLocker; stale-GPO cleanup; restore notice; schedule break-glass retirement; final report. |

### User data strategy: fresh profile + OneDrive KFM

The old domain profile is **left on disk untouched**. Users sign in with their
Entra ID (a new profile) and OneDrive Known Folder Move re-hydrates
Desktop/Documents/Pictures from the cloud. **KFM health is a Prepare gate** — the
migration refuses to start if KFM isn't confirmed syncing (override with
`-SkipKfmGate`, accepting the data-loss risk). There is **no in-place SID-remap**
in this version (see the design doc's decision D1 if that changes).

### Cleanup coverage (explicitly handled)

- **Old Intune enrollment** — Teardown removes the `EnterpriseMgmt\<GUID>`
  scheduled tasks + folder, the enrollment registry keys
  (`Enrollments`, `Enrollments\Status`, `EnterpriseResourceManager\Tracked`,
  `PolicyManager\{AdmxInstalled,Providers}`, `Provisioning\OMADM\*` incl. the
  no-dash `Accounts` variant), the `Microsoft Intune MDM Device CA` certs, and
  the `MmpcEnrollmentFlag` re-enrollment blocker.
- **Group Policy** — Finalize backs up then clears
  `%WinDir%\System32\GroupPolicy` + `GroupPolicyUsers`, the GP `History` key, and
  `State` children, so no stale `registry.pol` re-applies. **`HKLM\SOFTWARE\Policies`
  is deliberately left intact** (that's where Intune/OneDrive-KFM policy lives);
  the `BlockAADWorkplaceJoin` guard is re-asserted.
- **Re-enrollment** — the bulk-token join carries MDM enrollment; if it isn't
  present after join, Finalize's report flags it and the Join phase already tried
  `deviceenroller.exe /c /AutoEnrollMDM` after setting the `CloudDomainJoin`
  discovery URLs (Event 75 = success).

---

## Prerequisites

### Tenant-side runbook (the tool cannot do these — do them first)

`-Mode Migrate` **refuses to start without `-AcknowledgePrereqs`**, which asserts
the following are complete (design doc §4):

1. Migrating devices scoped **out of Entra Connect hybrid-join sync** (move the
   AD computer object to a non-synced OU / controlled validation). Otherwise the
   device hybrid-rejoins or the stale cloud object is resurrected.
2. Bulk enrollment token / `.ppkg` built (Windows Configuration Designer,
   ≤180-day expiry) **and** the `package_*` account excluded from MFA/Conditional
   Access (else `0xCAA2000C` or the device is deleted right after join).
3. **CA Token Protection** does not target migrated devices (bulk-token joins
   don't satisfy it — use `-JoinMode UserDriven` if it does).
4. Licensing: Entra ID P1 + Intune for every migrating user, users in MDM scope.
5. OneDrive KFM policy (`KFMSilentOptIn = <TenantId>`) deployed and **confirmed
   syncing** on the pilot ring.
6. Intune parity for anything GPO delivered today (Wi-Fi/802.1x, drive maps,
   certs). GPO settings are not rolled back — parity must come from Intune.

### Device-side

- Windows 10/11, TPM 2.0, run **elevated**.
- PowerShell 5.1+.
- Line-of-sight to a DC for a graceful unjoin (or `-OfflineUnjoin`).
- Internet egress to `login.microsoftonline.com` + `enrollment.manage.microsoft.com`.

---

## Usage

### 1. Assess first (always)

```powershell
.\Invoke-EntraCutover.ps1
```

Read-only. Emits a verdict object (`Ready`, `Blockers[]`, `Warnings[]`,
device/enrollment/KFM/profile facts). Export for a fleet:

```powershell
Get-ChildItem \\share\hosts.txt | ForEach-Object {
    Invoke-Command -ComputerName $_ -FilePath .\Invoke-EntraCutover.ps1
} | Export-Csv .\readiness.csv -NoTypeInformation
```

Blockers include: not hybrid-joined, pending reboot, unreachable DC/endpoints,
missing/expired ppkg, KFM unhealthy, **EFS-encrypted files present** (would be
unreadable post-migration — decrypt first).

### 2. Migrate

```powershell
$cred = Get-Credential CONTOSO\svc-unjoin
.\Invoke-EntraCutover.ps1 -Mode Migrate `
    -PpkgPath .\AADJ-Bulk-Expires-2026-12-31.ppkg `
    -TenantId 11111111-2222-3333-4444-555555555555 `
    -DomainCredential $cred `
    -AcknowledgePrereqs
```

Runs Assess → Prepare, then prompts once at the **point of no return** before
Teardown (suppress with `-Force` for unattended waves). **Record the break-glass
password shown during Prepare** — it is displayed once and is not recoverable
from the log or `state.json`.

### 3. Watch progress (works mid-migration, across reboots)

```powershell
.\Invoke-EntraCutover.ps1 -Mode Status
```

### 4. Roll back

```powershell
.\Invoke-EntraCutover.ps1 -Mode Rollback
```

- **Before the domain unjoin:** undoes Prepare (re-enables the join task,
  restores registry/BitLocker/notice, removes the break-glass account) and, if a
  `/reuse` blob was staged, re-applies it to restore machine trust.
- **After unjoin, join failed:** re-applies the `djoin` blob to rejoin the domain
  offline (no DC line-of-sight needed), re-enables hybrid re-registration, keeps
  the break-glass account.
- **After join verified:** refuses — the migration is forward-only from here; use
  tenant-side remediation.

### Key parameters

| Parameter | Purpose |
|-----------|---------|
| `-Mode` | `Assess` (default) / `Migrate` / `Resume` / `Status` / `Rollback` |
| `-PpkgPath` | Bulk-enrollment provisioning package (required for `Ppkg` join) |
| `-JoinMode` | `Ppkg` (default, silent) / `UserDriven` (user joins via Settings; immune to `package_*` CA issues and Token Protection) |
| `-DomainCredential` | For graceful unjoin + rollback-blob capture |
| `-OfflineUnjoin` | Permit unjoin without DC line-of-sight (NIC-isolated workgroup swap) |
| `-TenantId` | Expected tenant; validated pre- and post-join |
| `-SkipKfmGate` | Proceed with unhealthy KFM (accept data-loss risk) |
| `-FallbackRetentionDays` | Break-glass admin lifetime after Finalize (default 7) |
| `-AcknowledgePrereqs` | **Required for Migrate** — asserts the tenant runbook is done |
| `-Force` | Skip confirmations (implied in Resume) |

---

## Safeguards

- **Assess-first, blocker-gated** — Migrate aborts on any preflight blocker
  unless `-Force`.
- **Point-of-no-return confirmation** before the domain unjoin (attended runs).
- **Break-glass local admin** created in Prepare, retained through
  `-FallbackRetentionDays`, so you never lose local access if the new identity
  misbehaves.
- **Offline-join rollback blob** captured in Prepare → automatic domain rejoin if
  the Entra join fails after unjoin.
- **BitLocker**: suspended across the reboot chain, recovery key **re-escrowed to
  the new device object before any old-object cleanup** (Finalize hard-fails if
  escrow fails — the worst outcome is a recovery-key lockout).
- **Idempotent, resumable steps** — every action is recorded in `state.json`;
  completed steps are skipped on resume, so a crash or an extra reboot never
  repeats destructive work. (Verified by the state-machine tests.)
- **Registry backups** of all enrollment/WorkplaceJoin/CloudDomainJoin keys
  before teardown.

## Logging & output

- **Log**: CMTrace format at
  `%ProgramData%\Microsoft\IntuneManagementExtension\Logs\EntraCutover.log`
  (falls back to `%TEMP%`, rotates at 5 MB). Open with CMTrace. Every step,
  every external command exit code, and every registry mutation is logged.
- **Console**: phase banners, color-coded INFO/WARN/ERROR/SUCCESS lines,
  break-glass block, and a final report with the **tenant cleanup checklist** and
  **first-sign-in checklist**.
- **Report**: `%ProgramData%\EntraCutover\report.json` and the pipeline result
  object (`Export-Csv`-able).

### Exit codes

| Code | Meaning |
|------|---------|
| 0 | Success |
| 1 | Blocked in preflight |
| 2 | Failed before the point of no return (fully reverted) |
| 3 | Failed after PONR (rollback executed) |
| 4 | Failed after PONR — manual intervention required (break-glass account live) |

---

## What the user experiences — a NEW profile (communicate this before a wave)

This tool does **not** move or convert the user's profile. After it completes the
user signs in with their Entra ID and Windows creates a **brand-new, empty
profile** (`C:\Users\<newname>`). The old domain profile is left **fully intact**
at `C:\Users\<olduser>` — nothing is deleted, but nothing is carried into the new
profile either. There is no SID-remap and no "same-UPN reconnect": every app
starts from scratch, so this is not a "re-auth" story — it's a first-run story.

**Returns on its own (from the cloud):**

- Desktop / Documents / Pictures — re-hydrated by OneDrive KFM (the reason KFM is
  a hard Prepare gate).
- Edge / Chrome favorites + saved passwords — **only** if the user had browser
  sync to their account; they re-download after the user signs into the browser.
- Office / M365 apps — reactivate on the Entra sign-in (modern auth).

**Does NOT return — set up from scratch in the new profile:**

- OneDrive, Teams, Outlook, Authenticator, every other app — signed out; user
  re-signs in and reconfigures. Outlook rebuilds its OST/profile.
- Windows Hello (PIN / biometrics) — re-enrolled from scratch.
- Wi-Fi profiles, mapped drives, printers, VPN clients — re-provision (push Wi-Fi
  + drive/printer maps via Intune).
- Credential Manager entries and per-user certificates — a fresh profile has
  none. EFS-encrypted files can't be read in the new profile — Assess **blocks**
  on these so you decrypt first.
- **Everything left behind in the old profile**: all of `AppData` (Outlook
  signatures/templates, app settings), `Downloads`, PST files, and any loose
  folders not under KFM. Users who need these must copy them out of
  `C:\Users\<olduser>` (readable by a local admin) before that profile is
  eventually cleaned up.

Post-migration, complete the **tenant cleanup checklist** from the final report:
disable the on-prem AD computer account **first**, verify the BitLocker key on
the new object, then delete the stale hybrid object, assign the Intune primary
user, and switch the Autopilot profile to Entra-joined.

---

## Files

```
scripts/EntraCutover/
  Invoke-EntraCutover.ps1     entry point: state machine, logging, dispatcher, shared helpers
  lib/
    Phase1.Assess.ps1         read-only preflight (+ shared Get-ECKfmStatus)
    Phase2.Prepare.ps1        reversible staging
    Phase3.Teardown.ps1       enrollment purge + leave + unjoin (PONR)
    Phase4.Join.ps1           ppkg join + MDM enrollment + auto-rollback
    Phase5.Finalize.ps1       BitLocker escrow, GPO cleanup, report
    Rollback.ps1              pre/post-PONR rollback
  README.md                   this file
```

Status: **v0.1.0 — validated (parse/lint/state-machine tests pass); not yet
piloted on hardware.** Test on a throwaway VM ring before any production wave.
