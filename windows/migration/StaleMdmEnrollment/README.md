# StaleMdmEnrollment

**Version:** 1.1
**Author:** Anton Romanyuk

> **Disclaimer:** These scripts are provided "as-is" without warranty of any kind, express or implied. Use at your own risk. The author assumes no liability for any damage or data loss resulting from its use. Always test in a non-production environment before deployment.

| Script | Role |
|---|---|
| [`Get-StaleMdmEnrollment.ps1`](Get-StaleMdmEnrollment.ps1) | **Read-only** detector. Decides whether a device's MDM enrollment belongs to the tenant it is actually joined to. Changes nothing. |
| [`Remove-StaleMdmEnrollment.ps1`](Remove-StaleMdmEnrollment.ps1) | **Destructive** remediation. Purges one stale or orphaned enrollment so the device can re-enrol. Dry-run by default. |

The detector compares the tenant the device is currently Entra-joined to against the tenant that owns each live MDM enrollment, and reports a verdict plus the evidence needed to prove the old channel is still functioning.

## Problem

After a tenant-to-tenant migration a hybrid-joined device can end up in a **split state**:

- The Entra device object, the PRT and `dsregcmd`'s `TenantId` all belong to the **new** tenant.
- The MDM enrollment, the MDM client certificate and the OMA-DM sync channel still belong to the **old** tenant.

The device therefore keeps checking in to the **old** tenant's Intune — it remains fully manageable there, receives the old tenant's policy and apps, and can still be wiped from it — while appearing **unmanaged** in the new tenant. Nothing in either portal makes this obvious: the new tenant simply shows a device with no MDM, and the old tenant shows a device that looks perfectly healthy.

The usual symptom trail is Entra sign-in / `AAD.evtx` transactions from the device looping on **`AADSTS50020`** ("user from identity provider does not exist in tenant") against the old tenant ID, because the enrollment keeps presenting an identity the old tenant no longer knows.

This is hard to query centrally. Neither tenant's Intune inventory exposes "which tenant does this enrollment belong to", so the determination has to be made **on the device**.

### The second failure mode

The same migration can leave a device in a state that looks like the opposite problem. If the old tenant's Intune **retires** the device, the `MS DM Server` enrollment is removed but the declared-configuration enrollment is not — so the device shows up in *no* tenant's Intune, while Windows still believes it is MDM-enrolled and therefore never attempts to enrol anywhere. Auto-enrollment and `deviceenroller.exe` exit silently with no events, which reads as "nothing is even trying" rather than as a fault.

A detector that only inspects `MS DM Server` reports this device as `NotEnrolled` and stops. It is not simply unenrolled — it is blocked, and the surviving enrollment is usually still bound to the old tenant.

## The test

Two registry facts, compared:

| Side | Source |
|---|---|
| **Joined** tenant | `dsregcmd /status` → `TenantId`, cross-checked against `HKLM\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\JoinInfo\<certThumbprint>\TenantId` |
| **Enrolled** tenant | `HKLM\SOFTWARE\Microsoft\Enrollments\<enrollmentGuid>\AADTenantID`, on each subkey whose `ProviderID` is `MS DM Server` **or** `Microsoft Device Management` |

A channel's tenant differing from the joined tenant is the whole detection. Everything else the script collects is corroboration or triage context.

**Both channels are evaluated independently**, because each carries its own `AADTenantID` and they are enrolled and torn down separately. Checking only `MS DM Server` — as version 1.0 did — misses a cross-tenant declared-configuration enrollment entirely once Intune has retired the Intune enrollment away, because there is then no `MS DM Server` key left to compare against. See [The second channel](#the-second-channel--windows-declared-configuration) below.

`dsregcmd` wins over the registry for the joined tenant because a migrated device can retain a leftover `JoinInfo` subkey from the pre-migration tenant, and PowerShell 5.1 exposes no `LastWriteTime` on registry keys, so the subkeys cannot be ordered by recency.

### Why `ProviderID = 'MS DM Server'`

`HKLM\SOFTWARE\Microsoft\Enrollments` holds roughly thirty subkeys, most of them internal CSP plumbing with no `ProviderID` at all. `ProviderID` is the OMA-DM management-account identifier the MDM service assigns itself at enrollment — delivered in the `w7 APPLICATION` provisioning payload and surfaced by the DMClient CSP as `./Device/Vendor/MSFT/DMClient/Provider/{ProviderID}`. Intune uses the literal string `MS DM Server`, so filtering on it isolates the genuine MDM enrollment.

Other values seen on a real machine — `WMI_Bridge_SCCM_Server`, `Cloud Authority`, `Local Authority`, `Deploy Authority` — are pseudo-enrollments for internal CSP and policy-precedence plumbing. They carry a `ProviderID` but **no** `AADTenantID`, and are deliberately excluded. Filtering on "has a `ProviderID`" rather than on the two literal strings pulls them in and produces four extra rows with no tenant.

`Microsoft Device Management` is the exception — it is a real cloud channel and **is** collected. See below.

Note that `MS DM Server` identifies *an Intune enrollment*, **not which tenant's Intune**. If both the old and new tenant are Intune, both produce the same `ProviderID`; the `AADTenantID` on that key is what distinguishes them.

## The second channel — Windows declared configuration

Windows can hold a **second, fully live enrollment** on the declared-configuration channel (WinDC, enrollment name `MicrosoftManagementPlatformCloud`, also written MMP-C):

| Value | Observed |
|---|---|
| `ProviderID` | `Microsoft Device Management` |
| `EnrollmentType` | `0x1a` (26) |
| `DiscoveryServiceFullURL` | `https://discovery.dm.microsoft.com/EnrollmentConfiguration?api-version=1.0` |
| `AADResourceID` | `https://checkin.dm.microsoft.com` |

> **Those hostnames are the production ring.** A host on the preview ring carries `discovery.dm-beta.microsoft.com` and `checkin.dm-beta.microsoft.com` instead (measured). Match on `ProviderID`, never on the hostname.

Both the literal string `MicrosoftManagementPlatformCloud` and `Enroll Type: (0x1A)` appear verbatim in Microsoft's declared-configuration troubleshooting examples, so the enrollment name and the type code are Microsoft-documented — the *meaning* of the numeric `EnrollmentType` values is not, and no mapping table exists anywhere.

It carries the same footprint as a classic enrollment — `Enrollments` key, OMA-DM account, client certificate, and a complete `EnterpriseMgmt\<guid>\` task set. Microsoft documents it as a **dual** enrollment: *"This dual enrollment is only allowed if the device is already enrolled into a primary mobile device management (MDM) server."*

Two consequences matter for migrations:

**It appears to survive an Intune Retire.** ⚠️ *Unverified — this is our own field observation on one migrated device, and no Microsoft page or community write-up states it.* Microsoft's Retire article says only that the device is unenrolled from Intune, and is silent on the linked enrollment. What we measured is a live declared-configuration enrollment with no `MS DM Server` enrollment behind it; that a Retire is what produced that state is inference. It is, however, a state Microsoft's own precondition says should not exist. The device then appears unmanaged in Intune while Windows still believes it is MDM-enrolled, so auto-enrollment and `deviceenroller.exe` exit **silently, with no events at all**.

**It carries its own tenant.** Measured on a migrated device: the declared-configuration enrollment was still bound to the **old** tenant, with a UPN on the old domain and `IsFederated = 1` pointing at the old ADFS, while the device was Entra-joined to the **new** tenant. Any certificate renewal on that channel authenticates against the old tenant's IdP and can never succeed.

### Healthy vs. broken layout

| | Healthy | Broken (post-migration) |
|---|---|---|
| `MS DM Server` key | present | **absent** (retired away) |
| `Microsoft Device Management` key | present | present |
| Enrollment GUIDs | **different** from each other | — |
| `AADTenantID` on both | **same**, matching the join | old tenant, **not** matching the join |

> **Do not confuse the enrollment GUID with the tenant GUID.** The `Enrollments\<GUID>` subkey name is a locally-minted enrollment identifier, unique per device and per enrollment — it will never match a tenant ID or another machine. The tenant lives in the `AADTenantID` **value** inside that key. Comparing the key name on one machine against the `AADTenantID` on another is the easiest way to convince yourself of a contradiction that does not exist.

## Active enrollment vs. orphans

Windows keeps a **single active** MDM enrollment, but orphaned `Enrollments\<GUID>` keys survive a failed or partial unenrollment — which is why manual teardown has to clear eight separate registry roots. A device that has already had a half-finished cleanup can therefore hold an inert key from the old tenant *and* a healthy new one.

To stop that producing a false positive, the script scores every enrollment and lets only the **active** one drive the verdict:

| Signal | Weight | Why |
|---|---|---|
| `HKLM\SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\<id>` exists | +3 | the OMA-DM account is what the sync engine actually binds to |
| Scheduled tasks under `\Microsoft\Windows\EnterpriseMgmt\<id>\` | +2 | the sync schedule still exists |
| MDM client certificate resolvable via `DMPCertThumbPrint` | +2 | the channel has credentials |
| that certificate is not expired | +1 | and they are still usable |
| `DiscoveryServiceFullURL` matches `dsregcmd`'s `MdmUrl` | +1 | weak — see below |

Highest `LivenessScore` wins; ties are broken by the most recent `SyncTaskLastRun`. A top score of **0** means every key is residue, and the verdict becomes `NotEnrolled`. Everything not selected is counted in `OrphanEnrollmentCount` — visible, but never verdict-driving.

Ranking is done **per channel**. The two are independent, and a declared-configuration enrollment must never outrank and mask a live Intune one.

> **The URL match is deliberately the weakest signal.** Every Intune tenant uses the same enrollment discovery URL, so `DiscoveryServiceFullURL` matching `MdmUrl` cannot distinguish two Intune enrollments belonging to different tenants. It only rules out an enrollment pointing at a non-Intune MDM.

## What it deliberately does not use

**The MDM certificate's Subject CN.** It contains the **Intune** device ID, not the **Entra** device ID, and the two differ on perfectly healthy machines (measured on a known-good host: `CN=1742e0d5-…` against `dsregcmd DeviceId=4a71d882-…`). Comparing them would flag the entire fleet as stale.

**`LinkedEnrollment/EnrollStatus` as a liveness signal.** Measured on a fully healthy host: `EnrollStatus = 3` ("Enrollment Failed" per Microsoft's published mapping) and `LastError = 0`, while the declared-configuration enrollment it points at was live, certificated and syncing. The value evidently records some earlier attempt and is not refreshed on success. It is reported for triage and carries **no** weight in `LivenessScore` or the verdict.

**`MMPCLocked`.** Reported raw, never interpreted. It is not a documented CSP node and its meaning is unknown — the only community source describing it says outright that it is a guess.

> **Correction to earlier versions of this document.** They claimed `SslClientCertReference` "under the enrollment key is frequently blank". That was wrong on both counts. The value does not live under `Enrollments\<guid>` at all — it lives under `Provisioning\OMADM\Accounts\<guid>` in the form `MY;System;<thumbprint>`, and on a healthy host it is populated with the same thumbprint as `DMPCertThumbPrint`. The certificate is resolved from `DMPCertThumbPrint` simply because that value is on the key already being read.

## Detection parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `ExpectedTenantId` | `string` | the device's own joined tenant | The tenant this device *should* be managed by. Defaulting to the join state is the correct baseline after a migration and means the new tenant GUID never has to be hardcoded. Supply it explicitly only to test a hypothetical, or to audit devices that are not joined anywhere. |
| `StaleTenantId` | `string` | — | GUID of the known-bad (old) tenant. Purely a reporting refinement: a matching stale enrollment yields `StaleKnownTenant` instead of the generic `StaleOtherTenant`, which separates migration fallout from unrelated cross-tenant enrollments in the same export. |
| `Json` | `switch` | Off | Emit compressed JSON instead of a PowerShell object, for collection via ConfigMgr / a GPO startup script / an RMM into a file share. |

## Verdicts

| Verdict | Meaning |
|---|---|
| `Clean` | Active enrollment's tenant matches the joined tenant. Orphans may still be present — check `OrphanEnrollmentCount` and `Flags`. |
| `StaleKnownTenant` | Active Intune enrollment is cross-tenant **and** matches `-StaleTenantId`. Migration fallout. |
| `StaleOtherTenant` | Active Intune enrollment is cross-tenant, but not the tenant you named. Investigate. |
| `MmpcOrphan` | No live `MS DM Server` enrollment, but a live declared-configuration one **in the correct tenant**. Windows counts as enrolled, so re-enrollment silently no-ops. The orphan has to be torn down first. |
| `MmpcOrphanCrossTenant` | As `MmpcOrphan`, but the surviving enrollment belongs to a **different** tenant — migration residue that was never torn down. Worse: renewal authenticates against the old tenant's IdP and can never succeed. |
| `NotEnrolled` | No live enrollment on either channel. Not stale, but not managed anywhere either. |
| `Unknown` | No baseline to compare against (device not Entra joined and no `-ExpectedTenantId`), or the active enrollment exposes no tenant at all. |

### Flags

`Verdict` is a single string, but several conditions can be true at once — a device can be both orphaned *and* cross-tenant. `Flags` reports all of them:

| Flag | Meaning |
|---|---|
| `IntuneCrossTenant` | the live `MS DM Server` enrollment belongs to another tenant |
| `MmpcCrossTenant` | the live declared-configuration enrollment belongs to another tenant |
| `MmpcOrphan` | declared configuration is live with no Intune enrollment behind it |
| `MmpcUnlinked` | a live declared-configuration enrollment that **nothing points at** — no `LinkedEnrollmentId` on any Intune enrollment names it. The strongest available evidence that the parent was removed out from under the child. |
| `MmpcEnrollmentFlagBlocking` | `HKLM\SOFTWARE\Microsoft\Enrollments!MmpcEnrollmentFlag` is `2` — see below |
| `OrphanResidue` | inert enrollment keys left by an incomplete teardown |
| `MultipleJoinInfo` | more than one `CloudDomainJoin\JoinInfo` subkey — migration residue |

Triage on `Flags`, not on `Verdict` alone.

### `MmpcEnrollmentFlag`

A `REG_DWORD` sitting directly on the `Enrollments` root — not under any GUID. **Undocumented by Microsoft**, but independently reported by Quest (two support KBs, whose own cutover tooling deletes it), a Microsoft techcommunity thread, the fleetdm issue tracker, and several comments on the canonical teardown blog post.

| Value | Reported effect |
|---|---|
| `2` | auto-enrollment fails with `Auto MDM Enroll: Device Credential (0x1), Failed (Bad request (400).)` / `0x80190190` |
| `0`, or value deleted | enrollment succeeds |

> **Note the polarity.** `2` is the blocking state; `0` is healthy. Setting the value to `0` does **not** suppress enrollment — it is the fix, not the cause.

Three caveats. The meaning of the values is not documented and the community does not know it either. The "accepted answer" on the techcommunity thread asserts that `2` means *successfully enrolled* — it is contradicted by that thread's own author and by every other report, and is ignored here. And fleetdm reports devices where the value is already `0` or absent and enrollment still fails, so this is **a signal, not a diagnosis**.

The detector reports the raw value and raises the flag. It never writes it — clearing it is a separate, deliberate step, and it is machine-wide rather than per-enrollment so the remediation script leaves it alone too.

## Output

A `[pscustomobject]` (or compressed JSON with `-Json`):

| Property | Description |
|---|---|
| `ComputerName` | NetBIOS name |
| `Verdict` / `Flags` / `Reason` | verdict, every condition that applied, and a one-sentence explanation |
| `JoinedTenantId` / `ExpectedTenantId` | the two sides of the comparison |
| `JoinedUserEmail` / `EntraDeviceId` | joining identity and Entra device ID |
| `AzureAdJoined` / `DomainJoined` | both `YES` = hybrid |
| `MdmUrl` | enrollment discovery URL advertised by `dsregcmd` |
| `JoinInfoKeyCount` | `>1` means a leftover `JoinInfo` key from the old tenant |
| `EnrollmentCount` | number of `MS DM Server` enrollments found |
| `ActiveEnrollmentId` / `ActiveTenantId` | the live Intune enrollment and the tenant it belongs to |
| `OrphanEnrollmentCount` | enrollments that are not the live one — cleanup residue |
| `StaleEnrollmentCount` | `1` when the active Intune enrollment is cross-tenant, else `0` |
| `MmpcEnrollmentCount` | number of `Microsoft Device Management` enrollments |
| `MmpcActiveId` / `MmpcActiveTenantId` | the live declared-configuration enrollment and its tenant |
| `MmpcTenantMismatch` | `$true` when that tenant is not the baseline |
| `MmpcActiveUpn` | its UPN — the domain corroborates the tenant read |
| `MmpcActiveLastRun` | most recent `EnterpriseMgmt` task run for that enrollment |
| `MmpcEnrollmentFlag` | raw `Enrollments!MmpcEnrollmentFlag`; `2` is the reported blocking state, `0` healthy, `$null` absent |
| `LinkedEnrollmentId` | the GUID the **active Intune enrollment explicitly points at** as its declared-configuration child, read from `Enrollments\<guid>\LinkedEnrollment`. A hard pointer, not an inference. |
| `LinkedEnrollStatus` / `LinkedEnrollStatusText` | raw `EnrollStatus` (0–8) and its decode using Microsoft's published mapping (`3` = Enrollment Failed, `4` = Enrollment Succeeded, `8` = UnEnrollment Succeeded). **Not** a liveness signal — see above. |
| `LinkedLastError` | `LastError` from the same key, `0` when clean |
| `LinkedMmpcLocked` | raw `MMPCLocked` — reported, not interpreted |
| `LinkedDiscoveryEndpoint` | the declared-configuration discovery URL; a `dm-beta` hostname indicates a preview ring |
| `Enrollments` | per-enrollment detail, each carrying `ProviderId`, `Channel`, `LivenessScore` (0–9), `IsActive`, `DiscoveryUrlMatch`, `Upn`, `CertNotAfter`, `SyncTaskLastRun`, … |
| `CollectedUtc` | ISO-8601 collection timestamp |

## Detection usage

Run **as SYSTEM or elevated**. The values under `HKLM\SOFTWARE\Microsoft\Enrollments` are not readable by a standard user, and without them the script cannot distinguish `Clean` from `NotEnrolled`.

### Single-device triage

```powershell
.\Get-StaleMdmEnrollment.ps1
```

### Drill into the enrollment detail

Most useful when `EnrollmentCount > 1`:

```powershell
.\Get-StaleMdmEnrollment.ps1 | Select-Object -ExpandProperty Enrollments |
    Format-List EnrollmentId, IsActive, LivenessScore, EnrollmentTenantId, Upn, SyncTaskLastRun
```

### Flag a known old tenant

```powershell
.\Get-StaleMdmEnrollment.ps1 -StaleTenantId 00000000-1111-2222-3333-444444444444
```

### Fleet collection

```powershell
.\Get-StaleMdmEnrollment.ps1 -Json |
    Set-Content "\\fileserver\collect\$env:COMPUTERNAME.json" -Encoding UTF8
```

Aggregate afterwards:

```powershell
Get-ChildItem \\fileserver\collect\*.json |
    ForEach-Object { Get-Content $_ -Raw | ConvertFrom-Json } |
    Where-Object { $_.Verdict -ne 'Clean' } |
    Select-Object ComputerName, Verdict, @{n='Flags';e={$_.Flags -join '+'}},
                  JoinedTenantId, ActiveTenantId, MmpcActiveTenantId |
    Export-Csv .\stale-enrollments.csv -NoTypeInformation
```

### Minimum viable check

For copy-paste triage, without the liveness scoring or the orphan handling. Must run elevated. Note it checks **both** channels — a version that reads only `MS DM Server` returns a misleading `NOT-ENROLLED` on exactly the devices you are looking for:

```powershell
$joined = (dsregcmd /status |
    Select-String '^\s*TenantId\s*:\s*(\S+)' |
    Select-Object -First 1).Matches[0].Groups[1].Value

$all = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Enrollments' |
    ForEach-Object { Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue }

$intune = $all | Where-Object { $_.ProviderID -eq 'MS DM Server' }                | Select-Object -First 1
$windc  = $all | Where-Object { $_.ProviderID -eq 'Microsoft Device Management' } | Select-Object -First 1

$flags = @()
if (-not $intune -and -not $windc)                 { $flags += 'NOT-ENROLLED' }
if ($windc -and -not $intune)                      { $flags += 'MMPC-ORPHAN' }
if ($intune -and $intune.AADTenantID -ne $joined)  { $flags += 'STALE-INTUNE' }
if ($windc  -and $windc.AADTenantID  -ne $joined)  { $flags += 'STALE-WINDC' }
if (-not $flags)                                   { $flags = @('OK') }

'{0} | joined={1} | intune={2} | windc={3} | {4}' -f $env:COMPUTERNAME, $joined,
    $intune.AADTenantID, $windc.AADTenantID, ($flags -join '+')
```

A healthy device returns `OK` with both tenants equal to `joined`. A migrated, retired device returns `MMPC-ORPHAN+STALE-WINDC` with an empty `intune=`.

> The `$intune -and` / `$windc -and` guards are load-bearing. Without them, `$null -ne '<guid>'` evaluates to `$true` and every unenrolled device reports as stale. String `-ne` is case-insensitive by default, which usefully absorbs GUID-casing differences between `dsregcmd` and the registry — do not "fix" it to `-cne`.

## Interpreting the output

- **`StaleKnownTenant` with a recent `SyncTaskLastRun`** → the device is still actively syncing to the old tenant. It is manageable from there right now.
- **`StaleKnownTenant` with `CertPresent = False` or an expired `CertNotAfter`** → the old channel is degrading. It will eventually fail on its own, but the enrollment will not clean itself up.
- **`MmpcOrphanCrossTenant`** → the device is stuck. It cannot re-enrol (Windows thinks it already is) and it cannot renew (the old tenant's IdP will not authenticate it). Nothing will change without intervention. This is the state that produces *no* enrollment events at all, which is easily mistaken for "nothing is being attempted" — correct, and that is the symptom, not the cause.
- **`Clean` with `MmpcCrossTenant` in `Flags`** → Intune is healthy but the declared-configuration channel is still bound to the old tenant. Not yet service-affecting; will become `MmpcOrphanCrossTenant` if the Intune enrollment is ever retired.
- **`Clean` with `OrphanEnrollmentCount > 0`** → the device is correctly enrolled, but a previous cleanup left residue behind. Cosmetic unless the orphan later wins the liveness ranking.
- **`JoinInfoKeyCount > 1`** → leftover `CloudDomainJoin` residue from the pre-migration tenant. Independent corroboration that this device went through a migration.
- **`NotEnrolled` with `EnrollmentCount > 0`** → a teardown ran but did not complete. The device is managed nowhere.

# Remediation — `Remove-StaleMdmEnrollment.ps1`

> **Microsoft does not document or support this teardown.** Not the key list, not any part of it. A device repaired this way is in a state Microsoft has not sanctioned. **The supported remediation is a device reset.** Read the next section before using it.

## Why there is no supported alternative

Every documented way to remove an enrollment is **server-initiated** and needs a working enrollment to arrive through:

- Intune **Retire** / **Wipe** from the console.
- The DMClient CSP `Provider/{ProviderID}/Unenroll` Exec node.
- The DMClient CSP **`LinkedEnrollment/Unenroll`** Exec node — the specific supported teardown for a declared-configuration enrollment, which rolls back all WinDC settings cleanly.

Once the primary enrollment is gone, the server cannot reach the device at all, so none of these can be delivered. And the CSP cannot be driven locally either:

- The DMClient CSP is **not projected into the WMI Bridge**. Verified by enumerating all **465** classes in `root\cimv2\mdm\dmmap` as `NT AUTHORITY\SYSTEM` on Windows 11 26100 — no `MDM_DMClient*` class exists, and nothing matches `declared` or `linked`. The namespace is fully populated, so this is not a permissions artefact.
- This is consistent with the CSP documentation noting that DMClient custom URI settings "are not supported for IT admin management scenarios".

**That leaves exactly two options: reset the device, or purge by hand.**

| | Supported | When to use |
|---|---|---|
| **Autopilot Reset / Fresh Start / Windows reset / reimage**, then re-enrol clean | ✅ Yes | Default choice. Always viable, always defensible. |
| Manual registry + scheduled-task purge (this script) | ❌ **No** | Only when a reset is not viable — an irreplaceable device, a user who cannot afford the rebuild window, or a fleet too large to reset. Be prepared to reset anyway if it does not take. |

When the primary MDM enrollment has already been retired, the server can no longer reach the device to send the Exec either. The enrollment therefore **cannot be removed by any documented route**. Manual purge or reimage are the only remaining options; this script is the manual purge.

For completeness: a local trigger for the **enroll** direction does exist — `deviceenroller.exe /c /EnrollMmpc` (community-reported, undocumented). There is no counterpart for unenroll. No Microsoft page documents *any* `deviceenroller.exe` switch, including the widely-used `/c /AutoEnrollMDM`; a search for a `/DisenrollDevice` switch returns zero results in any source, and a scan of the binary's strings on build 26100 finds no `Disenroll` token. **It is not real.**

## What it removes

For one enrollment GUID:

- Scheduled tasks under `\Microsoft\Windows\EnterpriseMgmt\<guid>\`, then the task folder itself (the folder requires the `Schedule.Service` COM object — there is no `Unregister-ScheduledTaskFolder` cmdlet).
- Eight registry roots, all keyed by the **dashed** GUID:

  ```
  SOFTWARE\Microsoft\Enrollments\<guid>
  SOFTWARE\Microsoft\Enrollments\Status\<guid>
  SOFTWARE\Microsoft\EnterpriseResourceManager\Tracked\<guid>
  SOFTWARE\Microsoft\PolicyManager\AdmxInstalled\<guid>
  SOFTWARE\Microsoft\PolicyManager\Providers\<guid>
  SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\<guid>
  SOFTWARE\Microsoft\Provisioning\OMADM\Logger\<guid>
  SOFTWARE\Microsoft\Provisioning\OMADM\Sessions\<guid>
  ```

  All eight use the **dashed** form — verified on Windows 11 26100 by probing both variants for both the Intune and the declared-configuration enrollment: dashed matched everywhere, undashed matched nothing. The undashed fallback in the code is **defensive only**; no source, Microsoft or community, was found asserting that any build ever used it, and an earlier version of this document claimed older builds did — that claim was unsourced and is withdrawn.

  Two notes on this list. `PolicyManager\AdmxInstalled` and `PolicyManager\Providers` legitimately **do not exist** for a declared-configuration enrollment (verified) — absent is normal there, not a failure. And **Microsoft documents none of this.** The list traces to a single source — Maxime Rastello, [*Manually re-enroll a co-managed or Hybrid Azure AD Join Windows 10 PC to Microsoft Intune without loosing current configuration*](https://www.maximerastello.com/manually-re-enroll-a-co-managed-or-hybrid-azure-ad-join-windows-10-pc-to-microsoft-intune-without-loosing-current-configuration/) (2020) — which every later script and article copies; [steve-prentice's `Remove-IntuneCurrentEnrollment.ps1`](https://github.com/steve-prentice/powershell-scripts/blob/master/Remove-IntuneCurrentEnrollment.ps1) and [ztrhgf's `Reset-IntuneEnrollment.ps1`](https://github.com/ztrhgf/useful_powershell_functions/blob/master/INTUNE/Reset-IntuneEnrollment.ps1) both credit it by name. It is community consensus, **not** eight independent confirmations. Microsoft's own [enrollment diagnostics article](https://learn.microsoft.com/en-us/windows/client-management/mdm-diagnose-enrollment) names only the first key and offers a heuristic rather than a list — *"if you don't know which registry key to remove, go for the key that displays most entries"*. `PolicyManager`, `EnterpriseResourceManager` and `OMADM` are never mentioned in a teardown context anywhere on Learn.

  > **Ordering matters if you are purging both channels.** `Enrollments\<guid>` is deleted recursively, which takes the `LinkedEnrollment` subkey — the only explicit pointer to the declared-configuration child — with it. Record `LinkedEnrollmentId`, or purge the child GUID, **first**.
- The MDM client certificate in `Cert:\LocalMachine\My` named by `DMPCertThumbPrint` (skip with `-SkipCertificate`).

It does **not** touch `MmpcEnrollmentFlag`: that value is machine-wide rather than per-enrollment, so it is out of scope for a single-GUID purge. Clear it separately if the detector flags it.

Tasks are removed **first**, so an OMA-DM sync cannot fire mid-purge and rewrite keys that have just been deleted.

## Safety model

| Guard | Behaviour |
|---|---|
| Dry run by default | Nothing is touched without `-Execute`. Safe to hand to a customer for a first pass. |
| Elevation check | Refuses to run unelevated. |
| Real-channel check | Refuses any key whose `ProviderID` is not one of the two real MDM channels, so a mistyped GUID cannot destroy an internal CSP-provider subkey. |
| Healthy-enrollment check | Refuses an enrollment whose `AADTenantID` matches the joined tenant, unless `-Force`. This is the guard against running it on the wrong machine. |
| Backup | Every registry root (`reg.exe export`), every task definition (XML) and the certificate (public `.cer`) are written to `-BackupPath` with a `manifest.json` **before** anything is deleted. |
| `ShouldProcess` | Every destructive step supports `-WhatIf` and `-Confirm`. |

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `EnrollmentId` | `string` | — **required** | Enrollment GUID to remove, as reported in `ActiveEnrollmentId` or `MmpcActiveId`. Dashed form. |
| `BackupPath` | `string` | `%ProgramData%\MdmEnrollmentPurge` | Backup root. A `<guid>_<timestamp>` subfolder is created per run. |
| `Execute` | `switch` | Off | Actually perform the removal. Without it the script reports the plan and exits. |
| `SkipCertificate` | `switch` | Off | Leave the MDM client certificate in place — use when a second enrollment you are keeping references the same cert. |
| `Force` | `switch` | Off | Permit removal of an enrollment whose tenant matches the join. Only for deliberate full re-enrollment. |
| `Reenroll` | `switch` | Off | After a clean purge, trigger `deviceenroller.exe /c /AutoEnrollMDM`. Self-suppresses if any removal failed. |

## Usage

Run **elevated**; `SYSTEM` is preferable (see caveats).

```powershell
# 1. Detect
.\Get-StaleMdmEnrollment.ps1 | Select-Object Verdict, Flags, MmpcActiveId, MmpcActiveTenantId

# 2. Dry run against the GUID it reported
.\Remove-StaleMdmEnrollment.ps1 -EnrollmentId 00000000-1111-2222-3333-444444444444

# 3. Purge
.\Remove-StaleMdmEnrollment.ps1 -EnrollmentId 00000000-1111-2222-3333-444444444444 -Execute

# 4. Reboot, then confirm
.\Get-StaleMdmEnrollment.ps1 | Select-Object Verdict, Flags
```

Unattended, once validated on a pilot device:

```powershell
.\Remove-StaleMdmEnrollment.ps1 -EnrollmentId <guid> -Execute -Confirm:$false -Reenroll
```

## Caveats

- **A partial purge is worse than none.** Windows can still consider itself enrolled off a single surviving root. Always check the returned `FailedCount` is `0`, and re-run the detector to confirm the verdict moved to `NotEnrolled`.
- **Some keys may resist deletion as Administrator.** Subkeys under `EnterpriseResourceManager\Tracked` and `PolicyManager` can carry ACLs that deny Administrators write access. Re-run as SYSTEM (`PsExec -s -i`) before resorting to taking ownership — taking ownership of policy keys has its own side effects.
- **Reboot before evaluating re-enrollment.** The OMA-DM client caches enrollment state in memory.
- **Fix the enrollment path first.** If auto-enrollment authenticates against a federated IdP belonging to the *old* tenant, re-enrollment fails after a perfectly clean purge. Confirm a comparable device enrols successfully before purging at scale.

## Related

- [`../EntraCutover/`](../EntraCutover/README.md) — the Teardown phase implements a broader enrollment purge as part of a full cutover.
- [`../../../intune/mdm-enrollment/Repair-IntuneMdmCert.ps1`](../../../intune/mdm-enrollment/Repair-IntuneMdmCert.ps1) — teardown and re-enrollment for a broken MDM certificate.

## Sources & confidence

Every claim below was re-validated against Microsoft Learn and community sources rather than carried forward from earlier drafts. Where revalidation contradicted an earlier claim, the correction is called out rather than quietly applied.

### Verified on live Windows 11 hosts (our own measurement)

- The complete value set under `Enrollments\<guid>`, dumped on both channels: `ProviderID`, `AADTenantID`, `AADResourceID`, `UPN`, `DMPCertThumbPrint`, `EnrollmentType`, `EnrollmentState`, `EnrollmentFlags`, `IsFederated`, `DiscoveryServiceFullURL`, `CurCryptoProvider`, `RenewStatus`, and the cert-chain thumbprints.
- `ProviderID = 'MS DM Server'` isolating the Intune enrollment, and the full set of other `ProviderID` values present, including the pseudo-enrollments that carry no tenant.
- `EnrollmentType = 6` on the Intune channel, `26` (`0x1a`) on the declared-configuration channel.
- `SslClientCertReference` living under `Provisioning\OMADM\Accounts\<guid>` as `MY;System;<thumbprint>`, **populated**, and matching `DMPCertThumbPrint`.
- All eight teardown roots present under the **dashed** GUID; the undashed variant matching nothing.
- `PolicyManager\AdmxInstalled` and `PolicyManager\Providers` absent for the declared-configuration enrollment while present for the Intune one.
- A `Microsoft Device Management` enrollment surviving alone with its full task set, carrying an `AADTenantID` for a **different** tenant than the device is joined to, with a matching old-domain UPN and `IsFederated = 1`.
- Two healthy channels coexisting with **different** enrollment GUIDs but the **same** `AADTenantID`.
- `LinkedEnrollment` under the Intune enrollment key, with `LinkedEnrollmentId` resolving exactly to the declared-configuration enrollment GUID.
- `EnrollStatus = 3` ("Enrollment Failed") with `LastError = 0` on a host whose declared-configuration enrollment is demonstrably live — the basis for refusing to use it as a liveness signal.
- The `user@domain.com@<tenantGuid>` UPN form used as the fallback tenant read.
- The MDM certificate CN **not** matching the Entra device ID.
- `root\cimv2\mdm\dmmap` containing 465 classes and **no** `MDM_DMClient*` class, as SYSTEM.

### Documented on Microsoft Learn

- The `{ProviderID}` concept — *"the URI-encoded value of the bootstrapped device management account's Provider ID… set and controlled by the MDM server"* — [DMClient CSP](https://learn.microsoft.com/en-us/windows/client-management/mdm/dmclient-csp).
- The `LinkedEnrollment` nodes (`DiscoveryEndpoint`, `Enroll`, `Unenroll`, `EnrollStatus`, `LastError`) and the **`EnrollStatus` 0–8 mapping** — same page. `DiscoveryEndpoint` is flagged there as *Windows Insider Preview*.
- That `Enroll` fails with `ERROR_FILE_NOT_FOUND (0x80070002)` and *"there is no scheduled task created for dual enrollment"* when `DiscoveryEndpoint` is unset — same page. This is the documented link between the enrollment and its `EnterpriseMgmt` task set.
- *"This dual enrollment is only allowed if the device is already enrolled into a primary mobile device management (MDM) server"* — [Windows declared configuration](https://learn.microsoft.com/en-us/windows/client-management/declared-configuration).
- The literal strings `Enrollment Name: (MicrosoftManagementPlatformCloud)` and `Enroll Type: (0x1A)` — same page, in its troubleshooting examples.
- *"Can there be more than one MDM server…? No. Only one MDM is allowed."* — [MDM overview](https://learn.microsoft.com/en-us/windows/client-management/mdm-overview).

### Community-sourced, not Microsoft-documented

| Claim | Status |
|---|---|
| `ProviderID` is literally `MS DM Server` for Intune | Corroborated by community reports and a Citrix support article. Stable and universally relied upon, but **no Learn page states it**. A third-party MDM sets its own value. |
| The eight teardown registry roots | Community consensus, but tracing to a **single** origin post — [Rastello, 2020](https://www.maximerastello.com/manually-re-enroll-a-co-managed-or-hybrid-azure-ad-join-windows-10-pc-to-microsoft-intune-without-loosing-current-configuration/) — not eight independent confirmations. All eight confirmed present on our hosts. **Microsoft publishes no teardown key list and does not support this procedure.** |
| `deviceenroller.exe /c /AutoEnrollMDM`, `/c /AutoEnrollMDMUsingAADDeviceCredential`, `/c /EnrollMmpc` | Widely used and reproducible, **explicitly undocumented** — one community author notes the binary ships with no help output at all. |
| `MmpcEnrollmentFlag` semantics | Multiple independent reports (vendor KBs, MS forum thread, issue tracker, blog comments) that `2` blocks and `0`/deleted unblocks. The **meaning** of the values is unknown to everyone, including the people who found the fix. |
| `MMPCLocked` | Name confirmed on our host and in community write-ups. Meaning **unknown** — the only description of it is explicitly labelled a guess by its author. |
| Numeric `EnrollmentType` mapping (6, 26, …) | Values observed; **no mapping table exists** in any source. `0x1A` appears in one Microsoft log sample without explanation. |

### Not verified anywhere — treat as hypothesis

- **That an Intune Retire is what leaves the declared-configuration enrollment behind.** No source, Microsoft or community, states this. It is our inference from one device's end state.
- That an orphaned declared-configuration enrollment is what makes `deviceenroller.exe /c /AutoEnrollMDM` exit without logging. The silence plus the live enrollment is the measurement; the precondition check itself is not documented.
- That a device can hold more than one `MS DM Server` enrollment simultaneously. Windows keeps one active enrollment; orphan survival is inferred from the fact that manual teardown must clear eight roots by hand. The script handles it defensively rather than assuming it.
- `AADTenantID`, `EnrollmentState` and `IsFederated` as *named, contracted* values. All three are present and behave as described on our hosts, but no Microsoft or community source documents them — Microsoft has never published a schema for `HKLM\SOFTWARE\Microsoft\Enrollments\<guid>`. Where a value name matches a documented CSP node, that is the **CSP node** being documented, not the registry value.
- The liveness weights are a judgement call, not a Microsoft-defined ranking.

## Requirements

- Windows PowerShell **5.1**. No PowerShell 7+ syntax is used.
- Run **as SYSTEM or elevated** — `HKLM\SOFTWARE\Microsoft\Enrollments` values are not readable otherwise. SYSTEM is preferable for the remediation script (see [Caveats](#caveats)).
- `Get-StaleMdmEnrollment.ps1` is read-only: it creates, modifies and deletes **nothing**.
- `Remove-StaleMdmEnrollment.ps1` is destructive, and only with `-Execute`. Without that switch it reports its plan and exits.
