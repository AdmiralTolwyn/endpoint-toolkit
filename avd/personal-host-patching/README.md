# AVD Personal Host Pool Nightly Patching (Hotpatch-First & Hibernation-Safe)

[Invoke-AvdPersonalNightlyPatching.ps1](Invoke-AvdPersonalNightlyPatching.ps1) orchestrates nocturnal Windows Update, **Windows 11 24H2+ Hotpatch**, Microsoft Defender Antivirus signature updates, and Intune OMA-DM sync across **Azure Virtual Desktop (AVD) Personal** session hosts—without triggering **60-minute Autoscale Hibernation races** or wiping hibernated developer sessions via cold `deallocate`.

---

## 1. Why Personal AVD + Hibernation Needs Purpose-Built Servicing

1. **Azure Update Manager Doesn't Support Windows 10/11 Client OS:**
   * Per [Unsupported workloads for Azure Update Manager](https://learn.microsoft.com/en-us/azure/update-manager/unsupported-workloads), Windows 10/11 client images are unsupported in AUM; Microsoft recommends **Microsoft Intune / Windows Autopatch** ([Windows Autopatch on AVD workloads](https://learn.microsoft.com/en-us/windows/deployment/windows-autopatch/deploy/windows-autopatch-register-devices#windows-autopatch-on-azure-virtual-desktop-workloads) supports *Personal persistent virtual machines*).
   * However, Intune/Autopatch cannot wake a `Hibernated` or `Deallocated` Azure VM, leaving hibernated desktops to resume at 08:30 AM with stale Defender definitions and unapplied Cumulative Updates.
2. **AVD Autoscale's 60-Minute Hibernate Race (Fail-Closed Exclusion Gate):**
   * Per [AVD Autoscale FAQ](https://learn.microsoft.com/en-us/azure/virtual-desktop/autoscale-faq#what-happens-to-session-hosts-that-get-turned-on-but-dont-ever-get-signed-into), if an automation job turns on a Personal VM and nobody signs in (or a `Disconnected` session exceeds the Off-Peak timer), Autoscale **hibernates or deallocates the VM mid-install**.
   * This script resolves the active Scaling Plan's `properties.exclusionTag`—**failing closed** if no enabled Scaling Plan is visible (unless `-ScalingPlanResourceId`, `-ExclusionTagName`, or `-AllowNoScalingPlan` is explicitly passed) or if `exclusionTag` is empty/mismatched—stamps `<exclusionTag>` + `AvdPatchLock` *before* waking the VM, and performs **ownership-conditional, case-insensitive tag deletion** only after final power state is verified (preserving pre-existing or mid-run operator tags such as `OpsPin`).
3. **Windows 11 24H2+ Hotpatching + Hibernation Synergy:**
   * Per [Hotpatch updates in Windows Autopatch](https://learn.microsoft.com/en-us/windows/deployment/windows-autopatch/manage/windows-autopatch-hotpatch-updates) and [Hotpatch release notes](https://support.microsoft.com/en-us/servicing/os/hotpatch/windows-11/2025/release-notes-for-hotpatch-on-windows-11-enterprise-version-24h2), Windows 11 Enterprise 24H2+ devices on the current quarterly baseline with **VBS running** and `AllowRebootlessUpdates = 1` receive **rebootless Monthly B security updates in 8 out of 12 months** (`Feb, Mar, May, Jun, Aug, Sep, Nov, Dec`), requiring an OS restart in the **4 quarterly Baseline months** (`Jan, Apr, Jul, Oct`).
   * **Session-Preserving Servicing Rule:** When a `Disconnected` user session is preserved in RAM (within `-DisconnectedSessionGraceDays`, default `3` days), the guest script installs only rebootless updates (`InstallationRebootBehavior == 0` such as Defender definitions, plus Monthly B `Hotpatch` updates on `VBS + AllowRebootlessUpdates` hosts), defers reboot-requiring Baseline LCUs, triggers Intune `PushLaunch` sync, **skips OS restart**, and restores `HibernationState/Hibernated` via `POST .../deallocate?hibernate=true`.
   * **Platform Caveats:** Reboots still occur on 0-session hosts when `PreExistingRebootPending` (`CBS\RebootPending` / `WU\RebootRequired`) is already set, when quarterly Baseline LCUs or `.NET` updates require a restart, or once `-DisconnectedSessionGraceDays` expires. Note that Azure Compute itself can occasionally cold-boot a hibernated VM during host hardware maintenance (`ComponentStatus/VMBootState/Start/succeeded`, surfaced as a warning in `Detail`).

---

## 2. Architecture & Safety Guarantees

```text
Phase 0: Discovery, Sibling Guards & Self-Healing Lock Check
  ├── Verifies hostPoolType == 'Personal'
  ├── Validates active Scaling Plan properties.exclusionTag (fails closed if 0 plans visible or tag empty/mismatched)
  ├── Honors AvdLegalHold (any non-empty value except false/0/no) & AvdLifecycleState == 'Quarantined'
  ├── Self-heals expired idle AvdPatchLock tags (prunes finished Avd* RunCommands & carries forward original PriorState);
  │   leaves expired locks with still-Running RunCommands untouched as NeedsOperatorTriage
  └── Filters by deterministic SHA-256 Cohort (-CohortCount / -CohortIndex) & smart catch-up (-CatchUpSkipRecentHours)

Phase 1–2: Paced Sliding Window (-MaxConcurrentVMs 60, -MaxStartsPerMinute 15)
  ├── Refreshes live Broker userSessions & VM InstanceView at window admission (eliminating stale queued snapshots)
  ├── Merges <exclusionTag>='AvdNightlyPatch' + AvdPatchLock (reserving 1 slot below ARM's 50-tag ceiling for AvdPatchState)
  ├── Sets sessionHost allowNewSession = false (Drain Mode, when originally true)
  └── Wakes VM (POST /start); catches ZonalAllocationFailed/AllocationFailed without touching hibernated disk state

Phase 3: Unified In-Guest SYSTEM Servicing (Managed Run Command AvdPatch-<shortRun>)
  ├── Prunes terminal Avd* RunCommands first so VMs never approach Azure's 25-Managed-Run-Command quota
  ├── Authoritative Dual-Gate Session Check: Broker userSessions UNION Unicode Win32 WTSEnumerateSessionsW + HKEY_USERS/ProfileList hives
  ├── Refreshes Microsoft Defender signatures (Update-MpSignature, surfacing any failure in Detail)
  ├── Honors PauseQualityUpdates; queries WUA ssDefault (BrowseOnly=0, blocking Upgrades 25H2/26H2 & Feature Packs)
  ├── Session Present (within grace): installs ONLY rebootless updates (RebootBehavior == 0 or eligible Monthly B Hotpatch)
  ├── Zero Sessions (or grace expired): installs all approved Security/Critical/Rollup/Definition/Update KBs
  ├── TimedOut Guard: if RunCommand hits timeoutInSeconds, leaves VM Running (NeedsOperatorTriage) to never cut off TiWorker
  └── Triggers Intune PushLaunch & polls up to 90s for omadmclient.exe completion before emitting AVDPATCH_JSON

Phase 4: Conditional Guest-Coordinated Reboot (Strictly when RebootRequired / RebootPending)
  ├── Registers expiring task AvdNightlyPatchOneShotReboot (+20s, DeleteExpiredTaskAfter=PT10M, NO StartWhenAvailable)
  ├── Re-checks active sessions & loaded hives at T+20s execution time right before calling shutdown.exe /r /t 5
  └── Verifies LastBootUpTime > BaselineBootUtc + UBR + OMA-DM sync, then waits -PostRebootSettleSeconds (300s)

Phase 5: Power Restoration & Ownership-Conditional Unlock (+ Independent 05:15 UTC MorningUnlock)
  ├── Restores Hibernated (POST deallocate?hibernate=true; NEVER falls back to hibernate=false!), Deallocated, or Running
  ├── Restores prior allowNewSession and deletes AvdPatchLock + <exclusionTag> ONLY if values still match this run
  └── Persists AvdPatchState (last UTC timestamp, UBR, outcome, and deferral start epoch)
```

---

## 3. Required Permissions & Pipeline Prerequisites

### Custom Least-Privilege RBAC Role (`az role definition create`)

Do **not** assign built-in `Virtual Machine Contributor` (which grants `virtualMachines/delete` and `disks/*`). Create a scoped custom role (replacing `<SUBSCRIPTION_ID>`):

```json
{
  "Name": "AVD Personal Host Nightly Patch Orchestrator",
  "IsCustom": true,
  "Description": "Reads AVD host pools and scaling plans, toggles session host drain mode, manages VM tags, starts/stops/hibernates VMs, and executes Managed Run Commands without VM or disk deletion rights.",
  "Actions": [
    "Microsoft.DesktopVirtualization/hostpools/read",
    "Microsoft.DesktopVirtualization/hostpools/sessionhosts/read",
    "Microsoft.DesktopVirtualization/hostpools/sessionhosts/write",
    "Microsoft.DesktopVirtualization/hostpools/sessionhosts/usersessions/read",
    "Microsoft.DesktopVirtualization/scalingplans/read",
    "Microsoft.Compute/virtualMachines/read",
    "Microsoft.Compute/virtualMachines/instanceView/read",
    "Microsoft.Compute/virtualMachines/start/action",
    "Microsoft.Compute/virtualMachines/deallocate/action",
    "Microsoft.Compute/virtualMachines/powerOff/action",
    "Microsoft.Compute/virtualMachines/runCommands/read",
    "Microsoft.Compute/virtualMachines/runCommands/write",
    "Microsoft.Compute/virtualMachines/runCommands/delete",
    "Microsoft.Compute/locations/operations/read",
    "Microsoft.Resources/tags/read",
    "Microsoft.Resources/tags/write"
  ],
  "NotActions": [],
  "DataActions": [],
  "NotDataActions": [],
  "AssignableScopes": [
    "/subscriptions/<SUBSCRIPTION_ID>"
  ]
}
```

#### RBAC Assignment Scope Options & Security Notes
* **Option A — Resource-Group-Only Assignment (Strictest Least Privilege):** Assign the custom role strictly on the **Host Pool RG**, **Scaling Plan RG**, and **Session Host Compute RG(s)**, and populate `scalingPlanResourceIds` (`-ScalingPlanResourceId`). When subscription-level `Microsoft.Compute/locations/operations/read` is omitted, `Get-AvdPatchAsyncOperationStatus` automatically degrades on HTTP `403` to VM `instanceView` polling.
* **Option B — Auto-Discovery Across RGs:** Assign the custom role on the Host Pool and Compute RGs, plus `scalingPlans/read` and `locations/operations/read` at Subscription scope so Scaling Plans in separate platform RGs are discovered automatically.
* **Privilege Boundary Note:** `Microsoft.Compute/virtualMachines/runCommands/write` grants `NT AUTHORITY\SYSTEM` code execution on VMs in the target Compute RG, and `Microsoft.Resources/tags/write` can modify VM tags (`AvdLegalHold`, `AvdLifecycleState`) in that RG. Use a dedicated Managed Identity per environment rather than sharing identities across unrelated workloads.

### Azure DevOps Pipeline Prerequisites ([avd-nightly-personal-patching.yml](../pipelines/avd-nightly-personal-patching.yml))

1. **Agent Pool & Job Timeout:** Free Microsoft-hosted `windows-latest` jobs terminate at **60 minutes**. Leave `useSelfHostedPool: true` (`<YOURSELFHOSTEDPOOL>`, up to 295 min) for production fleets, or ensure paid parallel jobs are enabled before using hosted agents.
2. **Workload Identity Federation (>55 min runs):** Per [Troubleshoot workload identity service connections](https://learn.microsoft.com/en-us/azure/devops/pipelines/release/troubleshoot-workload-identity), Azure DevOps service connections using Workload Identity Federation backed by an **App Registration** hit `AADSTS700024` token-refresh failures after ~60 minutes. Back `<YOURSERVICECONNECTION>` with a **User-Assigned Managed Identity** (~24-hour token lifetime) for runs exceeding 55 minutes.
3. **Exclusive Lock Environment:** Create the Azure DevOps Environment `avd-personal-patching-prod` with an **Exclusive Lock** check. `lockBehavior: sequential` is scoped exclusively to the `PersonalHostServicing` stage, while `MorningDrainUnlockSafetyNet` declares `dependsOn: []` and no environment lock—guaranteeing immediate execution at `05:15 UTC` (or on manual `mode: MorningUnlock` runs) even if an earlier servicing run is hung.
4. **Companion Intune / Golden-Image Settings:**
   * **VM SKU & Image:** Trusted Launch + Nested Virtualization + Windows 11 Enterprise 24H2 build `>= 26100.3037` (required for VBS and Azure Hibernation to coexist per [Windows VM Hibernation docs](https://learn.microsoft.com/en-us/azure/virtual-machines/windows/hibernate-resume-windows)).
   * **Intune Quality Update Policy:** Set *"When available, apply without restarting the device ('Hotpatch')"* = **Allow** (and ensure devices install each quarterly Baseline LCU to stay Hotpatch-eligible).
   * **Update Ring Deadline Guard:** Set `ConfigureDeadlineNoAutoRebootForQualityUpdates = 1` so `UsoSvc` does not unilaterally reboot `Disconnected` sessions outside Active Hours.

---

## 4. Cohort Sizing & Cost Model

| Mode / Schedule | Wakes per VM / Month | Fleet Example (`1,200` VMs) | Notes |
| --- | --- | --- | --- |
| **`-CohortCount 5 -CohortIndex -1` *(Pipeline Default, Mon–Fri)*** | **~4.3× / month** *(~80% fewer wakes vs. daily Mon–Fri)* | **240 VMs / night** (`< 90 min` typical Hotpatch/Defender wave at `60` concurrent, `15` starts/min) | Deterministic SHA-256 assignment maps each host to 1 weekday (`Mon=0 .. Fri=4`), eliminating Friday full-fleet spikes. |
| **`-CohortCount 4 -CohortIndex -1` *(Mon–Thu + Smart Fri Catch-Up)*** | **~4.3–5.0× / month** | **300 VMs Mon–Thu** + **missed/deferred hosts only** on Friday | On Friday (`CatchUpSkipRecentHours = 96`), hosts whose `AvdPatchState` shows a clean run within 4 days return `SkippedCatchUpAlreadyHealthy` without waking. |
| **Orchestrator Compute** | **$0 incremental license** | Self-hosted Windows VM / runner | Native WUA COM (`Microsoft.Update.Session`) + ARM REST API. VM wake compute depends on SKU (e.g., ~15–20 min/week on `Standard_D4s_v5` PayG is ~€0.25–€0.45/VM/mo). |

---

## 5. Arguments

| Parameter | Type | Default | Description |
| --- | --- | --- | --- |
| `-HostPoolResourceId` | `string[]` | — | Full ARM Resource ID(s) of target AVD Personal Host Pools. |
| `-SubscriptionId` | `string` | Current context | Subscription ID when using `-ResourceGroupName` and `-HostPoolName`. |
| `-ResourceGroupName` | `string` | — | Resource Group containing the target AVD Personal Host Pool. |
| `-HostPoolName` | `string` | — | Name of the target AVD Personal Host Pool. |
| `-ScalingPlanResourceId` | `string[]` | — | Explicit Scaling Plan ARM ID(s) when subscription-wide `scalingPlans/read` is not granted. |
| `-ExclusionTagName` | `string` | — | Expected Scaling Plan `exclusionTag` name (validated against active plan or used when no plan is linked). |
| `-AllowNoScalingPlan` | `switch` | `$false` | Permits `Execute` on a Personal Host Pool with no enabled Scaling Plan (prevents silent fail-open). |
| `-Mode` | `string` | `Audit` | `Audit` (read-only), `Triage` (read-only lock/RunCommand inspection), `MorningUnlock`, `LocalGuestProbe` (local read-only guest test), or `Execute`. |
| `-ForceUnlockActive` | `switch` | `$false` | In `MorningUnlock`, unlocks/clears locks even if `ExpiresUtc` has not yet elapsed. |
| `-CohortCount` | `int` | `1` | Number of deterministic SHA-256 cohorts (`1..31`; pipeline uses `5`). |
| `-CohortIndex` | `int` | `0` | `0..CohortCount-1`, or `-1` for automatic weekday/daily rotation. |
| `-CatchUpSkipRecentHours` | `int` | `96` | Hours within which cleanly serviced hosts are skipped during catch-up windows (`0..336`). |
| `-MaxConcurrentVMs` | `int` | `60` | Maximum concurrent VMs active in the sliding window (`1..300`). |
| `-MaxStartsPerMinute` | `int` | `15` | Maximum VM wake (`POST /start`) calls per rolling 60 seconds (`1..120`). |
| `-DisconnectedSessionGraceDays` | `int` | `3` | Days to defer reboot-requiring Baseline LCUs when a Disconnected session is in RAM (`0..30`). |
| `-PostRebootSettleSeconds` | `int` | `300` | Post-reboot settle delay before calling `deallocate?hibernate=true` (`0..1800`). |
| `-PostRebootZeroSessionAction` | `string` | `Hibernate` | `Hibernate` or `Deallocate` after rebooting a 0-session host that started `Hibernated`. |
| `-StopNewWakeAtLocalHour` | `int` | `-1` | Local hour (`0..23` in `-MaintenanceTimeZoneId`, `-1` disabled) to stop new VM wakes. |
| `-MaintenanceTimeZoneId` | `string` | `W. Europe Standard Time` | Validated Windows `TimeZoneInfo` ID for `-StopNewWakeAtLocalHour`. |
| `-RunCommandTimeoutSeconds` | `int` | `2700` | Execution timeout for in-guest servicing RunCommand (`300..5400`). |
| `-LockTtlMinutes` | `int` | `180` | Minutes before `AvdPatchLock` expires (`30..720`). |
| `-MinFreeDiskSpaceGB` | `int` | `10` | Minimum free space on `C:` (GiB) required before WUA download (`4..100`). |
| `-MaxFailurePercent` | `int` | `20` | Circuit-breaker failure percentage threshold (`1..100`). |
| `-MaxFailureCount` | `int` | `5` | Minimum host failures before `MaxFailurePercent` trips the circuit breaker (`1..500`). |
| `-IncludeDrivers` | `switch` | `$false` | Permits WUA `Drivers` category updates when set. |
| `-IncludePreviewUpdates` | `switch` | `$false` | Permits `Preview`-titled cumulative updates when set. |
| `-IncludeQuarantined` | `switch` | `$false` | Permits servicing VMs tagged `AvdLifecycleState=Quarantined`. |
| `-TargetSessionHostName` | `string[]` | — | Optional short hostname filter for canary/targeted runs. |
| `-ManagedIdentityClientId` | `string` | — | Client ID when authenticating with a User-Assigned Managed Identity. |
| `-AuthMode` | `string` | `Auto` | `Auto`, `ManagedIdentity`, or `AzContext`. |

---

## 6. VM Tag Contract & Sibling Interoperability

| Tag Name | Lifecycle | Format / Values | Purpose |
| --- | --- | --- | --- |
| `<exclusionTag>` *(e.g. `exclude-sp-pers-01`)* | Transient (or operator-pinned) | `'AvdNightlyPatch'` (when stamped by this script) | Prevents AVD Autoscale from hibernating/deallocating the VM mid-install. Pre-existing or mid-run operator values (e.g. `OpsPin`) are preserved via case-insensitive ownership checks. |
| `AvdPatchLock` | Transient (`LockTtlMinutes`) | `v2\|run=<guid>\|vm=<vmId>\|ps=<state>\|ans=<0\|1>\|pex=<0\|1>\|tag=<excl>\|exp=<epoch>` | Single-writer lock + crash-recovery journal. Honored reciprocally by [Remove-InactiveAvdPersonalHosts.ps1](../personal-host-cleanup/Remove-InactiveAvdPersonalHosts.ps1) (`SkippedActivePatchLock`). |
| `AvdPatchState` | Persistent | `v2\|ok=<yyyyMMddTHHmmZ>\|ubr=<int>\|res=<status>\|def=<epoch>` | Tracks last successful servicing timestamp, OS UBR, and `Disconnected` Baseline deferral start time across runs. |
| `AvdLegalHold` | Read-only guard | Any non-empty value except `false`, `0`, `no` (e.g. `true`, `CASE-2026-42`) | Shared with [Remove-InactiveAvdPersonalHosts.ps1](../personal-host-cleanup/Remove-InactiveAvdPersonalHosts.ps1); skips all wake/patch/unlock mutations (`SkippedLegalHold`). |
| `AvdLifecycleState` | Read-only guard | `Quarantined` | Shared with [Remove-InactiveAvdPersonalHosts.ps1](../personal-host-cleanup/Remove-InactiveAvdPersonalHosts.ps1); skipped unless `-IncludeQuarantined` is passed (`SkippedQuarantined`). |

---

## 7. Terminal Status Reference & Operator Triage Runbook

### Key Terminal Status Codes
* **Healthy / Expected Skips:** `UpToDate`, `InstalledNoReboot`, `InstalledAndRebooted`, `RebootedClearedPendingState`, `HotpatchOrDefInstalled_BaselineDeferred`, `DeferredBaselineForUserSession`, `SkippedCohortFilter`, `SkippedCatchUpAlreadyHealthy`, `SkippedActiveSession`, `SkippedActiveSessionPostWake`, `SkippedActiveGuestSession`, `SkippedAutopatchPaused`, `SkippedLegalHold`, `SkippedQuarantined`, `SkippedMaintenanceWindowCutoff`.
* **Transient Warnings (`SucceededWithIssues`):** `WakeAllocationFailed` (regional capacity exhaustion; hibernated RAM untouched), `WakeTimeout`, `InsufficientDiskSpace`, `SkippedTagCeilingReached` (`>=48` existing VM tags), `PartialInstallWarning`, `RebootedPendingSecondPass`, `MorningUnlocked`.
* **Critical / Hard Pipeline Failures (`Failed`):**
  * `ScalingPlanExclusionTagMissingOrMismatch` — Active Scaling Plan has no `exclusionTag`, mismatches `-ExclusionTagName`, or no enabled plan was found for the pool without `-ExclusionTagName` or `-AllowNoScalingPlan`.
  * `NeedsOperatorTriage` / `StaleLockBusyRunCommand` — Guest RunCommand timed out (`TimedOut`, protecting potential in-flight `TiWorker.exe`) or an expired `AvdPatchLock` still has an active `Running`/`Pending` RunCommand. `allowNewSession` is restored automatically when `PriorAllowNewSession == true`.
  * `HibernateRestoreFailed_LeftRunning` — `POST .../deallocate?hibernate=true` failed after 3 settle retries; VM was intentionally left `Running` (never calling cold `deallocate`) to protect user RAM.
  * `AbortedByCircuitBreaker` / `OrchestratorError` / `PowerRestoreFailed` / `PowerRestoreTimeout` / `PostRebootAgentTimeout` / `PostRebootBootNotAdvanced` / `PostRebootVerifyFailed` / `RunCommandOutputParseFailed`.

### Operator Triage Runbook
1. **Inspect active or expired locks (100% read-only):**
   ```powershell
   .\avd\personal-host-patching\Invoke-AvdPersonalNightlyPatching.ps1 `
       -ResourceGroupName 'rg-avd-core' -HostPoolName 'vdpool-pers-01' -Mode Triage |
       Where-Object Status -ne 'NoStaleLock' | Format-Table -AutoSize
   ```
   * `StaleLockIdleRecoverable`: Previous run crashed/timed out and no `Avd*` RunCommand is running. Both `MorningUnlock` (`05:15 UTC`) and the next `Execute` run **self-heal this automatically** (pruning finished `Avd*` RunCommands and clearing `AvdPatchLock` + `AvdNightlyPatch`).
   * `StaleLockBusyRunCommand`: Guest `TiWorker` / RunCommand is still `Running`. Wait for completion before deallocating.
2. **Immediate pre-business-hours recovery (restores drain & clears idle expired locks):**
   ```powershell
   .\avd\personal-host-patching\Invoke-AvdPersonalNightlyPatching.ps1 `
       -ResourceGroupName 'rg-avd-core' -HostPoolName 'vdpool-pers-01' -Mode MorningUnlock -Confirm:$false
   ```
   *(Add `-ForceUnlockActive` only after confirming a prior pipeline run was canceled while its lock TTL has not yet expired).*

---

## 8. Usage Examples & Validation

### 1. Read-Only Audit (Default)
```powershell
.\avd\personal-host-patching\Invoke-AvdPersonalNightlyPatching.ps1 `
    -ResourceGroupName 'rg-avd-core' `
    -HostPoolName 'vdpool-pers-01' |
    Format-Table ShortHostName, AssignedUser, PriorPowerState, SessionDetected, Status, Detail -AutoSize
```

### 2. Nightly 5-Cohort Execution (Mon–Fri 20%/Night, Cutoff 05:00 Berlin)
```powershell
.\avd\personal-host-patching\Invoke-AvdPersonalNightlyPatching.ps1 `
    -HostPoolResourceId '/subscriptions/00000000-0000-0000-0000-000000000000/resourceGroups/rg-avd-core/providers/Microsoft.DesktopVirtualization/hostPools/vdpool-pers-01' `
    -ScalingPlanResourceId '/subscriptions/00000000-0000-0000-0000-000000000000/resourceGroups/rg-avd-core/providers/Microsoft.DesktopVirtualization/scalingPlans/sp-pers-01' `
    -Mode Execute `
    -CohortCount 5 `
    -CohortIndex -1 `
    -MaxConcurrentVMs 60 `
    -MaxStartsPerMinute 15 `
    -StopNewWakeAtLocalHour 5 `
    -Confirm:$false
```

### 3. Local Read-Only Guest Probe (Zero Azure Credentials Required)
```powershell
.\avd\personal-host-patching\Invoke-AvdPersonalNightlyPatching.ps1 -Mode LocalGuestProbe | Format-List *
```

### Canary Validation Checklist (`LAB-01` – `LAB-06`)
1. **`LAB-01` (Hibernated Session + Hotpatch Preservation):** Leave an unsaved file open in VS Code → disconnect → let Autoscale hibernate VM → run `-Mode Execute` during a Monthly B Hotpatch cycle → resume next morning and verify open window is intact, `RebootPerformed == False`, and confirm live WUA `RebootBehavior` / title matching in `InstalledKbs`.
2. **`LAB-02` (Autopatch Pause & 25H2/26H2 Upgrade Block):** Place canary VM in a Paused Autopatch ring → confirm `SkippedAutopatchPaused` and zero Feature Upgrades installed.
3. **`LAB-03` (0-Session Baseline LCU Reboot + 300s Settle Hibernate):** Run on unpatched 0-session VM → confirm scheduled reboot (`RebootPerformed == True`), post-boot OMA-DM sync, and `FinalPowerState == Hibernated`.
4. **`LAB-04` (WinHTTP / Zscaler Proxy under `SYSTEM`):** Confirm `Microsoft.Update.Session` download and `Update-MpSignature` succeed under `NT AUTHORITY\SYSTEM`.
5. **`LAB-05` (Case-Insensitive Admin `exclusionTag` Preservation):** Pre-tag VM with `Exclude-Sp-Pers-01 = OpsPin` → run `-Mode Execute` → verify `OpsPin` remains intact afterward.
6. **`LAB-06` (`MorningUnlock` & Self-Healing Lock Recovery):** Stamp an expired `AvdPatchLock` → verify `-Mode Triage` reports `StaleLockIdleRecoverable` with zero writes, and `-Mode MorningUnlock` restores `allowNewSession == true` and removes both `AvdPatchLock` and `AvdNightlyPatch` while leaving Quarantined VMs drained.

### Offline Regression Suite (195 Assertions)
```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\avd\personal-host-patching\Test-InvokeAvdPersonalNightlyPatching.ps1
```
