# AVD Personal Host Pool 30-Day Inactivity Quarantine & Decommissioning

[Remove-InactiveAvdPersonalHosts.ps1](Remove-InactiveAvdPersonalHosts.ps1) detects Azure Virtual Desktop (AVD) **Personal** session hosts that have had no user connections for at least `-InactiveDays` (default 30 days) and executes a fail-closed, two-phase quarantine, multi-disk snapshot, and decommissioning workflow across Azure Resource Manager (ARM), Azure Monitor Log Analytics, Microsoft Intune, and Microsoft Entra ID.

---

## 1. Problem & Architecture

Unlike Pooled AVD session hosts—which are stateless and replaced on image updates via [Remove-AvdHosts.ps1](../scripts/Remove-AvdHosts.ps1)—Personal session hosts are stateful 1:1 desktops where users may store data locally on the OS or attached data disks, and where `SessionHost.LastHeartBeat` or VM start/stop events tick during automated patching even when the assigned user has not connected for months.

### Execution Flow (Fail-Closed Two-Phase Lifecycle)

```text
1. ARM GET   /subscriptions/{sub}/resourceGroups/{rg}/providers/Microsoft.DesktopVirtualization/hostPools/{pool}
             └──► Verifies hostPoolType == 'Personal' (skips Pooled host pools)
2. LA POST   /v1/workspaces/{workspaceId}/query (Continuity Gate)
             └──► Verifies WVDConnections has data predating (Now - InactiveDays), recent events within 7 days,
                  and >= MinTelemetryDays active daily buckets out of 30 (fails closed on PartialError or missing logs)
3. LA POST   /v1/workspaces/{workspaceId}/query (Per-Host Activity over InactiveDays)
             └──► Correlates WVDConnections by both Compute vmId (SessionHostAzureVmId) and short hostname
                  • Connected/Completed > 0 ──► ActiveWithinWindow (restores allowNewSession=true if quarantined)
                  • Started-only > 0        ──► ConnectionFailuresDetected (broken host; protected from deletion)
4. ARM + Graph Safety Guards per Session Host:
             • AssignedUser empty           ──► UnassignedWarmHost (excluded)
             • AvdLegalHold tag enabled     ──► ExcludedByLegalHold
             • Live sessions > 0            ──► ActiveSessionsPresent (restores allowNewSession=true if quarantined)
             • VM timeCreated < InactiveDays──► RecentlyCreatedHost
             • User assigned/reassigned     ──► AssignmentObservationWindowPending / RecentlyAssignedUser
5. Phase 1 (Active ──► Quarantined, when -Mode QuarantineOnly or Execute):
             • Enforces MaxDecommissionPercent (default 15%) and MaxDecommissionCount (default 25) circuit breaker
             • PATCH sessionHost allowNewSession = false
             • POST VM /deallocate (stops compute billing immediately while retaining disks)
             • PATCH VM tags (Merge): AvdLifecycleState='Quarantined', AvdQuarantinedVmId, AvdDecommissionAfterUtc (+7d)
6. Phase 2 (Quarantined ──► Decommissioned, when -Mode Execute and UtcNow >= AvdDecommissionAfterUtc):
             • Verifies VM is still in PowerState/deallocated (resets to Active if booted during quarantine)
             • PUT incremental Standard_LRS snapshots of OS disk + all attached Data Disks (polls until Succeeded)
             • DELETE Intune managedDevice (verified via azureADDeviceId == Entra deviceId)
             • PATCH Entra device accountEnabled=false (default Disable, preserving BitLocker keys while snapshots exist;
               skips Hybrid ServerAD devices)
             • DELETE Azure VM, NICs, Disks, and AVD sessionHost registration
```

---

## 2. Required Permissions

| Plane | Role / Permission | Scope | Purpose |
| --- | --- | --- | --- |
| **Azure RBAC** | `Desktop Virtualization Session Host Operator`, `Virtual Machine Contributor`, `Network Contributor`, `Disk Snapshot Contributor`, `Tag Contributor` | Host Pool, Compute, and Snapshot Resource Groups | Read/drain/remove session hosts, deallocate/delete VMs/NICs/disks, create snapshots, merge tags |
| **Log Analytics** | `Log Analytics Reader` | Log Analytics Workspace | Query `WVDConnections` continuity and connection history |
| **Microsoft Graph** | `User.Read.All`, `Device.ReadWrite.All`, `DeviceManagementManagedDevices.ReadWrite.All` | Tenant (Application or Delegated) | Resolve `assignedUser` ObjectId, correlate and disable/delete Entra devices and Intune `managedDevices` |

---

## 3. Arguments

| Parameter | Type | Default | Description |
| --- | --- | --- | --- |
| `-HostPoolResourceId` | `string[]` | — | Full ARM Resource ID(s) of target AVD Personal Host Pools. |
| `-SubscriptionId` | `string` | Current context | Subscription ID when using `-ResourceGroupName` and `-HostPoolName`. |
| `-ResourceGroupName` | `string` | — | Resource Group containing the target AVD Personal Host Pool. |
| `-HostPoolName` | `string` | — | Name of the target AVD Personal Host Pool. |
| `-WorkspaceId` | `string` | *(Mandatory)* | Log Analytics Workspace GUID (`CustomerId`) receiving `WVDConnections` diagnostics. |
| `-Mode` | `string` | `Audit` | `Audit` (read-only), `QuarantineOnly` (Phase 1 drain + deallocate only), or `Execute` (Phase 1 + Phase 2 snapshot & teardown). |
| `-InactiveDays` | `int` | `30` | Minimum days without user connection activity before quarantine (`7..365`). |
| `-QuarantineGraceDays` | `int` | `7` | Days a VM remains deallocated in `Quarantined` state before Phase 2 deletion (`1..90`). |
| `-SnapshotRetentionDays` | `int` | `30` | Days stamped into `DeleteAfterUtc` on recovery disk snapshots (`7..365`). |
| `-SnapshotResourceGroupId` | `string` | VM Compute RG | Optional ARM Resource Group ID (`/subscriptions/<sub>/resourceGroups/<rg>`) for recovery snapshots. |
| `-MaxDecommissionPercent` | `int` | `15` | Maximum percentage of assigned hosts in a pool allowed in quarantine/decommission (`1..100`). |
| `-MaxDecommissionCount` | `int` | `25` | Maximum absolute number of hosts in a pool allowed in quarantine/decommission (`1..500`). |
| `-MinTelemetryDays` | `int` | `20` | Minimum distinct daily buckets (out of 30) required in `WVDConnections` (`1..30`). |
| `-EntraCleanupAction` | `string` | `Disable` | `Disable` (preserves escrowed BitLocker keys on Entra device while disk snapshots exist), `Delete`, or `Skip`. |
| `-ForceCircuitBreakerOverride` | `switch` | `$false` | Bypasses the blast-radius cap when explicitly approved. |
| `-AuthMode` | `string` | `Auto` | `Auto`, `ManagedIdentity`, or `AzContext`. |

---

## 4. Usage Examples

### Read-Only Audit (Default)
```powershell
.\avd\personal-host-cleanup\Remove-InactiveAvdPersonalHosts.ps1 `
    -ResourceGroupName 'rg-avd-core' `
    -HostPoolName 'vdpool-pers-01' `
    -WorkspaceId '11111111-2222-3333-4444-555555555555' |
    Format-Table ShortHostName, AssignedUser, LifecycleState, Action, Status, Detail -AutoSize
```

### Phase 1 Quarantine Only (Drain + Deallocate Unused Hosts)
```powershell
.\avd\personal-host-cleanup\Remove-InactiveAvdPersonalHosts.ps1 `
    -ResourceGroupName 'rg-avd-core' `
    -HostPoolName 'vdpool-pers-01' `
    -WorkspaceId '11111111-2222-3333-4444-555555555555' `
    -Mode QuarantineOnly `
    -Confirm:$false
```

### Full Execution (Phase 1 Quarantine + Phase 2 Snapshot & Decommission)
```powershell
.\avd\personal-host-cleanup\Remove-InactiveAvdPersonalHosts.ps1 `
    -ResourceGroupName 'rg-avd-core' `
    -HostPoolName 'vdpool-pers-01' `
    -WorkspaceId '11111111-2222-3333-4444-555555555555' `
    -Mode Execute `
    -Confirm:$false
```

---

## 5. Offline Regression Tests

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\avd\personal-host-cleanup\Test-RemoveInactiveAvdPersonalHosts.ps1
```
