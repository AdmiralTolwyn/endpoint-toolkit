# AVD Scripts

PowerShell scripts for AVD session host lifecycle management. Used standalone or called by the [pipelines](../pipelines/).

## Scripts

### Session Host Lifecycle

| Script | Purpose |
|--------|---------|
| [Get-AvdDetails.ps1](Get-AvdDetails.ps1) | Pre-flight orchestrator — finds latest gallery image, identifies outdated hosts, sizes the new fleet to the **outdated host count** (not total pool size), generates ISO-8601 week-based hostnames, and includes a circuit breaker requiring `-VmCountOverride` when 100% of a multi-host pool is outdated |
| [Set-AvdDrainMode.ps1](Set-AvdDrainMode.ps1) | Sets `AllowNewSession = $false` on outdated hosts to block new connections. Resolves short VM names to registered session host names via a pool-wide FQDN leaf-prefix fallback when an exact name match fails, and exits `1` if any target host fails to drain (so a partial drain can't silently let a canary/blast deployment proceed) |
| [Remove-AvdHosts.ps1](Remove-AvdHosts.ps1) | Decommissions drained hosts — force-logs-off any sessions remaining after the drain grace period (re-verifying zero sessions before proceeding), deletes the VM **first**, then the NIC/OS disk (resolved from the VM's own profile references), then the AVD registration and Entra ID device record. Supports a true dry-run `-Simulate` |

`Get-AvdDetails.ps1` still computes/generates a host pool registration token
internally (`Get`/`New-AzWvdRegistrationInfo`) and exposes it as a
`HostPoolToken` output variable, but the update pipelines no longer consume
that output — each deploy stage mints its own fresh token in-stage instead
(see [pipelines/README.md](../pipelines/README.md#registration-token-handling)).

### Hybrid Join

| Script | Purpose |
|--------|---------|
| [Invoke-HybridActivator.ps1](Invoke-HybridActivator.ps1) | Tag-driven scanner: finds VMs with `HybridStatus=Pending`, triggers `Automatic-Device-Join` task, validates via `dsregcmd`, undrains on success |

### FSLogix

| Script | Purpose |
|--------|---------|
| [Invoke-FslRepairDisk.ps1](Invoke-FslRepairDisk.ps1) | Enterprise-scale repair of dirty FSLogix profile/O365 VHD(x) disks — mounts, checks dirty bit, runs `chkdsk /f`, multi-threaded |
| [Remove-FSLogixTeamsArtifacts.ps1](Remove-FSLogixTeamsArtifacts.ps1) | Removes stale Teams classic + new cache paths left inside FSLogix containers after Redirections.xml changes |

### Telemetry

| Script | Purpose |
|--------|---------|
| [Write-DeploymentTelemetry.ps1](Write-DeploymentTelemetry.ps1) | Sends structured deployment events to Log Analytics via the HTTP Data Collector API. Accepts `WorkspaceId`/`SharedKey`/`EventData` as parameters or via `WORKSPACE_ID`/`SHARED_KEY`/`EVENT_DATA_JSON` env vars (the contract the pipelines use). **The HTTP Data Collector API is deprecated — Microsoft support ends 2026-09-14**; migration target is the Logs Ingestion API (DCR/DCE + Entra ID auth), not yet implemented here |

### Golden Image Provisioning

| Script | Purpose |
|--------|---------|
| [Get-StubAppPayloads.ps1](Get-StubAppPayloads.ps1) | Downloads Microsoft Store Stub App offline payloads via `winget download --source msstore` for side-loading during Packer image build. App list is data-driven via [StubApps.json](StubApps.json). Run locally with Entra ID auth. |
| [Install-AppxPayloads.ps1](Install-AppxPayloads.ps1) | Side-loads / re-provisions inbox AppX/MSIX packages from a local payload tree via `Add-AppxProvisionedPackage`. `-Mode Install` (default) for the stub-app fix; `-Mode UpdateProvisioned` to refresh built-in apps from a mounted FoD / Language ISO (legacy AIB workflow). Pairs with `Get-StubAppPayloads.ps1`. |

## Usage in Pipelines

The update pipelines call these scripts in the order the stages actually run —
**old hosts are not drained until after the canary deployment succeeds**:

```
Get-AvdDetails.ps1              ← Identify outdated hosts + generate new hostnames
    │
    ├─► Bicep deployment (Canary)  ← Deploy + health-gate one new session host
    ├─► Set-AvdDrainMode.ps1        ← Drain the canary for validation
    │
    ├─► Set-AvdDrainMode.ps1        ← THEN drain the rest of the outdated fleet
    │
    ├─► Bicep deployment (Blast)   ← Deploy remaining new session hosts
    ├─► Set-AvdDrainMode.ps1        ← Drain the blast batch
    │
    ├─► (on failure) inline Update-AzWvdSessionHost  ← Rollback: un-drain old hosts
    │
    ├─► Invoke-HybridActivator.ps1  ← (legacy AD only, scheduled pipeline)
    │
    └─► Remove-AvdHosts.ps1        ← Decommission old hosts after grace period (separate pipeline)
```

## Requirements

- Az PowerShell modules: `Az.Accounts`, `Az.DesktopVirtualization`, `Az.Compute`
- `Az.Resources` is also required by `Get-AvdDetails.ps1`, `Set-AvdDrainMode.ps1`,
  `Invoke-HybridActivator.ps1`, and `Remove-AvdHosts.ps1` (tag reads/writes via
  `Update-AzTag`/`Get-AzResource`, and Entra ID device cleanup in
  `Remove-AvdHosts.ps1`) — see each script's `#Requires -Modules` line
- `Write-DeploymentTelemetry.ps1` requires a Log Analytics workspace ID and shared key
- `Invoke-FslRepairDisk.ps1` requires local admin access and SMB access to the FSLogix share
