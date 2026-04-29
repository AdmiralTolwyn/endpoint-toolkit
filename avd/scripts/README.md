# AVD Scripts

PowerShell scripts for AVD session host lifecycle management. Used standalone or called by the [pipelines](../pipelines/).

## Scripts

### Session Host Lifecycle

| Script | Purpose |
|--------|---------|
| [Get-AvdDetails.ps1](Get-AvdDetails.ps1) | Pre-flight orchestrator — finds latest gallery image, identifies outdated hosts, generates ISO-8601 week-based hostnames, manages host pool registration tokens |
| [Set-AvdDrainMode.ps1](Set-AvdDrainMode.ps1) | Sets `AllowNewSession = $false` on outdated hosts to block new connections |
| [Remove-AvdHosts.ps1](Remove-AvdHosts.ps1) | Decommissions drained hosts — removes AVD registration, Entra ID device record, VM, NIC, and OS disk |

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
| [Write-DeploymentTelemetry.ps1](Write-DeploymentTelemetry.ps1) | Sends structured deployment events to Log Analytics via HTTP Data Collector API |

### Golden Image Provisioning

| Script | Purpose |
|--------|---------|
| [Get-StubAppPayloads.ps1](Get-StubAppPayloads.ps1) | Downloads Microsoft Store Stub App offline payloads via `winget download --source msstore` for side-loading during Packer image build. App list is data-driven via [StubApps.json](StubApps.json). Run locally with Entra ID auth. |

## Usage in Pipelines

The update and cleanup pipelines call these scripts in sequence:

```
Get-AvdDetails.ps1          ← Identify outdated hosts + generate new hostnames
    │
    ├─► Set-AvdDrainMode.ps1    ← Block new sessions on old hosts
    │
    ├─► Bicep deployment         ← Deploy new session hosts
    │
    ├─► Invoke-HybridActivator   ← (legacy AD only, scheduled pipeline)
    │
    └─► Remove-AvdHosts.ps1     ← Decommission old hosts after grace period
```

## Requirements

- Az PowerShell modules: `Az.Accounts`, `Az.DesktopVirtualization`, `Az.Compute`, `Az.Network`
- `Remove-AvdHosts.ps1` also requires `Az.Resources` (for Entra ID device cleanup)
- `Write-DeploymentTelemetry.ps1` requires a Log Analytics workspace ID and shared key
- `Invoke-FslRepairDisk.ps1` requires local admin access and SMB access to the FSLogix share
