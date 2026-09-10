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

#### Inbox App Manifest

[StubApps.json](StubApps.json) contains 19 inbox or commonly preinstalled Store
app candidates. This is a download catalog, not a recommended AVD application
baseline: inbox inclusion varies by Windows release and image SKU. Trim the list
to your image, particularly consumer apps such as Xbox and Phone Link. Store IDs
were checked against Microsoft's product listings on 2026-09-10; a valid listing
does not guarantee WinGet offline-download eligibility or AVD support.

| App | Store ID / Verified Listing |
|-----|-----------------------------|
| Microsoft Photos | [9WZDNCRFJBH4](https://apps.microsoft.com/detail/9WZDNCRFJBH4) |
| Windows Calculator | [9WZDNCRFHVN5](https://apps.microsoft.com/detail/9WZDNCRFHVN5) |
| Windows Camera | [9WZDNCRFJBBG](https://apps.microsoft.com/detail/9WZDNCRFJBBG) |
| Windows Clock | [9WZDNCRFJ3PR](https://apps.microsoft.com/detail/9WZDNCRFJ3PR) |
| Windows Notepad | [9MSMLRH6LZF3](https://apps.microsoft.com/detail/9MSMLRH6LZF3) |
| Paint | [9PCFS5B6T72H](https://apps.microsoft.com/detail/9PCFS5B6T72H) |
| Snipping Tool | [9MZ95KL8MR0L](https://apps.microsoft.com/detail/9MZ95KL8MR0L) |
| Windows Media Player | [9WZDNCRFJ3PT](https://apps.microsoft.com/detail/9WZDNCRFJ3PT) |
| Windows Terminal | [9N0DX20HK701](https://apps.microsoft.com/detail/9N0DX20HK701) |
| Microsoft Clipchamp | [9P1J8S7CCWWT](https://apps.microsoft.com/detail/9P1J8S7CCWWT) |
| Quick Assist | [9P7BP5VNWKX5](https://apps.microsoft.com/detail/9P7BP5VNWKX5) |
| Feedback Hub | [9NBLGGH4R32N](https://apps.microsoft.com/detail/9NBLGGH4R32N) |
| Microsoft To Do | [9NBLGGH5R558](https://apps.microsoft.com/detail/9NBLGGH5R558) |
| Phone Link | [9NMPJ99VJBWV](https://apps.microsoft.com/detail/9NMPJ99VJBWV) |
| Windows Sound Recorder | [9WZDNCRFHWKN](https://apps.microsoft.com/detail/9WZDNCRFHWKN) |
| Microsoft Sticky Notes | [9NBLGGH4QGHW](https://apps.microsoft.com/detail/9NBLGGH4QGHW) |
| Xbox App | [9MV0B5HZVK9Z](https://apps.microsoft.com/detail/9MV0B5HZVK9Z) |
| MSN Weather | [9WZDNCRFJ3Q2](https://apps.microsoft.com/detail/9WZDNCRFJ3Q2) |
| Microsoft News | [9WZDNCRFHVFW](https://apps.microsoft.com/detail/9WZDNCRFHVFW) |

Photos Legacy was replaced with current Photos. The former Dev Home entry was
removed because its ID `XP89DCGQ3K6VLD` actually identifies
[Microsoft PowerToys](https://apps.microsoft.com/detail/XP89DCGQ3K6VLD).
Power Automate was also excluded from this AVD catalog: its verified Store ID is
[9NFTCH6J7FHV](https://apps.microsoft.com/detail/9NFTCH6J7FHV), but Microsoft's
[installation documentation](https://learn.microsoft.com/power-automate/desktop-flows/install)
states that Windows multi-session is unsupported.

#### Download and Provision

Run the downloader locally in an interactive session with a current App Installer
that supports `winget download` and `--skip-license`. Offline licenses are retrieved
by default. WinGet authenticates an Entra ID account with License Administrator,
User Administrator, or Global Administrator privileges for license retrieval;
being signed in to the Store app alone is not the prerequisite check.

```powershell
.\Get-StubAppPayloads.ps1 -DownloadPath 'C:\BuildArtifacts\AVD Stubs'
```

Transfer the entire tree to the reference image. Keep one Store product per folder
and retain its `Dependencies` subfolder. The installer recognizes FoD
`<package-basename>.xml` and WinGet `<StoreId>_License.xml` licenses. Loose `.appx`
and `.msix` applications are distinguished from frameworks by their embedded
manifests, not by extension or license presence.

Run provisioning elevated on the reference image, not your download workstation:

```powershell
.\Install-AppxPayloads.ps1 -SourcePath 'C:\BuildArtifacts\AVD Stubs' -LogDirectory 'C:\BuildArtifacts\Logs'
```

The installer requests `-StubPackageOption InstallFull` and passes each app's
frameworks from the same directory or its descendants through
`-DependencyPackagePath`. Use a payload set appropriate to the target OS and
architecture; do not combine unrelated products or multiple releases in one app
folder. DISM remains responsible for signature, version, and dependency validation.

For FoD refreshes, add `-Mode UpdateProvisioned`. The installer compares embedded
Identity Name with provisioned DisplayName exactly, ignoring case. It does not use
the download filename or a name-prefix heuristic. Skipped apps do not cause their
dependencies to be provisioned independently. This selects existing apps; it is
not a version-newness check.

**License behavior changed:** omission is now explicit on both scripts. Use
`-SkipLicense` on the downloader and installer only for apps that permit offline
provisioning without a license on the target edition. Missing licenses fail
provisioning by default; ambiguous license files also fail. See
[WinGet download](https://learn.microsoft.com/windows/package-manager/winget/download)
and [DISM provisioning](https://learn.microsoft.com/powershell/module/dism/add-appxprovisionedpackage).

Both scripts emit result objects and exit `1` on failures. An empty or
framework-only installation tree also exits `1`. Validate provisioning and actual
launch with a newly created user before publishing the image. This is not a
guaranteed repair of existing users' registrations or a field-validated stub fix.

#### Local Validation

```powershell
Invoke-Pester -Path .\avd\scripts\tests\AppxPayloads.Tests.ps1 -Output Detailed
```

Requires Pester 5. Tests create synthetic package archives and mock WinGet/DISM;
they do not download apps or change Windows provisioning. The installer test copy
is built from its parsed parameter block and statements, omitting only script-level
requirements so no elevation is needed. Coverage includes both entry points,
manifest validation, path arguments, licenses, package classification, update
filtering, dependency isolation, logging-directory creation, and failure exits.
Live Store acquisition and provisioning on an AVD image remain separate checks.

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
