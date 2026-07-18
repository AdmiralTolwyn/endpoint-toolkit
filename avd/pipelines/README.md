# AVD Pipelines

Azure DevOps YAML pipelines for AVD image lifecycle and session host management.

## Pipelines

| Pipeline | Trigger | Purpose |
|----------|---------|---------|
| [img-build-custom-image.yml](img-build-custom-image.yml) | Manual | Build custom AVD image using **AIB Task v2** extension |
| [img-build-bicep-only.yml](img-build-bicep-only.yml) | Manual | Build custom AVD image using **Bicep + AzureCLI** (no extension) |
| [avd-update-hostpool-entraid.yml](avd-update-hostpool-entraid.yml) | Manual | Blue/green session host replacement (**Entra ID** join) |
| [avd-update-hostpool-legacy.yml](avd-update-hostpool-legacy.yml) | Manual | Blue/green session host replacement (**legacy AD** join) |
| [avd-cleanup-hostpool.yml](avd-cleanup-hostpool.yml) | Manual | Decommission outdated session hosts |
| [avd-activator.yml](avd-activator.yml) | Scheduled (15 min) | Hybrid join activation for pending session hosts |

## Image Build Approaches

Two pipeline options for Azure Image Builder — choose based on your constraints. Both produce the same output: a versioned image in an Azure Compute Gallery.

| Aspect | AIB Task v2 (`img-build-custom-image.yml`) | Bicep-Only (`img-build-bicep-only.yml`) |
|--------|--------------------------------------------|------------------------------------------|
| Marketplace dependency | Requires the [Azure VM Image Builder DevOps Task v2](https://marketplace.visualstudio.com/items?itemName=vacuumbreather.devOps-task-for-azure-image-builder-v2) extension | None (only built-in `AzureCLI@2` + `PowerShell@2`) |
| YAML complexity | ~40 lines for the build step | ~150 lines (prepare / deploy / poll / cleanup) |
| Customizer upload | Handled by the task (uploads `packagePath` to AIB staging blob) | Manual: pipeline zips repo + uploads to staging container |
| Template lifecycle | Task creates / runs / cleans the AIB template | Bicep `imageTemplate.bicep` deploys; `az image builder run` triggers; explicit cleanup step deletes (retries up to 5x with 60s backoff — AIB refuses template deletion while a run is in progress) |
| Progress reporting | Built-in phase transitions, elapsed time, heartbeats | Custom polling loop in the pipeline |
| Output variables | `imageUri`, `templateName`, `runOutput` | Parsed from `az` CLI output in pipeline scripts |
| Auth | Workload identity federation (OIDC) or service principal | Workload identity federation (OIDC) or service principal |

Both pipelines run an **identical customizer chain** (same 9 steps, same script arguments — see [Recommended pipeline order](../customizer/README.md#recommended-pipeline-order)) and pass full distribute parity (`optimizeVmBoot`, two target regions, versioning, error handling), so the choice between them is purely an infrastructure/dependency decision, not a feature trade-off. The Bicep-only artifact upload uses a user-delegation SAS (`--as-user`, 4h expiry) rather than a storage account key.

### Option 1: AIB Task v2 (`img-build-custom-image.yml`)

Uses the [`AzureImageBuilderTaskV2@2`](https://marketplace.visualstudio.com/items?itemName=vacuumbreather.devOps-task-for-azure-image-builder-v2) marketplace extension — a community-maintained refresh of Microsoft's original `AzureImageBuilderTask@1`. **Recommended** when you can install the extension in your Azure DevOps organization. The extension is currently distributed as a Marketplace **Preview** release.

Key v2 improvements over the deprecated v1 task:

- **Modern API**: `2024-02-01` instead of `2020-02-14`
- **Modern runtime**: Node 20 instead of Node 10 (EOL)
- **Federated auth (OIDC)**: workload identity federation in addition to service principal
- **Hardened SAS tokens**: timeout-based expiry (4–25 h), HTTPS-only — instead of 1-year HTTP-allowed
- **Modern Blob SDK**: `@azure/storage-blob` v12 instead of the deprecated `azure-storage` SDK
- **Multiple build VM identities** (e.g. for Key Vault access during customization)
- **Built-in validation** stage with `continueOnFailure`
- **Multi-region distribution** with replica counts
- **Image versioning**: automatic or explicit
- **VM boot optimization** (enabled by default)
- **Configurable error handling** (cleanup vs. abort)
- **Output variables**: `imageUri`, `templateName`, `runOutput`

See the [marketplace listing](https://marketplace.visualstudio.com/items?itemName=vacuumbreather.devOps-task-for-azure-image-builder-v2) for the full input reference and changelog.

The `BuildImage` job carries an explicit `timeoutInMinutes: 180` (the AIB build itself is capped separately at 150 min via `buildTimeout`), so a stuck build fails the job instead of hitting the Azure DevOps default 60-minute job timeout.

### Option 2: Bicep-Only (`img-build-bicep-only.yml`)

Zero extension dependency. Uses `AzureCLI@2` tasks + the shared [`imageTemplate.bicep`](../bicep/modules/imageTemplate.bicep) module. **Recommended** when extension installation in your Azure DevOps organization is restricted, or when you want full IaC ownership of the AIB template definition.

```
Prepare ──► Build ──► Cleanup
  │            │          │
  ├─ Checkout  ├─ Deploy  ├─ Delete template
  ├─ Zip       │  Bicep   ├─ Delete staging
  └─ Upload    ├─ Trigger │  container
     to blob   │  build   └─ Cleanup legacy
               └─ Poll       artifacts
                  status
```

## Host Pool Update Flow

The session host pipelines implement a blue/green deployment pattern across six
manually-triggered stages, sharing the helper scripts in
[avd/scripts](../scripts):

```
 ┌───────────────────────────┐   ┌────────────────────┐   ┌───────────────────────┐
 │  update-hostpool-*        │ → │   activator (15m)  │ → │  cleanup-hostpool     │
 │  (Preparation ▸ Validate ▸│   │  Hybrid join only  │   │  Decommission drained │
 │   Canary ▸ Drain Old ▸    │   │  legacy-AD pools   │   │  + outdated hosts     │
 │   Blast ▸ Rollback*)      │   │                     │   │                       │
 └───────────────────────────┘   └────────────────────┘   └───────────────────────┘
   * Rollback only runs if Canary / Drain Old / Blast fails
```

Old hosts are **not** touched until the canary proves out: the canary deploys
first, is health-gated, and is drained for validation — only then does the
`DrainOld` stage drain the rest of the outdated fleet, followed by the blast
deployment. If any of the three production stages (`DeployCanary`, `DrainOld`,
`DeployBlast`) fails, a `Rollback` stage restores `AllowNewSession = $true` on
the old fleet so users are never stranded on a half-replaced pool.

### Update pipelines — `avd-update-hostpool-entraid.yml` / `avd-update-hostpool-legacy.yml`

Both pipelines share the same six-stage skeleton; they differ only in **identity
join model**, **Bicep template**, and the **secrets / parameters** required for
domain join.

| Aspect | `…-entraid.yml` | `…-legacy.yml` |
|--------|------------------|----------------|
| Pipeline name | `…-AVD-SessionHost-Update` | `…-AVD-Legacy-Update` |
| Join model | Microsoft Entra ID join (optional Intune enrollment) | Hybrid Azure AD / on-prem AD domain join |
| Bicep entrypoint | [main-entraid.bicep](../bicep/main-entraid.bicep) | [main-legacy.bicep](../bicep/main-legacy.bicep) |
| Extra variables | `enableIntune` | `domainName`, `ouPath` |
| `baseName` / `vmCountOverride` | Both pipelines expose these (see [Key operational variables](#key-operational-variables)) | Both pipelines expose these |
| Extra Key Vault secrets | — | `avd-domain-join-user`, `avd-domain-join-password` |
| Default availability zones | `[]` (none) | `[1,2,3]` |
| Post-deploy activation | Native — Entra-joined hosts register immediately | Requires `avd-activator.yml` to flip `HybridStatus` tag |

Stage flow (identical for both):

| # | Stage | Job / Script | What it does |
|---|-------|--------------|--------------|
| 1 | **Preparation** | `Get-AvdDetails.ps1` | Discovers the latest gallery image version, lists outdated session hosts, splits them into a single **canary** + **blast** batch (sized to the number of *outdated* hosts, not total pool size), and generates ISO-8601 hostnames. Outputs are exposed as pipeline variables (`TargetImageVersion`, `TargetImageId`, `CanaryList`, `BlastList`, `OutdatedHostsList`). |
| 2 | **Validate** | inline script | Prints a human-readable deployment plan and asserts that the target image version was resolved and the canary list is non-empty. Registration tokens are no longer part of this stage's output — see [token handling](#registration-token-handling) below. |
| 3 | **DeployCanary** | `AzureCLI@2` + Bicep | Mints a fresh host pool registration token in-stage, deploys **one** new session host from the latest image, joins it (Entra ID or AD), and registers it with the host pool. A **health gate** then polls `Get-AzWvdSessionHost` for up to 10 minutes waiting for `Status -eq 'Available'` before the canary is drained for smoke-testing. Subsequent stages only run if this whole stage succeeds — old hosts are untouched at this point. |
| 4 | **DrainOld** | `Set-AvdDrainMode.ps1` | Runs only after `DeployCanary` succeeds. Sets `AllowNewSession = $false` on every outdated host so users start migrating off. Optionally writes a `Stage=Maintenance, Action=Drain` telemetry event. |
| 5 | **DeployBlast** | `AzureCLI@2` + Bicep | Mints another fresh registration token in-stage, deploys the remaining new session hosts in a single batch using the same Bicep template / parameters, then drains them as well so admins can validate before opening to users. Skipped automatically when `BlastList` is empty. |
| 6 | **Rollback** | inline `Update-AzWvdSessionHost` | Runs only if `DeployCanary`, `DrainOld`, or `DeployBlast` failed (and the run wasn't canceled). Restores `AllowNewSession = $true` on every host in `OutdatedHostsList` so the old fleet keeps serving users. `Set-AvdDrainMode.ps1` has no "undrain" direction, so this is done inline. |

#### Registration token handling

Registration tokens are minted **just-in-time inside each Deploy stage**
(`DeployCanary`, `DeployBlast`) — there is no cross-stage secret hand-off. Each
stage calls `Get-AzWvdRegistrationInfo`; if the existing token is missing or
expires within 6 hours, it mints a new 24-hour token via
`New-AzWvdRegistrationInfo`. This avoids a token minted in `Preparation`
expiring mid-rollout on later-joining hosts, and avoids passing a secret
through pipeline stage outputs.

#### Secrets and parameter files

Secrets (`avd-local-admin-password`, `avd-domain-join-user`,
`avd-domain-join-password`, and the minted `HostPoolToken`) are read via the
step's `env:` block and never macro-interpolated into scripts. The generated
`canary.json` / `blast.json` ARM parameter files are deleted by an `always()`
cleanup step at the end of each deploy job, regardless of success or failure.

#### Key operational variables

| Variable | Pipeline(s) | Default | Purpose |
|----------|-------------|---------|---------|
| `vmCountOverride` | `…-entraid.yml`, `…-legacy.yml` | `0` (auto-detect) | Circuit-breaker override. `Get-AvdDetails.ps1` aborts if 100% of a multi-host pool is outdated, unless a non-zero override confirms the full-fleet replacement. |
| `baseName` | `…-entraid.yml`, `…-legacy.yml` | `<YOURVMPREFIX>` | Prefix for generated ISO-8601 week-based hostnames. |
| `enableTelemetry` | All update/cleanup pipelines | `'false'` | Set to `'true'` to send `Write-DeploymentTelemetry.ps1` events to Log Analytics. |

Notes:

- The `Preparation`, `DrainOld`, `DeployCanary`, and `DeployBlast` stages emit
  optional Log Analytics telemetry via `Write-DeploymentTelemetry.ps1` when
  `enableTelemetry: 'true'`. **That script uses the HTTP Data Collector API,
  which Microsoft is retiring — support ends 2026-09-14.** The migration
  target is the Logs Ingestion API (Data Collection Rule / Data Collection
  Endpoint + Entra ID auth); this has not been implemented yet.
- Hosts deployed by these pipelines are **born drained** — flipping them into
  service is an explicit operator decision (typically after Canary validation).
- The legacy pipeline never undrains hosts on its own; for hybrid pools the
  activation step is owned by `avd-activator.yml` (see below).

#### Deployment diagnostics — async submit + poll

The `DeployCanary` / `DeployBlast` stages do **not** run a blocking
`az deployment group create`. A blocking call prints nothing while ARM waits on
the VM extension chain (domain join / AAD login → `GuestAttestation` →
`Microsoft.PowerShell.DSC` AVD agent). Extensions install **serially** and can
stall for a long time, so the VM reaches `Succeeded` while the deployment stays
`Running` — the run *looks* hung with no indication of the cause.

Instead each deploy step:

1. Submits the deployment with `--no-wait` and fails immediately if the submit
   itself is rejected (`$LASTEXITCODE`).
2. Polls every 30 s, logging a timestamped **deployment provisioning state** plus
   the **per-VM extension provisioning state** (`az vm extension list`) — so a
   stuck extension is named in the log.
3. On `Failed` / `Canceled`, dumps the failing operations
   (`az deployment operation group list`) and fails the task.
4. Enforces an internal **55-minute deadline** (under the task's
   `timeoutInMinutes: 60`) that dumps the still-`Running` operations before the
   agent hard-kills the task, so a genuine hang is still diagnosable.

Typical signature of the "VM replaced then hangs" symptom: the deployment sits at
`Running` with the VM's `GuestAttestation` or DSC extension stuck in
`Creating` / `Transitioning` — usually restricted outbound access from the
session-host subnet to the attestation or AVD registration endpoints.

### `avd-activator.yml` — Hybrid join activator (scheduled, every 15 min)

Lightweight scheduled pipeline that runs `Invoke-HybridActivator.ps1` on a
`*/15 * * * *` cron against `main`. The script scans the entire subscription for
VMs tagged `HybridStatus = Pending`, completes their hybrid Azure AD join
handshake (DSC/extension state, `dsregcmd` validation, AVD agent registration),
flips the tag to `Active`, and removes drain mode. **Only relevant for legacy AD
pipelines** — Entra ID-joined hosts activate themselves at provisioning time and
are skipped.

Run it standalone whenever a legacy update finishes: it has no inputs other than
`serviceConnection` and is safe to re-run; pending VMs are processed
idempotently.

### `avd-cleanup-hostpool.yml` — Decommission outdated hosts (manual)

A single-step pipeline that wraps `Remove-AvdHosts.ps1`. For each session host in
the target pool that is **already drained** *and* whose image version is older
than the latest published version in the Compute Gallery, the script:

1. If the host has active sessions, stamps a `PendingDrainTimestamp` tag on
   first encounter and warns connected users. On later runs, once
   `drainGracePeriodHours` has elapsed, it **force-logs-off any remaining
   sessions** and re-verifies the session count is zero before proceeding —
   a host that still shows live sessions after the forced logoff is skipped
   with an error, never decommissioned.
2. Deletes the underlying VM **first**. The AVD host registration is only
   removed once the VM delete is confirmed successful, so a failed VM delete
   never leaves an orphaned registration pointing at nothing (or vice versa).
3. Deletes the NIC and OS disk, resolved from the VM's own network/storage
   profile references (not a naming-pattern guess).
4. Removes the session host registration from the host pool, and the Entra ID
   device record for Entra-joined hosts (`Directory=EntraID` tag).
5. Optionally emits a `Stage=Cleanup, Action=Decommission` telemetry event.

Inputs are the host pool, compute RG, and gallery coordinates used to compute
"latest". This is the final step of the blue/green cycle and is intentionally
manual so operators can verify the new fleet is healthy before tearing down the
old one.

**Safety defaults:** the `simulate` pipeline variable defaults to `'true'` —
a real (destructive) run requires an explicit opt-out (`simulate: 'false'`)
when queuing the pipeline. `Simulate` mode is a true dry run: every delete is
logged but nothing is changed. `drainGracePeriodHours` (default `24`)
controls how long a host with active sessions is left alone before the
forced logoff kicks in.

## Configuration

All pipelines use `<YOUR...>` placeholders for environment-specific values. Search for `<YOUR` and replace:

| Placeholder | Description |
|-------------|-------------|
| `<YOURSERVICECONNECTION>` | Azure DevOps service connection name |
| `<YOURSUBSCRIPTIONID>` | Azure subscription ID |
| `<YOURRESOURCEGROUP-AIB>` | Resource group for AIB resources |
| `<YOURRESOURCEGROUP-CORE>` | Resource group for host pool and workspace |
| `<YOURRESOURCEGROUP-COMPUTE>` | Resource group for session host VMs |
| `<YOURKEYVAULTNAME>` | Key Vault storing secrets (admin password, storage keys) |
| `<YOURMANAGEDIDENTITY>` | User-assigned managed identity for AIB |
| `<YOURGALLERYNAME>` | Azure Compute Gallery name |
| `<YOURIMAGEDEFINITION>` | Gallery image definition name |
| `<YOURSTORAGEACCOUNT>` | Staging storage account for build artifacts |
| `<YOURREPOSITORYNAME>` | Azure DevOps repository with customizer scripts |
| `<YOURHOSTPOOLNAME>` | AVD host pool name |
| `<YOURVNETNAME>` | Virtual network name for session host NICs |
| `<YOURVMPREFIX>` | `baseName` prefix used to generate ISO-8601 week-based hostnames |
| `<YOURDOMAINNAME>` | AD domain FQDN (legacy join only) |
| `<YOUROUPATH>` | OU distinguished name for computer objects (legacy join only) |

See [Key operational variables](#key-operational-variables) above for
non-placeholder variables (`vmCountOverride`, `enableTelemetry`) and the
cleanup pipeline's `simulate` / `drainGracePeriodHours` variables.

## Prerequisites

- Azure DevOps project with service connection (workload identity federation recommended)
- Azure Compute Gallery with image definition
- User-assigned managed identity with Contributor + Storage Blob Data Contributor roles
- Storage account for staging artifacts
- Key Vault with admin credentials
- For AIB Task v2: [Azure VM Image Builder DevOps Task v2](https://marketplace.visualstudio.com/items?itemName=vacuumbreather.devOps-task-for-azure-image-builder-v2) extension installed (Azure DevOps agent **3.232.1+** — the first agent version with the Node 20 task handler; the extension itself is a Marketplace **Preview** release)

## Required outbound endpoints

Session host deployments run a chain of VM extensions that each need their own
outbound (egress) connectivity from the session-host subnet. A restricted
subnet is the most common cause of a deployment that reaches `Succeeded` on
the VM but hangs at the [async poll](#deployment-diagnostics--async-submit--poll)
stage. Microsoft's authoritative, always-current list is
[Required FQDNs and endpoints for Azure Virtual Desktop](https://learn.microsoft.com/azure/virtual-desktop/required-fqdn-endpoint) —
**verify against that doc**, since Microsoft revises regional/service
endpoints over time. The extensions polled by these pipelines map to it as
follows:

| Extension | Purpose | Representative endpoints (verify against the linked doc) |
|-----------|---------|-----------------------------------------------------------|
| `AADLoginForWindows` | Entra ID join | `login.microsoftonline.com`, `pas.windows.net`, `enterpriseregistration.windows.net`, `device.login.microsoftonline.com` |
| `GuestAttestation` | TPM attestation (TrustedLaunch) | Regional `*.attest.azure.net` |
| `Microsoft.PowerShell.DSC` (AVD agent) | AVD registration | `wvdportalstorageblob.blob.core.windows.net`, `catalogartifact.azureedge.net`, `*.wvd.microsoft.com` |

If a deployment stalls with the DSC/AVD agent extension stuck in
`Creating`/`Transitioning`, check the AVD agent's own event log first:
**WVD-Agent Event ID 3701** reports the specific region-scoped FQDN the agent
is trying (and failing) to reach — it is more precise than guessing from the
static list above.
