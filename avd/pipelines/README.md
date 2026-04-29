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
| Template lifecycle | Task creates / runs / cleans the AIB template | Bicep `imageTemplate.bicep` deploys; `az image builder run` triggers; explicit cleanup step deletes |
| Progress reporting | Built-in phase transitions, elapsed time, heartbeats | Custom polling loop in the pipeline |
| Output variables | `imageUri`, `templateName`, `runOutput` | Parsed from `az` CLI output in pipeline scripts |
| Auth | Workload identity federation (OIDC) or service principal | Workload identity federation (OIDC) or service principal |

### Option 1: AIB Task v2 (`img-build-custom-image.yml`)

Uses the [`AzureImageBuilderTaskV2@2`](https://marketplace.visualstudio.com/items?itemName=vacuumbreather.devOps-task-for-azure-image-builder-v2) marketplace extension — a community-maintained refresh of Microsoft's original `AzureImageBuilderTask@1`. **Recommended** when you can install the extension in your Azure DevOps organization.

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

The session host pipelines implement a blue/green deployment pattern across four
manually-triggered (or scheduled) stages, sharing the helper scripts in
[avd/scripts](../scripts):

```
 ┌───────────────────────┐   ┌────────────────────┐   ┌───────────────────────┐
 │  update-hostpool-*    │ → │   activator (15m)  │ → │  cleanup-hostpool     │
 │  (Plan ▸ Drain Old ▸  │   │  Hybrid join only  │   │  Decommission drained │
 │   Canary ▸ Blast)     │   │  legacy-AD pools   │   │  + outdated hosts     │
 └───────────────────────┘   └────────────────────┘   └───────────────────────┘
```

### Update pipelines — `avd-update-hostpool-entraid.yml` / `avd-update-hostpool-legacy.yml`

Both pipelines share the same five-stage skeleton; they differ only in **identity
join model**, **Bicep template**, and the **secrets / parameters** required for
domain join.

| Aspect | `…-entraid.yml` | `…-legacy.yml` |
|--------|------------------|----------------|
| Pipeline name | `…-AVD-SessionHost-Update` | `…-AVD-Legacy-Update` |
| Join model | Microsoft Entra ID join (optional Intune enrollment) | Hybrid Azure AD / on-prem AD domain join |
| Bicep entrypoint | [main-entraid.bicep](../bicep/main-entraid.bicep) | [main-legacy.bicep](../bicep/main-legacy.bicep) |
| Extra variables | `enableIntune` | `domainName`, `ouPath`, `baseName`, `vmCountOverride` |
| Extra Key Vault secrets | — | `avd-domain-join-user`, `avd-domain-join-password` |
| Default availability zones | `[]` (none) | `[1,2,3]` |
| Post-deploy activation | Native — Entra-joined hosts register immediately | Requires `avd-activator.yml` to flip `HybridStatus` tag |

Stage flow (identical for both):

| # | Stage | Job / Script | What it does |
|---|-------|--------------|--------------|
| 1 | **Preparation** | `Get-AvdDetails.ps1` | Discovers the latest gallery image version, lists outdated session hosts, splits them into a single **canary** + **blast** batch, generates ISO-8601 hostnames, and issues a fresh host pool registration token. Outputs are exposed as pipeline variables (`TargetImageVersion`, `TargetImageId`, `CanaryList`, `BlastList`, `OutdatedHostsList`, `HostPoolToken`). |
| 2 | **Validate** | inline script | Prints a human-readable deployment plan and asserts that the target version and registration token were resolved. Fails fast if the plan is empty or invalid. |
| 3 | **DrainOld** | `Set-AvdDrainMode.ps1` | Sets `AllowNewSession = $false` on every outdated host so users start migrating off. Optionally writes a `Stage=Maintenance, Action=Drain` telemetry event. |
| 4 | **DeployCanary** | `AzureCLI@2` + Bicep | Deploys **one** new session host from the latest image, joins it (Entra ID or AD), registers it with the host pool, and **immediately drains** it for smoke-testing. Subsequent stages only run if this succeeds. |
| 5 | **DeployBlast** | `AzureCLI@2` + Bicep | Deploys the remaining new session hosts in a single batch using the same Bicep template / parameters, then drains them as well so admins can validate before opening to users. Skipped automatically when `BlastList` is empty. |

Notes:

- All four "real work" stages (`DrainOld`, `DeployCanary`, `DeployBlast`, plus the
  preparation stage) emit optional Log Analytics telemetry via
  `Write-DeploymentTelemetry.ps1` when `enableTelemetry: 'true'`.
- Hosts deployed by these pipelines are **born drained** — flipping them into
  service is an explicit operator decision (typically after Canary validation).
- The legacy pipeline never undrains hosts on its own; for hybrid pools the
  activation step is owned by `avd-activator.yml` (see below).

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

1. Logs off any remaining sessions (grace period enforced inside the script).
2. Removes the session host registration from the host pool.
3. Deletes the underlying VM, NIC, and OS disk from the compute resource group.
4. Optionally emits a `Stage=Cleanup, Action=Decommission` telemetry event.

Inputs are the host pool, compute RG, and gallery coordinates used to compute
"latest". This is the final step of the blue/green cycle and is intentionally
manual so operators can verify the new fleet is healthy before tearing down the
old one.

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
| `<YOURDOMAINNAME>` | AD domain FQDN (legacy join only) |
| `<YOUROUPATH>` | OU distinguished name for computer objects (legacy join only) |

## Prerequisites

- Azure DevOps project with service connection (workload identity federation recommended)
- Azure Compute Gallery with image definition
- User-assigned managed identity with Contributor + Storage Blob Data Contributor roles
- Storage account for staging artifacts
- Key Vault with admin credentials
- For AIB Task v2: [Azure VM Image Builder DevOps Task v2](https://marketplace.visualstudio.com/items?itemName=vacuumbreather.devOps-task-for-azure-image-builder-v2) extension installed (Azure DevOps agent 2.144.0+)
