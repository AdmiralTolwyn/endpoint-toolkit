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

The update pipelines implement a blue/green deployment pattern:

1. **Pre-flight** (`Get-AvdDetails.ps1`) — Finds latest image, identifies outdated hosts, generates ISO-8601 hostnames, creates registration token
2. **Drain** (`Set-AvdDrainMode.ps1`) — Blocks new sessions on outdated hosts
3. **Deploy** (Bicep) — Provisions new session hosts from the latest image
4. **Activate** (`Invoke-HybridActivator.ps1`) — Hybrid join activation (legacy AD only, runs on schedule)
5. **Cleanup** (`Remove-AvdHosts.ps1`) — Decommissions old hosts after grace period

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
