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

Two pipeline options for Azure Image Builder — choose based on your constraints:

### Option 1: AIB Task v2 (`img-build-custom-image.yml`)

Uses the `AzureImageBuilderTask@2` marketplace extension. Simpler YAML, but requires the extension to be installed in your Azure DevOps organization.

### Option 2: Bicep-Only (`img-build-bicep-only.yml`)

Zero extension dependency. Uses `AzureCLI@2` tasks + the shared [`imageTemplate.bicep`](../bicep/modules/imageTemplate.bicep) module:

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

Both pipelines produce the same output: a versioned image in an Azure Compute Gallery.

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
- For AIB Task v2: [Azure Image Builder Task](https://marketplace.visualstudio.com/items?itemName=AzureImageBuilder.devOps-task-for-azure-image-builder) extension installed
