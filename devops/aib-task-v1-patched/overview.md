# Azure VM Image Builder DevOps Task v2

A community-maintained Azure DevOps pipeline task for [Azure Image Builder](https://learn.microsoft.com/en-us/azure/virtual-machines/image-builder-overview), updated to the latest API and modern Node.js runtime.

## Why v2?

The original Microsoft-published task uses API `2020-02-14`, Node 10 (EOL), and the deprecated `azure-storage` SDK. This v2 update brings:

| Feature | v1 (Original) | v2 (This Extension) |
|---|---|---|
| **API version** | 2020-02-14 | **2024-02-01** |
| **Node runtime** | Node 10 (EOL) | **Node 20** |
| **Auth** | Service principal only | **Workload identity federation** (OIDC) + service principal |
| **SAS tokens** | 1-year expiry, HTTP allowed | **Timeout-based expiry (4-25h), HTTPS-only** |
| **Blob SDK** | azure-storage (deprecated) | **@azure/storage-blob v12** |
| **Build VM identities** | Not supported | **Multiple user-assigned identities** |
| **Validation** | Not supported | **Built-in with continueOnFailure** |
| **Target regions** | Not configurable | **Multi-region + replica count** |
| **Image versioning** | Not supported | **Auto or explicit version** |
| **VM boot optimization** | Not supported | **Enabled by default** |
| **Error handling** | Basic | **Configurable error action** |
| **Progress** | Polling only | **Phase transitions, elapsed time, heartbeats** |
| **Output variables** | imageUri only | **imageUri, templateName, runOutput** |

## Task Inputs

### Identity
- **Template Identity Resource Id** — user-assigned managed identity for AIB
- **Build VM Identities** — additional identities for the build VM (e.g. Key Vault access)

### Source
- **Platform Image** (publisher/offer/SKU/version) with purchase plan support
- **Managed Image** by resource ID
- **Azure Compute Gallery** version by resource ID

### Customize
- **PowerShell** — inline script or package folder (uploaded to blob storage)
- **Shell** — inline script or package folder
- **Windows Restart** — with optional restart command and timeout
- **Windows Update** — built-in provisioner
- **File** — download files into the build VM

### Validation (Optional)
- Run source validation after customization
- `continueOnFailure` option to proceed even if validation fails

### Distribute
- **Managed Image** — distribute to a managed image
- **Azure Compute Gallery** — distribute to a gallery image definition
  - Target regions with replica counts
  - Automatic or explicit versioning
  - Exclude from latest flag

### VM Profile
- Custom VM size (e.g. `Standard_D4s_v5`)
- Build timeout (minutes)
- OS disk size override
- Proxy VM size for VNet-connected builds
- Container Instance subnet for isolated builds

### Optional Settings
- VM boot optimization (on by default)
- Error handling action (cleanup / abort)

## Output Variables

| Variable | Description |
|---|---|
| `imageUri` | Resource ID of the distributed image |
| `templateName` | Name of the created image template |
| `runOutput` | Full run output name |

## Requirements

- Azure DevOps agent 2.144.0+
- A user-assigned managed identity with appropriate RBAC roles
- A storage account for customizer package upload (if using package-based customizers)

## Links

- [Azure Image Builder documentation](https://learn.microsoft.com/en-us/azure/virtual-machines/image-builder-overview)
- [AIB API 2024-02-01 reference](https://learn.microsoft.com/en-us/rest/api/imagebuilder/)
- [Workload identity federation setup](https://learn.microsoft.com/en-us/azure/devops/pipelines/library/connect-to-azure?view=azure-devops#create-an-azure-resource-manager-service-connection-that-uses-workload-identity-federation)
