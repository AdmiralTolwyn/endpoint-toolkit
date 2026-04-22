# AVD Bicep Modules

Bicep templates for deploying AVD session hosts and Azure Image Builder image templates.

## Modules

```
bicep/
├── main-entraid.bicep              # Orchestrator: batch deploy (Entra ID join)
├── main-legacy.bicep               # Orchestrator: batch deploy (legacy AD join)
└── modules/
    ├── imageTemplate.bicep         # AIB image template (API 2024-02-01)
    ├── sessionHost-entraid.bicep   # Single session host (Entra ID join)
    └── sessionHost-legacy.bicep    # Single session host (legacy AD join)
```

## Session Host Deployment

Two orchestrators deploy session hosts in a loop from a `vmList` array parameter. Each calls the corresponding session host module per VM.

### Entra ID Join (`main-entraid.bicep`)

```
az deployment group create \
  --resource-group <rg> \
  --template-file main-entraid.bicep \
  --parameters vmList='["vm-01","vm-02"]' \
               imageResourceId=<imageVersionId> \
               hostPoolName=<hostpool> \
               hostPoolRg=<rg> \
               hostPoolToken=<token> \
               localAdminUser=<user> \
               localAdminPassword=<pass> \
               subnetId=<subnetResourceId>
```

### Legacy AD Join (`main-legacy.bicep`)

Same as above plus domain join parameters:

```
  --parameters domainName=<domain.fqdn> \
               ouPath='OU=AVD,DC=domain,DC=com' \
               domainJoinUser=<user@domain> \
               domainJoinPassword=<pass>
```

### Configurable Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `vmSize` | `Standard_B2as_v2` | VM SKU |
| `osDiskType` | `StandardSSD_LRS` | OS disk storage type |
| `osDiskSizeGB` | `128` | OS disk size |
| `availabilityZones` | `[]` | Zone pinning (round-robin across VMs) |
| `acceleratedNetworking` | `false` | Enable for supported SKUs |
| `enableIntune` | `false` | Auto-enroll in Intune (Entra ID only) |

## Image Template (`imageTemplate.bicep`)

Deploys a `Microsoft.VirtualMachineImages/imageTemplates` resource using API version **2024-02-01**. Used by both the AIB Task v2 and Bicep-only pipelines.

### Source Types

| Source | Parameter | Description |
|--------|-----------|-------------|
| `PlatformImage` | `imagePublisher`, `imageOffer`, `imageSku` | Azure Marketplace image |
| `ManagedImage` | `sourceImageId` | Custom managed image resource ID |
| `SharedImageVersion` | `sourceImageVersionId` | Gallery image version resource ID |

### Customization Flow

1. **Download artifacts** — Downloads and extracts a zip from a SAS URL (`artifactsBlobUrl`) to `C:\BuildArtifacts`
2. **Inline script** — Runs the user-supplied `inlineCustomizerScript` (PowerShell or Shell)
3. **Windows Update** — Optional `WindowsUpdate` customizer with a 5-minute pre-wait

### Distribution Types

| Type | Output |
|------|--------|
| `SharedImage` | Azure Compute Gallery image version (supports multi-region replication) |
| `ManagedImage` | Standalone managed image |
| `VHD` | Raw VHD blob in Azure Storage |

### Key Features (API 2024-02-01)

| Feature | Parameter | Description |
|---------|-----------|-------------|
| VM Boot Optimization | `optimizeVmBoot` | Faster image creation |
| Validation | `enableValidation`, `validationInlineScript` | Pre-distribute validation step |
| Auto-Run | `autoRun` | Start build on template creation |
| Error Handling | `errorHandlingOnCustomizerError` | `cleanup` or `abort` on failure |
| VNet Isolation | `vnetSubnetId`, `containerInstanceSubnetId` | Private build network |
| Staging Resource Group | `stagingResourceGroup` | Use pre-existing staging RG |
| Gallery Versioning | `distributeVersioningScheme` | `Latest` (auto-increment) or `Source` |
| Build VM Identity | `buildVmIdentities` | Assign identities for Key Vault access |

### Usage

Called by pipelines via `az deployment group create`:

```
az deployment group create \
  --resource-group <rg> \
  --template-file modules/imageTemplate.bicep \
  --parameters \
    templateName=t_12345 \
    location=westeurope \
    userAssignedIdentityId=<managedIdentityId> \
    sourceType=PlatformImage \
    imagePublisher=microsoftwindowsdesktop \
    imageOffer=office-365 \
    imageSku=win11-25h2-avd-m365 \
    artifactsBlobUrl=<sasUrl> \
    inlineCustomizerScript='Write-Host "Hello"' \
    distributeType=SharedImage \
    galleryImageId=<galleryImageDefId>
```
