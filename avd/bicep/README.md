# AVD Bicep Modules

Bicep templates for deploying AVD Management Plane objects (Personal & Pooled Host Pools, Desktop Application Groups, Workspaces, Personal Scaling Plans), Session Hosts, and Azure Image Builder (AIB) image templates.

## Modules & Structure

```
bicep/
├── main-personal-hostpool.bicep          # Full-stack orchestrator: Management Plane + Personal Scaling Plan + Session Hosts (Entra ID)
├── main-entraid.bicep                    # Session host batch orchestrator (Entra ID join — supports vmList OR 12-char prefix + 3-digit index)
├── main-legacy.bicep                     # Session host batch orchestrator (Legacy AD join)
├── modules/
│   ├── hostPool-entraid.bicep            # AVD Management Plane: Host Pool, Desktop App Group, Workspace, Diagnostics & RBAC
│   ├── scalingPlan-personal.bicep        # Personal Desktop Scaling Plan & personalSchedules child resources (API 2024-04-03)
│   ├── powerOnOffRbac-subscription.bicep # Subscription-scoped Power On/Off Contributor RBAC (required by Microsoft Autoscale)
│   ├── powerOnOffRbac.bicep              # Resource-Group-scoped Power On/Off Contributor RBAC (for Start VM on Connect)
│   ├── sessionHost-entraid.bicep         # Single session host (Entra ID join + LZA hardening/monitoring options)
│   ├── sessionHost-legacy.bicep          # Single session host (Legacy AD join)
│   └── imageTemplate.bicep               # AIB image template (API 2024-02-01)
└── parameters/
    └── personal-hostpool-dewc.bicepparam # Example parameter file for Personal Desktops in Germany West Central (dewc)
```

---

## Comparison with AVD Landing Zone Accelerator (LZA)

The modules in this directory adopt the core architectural patterns of the [AVD Landing Zone Accelerator (`Azure/avdaccelerator`)](https://github.com/Azure/avdaccelerator) while removing monolithic subscription-scoped complexity:

| Area | AVD Landing Zone Accelerator (`Azure/avdaccelerator`) | `avd/bicep/` (This Repository) |
|------|-------------------------------------------------------|--------------------------------|
| **Deployment Scope** | Subscription-scoped (`targetScope = 'subscription'`), tightly coupled to AVM wrapper modules, Key Vault creation, and VNet/DNS creation | Resource-group-scoped (`targetScope = 'resourceGroup'`), composable modules that plug into existing Hub-Spoke VNets and pipelines |
| **Session Host Prefix Limit** | Hardcodes `@maxLength(11)` on `avdSessionHostCustomNamePrefix` because it hardcodes 4-digit padding (`padLeft(..., 4, '0')`, e.g., `0001`) | Supports **up to 12-character prefixes** (`@maxLength(12)`) by using configurable 3-digit zero-padding (`vmIndexPadLength = 3`, `001`–`999`, e.g., `FSAVDP11002`) within the 15-char Windows `computerName` limit |
| **Registration Token Flow** | Fetches token via `hostPool.listRegistrationTokens().value[0].token` inside a CustomScriptExtension (`Set-SessionHostConfiguration.ps1`) | `hostPool-entraid.bicep` mints/updates the token via ARM, and `sessionHost-entraid.bicep` passes the token inside encrypted `protectedSettings.properties.registrationInfoToken` (either from an explicit `@secure() hostPoolToken` parameter or resolved via `existingHostPool.listRegistrationTokens().value[0].token`) |
| **Personal Desktops & Scaling** | Bundled inside `deploy-baseline.bicep` alongside Pooled schedules and FSLogix storage | Dedicated `hostPool-entraid.bicep` + `scalingPlan-personal.bicep` + `powerOnOffRbac.bicep` tailored for **Personal** desktops (`Automatic`/`Direct` assignment, `Persistent` load balancing, `StartVMOnConnect`, `scalingPlans/personalSchedules` child resources for Deallocate/Hibernate schedules, and Power On/Off RBAC across Management and Compute Resource Groups) |
| **Germany West Central (`germanywestcentral`)** | Maps `germanywestcentral` (`dewc`) to `Central Europe Standard Time` in `locations.json` | Defaults to `germanywestcentral` with Windows time zone `'W. Europe Standard Time'` (Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna) across Host Pool agent maintenance windows, Scaling Plans, and VM `osProfile` |
| **Zero-Trust & Ops (LZA Parity)** | Trusted Launch, `encryptionAtHost`, CMK Disk Encryption Set, ASG attachment, `IaaSAntimalware`, `AzureMonitorWindowsAgent` + DCR association, NVIDIA & AMD GPU drivers | All supported as opt-in parameters on `sessionHost-entraid.bicep`, `main-entraid.bicep`, and `main-personal-hostpool.bicep` |

---

## 12-Character Session Host Prefix (`vmPrefix`)

Windows enforces a strict **15-character NetBIOS `computerName` limit** (`osProfile.computerName`).

- Standard Azure Portal and AVD LZA templates enforce an **11-character prefix** because they append a 4-character suffix (`-001` or `0001`).
- Both `main-personal-hostpool.bicep` and `main-entraid.bicep` support **up to 12-character prefixes** (`@maxLength(12) param vmPrefix`) by pairing the prefix with a configurable `vmIndexPadLength` (default `3`, yielding `001` through `999` — up to 999 hosts per pool/prefix):

| `vmPrefix` | Prefix Length | `vmStartIndex` | `vmIndexPadLength` | Generated `vmName` / `computerName` | Total Length |
|------------|---------------|----------------|--------------------|-------------------------------------|--------------|
| `FSAVDP11` | 8 chars | `2` | `3` | `FSAVDP11002` | 11 chars (<= 15) |
| `FSAVDP11DEWC` | 12 chars | `2` | `3` | `FSAVDP11DEWC002` | 15 chars (<= 15) |

Both orchestrators validate every resolved VM name upfront before any module executes (compatible with Bicep 0.28+ as well as 0.48+) if any generated or explicit hostname exceeds 15 characters, preventing partial deployments.

---

## Full-Stack Personal Host Pool Deployment (`main-personal-hostpool.bicep`)

Deploys the complete AVD Personal Desktop stack in a single deployment (or modularly via `modules/powerOnOffRbac-subscription.bicep`, `modules/powerOnOffRbac.bicep`, `modules/hostPool-entraid.bicep`, `modules/scalingPlan-personal.bicep`, and `modules/sessionHost-entraid.bicep`):

1. **Power On/Off RBAC (`modules/powerOnOffRbac-subscription.bicep` / `modules/powerOnOffRbac.bicep`)**:
   - When `avdServicePrincipalObjectId` is supplied (the Object ID of the `Azure Virtual Desktop` enterprise application, `appId 9cdead84-a844-4324-93f2-b2e6bb768d07`) and either `deployScalingPlan` or `startVMOnConnect` is `true`, assigns `Desktop Virtualization Power On Off Contributor` at **Subscription scope** by default (`powerOnOffRbacScope = 'Subscription'`, as required by Microsoft Autoscale), or on the Management + Compute Resource Groups when `powerOnOffRbacScope = 'ResourceGroup'`.
2. **Management Plane (`modules/hostPool-entraid.bicep`)**:
   - Personal Host Pool (`hostPoolType: 'Personal'`, `personalDesktopAssignmentType: 'Automatic'` or `'Direct'`, `loadBalancerType: 'Persistent'`, `startVMOnConnect: true`)
   - Case-insensitively normalizes and enforces Entra ID RDP properties (`targetisaadjoined:i:1;enablerdsaadauth:i:1`)
   - Scheduled AVD agent maintenance windows in `'W. Europe Standard Time'`
   - Desktop Application Group (`vdag-desktop-*`) + optional `Desktop Virtualization User` RBAC assignments (`desktopUserPrincipalIds` + `desktopUserPrincipalType` supporting `'Group'`, `'User'`, or `'ServicePrincipal'`)
   - Optional AVD Workspace (`vdws-*`) + optional Log Analytics diagnostic settings
   - Uses collision-resistant 64-character name bounding (`<prefix>-<hash8>` only when a derived name would exceed 64 characters)
3. **Personal Scaling Plan (`modules/scalingPlan-personal.bicep`)**:
   - Deploys the parent `Microsoft.DesktopVirtualization/scalingPlans` resource (`hostPoolType: 'Personal'`, `schedules: []`) and child `Microsoft.DesktopVirtualization/scalingPlans/personalSchedules` resources for Weekday & Weekend schedules (`WithAssignedUser` ramp-up auto-start, `StartVMOnConnect` aligned with the Host Pool setting, configurable `Deallocate` or `Hibernate` on disconnect/logoff)
4. **Session Hosts (`modules/sessionHost-entraid.bicep`)**:
   - Batch-deployed (`@batchSize(5)`) using either `vmPrefix` + `vmCount` + `vmStartIndex` or an explicit `vmList`
   - Encrypts `registrationInfoToken` inside `protectedSettings.properties` of the `Microsoft.PowerShell.DSC` extension
   - Supports independent `vmUserLoginPrincipalType` and `vmAdminLoginPrincipalType` (e.g., `'Group'` for user login and `'User'` for personal admin login)
   - Supports deploying VMs into the same Resource Group or a dedicated `computeResourceGroupName`
   - Supports either a custom Azure Compute Gallery image (`imageResourceId`, automatically extracting both `ImageName` and `ImageVersion` tags from `/galleries/<gal>/images/<def>/versions/<ver>`) or an Azure Marketplace Windows 11 Enterprise image (`imagePublisher` / `imageOffer` / `imageSku`)

### Usage with `.bicepparam` (Germany West Central)

```bash
az deployment group create \
  --resource-group rg-avd-personal-dewc-001 \
  --template-file avd/bicep/main-personal-hostpool.bicep \
  --parameters avd/bicep/parameters/personal-hostpool-dewc.bicepparam \
  --parameters localAdminPassword='<secret>'
```

---

## Session Host Batch Deployment (`main-entraid.bicep` / `main-legacy.bicep`)

Two orchestrators deploy session hosts into an existing Host Pool. `main-entraid.bicep` accepts either an explicit `vmList` array (used by the blue/green update pipeline) **or** `vmPrefix` + `vmCount` + `vmStartIndex`:

### Entra ID Join (`main-entraid.bicep`)

#### Option A: Prefix + 3-digit sequence (up to 12-char prefix)

```bash
az deployment group create \
  --resource-group <rg> \
  --template-file main-entraid.bicep \
  --parameters vmPrefix='FSAVDP11' \
               vmCount=2 \
               vmStartIndex=1 \
               vmIndexPadLength=3 \
               imageResourceId=<imageVersionId> \
               hostPoolName=<hostpool> \
               hostPoolRg=<rg> \
               localAdminUser=<user> \
               localAdminPassword=<pass> \
               subnetId=<subnetResourceId>
```

> **Note on `hostPoolToken` in Option A:** When `hostPoolToken` is omitted, `sessionHost-entraid.bicep` retrieves the existing token via `existingHostPool.listRegistrationTokens().value[0].token` and places it in `protectedSettings.properties.registrationInfoToken`. Ensure the target Host Pool already has an active, non-expired registration token (or pass `hostPoolToken` explicitly as in Option B).

#### Option B: Explicit `vmList` (CI/CD Pipeline)

```bash
az deployment group create \
  --resource-group <rg> \
  --template-file main-entraid.bicep \
  --parameters vmList='["FSAVDP11001","FSAVDP11002"]' \
               imageResourceId=<imageVersionId> \
               hostPoolName=<hostpool> \
               hostPoolRg=<rg> \
               hostPoolToken=<token> \
               localAdminUser=<user> \
               localAdminPassword=<pass> \
               subnetId=<subnetResourceId>
```

### Legacy AD Join (`main-legacy.bicep`)

Same as Option B plus domain join parameters:

```
  --parameters domainName=<domain.fqdn> \
               ouPath='OU=AVD,DC=domain,DC=com' \
               domainJoinUser=<user@domain> \
               domainJoinPassword=<pass>
```

### Configurable Session Host Parameters (`main-entraid.bicep` & `sessionHost-entraid.bicep`)

| Parameter | Default | Description |
|-----------|---------|-------------|
| `vmPrefix` | `''` (`'FSAVDP11'` in `main-personal-hostpool.bicep`) | Session host prefix (up to 12 chars) used when `vmList` is empty |
| `vmCount` | `0` (`2` in `main-personal-hostpool.bicep`) | Number of session hosts to generate when `vmList` is empty |
| `vmStartIndex` | `1` | Starting sequence number (`1` -> `001`) |
| `vmIndexPadLength` | `3` | Zero-padding digits for sequence number (`3` -> `001`..`999`) |
| `vmSize` | `Standard_B2as_v2` (`Standard_D4s_v5` in Personal) | VM SKU |
| `osDiskType` | `StandardSSD_LRS` (`Premium_LRS` in Personal) | OS disk storage type |
| `osDiskSizeGB` | `128` | OS disk size |
| `availabilityZones` | `[]` (`['1','2','3']` in Personal) | Zone pinning (round-robin across VMs) |
| `acceleratedNetworking` | `false` (`true` in Personal) | Enable for supported SKUs |
| `enableIntune` | `false` (`true` in Personal) | Auto-enroll in Intune (Entra ID only) |
| `timeZone` | `'W. Europe Standard Time'` | Windows time zone configured in `osProfile.windowsConfiguration` |
| `securityType` | `'TrustedLaunch'` | `'TrustedLaunch'`, `'ConfidentialVM'`, or `'Standard'` |
| `encryptionAtHost` | `false` | LZA Zero-Trust host-level disk encryption |
| `diskEncryptionSetId` | `''` | Optional CMK Disk Encryption Set resource ID |
| `enableHibernation` | `false` | Enable VM hibernation capability (auto-enabled when Scaling Plan uses `Hibernate`) |
| `applicationSecurityGroupIds` | `[]` | Optional ASG resource IDs attached to the NIC |
| `deployAntiMalware` | `false` | Deploy `IaaSAntimalware` extension |
| `deployMonitoringAgent` | `false` | Deploy `AzureMonitorWindowsAgent` (AMA) extension |
| `dataCollectionRuleId` | `''` | Optional DCR resource ID to associate with the VM |
| `vmUserLoginPrincipalIds` | `[]` | Entra ID Principal IDs granted `Virtual Machine User Login` on the VM |
| `vmAdminLoginPrincipalIds` | `[]` | Entra ID Principal IDs granted `Virtual Machine Administrator Login` on the VM |
| `vmLoginPrincipalType` | `'Group'` | Default principal type (`'Group'`, `'User'`, or `'ServicePrincipal'`) for VM login role assignments |
| `vmUserLoginPrincipalType` | `vmLoginPrincipalType` | Principal type specifically for `vmUserLoginPrincipalIds` |
| `vmAdminLoginPrincipalType` | `vmLoginPrincipalType` | Principal type specifically for `vmAdminLoginPrincipalIds` (e.g., `'User'` for individual admin access) |
| `avdAgentPackageUrl` | `https://wvdportalstorageblob.blob.core.windows.net/galleryartifacts/Configuration_1.0.03299.1133.zip` | Pinned AVD DSC agent package with `-AadJoin` support (override to pin a different build without touching code) |

Session hosts are deployed via a module loop decorated with
`@batchSize(5)` — at most 5 VMs deploy concurrently per run, regardless of how
many names are in the list. Every generated NIC is named `nic-<vmName>` in
both `sessionHost-entraid.bicep` and `sessionHost-legacy.bicep`. The VM
(`Microsoft.Compute/virtualMachines`) and extension resources use API version
`2024-07-01`; the NIC (`Microsoft.Network/networkInterfaces`) uses
`2024-05-01`.

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
| Windows Update pacing | `windowsUpdateWaitSeconds` (default `300`) | Deliberate `Start-Sleep` before the `WindowsUpdate` customizer, to avoid the known AIB race where the Windows Update service isn't ready yet after prior customizers |
| Regional Replica Count | `regionalReplicaCount` (default `1`) | Replicas per target region for the distributed `SharedImage` gallery version |

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
