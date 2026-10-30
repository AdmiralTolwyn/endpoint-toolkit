/*
  SYNOPSIS: Single AVD Session Host (Entra ID Join)
  FILE:     bicep/modules/sessionHost-entraid.bicep
  DESCRIPTION:
    Deploys a single Entra ID-joined AVD Session Host (NIC, VM, Extensions,
    optional DCR association, and optional RBAC).
    Supports both Custom/Gallery Image IDs and Marketplace images, 12-char
    prefixes (up to 15-char computerName), and LZA hardening/monitoring options.
*/

param location string = 'germanywestcentral'
@minLength(1)
@maxLength(15)
param vmName string
@description('Custom or Azure Compute Gallery image version resource ID. Leave empty to use Marketplace image parameters.')
param imageResourceId string = ''
param subnetId string
param adminUsername string
@secure()
param adminPassword string
param hostPoolName string
param hostPoolRg string = resourceGroup().name
@secure()
@description('Host Pool registration token. When empty, resolved automatically from the existing Host Pool resource.')
param hostPoolToken string = ''
param tags object = {}

// --- CONFIGURATION PARAMETERS ---
param vmSize string = 'Standard_D4s_v5'
param osDiskType string = 'Premium_LRS'
param osDiskSizeGB int = 128
@allowed([
  'ReadWrite'
  'ReadOnly'
  'None'
])
param osDiskCaching string = 'ReadWrite'
param availabilityZone string = ''
param acceleratedNetworking bool = false
param enableIntune bool = false
// Pinned AVD DSC agent package that includes -AadJoin support (Configuration.zip is a 2020 package lacking -AadJoin).
// Override from the pipeline or parameter file to pin a different build without touching code.
#disable-next-line no-hardcoded-env-urls
param avdAgentPackageUrl string = 'https://wvdportalstorageblob.blob.core.windows.net/galleryartifacts/Configuration_1.0.03299.1133.zip'

// --- MARKETPLACE IMAGE FALLBACK PARAMETERS (Used when imageResourceId is empty) ---
param imagePublisher string = 'MicrosoftWindowsDesktop'
param imageOffer string = 'office-365'
param imageSku string = 'win11-24h2-avd-m365'
param marketplaceImageVersion string = 'latest'

// --- LZA HARDENING, TIMEZONE & OPERATIONS PARAMETERS ---
@description('Windows time zone ID for the VM (W. Europe Standard Time = Germany/Berlin/Frankfurt). Leave empty to keep image default.')
param timeZone string = 'W. Europe Standard Time'

@allowed([
  'TrustedLaunch'
  'ConfidentialVM'
  'Standard'
])
@description('VM security profile type. TrustedLaunch is default for Gen2 Win11 images.')
param securityType string = 'TrustedLaunch'
param secureBootEnabled bool = true
param vTpmEnabled bool = true
param enableGuestAttestation bool = true

@description('Enable Encryption at Host (LZA Zero-Trust recommendation; requires Microsoft.Compute/EncryptionAtHost feature registered)')
param encryptionAtHost bool = false

@description('Optional Disk Encryption Set (DES) resource ID for customer-managed key (CMK) OS disk encryption')
param diskEncryptionSetId string = ''

@description('Enable VM hibernation capability (required if Personal Scaling Plan uses Hibernate on disconnect)')
param enableHibernation bool = false

@description('Optional array of Application Security Group (ASG) resource IDs to attach to the session host NIC')
param applicationSecurityGroupIds array = []

@description('Deploy Microsoft IaaSAntimalware extension (LZA option)')
param deployAntiMalware bool = false

@description('Deploy Azure Monitor Windows Agent (AMA) extension (LZA monitoring option)')
param deployMonitoringAgent bool = false

@description('Optional Data Collection Rule (DCR) resource ID to associate with this session host when deployMonitoringAgent is true')
param dataCollectionRuleId string = ''

@description('Automatically install NVIDIA or AMD GPU driver extension when a GPU VM size is detected')
param installGpuDriver bool = true

@description('Optional array of Entra ID Principal IDs to grant "Virtual Machine User Login" on this VM')
param vmUserLoginPrincipalIds array = []

@description('Optional array of Entra ID Principal IDs to grant "Virtual Machine Administrator Login" on this VM (e.g., personal desktop admin access)')
param vmAdminLoginPrincipalIds array = []

@allowed([
  'Group'
  'User'
  'ServicePrincipal'
])
@description('Default principal type for VM login role assignments (used when vmUserLoginPrincipalType or vmAdminLoginPrincipalType is not overridden)')
param vmLoginPrincipalType string = 'Group'

@allowed([
  'Group'
  'User'
  'ServicePrincipal'
])
@description('Principal type for vmUserLoginPrincipalIds (defaults to vmLoginPrincipalType)')
param vmUserLoginPrincipalType string = vmLoginPrincipalType

@allowed([
  'Group'
  'User'
  'ServicePrincipal'
])
@description('Principal type for vmAdminLoginPrincipalIds (defaults to vmLoginPrincipalType; set to User when assigning individual admin users alongside a Group for user login)')
param vmAdminLoginPrincipalType string = vmLoginPrincipalType
// --------------------------------

// 0. HOST POOL REFERENCE (Used when hostPoolToken is not explicitly passed)
resource existingHostPool 'Microsoft.DesktopVirtualization/hostPools@2024-04-03' existing = {
  name: hostPoolName
  scope: resourceGroup(hostPoolRg)
}

// 1. CALCULATE LINK TAGS
var hostPoolResourceId = resourceId(hostPoolRg, 'Microsoft.DesktopVirtualization/hostpools', hostPoolName)
var finalTags = union(tags, { 'cm-resource-parent': hostPoolResourceId })

// GPU DETECTION (NVIDIA & AMD, aligned with LZA)
var requireNvidiaGpu = installGpuDriver && (startsWith(vmSize, 'Standard_NC') || contains(vmSize, '_A10_v5') || contains(vmSize, 'Standard_NV12s_v3') || contains(vmSize, 'Standard_NV24s_v3') || contains(vmSize, 'Standard_NV48s_v3'))
var requireAmdGpu = installGpuDriver && (contains(vmSize, 'Standard_NV4as_v4') || contains(vmSize, 'Standard_NV8as_v4') || contains(vmSize, 'Standard_NV16as_v4') || contains(vmSize, 'Standard_NV32as_v4') || contains(vmSize, '_V710_v5'))

// IMAGE & STORAGE PROFILE CONSTRUCTION
var resolvedImageReference = !empty(imageResourceId)
  ? {
      id: imageResourceId
    }
  : {
      publisher: imagePublisher
      offer: imageOffer
      sku: imageSku
      version: marketplaceImageVersion
    }

var resolvedManagedDisk = !empty(diskEncryptionSetId)
  ? {
      storageAccountType: osDiskType
      diskEncryptionSet: {
        id: diskEncryptionSetId
      }
    }
  : {
      storageAccountType: osDiskType
    }

var resolvedWindowsConfig = !empty(timeZone)
  ? {
      enableAutomaticUpdates: false
      provisionVMAgent: true
      timeZone: timeZone
    }
  : {
      enableAutomaticUpdates: false
      provisionVMAgent: true
    }

var nicAsgReferences = [
  for asgId in applicationSecurityGroupIds: {
    id: asgId
  }
]

// Built-in Role Definition IDs for Entra ID VM Login
var vmUserLoginRoleId = subscriptionResourceId(
  'Microsoft.Authorization/roleDefinitions',
  'fb879df8-f326-4884-ac11-06f6c05be272'
)
var vmAdminLoginRoleId = subscriptionResourceId(
  'Microsoft.Authorization/roleDefinitions',
  '1c0163c0-47e6-4577-8991-ea5c82e286e4'
)

// 2. NETWORK INTERFACE
resource nic 'Microsoft.Network/networkInterfaces@2024-05-01' = {
  name: 'nic-${vmName}'
  location: location
  tags: finalTags
  properties: {
    ipConfigurations: [
      {
        name: 'ipconfig1'
        properties: {
          privateIPAllocationMethod: 'Dynamic'
          subnet: { id: subnetId }
          applicationSecurityGroups: !empty(applicationSecurityGroupIds) ? nicAsgReferences : null
        }
      }
    ]
    enableAcceleratedNetworking: acceleratedNetworking
  }
}

// 3. VIRTUAL MACHINE
resource vm 'Microsoft.Compute/virtualMachines@2024-07-01' = {
  name: vmName
  location: location
  zones: availabilityZone != '' ? [availabilityZone] : []
  tags: finalTags
  identity: { type: 'SystemAssigned' }
  properties: {
    hardwareProfile: {
      vmSize: vmSize
    }
    additionalCapabilities: enableHibernation
      ? {
          hibernationEnabled: true
        }
      : null
    osProfile: {
      computerName: vmName
      adminUsername: adminUsername
      adminPassword: adminPassword
      windowsConfiguration: resolvedWindowsConfig
    }
    storageProfile: {
      imageReference: resolvedImageReference
      osDisk: {
        name: '${vmName}-OSDisk'
        createOption: 'FromImage'
        caching: osDiskCaching
        diskSizeGB: osDiskSizeGB > 0 ? osDiskSizeGB : null
        managedDisk: resolvedManagedDisk
        deleteOption: 'Delete'
      }
    }
    securityProfile: securityType != 'Standard'
      ? {
          securityType: securityType
          encryptionAtHost: encryptionAtHost
          uefiSettings: {
            secureBootEnabled: secureBootEnabled
            vTpmEnabled: vTpmEnabled
          }
        }
      : (encryptionAtHost
          ? {
              encryptionAtHost: true
            }
          : null)
    diagnosticsProfile: { bootDiagnostics: { enabled: true } }
    networkProfile: {
      networkInterfaces: [ { id: nic.id, properties: { deleteOption: 'Delete' } } ]
    }
    licenseType: 'Windows_Client'
  }
}

// 3b. OPTIONAL VM LOGIN RBAC ASSIGNMENTS (Personal Desktop User / Admin Login)
resource vmUserLoginAssignments 'Microsoft.Authorization/roleAssignments@2022-04-01' = [
  for principalId in vmUserLoginPrincipalIds: if (!empty(principalId)) {
    name: guid(vm.id, principalId, vmUserLoginRoleId)
    scope: vm
    properties: {
      roleDefinitionId: vmUserLoginRoleId
      principalId: principalId
      principalType: vmUserLoginPrincipalType
    }
  }
]

resource vmAdminLoginAssignments 'Microsoft.Authorization/roleAssignments@2022-04-01' = [
  for principalId in vmAdminLoginPrincipalIds: if (!empty(principalId)) {
    name: guid(vm.id, principalId, vmAdminLoginRoleId)
    scope: vm
    properties: {
      roleDefinitionId: vmAdminLoginRoleId
      principalId: principalId
      principalType: vmAdminLoginPrincipalType
    }
  }
]

// 4. EXTENSIONS (Standard Entra ID + LZA Set)
resource aadLogin 'Microsoft.Compute/virtualMachines/extensions@2024-07-01' = {
  parent: vm
  name: 'AADLoginForWindows'
  location: location
  properties: {
    publisher: 'Microsoft.Azure.ActiveDirectory'
    type: 'AADLoginForWindows'
    typeHandlerVersion: '2.0'
    autoUpgradeMinorVersion: true
    settings: enableIntune ? { mdmId: '0000000a-0000-0000-c000-000000000000' } : null
  }
}

resource guestAttestation 'Microsoft.Compute/virtualMachines/extensions@2024-07-01' = if (securityType != 'Standard' && enableGuestAttestation) {
  parent: vm
  name: 'GuestAttestation'
  location: location
  dependsOn: [ aadLogin ]
  properties: {
    publisher: 'Microsoft.Azure.Security.WindowsAttestation'
    type: 'GuestAttestation'
    typeHandlerVersion: '1.0'
    autoUpgradeMinorVersion: true
  }
}

// GPU DRIVERS (Conditional - NVIDIA)
resource gpuDriver 'Microsoft.Compute/virtualMachines/extensions@2024-07-01' = if (requireNvidiaGpu) {
  parent: vm
  name: 'NvidiaGpuDriverWindows'
  location: location
  dependsOn: [ guestAttestation ]
  properties: {
    publisher: 'Microsoft.HpcCompute'
    type: 'NvidiaGpuDriverWindows'
    typeHandlerVersion: '1.6'
    autoUpgradeMinorVersion: true
  }
}

// GPU DRIVERS (Conditional - AMD)
resource amdGpuDriver 'Microsoft.Compute/virtualMachines/extensions@2024-07-01' = if (requireAmdGpu) {
  parent: vm
  name: 'AmdGpuDriverWindows'
  location: location
  dependsOn: [ guestAttestation ]
  properties: {
    publisher: 'Microsoft.HpcCompute'
    type: 'AmdGpuDriverWindows'
    typeHandlerVersion: '1.1'
    autoUpgradeMinorVersion: true
  }
}

// ANTIMALWARE (Conditional - LZA)
resource antiMalware 'Microsoft.Compute/virtualMachines/extensions@2024-07-01' = if (deployAntiMalware) {
  parent: vm
  name: 'MicrosoftAntiMalware'
  location: location
  dependsOn: [ aadLogin ]
  properties: {
    publisher: 'Microsoft.Azure.Security'
    type: 'IaaSAntimalware'
    typeHandlerVersion: '1.3'
    autoUpgradeMinorVersion: true
    settings: {
      AntimalwareEnabled: true
      RealtimeProtectionEnabled: 'true'
      ScheduledScanSettings: {
        isEnabled: 'true'
        day: '7'
        time: '120'
        scanType: 'Quick'
      }
    }
  }
}

// AZURE MONITOR AGENT (Conditional - LZA)
resource azureMonitorAgent 'Microsoft.Compute/virtualMachines/extensions@2024-07-01' = if (deployMonitoringAgent) {
  parent: vm
  name: 'AzureMonitorWindowsAgent'
  location: location
  dependsOn: [ aadLogin ]
  properties: {
    publisher: 'Microsoft.Azure.Monitor'
    type: 'AzureMonitorWindowsAgent'
    typeHandlerVersion: '1.0'
    autoUpgradeMinorVersion: true
    enableAutomaticUpgrade: true
  }
}

resource dcrAssociation 'Microsoft.Insights/dataCollectionRuleAssociations@2022-06-01' = if (deployMonitoringAgent && !empty(dataCollectionRuleId)) {
  name: 'dcr-${vmName}'
  scope: vm
  dependsOn: [ azureMonitorAgent ]
  properties: {
    dataCollectionRuleId: dataCollectionRuleId
  }
}

#disable-next-line no-hardcoded-env-urls
resource avdAgent 'Microsoft.Compute/virtualMachines/extensions@2024-07-01' = {
  parent: vm
  name: 'Microsoft.PowerShell.DSC'
  location: location
  dependsOn: [
    aadLogin
    guestAttestation
    gpuDriver
    amdGpuDriver
    antiMalware
    azureMonitorAgent
  ]
  properties: {
    publisher: 'Microsoft.Powershell'
    type: 'DSC'
    typeHandlerVersion: '2.73'
    autoUpgradeMinorVersion: true
    settings: {
      modulesUrl: avdAgentPackageUrl
      configurationFunction: 'Configuration.ps1\\AddSessionHost'
      properties: {
        hostPoolName: hostPoolName
        aadJoin: true
      }
    }
    protectedSettings: {
      properties: {
        registrationInfoToken: !empty(hostPoolToken)
          ? hostPoolToken
          : existingHostPool.listRegistrationTokens().value[0].token
      }
    }
  }
}

output vmId string = vm.id
output vmName string = vm.name
output nicId string = nic.id
