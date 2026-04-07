/*
  SYNOPSIS: Single AVD Session Host (Entra ID Join)
  FILE:     bicep/modules/sessionHost-entraid.bicep
*/

param location string
@maxLength(15)
param vmName string
param imageResourceId string
param subnetId string
param adminUsername string
@secure()
param adminPassword string
param hostPoolName string
param hostPoolRg string 
@secure()
param hostPoolToken string
param tags object

// --- CONFIGURATION PARAMETERS ---
param vmSize string
param osDiskType string
param osDiskSizeGB int
param availabilityZone string = ''
param acceleratedNetworking bool = false
param enableIntune bool = false
// --------------------------------

// 1. CALCULATE LINK TAGS
var hostPoolResourceId = resourceId(hostPoolRg, 'Microsoft.DesktopVirtualization/hostpools', hostPoolName)
var finalTags = union(tags, { 'cm-resource-parent': hostPoolResourceId })

// GPU DETECTION
var requireNvidiaGpu = startsWith(vmSize, 'Standard_NC') || contains(vmSize, '_A10_v5')

// 2. NETWORK INTERFACE
resource nic 'Microsoft.Network/networkInterfaces@2023-04-01' = {
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
        }
      }
    ]
    enableAcceleratedNetworking: acceleratedNetworking
  }
}

// 3. VIRTUAL MACHINE
resource vm 'Microsoft.Compute/virtualMachines@2023-03-01' = {
  name: vmName
  location: location
  zones: availabilityZone != '' ? [availabilityZone] : []
  tags: finalTags
  identity: { type: 'SystemAssigned' }
  properties: {
    hardwareProfile: { 
        vmSize: vmSize 
    }
    osProfile: {
      computerName: vmName
      adminUsername: adminUsername
      adminPassword: adminPassword
      // FIXED: Properties must be on separate lines
      windowsConfiguration: { 
        enableAutomaticUpdates: false
        provisionVMAgent: true 
      }
    }
    storageProfile: {
      imageReference: { id: imageResourceId }
      osDisk: {
        name: '${vmName}-OSDisk'
        createOption: 'FromImage'
        diskSizeGB: osDiskSizeGB 
        managedDisk: { 
            storageAccountType: osDiskType 
        }
        deleteOption: 'Delete'
      }
    }
    securityProfile: {
      securityType: 'TrustedLaunch'
      // FIXED: Properties must be on separate lines
      uefiSettings: { 
        secureBootEnabled: true
        vTpmEnabled: true 
      }
    }
    diagnosticsProfile: { bootDiagnostics: { enabled: true } }
    networkProfile: {
      networkInterfaces: [ { id: nic.id, properties: { deleteOption: 'Delete' } } ]
    }
    licenseType: 'Windows_Client'
  }
}

// 4. EXTENSIONS (Standard Entra ID Set)
resource aadLogin 'Microsoft.Compute/virtualMachines/extensions@2023-03-01' = {
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

resource guestAttestation 'Microsoft.Compute/virtualMachines/extensions@2021-11-01' = {
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
resource gpuDriver 'Microsoft.Compute/virtualMachines/extensions@2023-03-01' = if (requireNvidiaGpu) {
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

#disable-next-line no-hardcoded-env-urls
resource avdAgent 'Microsoft.Compute/virtualMachines/extensions@2023-03-01' = {
  parent: vm
  name: 'Microsoft.PowerShell.DSC'
  location: location
  dependsOn: [ aadLogin, guestAttestation, gpuDriver ]
  properties: {
    publisher: 'Microsoft.Powershell'
    type: 'DSC'
    typeHandlerVersion: '2.73'
    autoUpgradeMinorVersion: true
    settings: {
      modulesUrl: 'https://wvdportalstorageblob.blob.core.windows.net/galleryartifacts/Configuration_1.0.03299.1133.zip'
      configurationFunction: 'Configuration.ps1\\AddSessionHost'
      properties: {
        hostPoolName: hostPoolName
        registrationInfoToken: hostPoolToken
        aadJoin: true
      }
    }
  }
}