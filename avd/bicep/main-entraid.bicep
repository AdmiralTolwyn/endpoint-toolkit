/*
  SYNOPSIS: AVD Session Host Batch Deployment (Entra ID Join)
  FILE:     bicep/main-entraid.bicep
*/

targetScope = 'resourceGroup'

// --- EXISTING PARAMETERS ---
param location string = resourceGroup().location
param vmList array
param imageResourceId string
param hostPoolName string
param hostPoolRg string
@secure()
param hostPoolToken string
param localAdminUser string
@secure()
param localAdminPassword string
param subnetId string

// --- NEW CONFIGURATION PARAMETERS (No longer hardcoded) ---
param vmSize string = 'Standard_B2as_v2'          // Default: Your current prod size
param osDiskType string = 'StandardSSD_LRS'       // Default: Your current disk type
param osDiskSizeGB int = 128                      // Default: Your requested size
param availabilityZones array = []                // Default: No zone pinning
param acceleratedNetworking bool = false          // Default: Off (enable for supported SKUs)
param enableIntune bool = false                   // Default: No Intune auto-enrollment
// ----------------------------------------------------------

param imageName string
param imageVersion string
param currentTimestamp string = utcNow('yyyy-MM-dd')

var defaultTags = {
  Environment: 'AVD-Factory'
  Workload:    'SessionHost'
  Directory:   'EntraID'
  ImageName:    imageName
  ImageVersion: imageVersion
  CreatedOn:    currentTimestamp
}

module sessionHosts 'modules/sessionHost-entraid.bicep' = [for (vmName, i) in vmList: {
  name: 'deploy-${vmName}'
  params: {
    location: location
    vmName: vmName
    imageResourceId: imageResourceId
    subnetId: subnetId
    adminUsername: localAdminUser
    adminPassword: localAdminPassword
    hostPoolName: hostPoolName
    hostPoolRg: hostPoolRg
    hostPoolToken: hostPoolToken
    tags: defaultTags
    
    // Configuration
    vmSize: vmSize
    osDiskType: osDiskType
    osDiskSizeGB: osDiskSizeGB
    availabilityZone: availabilityZones == [] ? '' : string(availabilityZones[i % length(availabilityZones)])
    acceleratedNetworking: acceleratedNetworking
    enableIntune: enableIntune
  }
}]