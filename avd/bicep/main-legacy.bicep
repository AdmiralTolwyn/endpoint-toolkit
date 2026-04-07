/*
  SYNOPSIS: Orchestrator for Legacy AD Join
  FILE:     bicep/main-legacy.bicep
*/
targetScope = 'resourceGroup'

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

// --- CONFIGURATION ---
param vmSize string = 'Standard_B2as_v2'
param osDiskType string = 'StandardSSD_LRS'
param osDiskSizeGB int = 128
param availabilityZones array = []                // Default: No zone pinning
param acceleratedNetworking bool = false          // Default: Off

// --- LEGACY AD PARAMS ---
param domainName string
param ouPath string
param domainJoinUser string
@secure()
param domainJoinPassword string
// -----------------------

param imageName string
param imageVersion string
param currentTimestamp string = utcNow('yyyy-MM-dd')

var defaultTags = {
  Environment: 'AVD-Factory'
  Workload:    'SessionHost'
  Directory:   'LegacyAD' // Updated tag
  ImageName:    imageName
  ImageVersion: imageVersion
  CreatedOn:    currentTimestamp
}

module sessionHosts 'modules/sessionHost-legacy.bicep' = [for (vmName, i) in vmList: {
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

    // Legacy AD
    domainName: domainName
    ouPath: ouPath
    domainJoinUser: domainJoinUser
    domainJoinPassword: domainJoinPassword
  }
}]