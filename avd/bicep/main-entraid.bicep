/*
  SYNOPSIS: AVD Session Host Batch Deployment (Entra ID Join)
  FILE:     bicep/main-entraid.bicep
  DESCRIPTION:
    Orchestrates batch deployment of Entra ID-joined AVD session hosts.
    Supports either:
      1. Explicit vmList array (e.g., ["FSAVDP11001", "FSAVDP11002"]), OR
      2. Prefix + 3-digit zero-padded sequence generation (vmPrefix up to 12 chars
         + padLeft(vmStartIndex + i, vmIndexPadLength, '0'), e.g. "FSAVDP11002").
*/

targetScope = 'resourceGroup'

// --- CORE PARAMETERS ---
param location string = resourceGroup().location

@description('Explicit list of session host VM names (max 15 chars each). When empty, names are generated from vmPrefix + vmCount + vmStartIndex.')
param vmList array = []

@maxLength(12)
@description('Session host prefix (up to 12 characters, e.g. "FSAVDP11" or a 12-char prefix) used when vmList is empty. Combined with 3-digit padding to stay within the 15-char Windows computerName limit.')
param vmPrefix string = ''

@minValue(0)
@maxValue(999)
@description('Number of session hosts to generate when vmList is empty')
param vmCount int = 0

@minValue(1)
@maxValue(999)
@description('Starting index for generated session host names (e.g. 1 -> 001, 2 -> 002)')
param vmStartIndex int = 1

@minValue(1)
@maxValue(4)
@description('Zero-padding width for generated session host index (default 3 -> 001..999, allowing a 12-character vmPrefix within the 15-char NetBIOS limit)')
param vmIndexPadLength int = 3

@description('Custom or Azure Compute Gallery image version resource ID. Leave empty to deploy from Azure Marketplace.')
param imageResourceId string = ''
param hostPoolName string
param hostPoolRg string = resourceGroup().name
@secure()
@description('Host Pool registration token. When empty, resolved automatically from the Host Pool resource.')
param hostPoolToken string = ''
param localAdminUser string
@secure()
param localAdminPassword string
param subnetId string

// --- VM & DISK CONFIGURATION PARAMETERS ---
param vmSize string = 'Standard_B2as_v2'          // Default: Current prod size
param osDiskType string = 'StandardSSD_LRS'       // Default: Current disk type
param osDiskSizeGB int = 128                      // Default: 128 GB
@allowed([
  'ReadWrite'
  'ReadOnly'
  'None'
])
param osDiskCaching string = 'ReadWrite'
param availabilityZones array = []                // Default: No zone pinning
param acceleratedNetworking bool = false          // Default: Off (enable for supported SKUs)
param enableIntune bool = false                   // Default: No Intune auto-enrollment
// Pinned AVD DSC agent package that includes -AadJoin support; override from the pipeline to pin a different build.
#disable-next-line no-hardcoded-env-urls
param avdAgentPackageUrl string = 'https://wvdportalstorageblob.blob.core.windows.net/galleryartifacts/Configuration_1.0.03299.1133.zip'

// --- MARKETPLACE IMAGE FALLBACK PARAMETERS (When imageResourceId is empty) ---
param imagePublisher string = 'MicrosoftWindowsDesktop'
param imageOffer string = 'office-365'
param imageSku string = 'win11-24h2-avd-m365'
param marketplaceImageVersion string = 'latest'

// --- LZA HARDENING, TIMEZONE & OPERATIONS PARAMETERS ---
@description('Windows time zone ID for the session hosts (W. Europe Standard Time = Germany/Berlin/Frankfurt)')
param timeZone string = 'W. Europe Standard Time'

@allowed([
  'TrustedLaunch'
  'ConfidentialVM'
  'Standard'
])
param securityType string = 'TrustedLaunch'
param secureBootEnabled bool = true
param vTpmEnabled bool = true
param enableGuestAttestation bool = true
param encryptionAtHost bool = false
param diskEncryptionSetId string = ''
param enableHibernation bool = false
param applicationSecurityGroupIds array = []
param deployAntiMalware bool = false
param deployMonitoringAgent bool = false
param dataCollectionRuleId string = ''
param installGpuDriver bool = true
param vmUserLoginPrincipalIds array = []
param vmAdminLoginPrincipalIds array = []
@allowed([
  'Group'
  'User'
  'ServicePrincipal'
])
param vmLoginPrincipalType string = 'Group'
@allowed([
  'Group'
  'User'
  'ServicePrincipal'
])
param vmUserLoginPrincipalType string = vmLoginPrincipalType
@allowed([
  'Group'
  'User'
  'ServicePrincipal'
])
param vmAdminLoginPrincipalType string = vmLoginPrincipalType
// ----------------------------------------------------------

param imageName string = ''
param imageVersion string = ''
param currentTimestamp string = utcNow('yyyy-MM-dd')
param customTags object = {}

var generatedVmList = [
  for i in range(0, vmCount): '${vmPrefix}${padLeft((vmStartIndex + i), vmIndexPadLength, '0')}'
]
var rawVmList = [for vmName in (!empty(vmList) ? vmList : generatedVmList): string(vmName)]
// Enforce the 15-character Windows NetBIOS computerName limit across all names upfront (compatible with Bicep 0.28+ without fail())
var invalidVmNames = filter(rawVmList, name => empty(name) || length(name) > 15)
var vmNameValidationGuard = empty(invalidVmNames)
  ? 0
  : int('InvalidSessionHostNameExceeds15CharLimit_${join(invalidVmNames, ',')}')
var resolvedVmList = vmNameValidationGuard == 0 ? rawVmList : []

var imageIdSegments = !empty(imageResourceId) ? split(imageResourceId, '/') : []
var isGalleryImageVersion = contains(toLower(imageResourceId), '/versions/') && length(imageIdSegments) >= 3
var resolvedImageName = !empty(imageName)
  ? imageName
  : (isGalleryImageVersion
      ? imageIdSegments[length(imageIdSegments) - 3]
      : (!empty(imageResourceId) ? last(imageIdSegments) : imageSku))
var resolvedImageVersion = !empty(imageVersion)
  ? imageVersion
  : (isGalleryImageVersion ? last(imageIdSegments) : marketplaceImageVersion)

var defaultTags = union(
  {
    Environment: 'AVD-Factory'
    Workload: 'SessionHost'
    Directory: 'EntraID'
    ImageName: resolvedImageName
    ImageVersion: resolvedImageVersion
    CreatedOn: currentTimestamp
  },
  customTags
)

@batchSize(5)
module sessionHosts 'modules/sessionHost-entraid.bicep' = [
  for (vmName, i) in resolvedVmList: {
    name: 'deploy-${vmName}'
    params: {
      location: location
      vmName: string(vmName)
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
      osDiskCaching: osDiskCaching
      availabilityZone: availabilityZones == [] ? '' : string(availabilityZones[i % length(availabilityZones)])
      acceleratedNetworking: acceleratedNetworking
      enableIntune: enableIntune
      avdAgentPackageUrl: avdAgentPackageUrl

      // Marketplace fallback
      imagePublisher: imagePublisher
      imageOffer: imageOffer
      imageSku: imageSku
      marketplaceImageVersion: marketplaceImageVersion

      // LZA Hardening & Operations
      timeZone: timeZone
      securityType: securityType
      secureBootEnabled: secureBootEnabled
      vTpmEnabled: vTpmEnabled
      enableGuestAttestation: enableGuestAttestation
      encryptionAtHost: encryptionAtHost
      diskEncryptionSetId: diskEncryptionSetId
      enableHibernation: enableHibernation
      applicationSecurityGroupIds: applicationSecurityGroupIds
      deployAntiMalware: deployAntiMalware
      deployMonitoringAgent: deployMonitoringAgent
      dataCollectionRuleId: dataCollectionRuleId
      installGpuDriver: installGpuDriver
      vmUserLoginPrincipalIds: vmUserLoginPrincipalIds
      vmAdminLoginPrincipalIds: vmAdminLoginPrincipalIds
      vmLoginPrincipalType: vmLoginPrincipalType
      vmUserLoginPrincipalType: vmUserLoginPrincipalType
      vmAdminLoginPrincipalType: vmAdminLoginPrincipalType
    }
  }
]

output deployedVmNames array = resolvedVmList
