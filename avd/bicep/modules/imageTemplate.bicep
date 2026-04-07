/*
  SYNOPSIS: Azure Image Builder - Image Template
  FILE:     bicep/modules/imageTemplate.bicep
  DESCRIPTION:
    Deploys a Microsoft.VirtualMachineImages/imageTemplates resource.
    Replaces the deprecated AzureImageBuilderTask@1 DevOps extension.
    Supports: PlatformImage, ManagedImage, SharedImageVersion sources.
    API version: 2024-02-01 (latest GA with errorHandling, optimize, validate,
    autoRun, managedResourceTags, stagingResourceGroup, containerInstanceSubnetId,
    targetRegions, versioning)
*/

// === REQUIRED PARAMETERS ===
param location string
param templateName string

@description('Resource ID of the user-assigned managed identity for AIB')
param userAssignedIdentityId string

// === SOURCE PARAMETERS ===
@allowed(['PlatformImage', 'ManagedImage', 'SharedImageVersion'])
param sourceType string = 'PlatformImage'

@description('For PlatformImage: publisher:offer:sku format or individual params')
param imagePublisher string = ''
param imageOffer string = ''
param imageSku string = ''
param imageVersion string = 'latest'

@description('For ManagedImage source: resource ID of the managed image')
param sourceImageId string = ''

@description('For SharedImageVersion source: resource ID of the image version')
param sourceImageVersionId string = ''

// === CUSTOMIZE PARAMETERS ===
@description('SAS URL to the build artifacts zip uploaded by the pipeline')
param artifactsBlobUrl string

@description('Inline PowerShell customizer script')
param inlineCustomizerScript string

@description('Name of the build folder inside C:\\BuildArtifacts')
param buildFolderName string = 'buildartifacts'

@allowed(['powershell', 'shell'])
param provisioner string = 'powershell'

param runElevated bool = true
param runAsSystem bool = false
param windowsUpdate bool = true

// === DISTRIBUTE PARAMETERS ===
@allowed(['SharedImage', 'ManagedImage', 'VHD'])
param distributeType string = 'SharedImage'

@description('For SharedImage: gallery image definition resource ID (without /versions)')
param galleryImageId string = ''

@description('Comma-separated target regions for gallery replication')
param targetRegions string = 'westeurope'

@description('Storage account type for gallery image version per target region')
@allowed(['', 'Standard_LRS', 'Standard_ZRS', 'Premium_LRS'])
param distributeStorageAccountType string = ''

@description('Exclude this version from being the latest in the gallery')
param excludeFromLatest bool = false

@description('Versioning scheme for gallery distribution: Latest or Source')
@allowed(['', 'Latest', 'Source'])
param distributeVersioningScheme string = ''

@description('Major version number when versioning scheme is Latest')
param distributeVersioningMajor int = -1

@description('For ManagedImage: resource ID for the output image')
param managedImageId string = ''

@description('For ManagedImage: location')
param managedImageLocation string = ''

@description('For VHD: custom Azure Storage URI for the output VHD blob')
param vhdUri string = ''

// === OPTIONAL SETTINGS ===
param vmSize string = 'Standard_D4_v5'
param buildTimeoutInMinutes int = 150
param osDiskSizeGB int = 0

@description('Subnet resource ID for the build VM (optional)')
param vnetSubnetId string = ''

@description('Subnet resource ID for ACI in isolated builds (new in 2024-02-01). Must be on the same VNet, delegated to ACI.')
param containerInstanceSubnetId string = ''

@description('Proxy VM size when using VNet without containerInstanceSubnetId')
param proxyVmSize string = ''

@description('User-assigned identity resource IDs to assign to the build VM (for Key Vault access etc.)')
param buildVmIdentities array = []

@description('Enable VM boot optimization to improve image creation time')
param optimizeVmBoot bool = false

@description('Auto-run the image build when the template is created')
param autoRun bool = false

@description('Resource ID of a pre-existing staging resource group')
param stagingResourceGroup string = ''

@description('Tags to apply to staging resources')
param managedResourceTags object = {}

@allowed(['cleanup', 'abort'])
param errorHandlingOnCustomizerError string = 'cleanup'

@allowed(['cleanup', 'abort'])
param errorHandlingOnValidationError string = 'cleanup'

@description('Enable image validation')
param enableValidation bool = false

@description('Validate source image directly without building')
param sourceValidationOnly bool = false

@description('Continue distributing even if validation fails')
param continueDistributeOnFailure bool = false

@description('Inline validation script (PowerShell or Shell depending on provisioner)')
param validationInlineScript string = ''

param tags object = {}

// === SOURCE CONSTRUCTION ===
var platformImageSource = {
  type: 'PlatformImage'
  publisher: imagePublisher
  offer: imageOffer
  sku: imageSku
  version: imageVersion
}

var managedImageSource = {
  type: 'ManagedImage'
  imageId: sourceImageId
}

var sharedImageVersionSource = {
  type: 'SharedImageVersion'
  imageVersionId: sourceImageVersionId
}

var source = sourceType == 'PlatformImage' ? platformImageSource : (sourceType == 'ManagedImage' ? managedImageSource : sharedImageVersionSource)

// === CUSTOMIZER CONSTRUCTION ===
// Step 1: Download and extract build artifacts (replicates AzureImageBuilderTask@1 packagePath behavior)
var downloadArtifactsScript = provisioner == 'powershell' ? [
  'New-Item -Path C:\\BuildArtifacts -ItemType Directory -Force'
  'Write-Host "Downloading build artifacts..."'
  '$ProgressPreference = "SilentlyContinue"'
  'Invoke-WebRequest -Uri \'${artifactsBlobUrl}\' -OutFile C:\\BuildArtifacts\\${buildFolderName}.zip -UseBasicParsing'
  'Write-Host "Extracting build artifacts..."'
  'Expand-Archive -Path C:\\BuildArtifacts\\${buildFolderName}.zip -DestinationPath C:\\BuildArtifacts -Force'
  'Remove-Item C:\\BuildArtifacts\\${buildFolderName}.zip -Force'
  'Write-Host "Build artifacts ready at C:\\BuildArtifacts"'
  'Get-ChildItem -Path C:\\BuildArtifacts -Recurse -Depth 2 | ForEach-Object { Write-Host $_.FullName }'
] : [
  'mkdir -p /tmp/buildartifacts'
  'wget -q -O /tmp/buildartifacts/${buildFolderName}.tar.gz \'${artifactsBlobUrl}\''
  'tar -xzf /tmp/buildartifacts/${buildFolderName}.tar.gz -C /tmp/buildartifacts'
  'rm -f /tmp/buildartifacts/${buildFolderName}.tar.gz'
]

var downloadCustomizer = {
  type: provisioner == 'powershell' ? 'PowerShell' : 'Shell'
  name: 'DownloadBuildArtifacts'
  runElevated: runElevated
  inline: downloadArtifactsScript
}

// Step 2: User's inline customizer script
var inlineCustomizer = provisioner == 'powershell' ? {
  type: 'PowerShell'
  name: 'RunCustomizations'
  runElevated: runElevated
  runAsSystem: runAsSystem
  inline: split(inlineCustomizerScript, '\n')
} : {
  type: 'Shell'
  name: 'RunCustomizations'
  inline: split(inlineCustomizerScript, '\n')
}

// Step 3: Windows Update (optional, PowerShell only)
var windowsUpdateWait = {
  type: 'PowerShell'
  name: 'PreWindowsUpdateWait'
  runElevated: true
  inline: [
    'Start-Sleep -Seconds 300'
  ]
}

var windowsUpdateCustomizer = {
  type: 'WindowsUpdate'
  searchCriteria: 'IsInstalled=0'
  filters: [
    'exclude:$_.Title -like \'*Preview*\''
    'include:$true'
  ]
}

var baseCustomizers = [
  downloadCustomizer
  inlineCustomizer
]

var customizers = windowsUpdate && provisioner == 'powershell' ? concat(baseCustomizers, [
  windowsUpdateWait
  windowsUpdateCustomizer
]) : baseCustomizers

// === DISTRIBUTE CONSTRUCTION ===
var regionNames = split(targetRegions, ',')

// Build targetRegions array with optional storageAccountType
var targetRegionsArray = [for region in regionNames: union(
  { name: trim(region) },
  !empty(distributeStorageAccountType) ? { storageAccountType: distributeStorageAccountType } : {}
)]

// Versioning for SharedImage
var versioningLatest = distributeVersioningMajor >= 0 ? { scheme: 'Latest', major: distributeVersioningMajor } : { scheme: 'Latest' }
var versioningSource = { scheme: 'Source' }
var versioning = distributeVersioningScheme == 'Latest' ? versioningLatest : (distributeVersioningScheme == 'Source' ? versioningSource : null)

var sharedImageDistributeBase = {
  type: 'SharedImage'
  galleryImageId: galleryImageId
  targetRegions: targetRegionsArray
  excludeFromLatest: excludeFromLatest
  runOutputName: 'SharedImage_distribute'
}

var sharedImageDistribute = versioning != null ? union(sharedImageDistributeBase, { versioning: versioning }) : sharedImageDistributeBase

var managedImageDistribute = {
  type: 'ManagedImage'
  imageId: managedImageId
  location: managedImageLocation
  runOutputName: 'ManagedImage_distribute'
}

var vhdDistributeBase = {
  type: 'VHD'
  runOutputName: 'VHD_distribute'
}

var vhdDistribute = !empty(vhdUri) ? union(vhdDistributeBase, { uri: vhdUri }) : vhdDistributeBase

var distribute = distributeType == 'SharedImage' ? sharedImageDistribute : (distributeType == 'ManagedImage' ? managedImageDistribute : vhdDistribute)

// === VM PROFILE ===
var baseVmProfile = union(
  { vmSize: vmSize },
  osDiskSizeGB > 0 ? { osDiskSizeGB: osDiskSizeGB } : {}
)

// VNet config with optional containerInstanceSubnetId and proxyVmSize
var vnetConfigBase = { subnetId: vnetSubnetId }
var vnetConfigWithAci = !empty(containerInstanceSubnetId) ? union(vnetConfigBase, { containerInstanceSubnetId: containerInstanceSubnetId }) : vnetConfigBase
var vnetConfigFull = (!empty(proxyVmSize) && empty(containerInstanceSubnetId)) ? union(vnetConfigWithAci, { proxyVmSize: proxyVmSize }) : vnetConfigWithAci

var vmProfileWithVnet = union(baseVmProfile, { vnetConfig: vnetConfigFull })

var vmProfilePreIdentity = !empty(vnetSubnetId) ? vmProfileWithVnet : baseVmProfile

// Build VM user-assigned identities
var vmProfile = !empty(buildVmIdentities) ? union(vmProfilePreIdentity, { userAssignedIdentities: buildVmIdentities }) : vmProfilePreIdentity

// === VALIDATION ===
var validationPsCustomizer = {
  type: 'PowerShell'
  name: 'ValidationScript'
  runElevated: runElevated
  inline: split(validationInlineScript, '\n')
}

var validationShellCustomizer = {
  type: 'Shell'
  name: 'ValidationScript'
  inline: split(validationInlineScript, '\n')
}

var validationCustomizer = provisioner == 'powershell' ? validationPsCustomizer : validationShellCustomizer

var validateBlock = {
  continueDistributeOnFailure: continueDistributeOnFailure
  sourceValidationOnly: sourceValidationOnly
  inVMValidations: !empty(validationInlineScript) ? [ validationCustomizer ] : []
}

// === IMAGE TEMPLATE RESOURCE ===
resource imageTemplate 'Microsoft.VirtualMachineImages/imageTemplates@2024-02-01' = {
  name: templateName
  location: location
  tags: tags
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${userAssignedIdentityId}': {}
    }
  }
  properties: union(
    {
      source: source
      customize: customizers
      distribute: [
        distribute
      ]
      vmProfile: vmProfile
      buildTimeoutInMinutes: buildTimeoutInMinutes
      errorHandling: {
        onCustomizerError: errorHandlingOnCustomizerError
        onValidationError: errorHandlingOnValidationError
      }
    },
    optimizeVmBoot ? { optimize: { vmBoot: { state: 'Enabled' } } } : {},
    enableValidation ? { validate: validateBlock } : {},
    autoRun ? { autoRun: { state: 'Enabled' } } : {},
    !empty(stagingResourceGroup) ? { stagingResourceGroup: stagingResourceGroup } : {},
    !empty(managedResourceTags) ? { managedResourceTags: managedResourceTags } : {}
  )
}

// === OUTPUTS ===
output templateId string = imageTemplate.id
output templateName string = imageTemplate.name
