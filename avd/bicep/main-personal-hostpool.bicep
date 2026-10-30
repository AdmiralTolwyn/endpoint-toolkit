/*
  SYNOPSIS: Full-Stack AVD Personal Host Pool Orchestrator (Entra ID Join)
  FILE:     bicep/main-personal-hostpool.bicep
  DESCRIPTION:
    End-to-end modular Bicep orchestrator comparing to and streamlining the
    AVD Landing Zone Accelerator (LZA) for Personal Desktops in Germany West Central:
      1. Management Plane (modules/hostPool-entraid.bicep):
         - Personal Host Pool (Automatic or Direct assignment, Persistent load balancing,
           Start VM on Connect, Entra ID RDP properties, Scheduled Agent Updates,
           in-template Registration Token generation)
         - Desktop Application Group + Entra ID User Group RBAC
         - Optional AVD Workspace + Log Analytics Diagnostic Settings
      2. Personal Scaling Plan (modules/scalingPlan-personal.bicep):
         - Personal desktop schedules (Deallocate/Hibernate on disconnect/logoff,
           Start VM on Connect, weekday auto-start in 'W. Europe Standard Time')
         - Optional Start VM on Connect / Autoscale RBAC role assignment
      3. Session Hosts (modules/sessionHost-entraid.bicep):
         - Supports up to 12-character vmPrefix + 3-digit zero-padded sequence
           (e.g., 'FSAVDP11' + 002 -> 'FSAVDP11002', or 12-char prefix + 001..999 = 15 chars)
         - Supports Custom Gallery Image ID or Azure Marketplace Windows 11 image
         - Supports separate Compute Resource Group or same Resource Group
         - LZA security & monitoring options (Trusted Launch, Encryption at Host,
           Customer-Managed Key DES, ASGs, AMA + DCR, Antimalware, Intune, VM Login RBAC)
*/

targetScope = 'resourceGroup'

// ============================================================================
// 1. REGION & TIMEZONE PARAMETERS (Default: West Germany / Frankfurt)
// ============================================================================
@description('Target Azure region for session host compute (germanywestcentral = West Germany / Frankfurt)')
param location string = 'germanywestcentral'

@description('Target Azure region for AVD management plane objects (defaults to location)')
param managementPlaneLocation string = location

@description('Windows time zone ID for session hosts, scaling plans, and maintenance windows (W. Europe Standard Time = Germany/Berlin/Frankfurt)')
param timeZone string = 'W. Europe Standard Time'

// ============================================================================
// 2. MANAGEMENT PLANE PARAMETERS (Host Pool, App Group, Workspace)
// ============================================================================
@minLength(3)
@maxLength(64)
@description('Name of the AVD Personal Host Pool (e.g., vdpool-fsavd-prod-dewc-001)')
param hostPoolName string

@description('Friendly (display) name for the Host Pool')
param hostPoolFriendlyName string = ''

@description('Description for the Host Pool')
param hostPoolDescription string = ''

@allowed([
  'Automatic'
  'Direct'
])
@description('Personal desktop assignment type (Automatic assigns first available host on connect; Direct requires explicit admin assignment)')
param personalDesktopAssignmentType string = 'Automatic'

@description('Enable Start VM on Connect on the Host Pool')
param startVMOnConnect bool = true

@description('Set Host Pool as a validation environment')
param validationEnvironment bool = false

@allowed([
  'Enabled'
  'EnabledForClientsOnly'
  'EnabledForSessionHostsOnly'
  'Disabled'
])
@description('Public network access setting for the Host Pool')
param hostPoolPublicNetworkAccess string = 'Enabled'

@description('Base custom RDP properties for the Host Pool (targetisaadjoined:i:1;enablerdsaadauth:i:1 appended automatically when enforceEntraIdRdpProperties is true)')
param customRdpProperties string = 'audiocapturemode:i:1;audiomode:i:0;drivestoredirect:s:;redirectclipboard:i:1;redirectcomports:i:1;redirectprinters:i:1;redirectsmartcards:i:1;screen mode id:i:2'

@description('Automatically enforce targetisaadjoined:i:1 and enablerdsaadauth:i:1 in customRdpProperties')
param enforceEntraIdRdpProperties bool = true

@minValue(1)
@maxValue(648)
@description('Registration token validity in hours (1 to 648h / 27 days)')
param tokenExpirationHours int = 24

@description('Enable scheduled maintenance windows for AVD agent updates')
param enableAgentScheduledUpdate bool = true

@description('Use session host local time zone for AVD agent maintenance windows')
param agentUpdateUseSessionHostLocalTime bool = true

@description('Maintenance windows for AVD agent updates')
param agentUpdateMaintenanceWindows array = [
  {
    dayOfWeek: 'Tuesday'
    hour: 18
  }
  {
    dayOfWeek: 'Friday'
    hour: 17
  }
]

@description('Deploy the Desktop Application Group for this Host Pool')
param deployApplicationGroup bool = true

@description('Custom name for the Desktop Application Group (defaults to vdag-desktop-<hostPoolName>)')
param applicationGroupName string = ''

@description('Friendly name displayed to users in Windows App / Remote Desktop feed')
param applicationGroupFriendlyName string = ''

@description('Optional array of Entra ID Group or User Principal (Object) IDs to assign "Desktop Virtualization User" on the Application Group')
param desktopUserPrincipalIds array = []

@allowed([
  'Group'
  'User'
  'ServicePrincipal'
])
@description('Principal type for desktopUserPrincipalIds (default: Group; set to User when assigning individual users)')
param desktopUserPrincipalType string = 'Group'

@description('Deploy an AVD Workspace and register the Desktop Application Group')
param deployWorkspace bool = true

@description('Custom name for the AVD Workspace (defaults to vdws-<hostPoolName>)')
param workspaceName string = ''

@description('Friendly name for the AVD Workspace')
param workspaceFriendlyName string = ''

@allowed([
  'Enabled'
  'Disabled'
])
@description('Public network access setting for the Workspace')
param workspacePublicNetworkAccess string = 'Enabled'

@description('Optional existing Application Group resource IDs to include in the Workspace alongside this pool')
param additionalApplicationGroupReferences array = []

// ============================================================================
// 3. PERSONAL SCALING PLAN PARAMETERS
// ============================================================================
@description('Deploy a Personal Desktop Scaling Plan and attach it to the Host Pool')
param deployScalingPlan bool = true

@description('Custom name for the Scaling Plan (defaults to vdscaling-<hostPoolName>)')
param scalingPlanName string = ''

@description('Friendly name for the Scaling Plan')
param scalingPlanFriendlyName string = ''

@allowed([
  'None'
  'WithAssignedUser'
  'All'
])
@description('Which personal desktops to auto-start at weekday ramp-up')
param weekdayRampUpAutoStartHosts string = 'WithAssignedUser'

@allowed([
  'None'
  'Hibernate'
  'Deallocate'
])
@description('Power action after user disconnect (if set to Hibernate, enableHibernation on VMs is also enabled automatically)')
param actionOnDisconnect string = 'Deallocate'

@minValue(0)
@description('Minutes to wait after disconnect before deallocating/hibernating')
param minutesToWaitOnDisconnect int = 30

@allowed([
  'None'
  'Hibernate'
  'Deallocate'
])
@description('Power action after user logoff')
param actionOnLogoff string = 'Deallocate'

@minValue(0)
@description('Minutes to wait after logoff before deallocating/hibernating')
param minutesToWaitOnLogoff int = 15

@description('Optional custom scaling schedule array override')
param customScalingSchedules array = []

@description('Optional Object ID of the "Azure Virtual Desktop" enterprise application in Entra ID to assign "Desktop Virtualization Power On Off Contributor"')
param avdServicePrincipalObjectId string = ''

@allowed([
  'Subscription'
  'ResourceGroup'
])
@description('Scope for the "Desktop Virtualization Power On Off Contributor" role assignment when avdServicePrincipalObjectId is provided. Microsoft Autoscale requires Subscription scope; ResourceGroup scope is supported when Subscription-level role assignment is pre-provisioned or only standalone Start VM on Connect is used.')
param powerOnOffRbacScope string = 'Subscription'

// ============================================================================
// 4. SESSION HOST NAMING (12-CHAR PREFIX SUPPORT) & COMPUTE PARAMETERS
// ============================================================================
@description('Deploy session host VMs in this deployment run')
param deploySessionHosts bool = true

@description('Resource Group where session host VMs and NICs will be deployed (defaults to the current Resource Group)')
param computeResourceGroupName string = resourceGroup().name

@maxLength(12)
@description('Session host prefix — supports up to 12 characters (e.g., "FSAVDP11" or a 12-char prefix). Combined with 3-digit padding (001..999) to stay within the 15-char Windows computerName limit.')
param vmPrefix string = 'FSAVDP11'

@minValue(0)
@maxValue(999)
@description('Number of session host VMs to deploy when vmList is empty')
param vmCount int = 2

@minValue(1)
@maxValue(999)
@description('Starting number for generated session hosts (e.g., 1 -> FSAVDP11001, 2 -> FSAVDP11002)')
param vmStartIndex int = 1

@minValue(1)
@maxValue(4)
@description('Zero-padding width for the VM sequence number. Default is 3 (001..999) so a 12-character vmPrefix fits within 15 characters.')
param vmIndexPadLength int = 3

@description('Optional explicit list of VM names (max 15 chars each). Overrides vmPrefix + vmCount when non-empty.')
param vmList array = []

@description('Subnet resource ID for session host NICs (required when deploySessionHosts is true)')
param subnetId string = ''

@description('Local administrator username for session hosts')
param localAdminUser string = 'avdadmin'

@secure()
@description('Local administrator password for session hosts (required when deploySessionHosts is true)')
param localAdminPassword string = ''

@description('VM size SKU for Personal Desktops')
param vmSize string = 'Standard_D4s_v5'

@description('OS disk storage type (Premium_LRS or StandardSSD_LRS recommended)')
param osDiskType string = 'Premium_LRS'

@description('OS disk size in GB')
param osDiskSizeGB int = 128

@allowed([
  'ReadWrite'
  'ReadOnly'
  'None'
])
param osDiskCaching string = 'ReadWrite'

@description('Availability Zones for round-robin placement in germanywestcentral (e.g., ["1", "2", "3"], or [] for regional)')
param availabilityZones array = [
  '1'
  '2'
  '3'
]

@description('Enable Accelerated Networking on session host NICs (supported on D/E-series v4/v5/v6 SKUs)')
param acceleratedNetworking bool = true

@description('Automatically enroll Entra ID-joined Personal Desktops in Microsoft Intune')
param enableIntune bool = true

// --- IMAGE SOURCE (Custom Gallery Image OR Marketplace Windows 11) ---
@description('Optional Azure Compute Gallery image version resource ID. Leave empty to deploy from Azure Marketplace.')
param imageResourceId string = ''

@description('Marketplace image publisher (used when imageResourceId is empty)')
param imagePublisher string = 'MicrosoftWindowsDesktop'

@description('Marketplace image offer (e.g., windows-11 for single-session or office-365 for M365 preinstalled)')
param imageOffer string = 'windows-11'

@description('Marketplace image SKU (e.g., win11-24h2-ent for Personal single-session or win11-24h2-avd-m365)')
param imageSku string = 'win11-24h2-ent'

@description('Marketplace image version')
param marketplaceImageVersion string = 'latest'

@description('Optional image name tag value (auto-derived if empty)')
param imageName string = ''

@description('Optional image version tag value (auto-derived if empty)')
param imageVersion string = ''

// --- LZA ZERO-TRUST, MONITORING & RBAC OPTIONS ---
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

@description('Optional Entra ID Principal IDs to grant "Virtual Machine User Login" on each session host (defaults to desktopUserPrincipalIds when empty)')
param vmUserLoginPrincipalIds array = []

@description('Optional Entra ID Principal IDs to grant "Virtual Machine Administrator Login" on each session host')
param vmAdminLoginPrincipalIds array = []

@allowed([
  'Group'
  'User'
  'ServicePrincipal'
])
@description('Default principal type for VM login role assignments (defaults to desktopUserPrincipalType)')
param vmLoginPrincipalType string = desktopUserPrincipalType

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

// Pinned AVD DSC agent package that includes -AadJoin support; override to pin a different build.
#disable-next-line no-hardcoded-env-urls
param avdAgentPackageUrl string = 'https://wvdportalstorageblob.blob.core.windows.net/galleryartifacts/Configuration_1.0.03299.1133.zip'

// ============================================================================
// 5. DIAGNOSTICS & TAGGING
// ============================================================================
@description('Optional Log Analytics Workspace resource ID for management plane and scaling plan diagnostics')
param logAnalyticsWorkspaceId string = ''

@description('Timestamp used for token calculation and CreatedOn tag')
param currentTimestamp string = utcNow('yyyy-MM-dd')

@description('Base UTC timestamp used for registration token expiration')
param baseTime string = utcNow('u')

@description('Additional custom tags to merge onto all deployed resources')
param customTags object = {}

// ============================================================================
// VARIABLES
// ============================================================================
var generatedVmList = [
  for i in range(0, vmCount): '${vmPrefix}${padLeft((vmStartIndex + i), vmIndexPadLength, '0')}'
]
var rawVmList = [for vmName in (!empty(vmList) ? vmList : generatedVmList): string(vmName)]
// Enforce the 15-character Windows NetBIOS computerName limit across all names upfront (compatible with Bicep 0.28+ without fail())
var invalidVmNames = deploySessionHosts ? filter(rawVmList, name => empty(name) || length(name) > 15) : []
var vmNameValidationGuard = empty(invalidVmNames)
  ? 0
  : int('InvalidSessionHostNameExceeds15CharLimit_${join(invalidVmNames, ',')}')
var resolvedVmList = vmNameValidationGuard == 0 ? rawVmList : []

var defaultScalingCandidate = 'vdscaling-${hostPoolName}'
var resolvedScalingPlanName = !empty(scalingPlanName)
  ? scalingPlanName
  : take(
      length(defaultScalingCandidate) <= 64
        ? defaultScalingCandidate
        : '${take(defaultScalingCandidate, 55)}-${uniqueString(hostPoolName)}',
      64
    )
var resolvedEnableHibernation = enableHibernation || actionOnDisconnect == 'Hibernate' || actionOnLogoff == 'Hibernate' || contains(toLower(string(customScalingSchedules)), '"hibernate"')
var resolvedVmUserLoginPrincipalIds = !empty(vmUserLoginPrincipalIds)
  ? vmUserLoginPrincipalIds
  : desktopUserPrincipalIds
var assignPowerOnOffRbac = (deployScalingPlan || startVMOnConnect) && !empty(avdServicePrincipalObjectId)
var isSeparateComputeRg = toLower(computeResourceGroupName) != toLower(resourceGroup().name)

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

var managementTags = union(
  {
    Environment: 'AVD-Personal'
    Workload: 'PersonalDesktop'
    Directory: 'EntraID'
    Region: location
    CreatedOn: currentTimestamp
  },
  customTags
)

var sessionHostTags = union(
  managementTags,
  {
    ImageName: resolvedImageName
    ImageVersion: resolvedImageVersion
  }
)

// ============================================================================
// MODULE 0: POWER ON/OFF RBAC (Subscription Scope Default, or Management + Compute RGs)
// ============================================================================
var subRbacDeployCandidate = 'deploy-powerrbac-sub-${hostPoolName}'
module powerRbacSubscription 'modules/powerOnOffRbac-subscription.bicep' = if (assignPowerOnOffRbac && powerOnOffRbacScope == 'Subscription') {
  name: take(
    length(subRbacDeployCandidate) <= 64
      ? subRbacDeployCandidate
      : '${take(subRbacDeployCandidate, 55)}-${uniqueString(hostPoolName)}',
    64
  )
  scope: subscription()
  params: {
    avdServicePrincipalObjectId: avdServicePrincipalObjectId
  }
}

var mgmtRbacDeployCandidate = 'deploy-powerrbac-mgmt-${hostPoolName}'
module powerRbacManagementRg 'modules/powerOnOffRbac.bicep' = if (assignPowerOnOffRbac && powerOnOffRbacScope == 'ResourceGroup') {
  name: take(
    length(mgmtRbacDeployCandidate) <= 64
      ? mgmtRbacDeployCandidate
      : '${take(mgmtRbacDeployCandidate, 55)}-${uniqueString(hostPoolName)}',
    64
  )
  params: {
    avdServicePrincipalObjectId: avdServicePrincipalObjectId
  }
}

var compRbacDeployCandidate = 'deploy-powerrbac-comp-${hostPoolName}'
module powerRbacComputeRg 'modules/powerOnOffRbac.bicep' = if (assignPowerOnOffRbac && powerOnOffRbacScope == 'ResourceGroup' && isSeparateComputeRg) {
  name: take(
    length(compRbacDeployCandidate) <= 64
      ? compRbacDeployCandidate
      : '${take(compRbacDeployCandidate, 55)}-${uniqueString(hostPoolName)}',
    64
  )
  scope: resourceGroup(computeResourceGroupName)
  params: {
    avdServicePrincipalObjectId: avdServicePrincipalObjectId
  }
}

// ============================================================================
// MODULE 1: AVD MANAGEMENT PLANE (Host Pool, App Group, Workspace)
// ============================================================================
var mgmtDeployCandidate = vmNameValidationGuard == 0 ? 'deploy-mgmt-${hostPoolName}' : 'deploy-mgmt-invalid'
module managementPlane 'modules/hostPool-entraid.bicep' = {
  name: take(
    length(mgmtDeployCandidate) <= 64
      ? mgmtDeployCandidate
      : '${take(mgmtDeployCandidate, 55)}-${uniqueString(hostPoolName)}',
    64
  )
  params: {
    location: managementPlaneLocation
    hostPoolName: hostPoolName
    hostPoolFriendlyName: hostPoolFriendlyName
    hostPoolDescription: hostPoolDescription
    hostPoolType: 'Personal'
    personalDesktopAssignmentType: personalDesktopAssignmentType
    loadBalancerType: 'Persistent'
    maxSessionLimit: 999999
    startVMOnConnect: startVMOnConnect
    validationEnvironment: validationEnvironment
    hostPoolPublicNetworkAccess: hostPoolPublicNetworkAccess
    customRdpProperties: customRdpProperties
    enforceEntraIdRdpProperties: enforceEntraIdRdpProperties
    tokenExpirationHours: tokenExpirationHours
    baseTime: baseTime
    enableAgentScheduledUpdate: enableAgentScheduledUpdate
    agentUpdateUseSessionHostLocalTime: agentUpdateUseSessionHostLocalTime
    agentUpdateTimeZone: timeZone
    agentUpdateMaintenanceWindows: agentUpdateMaintenanceWindows
    deployApplicationGroup: deployApplicationGroup
    applicationGroupName: applicationGroupName
    applicationGroupFriendlyName: applicationGroupFriendlyName
    desktopUserPrincipalIds: desktopUserPrincipalIds
    desktopUserPrincipalType: desktopUserPrincipalType
    deployWorkspace: deployWorkspace
    workspaceName: workspaceName
    workspaceFriendlyName: workspaceFriendlyName
    workspacePublicNetworkAccess: workspacePublicNetworkAccess
    additionalApplicationGroupReferences: additionalApplicationGroupReferences
    logAnalyticsWorkspaceId: logAnalyticsWorkspaceId
    tags: managementTags
  }
}

// ============================================================================
// MODULE 2: PERSONAL DESKTOP SCALING PLAN
// ============================================================================
var scalingDeployCandidate = 'deploy-scaling-${hostPoolName}'
module scalingPlan 'modules/scalingPlan-personal.bicep' = if (deployScalingPlan) {
  name: take(
    length(scalingDeployCandidate) <= 64
      ? scalingDeployCandidate
      : '${take(scalingDeployCandidate, 55)}-${uniqueString(hostPoolName)}',
    64
  )
  dependsOn: [
    powerRbacSubscription
    powerRbacManagementRg
    powerRbacComputeRg
  ]
  params: {
    location: managementPlaneLocation
    scalingPlanName: resolvedScalingPlanName
    scalingPlanFriendlyName: scalingPlanFriendlyName
    timeZone: timeZone
    hostPoolResourceId: managementPlane.outputs.hostPoolId
    scalingPlanEnabled: true
    startVMOnConnect: startVMOnConnect
    weekdayRampUpAutoStartHosts: weekdayRampUpAutoStartHosts
    actionOnDisconnect: actionOnDisconnect
    minutesToWaitOnDisconnect: minutesToWaitOnDisconnect
    actionOnLogoff: actionOnLogoff
    minutesToWaitOnLogoff: minutesToWaitOnLogoff
    customSchedules: customScalingSchedules
    // Handled by powerRbacSubscription / powerRbacManagementRg / powerRbacComputeRg above
    avdServicePrincipalObjectId: ''
    logAnalyticsWorkspaceId: logAnalyticsWorkspaceId
    tags: managementTags
  }
}

// ============================================================================
// MODULE 3: SESSION HOSTS (12-char prefix + 3-digit index, Entra ID Join)
// ============================================================================
@batchSize(5)
module sessionHosts 'modules/sessionHost-entraid.bicep' = [
  for (vmName, i) in (deploySessionHosts ? resolvedVmList : []): {
    name: 'deploy-${vmName}'
    scope: resourceGroup(computeResourceGroupName)
    params: {
      location: location
      vmName: string(vmName)
      imageResourceId: imageResourceId
      subnetId: subnetId
      adminUsername: localAdminUser
      adminPassword: localAdminPassword
      hostPoolName: managementPlane.outputs.hostPoolName
      hostPoolRg: resourceGroup().name
      // Empty token instructs sessionHost-entraid.bicep to securely fetch the
      // newly generated registration token directly from the Host Pool resource.
      hostPoolToken: ''
      tags: sessionHostTags

      // VM & Disk Specs
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
      enableHibernation: resolvedEnableHibernation
      applicationSecurityGroupIds: applicationSecurityGroupIds
      deployAntiMalware: deployAntiMalware
      deployMonitoringAgent: deployMonitoringAgent
      dataCollectionRuleId: dataCollectionRuleId
      installGpuDriver: installGpuDriver
      vmUserLoginPrincipalIds: resolvedVmUserLoginPrincipalIds
      vmAdminLoginPrincipalIds: vmAdminLoginPrincipalIds
      vmLoginPrincipalType: vmLoginPrincipalType
      vmUserLoginPrincipalType: vmUserLoginPrincipalType
      vmAdminLoginPrincipalType: vmAdminLoginPrincipalType
    }
  }
]

// ============================================================================
// OUTPUTS
// ============================================================================
output hostPoolId string = managementPlane.outputs.hostPoolId
output hostPoolName string = managementPlane.outputs.hostPoolName
output applicationGroupId string = managementPlane.outputs.applicationGroupId
output workspaceId string = managementPlane.outputs.workspaceId
output scalingPlanId string = deployScalingPlan ? scalingPlan!.outputs.scalingPlanId : ''
output deployedSessionHostNames array = deploySessionHosts ? resolvedVmList : []
