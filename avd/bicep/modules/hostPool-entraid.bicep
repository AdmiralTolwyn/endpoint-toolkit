/*
  SYNOPSIS: AVD Management Plane Module (Host Pool, Application Group, Workspace, Diagnostics & RBAC)
  FILE:     bicep/modules/hostPool-entraid.bicep
  DESCRIPTION:
    Deploys an Azure Virtual Desktop Host Pool (defaulting to Personal desktops
    in Germany West Central), Desktop Application Group, optional Workspace,
    Log Analytics Diagnostic Settings, and RBAC assignments.
    Aligned with Azure Virtual Desktop Landing Zone Accelerator (LZA) patterns
    while remaining lightweight, modular, and fully parametrized.
*/

targetScope = 'resourceGroup'

// === CORE PARAMETERS ===
@description('Azure region for AVD management plane objects (e.g., germanywestcentral)')
param location string = 'germanywestcentral'

@minLength(3)
@maxLength(64)
@description('Name of the AVD Host Pool')
param hostPoolName string

@description('Friendly (display) name for the Host Pool')
param hostPoolFriendlyName string = ''

@description('Description for the Host Pool')
param hostPoolDescription string = ''

// === HOST POOL CONFIGURATION ===
@allowed([
  'Personal'
  'Pooled'
])
@description('Host Pool type. Defaults to Personal for dedicated 1:1 desktops.')
param hostPoolType string = 'Personal'

@allowed([
  'Automatic'
  'Direct'
])
@description('Assignment type for Personal Host Pools (ignored when hostPoolType is Pooled)')
param personalDesktopAssignmentType string = 'Automatic'

@allowed([
  'Persistent'
  'BreadthFirst'
  'DepthFirst'
])
@description('Load balancer type. Automatically forced to Persistent when hostPoolType is Personal.')
param loadBalancerType string = 'Persistent'

@minValue(1)
@maxValue(999999)
@description('Maximum sessions per host (automatically set to 999999 for Personal pools)')
param maxSessionLimit int = 999999

@description('Enable Start VM on Connect (recommended for Personal desktops with deallocate/hibernate)')
param startVMOnConnect bool = true

@description('Set Host Pool as a validation environment to receive early AVD service updates')
param validationEnvironment bool = false

@allowed([
  'Enabled'
  'EnabledForClientsOnly'
  'EnabledForSessionHostsOnly'
  'Disabled'
])
@description('Public network access control for the Host Pool')
param hostPoolPublicNetworkAccess string = 'Enabled'

@description('Base custom RDP properties for the Host Pool')
param customRdpProperties string = 'audiocapturemode:i:1;audiomode:i:0;drivestoredirect:s:;redirectclipboard:i:1;redirectcomports:i:1;redirectprinters:i:1;redirectsmartcards:i:1;screen mode id:i:2'

@description('Automatically append targetisaadjoined:i:1;enablerdsaadauth:i:1 for Entra ID joined session hosts')
param enforceEntraIdRdpProperties bool = true

// === REGISTRATION TOKEN PARAMETERS ===
@minValue(1)
@maxValue(648)
@description('Validity duration in hours for the generated Host Pool registration token (1 to 648 hours / 27 days)')
param tokenExpirationHours int = 24

@description('Base UTC timestamp used to calculate registration token expiry')
param baseTime string = utcNow('u')

// === AGENT SCHEDULED UPDATES (LZA PATTERN) ===
@description('Enable scheduled maintenance windows for the AVD agent / side-by-side stack')
param enableAgentScheduledUpdate bool = true

@description('Use session host local time zone for AVD agent maintenance windows')
param agentUpdateUseSessionHostLocalTime bool = true

@description('Windows time zone name for AVD agent updates when not using session host local time (W. Europe Standard Time = Germany/Berlin)')
param agentUpdateTimeZone string = 'W. Europe Standard Time'

@description('Maintenance windows for AVD agent updates (array of { dayOfWeek, hour })')
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

// === APPLICATION GROUP PARAMETERS ===
@description('Deploy a Desktop Application Group linked to the Host Pool')
param deployApplicationGroup bool = true

@description('Name of the Desktop Application Group (defaults to vdag-desktop-<hostPoolName>)')
param applicationGroupName string = ''

@description('Friendly name displayed to users in the Windows App / AVD feed')
param applicationGroupFriendlyName string = ''

@description('Description for the Desktop Application Group')
param applicationGroupDescription string = ''

@description('Optional array of Entra ID Group or User Principal (Object) IDs to grant "Desktop Virtualization User" on the Application Group')
param desktopUserPrincipalIds array = []

@allowed([
  'Group'
  'User'
  'ServicePrincipal'
])
@description('Principal type for desktopUserPrincipalIds (default: Group; set to User when assigning individual users to a Personal Host Pool)')
param desktopUserPrincipalType string = 'Group'

// === WORKSPACE PARAMETERS ===
@description('Deploy an AVD Workspace and register the Application Group')
param deployWorkspace bool = true

@description('Name of the AVD Workspace (defaults to vdws-<hostPoolName>)')
param workspaceName string = ''

@description('Friendly name for the AVD Workspace')
param workspaceFriendlyName string = ''

@description('Description for the AVD Workspace')
param workspaceDescription string = ''

@allowed([
  'Enabled'
  'Disabled'
])
@description('Public network access control for the Workspace')
param workspacePublicNetworkAccess string = 'Enabled'

@description('Additional existing Application Group resource IDs to preserve/include in the Workspace')
param additionalApplicationGroupReferences array = []

// === MONITORING & TAGGING ===
@description('Optional Log Analytics Workspace resource ID for Host Pool, Application Group, and Workspace diagnostic logs')
param logAnalyticsWorkspaceId string = ''

@description('Tags to apply to all management plane resources')
param tags object = {}

// === VARIABLES ===
var isPersonal = hostPoolType == 'Personal'
var resolvedLoadBalancerType = isPersonal ? 'Persistent' : loadBalancerType
var resolvedMaxSessionLimit = isPersonal ? 999999 : maxSessionLimit

// Case-insensitively strip any existing targetisaadjoined:* or enablerdsaadauth:* entries when
// enforceEntraIdRdpProperties is true, then append canonical targetisaadjoined:i:1;enablerdsaadauth:i:1
var rawRdpEntries = [for entry in split(customRdpProperties, ';'): trim(entry)]
var keptRdpEntries = filter(
  rawRdpEntries,
  entry =>
    !empty(entry) && !(enforceEntraIdRdpProperties && (startsWith(toLower(entry), 'targetisaadjoined:') || startsWith(
      toLower(entry),
      'enablerdsaadauth:'
    )))
)
var baseRdpProperties = join(keptRdpEntries, ';')
var resolvedRdpProperties = enforceEntraIdRdpProperties
  ? (!empty(baseRdpProperties)
      ? '${baseRdpProperties};targetisaadjoined:i:1;enablerdsaadauth:i:1'
      : 'targetisaadjoined:i:1;enablerdsaadauth:i:1')
  : baseRdpProperties

var tokenExpirationTime = dateTimeAdd(baseTime, 'PT${tokenExpirationHours}H')

var defaultAppGroupCandidate = 'vdag-desktop-${hostPoolName}'
var resolvedAppGroupName = !empty(applicationGroupName)
  ? applicationGroupName
  : take(
      length(defaultAppGroupCandidate) <= 64
        ? defaultAppGroupCandidate
        : '${take(defaultAppGroupCandidate, 55)}-${uniqueString(hostPoolName)}',
      64
    )
var resolvedAppGroupFriendlyName = !empty(applicationGroupFriendlyName)
  ? applicationGroupFriendlyName
  : (!empty(hostPoolFriendlyName) ? hostPoolFriendlyName : 'Personal Desktop')

var defaultWorkspaceCandidate = 'vdws-${hostPoolName}'
var resolvedWorkspaceName = !empty(workspaceName)
  ? workspaceName
  : take(
      length(defaultWorkspaceCandidate) <= 64
        ? defaultWorkspaceCandidate
        : '${take(defaultWorkspaceCandidate, 55)}-${uniqueString(hostPoolName)}',
      64
    )
var resolvedWorkspaceFriendlyName = !empty(workspaceFriendlyName)
  ? workspaceFriendlyName
  : (!empty(hostPoolFriendlyName) ? '${hostPoolFriendlyName} Workspace' : '${hostPoolName} Workspace')

var workspaceAppGroupRefs = union(
  deployApplicationGroup ? [applicationGroup!.id] : [],
  additionalApplicationGroupReferences
)

// Built-in Role Definition ID: Desktop Virtualization User
var desktopVirtualizationUserRoleId = subscriptionResourceId(
  'Microsoft.Authorization/roleDefinitions',
  '1d18fff3-a72a-46b5-b4a9-0b38a3cd7e63'
)

// === 1. HOST POOL ===
resource hostPool 'Microsoft.DesktopVirtualization/hostPools@2024-04-03' = {
  name: hostPoolName
  location: location
  tags: tags
  properties: {
    friendlyName: !empty(hostPoolFriendlyName) ? hostPoolFriendlyName : hostPoolName
    description: hostPoolDescription
    hostPoolType: hostPoolType
    personalDesktopAssignmentType: isPersonal ? personalDesktopAssignmentType : null
    loadBalancerType: resolvedLoadBalancerType
    maxSessionLimit: resolvedMaxSessionLimit
    preferredAppGroupType: 'Desktop'
    startVMOnConnect: startVMOnConnect
    validationEnvironment: validationEnvironment
    publicNetworkAccess: hostPoolPublicNetworkAccess
    customRdpProperty: resolvedRdpProperties
    registrationInfo: {
      expirationTime: tokenExpirationTime
      registrationTokenOperation: 'Update'
    }
    agentUpdate: enableAgentScheduledUpdate
      ? {
          type: 'Scheduled'
          useSessionHostLocalTime: agentUpdateUseSessionHostLocalTime
          maintenanceWindowTimeZone: agentUpdateUseSessionHostLocalTime ? null : agentUpdateTimeZone
          maintenanceWindows: agentUpdateMaintenanceWindows
        }
      : {
          type: 'Default'
        }
  }
}

// === 2. DESKTOP APPLICATION GROUP ===
resource applicationGroup 'Microsoft.DesktopVirtualization/applicationGroups@2024-04-03' = if (deployApplicationGroup) {
  name: resolvedAppGroupName
  location: location
  tags: tags
  properties: {
    hostPoolArmPath: hostPool.id
    applicationGroupType: 'Desktop'
    friendlyName: resolvedAppGroupFriendlyName
    description: applicationGroupDescription
  }
}

// === 2b. APPLICATION GROUP RBAC (Desktop Virtualization User) ===
resource appGroupUserRoleAssignments 'Microsoft.Authorization/roleAssignments@2022-04-01' = [
  for principalId in desktopUserPrincipalIds: if (deployApplicationGroup && !empty(principalId)) {
    name: guid(applicationGroup.id, principalId, desktopVirtualizationUserRoleId)
    scope: applicationGroup
    properties: {
      roleDefinitionId: desktopVirtualizationUserRoleId
      principalId: principalId
      principalType: desktopUserPrincipalType
    }
  }
]

// === 3. WORKSPACE ===
resource workspace 'Microsoft.DesktopVirtualization/workspaces@2024-04-03' = if (deployWorkspace) {
  name: resolvedWorkspaceName
  location: location
  tags: tags
  properties: {
    friendlyName: resolvedWorkspaceFriendlyName
    description: workspaceDescription
    publicNetworkAccess: workspacePublicNetworkAccess
    applicationGroupReferences: workspaceAppGroupRefs
  }
}

// === 4. DIAGNOSTIC SETTINGS (LZA MONITORING PATTERN) ===
var hpDiagCandidate = 'diag-${hostPoolName}'
resource hostPoolDiagnostics 'Microsoft.Insights/diagnosticSettings@2021-05-01-preview' = if (!empty(logAnalyticsWorkspaceId)) {
  name: take(
    length(hpDiagCandidate) <= 64 ? hpDiagCandidate : '${take(hpDiagCandidate, 55)}-${uniqueString(hostPoolName)}',
    64
  )
  scope: hostPool
  properties: {
    workspaceId: logAnalyticsWorkspaceId
    logs: [
      {
        categoryGroup: 'allLogs'
        enabled: true
      }
    ]
  }
}

var agDiagCandidate = 'diag-${resolvedAppGroupName}'
resource appGroupDiagnostics 'Microsoft.Insights/diagnosticSettings@2021-05-01-preview' = if (deployApplicationGroup && !empty(logAnalyticsWorkspaceId)) {
  name: take(
    length(agDiagCandidate) <= 64
      ? agDiagCandidate
      : '${take(agDiagCandidate, 55)}-${uniqueString(resolvedAppGroupName)}',
    64
  )
  scope: applicationGroup
  properties: {
    workspaceId: logAnalyticsWorkspaceId
    logs: [
      {
        categoryGroup: 'allLogs'
        enabled: true
      }
    ]
  }
}

var wsDiagCandidate = 'diag-${resolvedWorkspaceName}'
resource workspaceDiagnostics 'Microsoft.Insights/diagnosticSettings@2021-05-01-preview' = if (deployWorkspace && !empty(logAnalyticsWorkspaceId)) {
  name: take(
    length(wsDiagCandidate) <= 64
      ? wsDiagCandidate
      : '${take(wsDiagCandidate, 55)}-${uniqueString(resolvedWorkspaceName)}',
    64
  )
  scope: workspace
  properties: {
    workspaceId: logAnalyticsWorkspaceId
    logs: [
      {
        categoryGroup: 'allLogs'
        enabled: true
      }
    ]
  }
}

// === OUTPUTS ===
output hostPoolId string = hostPool.id
output hostPoolName string = hostPool.name
output applicationGroupId string = deployApplicationGroup ? applicationGroup!.id : ''
output applicationGroupName string = deployApplicationGroup ? applicationGroup!.name : ''
output workspaceId string = deployWorkspace ? workspace!.id : ''
output workspaceName string = deployWorkspace ? workspace!.name : ''
