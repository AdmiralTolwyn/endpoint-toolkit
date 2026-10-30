/*
  SYNOPSIS: AVD Personal Desktop Scaling Plan & Power Management RBAC Module
  FILE:     bicep/modules/scalingPlan-personal.bicep
  DESCRIPTION:
    Deploys a Microsoft.DesktopVirtualization/scalingPlans resource configured for
    Personal Host Pools (supporting Start VM on Connect, auto-start during ramp-up,
    and Deallocate/Hibernate on disconnect or logoff), optional Log Analytics
    diagnostics, and optional RBAC assignment for the AVD service principal.
    Default time zone: 'W. Europe Standard Time' (Germany / Frankfurt / Berlin).
*/

targetScope = 'resourceGroup'

// === CORE PARAMETERS ===
@description('Azure region for the Scaling Plan (e.g., germanywestcentral)')
param location string = 'germanywestcentral'

@minLength(3)
@maxLength(64)
@description('Name of the AVD Scaling Plan')
param scalingPlanName string

@description('Friendly (display) name for the Scaling Plan')
param scalingPlanFriendlyName string = ''

@description('Description for the Scaling Plan')
param scalingPlanDescription string = ''

@description('Windows time zone ID for scaling schedules (W. Europe Standard Time = Germany/Berlin)')
param timeZone string = 'W. Europe Standard Time'

@description('Tag name used to exclude specific session hosts from scaling actions (defaults to exclude-<scalingPlanName>)')
param exclusionTag string = ''

@description('Optional Host Pool Resource ID to associate with this Scaling Plan')
param hostPoolResourceId string = ''

@description('Whether scaling is enabled on the associated Host Pool')
param scalingPlanEnabled bool = true

// === SCHEDULE CONFIGURATION (PERSONAL DESKTOPS) ===
@description('Enable Start VM on Connect across generated Personal scaling schedule phases (should match the Host Pool startVMOnConnect setting)')
param startVMOnConnect bool = true

@allowed([
  'None'
  'WithAssignedUser'
  'All'
])
@description('Which personal hosts to automatically start during weekday ramp-up')
param weekdayRampUpAutoStartHosts string = 'WithAssignedUser'

@allowed([
  'None'
  'Hibernate'
  'Deallocate'
])
@description('Power action after a user disconnects (Deallocate works on all SKUs; Hibernate requires hibernation-enabled VM/image)')
param actionOnDisconnect string = 'Deallocate'

@minValue(0)
@description('Minutes to wait after user disconnect before taking actionOnDisconnect')
param minutesToWaitOnDisconnect int = 30

@allowed([
  'None'
  'Hibernate'
  'Deallocate'
])
@description('Power action after a user logs off')
param actionOnLogoff string = 'Deallocate'

@minValue(0)
@description('Minutes to wait after user logoff before taking actionOnLogoff')
param minutesToWaitOnLogoff int = 15

@minValue(0)
@maxValue(23)
@description('Weekday ramp-up start hour (24h format)')
param rampUpStartHour int = 7

@minValue(0)
@maxValue(23)
@description('Weekday peak start hour (24h format)')
param peakStartHour int = 9

@minValue(0)
@maxValue(23)
@description('Weekday ramp-down start hour (24h format)')
param rampDownStartHour int = 18

@minValue(0)
@maxValue(23)
@description('Weekday off-peak start hour (24h format)')
param offPeakStartHour int = 20

@description('Optional custom schedules array override. Each entry should have { name: string, properties: { ...ScalingPlanPersonalScheduleProperties } }. When empty, LZA-aligned Weekday + Weekend Personal schedules are generated from the parameters above.')
param customSchedules array = []

// === RBAC & DIAGNOSTICS ===
@description('Optional Object ID of the "Azure Virtual Desktop" enterprise application (appId 9cdead84-a844-4324-93f2-b2e6bb768d07) in the tenant to assign "Desktop Virtualization Power On Off Contributor"')
param avdServicePrincipalObjectId string = ''

@allowed([
  'Subscription'
  'ResourceGroup'
])
@description('Scope for the Power On/Off Contributor role assignment when avdServicePrincipalObjectId is provided (Microsoft Autoscale requires Subscription scope)')
param powerOnOffRbacScope string = 'Subscription'

@description('Optional Log Analytics Workspace resource ID for Scaling Plan diagnostic logs')
param logAnalyticsWorkspaceId string = ''

@description('Tags to apply to the Scaling Plan')
param tags object = {}

// === VARIABLES ===
var resolvedExclusionTag = !empty(exclusionTag) ? exclusionTag : 'exclude-${scalingPlanName}'
var startVmOnConnectSetting = startVMOnConnect ? 'Enable' : 'Disable'
var weekdayCandidateName = 'Weekdays-${scalingPlanName}'
var weekendCandidateName = 'Weekend-${scalingPlanName}'

var defaultPersonalSchedules = [
  {
    name: take(
      length(weekdayCandidateName) <= 64
        ? weekdayCandidateName
        : '${take(weekdayCandidateName, 55)}-${uniqueString(scalingPlanName, 'Weekdays')}',
      64
    )
    properties: {
      daysOfWeek: [
        'Monday'
        'Tuesday'
        'Wednesday'
        'Thursday'
        'Friday'
      ]
      rampUpStartTime: {
        hour: rampUpStartHour
        minute: 0
      }
      rampUpAutoStartHosts: weekdayRampUpAutoStartHosts
      rampUpStartVMOnConnect: startVmOnConnectSetting
      rampUpActionOnDisconnect: actionOnDisconnect
      rampUpMinutesToWaitOnDisconnect: minutesToWaitOnDisconnect
      rampUpActionOnLogoff: actionOnLogoff
      rampUpMinutesToWaitOnLogoff: minutesToWaitOnLogoff
      peakStartTime: {
        hour: peakStartHour
        minute: 0
      }
      peakStartVMOnConnect: startVmOnConnectSetting
      peakActionOnDisconnect: actionOnDisconnect
      peakMinutesToWaitOnDisconnect: minutesToWaitOnDisconnect
      peakActionOnLogoff: actionOnLogoff
      peakMinutesToWaitOnLogoff: minutesToWaitOnLogoff
      rampDownStartTime: {
        hour: rampDownStartHour
        minute: 0
      }
      rampDownStartVMOnConnect: startVmOnConnectSetting
      rampDownActionOnDisconnect: actionOnDisconnect
      rampDownMinutesToWaitOnDisconnect: minutesToWaitOnDisconnect
      rampDownActionOnLogoff: actionOnLogoff
      rampDownMinutesToWaitOnLogoff: minutesToWaitOnLogoff
      offPeakStartTime: {
        hour: offPeakStartHour
        minute: 0
      }
      offPeakStartVMOnConnect: startVmOnConnectSetting
      offPeakActionOnDisconnect: actionOnDisconnect
      offPeakMinutesToWaitOnDisconnect: minutesToWaitOnDisconnect
      offPeakActionOnLogoff: actionOnLogoff
      offPeakMinutesToWaitOnLogoff: minutesToWaitOnLogoff
    }
  }
  {
    name: take(
      length(weekendCandidateName) <= 64
        ? weekendCandidateName
        : '${take(weekendCandidateName, 55)}-${uniqueString(scalingPlanName, 'Weekend')}',
      64
    )
    properties: {
      daysOfWeek: [
        'Saturday'
        'Sunday'
      ]
      rampUpStartTime: {
        hour: 9
        minute: 0
      }
      rampUpAutoStartHosts: 'None'
      rampUpStartVMOnConnect: startVmOnConnectSetting
      rampUpActionOnDisconnect: actionOnDisconnect
      rampUpMinutesToWaitOnDisconnect: minutesToWaitOnDisconnect
      rampUpActionOnLogoff: actionOnLogoff
      rampUpMinutesToWaitOnLogoff: minutesToWaitOnLogoff
      peakStartTime: {
        hour: 10
        minute: 0
      }
      peakStartVMOnConnect: startVmOnConnectSetting
      peakActionOnDisconnect: actionOnDisconnect
      peakMinutesToWaitOnDisconnect: minutesToWaitOnDisconnect
      peakActionOnLogoff: actionOnLogoff
      peakMinutesToWaitOnLogoff: minutesToWaitOnLogoff
      rampDownStartTime: {
        hour: 16
        minute: 0
      }
      rampDownStartVMOnConnect: startVmOnConnectSetting
      rampDownActionOnDisconnect: actionOnDisconnect
      rampDownMinutesToWaitOnDisconnect: minutesToWaitOnDisconnect
      rampDownActionOnLogoff: actionOnLogoff
      rampDownMinutesToWaitOnLogoff: minutesToWaitOnLogoff
      offPeakStartTime: {
        hour: 18
        minute: 0
      }
      offPeakStartVMOnConnect: startVmOnConnectSetting
      offPeakActionOnDisconnect: actionOnDisconnect
      offPeakMinutesToWaitOnDisconnect: minutesToWaitOnDisconnect
      offPeakActionOnLogoff: actionOnLogoff
      offPeakMinutesToWaitOnLogoff: minutesToWaitOnLogoff
    }
  }
]

var resolvedSchedules = !empty(customSchedules) ? customSchedules : defaultPersonalSchedules

// Built-in Role Definition ID: Desktop Virtualization Power On Off Contributor
var powerOnOffContributorRoleId = subscriptionResourceId(
  'Microsoft.Authorization/roleDefinitions',
  '40c5ff49-9181-41f8-ae61-143b0e78555e'
)

// === 1. RBAC FOR START VM ON CONNECT / AUTOSCALE ===
var subPowerRbacDeployCandidate = 'deploy-sub-powerrbac-${scalingPlanName}'
module avdPowerSubscriptionRoleAssignment 'powerOnOffRbac-subscription.bicep' = if (!empty(avdServicePrincipalObjectId) && powerOnOffRbacScope == 'Subscription') {
  name: take(
    length(subPowerRbacDeployCandidate) <= 64
      ? subPowerRbacDeployCandidate
      : '${take(subPowerRbacDeployCandidate, 55)}-${uniqueString(scalingPlanName)}',
    64
  )
  scope: subscription()
  params: {
    avdServicePrincipalObjectId: avdServicePrincipalObjectId
  }
}

resource avdPowerRoleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (!empty(avdServicePrincipalObjectId) && powerOnOffRbacScope == 'ResourceGroup') {
  name: guid(resourceGroup().id, avdServicePrincipalObjectId, powerOnOffContributorRoleId)
  scope: resourceGroup()
  properties: {
    roleDefinitionId: powerOnOffContributorRoleId
    principalId: avdServicePrincipalObjectId
    principalType: 'ServicePrincipal'
  }
}

// === 2. PERSONAL SCALING PLAN ===
resource scalingPlan 'Microsoft.DesktopVirtualization/scalingPlans@2024-04-03' = {
  name: scalingPlanName
  location: location
  tags: tags
  dependsOn: [
    avdPowerSubscriptionRoleAssignment
    avdPowerRoleAssignment
  ]
  properties: {
    friendlyName: !empty(scalingPlanFriendlyName) ? scalingPlanFriendlyName : scalingPlanName
    description: scalingPlanDescription
    hostPoolType: 'Personal'
    timeZone: timeZone
    exclusionTag: resolvedExclusionTag
    // Inline schedules on the parent scalingPlans resource apply only to Pooled
    // plans; Personal schedules are deployed via scalingPlans/personalSchedules.
    schedules: []
    hostPoolReferences: !empty(hostPoolResourceId)
      ? [
          {
            hostPoolArmPath: hostPoolResourceId
            scalingPlanEnabled: scalingPlanEnabled
          }
        ]
      : []
  }
}

// === 2b. PERSONAL SCALING PLAN SCHEDULES (Child Resources) ===
resource personalSchedules 'Microsoft.DesktopVirtualization/scalingPlans/personalSchedules@2024-04-03' = [
  for (schedule, i) in resolvedSchedules: {
    parent: scalingPlan
    name: !empty(schedule.?name) ? substring(string(schedule.name), 0, min(length(string(schedule.name)), 64)) : 'Schedule-${i + 1}'
    properties: schedule.?properties ?? schedule
  }
]

// === 3. DIAGNOSTIC SETTINGS ===
var spDiagCandidate = 'diag-${scalingPlanName}'
resource scalingPlanDiagnostics 'Microsoft.Insights/diagnosticSettings@2021-05-01-preview' = if (!empty(logAnalyticsWorkspaceId)) {
  name: take(
    length(spDiagCandidate) <= 64 ? spDiagCandidate : '${take(spDiagCandidate, 55)}-${uniqueString(scalingPlanName)}',
    64
  )
  scope: scalingPlan
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
output scalingPlanId string = scalingPlan.id
output scalingPlanName string = scalingPlan.name
output exclusionTag string = resolvedExclusionTag
