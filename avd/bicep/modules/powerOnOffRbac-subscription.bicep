/*
  SYNOPSIS: AVD Power On/Off RBAC Role Assignment Module (Subscription Scope)
  FILE:     bicep/modules/powerOnOffRbac-subscription.bicep
  DESCRIPTION:
    Assigns the built-in "Desktop Virtualization Power On Off Contributor" role
    (40c5ff49-9181-41f8-ae61-143b0e78555e) to the Azure Virtual Desktop service
    principal at Subscription scope, as required by Microsoft Autoscale for
    AVD Scaling Plans.
*/

targetScope = 'subscription'

@description('Object ID of the "Azure Virtual Desktop" enterprise application (appId 9cdead84-a844-4324-93f2-b2e6bb768d07) in the tenant')
param avdServicePrincipalObjectId string

var powerOnOffContributorRoleId = subscriptionResourceId(
  'Microsoft.Authorization/roleDefinitions',
  '40c5ff49-9181-41f8-ae61-143b0e78555e'
)

resource avdPowerRoleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (!empty(avdServicePrincipalObjectId)) {
  name: guid(subscription().id, avdServicePrincipalObjectId, powerOnOffContributorRoleId)
  properties: {
    roleDefinitionId: powerOnOffContributorRoleId
    principalId: avdServicePrincipalObjectId
    principalType: 'ServicePrincipal'
  }
}

output roleAssignmentId string = !empty(avdServicePrincipalObjectId) ? avdPowerRoleAssignment.id : ''
