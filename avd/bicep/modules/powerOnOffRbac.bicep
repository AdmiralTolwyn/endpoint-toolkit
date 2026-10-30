/*
  SYNOPSIS: AVD Power On/Off RBAC Role Assignment Module
  FILE:     bicep/modules/powerOnOffRbac.bicep
  DESCRIPTION:
    Assigns the built-in "Desktop Virtualization Power On Off Contributor" role
    (40c5ff49-9181-41f8-ae61-143b0e78555e) to the Azure Virtual Desktop service
    principal on the target Resource Group. Used for Start VM on Connect and
    Personal/Pooled Autoscale across Management and Compute Resource Groups.
*/

targetScope = 'resourceGroup'

@description('Object ID of the "Azure Virtual Desktop" enterprise application (appId 9cdead84-a844-4324-93f2-b2e6bb768d07) in the tenant')
param avdServicePrincipalObjectId string

var powerOnOffContributorRoleId = subscriptionResourceId(
  'Microsoft.Authorization/roleDefinitions',
  '40c5ff49-9181-41f8-ae61-143b0e78555e'
)

resource avdPowerRoleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (!empty(avdServicePrincipalObjectId)) {
  name: guid(resourceGroup().id, avdServicePrincipalObjectId, powerOnOffContributorRoleId)
  scope: resourceGroup()
  properties: {
    roleDefinitionId: powerOnOffContributorRoleId
    principalId: avdServicePrincipalObjectId
    principalType: 'ServicePrincipal'
  }
}

output roleAssignmentId string = !empty(avdServicePrincipalObjectId) ? avdPowerRoleAssignment.id : ''
