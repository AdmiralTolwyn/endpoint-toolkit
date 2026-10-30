/*
  SYNOPSIS: Parameter file for AVD Personal Host Pool + Session Hosts in West Germany
  FILE:     bicep/parameters/personal-hostpool-dewc.bicepparam
  USAGE:
    az deployment group create \
      --resource-group rg-avd-personal-dewc-001 \
      --template-file avd/bicep/main-personal-hostpool.bicep \
      --parameters avd/bicep/parameters/personal-hostpool-dewc.bicepparam \
      --parameters localAdminPassword='<secret>'
*/

using '../main-personal-hostpool.bicep'

// --- 1. Region & Time Zone (West Germany / Frankfurt) ---
param location = 'germanywestcentral'
param managementPlaneLocation = 'germanywestcentral'
param timeZone = 'W. Europe Standard Time'

// --- 2. Management Plane (Personal Host Pool, Desktop App Group, Workspace) ---
param hostPoolName = 'vdpool-fsavd-prod-dewc-001'
param hostPoolFriendlyName = 'FS Personal Desktops (DEWC)'
param hostPoolDescription = 'Personal Windows 11 Desktops in Germany West Central (Entra ID Joined)'
param personalDesktopAssignmentType = 'Automatic'
param startVMOnConnect = true
param hostPoolPublicNetworkAccess = 'Enabled'
param enforceEntraIdRdpProperties = true
param tokenExpirationHours = 24

param deployApplicationGroup = true
param applicationGroupName = 'vdag-desktop-fsavd-prod-dewc-001'
param applicationGroupFriendlyName = 'Personal Desktop (DEWC)'
param desktopUserPrincipalIds = [] // Populate with Entra ID Security Group or User Object IDs for user assignment
param desktopUserPrincipalType = 'Group' // Set to 'User' when assigning individual User Object IDs

param deployWorkspace = true
param workspaceName = 'vdws-fsavd-prod-dewc-001'
param workspaceFriendlyName = 'FS AVD Workspace (DEWC)'

// --- 3. Personal Scaling Plan ---
param deployScalingPlan = true
param scalingPlanName = 'vdscaling-fsavd-prod-dewc-001'
param scalingPlanFriendlyName = 'FS Personal Desktops Scaling (DEWC)'
param weekdayRampUpAutoStartHosts = 'WithAssignedUser'
param actionOnDisconnect = 'Deallocate'
param minutesToWaitOnDisconnect = 30
param actionOnLogoff = 'Deallocate'
param minutesToWaitOnLogoff = 15
// Object ID of the 'Azure Virtual Desktop' Enterprise App (appId 9cdead84-a844-4324-93f2-b2e6bb768d07)
// Retrieve via: az ad sp show --id 9cdead84-a844-4324-93f2-b2e6bb768d07 --query id -o tsv
// When provided, assigns 'Desktop Virtualization Power On Off Contributor' at Subscription scope by default
// (required by Microsoft Autoscale; requires Subscription Owner / User Access Administrator permissions),
// or on Management + Compute RGs when powerOnOffRbacScope = 'ResourceGroup' (for standalone Start VM on Connect).
// Leave avdServicePrincipalObjectId empty if the role is already pre-assigned in the subscription.
param avdServicePrincipalObjectId = ''
param powerOnOffRbacScope = 'Subscription'

// --- 4. Session Host Naming (Up to 12-char prefix + 3-digit sequence -> e.g. FSAVDP11001, FSAVDP11002) ---
param deploySessionHosts = true
param vmPrefix = 'FSAVDP11'
param vmCount = 2
param vmStartIndex = 1
param vmIndexPadLength = 3

// --- 5. Compute, Networking & Image Configuration ---
param subnetId = '/subscriptions/00000000-0000-0000-0000-000000000000/resourceGroups/rg-network-dewc-001/providers/Microsoft.Network/virtualNetworks/vnet-avd-dewc-001/subnets/snet-avd-001'
param localAdminUser = 'avdadmin'
param vmSize = 'Standard_D4s_v5'
param osDiskType = 'Premium_LRS'
param osDiskSizeGB = 128
param availabilityZones = [
  '1'
  '2'
  '3'
]
param acceleratedNetworking = true
param enableIntune = true

// Leave imageResourceId empty to use Azure Marketplace Windows 11 single-session,
// or set imageResourceId to an Azure Compute Gallery image version ID.
param imageResourceId = ''
param imagePublisher = 'MicrosoftWindowsDesktop'
param imageOffer = 'windows-11'
param imageSku = 'win11-24h2-ent'
param marketplaceImageVersion = 'latest'

// --- 6. LZA Security Hardening & Monitoring ---
param securityType = 'TrustedLaunch'
param secureBootEnabled = true
param vTpmEnabled = true
param enableGuestAttestation = true
param encryptionAtHost = false
param deployAntiMalware = false
param deployMonitoringAgent = false
param dataCollectionRuleId = ''
param logAnalyticsWorkspaceId = ''

param customTags = {
  Customer: 'FS'
  PoolType: 'Personal'
}
