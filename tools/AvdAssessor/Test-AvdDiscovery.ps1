#Requires -Version 5.1
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$CollectorPath = Join-Path $PSScriptRoot 'Invoke-AvdDiscovery.ps1'
$ParseTokens = $null
$ParseErrors = $null
$CollectorAst = [System.Management.Automation.Language.Parser]::ParseFile($CollectorPath, [ref]$ParseTokens, [ref]$ParseErrors)
if ($ParseErrors.Count -gt 0) { throw ($ParseErrors | Out-String) }

function Assert-Equal {
    param($Actual, $Expected, [string]$Label)
    if ($Actual -ne $Expected) { throw "$Label : expected '$Expected', got '$Actual'" }
}

$ResultFunction = $CollectorAst.Find({
    param($Node)
    $Node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -eq 'New-CheckResult'
}, $false)
. ([scriptblock]::Create($ResultFunction.Extent.Text))
foreach ($FunctionName in @('Get-AvdArmList', 'Get-AvdReservationScopeMatch')) {
    $HelperFunction = $CollectorAst.Find({
        param($Node)
        $Node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -eq $FunctionName
    }, $false)
    . ([scriptblock]::Create($HelperFunction.Extent.Text))
}

$CapacityLoop = $CollectorAst.Find({
    param($Node)
    $Node -is [System.Management.Automation.Language.ForEachStatementAst] -and
    $Node.Condition.Extent.Text -eq '$VNet.Subnets' -and $Node.Extent.Text.Contains('NET-SUBCAP-')
}, $true)
$PeeringLoop = $CollectorAst.Find({
    param($Node)
    $Node -is [System.Management.Automation.Language.ForEachStatementAst] -and
    $Node.Condition.Extent.Text -eq '$VNet.VirtualNetworkPeerings'
}, $true)
$NetworkChecks = [scriptblock]::Create($CapacityLoop.Extent.Text + "`n" + $PeeringLoop.Extent.Text)

$Cases = @(
    @{ Name = 'scalar'; Prefix = '10.0.0.0/24'; Status = 'Pass'; Total = 251 },
    @{ Name = 'array'; Prefix = @('10.0.0.0/24'); Status = 'Pass'; Total = 251 },
    @{ Name = 'list'; Prefix = [System.Collections.Generic.List[string]]@('10.0.0.0/24'); Status = 'Pass'; Total = 251 },
    @{ Name = 'large'; Prefix = '64.0.0.0/2'; Status = 'Pass'; Total = 1073741819 },
    @{ Name = 'multiple'; Prefix = @('10.0.0.0/24', '10.0.1.0/24'); Status = 'Pass'; Total = 502 },
    @{ Name = 'zero-prefix'; Prefix = '0.0.0.0/0'; Status = 'Error' },
    @{ Name = 'invalid'; Prefix = 'invalid/24'; Status = 'Error' },
    @{ Name = 'empty'; Prefix = @(); Status = 'Error' },
    @{ Name = 'ipv6'; Prefix = 'fd00::/64'; Status = 'Error' },
    @{ Name = 'dual-stack'; Prefix = @('10.0.0.0/24', 'fd00::/64'); Status = 'Error' },
    @{ Name = 'too-small'; Prefix = '10.0.0.0/32'; Status = 'Error' }
)
foreach ($Case in $Cases) {
    $AllChecks = [System.Collections.ArrayList]::new()
    $Matches = @{ 1 = '0' }
    $VNetName = 'test-vnet'
    $VNet = [pscustomobject]@{
        Subnets = @([pscustomobject]@{ Name = $Case.Name; AddressPrefix = $Case.Prefix; IpConfigurations = @('nic-1') })
        VirtualNetworkPeerings = @([pscustomobject]@{
            Name = 'hub'; PeeringState = 'Connected'; RemoteVirtualNetwork = @{ Id = '/virtualNetworks/hub' }
            AllowForwardedTraffic = $true; AllowGatewayTransit = $true
        })
    }
    . $NetworkChecks
    Assert-Equal $AllChecks.Count 2 "$($Case.Name): capacity and peering emitted"
    Assert-Equal $AllChecks[0].Status $Case.Status "$($Case.Name): capacity status"
    Assert-Equal $AllChecks[1].Status 'Pass' "$($Case.Name): peering still evaluated"
    if ($Case.ContainsKey('Total')) { Assert-Equal $AllChecks[0].Evidence.Total $Case.Total "$($Case.Name): capacity" }
}
Write-Host "PASS: $($Cases.Count) subnet capacity cases, peering continuation, and collector AST parsing."

$SharedLoop = $CollectorAst.Find({
    param($Node)
    $Node -is [System.Management.Automation.Language.ForEachStatementAst] -and
    $Node.Variable.VariablePath.UserPath -eq 'SharedSubscription'
}, $true)
$HubLoop = $CollectorAst.Find({
    param($Node)
    $Node -is [System.Management.Automation.Language.ForEachStatementAst] -and
    $Node.Variable.VariablePath.UserPath -eq 'AvdSubscription'
}, $true)
Assert-Equal ($SharedLoop.Parent -eq $CollectorAst.EndBlock) $true 'shared discovery is outside the host-pool gate'
$SharedChecks = [scriptblock]::Create($SharedLoop.Extent.Text + "`n" + $HubLoop.Extent.Text)
$HubVNetId = '/subscriptions/hub/resourceGroups/network/providers/Microsoft.Network/virtualNetworks/hub'
function Write-Status { param($Message, $Level) }
function Set-AzContext { param($SubscriptionId, $ErrorAction, $WarningAction) $script:TestSubscription = $SubscriptionId }
function Get-AzFirewall {
    param($ErrorAction)
    if ($script:TestSubscription -eq 'hub') {
        if ($script:DenyFirewall) { throw '403 firewall access denied' }
        [pscustomobject]@{
            Id = '/subscriptions/hub/resourceGroups/network/providers/Microsoft.Network/azureFirewalls/firewall'
            Name = 'firewall'; Sku = @{ Tier = 'Standard' }; IpConfigurations = @(@{ Subnet = @{ Id = "$HubVNetId/subnets/AzureFirewallSubnet" } })
        }
    }
}
function Get-AzVirtualNetworkGateway {
    param($ResourceGroupName, $Name, $ErrorAction)
    [pscustomobject]@{
        Id = '/subscriptions/hub/resourceGroups/network/providers/Microsoft.Network/virtualNetworkGateways/gateway'
        Name = 'gateway'; GatewayType = 'ExpressRoute'; Sku = @{ Name = 'Standard' }
        IpConfigurations = @(@{ Subnet = @{ Id = "$HubVNetId/subnets/GatewaySubnet" } })
    }
}
function Get-AzResource {
    param($ResourceType, $ResourceId, [switch]$ExpandProperties, $ErrorAction)
    if ($ResourceId) {
        return [pscustomobject]@{ Properties = @{ provisioningState = 'Succeeded'; serviceProviderProperties = @{ serviceProviderName = 'test'; bandwidthInMbps = 1000 } } }
    }
    if ($script:TestSubscription -eq 'hub') {
        [pscustomobject]@{ ResourceId = "/subscriptions/hub/resourceGroups/network/providers/$ResourceType/resource"; Name = 'resource'; ResourceGroupName = 'network' }
    }
}
function Get-AzStorageAccount {
    param($ErrorAction)
    if ($script:TestSubscription -eq 'hub') {
        [pscustomobject]@{
            Id = '/subscriptions/hub/resourceGroups/storage/providers/Microsoft.Storage/storageAccounts/shared'
            StorageAccountName = 'shared'; ResourceGroupName = 'storage'; Sku = @{ Name = 'Standard_LRS' }
        }
    }
}
function Get-AzRmStorageShare {
    param($ResourceGroupName, $StorageAccountName, $ErrorAction)
    if ($script:ProfileShare) { [pscustomobject]@{ Name = 'profiles' } }
}
function Get-AzStorageFileServiceProperty {
    param($ResourceGroupName, $StorageAccountName, $ErrorAction, $WarningAction)
    [pscustomobject]@{ ShareDeleteRetentionPolicy = @{ Enabled = $true }; ProtocolSetting = $null }
}
function Invoke-AzRestMethod {
    param($Path, $Method, $ErrorAction)
    $Values = @()
    if ($script:TestSubscription -eq 'hub') {
        $Values = @(@{
            id = '/subscriptions/hub/resourceGroups/reserved/providers/Microsoft.Compute/capacityReservationGroups/group'
            name = 'group'; properties = @{ virtualMachinesAssociated = @(@{ id = '/subscriptions/avd/resourceGroups/hosts/providers/Microsoft.Compute/virtualMachines/host' }) }
        })
    }
    [pscustomobject]@{ StatusCode = 200; Content = ConvertTo-Json -InputObject @{ value = $Values } -Depth 8 }
}

function New-TestDiscovery {
    param([string[]]$SubscriptionIds)
    [pscustomobject]@{
        Subscriptions = @($SubscriptionIds | ForEach-Object { [pscustomobject]@{ Id = $_; Name = $_; TenantId = 'tenant' } })
        CollectionStatus = @(); Errors = @()
        Inventory = [pscustomobject]@{
            HostPools = @([pscustomobject]@{ SubscriptionId = 'avd' })
            SessionHosts = @([pscustomobject]@{ ResourceId = '/subscriptions/avd/resourceGroups/hosts/providers/Microsoft.Compute/virtualMachines/host' })
            VNets = @([pscustomobject]@{
                Id = '/subscriptions/avd/resourceGroups/avd/providers/Microsoft.Network/virtualNetworks/spoke'
                SubscriptionId = 'avd'; Peerings = @(@{ RemoteVNet = $HubVNetId })
            })
            Firewalls = @(); VPNGateways = @(); ExpressRouteCircuits = @(); StorageAccounts = @(); CapacityReservations = @(); Reservations = @()
        }
    }
}
foreach ($Order in @('avd,hub', 'hub,avd')) {
    $Discovery = New-TestDiscovery ($Order -split ',')
    $AllChecks = [System.Collections.ArrayList]::new()
    $script:DenyFirewall = $false
    . $SharedChecks
    Assert-Equal $Discovery.Errors.Count 0 "$Order : shared discovery errors"
    foreach ($Kind in @('Firewalls', 'VPNGateways', 'ExpressRouteCircuits', 'StorageAccounts', 'CapacityReservations')) {
        Assert-Equal $Discovery.Inventory.$Kind.Count 1 "$Order : $Kind in subscription without host pools"
        Assert-Equal $Discovery.Inventory.$Kind[0].SubscriptionId 'hub' "$Order : resource subscription"
    }
    Assert-Equal @($AllChecks | Where-Object { $_.Id -like 'NET-HUB*' -and $_.Status -eq 'Pass' }).Count 2 "$Order : related hub checks"
    Assert-Equal @($AllChecks | Where-Object { $_.Id -like 'GOV-CAPRESERV*' })[0].Evidence.AssociatedGroupIds.Count 1 "$Order : cross-subscription capacity association"
}
$Discovery = New-TestDiscovery @('avd', 'hub')
$AllChecks = [System.Collections.ArrayList]::new()
$script:DenyFirewall = $true
. $SharedChecks
Assert-Equal @($AllChecks | Where-Object { $_.Id -like 'NET-HUBFW*' })[0].Status 'Error' 'firewall denial is missing evidence'
Assert-Equal $Discovery.Inventory.ExpressRouteCircuits.Count 1 'firewall failure does not block circuits'
Assert-Equal $Discovery.Inventory.StorageAccounts.Count 1 'firewall failure does not block storage'
$script:DenyFirewall = $false
$Discovery = New-TestDiscovery @('avd')
$AllChecks = [System.Collections.ArrayList]::new()
. $SharedChecks
Assert-Equal @($AllChecks | Where-Object { $_.Id -like 'NET-HUB*' -and $_.Status -eq 'Error' }).Count 2 'unselected hub is missing evidence'
Write-Host 'PASS: shared resources without host pools, both subscription orders, access denial, and unselected hub.'

$script:ProfileShare = $true
$Discovery = New-TestDiscovery @('avd', 'hub')
$AllChecks = [System.Collections.ArrayList]::new()
. $SharedChecks
Assert-Equal $Discovery.Errors.Count 0 'FSLogix candidate discovery errors'
Assert-Equal $Discovery.Inventory.StorageAccounts[0].LikelyFSLogix $true 'profile share identifies FSLogix candidate in hub'
Assert-Equal ($Discovery.Inventory.StorageAccounts[0].FSLogixEvidence -contains 'share name(s): profiles') $true 'FSLogix classification evidence retained'
Assert-Equal @($AllChecks | Where-Object { $_.Id -eq 'PROF-SD-shared' })[0].Status 'Pass' 'profile storage checks execute in hub subscription'
$script:ProfileShare = $false
Write-Host 'PASS: FSLogix share-name classification and file-service checks in a subscription without host pools.'

$ReservationLoop = $CollectorAst.Find({
    param($Node)
    $Node -is [System.Management.Automation.Language.ForEachStatementAst] -and
    $Node.Variable.VariablePath.UserPath -eq 'ReservationTenant'
}, $true)
Assert-Equal ($ReservationLoop.Parent -eq $CollectorAst.EndBlock) $true 'reservation orders are queried outside subscription sweeps'
$ReservationChecks = [scriptblock]::Create($ReservationLoop.Extent.Text)
$OrdersPath = '/providers/Microsoft.Capacity/reservationOrders?api-version=2022-11-01'
$FirstOrder = '/providers/Microsoft.Capacity/reservationOrders/first'
$SecondOrder = '/providers/Microsoft.Capacity/reservationOrders/second'
$FirstChildren = "$FirstOrder/reservations?api-version=2022-11-01"
$SecondChildren = "$SecondOrder/reservations?api-version=2022-11-01"
function New-TestReservation {
    param([string]$Name, [string]$ScopeType, [string]$Scope, [string]$State = 'Succeeded', [string]$ResourceType = 'VirtualMachines')
    @{
        id = "$FirstOrder/reservations/$Name"; name = $Name; location = 'westeurope'; sku = @{ name = 'Standard_D4s_v5' }
        properties = @{
            reservedResourceType = $ResourceType; appliedScopeType = $ScopeType; appliedScopes = @($Scope)
            provisioningState = $State; quantity = 2; term = 'P1Y'
        }
    }
}
function New-TestPage {
    param([object[]]$Values, [string]$NextLink)
    [pscustomobject]@{ StatusCode = 200; Content = ConvertTo-Json -InputObject @{ value = @($Values); nextLink = $NextLink } -Depth 12 }
}
$SingleReservation = New-TestReservation 'single' 'Single' '/subscriptions/avd'
$GroupReservation = New-TestReservation 'group' 'Single' ''
$GroupReservation.properties.appliedScopeProperties = @{ resourceGroupId = '/subscriptions/avd/resourceGroups/hosts'; subscriptionId = '/subscriptions/avd' }
$BaseResponses = @{}
$BaseResponses[$OrdersPath] = New-TestPage @(@{ id = $FirstOrder }) "https://management.azure.com$OrdersPath&page=2"
$BaseResponses["$OrdersPath&page=2"] = New-TestPage @(@{ id = $FirstOrder }, @{ id = $SecondOrder })
$BaseResponses[$FirstChildren] = New-TestPage @($SingleReservation, (New-TestReservation 'shared' 'Shared' '')) "https://management.azure.com$FirstChildren&page=2"
$BaseResponses["$FirstChildren&page=2"] = New-TestPage @(
    $GroupReservation,
    (New-TestReservation 'other-group' 'Single' '/subscriptions/avd/resourceGroups/hosts-extra'),
    (New-TestReservation 'other-subscription' 'Single' '/subscriptions/other'),
    (New-TestReservation 'management-group' 'ManagementGroup' ''),
    (New-TestReservation 'expired' 'Single' '/subscriptions/avd' 'Expired'),
    (New-TestReservation 'sql' 'Single' '/subscriptions/avd' 'Succeeded' 'SqlDatabases')
)
$BaseResponses[$SecondChildren] = New-TestPage @($SingleReservation)
function Invoke-AzRestMethod {
    param($Path, $Method, $ErrorAction)
    if (-not $script:Responses.ContainsKey($Path)) { throw "Unexpected ARM request: $Path" }
    $script:Requests += $Path
    $script:Responses[$Path]
}
function Reset-ReservationTest {
    $script:Responses = $BaseResponses.Clone()
    $script:Requests = @()
}
Reset-ReservationTest
$Discovery = New-TestDiscovery @('avd', 'hub')
$AllChecks = [System.Collections.ArrayList]::new()
. $ReservationChecks
Assert-Equal $Discovery.Errors.Count 0 'reservation discovery errors'
Assert-Equal $Discovery.Inventory.Reservations.Count 7 'child VM reservations retained, duplicates and non-VM excluded'
Assert-Equal @($script:Requests | Where-Object { $_ -eq $OrdersPath }).Count 1 'orders queried once per tenant'
Assert-Equal @($script:Requests | Where-Object { $_ -eq $FirstChildren }).Count 1 'duplicate orders queried once'
Assert-Equal $script:Requests.Count 5 'orders and children paginated'
Assert-Equal @($Discovery.Inventory.Reservations | Where-Object { $_.ScopeMatch -eq 'MatchesSelectedScope' }).Count 3 'subscription and exact resource-group matches'
Assert-Equal @($Discovery.Inventory.Reservations | Where-Object { $_.ScopeMatch -eq 'SharedScopeUnverified' }).Count 1 'shared scope is not assumed to cover AVD'
Assert-Equal @($Discovery.Inventory.Reservations | Where-Object { $_.ScopeMatch -eq 'ManagementGroupScopeUnverified' }).Count 1 'management-group scope stays unverified'
Assert-Equal @($Discovery.Inventory.Reservations | Where-Object { $_.ProvisioningState -eq 'Expired' }).Count 1 'lifecycle state retained'
Assert-Equal $AllChecks[0].Status 'Error' 'inventory alone does not prove discount coverage'

Reset-ReservationTest
$script:Responses[$FirstChildren] = [pscustomobject]@{ StatusCode = 403; Content = '{}' }
$Discovery = New-TestDiscovery @('avd', 'hub')
$AllChecks = [System.Collections.ArrayList]::new()
. $ReservationChecks
Assert-Equal $Discovery.Inventory.Reservations.Count 1 'one denied order does not discard other accessible orders'
Assert-Equal $AllChecks[0].Evidence.CollectionStatus 'Error' 'partial reservation access is recorded'

Reset-ReservationTest
$script:Responses[$OrdersPath] = [pscustomobject]@{ StatusCode = 403; Content = '{}' }
$Discovery = New-TestDiscovery @('avd', 'hub')
$AllChecks = [System.Collections.ArrayList]::new()
. $ReservationChecks
Assert-Equal $Discovery.Inventory.Reservations.Count 0 'denied reservation list'
Assert-Equal $AllChecks[0].Evidence.CollectionStatus 'Error' 'denial differs from empty inventory'

Reset-ReservationTest
$script:Responses[$OrdersPath] = New-TestPage @()
$Discovery = New-TestDiscovery @('avd', 'hub')
$AllChecks = [System.Collections.ArrayList]::new()
. $ReservationChecks
Assert-Equal $AllChecks[0].Evidence.CollectionStatus 'Complete' 'empty visible list is distinguished from denial'
Assert-Equal $AllChecks[0].Status 'Error' 'empty visible inventory does not prove workload coverage'

Reset-ReservationTest
$script:Responses[$OrdersPath] = New-TestPage @() "https://management.azure.com$OrdersPath"
$RejectedCycle = $false
try { $null = @(Get-AvdArmList $OrdersPath) } catch { $RejectedCycle = $true }
Assert-Equal $RejectedCycle $true 'pagination cycles terminate with an error'
Write-Host 'PASS: reservation child enumeration, pagination, deduplication, scope matching, lifecycle state, and partial/denied/empty responses.'

$DrCheck = $CollectorAst.Find({
    param($Node)
    $Node -is [System.Management.Automation.Language.IfStatementAst] -and $Node.Extent.Text.Contains('BCDR-DRCAP')
}, $true)
$AllChecks = [System.Collections.ArrayList]::new()
$HPLocations = @('westeurope', 'northeurope')
$AllCapResRegions = @('unrelated-region')
. ([scriptblock]::Create($DrCheck.Extent.Text))
Assert-Equal $AllChecks[0].Status 'Error' 'unrelated capacity groups cannot prove AVD disaster recovery capacity'
Write-Host 'PASS: DR capacity remains unassessed without workload capacity evidence.'