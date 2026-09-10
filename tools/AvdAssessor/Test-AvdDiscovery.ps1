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
foreach ($FunctionName in @('Get-AvdArmList', 'Get-AvdReservationScopeMatch', 'Get-AvdStorageZrsAvailability', 'Get-AvdStorageReplicationAssessment')) {
    $HelperFunction = $CollectorAst.Find({
        param($Node)
        $Node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -eq $FunctionName
    }, $false)
    . ([scriptblock]::Create($HelperFunction.Extent.Text))
}
$StorageSkuCache = @{}

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

$PrivateLinkBlocks = @($CollectorAst.FindAll({
    param($Node)
    $Node -is [System.Management.Automation.Language.TryStatementAst] -and
    ($Node.Body.Extent.Text.Contains('-Id "SEC-HPPL-') -or $Node.Body.Extent.Text.Contains('-Id "NET-PL-')) -and
    -not $Node.Body.Extent.Text.Contains('foreach ($HP in $HostPools)')
}, $true))
Assert-Equal $PrivateLinkBlocks.Count 2 'both Private Link checks are exercised'
$PrivateLinkChecks = [scriptblock]::Create(($PrivateLinkBlocks | ForEach-Object { $_.Extent.Text }) -join "`n")
function Get-AzResource {
    param($ResourceId, [switch]$ExpandProperties, $ErrorAction)
    Assert-Equal $ExpandProperties.IsPresent $true 'Private Link resource properties expanded'
    Assert-Equal $ErrorAction 'Stop' 'Private Link errors are not suppressed'
    switch ($script:PrivateLinkScenario) {
        'denied' { throw '403 access denied' }
        'null' { return $null }
        'missing-properties' { return [pscustomobject]@{ ResourceId = $ResourceId } }
        default {
            $Connections = @()
            if ($script:PrivateLinkScenario -in @('approved', 'pending')) {
                $Connections = @(@{ properties = @{ privateLinkServiceConnectionState = @{ status = $script:PrivateLinkScenario } } })
            }
            return [pscustomobject]@{ Properties = @{ privateEndpointConnections = $Connections; publicNetworkAccess = 'Enabled' } }
        }
    }
}
$HP = [pscustomobject]@{ Name = 'pool'; Id = '/subscriptions/avd/resourceGroups/hosts/providers/Microsoft.DesktopVirtualization/hostPools/pool' }
foreach ($Scenario in @('zero', 'approved', 'pending', 'denied', 'null', 'missing-properties')) {
    $script:PrivateLinkScenario = $Scenario
    $AllChecks = [System.Collections.ArrayList]::new()
    . $PrivateLinkChecks
    Assert-Equal $AllChecks.Count 2 "$Scenario : both Private Link outcomes"
    foreach ($Check in $AllChecks) {
        Assert-Equal $Check.Status 'Error' "$Scenario : neither absence nor presence establishes applicability/compliance"
        Assert-Equal $Check.Evidence.Applicability 'Unknown' "$Scenario : applicability unknown"
        if ($Scenario -in @('denied', 'null', 'missing-properties')) {
            Assert-Equal $Check.Evidence.CollectionStatus 'Error' "$Scenario : collection failure recorded"
            Assert-Equal $Check.Evidence.ContainsKey('PECount') $false "$Scenario : no invented zero count"
        } else {
            $ExpectedCount = if ($Scenario -eq 'zero') { 0 } else { 1 }
            Assert-Equal $Check.Evidence.PECount $ExpectedCount "$Scenario : endpoint count retained"
            Assert-Equal $Check.Evidence.CollectionStatus 'Complete' "$Scenario : inventory successfully read"
            Assert-Equal ($Check.Details -like '*Applicability not assessed*') $true "$Scenario : applicability wording"
        }
    }
}
Write-Host 'PASS: Private Link inventory cannot imply applicability/compliance; failed reads cannot imply zero endpoints.'

$AmaFunction = $CollectorAst.Find({
    param($Node)
    $Node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -eq 'Get-AvdAmaEvidence'
}, $false)
. ([scriptblock]::Create($AmaFunction.Extent.Text))
function Invoke-AzRestMethod {
    param($Path, $Method, $ErrorAction)
    $script:AmaRequests += $Path
    if ($script:AmaScenario -in @('denied', 'not-found', 'throttled', 'server-error')) {
        $Code = switch ($script:AmaScenario) { 'denied' { 403 }; 'not-found' { 404 }; 'throttled' { 429 }; default { 500 } }
        return [pscustomobject]@{ StatusCode = $Code; Content = '{}' }
    }
    if ($script:AmaScenario -eq 'malformed') { return [pscustomobject]@{ StatusCode = 200; Content = '{}' } }
    $Values = @()
    if ($script:AmaScenario -ne 'empty') {
        $Publisher = if ($script:AmaScenario -eq 'wrong-publisher') { 'Example.Other' } else { 'Microsoft.Azure.Monitor' }
        $Type = if ($script:AmaScenario -eq 'mma') { 'MicrosoftMonitoringAgent' } else { 'AzureMonitorWindowsAgent' }
        $State = switch ($script:AmaScenario) { 'failed' { 'Failed' }; 'creating' { 'Creating' }; 'missing-state' { $null }; default { 'Succeeded' } }
        $Properties = @{ publisher = $Publisher; type = $Type; provisioningState = $State; typeHandlerVersion = '1.38'; settings = @{ NeverExport = 'sensitive-setting' } }
        if ($script:AmaScenario -eq 'unknown-identity') { $Properties = @{} }
        $Values = @(@{ id = "$Path/custom-extension-name"; name = 'custom-extension-name'; properties = $Properties })
    }
    if ($script:AmaScenario -eq 'paged' -and $Path -notlike '*page=2') {
        return (New-TestPage @(@{ properties = @{ publisher = 'Microsoft.Compute'; type = 'CustomScriptExtension' } }) "$Path&page=2")
    }
    New-TestPage $Values
}
$AmaCases = @(
    @{ Name = 'custom-name'; Status = 'Pass'; Installed = $true; Collection = 'Complete' },
    @{ Name = 'empty'; Status = 'Fail'; Installed = $false; Collection = 'Complete' },
    @{ Name = 'wrong-publisher'; Status = 'Fail'; Installed = $false; Collection = 'Complete' },
    @{ Name = 'mma'; Status = 'Fail'; Installed = $false; Collection = 'Complete' },
    @{ Name = 'failed'; Status = 'Fail'; Installed = $null; Collection = 'Complete' },
    @{ Name = 'creating'; Status = 'Warning'; Installed = $null; Collection = 'Complete' },
    @{ Name = 'missing-state'; Status = 'Error'; Installed = $null; Collection = 'Complete' },
    @{ Name = 'unknown-identity'; Status = 'Error'; Installed = $null; Collection = 'Partial' },
    @{ Name = 'denied'; Status = 'Error'; Installed = $null; Collection = 'Error' },
    @{ Name = 'not-found'; Status = 'Error'; Installed = $null; Collection = 'Error' },
    @{ Name = 'throttled'; Status = 'Error'; Installed = $null; Collection = 'Error' },
    @{ Name = 'server-error'; Status = 'Error'; Installed = $null; Collection = 'Error' },
    @{ Name = 'malformed'; Status = 'Error'; Installed = $null; Collection = 'Error' }
)
$VmId = '/subscriptions/first/resourceGroups/shared/providers/Microsoft.Compute/virtualMachines/host'
foreach ($Case in $AmaCases) {
    $script:AmaScenario = $Case.Name
    $script:AmaRequests = @()
    $Cache = @{}
    $Actual = Get-AvdAmaEvidence -ResourceId $VmId -Cache $Cache
    Assert-Equal $Actual.Status $Case.Status "$($Case.Name): AMA verdict"
    Assert-Equal $Actual.Installed $Case.Installed "$($Case.Name): installed is not fabricated"
    Assert-Equal $Actual.CollectionStatus $Case.Collection "$($Case.Name): collection coverage"
    Assert-Equal $Actual.RuntimeHealth 'NotAssessed' "$($Case.Name): no inferred runtime health"
    Assert-Equal ($script:AmaRequests[0] -eq "$($VmId.ToLowerInvariant())/extensions?api-version=2024-07-01") $true 'query scoped to full VM resource ID'
    Assert-Equal (($Actual | ConvertTo-Json -Depth 10) -like '*sensitive-setting*') $false 'extension settings excluded'
    $null = Get-AvdAmaEvidence -ResourceId $VmId.ToUpperInvariant() -Cache $Cache
    Assert-Equal $script:AmaRequests.Count 1 'case-insensitive full-resource cache'
    $null = Get-AvdAmaEvidence -ResourceId ($VmId -replace '/first/', '/second/') -Cache $Cache
    Assert-Equal $script:AmaRequests.Count 2 'same VM name in another subscription has separate evidence'
}
Write-Host 'PASS: authoritative AMA extension identity/provisioning, absent versus unreadable inventory, full-resource scoping, and evidence redaction.'
$script:AmaScenario = 'paged'
$script:AmaRequests = @()
$PagedAma = Get-AvdAmaEvidence -ResourceId $VmId -Cache @{}
Assert-Equal $PagedAma.Status 'Pass' 'AMA found on second extension page'
Assert-Equal $script:AmaRequests.Count 2 'extension pagination completed'
Assert-Equal (Get-AvdAmaEvidence -ResourceId '' -Cache @{}).Status 'Error' 'missing VM resource ID is not agent absence'

foreach ($FunctionName in @('Get-AvdAmaConfiguration', 'Update-AvdAmaHeartbeats')) {
    $HelperFunction = $CollectorAst.Find({
        param($Node)
        $Node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -eq $FunctionName
    }, $false)
    . ([scriptblock]::Create($HelperFunction.Extent.Text))
}
$RuleId = '/subscriptions/monitoring/resourceGroups/config/providers/Microsoft.Insights/dataCollectionRules/avd-rule'
$WorkspaceId = '/subscriptions/logs/resourceGroups/data/providers/Microsoft.OperationalInsights/workspaces/avd-logs'
function Invoke-AzRestMethod {
    param($Path, $Method, $ErrorAction)
    $script:DcrRequests += $Path
    if ($Path -like '*dataCollectionRuleAssociations*') {
        switch ($script:DcrScenario) {
            'denied' { return [pscustomobject]@{ StatusCode = 403; Content = '{}' } }
            'none' { return (New-TestPage @()) }
            'dce-only' { return (New-TestPage @(@{ properties = @{ dataCollectionEndpointId = '/dce/endpoint' } })) }
            default { return (New-TestPage @(@{ properties = @{ dataCollectionRuleId = $RuleId } }, @{ properties = @{ dataCollectionRuleId = $RuleId } })) }
        }
    }
    if ($script:DcrScenario -eq 'rule-denied') { return [pscustomobject]@{ StatusCode = 403; Content = '{}' } }
    $Destinations = if ($script:DcrScenario -eq 'metrics-only') { @('metrics') } else { @('active') }
    [pscustomobject]@{ StatusCode = 200; Content = ConvertTo-Json -Depth 12 -InputObject @{
        id = $RuleId; properties = @{
            provisioningState = 'Succeeded'
            dataSources = @{ performanceCounters = @(@{ streams = @('Microsoft-Perf'); counterSpecifiers = @('\Processor(_Total)\% Processor Time'); samplingFrequencyInSeconds = 60 }) }
            dataFlows = @(@{ streams = @('Microsoft-Perf'); destinations = $Destinations })
            destinations = @{ logAnalytics = @(@{ name = 'active'; workspaceResourceId = $WorkspaceId }, @{ name = 'unused'; workspaceResourceId = '/unrelated/workspace' }); azureMonitorMetrics = @{ name = 'metrics' } }
        }
    } }
}
foreach ($Scenario in @('associated', 'none', 'dce-only', 'metrics-only', 'denied', 'rule-denied')) {
    $script:DcrScenario = $Scenario
    $script:DcrRequests = @()
    $RuleCache = @{}
    $Config = Get-AvdAmaConfiguration -ResourceId $VmId -RuleCache $RuleCache
    $ExpectedStatus = switch ($Scenario) { 'denied' { 'Error' }; 'rule-denied' { 'Partial' }; default { 'Complete' } }
    Assert-Equal $Config.CollectionStatus $ExpectedStatus "$Scenario : DCR collection status"
    $ExpectedDestinations = if ($Scenario -eq 'associated') { 1 } else { 0 }
    Assert-Equal $Config.LogAnalyticsWorkspaceIds.Count $ExpectedDestinations "$Scenario : only destinations used by data flows"
    if ($Scenario -eq 'associated') {
        $AssociatedConfig = $Config
        Assert-Equal $Config.LogAnalyticsWorkspaceIds[0] $WorkspaceId 'cross-subscription DCR workspace resolved'
        Assert-Equal $Config.Rules.Count 1 'duplicate DCR references deduplicated'
        $null = Get-AvdAmaConfiguration -ResourceId ($VmId -replace '/first/', '/second/') -RuleCache $RuleCache
        Assert-Equal @($script:DcrRequests | Where-Object { $_ -like '*dataCollectionRules/*' }).Count 1 'DCR read cached across VMs'
    }
}
function Get-AzContext { param($ErrorAction) 'original-context' }
function Set-AzContext { param($Context, $ErrorAction, $WarningAction) $script:RestoredContext = $Context }
function Invoke-AvdLaQuery {
    param($WorkspaceResourceId, $Query, $TimespanDays)
    Assert-Equal ($Query.Contains("Category == 'Azure Monitor Agent'")) $true 'only AMA heartbeat category'
    Assert-Equal ($Query.Contains('tolower(_ResourceId) in (avdResources)')) $true 'heartbeat constrained by VM resource IDs'
    Assert-Equal $TimespanDays 1 'explicit heartbeat observation window'
    $script:HeartbeatRequests++
    if ($script:HeartbeatScenario -eq 'denied') { return @{ Ok = $false; Error = '403 Log Analytics'; Rows = @() } }
    $Rows = @(@{ ResourceId = '/subscriptions/unrelated/resourceGroups/shared/providers/Microsoft.Compute/virtualMachines/host'; LastSeen = '2026-09-10T10:00:00Z'; AgeMinutes = 1 })
    if ($script:HeartbeatScenario -in @('recent', 'stale')) {
        $Rows += @{ ResourceId = $VmId.ToUpperInvariant(); LastSeen = '2026-09-10T10:00:00Z'; AgeMinutes = if ($script:HeartbeatScenario -eq 'recent') { 2 } else { 60 } }
    }
    @{ Ok = $true; Error = $null; Rows = $Rows }
}
foreach ($Scenario in @('recent', 'stale', 'missing', 'denied', 'stopped', 'unknown-power')) {
    $script:HeartbeatScenario = $Scenario
    $script:HeartbeatRequests = 0
    $script:RestoredContext = $null
    $Ama = [pscustomobject]@{ ResourceId = $VmId; Details = 'Extension provisioned.'; Status = 'Pass'; RuntimeHealth = 'NotAssessed'; Heartbeats = @(); Configuration = [pscustomobject]@{ CollectionStatus = 'Complete'; Rules = @($RuleId); LogAnalyticsWorkspaceIds = @($WorkspaceId); Errors = @() } }
    $HostEntry = [pscustomobject]@{ ResourceId = $VmId; AMADiscovery = $Ama; PowerStateCode = switch ($Scenario) { 'stopped' { 'PowerState/deallocated' }; 'unknown-power' { $null }; default { 'PowerState/running' } } }
    $AllChecks = [System.Collections.ArrayList]::new()
    [void]$AllChecks.Add([pscustomobject]@{ Id = 'MON-AMA-first-shared-host'; Status = 'Pass'; Details = ''; Evidence = $Ama })
    Update-AvdAmaHeartbeats -SessionHosts @($HostEntry, $HostEntry) -Checks $AllChecks
    $ExpectedRuntime = switch ($Scenario) { 'recent' { 'RecentHeartbeatObserved' }; 'denied' { 'NotAssessed' }; 'stopped' { 'NotAssessedVmStopped' }; 'unknown-power' { 'NotAssessedPowerStateUnknown' }; default { 'NoRecentHeartbeatObserved' } }
    Assert-Equal $Ama.RuntimeHealth $ExpectedRuntime "$Scenario : heartbeat interpretation"
    Assert-Equal $script:RestoredContext 'original-context' 'Log Analytics context restored'
    Assert-Equal $script:HeartbeatRequests 1 'one workspace query, deduplicated hosts'
    Assert-Equal $AllChecks[0].Status 'Pass' 'installation result is distinct from telemetry'
    Assert-Equal ($AllChecks[0].Details -like "*$ExpectedRuntime*") $true 'runtime qualification exposed in check details'
}
Write-Host 'PASS: VM-associated DCRs, referenced destinations, caching, resource-specific AMA heartbeats, stopped/unknown VM states, and query access failures.'

$script:HeartbeatScenario = 'missing'
$script:HeartbeatRequests = 0
$ManyHosts = @(foreach ($Index in 1..101) {
    [pscustomobject]@{
        ResourceId = "$VmId-$Index"; PowerStateCode = 'PowerState/running'
        AMADiscovery = [pscustomobject]@{ ResourceId = "$VmId-$Index"; Details = 'Extension provisioned.'; RuntimeHealth = 'NotAssessed'; Heartbeats = @(); Configuration = [pscustomobject]@{ CollectionStatus = 'Complete'; Rules = @($RuleId); LogAnalyticsWorkspaceIds = @($WorkspaceId); Errors = @() } }
    }
})
Update-AvdAmaHeartbeats -SessionHosts $ManyHosts -Checks ([System.Collections.ArrayList]::new())
Assert-Equal $script:HeartbeatRequests 2 'heartbeat queries are bounded to 100 hosts per batch'
$NoLawHost = $ManyHosts[0]
$NoLawHost.AMADiscovery.Configuration.LogAnalyticsWorkspaceIds = @()
Update-AvdAmaHeartbeats -SessionHosts @($NoLawHost) -Checks ([System.Collections.ArrayList]::new())
Assert-Equal $NoLawHost.AMADiscovery.RuntimeHealth 'NoLogAnalyticsDestination' 'no LAW destination does not claim a dead agent'

$LaFunction = $CollectorAst.Find({
    param($Node)
    $Node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -eq 'Invoke-AvdLaQuery'
}, $false)
. ([scriptblock]::Create($LaFunction.Extent.Text))
function Get-Module { param([switch]$ListAvailable, $Name, $ErrorAction) 'available' }
function Set-AzContext { param($SubscriptionId, $Context, $ErrorAction, $WarningAction) }
function Get-AzOperationalInsightsWorkspace { param($ResourceGroupName, $Name, $ErrorAction) [pscustomobject]@{ CustomerId = 'test-workspace' } }
function Invoke-AzOperationalInsightsQuery {
    param($WorkspaceId, $Query, $Timespan, $ErrorAction)
    switch ($script:LaScenario) {
        'partial' { [pscustomobject]@{ Results = @(@{ ResourceId = $VmId }); Error = 'Partial query failure' } }
        'null' { $null }
        default { [pscustomobject]@{ Results = @(); Error = $null } }
    }
}
foreach ($Scenario in @('partial', 'null', 'empty')) {
    $script:LaScenario = $Scenario
    $Response = Invoke-AvdLaQuery -WorkspaceResourceId $WorkspaceId -Query 'Heartbeat' -TimespanDays 1
    Assert-Equal $Response.Ok ($Scenario -eq 'empty') "$Scenario : incomplete KQL is not successful empty telemetry"
}
Write-Host 'PASS: AMA extension pagination, invalid IDs, bounded heartbeat queries, no LAW destination, and incomplete KQL responses.'

$ExportAma = [pscustomobject]@{
    ResourceId = $VmId
    VM = 'host'
    CollectionStatus = 'Complete'
    Installed = $true
    Extensions = @(@{ Publisher = 'Microsoft.Azure.Monitor'; Type = 'AzureMonitorWindowsAgent'; ProvisioningState = 'Succeeded' })
    Configuration = $AssociatedConfig
    Heartbeats = @(@{ WorkspaceResourceId = $WorkspaceId; State = 'RecentHeartbeat'; AgeMinutes = 2; LastSeen = '2026-09-10T10:00:00Z' })
}
$Export = [pscustomobject]@{
    Inventory = @{ SessionHosts = @(@{ ResourceId = $VmId; AMAInstalled = $true; AMADiscovery = $ExportAma }) }
    CheckResults = @(@{ Id = 'MON-AMA-first-shared-host'; Status = 'Pass'; Evidence = $ExportAma })
}
$RoundTrip = ConvertFrom-Json -InputObject (ConvertTo-Json -InputObject $Export -Depth 10 -WarningAction Stop)
Assert-Equal $RoundTrip.Inventory.SessionHosts[0].AMADiscovery.Configuration.Rules[0].PerformanceCounters[0].CounterSpecifiers[0] '\Processor(_Total)\% Processor Time' 'nested counter evidence survives export'
Assert-Equal ([DateTimeOffset]$RoundTrip.CheckResults[0].Evidence.Heartbeats[0].LastSeen) ([DateTimeOffset]'2026-09-10T10:00:00Z') 'heartbeat timestamp survives export'
Write-Host 'PASS: nested AMA configuration and heartbeat evidence survive discovery JSON serialization.'

function Invoke-AzRestMethod {
    param($Path, $Method, $ErrorAction)
    $script:SkuRequests += $Path
    if ($script:SkuScenario -eq 'denied') { return [pscustomobject]@{ StatusCode = 403; Content = '{}' } }
    if ($script:SkuScenario -eq 'malformed') { return [pscustomobject]@{ StatusCode = 200; Content = '{}' } }
    if ($script:SkuScenario -eq 'empty') { return (New-TestPage @()) }
    $Region = if ($script:SkuScenario -eq 'other-region') { 'elsewhere' } else { 'testregion' }
    $Kind = if ($script:SkuScenario -eq 'wrong-kind') { 'StorageV2' } else { 'FileStorage' }
    $Name = if ($script:SkuScenario -eq 'wrong-tier') { 'Standard_ZRS' } elseif ($script:SkuScenario -eq 'v2') { 'PremiumV2_ZRS' } else { 'Premium_ZRS' }
    $Restrictions = @()
    if ($script:SkuScenario -in @('restricted', 'quota', 'other-restriction', 'unknown-restriction', 'unscoped-restriction')) {
        $RestrictionRegion = if ($script:SkuScenario -eq 'other-restriction') { 'elsewhere' } else { 'testregion' }
        $Restrictions = @(@{ type = if ($script:SkuScenario -eq 'unknown-restriction') { 'zone' } else { 'location' }; values = @(if ($script:SkuScenario -ne 'unscoped-restriction') { $RestrictionRegion }); reasonCode = if ($script:SkuScenario -eq 'quota') { 'QuotaId' } else { 'NotAvailableForSubscription' } })
    }
    $Sku = @{ name = $Name; kind = $Kind; resourceType = 'storageAccounts'; locations = @($Region); restrictions = $Restrictions }
    if ($script:SkuScenario -eq 'location-info') { $Sku.locations = @(); $Sku.locationInfo = @(@{ location = 'Test Region'; zones = @('1', '2', '3') }) }
    if ($script:SkuScenario -eq 'missing-location') { $Sku.locations = @() }
    if ($script:SkuScenario -eq 'missing-restrictions') { $Sku.Remove('restrictions') }
    if ($script:SkuScenario -eq 'paged' -and $Path -notlike '*page=2') { return (New-TestPage @(@{ name = 'Premium_LRS'; kind = 'FileStorage' }) "$Path&page=2") }
    New-TestPage @($Sku)
}
$StorageAccount = [pscustomobject]@{ StorageAccountName = 'profiles'; Id = '/subscriptions/storage/resourceGroups/files/providers/Microsoft.Storage/storageAccounts/profiles'; PrimaryLocation = 'Test Region'; Kind = 'FileStorage'; Sku = @{ Name = 'Premium_LRS' } }
$SkuCases = @(
    @{ Name = 'available'; State = 'Available'; Status = 'Warning' },
    @{ Name = 'other-region'; State = 'NotOfferedInRegion'; Status = 'Warning' },
    @{ Name = 'restricted'; State = 'RestrictedForSubscription'; Status = 'Warning' },
    @{ Name = 'quota'; State = 'RestrictedForSubscription'; Status = 'Warning' },
    @{ Name = 'other-restriction'; State = 'Available'; Status = 'Warning' },
    @{ Name = 'unknown-restriction'; State = 'Unknown'; Status = 'Error' },
    @{ Name = 'unscoped-restriction'; State = 'Unknown'; Status = 'Error' },
    @{ Name = 'wrong-kind'; State = 'Unknown'; Status = 'Error' },
    @{ Name = 'wrong-tier'; State = 'Unknown'; Status = 'Error' },
    @{ Name = 'denied'; State = 'Unknown'; Status = 'Error' },
    @{ Name = 'malformed'; State = 'Unknown'; Status = 'Error' },
    @{ Name = 'empty'; State = 'Unknown'; Status = 'Error' },
    @{ Name = 'missing-location'; State = 'Unknown'; Status = 'Error' },
    @{ Name = 'missing-restrictions'; State = 'Unknown'; Status = 'Error' },
    @{ Name = 'location-info'; State = 'Available'; Status = 'Warning' },
    @{ Name = 'paged'; State = 'Available'; Status = 'Warning' }
)
foreach ($Case in $SkuCases) {
    $script:SkuScenario = $Case.Name
    $script:SkuRequests = @()
    $Cache = @{}
    $Assessment = Get-AvdStorageReplicationAssessment -StorageAccount $StorageAccount -Cache $Cache
    Assert-Equal $Assessment.Evidence.ZrsAvailability.State $Case.State "$($Case.Name): ZRS availability; SKU response: $($Cache | ConvertTo-Json -Depth 12 -Compress)"
    Assert-Equal $Assessment.Status $Case.Status "$($Case.Name): no automatic exemption for LRS"
    Assert-Equal $Assessment.Evidence.ZrsAvailability.TargetSku 'Premium_ZRS' 'premium family preserved'
    Assert-Equal ($Assessment.Recommendation -like '*Use GRS*') $false 'no unsupported Premium GRS recommendation'
    $RequestCount = $script:SkuRequests.Count
    $null = Get-AvdStorageReplicationAssessment -StorageAccount $StorageAccount -Cache $Cache
    Assert-Equal $script:SkuRequests.Count $RequestCount 'SKU inventory cached per subscription'
    if ($Case.Name -eq 'paged') { Assert-Equal $RequestCount 2 'SKU pagination followed' }
}
$script:SkuScenario = 'v2'
$StorageAccount.Sku.Name = 'PremiumV2_LRS'
$V2Assessment = Get-AvdStorageReplicationAssessment -StorageAccount $StorageAccount -Cache @{}
Assert-Equal $V2Assessment.Evidence.ZrsAvailability.TargetSku 'PremiumV2_ZRS' 'V2 SKU family retained'
Assert-Equal $V2Assessment.Evidence.ZrsAvailability.State 'Available' 'V2 SKU availability matched'
$StorageAccount.Sku.Name = 'Premium_LRS'
$script:SkuScenario = 'available'
$script:SkuRequests = @()
$Cache = @{}
$null = Get-AvdStorageReplicationAssessment -StorageAccount $StorageAccount -Cache $Cache
$StorageAccount.Id = '/subscriptions/another/resourceGroups/files/providers/Microsoft.Storage/storageAccounts/profiles'
$null = Get-AvdStorageReplicationAssessment -StorageAccount $StorageAccount -Cache $Cache
Assert-Equal $script:SkuRequests.Count 2 'SKU availability is not reused across subscriptions'
$StorageAccount.Kind = 'StorageV2'
$StorageAccount.Sku.Name = 'Standard_LRS'
$script:SkuScenario = 'wrong-kind'
$StandardPremiumMismatch = Get-AvdStorageReplicationAssessment -StorageAccount $StorageAccount -Cache @{}
Assert-Equal $StandardPremiumMismatch.Evidence.ZrsAvailability.State 'Unknown' 'Premium availability cannot satisfy a Standard account'
$StorageAccount.Kind = 'FileStorage'
$StorageAccount.Sku.Name = 'Premium_ZRS'
$script:SkuRequests = @()
$Assessment = Get-AvdStorageReplicationAssessment -StorageAccount $StorageAccount -Cache @{}
Assert-Equal $Assessment.Status 'Pass' 'existing ZRS is not downgraded by unavailable SKU queries'
Assert-Equal $Assessment.Evidence.ZrsAvailability.State 'AlreadyConfigured' 'existing zone redundancy recorded'
Assert-Equal $script:SkuRequests.Count 0 'no availability lookup needed for existing ZRS'
$StorageAccount.Sku.Name = 'Standard_GRS'
$StorageAccount.Kind = 'StorageV2'
$script:SkuScenario = 'denied'
Assert-Equal (Get-AvdStorageReplicationAssessment -StorageAccount $StorageAccount -Cache @{}).Status 'Pass' 'configured geo redundancy remains separate from ZRS availability'
Write-Host 'PASS: ZRS availability by subscription, region, kind and SKU family; restrictions, unknowns, caching and pagination.'

$MdeFunction = $CollectorAst.Find({
    param($Node)
    $Node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -eq 'Get-AvdMdeEvidence'
}, $false)
. ([scriptblock]::Create($MdeFunction.Extent.Text))
function Invoke-AzRestMethod {
    param($Path, $Method, $ErrorAction)
    $script:MdeExtensionRequests += $Path
    if ($script:MdeExtensionScenario -eq 'denied') { return [pscustomobject]@{ StatusCode = 403; Content = '{}' } }
    if ($script:MdeExtensionScenario -eq 'empty') { return (New-TestPage @()) }
    $Publisher = if ($script:MdeExtensionScenario -eq 'wrong-publisher') { 'Example.Other' } else { 'Microsoft.Azure.AzureDefenderForServers' }
    $ExtensionType = if ($script:MdeExtensionScenario -eq 'mma') { 'MicrosoftMonitoringAgent' } else { 'MDE.Windows' }
    $State = if ($script:MdeExtensionScenario -eq 'failed') { 'Failed' } else { 'Succeeded' }
    $Properties = @{ publisher = $Publisher; type = $ExtensionType; provisioningState = $State; settings = @{ Secret = 'never-export-this' } }
    if ($script:MdeExtensionScenario -eq 'partial') { $Properties = @{} }
    New-TestPage @(@{ id = "$VmId/extensions/custom"; name = 'custom'; properties = $Properties })
}
foreach ($Scenario in @('renamed', 'failed', 'empty', 'wrong-publisher', 'mma', 'partial', 'denied')) {
    $script:MdeExtensionScenario = $Scenario
    $script:MdeExtensionRequests = @()
    $Cache = @{}
    $Mde = Get-AvdMdeEvidence -ResourceId $VmId -Cache $Cache
    Assert-Equal $Mde.Status 'Error' "$Scenario : extension is not EDR onboarding evidence"
    Assert-Equal $Mde.Onboarded $null "$Scenario : onboarding not fabricated"
    Assert-Equal $Mde.SensorHealth 'NotAssessed' "$Scenario : sensor health not fabricated"
    $ExpectedPresent = switch ($Scenario) { 'renamed' { $true }; 'failed' { $true }; 'partial' { $null }; 'denied' { $null }; default { $false } }
    Assert-Equal $Mde.DeploymentExtensionPresent $ExpectedPresent "$Scenario : exact publisher/type detection"
    Assert-Equal (($Mde | ConvertTo-Json -Depth 10) -like '*never-export-this*') $false 'MDE settings not exported'
    $null = Get-AvdMdeEvidence -ResourceId $VmId.ToUpperInvariant() -Cache $Cache
    Assert-Equal $script:MdeExtensionRequests.Count 1 'MDE per-resource cache'
    $null = Get-AvdMdeEvidence -ResourceId ($VmId -replace '/first/', '/second/') -Cache $Cache
    Assert-Equal $script:MdeExtensionRequests.Count 2 'same-name VM in another subscription stays distinct'
}
Write-Host 'PASS: MDE extension inventory is a deployment hint only; missing or failed extensions do not fabricate EDR state.'

foreach ($FunctionName in @('Invoke-AvdMdeHuntingQuery', 'Set-AvdMdeDeviceEvidence', 'Update-AvdMdeDeviceInventory')) {
    $HelperFunction = $CollectorAst.Find({
        param($Node)
        $Node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -eq $FunctionName
    }, $false)
    . ([scriptblock]::Create($HelperFunction.Extent.Text))
}
function Invoke-RestMethod {
    param($Uri, $Method, $Headers, $ContentType, $Body, $TimeoutSec, $MaximumRedirection, $ErrorAction)
    $script:HuntingCalls++
    Assert-Equal $Uri 'https://graph.microsoft.com/v1.0/security/runHuntingQuery' 'supported Graph hunting API'
    Assert-Equal $Method 'POST' 'read-only hunting query method'
    Assert-Equal $MaximumRedirection 0 'tokens not forwarded to redirected endpoints'
    $Request = ConvertFrom-Json -InputObject $Body
    Assert-Equal $Request.Timespan 'P7D' 'bounded DeviceInfo window'
    Assert-Equal $Request.Query.Contains('summarize arg_max(SnapshotIngestedAt, *) by DeviceId | where tolower(AzureResourceId) in (avdResources)') $true 'latest device record selected before matching resource'
    Assert-Equal $Request.Query.Contains('DeviceName ==') $false 'no hostname identity fallback'
    switch ($script:HuntingScenario) {
        'denied' { throw 'Do not export error body or bearer test-token' }
        'malformed' { [pscustomobject]@{ results = $null } }
        'partial' { [pscustomobject]@{ results = @(); error = 'partial results' } }
        'limit' { [pscustomobject]@{ results = @(1..10001) } }
        default { [pscustomobject]@{ results = @() } }
    }
}
foreach ($Scenario in @('empty', 'denied', 'malformed', 'partial', 'limit')) {
    $script:HuntingScenario = $Scenario
    $Response = Invoke-AvdMdeHuntingQuery -ResourceIds @($VmId) -Token 'test-token'
    Assert-Equal $Response.Ok ($Scenario -eq 'empty') "$Scenario : hunting response completeness"
    Assert-Equal (($Response | ConvertTo-Json -Depth 10).Contains('test-token')) $false 'no token or raw exception export'
}
$script:HuntingCalls = 0
$null = Invoke-AvdMdeHuntingQuery -ResourceIds @($VmId) -Token ''
Assert-Equal $script:HuntingCalls 0 'no request without credentials'

$InstanceId = '18b20622-80f2-4500-a856-cd6dbf6f18bf'
$Now = [DateTimeOffset]'2026-09-10T12:00:00Z'
foreach ($Scenario in @('healthy', 'impaired', 'inactive', 'unknown-health', 'not-onboarded', 'unsupported', 'stale', 'future', 'missing-timestamp', 'invalid-timestamp', 'stopped', 'unknown-power', 'no-match', 'wrong-resource', 'wrong-vmid', 'no-vmid', 'duplicate', 'merged-only', 'merged-and-current', 'query-denied')) {
    $script:MdeExtensionScenario = 'empty'
    $Evidence = Get-AvdMdeEvidence -ResourceId $VmId -Cache @{}
    $HostEntry = [pscustomobject]@{ ResourceId = $VmId; AzureVmId = $InstanceId; PowerStateCode = 'PowerState/running'; MDEInstalled = $null; MDEDiscovery = $Evidence }
    $Device = [pscustomobject]@{ DeviceId = 'device-1'; DeviceName = 'host'; AzureResourceId = $VmId.ToUpperInvariant(); AzureVmId = $InstanceId; Timestamp = '2026-09-10T10:00:00Z'; OnboardingStatus = 'Onboarded'; SensorHealthState = 'Active'; ClientVersion = '1'; MergedToDeviceId = '' }
    $Rows = @($Device)
    switch ($Scenario) {
        'impaired' { $Device.SensorHealthState = 'ImpairedCommunication' }
        'inactive' { $Device.SensorHealthState = 'Inactive' }
        'unknown-health' { $Device.SensorHealthState = 'Unknown' }
        'not-onboarded' { $Device.OnboardingStatus = 'Can be onboarded' }
        'unsupported' { $Device.OnboardingStatus = 'Unsupported' }
        'stale' { $Device.Timestamp = '2026-09-07T10:00:00Z' }
        'future' { $Device.Timestamp = '2026-09-11T10:00:00Z' }
        'missing-timestamp' { $Device.Timestamp = $null }
        'invalid-timestamp' { $Device.Timestamp = 'invalid' }
        'stopped' { $HostEntry.PowerStateCode = 'PowerState/deallocated' }
        'unknown-power' { $HostEntry.PowerStateCode = $null }
        'no-match' { $Rows = @() }
        'wrong-resource' { $Device.AzureResourceId = $VmId -replace '/first/', '/second/' }
        'wrong-vmid' { $Device.AzureVmId = '9224df28-f65e-40b3-bf80-3e00da9d0c44' }
        'no-vmid' { $Device.AzureVmId = $null }
        'duplicate' { $Rows += $Device.PSObject.Copy(); $Rows[1].DeviceId = 'device-2' }
        'merged-only' { $Device.MergedToDeviceId = 'device-2' }
        'merged-and-current' { $Rows += $Device.PSObject.Copy(); $Rows[1].DeviceId = 'old-device'; $Rows[1].MergedToDeviceId = 'device-1' }
    }
    Set-AvdMdeDeviceEvidence -SessionHost $HostEntry -QueryResult ([pscustomobject]@{ Ok = ($Scenario -ne 'query-denied'); Rows = $Rows; Error = '403 denied' }) -Now $Now
    $Expected = switch ($Scenario) { 'healthy' { 'Pass' }; 'merged-and-current' { 'Pass' }; 'impaired' { 'Warning' }; 'inactive' { 'Warning' }; 'not-onboarded' { 'Fail' }; default { 'Error' } }
    Assert-Equal $Evidence.Status $Expected "$Scenario : EDR state"
    Assert-Equal $Evidence.DeploymentExtensionPresent $false 'portal onboarding may be verified without extension'
    if ($Expected -eq 'Pass') { Assert-Equal $Evidence.Onboarded $true 'verified service onboarding'; Assert-Equal $Evidence.DeviceMatch 'ResourceAndVmId' 'VM incarnation verified' }
}
Write-Host 'PASS: Defender hunting identity, freshness, onboarding and sensor states; no matches/duplicates/errors are not EDR absence.'

$script:HuntingBatches = @()
$script:TenantTokenRequests = @()
$script:MdeTestTenant = 'original'
$script:MdeTestEnvironment = 'AzureCloud'
$script:MdeRestored = $false
function Get-AzContext {
    param($ErrorAction)
    [pscustomobject]@{ Tenant = @{ Id = $script:MdeTestTenant }; Environment = @{ Name = $script:MdeTestEnvironment } }
}
function Set-AzContext {
    param($Context, $SubscriptionId, $ErrorAction, $WarningAction)
    if ($Context) { $script:MdeTestTenant = $Context.Tenant.Id; $script:MdeRestored = $true }
    else { $script:MdeTestTenant = if ($SubscriptionId -eq 'first') { 'tenant-first' } else { 'tenant-second' } }
}
function Get-GraphTokenString {
    $script:TenantTokenRequests += $script:MdeTestTenant
    "token-for-$script:MdeTestTenant"
}
function Invoke-AvdMdeHuntingQuery {
    param([string[]]$ResourceIds, [string]$Token)
    Assert-Equal ($ResourceIds.Count -le 100) $true 'hunting batch size bounded'
    Assert-Equal $Token "token-for-$script:MdeTestTenant" 'matching tenant token'
    $ExpectedSubscription = if ($script:MdeTestTenant -eq 'tenant-first') { 'first' } else { 'second' }
    foreach ($ResourceId in $ResourceIds) { Assert-Equal (($ResourceId -split '/')[2]) $ExpectedSubscription 'no tenant mixing in hunting batch' }
    $script:HuntingBatches += ,$ResourceIds
    $Rows = @($ResourceIds | ForEach-Object {
        [pscustomobject]@{ DeviceId = $_; AzureResourceId = $_; AzureVmId = $InstanceId; Timestamp = [DateTimeOffset]::UtcNow.AddHours(-1).ToString('o'); OnboardingStatus = 'Onboarded'; SensorHealthState = 'Active'; MergedToDeviceId = '' }
    })
    [pscustomobject]@{ Ok = $true; Rows = $Rows; Error = $null }
}
$script:MdeExtensionScenario = 'empty'
$MdeHosts = @(foreach ($Index in 1..101) {
    $ResourceId = "$VmId-$Index"
    [pscustomobject]@{ ResourceId = $ResourceId; AzureVmId = $InstanceId; PowerStateCode = 'PowerState/running'; MDEInstalled = $null; MDEDiscovery = (Get-AvdMdeEvidence -ResourceId $ResourceId -Cache @{}) }
})
$SecondHost = $MdeHosts[0].PSObject.Copy()
$SecondHost.ResourceId = $VmId -replace '/first/', '/second/'
$SecondHost.MDEDiscovery = Get-AvdMdeEvidence -ResourceId $SecondHost.ResourceId -Cache @{}
$MdeHosts += $SecondHost
$MdeChecks = [System.Collections.ArrayList]::new()
foreach ($SessionHost in $MdeHosts) {
    [void]$MdeChecks.Add([pscustomobject]@{ Id = "SEC-MDE-$($SessionHost.ResourceId)"; Status = 'Error'; Details = ''; Evidence = $SessionHost.MDEDiscovery })
}
$MdeSubscriptions = @([pscustomobject]@{ Id = 'first'; TenantId = 'tenant-first' }, [pscustomobject]@{ Id = 'second'; TenantId = 'tenant-second' })
Update-AvdMdeDeviceInventory -SessionHosts $MdeHosts -Subscriptions $MdeSubscriptions -Checks $MdeChecks
Assert-Equal $script:HuntingBatches.Count 3 '102 hosts in two tenants use three bounded queries'
Assert-Equal $script:TenantTokenRequests.Count 2 'Graph token acquired per tenant'
Assert-Equal $script:MdeRestored $true 'Azure context restored after MDE queries'
Assert-Equal $script:MdeTestTenant 'original' 'original tenant restored'
Assert-Equal @($MdeChecks | Where-Object { $_.Status -eq 'Pass' }).Count 102 'verified Defender evidence reaches emitted checks'
Assert-Equal @($MdeHosts | Where-Object { $_.MDEInstalled -eq $true }).Count 102 'onboarding snapshot reaches inventory'
$script:HuntingBatches = @()
$script:MdeTestEnvironment = 'AzureUSGovernment'
Update-AvdMdeDeviceInventory -SessionHosts @($SecondHost) -Subscriptions $MdeSubscriptions -Checks $MdeChecks
Assert-Equal $script:HuntingBatches.Count 0 'no public Graph request for unsupported cloud'
Assert-Equal $SecondHost.MDEDiscovery.Status 'Error' 'unsupported cloud is unavailable evidence'
$MdeExport = ConvertFrom-Json -InputObject (ConvertTo-Json -Depth 10 -WarningAction Stop -InputObject @{ Inventory = @{ SessionHosts = @($MdeHosts[0]) }; CheckResults = @($MdeChecks[0]) })
Assert-Equal $MdeExport.Inventory.SessionHosts[0].MDEDiscovery.Devices[0].AzureVmId $InstanceId 'Defender identity survives export'
Assert-Equal $MdeExport.CheckResults[0].Evidence.DeviceMatch 'ResourceAndVmId' 'Defender evidence match survives export'
Write-Host 'PASS: tenant-specific Defender tokens, bounded batches, context restoration, cloud guard, and exported verdicts.'