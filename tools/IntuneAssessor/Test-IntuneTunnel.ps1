#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Invoke-IntuneDiscovery.ps1') -LibraryOnly
function Assert-Tunnel([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
$Root = 'https://graph.microsoft.com/beta/deviceManagement/microsoftTunnelSites'
Assert-Tunnel (-not (Test-IntuneUri $Root)) 'Tunnel became a default endpoint'
Assert-Tunnel (Test-IntuneUri $Root -AllowExpansion) 'Tunnel list blocked'
foreach ($Address in @($Root + '/site/requestUpgrade', $Root + '/site/microsoftTunnelServers/server/getHealthMetrics', $Root + '/site/microsoftTunnelServers/server/createServerLogCollectionRequest')) {
    Assert-Tunnel (-not (Test-IntuneUri $Address -AllowExpansion)) 'Tunnel action permitted'
}
$Inventory = [ordered]@{}
$States = [ordered]@{}
$Request = {
    param($Address)
    if ($Address -eq $Root) { return @{ StatusCode = 200; Body = ('{"value":[{"id":"site","upgradeAutomatically":false,"upgradeAvailable":true,"internalNetworkProbeUrl":"SECRET"}]}' | ConvertFrom-Json) } }
    return @{ StatusCode = 200; Body = ('{"value":[{"id":"server","tunnelServerHealthStatus":"healthy","lastCheckinDateTime":"2026-09-17T09:00:00Z","secret":"SECRET"}]}' | ConvertFrom-Json) }
}
Invoke-IntuneTunnel -Inventory $Inventory -States $States -Request $Request -Delay {} -Deadline ([datetime]::UtcNow.AddMinutes(1))
Assert-Tunnel ($States.TunnelServers.State -eq 'Complete' -and $States.TunnelServers.CompletedParentIds -contains 'site') 'Tunnel parent coverage missing'
Assert-Tunnel ($Inventory.TunnelServers[0].parentId -eq 'site' -and $Inventory.TunnelServers[0].lastCheckinDateTime) 'Tunnel identity/time lost'
Assert-Tunnel (-not ($Inventory | ConvertTo-Json -Depth 10).Contains('SECRET')) 'Unreviewed Tunnel fields leaked'
Invoke-IntuneTunnel -Inventory $Inventory -States $States -Request { param($Address) if ($Address -eq 'https://graph.microsoft.com/beta/deviceManagement/microsoftTunnelSites') { @{ StatusCode = 200; Body = @{ value = @(@{ id = 'site' }) } } } else { @{ StatusCode = 403 } } } -Delay {} -Deadline ([datetime]::UtcNow.AddMinutes(1))
Assert-Tunnel ($States.TunnelServers.State -eq 'Partial' -and $States.TunnelServers.CompletedParentIds.Count -eq 0) 'Failed Tunnel server list inferred complete'
Write-Output 'PASS: Tunnel GET-only boundaries, metadata projection, timestamps and failed-child coverage.'
$Document = Invoke-IntuneDiscoveryCore -SelectedTenant '22222222-2222-4222-8222-222222222222' -Tunnel $true -Request {
    param($Address)
    switch (([uri]$Address).AbsolutePath) {
        '/beta/deviceManagement/microsoftTunnelSites' { return @{ StatusCode = 200; Body = @{ value = @(@{ id = 'site'; upgradeAutomatically = $false; upgradeAvailable = $true }) } } }
        '/beta/deviceManagement/microsoftTunnelSites/site/microsoftTunnelServers' { return @{ StatusCode = 200; Body = @{ value = @(@{ id = 'server'; tunnelServerHealthStatus = 'healthy'; lastCheckinDateTime = [datetime]::UtcNow.ToString('o') }) } } }
        default { return @{ StatusCode = 200; Body = @{ value = @() } } }
    }
} -Requirements @{ ScopeConfirmed = $true; Assessor = 'Offline Tunnel test'; ScopeDescription = 'Selected synthetic site'; MaxCollectionAgeHours = 24; ConfigurationReview = @{ TunnelSiteIds = @('site'); MaxTunnelCheckinAgeHours = 2 } }
Assert-Tunnel ($Document.CollectionStatus.TunnelSites.State -eq 'Complete') 'Tunnel core switch not wired'
if ($env:ASSAY_INTUNE_TUNNEL_FIXTURE) { [IO.File]::WriteAllText($env:ASSAY_INTUNE_TUNNEL_FIXTURE, ($Document | ConvertTo-Json -Depth 30), [Text.UTF8Encoding]::new($false)) }
$Default = Invoke-IntuneDiscoveryCore -SelectedTenant '22222222-2222-4222-8222-222222222222' -Request { param($Address) if ($Address -match 'microsoftTunnel') { throw 'Unexpected Tunnel request' }; @{ StatusCode = 200; Body = @{ value = @() } } }
Assert-Tunnel ($Default.CollectionStatus.TunnelSites.State -eq 'NotRequested') 'Tunnel was enabled by default'