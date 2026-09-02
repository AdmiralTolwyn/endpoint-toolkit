<#
.SYNOPSIS
    Reports local Delivery Optimization statistics as a console table, PowerShell
    object, or single-line JSON for Grafana / Loki / Telegraf ingestion.

.DESCRIPTION
    Read-only. Collects the effective Delivery Optimization configuration, current
    service state, and month-to-date byte counters directly from the
    root/Microsoft/Windows/DeliveryOptimization CIM provider.

    The monthly download fields and efficiency calculations follow the source
    categories used by Windows Update for Business Delivery Optimization reports:
    BytesFromCDN, BytesFromCache, BytesFromPeers, BytesFromGroupPeers, TotalBytes,
    BandwidthSavingsPct, P2PEfficiencyPct, and ConnectedCacheEfficiencyPct.

    The local provider's MonthlyCdnBytes and per-transfer BytesFromHttp counters
    are inclusive HTTP totals: original CDN plus Microsoft Connected Cache. The
    script retains that total as BytesFromHTTP and derives direct BytesFromCDN by
    subtracting cache-host bytes so Connected Cache traffic is not counted twice.

    The local CIM provider also exposes Link-Local and Internet peer bytes. These
    are retained in separate all-source totals. The WUfB-compatible calculations
    use only LAN, Group, CDN, and Connected Cache bytes, matching the published
    report formulas. Neither local calculation has the same rolling 28-day window
    or cloud-side ISP cache classification used by WUfB reports.

.PARAMETER AsObject
    Emit the structured result object instead of the formatted console report.

.PARAMETER AsJson
    Emit the structured result as JSON instead of the formatted console report.
    JSON is compressed to one line unless -Pretty is also specified.

.PARAMETER OutputPath
    Optional. Append the compact JSON result to this file using UTF-8 without a BOM
    and a Unix newline. The selected console output is still produced.

.PARAMETER Pretty
    Emit indented JSON with -AsJson. This does not affect the compact JSON appended
    by -OutputPath.

.EXAMPLE
    .\Get-DeliveryOptimizationStatistics.ps1
    Display a formatted configuration, efficiency, download-source, upload, and
    current-state report.

.EXAMPLE
    .\Get-DeliveryOptimizationStatistics.ps1 -AsObject
    Return a structured object for PowerShell processing.

.EXAMPLE
    .\Get-DeliveryOptimizationStatistics.ps1 -AsJson
    Emit one compact JSON object for a Grafana Agent, Loki, or Telegraf exec input.

.EXAMPLE
    .\Get-DeliveryOptimizationStatistics.ps1 -OutputPath 'C:\ProgramData\DOMonitor\do_stats.log'
    Display the table and append the same result as one JSON line.

.NOTES
    File:     windows/diagnostics/DeliveryOptimizationStatistics/Get-DeliveryOptimizationStatistics.ps1
    Author:   Anton Romanyuk
    Version:  1.3.0
    Requires: PowerShell 5.1+, DeliveryOptimization module, CIM access.

    Data scope:
      - Monthly byte counters: local calendar month to collection time
      - Current service fields: point-in-time snapshot
      - WUfB report comparison: WUfB uses cloud telemetry over rolling 7/28-day windows

    Exit codes:
      0 - Statistics collected successfully
      4 - Collection failed; a structured error is emitted

    References:
      https://learn.microsoft.com/windows/deployment/update/wufb-reports-do
      https://learn.microsoft.com/windows/deployment/update/wufb-reports-schema-ucdostatus

.DISCLAIMER
    THIS SCRIPT IS PROVIDED "AS-IS" WITHOUT WARRANTY OF ANY KIND. Test against a
    representative device before deploying through an endpoint management platform.
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [switch] $AsObject,

    [switch] $AsJson,

    [string] $OutputPath,

    [switch] $Pretty
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$DO_NAMESPACE = 'root/Microsoft/Windows/DeliveryOptimization'
$EXIT_OK       = 0
$EXIT_ERROR    = 4

function ConvertTo-Percentage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [double] $Numerator,
        [Parameter(Mandatory)] [double] $Denominator
    )

    if ($Denominator -le 0) { return 0.0 }
    return [Math]::Round(($Numerator / $Denominator) * 100.0, 2)
}

function ConvertTo-GiB {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [double] $Bytes
    )

    return [Math]::Round($Bytes / 1GB, 3)
}

function ConvertTo-DisplaySize {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [double] $Bytes
    )

    if ($Bytes -eq 0) { return '0 B' }
    if ($Bytes -lt 1MB) { return ('{0:N2} KiB' -f ($Bytes / 1KB)) }
    if ($Bytes -lt 1GB) { return ('{0:N2} MiB' -f ($Bytes / 1MB)) }
    return ('{0:N2} GiB' -f ($Bytes / 1GB))
}

function ConvertTo-FriendlyProvider {
    [CmdletBinding()]
    param(
        [AllowNull()] [string] $Provider
    )

    switch ($Provider) {
        'MdmProvider'      { return 'MDM' }
        'RegistryProvider' { return 'Registry/GP' }
        'SettingsProvider' { return 'Settings' }
        'AdminProvider'    { return 'Admin' }
        'DefaultProvider'  { return 'Default' }
        default            { if ($Provider) { return $Provider } else { return '-' } }
    }
}

function Write-ReportTable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object[]] $Rows
    )

    if ($Rows.Count -eq 0) { return }
    Write-Host (($Rows | Format-Table -AutoSize | Out-String).Trim())
}

function ConvertTo-DurationDisplay {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [double] $Seconds
    )

    if ($Seconds -ge 86400 -and ($Seconds % 86400) -eq 0) {
        return ('{0:N0} days ({1:N0} seconds)' -f ($Seconds / 86400), $Seconds)
    }
    if ($Seconds -ge 60 -and ($Seconds % 60) -eq 0) {
        return ('{0:N0} minutes ({1:N0} seconds)' -f ($Seconds / 60), $Seconds)
    }
    return ('{0:N0} seconds' -f $Seconds)
}

function ConvertTo-DOPolicyDisplay {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Name,
        [Parameter(Mandatory)] [AllowNull()] $Value
    )

    $displayValue = "$Value"
    switch ($Name.ToLowerInvariant()) {
        'doabsolutemaxcachesize'                   { $displayValue = '{0} GB' -f $Value }
        'doallowvpnpeercaching'                    { $displayValue = if ([int] $Value -eq 1) { 'Allowed' } else { 'Not allowed' } }
        'docachehostsource'                        { $displayValue = switch ([int] $Value) { 1 { 'DHCP option 235 (1)' } 2 { 'DHCP option 235 force (2)' } default { "Unknown ($Value)" } } }
        'dodelaybackgrounddownloadfromhttp'        { $displayValue = ConvertTo-DurationDisplay -Seconds ([double] $Value) }
        'dodelayforegrounddownloadfromhttp'        { $displayValue = ConvertTo-DurationDisplay -Seconds ([double] $Value) }
        'dodelaycacheserverfallbackbackground'     { $displayValue = ConvertTo-DurationDisplay -Seconds ([double] $Value) }
        'dodelaycacheserverfallbackforeground'     { $displayValue = ConvertTo-DurationDisplay -Seconds ([double] $Value) }
        'dodisallowcacheserverdownloadsonvpn'      { $displayValue = if ([int] $Value -eq 1) { 'Disallowed' } else { 'Allowed (default)' } }
        'dodownloadmode'                           { $displayValue = '{0} ({1})' -f (Resolve-DownloadMode -Value ([int] $Value)), $Value }
        'dogroupidsource'                          { $displayValue = switch ([int] $Value) { 0 { 'Not set (0)' } 1 { 'AD site (1)' } 2 { 'Authenticated domain SID (2)' } 3 { 'DHCP option 234 (3)' } 4 { 'DNS suffix (4)' } 5 { 'Entra tenant ID (5)' } default { "Unknown ($Value)" } } }
        'domaxbackgrounddownloadbandwidth'         { $displayValue = '{0} KB/s' -f $Value }
        'domaxcacheage'                            { $displayValue = ConvertTo-DurationDisplay -Seconds ([double] $Value) }
        'domaxcachesize'                           { $displayValue = '{0}% of available disk space' -f $Value }
        'domaxforegrounddownloadbandwidth'         { $displayValue = '{0} KB/s' -f $Value }
        'dominbackgroundqos'                       { $displayValue = '{0} KB/s' -f $Value }
        'dominbatterypercentageallowedtoupload'    { $displayValue = 'Allowed at >= {0}%' -f $Value }
        'domindisksizeallowedtopeer'               { $displayValue = '{0} GB' -f $Value }
        'dominfilesizetocache'                     { $displayValue = '{0} MB' -f $Value }
        'dominramallowedtopeer'                    { $displayValue = '{0} GB' -f $Value }
        'domonthlyuploaddatacap'                   { $displayValue = '{0} GB' -f $Value }
        'dopercentagemaxbackgroundbandwidth'       { $displayValue = '{0}%' -f $Value }
        'dopercentagemaxforegroundbandwidth'       { $displayValue = '{0}%' -f $Value }
        'dorestrictpeerselectionby'                { $displayValue = switch ([int] $Value) { 0 { 'None (0)' } 1 { 'Subnet mask (1)' } 2 { 'Local discovery / DNS-SD (2)' } default { "Unknown ($Value)" } } }
    }

    return $displayValue
}

function Get-BandwidthSavingsStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [double] $BandwidthSavingsPct,
        [Parameter(Mandatory)] [double] $TotalBytes
    )

    if ($TotalBytes -le 0) { return 'NO_DATA' }
    if ($BandwidthSavingsPct -lt 10) { return 'ERROR' }
    if ($BandwidthSavingsPct -le 60) { return 'WARNING' }
    return 'OK'
}

function Resolve-DownloadMode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [int] $Value
    )

    switch ($Value) {
        0       { return 'HTTP Only' }
        1       { return 'LAN' }
        2       { return 'Group' }
        3       { return 'Internet' }
        99      { return 'Simple' }
        100     { return 'Bypass (deprecated)' }
        default { return "Unknown ($Value)" }
    }
}

function Resolve-SwarmStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [int] $Value
    )

    switch ($Value) {
        0       { return 'Downloading' }
        1       { return 'Complete' }
        2       { return 'Caching' }
        3       { return 'Paused' }
        default { return "Unknown ($Value)" }
    }
}

function Resolve-PeerType {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [int] $Value
    )

    switch ($Value) {
        0       { return 'None' }
        1       { return 'LAN' }
        2       { return 'CDN' }
        3       { return 'Internet' }
        4       { return 'Group' }
        5       { return 'DOINC' }
        6       { return 'LocalCache' }
        7       { return 'LinkLocal' }
        default { return "Unknown ($Value)" }
    }
}

function Get-EntraIdentity {
    [CmdletBinding()]
    param()

    $identity = [ordered]@{
        AzureAdJoined  = $false
        AzureADDeviceId = $null
        AzureADTenantId = $null
        TenantName      = $null
    }

    $joinRoot = 'HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\JoinInfo'
    if (-not (Test-Path -LiteralPath $joinRoot)) { return [pscustomobject] $identity }

    foreach ($joinKey in Get-ChildItem -LiteralPath $joinRoot -ErrorAction SilentlyContinue) {
        $join = Get-ItemProperty -LiteralPath $joinKey.PSPath -ErrorAction SilentlyContinue
        $certificate = Get-Item -LiteralPath ('Cert:\LocalMachine\My\' + $joinKey.PSChildName) -ErrorAction SilentlyContinue
        if ($null -eq $join -or $null -eq $certificate) { continue }

        $deviceMatch = [regex]::Match($certificate.Subject, '(?i)^CN=(?<id>[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})$')
        if (-not $deviceMatch.Success) { continue }

        $identity.AzureAdJoined = $true
        $identity.AzureADDeviceId = $deviceMatch.Groups['id'].Value.ToLowerInvariant()
        $identity.AzureADTenantId = "$($join.TenantId)".ToLowerInvariant()

        $tenantPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\TenantInfo\' + $identity.AzureADTenantId
        if (Test-Path -LiteralPath $tenantPath) {
            $tenant = Get-ItemProperty -LiteralPath $tenantPath -ErrorAction SilentlyContinue
            if ($tenant) { $identity.TenantName = $tenant.DisplayName }
        }
        break
    }

    return [pscustomobject] $identity
}

function Get-DOPolicyValues {
    [CmdletBinding()]
    param()

    $sources = [ordered]@{
        MDM         = [ordered]@{}
        GroupPolicy = [ordered]@{}
    }
    $paths = [ordered]@{
        MDM         = 'HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\DeliveryOptimization'
        GroupPolicy = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization'
    }

    foreach ($sourceName in $paths.Keys) {
        $path = $paths[$sourceName]
        if (-not (Test-Path -LiteralPath $path)) { continue }

        $item = Get-ItemProperty -LiteralPath $path -ErrorAction Stop
        $properties = $item.PSObject.Properties |
            Where-Object { $_.Name -notlike 'PS*' -and $_.Name -notmatch '_(ProviderSet|WinningProvider)$' } |
            Sort-Object Name

        foreach ($property in $properties) {
            $sources[$sourceName][$property.Name] = $property.Value
        }
    }

    return [pscustomobject]@{
        MDM         = [pscustomobject] $sources.MDM
        GroupPolicy = [pscustomobject] $sources.GroupPolicy
    }
}

function Get-ObjectPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowNull()] $InputObject,
        [Parameter(Mandatory)] [string] $Name
    )

    if ($null -eq $InputObject) { return $null }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Get-WUfBGroupIdHash {
    [CmdletBinding()]
    param(
        [AllowNull()] [string] $GroupId
    )

    if ([string]::IsNullOrWhiteSpace($GroupId)) { return $null }
    $algorithm = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::Unicode.GetBytes($GroupId + [char] 0)
        return [Convert]::ToBase64String($algorithm.ComputeHash($bytes))
    }
    finally {
        $algorithm.Dispose()
    }
}

function ConvertTo-CurrentTransfer {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Transfer,
        [AllowNull()] $PeerInfo
    )

    $peerRows = @()
    if ($PeerInfo -and $PeerInfo.IPs) {
        for ($index = 0; $index -lt $PeerInfo.IPs.Count; $index++) {
            $peerTypeId = [int] $PeerInfo.PeerTypes[$index]
            $peerRows += [pscustomobject]@{
                PeerTypeId           = $peerTypeId
                PeerType             = Resolve-PeerType -Value $peerTypeId
                ConnectionEstablished = [bool] $PeerInfo.ConnectionEstablished[$index]
                BytesSent            = [uint64] $PeerInfo.BytesSent[$index]
                BytesReceived        = [uint64] $PeerInfo.BytesReceived[$index]
                UploadRateBytes      = [uint32] $PeerInfo.UploadRates[$index]
                DownloadRateBytes    = [uint32] $PeerInfo.DownloadRates[$index]
            }
        }
    }

    $peerTypeCounts = @(
        $peerRows | Group-Object PeerTypeId | ForEach-Object {
            [pscustomobject]@{
                PeerTypeId = [int] $_.Name
                PeerType   = Resolve-PeerType -Value ([int] $_.Name)
                Count      = [int] $_.Count
            }
        }
    )

    $sourceHost = $null
    if ($Transfer.SourceURL) {
        try { $sourceHost = ([uri] $Transfer.SourceURL).Host } catch { }
    }
    $cacheHost = $null
    if ($Transfer.CacheHost) {
        try { $cacheHost = ([uri] $Transfer.CacheHost).Host } catch { $cacheHost = "$($Transfer.CacheHost)" }
    }

    $peerBytes = [uint64] ($Transfer.BytesFromLanPeers + $Transfer.BytesFromGroupPeers + $Transfer.BytesFromInternetPeers + $Transfer.BytesFromLinkLocalPeers)
    $downloadModeId = [int] $Transfer.DownloadMode
    $expireOn = $null
    if ($Transfer.ExpireOn -and ([datetime] $Transfer.ExpireOn).Year -gt 1) {
        $expireOn = ([datetime] $Transfer.ExpireOn).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
    }

    return [pscustomobject]@{
        FileId                     = "$($Transfer.FileId)"
        Status                     = Resolve-SwarmStatus -Value ([int] $Transfer.Status)
        StatusId                   = [int] $Transfer.Status
        Priority                   = if ($Transfer.IsBackground) { 'Background' } else { 'Foreground' }
        DownloadMode               = Resolve-DownloadMode -Value $downloadModeId
        DownloadModeId             = $downloadModeId
        PeerEligible               = ($downloadModeId -in @(1, 2, 3))
        CallerApplication          = "$($Transfer.PredefinedCallerApplication)"
        SourceHost                 = $sourceHost
        CacheHost                  = $cacheHost
        FileSize                   = [uint64] $Transfer.FileSize
        FileSizeInCache            = [uint64] $Transfer.FileSizeInCache
        TotalBytesDownloaded       = [uint64] $Transfer.TotalBytesDownloaded
        BytesFromHTTP              = [uint64] $Transfer.BytesFromHttp
        BytesFromCDN               = if ([uint64] $Transfer.BytesFromHttp -ge [uint64] $Transfer.BytesFromCacheServer) { [uint64] ([uint64] $Transfer.BytesFromHttp - [uint64] $Transfer.BytesFromCacheServer) } else { [uint64] 0 }
        BytesFromCache             = [uint64] $Transfer.BytesFromCacheServer
        BytesFromPeers             = [uint64] $Transfer.BytesFromLanPeers
        BytesFromGroupPeers        = [uint64] $Transfer.BytesFromGroupPeers
        BytesFromIntPeers          = [uint64] $Transfer.BytesFromInternetPeers
        BytesFromLinkLocal         = [uint64] $Transfer.BytesFromLinkLocalPeers
        TotalPeerBytes             = $peerBytes
        BytesToLanPeers            = [uint64] $Transfer.BytesToLanPeers
        BytesToGroupPeers          = [uint64] $Transfer.BytesToGroupPeers
        BytesToInternetPeers       = [uint64] $Transfer.BytesToInternetPeers
        BytesToLinkLocal           = [uint64] $Transfer.BytesToLinkLocalPeers
        DownloadDurationSeconds    = [Math]::Round(([double] $Transfer.DownloadDurationMsecs / 1000.0), 3)
        HTTPConnections            = [uint32] $Transfer.HttpConnectionCount
        CacheConnections           = [uint32] $Transfer.CacheServerConnectionCount
        LANConnections             = [uint32] $Transfer.LanConnectionCount
        LinkLocalConnections       = [uint32] $Transfer.LinkLocalConnectionCount
        GroupConnections           = [uint32] $Transfer.GroupConnectionCount
        InternetConnections        = [uint32] $Transfer.InternetConnectionCount
        PeerCount                  = [uint32] $Transfer.PeerCount
        PeerDetailsCaptured        = [int] $peerRows.Count
        PeerConnectionsEstablished = [int] @($peerRows | Where-Object ConnectionEstablished).Count
        PeerConnectionsFailed      = [int] @($peerRows | Where-Object { -not $_.ConnectionEstablished }).Count
        PeerTypes                  = $peerTypeCounts
        ExpireOn                   = $expireOn
        IsPinned                   = [bool] $Transfer.IsPinned
    }
}

function Write-JsonLine {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $InputObject,
        [Parameter(Mandatory)] [string] $Path
    )

    $directory = Split-Path -Path $Path -Parent
    if ($directory -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $json = $InputObject | ConvertTo-Json -Depth 6 -Compress
    [System.IO.File]::AppendAllText(
        $Path,
        ($json + "`n"),
        (New-Object System.Text.UTF8Encoding $false)
    )
}

function Write-ConsoleReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Result
    )

    $configuredModeProvider = ConvertTo-FriendlyProvider $Result.configuration.DownloadModeProvider
    $effective = $Result.configuration.Effective
    $mdm = $Result.configuration.PolicyValues.MDM
    $groupPolicy = $Result.configuration.PolicyValues.GroupPolicy

    Write-Host ''
    Write-Host '=== Delivery Optimization Statistics ===' -ForegroundColor Cyan
    Write-Host ("Computer: {0}   OS: {1} ({2})   Collected: {3}" -f `
        $Result.hostname, $Result.display_version, $Result.os_build, $Result.period_end_local) -ForegroundColor DarkGray
    Write-Host ("Period:   {0} to {1} ({2}; local calendar month)" -f `
        $Result.period_start_local, $Result.period_end_local, $Result.period_timezone) -ForegroundColor DarkGray
    Write-Host 'Scope:    Local counters; not the WUfB rolling 7/28-day cloud telemetry.' -ForegroundColor DarkGray

    Write-Host ''
    Write-Host 'Effective configuration' -ForegroundColor Cyan
    $configurationRows = @(
        [pscustomobject]@{ Setting = 'Download mode';                 Value = '{0} ({1})' -f $Result.configuration.DownloadMode, $Result.configuration.DownloadModeId; Source = $configuredModeProvider }
        [pscustomobject]@{ Setting = 'Peering';                       Value = $Result.configuration.PeeringStatus; Source = 'Derived from mode' }
        [pscustomobject]@{ Setting = 'Minimum RAM for peering';       Value = '{0} GB' -f $effective.MinTotalRAMGB; Source = ConvertTo-FriendlyProvider $effective.MinTotalRAMProvider }
        [pscustomobject]@{ Setting = 'Minimum disk for peering';      Value = '{0} GB' -f $effective.MinTotalDiskSizeGB; Source = ConvertTo-FriendlyProvider $effective.MinTotalDiskSizeProvider }
        [pscustomobject]@{ Setting = 'Upload on battery';             Value = 'Allowed at >= {0}%' -f $effective.BatteryPctToSeed; Source = ConvertTo-FriendlyProvider $effective.BatteryPctToSeedProvider }
        [pscustomobject]@{ Setting = 'Peer caching on VPN';           Value = if ($effective.VpnPeerCachingAllowed) { 'Allowed' } else { 'Blocked' }; Source = ConvertTo-FriendlyProvider $effective.VpnPeerCachingAllowedProvider }
        [pscustomobject]@{ Setting = 'Background bandwidth limit';    Value = '{0}%' -f $effective.BackgroundDownloadLimitPct; Source = ConvertTo-FriendlyProvider $effective.BackgroundDownloadLimitPctProvider }
        [pscustomobject]@{ Setting = 'Foreground bandwidth limit';    Value = '{0}%' -f $effective.ForegroundDownloadLimitPct; Source = ConvertTo-FriendlyProvider $effective.ForegroundDownloadLimitPctProvider }
    )
    Write-ReportTable -Rows $configurationRows

    $policyRows = @()
    foreach ($source in @(
        [pscustomobject]@{ Name = 'MDM'; Values = $mdm }
        [pscustomobject]@{ Name = 'Group Policy'; Values = $groupPolicy }
    )) {
        foreach ($property in @($source.Values.PSObject.Properties | Sort-Object Name)) {
            $policyRows += [pscustomobject]@{
                Policy = $property.Name
                Value  = ConvertTo-DOPolicyDisplay -Name $property.Name -Value $property.Value
                Source = $source.Name
            }
        }
    }
    if ($policyRows.Count -gt 0) {
        Write-Host ''
        Write-Host 'Applied policy overrides' -ForegroundColor Cyan
        Write-ReportTable -Rows $policyRows
    }

    Write-Host ''
    Write-Host 'Month-to-date summary' -ForegroundColor Cyan
    Write-Host ('Downloaded {0}; alternate sources {1} ({2:N2}%); peers {3} ({4:N2}%); uploaded {5}.' -f `
        (ConvertTo-DisplaySize $Result.download.TotalBytes),
        (ConvertTo-DisplaySize $Result.download.WUfBLocalSourceBytes),
        $Result.efficiency.BandwidthSavingsPct,
        (ConvertTo-DisplaySize $Result.download.WUfBPeerBytes),
        $Result.efficiency.P2PEfficiencyPct,
        (ConvertTo-DisplaySize $Result.upload.TotalBytes))

    $signalColor = switch ($Result.efficiency.BandwidthSavingsStatus) {
        'ERROR'   { 'Red' }
        'WARNING' { 'Yellow' }
        'OK'      { 'Green' }
        default   { 'DarkGray' }
    }
    Write-Host ('Workbook signal: {0} (Microsoft flags savings <= 60%; comparative signal, not device health).' -f `
        $Result.efficiency.BandwidthSavingsStatus) -ForegroundColor $signalColor
    $observations = New-Object System.Collections.Generic.List[string]
    if ($Result.configuration.PeerConfigured -and -not $Result.configuration.PeerUsedThisMonth) {
        if ($Result.configuration.MCCUsedThisMonth) {
            $observations.Add('Peering is enabled, but no peer traffic is recorded this month; alternate-source savings came from Connected Cache.')
        } else {
            $observations.Add('Peering is enabled, but no peer traffic is recorded this month.')
        }
    }
    if (-not $Result.configuration.PeerConfigured -and $Result.configuration.MCCUsedThisMonth) {
        if ($Result.configuration.CacheHostConfigured) {
            $observations.Add(('Peering is off, but {0:N2}% of downloads came from the configured Microsoft Connected Cache host ({1}). MCC serves content over HTTP independently of peer-to-peer, so this is expected.' -f $Result.efficiency.ConnectedCacheEfficiencyPct, $Result.configuration.ConfiguredCacheHost))
        } else {
            $observations.Add(('Peering is off, but {0:N2}% of downloads came from Microsoft Connected Cache with no enterprise cache host configured - this is an ISP-operated MCC (e.g. Deutsche Telekom) assigned automatically by the DO cloud service. WUfB reports exclude ISP-hosted MCC from BytesFromCache, so this local percentage will not match the WUfB report.' -f $Result.efficiency.ConnectedCacheEfficiencyPct))
        }
        if (@($Result.configuration.ObservedCacheHosts).Count -gt 0) {
            $observations.Add(('Observed Connected Cache host(s) this run: {0}.' -f (@($Result.configuration.ObservedCacheHosts) -join ', ')))
        }
    }
    foreach ($obs in $observations) {
        Write-Host ('Observation: {0}' -f $obs) -ForegroundColor Yellow
    }

    Write-Host ''
    Write-Host 'Efficiency (WUfB workbook formula, local month-to-date)' -ForegroundColor Cyan
    $efficiencyRows = @(
        [pscustomobject]@{ Metric = 'Bandwidth savings'; Percent = $Result.efficiency.BandwidthSavingsPct }
        [pscustomobject]@{ Metric = 'P2P efficiency'; Percent = $Result.efficiency.P2PEfficiencyPct }
        [pscustomobject]@{ Metric = 'Connected Cache efficiency'; Percent = $Result.efficiency.ConnectedCacheEfficiencyPct }
        [pscustomobject]@{ Metric = 'CDN share'; Percent = $Result.efficiency.CDNPercentage }
    )
    Write-ReportTable -Rows $efficiencyRows

    if ($Result.download.BytesFromIntPeers -gt 0 -or $Result.download.BytesFromLinkLocal -gt 0) {
        Write-Host ('All local sources including Internet/Link-Local peers: savings {0:N2}%, P2P {1:N2}%.' -f `
            $Result.efficiency.AllSourcesBandwidthSavingsPct, $Result.efficiency.AllSourcesP2PEfficiencyPct) -ForegroundColor DarkGray
    }

    Write-Host ''
    Write-Host 'Month-to-date downloads' -ForegroundColor Cyan
    $downloadRows = @()
    foreach ($source in @(
        [pscustomobject]@{ Name = 'Direct CDN';       Bytes = $Result.download.BytesFromCDN }
        [pscustomobject]@{ Name = 'Connected Cache';  Bytes = $Result.download.BytesFromCache }
        [pscustomobject]@{ Name = 'LAN peers';        Bytes = $Result.download.BytesFromPeers }
        [pscustomobject]@{ Name = 'Group peers';      Bytes = $Result.download.BytesFromGroupPeers }
        [pscustomobject]@{ Name = 'Internet peers';   Bytes = $Result.download.BytesFromIntPeers }
        [pscustomobject]@{ Name = 'Link-Local peers'; Bytes = $Result.download.BytesFromLinkLocal }
    )) {
        if ($source.Bytes -le 0) { continue }
        $downloadRows += [pscustomobject]@{
            Source   = $source.Name
            Amount   = ConvertTo-DisplaySize $source.Bytes
            SharePct = ConvertTo-Percentage $source.Bytes $Result.download.TotalBytes
        }
    }
    $downloadRows += [pscustomobject]@{ Source = 'TOTAL'; Amount = ConvertTo-DisplaySize $Result.download.TotalBytes; SharePct = ConvertTo-Percentage $Result.download.TotalBytes $Result.download.TotalBytes }
    Write-ReportTable -Rows $downloadRows
    Write-Host ('Average download rates: foreground {0:N2} Kbps; background {1:N2} Kbps.' -f `
        $Result.download.AverageForegroundKbps, $Result.download.AverageBackgroundKbps) -ForegroundColor DarkGray

    Write-Host ''
    Write-Host 'Month-to-date uploads' -ForegroundColor Cyan
    if ($Result.upload.TotalBytes -eq 0) {
        Write-Host 'No peer uploads recorded this month.' -ForegroundColor DarkGray
    } else {
        $uploadRows = @(
            [pscustomobject]@{ Destination = 'LAN peers';        Amount = ConvertTo-DisplaySize $Result.upload.BytesToLanPeers }
            [pscustomobject]@{ Destination = 'Group peers';      Amount = ConvertTo-DisplaySize $Result.upload.BytesToGroupPeers }
            [pscustomobject]@{ Destination = 'Internet peers';   Amount = ConvertTo-DisplaySize $Result.upload.BytesToInternetPeers }
            [pscustomobject]@{ Destination = 'Link-Local peers'; Amount = ConvertTo-DisplaySize $Result.upload.BytesToLinkLocal }
            [pscustomobject]@{ Destination = 'TOTAL';            Amount = ConvertTo-DisplaySize $Result.upload.TotalBytes }
        )
        Write-ReportTable -Rows $uploadRows
    }

    Write-Host ''
    Write-Host 'Current service state' -ForegroundColor Cyan
    $diskFreePct = ConvertTo-Percentage $Result.current.DiskAvailableBytes $Result.current.DiskTotalBytes
    $currentRows = @(
        [pscustomobject]@{ Metric = 'Status';            Value = $Result.current.DOStatusDescription }
        [pscustomobject]@{ Metric = 'Resident DO cache'; Value = ConvertTo-DisplaySize $Result.current.CacheSizeBytes }
        [pscustomobject]@{ Metric = 'Cache-drive free';  Value = '{0} ({1:N2}%)' -f (ConvertTo-DisplaySize $Result.current.DiskAvailableBytes), $diskFreePct }
        [pscustomobject]@{ Metric = 'Peers / connections'; Value = '{0} / {1}' -f $Result.current.PeerCount, $Result.current.TotalConnections }
        [pscustomobject]@{ Metric = 'Active / pending jobs'; Value = '{0} / {1}' -f $Result.current.ActiveJobs, $Result.current.PendingJobs }
        [pscustomobject]@{ Metric = 'DO service CPU';    Value = '{0:N4}%' -f $Result.current.CpuUsagePct }
        [pscustomobject]@{ Metric = 'DO service memory'; Value = '{0:N2} MB' -f ([double] $Result.current.MemUsageKB / 1KB) }
    )
    Write-ReportTable -Rows $currentRows

    if ($Result.current.TransferCount -gt 0) {
        Write-Host 'Current transfers' -ForegroundColor Cyan
        $Result.transfers |
            Select-Object Status, Priority, DownloadMode, CallerApplication, TotalBytesDownloaded, TotalPeerBytes, PeerCount, DownloadDurationSeconds |
            Format-Table -AutoSize
    }

}

if ($AsObject -and $AsJson) {
    throw 'Specify either -AsObject or -AsJson, not both.'
}

try {
    Import-Module DeliveryOptimization -DisableNameChecking -ErrorAction Stop

    $rawConfig    = Get-CimInstance -Namespace $DO_NAMESPACE -ClassName MSFT_DeliveryOptimizationConfig -ErrorAction Stop
    $download     = Get-CimInstance -Namespace $DO_NAMESPACE -ClassName MSFT_DODownloadUsage -ErrorAction Stop
    $upload       = Get-CimInstance -Namespace $DO_NAMESPACE -ClassName MSFT_DOUploadUsage -ErrorAction Stop
    $current      = Get-CimInstance -Namespace $DO_NAMESPACE -ClassName MSFT_DOCurrentStatus -ErrorAction Stop
    $friendly     = Get-DOConfig -Verbose -ErrorAction Stop 4>$null
    $rawTransfers = @(Get-CimInstance -Namespace $DO_NAMESPACE -ClassName MSFT_DeliveryOptimizationFile -WarningAction SilentlyContinue -ErrorAction Stop)
    $rawPeerInfo  = @(Get-CimInstance -Namespace $DO_NAMESPACE -ClassName MSFT_DeliveryOptimizationFilePeerInfo -WarningAction SilentlyContinue -ErrorAction Stop)
    $entraIdentity = Get-EntraIdentity
    $policyValues = Get-DOPolicyValues
    $os           = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    $currentBuild = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop

    $bytesFromHTTP       = [uint64] $download.MonthlyCdnBytes
    $bytesFromCache      = [uint64] $download.MonthlyCacheHostBytes
    $bytesFromCDN        = if ($bytesFromHTTP -ge $bytesFromCache) { [uint64] ($bytesFromHTTP - $bytesFromCache) } else { [uint64] 0 }
    $bytesFromPeers      = [uint64] $download.MonthlyLanBytes
    $bytesFromGroup      = [uint64] $download.MonthlyGroupBytes
    $bytesFromInternet   = [uint64] $download.MonthlyInternetBytes
    $bytesFromLinkLocal  = [uint64] $download.MonthlyLinkLocalBytes
    $wufbPeerBytes       = [double] ($bytesFromPeers + $bytesFromGroup)
    $wufbTotalBytes      = [double] ($bytesFromHTTP + $wufbPeerBytes)
    $wufbLocalBytes      = [double] ($bytesFromCache + $wufbPeerBytes)
    $totalPeerBytes      = [double] ($bytesFromPeers + $bytesFromGroup + $bytesFromInternet + $bytesFromLinkLocal)
    $totalDownloadBytes  = [double] ($bytesFromHTTP + $totalPeerBytes)
    $localSourceBytes    = [double] ($bytesFromCache + $totalPeerBytes)
    $wufbSavingsPct      = ConvertTo-Percentage -Numerator $wufbLocalBytes -Denominator $wufbTotalBytes

    $uploadLanBytes       = [uint64] $upload.MonthlyLanBytes
    $uploadGroupBytes     = [uint64] $upload.MonthlyGroupBytes
    $uploadInternetBytes  = [uint64] $upload.MonthlyInternetBytes
    $uploadLinkLocalBytes = [uint64] $upload.MonthlyLinkLocalBytes
    $totalUploadBytes     = [double] ($uploadLanBytes + $uploadGroupBytes + $uploadInternetBytes + $uploadLinkLocalBytes)

    $downloadModeId = [int] $rawConfig.DownloadMode
    $downloadMode   = Resolve-DownloadMode -Value $downloadModeId
    $peeringStatus  = if ($downloadModeId -in @(1, 2, 3)) { 'On' } else { 'Off' }
    $collectedAtLocal = Get-Date
    $collectedAt    = $collectedAtLocal.ToUniversalTime()
    $monthStart     = Get-Date -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
    $configuredGroupId = Get-ObjectPropertyValue -InputObject $policyValues.MDM -Name 'DOGroupID'
    if ($null -eq $configuredGroupId) {
        $configuredGroupId = Get-ObjectPropertyValue -InputObject $policyValues.GroupPolicy -Name 'DOGroupID'
    }
    $configuredGroupIdSource = Get-ObjectPropertyValue -InputObject $policyValues.MDM -Name 'DOGroupIDSource'
    if ($null -eq $configuredGroupIdSource) {
        $configuredGroupIdSource = Get-ObjectPropertyValue -InputObject $policyValues.GroupPolicy -Name 'DOGroupIDSource'
    }
    # An enterprise/self-hosted MCC is set via DOCacheHost; without it, Connected Cache bytes come from an ISP-hosted MCC the DO cloud service assigns automatically.
    $configuredCacheHost = Get-ObjectPropertyValue -InputObject $policyValues.MDM -Name 'DOCacheHost'
    if ($null -eq $configuredCacheHost) {
        $configuredCacheHost = Get-ObjectPropertyValue -InputObject $policyValues.GroupPolicy -Name 'DOCacheHost'
    }
    $transferPayload = @(
        foreach ($transfer in $rawTransfers) {
            $matchingPeerInfo = $rawPeerInfo | Where-Object { $_.FileId -eq $transfer.FileId } | Select-Object -First 1
            ConvertTo-CurrentTransfer -Transfer $transfer -PeerInfo $matchingPeerInfo
        }
    )
    $observedCacheHosts = @($transferPayload | ForEach-Object { $_.CacheHost } | Where-Object { $_ } | Sort-Object -Unique)
    $currentStatusDescription = if ($transferPayload.Count -gt 0) {
        (@($transferPayload.Status | Sort-Object -Unique) -join ', ')
    } else {
        'Idle'
    }

    $result = [pscustomobject]@{
        type            = 'LocalDeliveryOptimizationStatistics'
        timestamp       = $collectedAt.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
        hostname        = $env:COMPUTERNAME
        os_build        = '{0}.{1}' -f $os.BuildNumber, $currentBuild.UBR
        display_version = $currentBuild.DisplayVersion
        period          = 'LocalCalendarMonthToDate'
        period_start    = $monthStart.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
        period_end      = $collectedAt.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
        period_timezone = [TimeZoneInfo]::Local.Id
        period_start_local = $monthStart.ToString('yyyy-MM-dd HH:mm:ss zzz')
        period_end_local = $collectedAtLocal.ToString('yyyy-MM-dd HH:mm:ss zzz')
        identity        = [pscustomobject]@{
            DeviceName      = $env:COMPUTERNAME
            AzureAdJoined   = $entraIdentity.AzureAdJoined
            AzureADDeviceId = $entraIdentity.AzureADDeviceId
            AzureADTenantId = $entraIdentity.AzureADTenantId
            TenantName      = $entraIdentity.TenantName
        }
        configuration   = [pscustomobject]@{
            DownloadMode         = $downloadMode
            DownloadModeId       = $downloadModeId
            DownloadModeProvider = "$($friendly.DownloadModeProvider)"
            PeeringStatus        = $peeringStatus
            PeerConfigured       = ($downloadModeId -in @(1, 2, 3))
            PeerUsedThisMonth    = ($totalPeerBytes -gt 0)
            MCCUsedThisMonth     = ($bytesFromCache -gt 0)
            ConfiguredCacheHost  = $configuredCacheHost
            CacheHostConfigured  = (-not [string]::IsNullOrWhiteSpace([string] $configuredCacheHost))
            ObservedCacheHosts   = $observedCacheHosts
            ConfiguredGroupID    = $configuredGroupId
            ConfiguredGroupIDHash = Get-WUfBGroupIdHash -GroupId $configuredGroupId
            ConfiguredGroupIDSource = $configuredGroupIdSource
            Effective            = [pscustomobject]@{
                BackgroundDownloadLimitBps          = [uint32] $friendly.DownBackLimitBps
                BackgroundDownloadLimitBpsProvider  = "$($friendly.DownBackLimitBpsProvider)"
                BackgroundDownloadLimitPct          = [uint32] $friendly.DownBackLimitPct
                BackgroundDownloadLimitPctProvider  = "$($friendly.DownBackLimitPctProvider)"
                ForegroundDownloadLimitBps          = [uint32] $friendly.DownloadForegroundLimitBps
                ForegroundDownloadLimitBpsProvider  = "$($friendly.DownloadForegroundLimitBpsProvider)"
                ForegroundDownloadLimitPct          = [uint32] $friendly.DownloadForegroundLimitPct
                ForegroundDownloadLimitPctProvider  = "$($friendly.DownloadForegroundLimitPctProvider)"
                MaxUploadRatePct                    = [uint32] $friendly.MaxUploadRatePct
                MaxUploadRateProvider               = "$($friendly.MaxUploadRateProvider)"
                UploadLimitMonthlyGB                = [double] $friendly.UploadLimitMonthlyGB
                UploadLimitMonthlyGBProvider        = "$($friendly.UploadLimitMonthlyGBProvider)"
                BatteryPctToSeed                    = [uint32] $friendly.BatteryPctToSeed
                BatteryPctToSeedProvider            = "$($friendly.BatteryPctToSeedProvider)"
                MinTotalDiskSizeGB                  = [uint32] $friendly.MinTotalDiskSize
                MinTotalDiskSizeProvider            = "$($friendly.MinTotalDiskSizeProvider)"
                MinTotalRAMGB                       = [uint32] $friendly.MinTotalRAM
                MinTotalRAMProvider                 = "$($friendly.MinTotalRAMProvider)"
                VpnPeerCachingAllowed               = [bool] $friendly.VpnPeerCachingAllowed
                VpnPeerCachingAllowedProvider       = "$($friendly.VpnPeerCachingAllowedProvider)"
                VpnKeywords                         = "$($friendly.VpnKeywords)"
                VpnKeywordsProvider                 = "$($friendly.VpnKeywordsProvider)"
                WorkingDirectory                    = "$($friendly.WorkingDirectory)"
                WorkingDirectoryProvider            = "$($friendly.WorkingDirectoryProvider)"
                SetHoursToLimitDownloadBackground   = "$($friendly.SetHoursToLimitDownloadBackground)"
                SetHoursToLimitDownloadForeground   = "$($friendly.SetHoursToLimitDownloadForeground)"
            }
            PolicyValues          = $policyValues
        }
        efficiency      = [pscustomobject]@{
            BandwidthSavingsPct             = $wufbSavingsPct
            BandwidthSavingsStatus          = Get-BandwidthSavingsStatus -BandwidthSavingsPct $wufbSavingsPct -TotalBytes $wufbTotalBytes
            P2PEfficiencyPct                = ConvertTo-Percentage -Numerator $wufbPeerBytes -Denominator $wufbTotalBytes
            ConnectedCacheEfficiencyPct     = ConvertTo-Percentage -Numerator $bytesFromCache -Denominator $wufbTotalBytes
            CDNPercentage                   = ConvertTo-Percentage -Numerator $bytesFromCDN -Denominator $wufbTotalBytes
            AllSourcesBandwidthSavingsPct   = ConvertTo-Percentage -Numerator $localSourceBytes -Denominator $totalDownloadBytes
            AllSourcesP2PEfficiencyPct      = ConvertTo-Percentage -Numerator $totalPeerBytes -Denominator $totalDownloadBytes
            AllSourcesCDNPercentage         = ConvertTo-Percentage -Numerator $bytesFromCDN -Denominator $totalDownloadBytes
        }
        download        = [pscustomobject]@{
            BytesFromHTTP            = $bytesFromHTTP
            BytesFromCDN             = $bytesFromCDN
            BytesFromCache           = $bytesFromCache
            BytesFromPeers           = $bytesFromPeers
            BytesFromGroupPeers      = $bytesFromGroup
            BytesFromIntPeers        = $bytesFromInternet
            BytesFromLinkLocal       = $bytesFromLinkLocal
            WUfBPeerBytes            = [uint64] $wufbPeerBytes
            WUfBLocalSourceBytes     = [uint64] $wufbLocalBytes
            WUfBTotalBytes           = [uint64] $wufbTotalBytes
            TotalPeerBytes           = [uint64] $totalPeerBytes
            LocalSourceBytes         = [uint64] $localSourceBytes
            TotalBytes               = [uint64] $totalDownloadBytes
            AverageForegroundKbps    = [Math]::Round(([double] $download.MonthlyFrRateBps * 8.0) / 1KB, 2)
            AverageBackgroundKbps    = [Math]::Round(([double] $download.MonthlyBkRateBps * 8.0) / 1KB, 2)
        }
        upload          = [pscustomobject]@{
            BytesToLanPeers       = $uploadLanBytes
            BytesToGroupPeers     = $uploadGroupBytes
            BytesToInternetPeers  = $uploadInternetBytes
            BytesToLinkLocal      = $uploadLinkLocalBytes
            TotalBytes            = [uint64] $totalUploadBytes
            UploadLimitReached    = ([int] $upload.MonthlyUploadRestriction -ne 0)
        }
        current         = [pscustomobject]@{
            DOStatusDescription = $currentStatusDescription
            TransferCount       = [int] $transferPayload.Count
            CacheSizeBytes      = [uint64] $current.CacheSizeBytes
            DiskTotalBytes      = [uint64] $current.DiskTotalBytes
            DiskAvailableBytes  = [uint64] $current.DiskAvailableBytes
            PeerCount           = [uint32] $current.PeerInfoCount
            CacheConnections    = [uint32] $current.CacheServerConnections
            CDNConnections      = [uint32] $current.CdnConnections
            LANConnections      = [uint32] $current.LanConnections
            LinkLocalConnections = [uint32] $current.LinkLocalConnections
            GroupConnections    = [uint32] $current.GroupConnections
            InternetConnections = [uint32] $current.InternetConnections
            TotalConnections    = [uint32] ($current.CacheServerConnections + $current.CdnConnections + $current.LanConnections + $current.LinkLocalConnections + $current.GroupConnections + $current.InternetConnections)
            ActiveJobs          = [uint32] ($download.PriorityDownloads + $download.NormalDownloads + $upload.Uploads)
            PendingJobs         = [uint32] ($download.PriorityDownloadsPending + $download.NormalDownloadsPending)
            CpuUsagePct         = [Math]::Round([double] $current.CpuUsagePct, 4)
            MemUsageKB          = [uint64] $current.MemUsageKBytes
        }
        transfers       = $transferPayload
        coverage        = [pscustomobject]@{
            MonthlyHTTPBytes         = 'CapturedInclusiveOfConnectedCache'
            DirectCDNBytes           = 'DerivedAsHTTPMinusConnectedCache'
            MonthlySourceBytes       = 'CapturedLocalCalendarMonth'
            CurrentTransferDetails   = 'CapturedWhenActive'
            CurrentPeerOutcomes      = 'CapturedWhenActive'
            AzureADIdentity          = 'CapturedWhenDeviceJoined'
            ConfiguredGroupID        = 'CapturedWhenExplicitlyConfigured'
            Rolling7DayMetrics       = 'UnavailableLocallyCloudTelemetryOnly'
            Rolling28DayMetrics      = 'UnavailableLocallyCloudTelemetryOnly'
            ContentTypeHistory       = 'UnavailableLocallyCloudTelemetryOnly'
            GeographyAndISP          = 'UnavailableLocallyCloudTelemetryOnly'
            LastCensusSeenTime       = 'UnavailableLocallyCloudTelemetryOnly'
            FleetDeviceCounts        = 'UnavailableLocallyTenantAggregationOnly'
            GlobalDeviceId           = 'UnavailableLocallyMicrosoftInternal'
            HistoricalPeerOutcomes   = 'UnavailableLocallyCloudTelemetryOnly'
        }
    }

    if ($OutputPath) {
        Write-JsonLine -InputObject $result -Path $OutputPath
    }

    if ($AsObject) {
        Write-Output $result
    }
    elseif ($AsJson) {
        if ($Pretty) {
            Write-Output ($result | ConvertTo-Json -Depth 6)
        } else {
            Write-Output ($result | ConvertTo-Json -Depth 6 -Compress)
        }
    }
    else {
        Write-ConsoleReport -Result $result
    }

    exit $EXIT_OK
}
catch {
    $errorResult = [pscustomobject]@{
        timestamp      = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
        hostname       = $env:COMPUTERNAME
        overall_status = 'ERROR'
        error          = $_.Exception.Message
    }

    if ($AsObject) {
        Write-Output $errorResult
    }
    elseif ($AsJson) {
        Write-Output ($errorResult | ConvertTo-Json -Depth 3 -Compress)
    }
    else {
        Write-Host ''
        Write-Host 'Delivery Optimization statistics collection failed.' -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }

    exit $EXIT_ERROR
}