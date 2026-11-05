<#
.SYNOPSIS
    Evaluates AVD session latency, bandwidth, transport protocol, proxy hairpinning,
    and host encoding bottlenecks from Azure Monitor Log Analytics.
.DESCRIPTION
    Queries WVDConnections and WVDConnectionNetworkData (plus optional WVDCheckpoints
    and AMA Perf counters) to produce session-level or session-weighted summary
    evaluations (Session, HostPoolSummary, UserSummary, RegionSummary).
.PARAMETER WorkspaceId
    Mandatory Log Analytics Workspace GUID.
.PARAMETER LookbackHours
    Time window in hours to evaluate. Default: 168 (7 days).
.PARAMETER HostPoolResourceId
    Optional ARM Resource ID(s) of AVD Host Pools to filter.
.PARAMETER UserPrincipalName
    Optional user UPN filter.
.PARAMETER MinSampleCount
    Minimum 2-minute network samples required to classify a session. Default: 3.
.PARAMETER CorporateEgressCidrs
    Optional IPv4/IPv6 CIDR ranges of corporate SWG/VPN egress IPs.
.PARAMETER ExpectedGatewayRegions
    Optional expected AVD Gateway region(s) (e.g. 'westeurope').
.PARAMETER IncludeLogonDuration
    Optional switch to join WVDCheckpoints logon duration per session.
.PARAMETER IncludeGuestPerfCounters
    Optional switch to join AMA Perf counters (RemoteFX Graphics & User Input Delay).
.PARAMETER OutputView
    Session (default), UserSummary, HostPoolSummary, or RegionSummary.
.PARAMETER ManagedIdentityClientId
    Optional Client ID of a User-Assigned Managed Identity.
.PARAMETER LogAnalyticsAccessToken
    Optional pre-acquired Log Analytics bearer token.
.PARAMETER AuthMode
    Auto (default), ManagedIdentity, or AzContext.
.OUTPUTS
    PSCustomObject[]
.EXAMPLE
    .\Get-AvdConnectionQuality.ps1 -WorkspaceId '11111111-2222-3333-4444-555555555555' -OutputView HostPoolSummary
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]  [string] $WorkspaceId,
    [Parameter(Mandatory = $false)] [ValidateRange(1, 2160)] [int] $LookbackHours = 168,
    [Parameter(Mandatory = $false)] [string[]] $HostPoolResourceId,
    [Parameter(Mandatory = $false)] [string] $UserPrincipalName,
    [Parameter(Mandatory = $false)] [ValidateRange(1, 100)] [int] $MinSampleCount = 3,
    [Parameter(Mandatory = $false)] [string[]] $CorporateEgressCidrs,
    [Parameter(Mandatory = $false)] [string[]] $ExpectedGatewayRegions,
    [Parameter(Mandatory = $false)] [switch] $IncludeLogonDuration,
    [Parameter(Mandatory = $false)] [switch] $IncludeGuestPerfCounters,
    [Parameter(Mandatory = $false)] [ValidateSet('Session', 'UserSummary', 'HostPoolSummary', 'RegionSummary')] [string] $OutputView = 'Session',
    [Parameter(Mandatory = $false)] [string] $ManagedIdentityClientId,
    [Parameter(Mandatory = $false)] [string] $LogAnalyticsAccessToken,
    [Parameter(Mandatory = $false)] [ValidateSet('Auto', 'ManagedIdentity', 'AzContext')] [string] $AuthMode = 'Auto'
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

function Get-AvdQualityProp {
    <#
    .SYNOPSIS
        Reads a property or dictionary key safely under Set-StrictMode 2.0.
    .DESCRIPTION
        Returns the value of Name from InputObject or $null when missing.
    .PARAMETER InputObject
        Object or IDictionary to inspect.
    .PARAMETER Name
        Property or key name.
    .OUTPUTS
        System.Object
    #>
    param([Parameter(Mandatory = $false)] $InputObject, [Parameter(Mandatory = $true)] [string] $Name)
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Collections.IDictionary]) {
        foreach ($k in $InputObject.Keys) {
            if ([string]::Equals([string]$k, $Name, [System.StringComparison]::OrdinalIgnoreCase)) { return $InputObject[$k] }
        }
        return $null
    }
    $p = $InputObject.PSObject.Properties[$Name]
    if ($null -ne $p) { return $p.Value }
    return $null
}

function ConvertTo-AvdQualityKqlLiteral {
    <#
    .SYNOPSIS
        Formats a string as a safe Kusto (KQL) single-quoted literal.
    .DESCRIPTION
        Rejects newlines/nulls and escapes backslashes and single quotes for KQL.
    .PARAMETER Value
        String value to encode.
    .OUTPUTS
        System.String
    #>
    param([Parameter(Mandatory = $true)] [AllowEmptyString()] [string] $Value)
    if ($Value -match "[\r\n`0]") { throw "Refusing to interpolate control characters into KQL." }
    return "'$($Value.Replace('\', '\\').Replace("'", "\'"))'"
}

function ConvertTo-AvdQualityDoubleOrNull {
    <#
    .SYNOPSIS
        Converts a Log Analytics numeric cell to a rounded double or $null.
    .DESCRIPTION
        Parses Value using InvariantCulture and returns $null for empty/NaN/Infinity.
    .PARAMETER Value
        Raw cell value.
    .OUTPUTS
        System.Nullable[double]
    #>
    param([Parameter(Mandatory = $false)] $Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    $d = [double]0
    if ([double]::TryParse(([string]$Value).Trim(), [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) {
        if (-not [double]::IsNaN($d) -and -not [double]::IsInfinity($d)) { return [Math]::Round($d, 2) }
    }
    return $null
}

function Invoke-AvdQualityHttpRequest {
    <#
    .SYNOPSIS
        Low-level HTTP transport wrapper around Invoke-RestMethod.
    .DESCRIPTION
        Executes an HTTP request with UTF-8 JSON body encoding and a 210s timeout.
    .PARAMETER Method
        HTTP method.
    .PARAMETER Uri
        Request URI.
    .PARAMETER Headers
        Request headers hashtable.
    .PARAMETER Body
        Optional JSON request body.
    .OUTPUTS
        System.Object
    #>
    param(
        [Parameter(Mandatory = $true)]  [string] $Method,
        [Parameter(Mandatory = $true)]  [string] $Uri,
        [Parameter(Mandatory = $true)]  [hashtable] $Headers,
        [Parameter(Mandatory = $false)] [string] $Body
    )
    $p = @{ Method = $Method; Uri = $Uri; Headers = $Headers; UseBasicParsing = $true; TimeoutSec = 210 }
    if (-not [string]::IsNullOrEmpty($Body)) {
        $p['Body']        = [System.Text.Encoding]::UTF8.GetBytes($Body)
        $p['ContentType'] = 'application/json; charset=utf-8'
    }
    return (Invoke-RestMethod @p)
}

function Get-AvdQualityAccessToken {
    <#
    .SYNOPSIS
        Acquires an OAuth2 bearer token for Azure Monitor Log Analytics.
    .DESCRIPTION
        Supports Managed Identity (IDENTITY_ENDPOINT / IMDS) and Az.Accounts (Get-AzAccessToken).
    .PARAMETER ResourceUri
        Target audience URI.
    .PARAMETER SelectedAuthMode
        Auto, ManagedIdentity, or AzContext.
    .PARAMETER ClientId
        Optional User-Assigned Managed Identity Client ID.
    .OUTPUTS
        System.String
    #>
    param(
        [Parameter(Mandatory = $true)]  [string] $ResourceUri,
        [Parameter(Mandatory = $true)]  [string] $SelectedAuthMode,
        [Parameter(Mandatory = $false)] [string] $ClientId
    )
    $res = $ResourceUri.TrimEnd('/')
    $useMi = ($SelectedAuthMode -eq 'ManagedIdentity') -or ($SelectedAuthMode -eq 'Auto' -and -not [string]::IsNullOrWhiteSpace($env:IDENTITY_ENDPOINT))
    if ($useMi) {
        $ep = if (-not [string]::IsNullOrWhiteSpace($env:IDENTITY_ENDPOINT)) { $env:IDENTITY_ENDPOINT } else { 'http://169.254.169.254/metadata/identity/oauth2/token' }
        $ver = if (-not [string]::IsNullOrWhiteSpace($env:IDENTITY_ENDPOINT)) { '2019-08-01' } else { '2018-02-01' }
        $q = "api-version=$ver&resource=$([uri]::EscapeDataString($res))"
        if (-not [string]::IsNullOrWhiteSpace($ClientId)) { $q += "&client_id=$([uri]::EscapeDataString($ClientId.Trim()))" }
        $hdr = @{ Metadata = 'true' }
        if (-not [string]::IsNullOrWhiteSpace($env:IDENTITY_HEADER)) { $hdr['X-IDENTITY-HEADER'] = $env:IDENTITY_HEADER }
        $resp = Invoke-AvdQualityHttpRequest -Method Get -Uri ("{0}?{1}" -f $ep, $q) -Headers $hdr
        return [string](Get-AvdQualityProp -InputObject $resp -Name 'access_token')
    }
    $tokObj = Get-AzAccessToken -ResourceUrl $res -ErrorAction Stop
    $rawTok = Get-AvdQualityProp -InputObject $tokObj -Name 'Token'
    if ($rawTok -is [System.Security.SecureString]) {
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($rawTok)
        try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    }
    return [string]$rawTok
}

function Invoke-AvdQualityRest {
    <#
    .SYNOPSIS
        Invokes a Log Analytics REST request with HTTP 429 retry and 401 token refresh.
    .DESCRIPTION
        Extracts status codes and Retry-After from WebException.Response or mocked Exception.Data.
    .PARAMETER Method
        HTTP method (Get, Post).
    .PARAMETER Uri
        Full request URI.
    .PARAMETER Headers
        Request headers hashtable.
    .PARAMETER Body
        Optional JSON request body.
    .OUTPUTS
        System.Object
    #>
    param(
        [Parameter(Mandatory = $true)]  [ValidateSet('Get', 'Post')] [string] $Method,
        [Parameter(Mandatory = $true)]  [string] $Uri,
        [Parameter(Mandatory = $true)]  [hashtable] $Headers,
        [Parameter(Mandatory = $false)] [string] $Body
    )
    $attempt = 0
    $refreshed = $false
    while ($true) {
        $attempt++
        try {
            return (Invoke-AvdQualityHttpRequest -Method $Method -Uri $Uri -Headers $Headers -Body $Body)
        } catch {
            $status = 0
            $retrySec = 0
            $resp = Get-AvdQualityProp -InputObject $_.Exception -Name 'Response'
            if ($null -ne $resp) {
                try { $status = [int]$resp.StatusCode } catch { $status = 0 }
                try {
                    $ra = $resp.Headers['Retry-After']
                    if (-not [string]::IsNullOrWhiteSpace([string]$ra)) { [void][int]::TryParse([string]$ra, [ref]$retrySec) }
                } catch { $retrySec = 0 }
            }
            if ($status -eq 0 -and $null -ne $_.Exception.Data -and $_.Exception.Data.Contains('HttpStatusCode')) {
                $status = [int]$_.Exception.Data['HttpStatusCode']
            }
            if ($retrySec -le 0 -and $null -ne $_.Exception.Data -and $_.Exception.Data.Contains('RetryAfterSeconds')) {
                $retrySec = [int]$_.Exception.Data['RetryAfterSeconds']
            }
            if ($status -eq 401 -and -not $refreshed) {
                $refreshed = $true
                $amVar = Get-Variable -Name 'AuthMode' -ErrorAction SilentlyContinue
                $miVar = Get-Variable -Name 'ManagedIdentityClientId' -ErrorAction SilentlyContinue
                if ($null -ne $amVar) {
                    try {
                        $miClient = if ($null -ne $miVar) { [string]$miVar.Value } else { '' }
                        $newTok = Get-AvdQualityAccessToken -ResourceUri 'https://api.loganalytics.io/' -SelectedAuthMode ([string]$amVar.Value) -ClientId $miClient
                        if (-not [string]::IsNullOrWhiteSpace($newTok)) { $Headers['Authorization'] = "Bearer $newTok"; continue }
                    } catch { }
                }
            }
            if (($status -in @(429, 502, 503, 504)) -and $attempt -lt 4) {
                $wait = if ($retrySec -gt 0) { [Math]::Min($retrySec, 300) } else { [int][Math]::Pow(2, $attempt) }
                if ($wait -gt 0) { Start-Sleep -Seconds $wait }
                continue
            }
            throw
        }
    }
}

function ConvertTo-AvdTransportCategory {
    <#
    .SYNOPSIS
        Normalizes WVDConnections TransportType or legacy UdpUse into a canonical category.
    .DESCRIPTION
        Maps into DirectUdpShortpath, RelayedUdpTurn, TcpWebsocket, UdpSubtypeUnknown, or Unknown.
    .PARAMETER RawTransportType
        Value of TransportType column.
    .PARAMETER RawUdpUse
        Value of legacy UdpUse column.
    .OUTPUTS
        System.String
    #>
    param([Parameter(Mandatory = $false)] [string] $RawTransportType, [Parameter(Mandatory = $false)] [string] $RawUdpUse)
    if (-not [string]::IsNullOrWhiteSpace($RawTransportType)) {
        $t = $RawTransportType.Trim().ToLowerInvariant()
        if ($t -in @('shortpath', 'udp-shortpath', 'directudp')) { return 'DirectUdpShortpath' }
        if ($t -in @('turn', 'relayedudp', 'relay'))             { return 'RelayedUdpTurn' }
        if ($t -in @('websocket', 'tcp', 'https'))               { return 'TcpWebsocket' }
        if ($t -eq 'udp')                                        { return 'UdpSubtypeUnknown' }
    }
    if (-not [string]::IsNullOrWhiteSpace($RawUdpUse)) {
        $u = $RawUdpUse.Trim().ToLowerInvariant()
        if ($u -eq 'shortpath')                              { return 'DirectUdpShortpath' }
        if ($u -eq 'turn')                                   { return 'RelayedUdpTurn' }
        if ($u -in @('1', 'true', 'yes', 'udp'))             { return 'UdpSubtypeUnknown' }
        if ($u -in @('0', 'false', 'no', 'tcp', 'websocket')) { return 'TcpWebsocket' }
    }
    return 'Unknown'
}

function Test-AvdIpInCidr {
    <#
    .SYNOPSIS
        Tests whether an IPv4 or IPv6 address falls within any of the supplied CIDR prefixes.
    .DESCRIPTION
        Performs byte-wise comparison supporting /0../32 (IPv4), /0../128 (IPv6), and IPv4-mapped IPv6.
    .PARAMETER IpString
        Client IP address string.
    .PARAMETER CidrList
        Array of CIDR strings.
    .OUTPUTS
        System.Boolean
    #>
    param([Parameter(Mandatory = $false)] [string] $IpString, [Parameter(Mandatory = $false)] [string[]] $CidrList)
    if ($null -eq $CidrList -or $CidrList.Count -eq 0) { return $false }
    $parsed = New-Object 'System.Collections.Generic.List[object]'
    foreach ($cidr in $CidrList) {
        if ([string]::IsNullOrWhiteSpace($cidr)) { continue }
        $parts = $cidr.Trim() -split '/'
        $netObj = $null; $pLen = -1
        if ($parts.Count -ne 2 -or -not [System.Net.IPAddress]::TryParse($parts[0].Trim(), [ref]$netObj) -or -not [int]::TryParse($parts[1].Trim(), [ref]$pLen)) {
            throw "Invalid CIDR '$cidr'."
        }
        if ($netObj.IsIPv4MappedToIPv6) { $netObj = $netObj.MapToIPv4() }
        $nb = $netObj.GetAddressBytes()
        if ($pLen -lt 0 -or $pLen -gt ($nb.Length * 8)) { throw "Invalid CIDR prefix '/$pLen' in '$cidr'." }
        [void]$parsed.Add(@{ Bytes = $nb; Prefix = $pLen })
    }
    if ([string]::IsNullOrWhiteSpace($IpString)) { return $false }
    $clean = $IpString.Trim()
    if ($clean -match '^(\d{1,3}(?:\.\d{1,3}){3}):\d+$') { $clean = $Matches[1] }
    $ipObj = $null
    if (-not [System.Net.IPAddress]::TryParse($clean, [ref]$ipObj)) { return $false }
    if ($ipObj.IsIPv4MappedToIPv6) { $ipObj = $ipObj.MapToIPv4() }
    $ib = $ipObj.GetAddressBytes()
    foreach ($e in $parsed) {
        if ($ib.Length -ne $e.Bytes.Length) { continue }
        $full = [Math]::Floor($e.Prefix / 8); $rem = $e.Prefix % 8; $ok = $true
        for ($i = 0; $i -lt $full; $i++) { if ($ib[$i] -ne $e.Bytes[$i]) { $ok = $false; break } }
        if ($ok -and $rem -gt 0) {
            $mask = [byte]((0xFF -shl (8 - $rem)) -band 0xFF)
            if (($ib[$full] -band $mask) -ne ($e.Bytes[$full] -band $mask)) { $ok = $false }
        }
        if ($ok) { return $true }
    }
    return $false
}

function Get-AvdNumericPercentile {
    <#
    .SYNOPSIS
        Computes a linearly interpolated percentile over an array of doubles.
    .DESCRIPTION
        Used to compute session-weighted P50/P90 summaries across per-session metrics.
    .PARAMETER Values
        Numeric array.
    .PARAMETER Percentile
        Target percentile (0..100).
    .OUTPUTS
        System.Nullable[double]
    #>
    param([Parameter(Mandatory = $false)] [double[]] $Values, [Parameter(Mandatory = $true)] [ValidateRange(0, 100)] [double] $Percentile)
    if ($null -eq $Values -or $Values.Count -eq 0) { return $null }
    $s = @($Values | Sort-Object)
    if ($s.Count -eq 1) { return [Math]::Round([double]$s[0], 2) }
    $r = ($Percentile / 100.0) * ($s.Count - 1)
    $lo = [int][Math]::Floor($r); $hi = [int][Math]::Ceiling($r)
    if ($lo -eq $hi) { return [Math]::Round([double]$s[$lo], 2) }
    return [Math]::Round(([double]$s[$lo] * (1.0 - ($r - $lo))) + ([double]$s[$hi] * ($r - $lo)), 2)
}

# --- MAIN EXECUTION ---
$poolIdPattern = '^/subscriptions/[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}/resourceGroups/[A-Za-z0-9_.()\-]{1,90}/providers/Microsoft\.DesktopVirtualization/hostPools/[A-Za-z0-9_.\-]{1,64}$'
$cleanWsId = $WorkspaceId.Trim().Trim('{', '}').ToLowerInvariant()
if ($cleanWsId -notmatch '^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$') { throw "Invalid -WorkspaceId '$WorkspaceId'." }
if ($null -ne $CorporateEgressCidrs -and $CorporateEgressCidrs.Count -gt 0) { $null = Test-AvdIpInCidr -IpString '' -CidrList $CorporateEgressCidrs }

$poolClause = ''
if ($null -ne $HostPoolResourceId -and $HostPoolResourceId.Count -gt 0) {
    $lits = foreach ($hp in $HostPoolResourceId) {
        if (-not [string]::IsNullOrWhiteSpace($hp)) {
            $c = $hp.Trim()
            if ($c -notmatch $poolIdPattern) { throw "Invalid -HostPoolResourceId '$c'." }
            ConvertTo-AvdQualityKqlLiteral -Value $c.ToLowerInvariant()
        }
    }
    if (@($lits).Count -gt 0) { $poolClause = "| where tolower(_ResourceId) in ($($lits -join ', '))" }
}
$userClause = if (-not [string]::IsNullOrWhiteSpace($UserPrincipalName)) { "| where tolower(UserName) == $(ConvertTo-AvdQualityKqlLiteral -Value $UserPrincipalName.Trim().ToLowerInvariant())" } else { '' }

$nowUtc       = [DateTimeOffset]::UtcNow
$endIso       = $nowUtc.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ss.fffZ', [Globalization.CultureInfo]::InvariantCulture)
$startIso     = $nowUtc.AddHours(-$LookbackHours).UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ss.fffZ', [Globalization.CultureInfo]::InvariantCulture)
$connStartIso = $nowUtc.AddHours(-($LookbackHours + 24)).UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ss.fffZ', [Globalization.CultureInfo]::InvariantCulture)

$kql = @"
let winStart = datetime($startIso);
let winEnd = datetime($endIso);
let connLookback = datetime($connStartIso);
let connections = WVDConnections
    | where TimeGenerated between (connLookback .. winEnd)
    $poolClause
    $userClause
    | where State in ('Connected', 'Completed')
    | extend RawTransport = tostring(column_ifexists('TransportType', '')), RawUdp = tostring(column_ifexists('UdpUse', '')), VmId = tolower(tostring(column_ifexists('SessionHostAzureVmId', '')))
    | summarize
        ConnectTimeUtc = minif(TimeGenerated, State == 'Connected'),
        EndTimeUtc = maxif(TimeGenerated, State == 'Completed'),
        UserName = take_anyif(UserName, isnotempty(UserName)),
        SessionHostName = take_anyif(SessionHostName, isnotempty(SessionHostName)),
        SessionHostAzureVmId = take_anyif(VmId, isnotempty(VmId)),
        ClientSideIPAddress = take_anyif(ClientSideIPAddress, isnotempty(ClientSideIPAddress)),
        ClientOS = take_anyif(ClientOS, isnotempty(ClientOS)),
        GatewayRegion = take_anyif(GatewayRegion, isnotempty(GatewayRegion)),
        TransportType = take_anyif(RawTransport, isnotempty(RawTransport)),
        UdpUse = take_anyif(RawUdp, isnotempty(RawUdp))
      by CorrelationId, _ResourceId;
let netSummary = WVDConnectionNetworkData
    | where TimeGenerated between (winStart .. winEnd) and isnotnull(EstRoundTripTimeInMs)
    | summarize
        SampleCount = count(),
        SessionP50RttMs = percentile(EstRoundTripTimeInMs, 50),
        SessionP95RttMs = percentile(EstRoundTripTimeInMs, 95),
        SessionMaxRttMs = max(EstRoundTripTimeInMs),
        SessionP10BandwidthKBps = percentile(EstAvailableBandwidthKBps, 10),
        SessionAvgBandwidthKBps = avg(EstAvailableBandwidthKBps)
      by CorrelationId;
connections
| join kind=inner netSummary on CorrelationId
| extend HostKey = toupper(tostring(split(SessionHostName, '.')[0]))
"@

if ($IncludeLogonDuration) {
    $kql += @"

| join kind=leftouter (
    union isfuzzy=true (datatable(CorrelationId:string, Name:string, TimeGenerated:datetime)[]), (WVDCheckpoints | where TimeGenerated between (connLookback .. winEnd) | project CorrelationId, Name, TimeGenerated)
    | where Name in ('LoadBalancedNewConnection', 'RdpStackConnected', 'Logon', 'OnConnected', 'ShellReady', 'LogonCompleted')
    | summarize LogonDurationSec = datetime_diff('second', max(TimeGenerated), min(TimeGenerated)) by CorrelationId
) on CorrelationId
"@
}
if ($IncludeGuestPerfCounters) {
    $kql += @"

| join kind=leftouter (
    union isfuzzy=true (datatable(Computer:string, ObjectName:string, CounterName:string, CounterValue:real, TimeGenerated:datetime)[]), (Perf | where TimeGenerated between (winStart .. winEnd) and ObjectName in ('RemoteFX Graphics', 'User Input Delay per Session') | project Computer, ObjectName, CounterName, CounterValue, TimeGenerated)
    | extend HostKey = toupper(tostring(split(Computer, '.')[0]))
    | summarize
        HostP95EncodingTimeMs = percentileif(CounterValue, 95, ObjectName == 'RemoteFX Graphics' and CounterName == 'Average Encoding Time'),
        HostP95InputDelayMs = percentileif(CounterValue, 95, ObjectName == 'User Input Delay per Session' and CounterName == 'Max Input Delay')
      by HostKey
) on HostKey
"@
}

$laTok = if (-not [string]::IsNullOrWhiteSpace($LogAnalyticsAccessToken)) { $LogAnalyticsAccessToken.Trim() } else { Get-AvdQualityAccessToken -ResourceUri 'https://api.loganalytics.io/' -SelectedAuthMode $AuthMode -ClientId $ManagedIdentityClientId }
$resp = Invoke-AvdQualityRest -Method Post -Uri ("https://api.loganalytics.io/v1/workspaces/{0}/query" -f $cleanWsId) -Headers @{ Authorization = "Bearer $laTok"; Prefer = 'response-v1=true,wait=180' } -Body (@{ query = $kql; timespan = ("PT{0}H" -f ($LookbackHours + 24)) } | ConvertTo-Json -Compress)
$err = Get-AvdQualityProp -InputObject $resp -Name 'error'
if ($null -eq $err) { $err = Get-AvdQualityProp -InputObject $resp -Name 'partialError' }
if ($null -ne $err) { throw "Log Analytics query returned partial/error state (code='$(Get-AvdQualityProp -InputObject $err -Name 'code')'): $(Get-AvdQualityProp -InputObject $err -Name 'message')" }

$tables = @(Get-AvdQualityProp -InputObject $resp -Name 'tables')
if ($tables.Count -eq 0 -or $null -eq $tables[0]) { return @() }
$tbl = $tables[0]
foreach ($t in $tables) { if ([string](Get-AvdQualityProp -InputObject $t -Name 'name') -eq 'PrimaryResult') { $tbl = $t; break } }
$cols = @(Get-AvdQualityProp -InputObject $tbl -Name 'columns')
$rows = @(Get-AvdQualityProp -InputObject $tbl -Name 'rows')

$expectedRegions = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
if ($null -ne $ExpectedGatewayRegions) { foreach ($r in $ExpectedGatewayRegions) { if (-not [string]::IsNullOrWhiteSpace($r)) { [void]$expectedRegions.Add($r.Trim()) } } }

$sessions = New-Object 'System.Collections.Generic.List[object]'
foreach ($row in $rows) {
    if ($null -eq $row) { continue }
    $arr = @($row); $m = @{}
    for ($i = 0; $i -lt $cols.Count; $i++) {
        $cn = [string](Get-AvdQualityProp -InputObject $cols[$i] -Name 'name')
        if (-not [string]::IsNullOrWhiteSpace($cn) -and -not $m.ContainsKey($cn)) { $m[$cn] = if ($i -lt $arr.Count) { $arr[$i] } else { $null } }
    }
    $resId    = [string](Get-AvdQualityProp -InputObject $m -Name '_ResourceId')
    $normPool = if (-not [string]::IsNullOrWhiteSpace($resId)) { $resId.Trim().ToLowerInvariant() } else { 'unknown-pool-id' }
    $poolName = if (-not [string]::IsNullOrWhiteSpace($resId)) { ($resId -split '/')[-1] } else { 'UnknownPool' }
    $shName   = [string](Get-AvdQualityProp -InputObject $m -Name 'SessionHostName')
    $clientIp = [string](Get-AvdQualityProp -InputObject $m -Name 'ClientSideIPAddress')
    $gwRegion = [string](Get-AvdQualityProp -InputObject $m -Name 'GatewayRegion')
    $tCat     = ConvertTo-AvdTransportCategory -RawTransportType ([string](Get-AvdQualityProp -InputObject $m -Name 'TransportType')) -RawUdpUse ([string](Get-AvdQualityProp -InputObject $m -Name 'UdpUse'))
    $sCnt     = [int](Get-AvdQualityProp -InputObject $m -Name 'SampleCount')
    $p50      = ConvertTo-AvdQualityDoubleOrNull -Value (Get-AvdQualityProp -InputObject $m -Name 'SessionP50RttMs')
    $p95      = ConvertTo-AvdQualityDoubleOrNull -Value (Get-AvdQualityProp -InputObject $m -Name 'SessionP95RttMs')
    $spread   = if ($null -ne $p95 -and $null -ne $p50) { [Math]::Round(($p95 - $p50), 2) } else { $null }
    $encMs    = ConvertTo-AvdQualityDoubleOrNull -Value (Get-AvdQualityProp -InputObject $m -Name 'HostP95EncodingTimeMs')
    $inDelay  = ConvertTo-AvdQualityDoubleOrNull -Value (Get-AvdQualityProp -InputObject $m -Name 'HostP95InputDelayMs')

    $qual     = ($sCnt -ge $MinSampleCount -and $null -ne $p95)
    $isEgress = Test-AvdIpInCidr -IpString $clientIp -CidrList $CorporateEgressCidrs
    $isGwMis  = ($expectedRegions.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($gwRegion) -and -not $expectedRegions.Contains($gwRegion.Trim()))
    $isTcp    = ($tCat -eq 'TcpWebsocket')
    $hairpin  = ($qual -and $p95 -ge 150.0 -and (($isEgress -and $tCat -in @('TcpWebsocket', 'RelayedUdpTurn')) -or $isGwMis))
    $band     = if (-not $qual) { 'InsufficientSamples' } elseif ($p95 -lt 50.0 -and ($null -eq $spread -or $spread -lt 25.0)) { 'Excellent' } elseif ($p95 -lt 100.0 -and ($null -eq $spread -or $spread -lt 50.0)) { 'Good' } elseif ($p95 -lt 150.0) { 'Fair' } elseif ($p95 -lt 200.0) { 'Degraded' } else { 'Poor' }

    [void]$sessions.Add([PSCustomObject]@{
        HostPoolName             = $poolName
        HostPoolResourceId       = $normPool
        CorrelationId            = [string](Get-AvdQualityProp -InputObject $m -Name 'CorrelationId')
        UserName                 = [string](Get-AvdQualityProp -InputObject $m -Name 'UserName')
        SessionHostName          = $shName
        ShortHostName            = if (-not [string]::IsNullOrWhiteSpace($shName)) { (($shName -split '\.')[0]).Trim().ToUpperInvariant() } else { '' }
        SessionHostAzureVmId     = [string](Get-AvdQualityProp -InputObject $m -Name 'SessionHostAzureVmId')
        ClientSideIPAddress      = $clientIp
        ClientOS                 = [string](Get-AvdQualityProp -InputObject $m -Name 'ClientOS')
        GatewayRegion            = $gwRegion
        TransportCategory        = $tCat
        SampleCount              = $sCnt
        SessionP50RttMs          = $p50
        SessionP95RttMs          = $p95
        SessionMaxRttMs          = ConvertTo-AvdQualityDoubleOrNull -Value (Get-AvdQualityProp -InputObject $m -Name 'SessionMaxRttMs')
        SessionRttSpreadMs       = $spread
        SessionP10BandwidthKBps  = ConvertTo-AvdQualityDoubleOrNull -Value (Get-AvdQualityProp -InputObject $m -Name 'SessionP10BandwidthKBps')
        SessionAvgBandwidthKBps  = ConvertTo-AvdQualityDoubleOrNull -Value (Get-AvdQualityProp -InputObject $m -Name 'SessionAvgBandwidthKBps')
        LogonDurationSec         = ConvertTo-AvdQualityDoubleOrNull -Value (Get-AvdQualityProp -InputObject $m -Name 'LogonDurationSec')
        HostP95EncodingTimeMs    = $encMs
        HostP95InputDelayMs      = $inDelay
        IsCorporateEgressIp      = $isEgress
        IsGatewayRegionMismatch  = $isGwMis
        IsTcpFallback            = $isTcp
        SuspectedProxyHairpin    = $hairpin
        HostEncodingBottleneck   = ($qual -and $null -ne $encMs -and $encMs -gt 33.0)
        HostInputDelayBottleneck = ($qual -and $null -ne $inDelay -and $inDelay -gt 100.0)
        QualityBand              = $band
    })
}

if ($OutputView -eq 'Session') { return $sessions.ToArray() }

$gProp = switch ($OutputView) { 'UserSummary' { 'UserName' } 'HostPoolSummary' { 'HostPoolResourceId' } 'RegionSummary' { 'GatewayRegion' } }
$summaries = New-Object 'System.Collections.Generic.List[object]'
foreach ($g in ($sessions | Group-Object -Property $gProp)) {
    $all = @($g.Group)
    $val = @($all | Where-Object { $_.QualityBand -ne 'InsufficientSamples' -and $null -ne $_.SessionP95RttMs })
    $p50s = @($val | Where-Object { $null -ne $_.SessionP50RttMs } | ForEach-Object { [double]$_.SessionP50RttMs })
    $p95s = @($val | ForEach-Object { [double]$_.SessionP95RttMs })
    $bws  = @($val | Where-Object { $null -ne $_.SessionP10BandwidthKBps } | ForEach-Object { [double]$_.SessionP10BandwidthKBps })
    $qCnt = $val.Count
    [void]$summaries.Add([PSCustomObject]@{
        GroupType                    = $OutputView
        GroupKey                     = if ([string]::IsNullOrWhiteSpace($g.Name)) { 'Unknown' } else { $g.Name }
        HostPoolName                 = if ($all.Count -gt 0) { $all[0].HostPoolName } else { 'UnknownPool' }
        TotalSessions                = $all.Count
        QualifiedSessions            = $qCnt
        AvgOfSessionP50RttMs         = if ($p50s.Count -gt 0) { [Math]::Round((($p50s | Measure-Object -Average).Average), 2) } else { $null }
        MedianOfSessionP95RttMs      = Get-AvdNumericPercentile -Values $p95s -Percentile 50
        P90OfSessionP95RttMs         = Get-AvdNumericPercentile -Values $p95s -Percentile 90
        MedianOfSessionP10BwKBps     = Get-AvdNumericPercentile -Values $bws -Percentile 50
        DirectUdpShortpathPct        = if ($qCnt -gt 0) { [Math]::Round((@($val | Where-Object { $_.TransportCategory -eq 'DirectUdpShortpath' }).Count * 100.0) / $qCnt, 1) } else { 0.0 }
        RelayedUdpTurnPct            = if ($qCnt -gt 0) { [Math]::Round((@($val | Where-Object { $_.TransportCategory -eq 'RelayedUdpTurn' }).Count * 100.0) / $qCnt, 1) } else { 0.0 }
        TcpWebsocketPct              = if ($qCnt -gt 0) { [Math]::Round((@($val | Where-Object { $_.TransportCategory -eq 'TcpWebsocket' }).Count * 100.0) / $qCnt, 1) } else { 0.0 }
        DegradedOrPoorSessionCount   = @($val | Where-Object { $_.QualityBand -in @('Degraded', 'Poor') }).Count
        SuspectedHairpinSessionCount = @($val | Where-Object { $_.SuspectedProxyHairpin }).Count
        HostBottleneckSessionCount   = @($val | Where-Object { $_.HostEncodingBottleneck -or $_.HostInputDelayBottleneck }).Count
    })
}
return $summaries.ToArray()
