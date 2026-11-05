<#
.SYNOPSIS
    Detects, quarantines, snapshots, and decommissions AVD Personal Host Pool VMs
    unused for N days (default 30).
.DESCRIPTION
    Evaluates AVD Personal Host Pools for session hosts with no user connections
    over -InactiveDays (default 30) and runs a fail-closed two-phase lifecycle:
    1. Telemetry Continuity Gate: Verifies WVDConnections has data spanning back at
       least InactiveDays, recent events within 7 days, and at least MinTelemetryDays
       active daily buckets over the last 30 days (failing closed on PartialError).
    2. Activity & Safety Guards: Correlates WVDConnections by Compute vmId
       (SessionHostAzureVmId) and short hostname. Excludes unassigned warm hosts,
       hosts with live sessions (sessions > 0), hosts with Started-only failed
       connections, VMs created or assigned within InactiveDays, and hosts with
       AvdLegalHold enabled. Restores allowNewSession = true if a quarantined host
       shows new activity or user reassignment.
    3. Phase 1 (Active -> Quarantined): Drains the session host (allowNewSession = false),
       deallocates the VM, and merges AvdLifecycleState = 'Quarantined' and
       AvdDecommissionAfterUtc (+QuarantineGraceDays, default 7d) onto the VM tags.
    4. Phase 2 (Quarantined -> Decommissioned, -Mode Execute): After QuarantineGraceDays,
       confirms the VM is still deallocated, creates incremental snapshots of the OS
       disk and any data disks, removes the Intune managedDevice, disables (default,
       preserving BitLocker keys) or deletes the Entra ID device (skipping Hybrid
       ServerAD devices), deletes the VM and exclusively owned NICs/disks, and removes
       the AVD sessionHost registration.
.PARAMETER SubscriptionId
    Optional Azure Subscription ID used with -ResourceGroupName and -HostPoolName.
.PARAMETER ResourceGroupName
    Resource Group containing the target AVD Personal Host Pool.
.PARAMETER HostPoolName
    Name of the target AVD Personal Host Pool.
.PARAMETER HostPoolResourceId
    Full ARM Resource ID of one or more AVD Personal Host Pools.
.PARAMETER WorkspaceId
    Log Analytics Workspace GUID receiving WVDConnections diagnostics.
.PARAMETER Mode
    Audit (default, read-only), QuarantineOnly (Phase 1 only), or Execute (Phase 1 + Phase 2).
.PARAMETER InactiveDays
    Days without user connections before quarantine eligibility. Default: 30.
.PARAMETER QuarantineGraceDays
    Days a VM stays deallocated in Quarantined state before Phase 2 deletion. Default: 7.
.PARAMETER SnapshotRetentionDays
    Days stamped into DeleteAfterUtc on recovery disk snapshots. Default: 30.
.PARAMETER SnapshotResourceGroupId
    Optional ARM Resource Group ID for snapshots. Defaults to the VM's Compute RG.
.PARAMETER MaxDecommissionPercent
    Max percentage of assigned hosts in the pool allowed in quarantine/decommission per window. Default: 15.
.PARAMETER MaxDecommissionCount
    Max absolute number of hosts in the pool allowed in quarantine/decommission per window. Default: 25.
.PARAMETER MinTelemetryDays
    Minimum distinct daily buckets (out of 30) required in WVDConnections. Default: 20.
.PARAMETER EntraCleanupAction
    Disable (default, preserves BitLocker keys), Delete, or Skip.
.PARAMETER ForceCircuitBreakerOverride
    Bypasses the blast-radius cap when explicitly approved.
.PARAMETER ManagedIdentityClientId
    Optional Client ID of a User-Assigned Managed Identity.
.PARAMETER ArmAccessToken
    Optional pre-acquired ARM bearer token.
.PARAMETER LogAnalyticsAccessToken
    Optional pre-acquired Log Analytics bearer token.
.PARAMETER GraphAccessToken
    Optional pre-acquired Microsoft Graph bearer token.
.PARAMETER AuthMode
    Auto (default), ManagedIdentity, or AzContext.
.OUTPUTS
    PSCustomObject[]
.EXAMPLE
    .\Remove-InactiveAvdPersonalHosts.ps1 -ResourceGroupName 'rg-avd' -HostPoolName 'vdpool-pers-01' -WorkspaceId '11111111-2222-3333-4444-555555555555'
#>

#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $false)] [string] $SubscriptionId,
    [Parameter(Mandatory = $false)] [string] $ResourceGroupName,
    [Parameter(Mandatory = $false)] [string] $HostPoolName,
    [Parameter(Mandatory = $false)] [string[]] $HostPoolResourceId,
    [Parameter(Mandatory = $true)]  [string] $WorkspaceId,
    [Parameter(Mandatory = $false)] [ValidateSet('Audit', 'QuarantineOnly', 'Execute')] [string] $Mode = 'Audit',
    [Parameter(Mandatory = $false)] [ValidateRange(7, 365)]  [int] $InactiveDays = 30,
    [Parameter(Mandatory = $false)] [ValidateRange(1, 90)]   [int] $QuarantineGraceDays = 7,
    [Parameter(Mandatory = $false)] [ValidateRange(7, 365)]  [int] $SnapshotRetentionDays = 30,
    [Parameter(Mandatory = $false)] [string] $SnapshotResourceGroupId,
    [Parameter(Mandatory = $false)] [ValidateRange(1, 100)]  [int] $MaxDecommissionPercent = 15,
    [Parameter(Mandatory = $false)] [ValidateRange(1, 500)]  [int] $MaxDecommissionCount = 25,
    [Parameter(Mandatory = $false)] [ValidateRange(1, 30)]   [int] $MinTelemetryDays = 20,
    [Parameter(Mandatory = $false)] [ValidateSet('Disable', 'Delete', 'Skip')] [string] $EntraCleanupAction = 'Disable',
    [Parameter(Mandatory = $false)] [switch] $ForceCircuitBreakerOverride,
    [Parameter(Mandatory = $false)] [string] $ManagedIdentityClientId,
    [Parameter(Mandatory = $false)] [string] $ArmAccessToken,
    [Parameter(Mandatory = $false)] [string] $LogAnalyticsAccessToken,
    [Parameter(Mandatory = $false)] [string] $GraphAccessToken,
    [Parameter(Mandatory = $false)] [ValidateSet('Auto', 'ManagedIdentity', 'AzContext')] [string] $AuthMode = 'Auto'
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

function Get-AvdCleanupProp {
    <#
    .SYNOPSIS
        Reads a property or dictionary key safely under Set-StrictMode 2.0.
    .DESCRIPTION
        Returns the value of Name from InputObject or $null when missing.
    .PARAMETER InputObject
        Object or IDictionary to read.
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

function ConvertTo-AvdCleanupGuid {
    <#
    .SYNOPSIS
        Normalizes a GUID string and filters zero GUIDs or path characters.
    .DESCRIPTION
        Trims braces and whitespace, rejects zero GUIDs and path traversal, and returns lowercase.
    .PARAMETER Value
        Raw GUID or identifier string.
    .OUTPUTS
        System.String
    #>
    param([Parameter(Mandatory = $false)] [string] $Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $t = $Value.Trim().Trim('{', '}').Trim()
    if ([string]::IsNullOrWhiteSpace($t) -or $t -match '^0{8}-0{4}-0{4}-0{4}-0{12}$' -or $t -match '[/\\?#]' -or $t.Contains('..')) { return $null }
    return $t.ToLowerInvariant()
}

function ConvertTo-AvdKqlLiteral {
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

function ConvertTo-AvdUtcTicks {
    <#
    .SYNOPSIS
        Parses a DateTime, DateTimeOffset, or ISO-8601 string into UTC ticks.
    .DESCRIPTION
        Converts timestamps using InvariantCulture and AssumeUniversal semantics.
    .PARAMETER DateValue
        Timestamp value to parse.
    .OUTPUTS
        System.Int64
    #>
    param([Parameter(Mandatory = $false)] $DateValue)
    if ($null -eq $DateValue) { return [int64]0 }
    if ($DateValue -is [DateTimeOffset]) { return $DateValue.UtcTicks }
    if ($DateValue -is [DateTime]) {
        if ($DateValue.Kind -eq [DateTimeKind]::Unspecified) { return ([DateTime]::SpecifyKind($DateValue, [DateTimeKind]::Utc)).Ticks }
        return $DateValue.ToUniversalTime().Ticks
    }
    $dto = [DateTimeOffset]::MinValue
    if ([DateTimeOffset]::TryParse(([string]$DateValue).Trim(), [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$dto)) {
        return $dto.UtcTicks
    }
    return [int64]0
}

function Get-AvdCleanupNowUtc {
    <#
    .SYNOPSIS
        Returns current UTC time (mockable for deterministic tests).
    .DESCRIPTION
        Returns [DateTimeOffset]::UtcNow.
    .OUTPUTS
        System.DateTimeOffset
    #>
    param()
    return [DateTimeOffset]::UtcNow
}

function New-AvdCleanupResult {
    <#
    .SYNOPSIS
        Creates a standardized output record for a session host or pool evaluation.
    .DESCRIPTION
        Constructs the PSCustomObject emitted by Remove-InactiveAvdPersonalHosts.ps1.
    .PARAMETER HostPoolName
        Name of the AVD Host Pool.
    .PARAMETER SessionHostName
        Leaf name of the session host.
    .PARAMETER ShortHostName
        Uppercase NetBIOS/short computer name.
    .PARAMETER VmId
        Normalized Compute vmId GUID.
    .PARAMETER AssignedUser
        Assigned user UPN.
    .PARAMETER LifecycleState
        Current lifecycle state (Active, Quarantined, Decommissioned, etc.).
    .PARAMETER Action
        Action performed during this run.
    .PARAMETER Status
        Evaluation or transition outcome code.
    .PARAMETER SnapshotIds
        Array of recovery snapshot ARM resource IDs created.
    .PARAMETER Detail
        Human-readable explanation.
    .OUTPUTS
        PSCustomObject
    #>
    param(
        [Parameter(Mandatory = $true)]  [string] $HostPoolName,
        [Parameter(Mandatory = $false)] [string] $SessionHostName,
        [Parameter(Mandatory = $false)] [string] $ShortHostName,
        [Parameter(Mandatory = $false)] [string] $VmId,
        [Parameter(Mandatory = $false)] [string] $AssignedUser,
        [Parameter(Mandatory = $true)]  [string] $LifecycleState,
        [Parameter(Mandatory = $true)]  [string] $Action,
        [Parameter(Mandatory = $true)]  [string] $Status,
        [Parameter(Mandatory = $false)] [string[]] $SnapshotIds = @(),
        [Parameter(Mandatory = $true)]  [string] $Detail
    )
    return [PSCustomObject]@{
        HostPoolName    = $HostPoolName
        SessionHostName = $SessionHostName
        ShortHostName   = $ShortHostName
        VmId            = $VmId
        AssignedUser    = $AssignedUser
        LifecycleState  = $LifecycleState
        Action          = $Action
        Status          = $Status
        SnapshotIds     = @($SnapshotIds)
        Detail          = $Detail
    }
}

function Invoke-AvdCleanupHttpRequest {
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

function Get-AvdCleanupAccessToken {
    <#
    .SYNOPSIS
        Acquires an OAuth2 bearer token for ARM, Log Analytics, or Microsoft Graph.
    .DESCRIPTION
        Supports Managed Identity (IDENTITY_ENDPOINT / IMDS) and Az.Accounts (Get-AzAccessToken).
    .PARAMETER ResourceUri
        Target resource audience URI.
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
        $resp = Invoke-AvdCleanupHttpRequest -Method Get -Uri ("{0}?{1}" -f $ep, $q) -Headers $hdr
        $tok = [string](Get-AvdCleanupProp -InputObject $resp -Name 'access_token')
        if ([string]::IsNullOrWhiteSpace($tok)) { throw "Managed Identity returned empty token for '$ResourceUri'." }
        return $tok
    }
    $tokObj = Get-AzAccessToken -ResourceUrl $res -ErrorAction Stop
    $rawTok = Get-AvdCleanupProp -InputObject $tokObj -Name 'Token'
    if ($rawTok -is [System.Security.SecureString]) {
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($rawTok)
        try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    }
    return [string]$rawTok
}

function Invoke-AvdCleanupRest {
    <#
    .SYNOPSIS
        Invokes a REST request with HTTP 429/5xx retry and single 401 token refresh.
    .DESCRIPTION
        Extracts status codes from WebException.Response or mocked Exception.Data.
    .PARAMETER Method
        HTTP method (Get, Post, Delete, Patch, Put).
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
        [Parameter(Mandatory = $true)]  [ValidateSet('Get', 'Post', 'Delete', 'Patch', 'Put')] [string] $Method,
        [Parameter(Mandatory = $true)]  [string] $Uri,
        [Parameter(Mandatory = $true)]  [hashtable] $Headers,
        [Parameter(Mandatory = $false)] [string] $Body
    )
    $attempt = 0
    $refreshed = $false
    while ($true) {
        $attempt++
        try {
            return (Invoke-AvdCleanupHttpRequest -Method $Method -Uri $Uri -Headers $Headers -Body $Body)
        } catch {
            $status = 0
            $retrySec = 0
            $resp = Get-AvdCleanupProp -InputObject $_.Exception -Name 'Response'
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
                $aud = if ($Uri -like 'https://management.azure.com/*') { 'https://management.azure.com/' } elseif ($Uri -like 'https://api.loganalytics.io/*') { 'https://api.loganalytics.io/' } elseif ($Uri -like 'https://graph.microsoft.com/*') { 'https://graph.microsoft.com/' } else { '' }
                $amVar = Get-Variable -Name 'AuthMode' -ErrorAction SilentlyContinue
                $miVar = Get-Variable -Name 'ManagedIdentityClientId' -ErrorAction SilentlyContinue
                if (-not [string]::IsNullOrWhiteSpace($aud) -and $null -ne $amVar) {
                    try {
                        $miClient = if ($null -ne $miVar) { [string]$miVar.Value } else { '' }
                        $newTok = Get-AvdCleanupAccessToken -ResourceUri $aud -SelectedAuthMode ([string]$amVar.Value) -ClientId $miClient
                        if (-not [string]::IsNullOrWhiteSpace($newTok)) { $Headers['Authorization'] = "Bearer $newTok"; continue }
                    } catch { }
                }
            }
            if (($status -in @(429, 502, 503, 504)) -and $attempt -lt 4) {
                $wait = if ($retrySec -gt 0) { [Math]::Min($retrySec, 300) } else { [int][Math]::Pow(2, $attempt) }
                if ($wait -gt 0) { Start-Sleep -Seconds $wait }
                continue
            }
            $ex = New-Object System.InvalidOperationException("REST $Method '$Uri' failed (HTTP $status): $($_.Exception.Message)", $_.Exception)
            if ($status -gt 0) { $ex.Data['HttpStatusCode'] = $status }
            throw $ex
        }
    }
}

function Invoke-AvdCleanupPagedGet {
    <#
    .SYNOPSIS
        Follows ARM or Microsoft Graph nextLink / @odata.nextLink pagination on the same HTTPS origin.
    .DESCRIPTION
        Collects all value items across pages while verifying same-origin HTTPS links.
    .PARAMETER Uri
        Initial collection GET URI.
    .PARAMETER Headers
        Request headers hashtable.
    .OUTPUTS
        System.Object[]
    #>
    param([Parameter(Mandatory = $true)] [string] $Uri, [Parameter(Mandatory = $true)] [hashtable] $Headers)
    $baseUri = [System.Uri]$Uri
    $items = New-Object 'System.Collections.Generic.List[object]'
    $next = $Uri
    while (-not [string]::IsNullOrWhiteSpace($next)) {
        $u = [System.Uri]$next
        if ($u.Scheme -ne 'https' -or -not $u.Authority.Equals($baseUri.Authority, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing to follow off-origin pagination link '$next'."
        }
        $resp = Invoke-AvdCleanupRest -Method Get -Uri $next -Headers $Headers
        foreach ($e in @(Get-AvdCleanupProp -InputObject $resp -Name 'value')) {
            if ($null -ne $e) { [void]$items.Add($e) }
        }
        $nl = Get-AvdCleanupProp -InputObject $resp -Name 'nextLink'
        if ([string]::IsNullOrWhiteSpace([string]$nl)) { $nl = Get-AvdCleanupProp -InputObject $resp -Name '@odata.nextLink' }
        $next = [string]$nl
    }
    return $items.ToArray()
}

function Invoke-AvdLogAnalyticsQuery {
    <#
    .SYNOPSIS
        Executes a KQL query against Log Analytics and fails closed on PartialError.
    .DESCRIPTION
        POSTs to /v1/workspaces/{WorkspaceGuid}/query, rejects any response with
        error or partialError, and converts PrimaryResult rows into PSCustomObjects.
    .PARAMETER WorkspaceGuid
        Log Analytics Workspace GUID.
    .PARAMETER Query
        KQL query text.
    .PARAMETER TimespanIso
        ISO-8601 duration timespan (e.g. P35D).
    .PARAMETER LaHeaders
        Log Analytics request headers hashtable.
    .OUTPUTS
        PSCustomObject[]
    #>
    param(
        [Parameter(Mandatory = $true)]  [string] $WorkspaceGuid,
        [Parameter(Mandatory = $true)]  [string] $Query,
        [Parameter(Mandatory = $false)] [string] $TimespanIso,
        [Parameter(Mandatory = $true)]  [hashtable] $LaHeaders
    )
    $uri = "https://api.loganalytics.io/v1/workspaces/{0}/query" -f $WorkspaceGuid
    $hdrs = @{ Authorization = $LaHeaders['Authorization']; Prefer = 'response-v1=true,wait=180' }
    $body = @{ query = $Query }
    if (-not [string]::IsNullOrWhiteSpace($TimespanIso)) { $body['timespan'] = $TimespanIso }
    $resp = Invoke-AvdCleanupRest -Method Post -Uri $uri -Headers $hdrs -Body ($body | ConvertTo-Json -Compress)
    $err = Get-AvdCleanupProp -InputObject $resp -Name 'error'
    if ($null -eq $err) { $err = Get-AvdCleanupProp -InputObject $resp -Name 'partialError' }
    if ($null -ne $err) {
        throw "Log Analytics query returned partial/error state (code='$(Get-AvdCleanupProp -InputObject $err -Name 'code')'): $(Get-AvdCleanupProp -InputObject $err -Name 'message')"
    }
    $tables = @(Get-AvdCleanupProp -InputObject $resp -Name 'tables')
    if ($tables.Count -eq 0 -or $null -eq $tables[0]) { throw "Log Analytics response contained no tables." }
    $tbl = $tables[0]
    foreach ($t in $tables) {
        if ([string](Get-AvdCleanupProp -InputObject $t -Name 'name') -eq 'PrimaryResult') { $tbl = $t; break }
    }
    $cols = @(Get-AvdCleanupProp -InputObject $tbl -Name 'columns')
    $rows = @(Get-AvdCleanupProp -InputObject $tbl -Name 'rows')
    $out = New-Object 'System.Collections.Generic.List[object]'
    foreach ($r in $rows) {
        if ($null -eq $r) { continue }
        $arr = @($r)
        $h = [ordered]@{}
        for ($i = 0; $i -lt $cols.Count; $i++) {
            $cName = [string](Get-AvdCleanupProp -InputObject $cols[$i] -Name 'name')
            $h[$cName] = if ($i -lt $arr.Count) { $arr[$i] } else { $null }
        }
        [void]$out.Add([PSCustomObject]$h)
    }
    return $out.ToArray()
}

function Test-AvdPoolTelemetryContinuity {
    <#
    .SYNOPSIS
        Verifies WVDConnections telemetry continuity for the Host Pool before inactivity evaluation.
    .DESCRIPTION
        Requires WVDConnections to have events predating (NowUtc - InactiveDays),
        recent events within the last 7 days, and at least RequiredDaysCount daily
        buckets over the last 30 days.
    .PARAMETER WorkspaceGuid
        Log Analytics Workspace GUID.
    .PARAMETER HostPoolArmId
        Host Pool ARM Resource ID.
    .PARAMETER InactiveDaysCount
        Inactivity threshold in days.
    .PARAMETER RequiredDaysCount
        Minimum populated daily buckets out of 30.
    .PARAMETER NowUtc
        Reference UTC timestamp.
    .PARAMETER LaHeaders
        Log Analytics headers hashtable.
    .OUTPUTS
        PSCustomObject
    #>
    param(
        [Parameter(Mandatory = $true)] [string] $WorkspaceGuid,
        [Parameter(Mandatory = $true)] [string] $HostPoolArmId,
        [Parameter(Mandatory = $true)] [int] $InactiveDaysCount,
        [Parameter(Mandatory = $true)] [int] $RequiredDaysCount,
        [Parameter(Mandatory = $true)] [DateTimeOffset] $NowUtc,
        [Parameter(Mandatory = $true)] [hashtable] $LaHeaders
    )
    $poolLit  = ConvertTo-AvdKqlLiteral -Value $HostPoolArmId.Trim().ToLowerInvariant()
    $lookback = [Math]::Max($InactiveDaysCount, 30) + 5
    $reqDays  = [Math]::Min($RequiredDaysCount, $InactiveDaysCount)
    $nowIso   = $NowUtc.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ss.fffZ', [Globalization.CultureInfo]::InvariantCulture)
    $kql = @"
let refNow = datetime($nowIso);
WVDConnections
| where TimeGenerated between ((refNow - $($lookback)d) .. refNow)
| where tolower(_ResourceId) == $poolLit
| summarize
    ConnEventCount = count(),
    EarliestConnUtc = min(TimeGenerated),
    LatestConnUtc = max(TimeGenerated),
    ActiveDays30d = dcountif(bin_at(TimeGenerated, 1d, refNow), TimeGenerated >= (refNow - 30d))
"@
    $rows = @(Invoke-AvdLogAnalyticsQuery -WorkspaceGuid $WorkspaceGuid -Query $kql -TimespanIso ("P{0}D" -f $lookback) -LaHeaders $LaHeaders)
    if ($rows.Count -eq 0) { return [PSCustomObject]@{ IsHealthy = $false; Reason = 'Continuity query returned no rows.' } }
    $r = $rows[0]
    $cnt = [int](Get-AvdCleanupProp -InputObject $r -Name 'ConnEventCount')
    $days = [int](Get-AvdCleanupProp -InputObject $r -Name 'ActiveDays30d')
    $earliestTicks = ConvertTo-AvdUtcTicks -DateValue (Get-AvdCleanupProp -InputObject $r -Name 'EarliestConnUtc')
    $latestTicks   = ConvertTo-AvdUtcTicks -DateValue (Get-AvdCleanupProp -InputObject $r -Name 'LatestConnUtc')
    if ($cnt -le 0 -or $earliestTicks -le 0) {
        return [PSCustomObject]@{ IsHealthy = $false; Reason = "WVDConnections has 0 events for pool '$HostPoolArmId'." }
    }
    if ($earliestTicks -gt $NowUtc.AddDays(-$InactiveDaysCount).UtcTicks) {
        return [PSCustomObject]@{ IsHealthy = $false; Reason = "Earliest WVDConnections event is newer than the $InactiveDaysCount-day threshold." }
    }
    if ($latestTicks -lt $NowUtc.AddDays(-7).UtcTicks) {
        return [PSCustomObject]@{ IsHealthy = $false; Reason = "Latest WVDConnections event is older than 7 days; diagnostic ingestion may be stopped." }
    }
    if ($days -lt $reqDays) {
        return [PSCustomObject]@{ IsHealthy = $false; Reason = "WVDConnections covered $days/30 daily buckets (minimum required: $reqDays)." }
    }
    return [PSCustomObject]@{ IsHealthy = $true; Reason = "Continuity verified ($days/30 active days, $cnt connection events)." }
}

function Merge-AvdResourceTags {
    <#
    .SYNOPSIS
        Merges tags onto an Azure resource via Microsoft.Resources/tags/default.
    .DESCRIPTION
        Uses operation='Merge' so existing VM tags are preserved.
    .PARAMETER ResourceId
        Target ARM Resource ID.
    .PARAMETER Tags
        Hashtable of tags to merge.
    .PARAMETER ArmHeaders
        ARM headers hashtable.
    .OUTPUTS
        None
    #>
    param([Parameter(Mandatory = $true)] [string] $ResourceId, [Parameter(Mandatory = $true)] [hashtable] $Tags, [Parameter(Mandatory = $true)] [hashtable] $ArmHeaders)
    $uri = "https://management.azure.com{0}/providers/Microsoft.Resources/tags/default?api-version=2021-04-01" -f $ResourceId.TrimEnd('/')
    $body = @{ operation = 'Merge'; properties = @{ tags = $Tags } } | ConvertTo-Json -Compress -Depth 5
    $null = Invoke-AvdCleanupRest -Method Patch -Uri $uri -Headers $ArmHeaders -Body $body
}

function Reset-AvdHostToActiveState {
    <#
    .SYNOPSIS
        Restores allowNewSession=true on the session host and resets VM quarantine tags to Active.
    .DESCRIPTION
        Called when a Quarantined host shows live sessions, recent connections, user
        reassignment, mismatched vmId, or is no longer deallocated.
    .PARAMETER HostPoolArmId
        Host Pool ARM Resource ID.
    .PARAMETER SessionHostLeafName
        Leaf name of the session host.
    .PARAMETER VmResourceId
        ARM Resource ID of the VM.
    .PARAMETER CurrentAllowNewSession
        Current allowNewSession boolean value.
    .PARAMETER AdditionalTags
        Optional extra tags to merge.
    .PARAMETER ArmHeaders
        ARM headers hashtable.
    .OUTPUTS
        None
    #>
    param(
        [Parameter(Mandatory = $true)]  [string] $HostPoolArmId,
        [Parameter(Mandatory = $true)]  [string] $SessionHostLeafName,
        [Parameter(Mandatory = $true)]  [string] $VmResourceId,
        [Parameter(Mandatory = $true)]  [bool] $CurrentAllowNewSession,
        [Parameter(Mandatory = $false)] [hashtable] $AdditionalTags,
        [Parameter(Mandatory = $true)]  [hashtable] $ArmHeaders
    )
    if (-not $CurrentAllowNewSession) {
        $shUri = "https://management.azure.com{0}/sessionHosts/{1}?api-version=2022-09-09" -f $HostPoolArmId.TrimEnd('/'), [uri]::EscapeDataString($SessionHostLeafName)
        $null = Invoke-AvdCleanupRest -Method Patch -Uri $shUri -Headers $ArmHeaders -Body (@{ properties = @{ allowNewSession = $true } } | ConvertTo-Json -Compress)
    }
    $tags = @{ AvdLifecycleState = 'Active'; AvdQuarantinedVmId = ''; AvdQuarantinedAtUtc = ''; AvdDecommissionAfterUtc = '' }
    if ($null -ne $AdditionalTags) { foreach ($k in $AdditionalTags.Keys) { $tags[$k] = $AdditionalTags[$k] } }
    Merge-AvdResourceTags -ResourceId $VmResourceId -Tags $tags -ArmHeaders $ArmHeaders
}

# --- MAIN EXECUTION ---
$poolIdPattern = '^/subscriptions/[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}/resourceGroups/[A-Za-z0-9_.()\-]{1,90}/providers/Microsoft\.DesktopVirtualization/hostPools/[A-Za-z0-9_.\-]{1,64}$'
$guidRegex     = '^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$'
$normWsId      = ConvertTo-AvdCleanupGuid -Value $WorkspaceId
if ([string]::IsNullOrWhiteSpace($normWsId) -or $normWsId -notmatch $guidRegex) {
    throw "Invalid -WorkspaceId '$WorkspaceId'. Expected a Log Analytics Workspace GUID."
}

$targetPools = New-Object 'System.Collections.Generic.List[string]'
if ($null -ne $HostPoolResourceId) {
    foreach ($id in $HostPoolResourceId) {
        if (-not [string]::IsNullOrWhiteSpace($id)) {
            $clean = $id.Trim()
            if ($clean -notmatch $poolIdPattern) { throw "Invalid HostPoolResourceId '$clean'." }
            if (-not $targetPools.Contains($clean)) { [void]$targetPools.Add($clean) }
        }
    }
}
if (-not [string]::IsNullOrWhiteSpace($HostPoolName) -or -not [string]::IsNullOrWhiteSpace($ResourceGroupName)) {
    if ([string]::IsNullOrWhiteSpace($HostPoolName) -or [string]::IsNullOrWhiteSpace($ResourceGroupName)) {
        throw "Both -HostPoolName and -ResourceGroupName must be supplied together."
    }
    $sub = if (-not [string]::IsNullOrWhiteSpace($SubscriptionId)) { $SubscriptionId.Trim() } else { [string](Get-AzContext).Subscription.Id }
    $built = "/subscriptions/$sub/resourceGroups/$($ResourceGroupName.Trim())/providers/Microsoft.DesktopVirtualization/hostPools/$($HostPoolName.Trim())"
    if ($built -notmatch $poolIdPattern) { throw "Invalid Host Pool ARM ID '$built'." }
    if (-not $targetPools.Contains($built)) { [void]$targetPools.Add($built) }
}
if ($targetPools.Count -eq 0) { throw "Specify -HostPoolResourceId or both -ResourceGroupName and -HostPoolName." }

$armTok = if (-not [string]::IsNullOrWhiteSpace($ArmAccessToken)) { $ArmAccessToken.Trim() } else { Get-AvdCleanupAccessToken -ResourceUri 'https://management.azure.com/' -SelectedAuthMode $AuthMode -ClientId $ManagedIdentityClientId }
$laTok  = if (-not [string]::IsNullOrWhiteSpace($LogAnalyticsAccessToken)) { $LogAnalyticsAccessToken.Trim() } else { Get-AvdCleanupAccessToken -ResourceUri 'https://api.loganalytics.io/' -SelectedAuthMode $AuthMode -ClientId $ManagedIdentityClientId }
$grpTok = if (-not [string]::IsNullOrWhiteSpace($GraphAccessToken)) { $GraphAccessToken.Trim() } else { Get-AvdCleanupAccessToken -ResourceUri 'https://graph.microsoft.com/' -SelectedAuthMode $AuthMode -ClientId $ManagedIdentityClientId }

$armHeaders   = @{ Authorization = "Bearer $armTok" }
$laHeaders    = @{ Authorization = "Bearer $laTok" }
$graphHeaders = @{ Authorization = "Bearer $grpTok" }
$nowUtc       = Get-AvdCleanupNowUtc
$nowTicks     = $nowUtc.UtcTicks
$nowIso       = $nowUtc.UtcDateTime.ToString('o', [Globalization.CultureInfo]::InvariantCulture)
$minAgeTicks  = $nowUtc.AddDays(-$InactiveDays).UtcTicks
$results      = New-Object 'System.Collections.Generic.List[object]'

foreach ($poolId in $targetPools) {
    $poolName = ($poolId -split '/')[-1]
    $poolObj  = Invoke-AvdCleanupRest -Method Get -Uri ("https://management.azure.com{0}?api-version=2022-09-09" -f $poolId) -Headers $armHeaders
    $poolType = [string](Get-AvdCleanupProp -InputObject (Get-AvdCleanupProp -InputObject $poolObj -Name 'properties') -Name 'hostPoolType')
    if (-not [string]::Equals($poolType, 'Personal', [System.StringComparison]::OrdinalIgnoreCase)) {
        [void]$results.Add((New-AvdCleanupResult -HostPoolName $poolName -LifecycleState 'N/A' -Action 'None' -Status 'SkippedNonPersonalHostPool' -Detail "Pool '$poolName' is '$poolType' (only Personal pools are eligible)."))
        continue
    }

    # 1. Telemetry continuity gate
    $gate = Test-AvdPoolTelemetryContinuity -WorkspaceGuid $normWsId -HostPoolArmId $poolId -InactiveDaysCount $InactiveDays -RequiredDaysCount $MinTelemetryDays -NowUtc $nowUtc -LaHeaders $laHeaders
    if (-not $gate.IsHealthy) {
        [void]$results.Add((New-AvdCleanupResult -HostPoolName $poolName -LifecycleState 'N/A' -Action 'None' -Status 'TelemetryContinuityGateFailed' -Detail $gate.Reason))
        continue
    }

    # 2. Query WVDConnections activity over InactiveDays
    $poolLit = ConvertTo-AvdKqlLiteral -Value $poolId.Trim().ToLowerInvariant()
    $actKql = @"
let refNow = datetime($($nowUtc.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ss.fffZ', [Globalization.CultureInfo]::InvariantCulture)));
WVDConnections
| where TimeGenerated between ((refNow - $($InactiveDays)d) .. refNow)
| where tolower(_ResourceId) == $poolLit
| where State in ('Started', 'Connected', 'Completed')
| extend VmIdKey = tolower(tostring(column_ifexists('SessionHostAzureVmId', '')))
| extend HostKey = toupper(tostring(split(SessionHostName, '.')[0]))
| summarize ConnectedCount = countif(State in ('Connected', 'Completed')), StartedOnlyCount = countif(State == 'Started') by VmIdKey, HostKey
"@
    $actRows = @(Invoke-AvdLogAnalyticsQuery -WorkspaceGuid $normWsId -Query $actKql -TimespanIso ("P{0}D" -f ($InactiveDays + 1)) -LaHeaders $laHeaders)
    $byVmId = @{}
    $byHost = @{}
    foreach ($row in $actRows) {
        $vk = ConvertTo-AvdCleanupGuid -Value ([string](Get-AvdCleanupProp -InputObject $row -Name 'VmIdKey'))
        $hk = ([string](Get-AvdCleanupProp -InputObject $row -Name 'HostKey')).Trim().ToUpperInvariant()
        $cCnt = [int](Get-AvdCleanupProp -InputObject $row -Name 'ConnectedCount')
        $sCnt = [int](Get-AvdCleanupProp -InputObject $row -Name 'StartedOnlyCount')
        if (-not [string]::IsNullOrWhiteSpace($vk)) {
            if (-not $byVmId.ContainsKey($vk)) { $byVmId[$vk] = @{ Connected = 0; Started = 0 } }
            $byVmId[$vk].Connected += $cCnt; $byVmId[$vk].Started += $sCnt
        }
        if (-not [string]::IsNullOrWhiteSpace($hk)) {
            if (-not $byHost.ContainsKey($hk)) { $byHost[$hk] = @{ Connected = 0; Started = 0 } }
            $byHost[$hk].Connected += $cCnt; $byHost[$hk].Started += $sCnt
        }
    }

    # 3. Enumerate session hosts and VMs
    $sessionHosts = @(Invoke-AvdCleanupPagedGet -Uri ("https://management.azure.com{0}/sessionHosts?api-version=2022-09-09" -f $poolId) -Headers $armHeaders)
    $hostItems = New-Object 'System.Collections.Generic.List[object]'
    $assignedCount = 0
    $quarantinedCount = 0

    foreach ($sh in $sessionHosts) {
        $shLeaf = (([string](Get-AvdCleanupProp -InputObject $sh -Name 'name')) -split '/')[-1].Trim()
        $short  = (($shLeaf -split '\.')[0]).Trim().ToUpperInvariant()
        $shProps = Get-AvdCleanupProp -InputObject $sh -Name 'properties'
        $upn = [string](Get-AvdCleanupProp -InputObject $shProps -Name 'assignedUser')
        if ([string]::IsNullOrWhiteSpace($upn)) {
            [void]$results.Add((New-AvdCleanupResult -HostPoolName $poolName -SessionHostName $shLeaf -ShortHostName $short -LifecycleState 'Active' -Action 'None' -Status 'UnassignedWarmHost' -Detail "No assignedUser; excluded as warm buffer."))
            continue
        }
        $assignedCount++
        $vmResId = [string](Get-AvdCleanupProp -InputObject $shProps -Name 'resourceId')
        $vmObj = $null
        try {
            $vmObj = Invoke-AvdCleanupRest -Method Get -Uri ("https://management.azure.com{0}?`$expand=instanceView&api-version=2023-09-01" -f $vmResId) -Headers $armHeaders
        } catch {
            $code = if ($null -ne $_.Exception.Data -and $_.Exception.Data.Contains('HttpStatusCode')) { [int]$_.Exception.Data['HttpStatusCode'] } else { 0 }
            $errStat = if ($code -eq 404) { 'OrphanedSessionHost' } else { 'VmLookupFailed' }
            [void]$results.Add((New-AvdCleanupResult -HostPoolName $poolName -SessionHostName $shLeaf -ShortHostName $short -AssignedUser $upn -LifecycleState 'Orphaned' -Action 'None' -Status $errStat -Detail $_.Exception.Message))
            continue
        }
        $st = [string](Get-AvdCleanupProp -InputObject (Get-AvdCleanupProp -InputObject $vmObj -Name 'tags') -Name 'AvdLifecycleState')
        if (-not [string]::IsNullOrWhiteSpace($st) -and $st.Trim() -eq 'Quarantined') { $quarantinedCount++ }
        [void]$hostItems.Add([PSCustomObject]@{ SessionHost = $sh; Leaf = $shLeaf; Short = $short; Upn = $upn.Trim(); VmResId = $vmResId; Vm = $vmObj })
    }

    $budgetCap = [Math]::Min($MaxDecommissionCount, [Math]::Max(1, [int][Math]::Floor($assignedCount * ($MaxDecommissionPercent / 100.0))))
    $budgetUsed = $quarantinedCount
    $decomCount = 0

    foreach ($item in $hostItems) {
        $shLeaf  = $item.Leaf
        $short   = $item.Short
        $upn     = $item.Upn
        $vmResId = $item.VmResId
        $vmObj   = $item.Vm
        $shProps = Get-AvdCleanupProp -InputObject $item.SessionHost -Name 'properties'
        $vmProps = Get-AvdCleanupProp -InputObject $vmObj -Name 'properties'
        $vmTags  = Get-AvdCleanupProp -InputObject $vmObj -Name 'tags'
        $vmId    = ConvertTo-AvdCleanupGuid -Value ([string](Get-AvdCleanupProp -InputObject $vmProps -Name 'vmId'))
        $allowNew = [bool](Get-AvdCleanupProp -InputObject $shProps -Name 'allowNewSession')
        $curState = [string](Get-AvdCleanupProp -InputObject $vmTags -Name 'AvdLifecycleState')
        if ([string]::IsNullOrWhiteSpace($curState)) { $curState = 'Active' } else { $curState = $curState.Trim() }

        # Guard 1: Legal hold
        $hold = [string](Get-AvdCleanupProp -InputObject $vmTags -Name 'AvdLegalHold')
        if (-not [string]::IsNullOrWhiteSpace($hold) -and $hold.Trim().ToLowerInvariant() -notin @('false', '0', 'no')) {
            [void]$results.Add((New-AvdCleanupResult -HostPoolName $poolName -SessionHostName $shLeaf -ShortHostName $short -VmId $vmId -AssignedUser $upn -LifecycleState 'LegalHold' -Action 'None' -Status 'ExcludedByLegalHold' -Detail "VM has AvdLegalHold='$($hold.Trim())'."))
            continue
        }

        # Guard 2: Live sessions
        $sessRaw = Get-AvdCleanupProp -InputObject $shProps -Name 'sessions'
        if ($null -eq $sessRaw -or [int]$sessRaw -gt 0) {
            $act = 'None'
            if ($null -ne $sessRaw -and $Mode -ne 'Audit' -and $curState -eq 'Quarantined' -and $PSCmdlet.ShouldProcess("VM '$short'", "Restore allowNewSession=true due to active session")) {
                Reset-AvdHostToActiveState -HostPoolArmId $poolId -SessionHostLeafName $shLeaf -VmResourceId $vmResId -CurrentAllowNewSession $allowNew -ArmHeaders $armHeaders
                $act = 'RestoredToActive'
            }
            $stCode = if ($null -eq $sessRaw) { 'SessionCountUnavailable' } else { 'ActiveSessionsPresent' }
            [void]$results.Add((New-AvdCleanupResult -HostPoolName $poolName -SessionHostName $shLeaf -ShortHostName $short -VmId $vmId -AssignedUser $upn -LifecycleState 'Active' -Action $act -Status $stCode -Detail "Active or unavailable session count ($sessRaw)."))
            continue
        }

        # Guard 3: WVDConnections activity (union of vmId and short hostname)
        $vAct = if ($byVmId.ContainsKey($vmId)) { $byVmId[$vmId] } else { @{ Connected = 0; Started = 0 } }
        $hAct = if ($byHost.ContainsKey($short)) { $byHost[$short] } else { @{ Connected = 0; Started = 0 } }
        $connCnt  = [Math]::Max([int]$vAct.Connected, [int]$hAct.Connected)
        $startCnt = [Math]::Max([int]$vAct.Started, [int]$hAct.Started)
        if ($connCnt -gt 0 -or $startCnt -gt 0) {
            $act = 'None'
            if ($Mode -ne 'Audit' -and $curState -eq 'Quarantined' -and $PSCmdlet.ShouldProcess("VM '$short'", "Restore allowNewSession=true due to connection activity")) {
                Reset-AvdHostToActiveState -HostPoolArmId $poolId -SessionHostLeafName $shLeaf -VmResourceId $vmResId -CurrentAllowNewSession $allowNew -ArmHeaders $armHeaders
                $act = 'RestoredToActive'
            }
            $stCode = if ($connCnt -gt 0) { 'ActiveWithinWindow' } else { 'ConnectionFailuresDetected' }
            [void]$results.Add((New-AvdCleanupResult -HostPoolName $poolName -SessionHostName $shLeaf -ShortHostName $short -VmId $vmId -AssignedUser $upn -LifecycleState 'Active' -Action $act -Status $stCode -Detail "Recorded $connCnt connected and $startCnt started event(s) in last $InactiveDays days."))
            continue
        }

        # Guard 4: Creation age
        $vmCreated = ConvertTo-AvdUtcTicks -DateValue (Get-AvdCleanupProp -InputObject $vmProps -Name 'timeCreated')
        if ($vmCreated -le 0 -or $vmCreated -gt $minAgeTicks) {
            [void]$results.Add((New-AvdCleanupResult -HostPoolName $poolName -SessionHostName $shLeaf -ShortHostName $short -VmId $vmId -AssignedUser $upn -LifecycleState 'Active' -Action 'None' -Status 'RecentlyCreatedHost' -Detail "VM created within the last $InactiveDays days."))
            continue
        }

        # Guard 5: User resolution & assignment age (resolve current UPN first to detect reassignments)
        $taggedUserId = ConvertTo-AvdCleanupGuid -Value ([string](Get-AvdCleanupProp -InputObject $vmTags -Name 'AvdAssignedUserObjectId'))
        $assignTicks  = ConvertTo-AvdUtcTicks -DateValue (Get-AvdCleanupProp -InputObject $vmTags -Name 'AvdAssignedAtUtc')
        $userObj = $null
        try {
            $userObj = Invoke-AvdCleanupRest -Method Get -Uri ("https://graph.microsoft.com/v1.0/users/{0}?`$select=id,userPrincipalName,accountEnabled" -f [uri]::EscapeDataString($upn)) -Headers $graphHeaders
        } catch {
            $uSt = if ($null -ne $_.Exception.Data -and $_.Exception.Data.Contains('HttpStatusCode')) { [int]$_.Exception.Data['HttpStatusCode'] } else { 0 }
            if ($uSt -ne 404) { throw }
        }
        $resolvedUserId = if ($null -ne $userObj) { ConvertTo-AvdCleanupGuid -Value ([string](Get-AvdCleanupProp -InputObject $userObj -Name 'id')) } else { $taggedUserId }
        $userChanged = (-not [string]::IsNullOrWhiteSpace($taggedUserId)) -and (-not [string]::IsNullOrWhiteSpace($resolvedUserId)) -and ($taggedUserId -ne $resolvedUserId)
        if ($assignTicks -le 0 -or $userChanged) {
            $act = 'None'
            if ($Mode -ne 'Audit' -and $PSCmdlet.ShouldProcess("VM '$short'", "Stamp AvdAssignedAtUtc")) {
                $assignTags = @{ AvdAssignedAtUtc = $nowIso; AvdAssignedUserObjectId = [string]$resolvedUserId }
                if ($curState -eq 'Quarantined') {
                    Reset-AvdHostToActiveState -HostPoolArmId $poolId -SessionHostLeafName $shLeaf -VmResourceId $vmResId -CurrentAllowNewSession $allowNew -AdditionalTags $assignTags -ArmHeaders $armHeaders
                } else {
                    Merge-AvdResourceTags -ResourceId $vmResId -Tags $assignTags -ArmHeaders $armHeaders
                }
                $act = 'StampedAssignmentTimestamp'
            }
            [void]$results.Add((New-AvdCleanupResult -HostPoolName $poolName -SessionHostName $shLeaf -ShortHostName $short -VmId $vmId -AssignedUser $upn -LifecycleState 'Active' -Action $act -Status 'AssignmentObservationWindowPending' -Detail "Missing or reassigned user timestamp; observation window started."))
            continue
        }
        if ($assignTicks -gt $minAgeTicks) {
            [void]$results.Add((New-AvdCleanupResult -HostPoolName $poolName -SessionHostName $shLeaf -ShortHostName $short -VmId $vmId -AssignedUser $upn -LifecycleState 'Active' -Action 'None' -Status 'RecentlyAssignedUser' -Detail "Assigned less than $InactiveDays days ago."))
            continue
        }

        # Guard 6: Validate existing Quarantine state binding
        $qVmId = ConvertTo-AvdCleanupGuid -Value ([string](Get-AvdCleanupProp -InputObject $vmTags -Name 'AvdQuarantinedVmId'))
        if ($curState -eq 'Quarantined' -and ($qVmId -ne $vmId -or $allowNew)) {
            $act = 'None'
            if ($Mode -ne 'Audit' -and $PSCmdlet.ShouldProcess("VM '$short'", "Reset invalidated quarantine tags to Active")) {
                Reset-AvdHostToActiveState -HostPoolArmId $poolId -SessionHostLeafName $shLeaf -VmResourceId $vmResId -CurrentAllowNewSession $allowNew -ArmHeaders $armHeaders
                $act = 'ResetToActive'
            }
            [void]$results.Add((New-AvdCleanupResult -HostPoolName $poolName -SessionHostName $shLeaf -ShortHostName $short -VmId $vmId -AssignedUser $upn -LifecycleState 'Active' -Action $act -Status 'QuarantineInvalidatedResetToActive' -Detail "Quarantine tag vmId mismatch or allowNewSession=true; reset to Active."))
            continue
        }

        # PHASE 1: Active -> Quarantined (Drain + Deallocate)
        if ($curState -eq 'Active') {
            if ($budgetUsed -ge $budgetCap -and -not $ForceCircuitBreakerOverride) {
                [void]$results.Add((New-AvdCleanupResult -HostPoolName $poolName -SessionHostName $shLeaf -ShortHostName $short -VmId $vmId -AssignedUser $upn -LifecycleState 'Active' -Action 'None' -Status 'CircuitBreakerBlocked' -Detail "Pool quarantine budget reached ($budgetUsed/$budgetCap)."))
                continue
            }
            if ($Mode -eq 'Audit') {
                $budgetUsed++
                [void]$results.Add((New-AvdCleanupResult -HostPoolName $poolName -SessionHostName $shLeaf -ShortHostName $short -VmId $vmId -AssignedUser $upn -LifecycleState 'Active' -Action 'AuditEligibleForQuarantine' -Status 'EligibleForQuarantine' -Detail "Unused for >= $InactiveDays days; eligible for Phase 1 quarantine."))
                continue
            }
            if ($PSCmdlet.ShouldProcess("VM '$short' ($vmId)", "Phase 1: Drain session host, deallocate VM, and mark Quarantined")) {
                $shUri = "https://management.azure.com{0}/sessionHosts/{1}?api-version=2022-09-09" -f $poolId, [uri]::EscapeDataString($shLeaf)
                $null = Invoke-AvdCleanupRest -Method Patch -Uri $shUri -Headers $armHeaders -Body (@{ properties = @{ allowNewSession = $false } } | ConvertTo-Json -Compress)
                $null = Invoke-AvdCleanupRest -Method Post -Uri ("https://management.azure.com{0}/deallocate?api-version=2023-09-01" -f $vmResId) -Headers $armHeaders
                $decomAfterIso = $nowUtc.AddDays($QuarantineGraceDays).UtcDateTime.ToString('o', [Globalization.CultureInfo]::InvariantCulture)
                Merge-AvdResourceTags -ResourceId $vmResId -Tags @{
                    AvdLifecycleState       = 'Quarantined'
                    AvdQuarantinedVmId      = $vmId
                    AvdQuarantinedAtUtc     = $nowIso
                    AvdDecommissionAfterUtc = $decomAfterIso
                    AvdAssignedUserObjectId = [string]$resolvedUserId
                } -ArmHeaders $armHeaders
                $budgetUsed++
                [void]$results.Add((New-AvdCleanupResult -HostPoolName $poolName -SessionHostName $shLeaf -ShortHostName $short -VmId $vmId -AssignedUser $upn -LifecycleState 'Quarantined' -Action 'DrainedAndQuarantined' -Status 'TransitionedToQuarantined' -Detail "Drained, deallocated, and quarantined until $decomAfterIso."))
            } else {
                [void]$results.Add((New-AvdCleanupResult -HostPoolName $poolName -SessionHostName $shLeaf -ShortHostName $short -VmId $vmId -AssignedUser $upn -LifecycleState 'Active' -Action 'WhatIfSkipped' -Status 'EligibleForQuarantine' -Detail "Skipped Phase 1 quarantine (WhatIf)."))
            }
            continue
        }

        # PHASE 2: Quarantined -> Decommissioned (Snapshot + Teardown)
        if ($curState -eq 'Quarantined') {
            $decomAfterTicks = ConvertTo-AvdUtcTicks -DateValue (Get-AvdCleanupProp -InputObject $vmTags -Name 'AvdDecommissionAfterUtc')
            if ($decomAfterTicks -le 0 -or $nowTicks -lt $decomAfterTicks) {
                [void]$results.Add((New-AvdCleanupResult -HostPoolName $poolName -SessionHostName $shLeaf -ShortHostName $short -VmId $vmId -AssignedUser $upn -LifecycleState 'Quarantined' -Action 'None' -Status 'InQuarantineGracePeriod' -Detail "Quarantine grace period active until $(Get-AvdCleanupProp -InputObject $vmTags -Name 'AvdDecommissionAfterUtc')."))
                continue
            }
            $statuses = @(Get-AvdCleanupProp -InputObject (Get-AvdCleanupProp -InputObject $vmProps -Name 'instanceView') -Name 'statuses')
            $isDeallocated = $false
            foreach ($s in $statuses) {
                if ([string]::Equals([string](Get-AvdCleanupProp -InputObject $s -Name 'code'), 'PowerState/deallocated', [System.StringComparison]::OrdinalIgnoreCase)) { $isDeallocated = $true; break }
            }
            if (-not $isDeallocated) {
                $act = 'None'
                if ($Mode -ne 'Audit' -and $PSCmdlet.ShouldProcess("VM '$short'", "Reset quarantine because VM is running")) {
                    Reset-AvdHostToActiveState -HostPoolArmId $poolId -SessionHostLeafName $shLeaf -VmResourceId $vmResId -CurrentAllowNewSession $allowNew -ArmHeaders $armHeaders
                    $act = 'ResetToActive'
                }
                [void]$results.Add((New-AvdCleanupResult -HostPoolName $poolName -SessionHostName $shLeaf -ShortHostName $short -VmId $vmId -AssignedUser $upn -LifecycleState 'Active' -Action $act -Status 'QuarantinedVmNotDeallocated' -Detail "VM is not deallocated; aborted Phase 2."))
                continue
            }
            if ($Mode -ne 'Execute') {
                [void]$results.Add((New-AvdCleanupResult -HostPoolName $poolName -SessionHostName $shLeaf -ShortHostName $short -VmId $vmId -AssignedUser $upn -LifecycleState 'Quarantined' -Action 'AuditEligibleForDecommission' -Status 'EligibleForDecommission' -Detail "Quarantine expired; run with -Mode Execute to snapshot and remove."))
                continue
            }
            if ($decomCount -ge $budgetCap -and -not $ForceCircuitBreakerOverride) {
                [void]$results.Add((New-AvdCleanupResult -HostPoolName $poolName -SessionHostName $shLeaf -ShortHostName $short -VmId $vmId -AssignedUser $upn -LifecycleState 'Quarantined' -Action 'None' -Status 'CircuitBreakerBlocked' -Detail "Phase 2 decommission budget reached ($decomCount/$budgetCap)."))
                continue
            }

            if ($PSCmdlet.ShouldProcess("VM '$short' ($vmId)", "Phase 2: Snapshot disks and remove Intune, Entra, VM, and SessionHost")) {
                # 1. Snapshot OS disk and any data disks
                $storage = Get-AvdCleanupProp -InputObject $vmProps -Name 'storageProfile'
                $osDiskId = [string](Get-AvdCleanupProp -InputObject (Get-AvdCleanupProp -InputObject (Get-AvdCleanupProp -InputObject $storage -Name 'osDisk') -Name 'managedDisk') -Name 'id')
                $disks = New-Object 'System.Collections.Generic.List[object]'
                [void]$disks.Add(@{ Id = $osDiskId; Label = 'os' })
                foreach ($dd in @(Get-AvdCleanupProp -InputObject $storage -Name 'dataDisks')) {
                    if ($null -eq $dd) { continue }
                    $ddId = [string](Get-AvdCleanupProp -InputObject (Get-AvdCleanupProp -InputObject $dd -Name 'managedDisk') -Name 'id')
                    $lun  = [string](Get-AvdCleanupProp -InputObject $dd -Name 'lun')
                    [void]$disks.Add(@{ Id = $ddId; Label = "lun$lun" })
                }
                $snapRg = if (-not [string]::IsNullOrWhiteSpace($SnapshotResourceGroupId)) { $SnapshotResourceGroupId.Trim() } else { ([regex]::Match($vmResId, '^(/subscriptions/[^/]+/resourceGroups/[^/]+)/')).Groups[1].Value }
                $vmPrefix = ($vmId -replace '[^0-9a-f]', '').Substring(0, 8)
                $delAfter = $nowUtc.AddDays($SnapshotRetentionDays).UtcDateTime.ToString('o', [Globalization.CultureInfo]::InvariantCulture)
                $snapIds = New-Object 'System.Collections.Generic.List[string]'
                foreach ($d in $disks) {
                    $sName = "snap-$($short.ToLowerInvariant())-$vmPrefix-$($d.Label)"
                    $sId   = "$($snapRg.TrimEnd('/'))/providers/Microsoft.Compute/snapshots/$sName"
                    $sUri  = "https://management.azure.com{0}?api-version=2023-04-02" -f $sId
                    $sBody = @{
                        location   = [string](Get-AvdCleanupProp -InputObject $vmObj -Name 'location')
                        sku        = @{ name = 'Standard_LRS' }
                        tags       = @{ SourceVmId = $vmId; AssignedUserObjectId = [string]$resolvedUserId; DeleteAfterUtc = $delAfter }
                        properties = @{ incremental = $true; creationData = @{ createOption = 'Copy'; sourceResourceId = $d.Id } }
                    } | ConvertTo-Json -Compress -Depth 5
                    $sResp = Invoke-AvdCleanupRest -Method Put -Uri $sUri -Headers $armHeaders -Body $sBody
                    $pState = [string](Get-AvdCleanupProp -InputObject (Get-AvdCleanupProp -InputObject $sResp -Name 'properties') -Name 'provisioningState')
                    for ($poll = 0; $pState -in @('Creating', 'Updating') -and $poll -lt 20; $poll++) {
                        $sResp = Invoke-AvdCleanupRest -Method Get -Uri $sUri -Headers $armHeaders
                        $pState = [string](Get-AvdCleanupProp -InputObject (Get-AvdCleanupProp -InputObject $sResp -Name 'properties') -Name 'provisioningState')
                    }
                    if ($pState -ne 'Succeeded') { throw "Snapshot '$sName' state is '$pState' (expected 'Succeeded')." }
                    [void]$snapIds.Add($sId)
                }

                # 2. Correlate and clean up Intune + Entra ID device records
                $escHost = $short.Replace("'", "''")
                $entraDevs = @(Invoke-AvdCleanupPagedGet -Uri ("https://graph.microsoft.com/v1.0/devices?`$filter=$([uri]::EscapeDataString("displayName eq '$escHost'"))&`$select=id,deviceId,trustType") -Headers $graphHeaders)
                $intuneDevs = @(Invoke-AvdCleanupPagedGet -Uri ("https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=$([uri]::EscapeDataString("operatingSystem eq 'Windows' and deviceName eq '$escHost'"))&`$select=id,azureADDeviceId") -Headers $graphHeaders)
                if ($entraDevs.Count -gt 1 -or $intuneDevs.Count -gt 1) {
                    throw "Ambiguous directory/Intune records for '$short' (Entra=$($entraDevs.Count), Intune=$($intuneDevs.Count))."
                }
                if ($intuneDevs.Count -eq 1 -and $entraDevs.Count -eq 1) {
                    $iAadId = ConvertTo-AvdCleanupGuid -Value ([string](Get-AvdCleanupProp -InputObject $intuneDevs[0] -Name 'azureADDeviceId'))
                    $eDevId = ConvertTo-AvdCleanupGuid -Value ([string](Get-AvdCleanupProp -InputObject $entraDevs[0] -Name 'deviceId'))
                    if ($iAadId -ne $eDevId) { throw "Intune azureADDeviceId '$iAadId' does not match Entra deviceId '$eDevId' for '$short'." }
                }
                if ($intuneDevs.Count -eq 1) {
                    $intuneId = [string](Get-AvdCleanupProp -InputObject $intuneDevs[0] -Name 'id')
                    $null = Invoke-AvdCleanupRest -Method Delete -Uri ("https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/{0}" -f [uri]::EscapeDataString($intuneId)) -Headers $graphHeaders
                }
                if ($entraDevs.Count -eq 1) {
                    $eObjId = [string](Get-AvdCleanupProp -InputObject $entraDevs[0] -Name 'id')
                    $tType  = [string](Get-AvdCleanupProp -InputObject $entraDevs[0] -Name 'trustType')
                    if (-not [string]::Equals($tType, 'ServerAD', [System.StringComparison]::OrdinalIgnoreCase)) {
                        $eUri = "https://graph.microsoft.com/v1.0/devices/{0}" -f [uri]::EscapeDataString($eObjId)
                        if ($EntraCleanupAction -eq 'Disable') {
                            $null = Invoke-AvdCleanupRest -Method Patch -Uri $eUri -Headers $graphHeaders -Body (@{ accountEnabled = $false } | ConvertTo-Json -Compress)
                        } elseif ($EntraCleanupAction -eq 'Delete') {
                            $null = Invoke-AvdCleanupRest -Method Delete -Uri $eUri -Headers $graphHeaders
                        }
                    }
                }

                # 3. Delete VM, NICs, Disks, and AVD SessionHost
                $null = Invoke-AvdCleanupRest -Method Delete -Uri ("https://management.azure.com{0}?api-version=2023-09-01" -f $vmResId) -Headers $armHeaders
                foreach ($nic in @(Get-AvdCleanupProp -InputObject (Get-AvdCleanupProp -InputObject $vmProps -Name 'networkProfile') -Name 'networkInterfaces')) {
                    $nicId = [string](Get-AvdCleanupProp -InputObject $nic -Name 'id')
                    if (-not [string]::IsNullOrWhiteSpace($nicId)) {
                        try { $null = Invoke-AvdCleanupRest -Method Delete -Uri ("https://management.azure.com{0}?api-version=2023-04-01" -f $nicId) -Headers $armHeaders } catch { }
                    }
                }
                foreach ($d in $disks) {
                    try { $null = Invoke-AvdCleanupRest -Method Delete -Uri ("https://management.azure.com{0}?api-version=2023-04-02" -f $d.Id) -Headers $armHeaders } catch { }
                }
                $null = Invoke-AvdCleanupRest -Method Delete -Uri ("https://management.azure.com{0}/sessionHosts/{1}?api-version=2022-09-09&force=true" -f $poolId, [uri]::EscapeDataString($shLeaf)) -Headers $armHeaders
                $decomCount++

                [void]$results.Add((New-AvdCleanupResult -HostPoolName $poolName -SessionHostName $shLeaf -ShortHostName $short -VmId $vmId -AssignedUser $upn -LifecycleState 'Decommissioned' -Action 'SnapshottedAndDecommissioned' -Status 'Decommissioned' -SnapshotIds $snapIds.ToArray() -Detail "Created $($snapIds.Count) snapshot(s) and removed Intune, Entra ($EntraCleanupAction), VM, and SessionHost."))
            } else {
                [void]$results.Add((New-AvdCleanupResult -HostPoolName $poolName -SessionHostName $shLeaf -ShortHostName $short -VmId $vmId -AssignedUser $upn -LifecycleState 'Quarantined' -Action 'WhatIfSkipped' -Status 'EligibleForDecommission' -Detail "Skipped Phase 2 teardown (WhatIf)."))
            }
        }
    }
}

return $results.ToArray()
