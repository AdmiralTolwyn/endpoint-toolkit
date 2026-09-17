#Requires -Version 5.1
[CmdletBinding()]
param([string]$TenantId, [securestring]$AccessToken, [string]$OutputPath, [switch]$LibraryOnly)

function Test-IntuneDefenderUri {
    param([string]$Address)
    $Parsed = $null
    return [uri]::TryCreate($Address, [UriKind]::Absolute, [ref]$Parsed) -and $Parsed.Scheme -eq 'https' -and $Parsed.Host -eq 'api.security.microsoft.com' -and $Parsed.IsDefaultPort -and -not $Parsed.UserInfo -and -not $Parsed.Fragment -and $Parsed.AbsolutePath -ceq '/api/machines'
}

function Read-IntuneDefenderMachines {
    param([scriptblock]$Request, [scriptblock]$Delay = { param($Seconds) [Threading.Tasks.Task]::Delay([timespan]::FromSeconds($Seconds)).GetAwaiter().GetResult() })
    $Next = 'https://api.security.microsoft.com/api/machines?$top=1000'
    $Visited = [Collections.Generic.HashSet[string]]::new()
    $Ids = [Collections.Generic.HashSet[string]]::new()
    $Rows = [Collections.Generic.List[object]]::new()
    $Failure = $null
    $Pages = 0
    $Deadline = [datetime]::UtcNow.AddMinutes(30)
    while ($Next) {
        if (-not (Test-IntuneDefenderUri $Next)) { $Failure = 'BlockedEndpoint'; break }
        if (-not $Visited.Add($Next)) { $Failure = 'PaginationCycle'; break }
        if ($Pages -ge 1000 -or [datetime]::UtcNow -ge $Deadline) { $Failure = 'CollectionBudget'; break }
        $Response = $null
        for ($Attempt = 0; $Attempt -lt 4; $Attempt++) {
            try { $Response = & $Request $Next } catch { $Response = @{ StatusCode = 0 } }
            if ($Response.StatusCode -notin @(429, 503, 504) -or $Attempt -eq 3) { break }
            $Seconds = [double]$Response.RetryAfter
            if ($Seconds -le 0 -or [double]::IsNaN($Seconds) -or [double]::IsInfinity($Seconds)) { $Seconds = [math]::Pow(2, $Attempt + 1) }
            if ($Seconds -gt 1800) { break }
            if ([datetime]::UtcNow.AddSeconds($Seconds) -ge $Deadline) { break }
            & $Delay $Seconds
        }
        if ($Response.StatusCode -ne 200) { $Failure = 'HTTP' + $Response.StatusCode; break }
        if ($null -eq $Response.Body -or $null -eq $Response.Body.PSObject.Properties['value'] -or $Response.Body.value -isnot [array] -or $null -ne $Response.Body.PSObject.Properties['error']) { $Failure = 'InvalidResponse'; break }
        $Pages++
        foreach ($Row in $Response.Body.value) {
            if ($Rows.Count -ge 100000) { $Failure = 'RowLimit'; break }
            if ($null -eq $Row.PSObject.Properties['id'] -or -not $Row.id -or -not $Ids.Add([string]$Row.id)) { $Failure = 'InvalidIdentity'; break }
            $Safe = [ordered]@{}
            foreach ($Field in @('id', 'aadDeviceId', 'onboardingStatus', 'healthStatus', 'firstSeen', 'lastSeen', 'osPlatform', 'osBuild', 'version')) {
                if ($null -ne $Row.PSObject.Properties[$Field]) {
                    $Value = $Row.$Field
                    if (($Value -is [string] -and $Value.Length -le 256) -or $Value -is [int] -or $Value -is [long]) { $Safe[$Field] = $Value }
                    elseif ($Value -is [datetime] -or $Value -is [datetimeoffset]) { $Safe[$Field] = $Value.ToUniversalTime().ToString('o') }
                }
            }
            $Rows.Add($Safe)
        }
        if ($Failure) { break }
        $Next = $null
        if ($null -ne $Response.Body.PSObject.Properties['@odata.nextLink']) {
            $Next = $Response.Body.'@odata.nextLink'
            if ($Next -isnot [string] -or -not $Next) { $Failure = 'InvalidNextLink'; break }
        } elseif ($Response.Body.value.Count -eq 1000) {
            $Next = 'https://api.security.microsoft.com/api/machines?$top=1000&$skip=' + $Rows.Count
        }
    }
    return @{ Rows = @($Rows.ToArray()); State = $(if (-not $Failure) { 'Complete' } elseif ($Pages) { 'Partial' } else { 'Error' }); PagesRead = $Pages; ErrorCode = $Failure }
}

if ($LibraryOnly) { return }
$ErrorActionPreference = 'Stop'
if (-not $AccessToken -or -not $OutputPath -or (Test-Path -LiteralPath $OutputPath)) { throw 'Supply a SecureString Defender token and a new output path.' }
$SelectedTenant = [guid]::Empty
if (-not [guid]::TryParse($TenantId, [ref]$SelectedTenant) -or $SelectedTenant -eq [guid]::Empty) { throw 'Provide an explicit tenant GUID.' }
$Client = $null
$Handler = $null
$PlainToken = $null
try {
    $PlainToken = [Net.NetworkCredential]::new('', $AccessToken).Password
    $Parts = $PlainToken.Split('.')
    if ($Parts.Count -ne 3) { throw 'Unsupported token metadata.' }
    $Payload = $Parts[1].Replace('-', '+').Replace('_', '/')
    $Payload = $Payload.PadRight($Payload.Length + ((4 - $Payload.Length % 4) % 4), '=')
    try { $Claims = [Text.Encoding]::UTF8.GetString([convert]::FromBase64String($Payload)) | ConvertFrom-Json } catch { throw 'Invalid token metadata.' }
    if ($Claims.tid -ne $TenantId -or $Claims.aud -notin @('https://api.security.microsoft.com', 'https://api.securitycenter.microsoft.com') -or $Claims.exp -le [datetimeoffset]::UtcNow.ToUnixTimeSeconds() -or 'Machine.Read' -notin ([string]$Claims.scp).Split(' ')) { throw 'Expected an unexpired delegated Defender Machine.Read token for the selected tenant. No permissions were granted.' }
    Add-Type -AssemblyName System.Net.Http
    $Handler = [Net.Http.HttpClientHandler]::new()
    $Handler.AllowAutoRedirect = $false
    $Client = [Net.Http.HttpClient]::new($Handler)
    $Client.Timeout = [timespan]::FromSeconds(120)
    $Client.MaxResponseContentBufferSize = 16MB
    $Client.DefaultRequestHeaders.Authorization = [Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer', $PlainToken)
    $PlainToken = $null
    $Started = [datetime]::UtcNow.ToString('o')
    $Result = Read-IntuneDefenderMachines -Request {
        param($Address)
        if (-not (Test-IntuneDefenderUri $Address)) { throw 'Blocked endpoint' }
        $Response = $Client.GetAsync($Address).GetAwaiter().GetResult()
        try {
            $Body = $null
            if ([int]$Response.StatusCode -eq 200) { $Body = $Response.Content.ReadAsStringAsync().GetAwaiter().GetResult() | ConvertFrom-Json }
            $RetryAfter = 0
            if ($Response.Headers.RetryAfter) {
                if ($Response.Headers.RetryAfter.Delta) { $RetryAfter = [math]::Ceiling($Response.Headers.RetryAfter.Delta.TotalSeconds) }
                elseif ($Response.Headers.RetryAfter.Date) { $RetryAfter = [math]::Max(0, [math]::Ceiling(($Response.Headers.RetryAfter.Date - [datetimeoffset]::UtcNow).TotalSeconds)) }
            }
            return @{ StatusCode = [int]$Response.StatusCode; Body = $Body; RetryAfter = $RetryAfter }
        } finally { $Response.Dispose() }
    }
    $Document = @{ SchemaVersion = '1.0'; TenantId = $TenantId; StartedAtUtc = $Started; CollectedAtUtc = [datetime]::UtcNow.ToString('o'); Result = $Result }
    $Bytes = [Text.UTF8Encoding]::new($false).GetBytes(($Document | ConvertTo-Json -Depth 10))
    if ($Bytes.Length -gt 64MB) { throw 'Defender export exceeds 64 MB.' }
    $Stream = [IO.File]::Open([IO.Path]::GetFullPath($OutputPath), [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    try { $Stream.Write($Bytes, 0, $Bytes.Length) } finally { $Stream.Dispose() }
    Write-Output ('Defender evidence exported: ' + $Result.State + '. Device-group visibility and retention limit coverage.')
} finally {
    $PlainToken = $null
    if ($null -ne $Client) { $Client.Dispose() }
    if ($null -ne $Handler) { $Handler.Dispose() }
}