#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Get-IntuneDefenderEvidence.ps1') -LibraryOnly
function Assert-Defender([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
foreach ($Address in @('https://graph.microsoft.com/api/machines', 'https://api.security.microsoft.com/api/machines/id/isolate', 'http://api.security.microsoft.com/api/machines')) {
    Assert-Defender (-not (Test-IntuneDefenderUri $Address)) 'Unsafe endpoint accepted'
}
Assert-Defender (Test-IntuneDefenderUri 'https://api.security.microsoft.com/api/machines?$top=1000&$skip=1000') 'Documented paging blocked'
foreach ($Query in @('$top=1000&$filter=healthStatus%20eq%20%27Active%27', '$top=1000&%24expand=alerts', '$top=1000&%24top=1', '$top=1000&$skip=-1', '$top=1000&$skip=100001', '$top=1', '$top=1000&$skiptoken=opaque', '$top=1000&unknown=value')) {
    Assert-Defender (-not (Test-IntuneDefenderUri ('https://api.security.microsoft.com/api/machines?' + $Query))) 'Unexpected Defender query accepted'
}
$script:PagingRequests = @()
$Result = Read-IntuneDefenderMachines -Request {
    param($Address)
    $script:PagingRequests += $Address
    if ($script:PagingRequests.Count -eq 1) {
        $Rows = @(foreach ($Index in 1..1000) { [pscustomobject]@{ id = "machine-$Index" } })
        return @{ StatusCode = 200; Body = [pscustomobject]@{ value = $Rows } }
    }
    return @{ StatusCode = 200; Body = [pscustomobject]@{ value = @() } }
}
Assert-Defender ($Result.State -eq 'Complete' -and $Result.Rows.Count -eq 1000 -and $script:PagingRequests[1] -ceq 'https://api.security.microsoft.com/api/machines?$top=1000&$skip=1000') 'Documented skip pagination failed'
$script:PagingRequests = @()
$Result = Read-IntuneDefenderMachines -Request {
    param($Address)
    $script:PagingRequests += $Address
    @{ StatusCode = 200; Body = [pscustomobject]@{ value = @([pscustomobject]@{ id = 'machine' }); '@odata.nextLink' = 'https://api.security.microsoft.com/api/machines?$top=1000&$skip=900' } }
}
Assert-Defender ($Result.State -eq 'Partial' -and $Result.ErrorCode -eq 'InvalidContinuation' -and $script:PagingRequests.Count -eq 1) 'Continuation skipped unseen records'
$Result = Read-IntuneDefenderMachines -Request { @{ StatusCode = 200; Body = ('{"value":[{"id":"mde","aadDeviceId":"entra","healthStatus":"Active","lastSeen":"2026-09-17T08:00:00Z","lastIpAddress":"SECRET"}]}' | ConvertFrom-Json) } }
Assert-Defender ($Result.State -eq 'Complete' -and $Result.Rows.Count -eq 1) 'Machine collection failed'
Assert-Defender ([datetime]$Result.Rows[0].lastSeen -eq [datetime]'2026-09-17T08:00:00Z') 'Defender timestamp lost after JSON parsing'
Assert-Defender (-not (($Result | ConvertTo-Json -Depth 10).Contains('SECRET'))) 'Unnecessary machine data leaked'
$Result = Read-IntuneDefenderMachines -Request { @{ StatusCode = 200; Body = ('{"value":[],"@odata.nextLink":"https://example.com/secret"}' | ConvertFrom-Json) } }
Assert-Defender ($Result.State -eq 'Partial' -and $Result.ErrorCode -eq 'BlockedEndpoint') 'Cross-host paging accepted'
$Result = Read-IntuneDefenderMachines -Request { @{ StatusCode = 404 } }
Assert-Defender ($Result.State -eq 'Error') '404 incorrectly inferred clean absence'
$script:Attempts = 0
$Result = Read-IntuneDefenderMachines -Request { $script:Attempts++; @{ StatusCode = 429 } } -Delay { param($Seconds) }
Assert-Defender ($script:Attempts -eq 4 -and $Result.State -eq 'Error') 'Retry cap failed'
$script:Delays = @()
$Result = Read-IntuneDefenderMachines -Request { @{ StatusCode = 429; RetryAfter = 7 } } -Delay { param($Seconds) $script:Delays += $Seconds }
Assert-Defender ($script:Delays.Count -eq 3 -and @($script:Delays | Where-Object { $_ -ne 7 }).Count -eq 0) 'Retry-After ignored'
Write-Output 'PASS: Defender read-only host boundary, metadata projection, error states and bounded retries.'