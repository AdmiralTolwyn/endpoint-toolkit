#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Get-IntuneDefenderEvidence.ps1') -LibraryOnly
function Assert-Defender([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
foreach ($Address in @('https://graph.microsoft.com/api/machines', 'https://api.security.microsoft.com/api/machines/id/isolate', 'http://api.security.microsoft.com/api/machines')) {
    Assert-Defender (-not (Test-IntuneDefenderUri $Address)) 'Unsafe endpoint accepted'
}
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