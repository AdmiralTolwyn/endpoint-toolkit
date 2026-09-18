#Requires -Version 5.1
[CmdletBinding()]
param([Parameter(Mandatory)][string]$EvidenceDirectory)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Invoke-IntuneDiscovery.ps1') -LibraryOnly
$Contracts = Get-Content (Join-Path $PSScriptRoot 'PolicyContracts.json') -Raw | ConvertFrom-Json
$Results = [Collections.Generic.List[object]]::new()
foreach ($Contract in $Contracts) {
    $Source = 'https://learn.microsoft.com/en-us/windows/client-management/mdm/' + $Contract.doc
    $File = Join-Path $EvidenceDirectory ('csp-' + $Contract.doc + '.md')
    if (-not (Test-Path -LiteralPath $File)) { Invoke-WebRequest -UseBasicParsing -Uri ($Source + '?accept=text/markdown') -OutFile $File -TimeoutSec 60 }
    $Text = (Get-Content -LiteralPath $File -Raw).Replace('\_', '_')
    $Hash = (Get-FileHash -LiteralPath $File -Algorithm SHA256).Hash
    foreach ($Node in $Contract.nodes) {
        $Path = $Contract.prefix + $Node
        $Pattern = '\./(?:Device/)?Vendor/MSFT/' + [regex]::Escape($Path) + '(?![A-Za-z0-9_/])'
        $Match = [regex]::Match($Text, $Pattern)
        $Format = $null
        if ($Match.Success) {
            $Tail = $Text.Substring($Match.Index)
            $NextSection = [regex]::Match($Tail, '(?m)^## ')
            if ($NextSection.Success) { $Tail = $Tail.Substring(0, $NextSection.Index) }
            $FormatMatch = [regex]::Match($Tail, '(?im)^\|\s*Format\s*\|\s*([^|\r\n]+)')
            if ($FormatMatch.Success) { $Format = $FormatMatch.Groups[1].Value.Trim() }
        }
        $Allowed = (Test-IntuneSecurityPath $Path) -or (Test-IntuneExtendedPath $Path)
        $Results.Add(@{ Path = $Path; Source = $Source; Sha256 = $Hash; DocumentedUri = $Match.Success; Format = $Format; CollectorAllowed = $Allowed })
    }
}
$Output = Join-Path $EvidenceDirectory 'intune-policy-contract-audit.json'
[IO.File]::WriteAllText($Output, (@{ CheckedAtUtc = [datetime]::UtcNow.ToString('o'); Nodes = @($Results.ToArray()); Limit = 'URI and published format evidence only; not recommendation values or effective policy.' } | ConvertTo-Json -Depth 10), [Text.UTF8Encoding]::new($false))
$Failures = @($Results | Where-Object { -not $_.DocumentedUri -or -not $_.CollectorAllowed -or -not $_.Format })
foreach ($Failure in $Failures) { Write-Output ($Failure | ConvertTo-Json -Compress) }
Write-Output ('Audited {0} fixed or templated CSP paths; {1} require review. Evidence: {2}' -f $Results.Count, $Failures.Count, $Output)
if ($Failures.Count) { throw 'CSP provenance incomplete; review the report' }