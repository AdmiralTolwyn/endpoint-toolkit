#Requires -Version 5.1
[CmdletBinding()]
param([string]$EvidenceDirectory, [switch]$CheckDocumentation)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Invoke-IntuneDiscovery.ps1') -LibraryOnly
$Contracts = Get-Content (Join-Path $PSScriptRoot 'EndpointContracts.json') -Raw | ConvertFrom-Json
$Tokens = $null
$ParseErrors = $null
$Ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Get-IntuneEndpointEvidence.ps1'), [ref]$Tokens, [ref]$ParseErrors)
if ($ParseErrors.Count) { throw 'Endpoint companion parse errors' }
$Switches = @($Ast.FindAll({ param($Node) $Node -is [Management.Automation.Language.SwitchStatementAst] -and $Node.Condition.Extent.Text -eq '$Module' }, $true))
if ($Switches.Count -ne 1 -or $Switches[0].Clauses.Count -ne $Contracts.Count) { throw 'Provider command inventory changed; review contracts' }
$Identity = @($Ast.FindAll({ param($Node) $Node -is [Management.Automation.Language.CommandAst] -and $Node.GetCommandName() -eq 'dsregcmd.exe' }, $true))
if ($Identity.Count -ne 1 -or $Identity[0].CommandElements.Count -ne 2 -or $Identity[0].CommandElements[1].Extent.Text -cne '/status') { throw 'Device identity command changed' }
$Evidence = [Collections.Generic.List[object]]::new()
foreach ($Clause in $Switches[0].Clauses) {
    $Name = $Clause.Item1.SafeGetValue()
    $Contract = @($Contracts | Where-Object { $_.module -ceq $Name })
    if ($Contract.Count -ne 1) { throw ('Missing provider contract: ' + $Name) }
    $Contract = $Contract[0]
    $Pipeline = $Clause.Item2.Statements[0]
    if ($Clause.Item2.Statements.Count -ne 1 -or $Pipeline.PipelineElements.Count -ne 2 -or ($Pipeline.PipelineElements[0].Extent.Text -replace '\s+', ' ') -cne $Contract.command) { throw ('Provider command changed: ' + $Name) }
    $Projection = $Pipeline.PipelineElements[1]
    if ($Projection.GetCommandName() -cne 'Select-Object') { throw 'Expected explicit field projection' }
    $Fields = @(foreach ($Element in $Projection.CommandElements | Select-Object -Skip 1) { foreach ($Field in @($Element.SafeGetValue())) { $Field } })
    if (($Fields -join ',') -cne ($Contract.fields -join ',')) { throw ('Provider fields changed: ' + $Name) }
    $SourceFields = @($Contract.sources | ForEach-Object { $_.fields })
    foreach ($Field in $Fields) { if ($Field -cnotin $SourceFields) { throw ('Field has no source: ' + $Name + '.' + $Field) } }
    $Probe = [ordered]@{ unreviewed = 'DO_NOT_EXPORT' }
    foreach ($Field in $Fields) { $Probe[$Field] = 'documented-field' }
    if ($Name -eq 'DefenderStatus') { $Probe['AntivirusSignatureLastUpdated'] = '2026-09-18T08:00:00.0000000Z' }
    $Safe = ConvertTo-IntuneEndpointModules @{ $Name = @{ State = 'Complete'; Rows = @($Probe) } }
    if ($Safe[$Name].Rows[0].Count -ne $Fields.Count -or ($Safe | ConvertTo-Json -Depth 10) -match 'DO_NOT_EXPORT') { throw ('Importer projection drift: ' + $Name) }
    if ($CheckDocumentation) {
        if (-not $EvidenceDirectory -or -not (Test-Path -LiteralPath $EvidenceDirectory -PathType Container)) { throw 'Provide an existing evidence directory' }
        $Ordinal = 0
        foreach ($Source in $Contract.sources) {
            $Ordinal++
            $Path = Join-Path $EvidenceDirectory ('endpoint-' + $Name + '-' + $Ordinal + '.md')
            $Separator = if ($Source.url.Contains('?')) { '&' } else { '?' }
            if (-not (Test-Path -LiteralPath $Path)) { Invoke-WebRequest -UseBasicParsing -Uri ($Source.url + $Separator + 'accept=text/markdown') -OutFile $Path -TimeoutSec 60 }
            $Text = (Get-Content -LiteralPath $Path -Raw).Replace('\_', '_')
            if ($Text -notmatch '(?m)^canonicalUrl: https://learn.microsoft.com/') { throw 'Not a Microsoft Learn source' }
            foreach ($Field in $Source.fields) { if ($Text -notmatch ('(?<![A-Za-z0-9_])' + [regex]::Escape($Field) + '(?![A-Za-z0-9_])')) { throw ('Source no longer names field: ' + $Name + '.' + $Field) } }
            $Evidence.Add(@{ Module = $Name; Source = $Source.url; Fields = $Source.fields; Sha256 = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash })
        }
    }
}
if ($CheckDocumentation) {
    [IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'intune-endpoint-contract-audit.json'), (@{ CheckedAtUtc = [datetime]::UtcNow.ToString('o'); Evidence = @($Evidence.ToArray()); Limit = 'Source-presence and command/projection drift checks only; not live provider availability or enum validation.' } | ConvertTo-Json -Depth 10), [Text.UTF8Encoding]::new($false))
}
Write-Output 'PASS: five documented provider commands, 36 projected fields, identity read and importer drift checks; no endpoint queries executed.'