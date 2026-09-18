#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
$Tokens = $null
$ParseErrors = $null
$Ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Invoke-BaselineCollection.ps1'), [ref]$Tokens, [ref]$ParseErrors)
if ($ParseErrors.Count) { throw 'Collector parse errors' }
$BindParameters = [scriptblock]::Create($Ast.ParamBlock.Extent.Text + "`n`$AssessmentProfile")
$OutputAssignment = @($Ast.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.AssignmentStatementAst] -and $_.Left.Extent.Text -ceq '$output' })
$ProfileDeclaration = @($Ast.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.IfStatementAst] -and $_.Extent.Text.StartsWith("if (`$AssessmentProfile -eq 'Windows365CloudPc')") })
if ($OutputAssignment.Count -ne 1 -or $ProfileDeclaration.Count -ne 1) { throw 'Expected production export and profile declaration' }
$BuildOutput = [scriptblock]::Create($OutputAssignment[0].Extent.Text + "`n" + $ProfileDeclaration[0].Extent.Text + "`n`$output")
$Script:StartTime = [datetime]::UtcNow
$Script:CollectorVersion = 'synthetic'
$Script:Errors = @()
$Script:TotalAreas = 0
$systemInfo = @{ hostname = 'synthetic-cloud-pc'; isServer = $false; isDomainController = $false; productType = 1 }
$bitlocker = @{ osProtectionStatus = 0 }
$AssessmentProfile = & $BindParameters
if ($AssessmentProfile -cne 'Generic') { throw 'Default profile changed' }
$Generic = & $BuildOutput
if ($Generic.Contains('assessmentContext')) { throw 'Default profile implicitly declared a Cloud PC' }
foreach ($Selection in @('Generic','generic')) {
    $AssessmentProfile = & $BindParameters -AssessmentProfile $Selection
    $Result = & $BuildOutput
    if ($Result.Contains('assessmentContext')) { throw 'Generic selection emitted Cloud PC exceptions' }
}
foreach ($Selection in @('Windows365CloudPc','windows365cloudpc')) {
    $AssessmentProfile = & $BindParameters -AssessmentProfile $Selection
    $Result = (& $BuildOutput | ConvertTo-Json -Depth 10) | ConvertFrom-Json
    $Context = $Result.assessmentContext
    if ($Context.schemaVersion -cne '1.0' -or $Context.profileId -cne 'windows365-cloud-pc' -or $Context.profileVersion -cne '2026-09-18' -or $Context.selectionSource -cne 'Operator' -or @($Context.PSObject.Properties).Count -ne 4) { throw 'Profile contract changed or lost through export' }
    if ($Result.systemInfo.isServer -ne $false -or $Result.bitlocker.osProtectionStatus -ne 0 -or $Result._metadata.parameters.assessmentProfile -cne $Selection) { throw 'Profile changed observed evidence or lost parameter provenance' }
}
foreach ($Selection in @('Unknown','WindowsServer','', ' Windows365CloudPc')) {
    $Rejected = $false
    try { & $BindParameters -AssessmentProfile $Selection | Out-Null } catch { $Rejected = $true }
    if (-not $Rejected) { throw 'Unreviewed profile parameter accepted' }
}
if ($env:ASSAY_BASELINE_APPLICABILITY_FIXTURE) {
    [IO.File]::WriteAllText($env:ASSAY_BASELINE_APPLICABILITY_FIXTURE, ($Result | ConvertTo-Json -Depth 10), [Text.UTF8Encoding]::new($false))
}
Write-Output 'PASS: explicit versioned Cloud PC declaration, generic default, parameter rejection and production JSON export; no collection or endpoint commands executed.'