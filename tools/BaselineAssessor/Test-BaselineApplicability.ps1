#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
$Tokens = $null
$ParseErrors = $null
$Ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Invoke-BaselineCollection.ps1'), [ref]$Tokens, [ref]$ParseErrors)
if ($ParseErrors.Count) { throw 'Collector parse errors' }
$BindParameters = [scriptblock]::Create($Ast.ParamBlock.Extent.Text + "`n`$AssessmentProfile")
$OutputAssignment = @($Ast.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.AssignmentStatementAst] -and $_.Left.Extent.Text -ceq '$output' })
$ProfileDeclaration = @($Ast.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.IfStatementAst] -and $_.Extent.Text.StartsWith("if (`$AssessmentProfile -ne 'Generic')") })
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
$PlatformFixtures = [Collections.Generic.List[object]]::new()
$AuditAssignment = @($Ast.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.AssignmentStatementAst] -and $_.Left.Extent.Text -ceq '$auditPolicy' })
if ($AuditAssignment.Count -ne 1) { throw 'Expected production audit collector' }
$AuditBlock = @($AuditAssignment[0].Right.PipelineElements[0].CommandElements | Where-Object { $_ -is [Management.Automation.Language.ScriptBlockExpressionAst] })
if ($AuditBlock.Count -ne 1) { throw 'Expected audit provider block' }
$AuditText = $AuditBlock[0].ScriptBlock.Extent.Text.Trim()
$ReadAudit = [scriptblock]::Create($AuditText.Substring(1, $AuditText.Length - 2))
function auditpol {
    if (($args -join ' ') -cne '/get /category:* /r') { throw 'Unexpected audit command' }
    'Machine Name,Policy Target,Subcategory,Subcategory GUID,Inclusion Setting,Exclusion Setting'
    "synthetic,System,Sensitive Privilege Use,{0cce9228-69ae-11d9-bed3-505054503030},$script:AuditValue,"
    "synthetic,System,Audit Policy Change,{0cce922f-69ae-11d9-bed3-505054503030},$script:AuditValue,"
}
try {
foreach ($Case in @(
    @{ Selection = 'Windows11_25H2'; Id = 'windows-11-25h2'; Build = '26200'; ProductType = 1; Server = $false; DC = $false; Audit = 'Success' },
    @{ Selection = 'WindowsServer2025Member'; Id = 'windows-server-2025-member'; Build = '26100'; ProductType = 3; Server = $true; DC = $false; Audit = 'Success and Failure' },
    @{ Selection = 'WindowsServer2025DC'; Id = 'windows-server-2025-dc'; Build = '26100'; ProductType = 2; Server = $true; DC = $true; Audit = 'Success and Failure' }
)) {
    $systemInfo = @{ hostname = $Case.Id; osBuild = $Case.Build; productType = $Case.ProductType; isServer = $Case.Server; isDomainController = $Case.DC }
    $script:AuditValue = $Case.Audit
    $auditPolicy = & $ReadAudit
    if ($auditPolicy['0cce9228-69ae-11d9-bed3-505054503030'] -cne $Case.Audit -or $auditPolicy['0cce922f-69ae-11d9-bed3-505054503030'] -cne $Case.Audit) { throw 'Production GUID serialization changed' }
    foreach ($Selection in @($Case.Selection, $Case.Selection.ToLowerInvariant())) {
        $AssessmentProfile = & $BindParameters -AssessmentProfile $Selection
        $Result = (& $BuildOutput | ConvertTo-Json -Depth 10) | ConvertFrom-Json
        if ($Result.assessmentContext.profileId -cne $Case.Id -or $Result.assessmentContext.profileVersion -cne '2026-09-18' -or $Result.assessmentContext.selectionSource -cne 'Operator') { throw 'Platform profile declaration changed' }
        if ($Result.systemInfo.osBuild -cne $Case.Build -or $Result.systemInfo.productType -ne $Case.ProductType -or $Result.auditPolicy.'Audit Policy Change' -cne $Case.Audit) { throw 'Profile changed observed platform/audit evidence' }
    }
    $PlatformFixtures.Add($Result)
}
} finally { Remove-Item Function:\auditpol }
if ($env:ASSAY_BASELINE_PLATFORM_FIXTURES) {
    [IO.File]::WriteAllText($env:ASSAY_BASELINE_PLATFORM_FIXTURES, (ConvertTo-Json -InputObject $PlatformFixtures.ToArray() -Depth 10), [Text.UTF8Encoding]::new($false))
}
Write-Output 'PASS: explicit versioned Cloud PC/client/member/DC declarations, generic default, parameter rejection and production JSON export; no collection or endpoint commands executed.'