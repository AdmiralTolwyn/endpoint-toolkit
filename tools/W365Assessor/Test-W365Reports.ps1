#Requires -Version 5.1
[CmdletBinding()]
param([string]$MetadataPath)
$ErrorActionPreference = 'Stop'
$Tokens = $null
$ParseErrors = $null
$Ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Invoke-W365Discovery.ps1'), [ref]$Tokens, [ref]$ParseErrors)
if ($ParseErrors.Count) { throw 'Collector parse failure' }
foreach ($Name in @('New-CheckResult','Invoke-GraphReport')) {
    $Definition = $Ast.Find({ param($Node) $Node -is [Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -ceq $Name }, $true)
    if ($Definition) { . ([scriptblock]::Create($Definition.Extent.Text)) }
}
function Write-Status { param($Message, $Level) }
function Invoke-MgGraphRequest {
    param($Method, $Uri, $Body, $ContentType)
    $script:Requests.Add(@{ Method = $Method; Uri = $Uri; Body = ($Body | ConvertFrom-Json); ContentType = $ContentType })
    if ($script:ReportFails) { throw 'Synthetic report failure' }
    return ,$script:ReportResponse
}
$script:Requests = [Collections.Generic.List[object]]::new()
$script:ReportFails = $false
$CpcNoLoginCount = 1
$ScriptRoot = $PSScriptRoot
$Discovery = [pscustomobject]@{ Inventory = [ordered]@{} }
$Blocks = @{}
foreach ($Action in @('retrieveConnectionQualityReports','retrieveCloudPcTenantMetricsReport')) {
    $Needle = "Invoke-GraphReport -Action '$Action'"
    $Found = @($Ast.FindAll({ param($Node) $Node -is [Management.Automation.Language.TryStatementAst] -and $Node.Body.Extent.Text.Contains($Needle) }, $true))
    if ($Found.Count -ne 1) { throw ('Expected one report block: ' + $Action) }
    $Blocks[$Action] = [scriptblock]::Create($Found[0].Extent.Text)
}
foreach ($Response in @(@{},@{ TotalRowCount = '0' },@{ TotalRowCount = -1 },@{ Values = 'not rows' })) {
    $script:ReportResponse = $Response
    foreach ($Action in @('retrieveConnectionQualityReports','retrieveCloudPcTenantMetricsReport')) {
        $AllChecks = [Collections.Generic.List[object]]::new()
        & $Blocks[$Action]
        if ($AllChecks.Count -eq 0 -or @($AllChecks | Where-Object Status -ne 'Error').Count) { throw ('Malformed report produced a verdict: ' + $Action) }
    }
}
. (Join-Path $ScriptRoot 'W365Reports.ps1')
function Assert-Report([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function New-TestReport([string]$Action, [int]$Total = 1) {
    $Contract = Get-W365ReportContract $Action
    $Schema = @($Contract.Columns.Keys | ForEach-Object { @{ Column = $_; PropertyType = $Contract.Columns[$_] } })
    $Rows = [Collections.Generic.List[object]]::new()
    $Rows.Add(@('PRIVATE_CELL','25.5','10'))
    @{ TotalRowCount = $Total; Schema = $Schema; Values = $Rows.ToArray() }
}
foreach ($Action in @('retrieveConnectionQualityReports','retrieveCloudPcTenantMetricsReport')) {
    $Contract = Get-W365ReportContract $Action
    $Report = New-TestReport $Action 40
    foreach ($Form in @('object','pscustom','string','bytes','stream')) {
        $Json = $Report | ConvertTo-Json -Depth 10
        $script:ReportResponse = switch ($Form) {
            'object' { $Report }
            'pscustom' { $Json | ConvertFrom-Json }
            'string' { $Json }
            'bytes' { ,([Text.Encoding]::UTF8.GetBytes($Json)) }
            'stream' { [IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes($Json)) }
        }
        $AllChecks.Clear()
        & $Blocks[$Action]
        $Evidence = $AllChecks[0].Evidence
        Assert-Report ($Evidence.RowsReturned -eq 1 -and $Evidence.ReportedTotalRowCount -eq 40 -and $Evidence.PageCoverage -eq 'PartialPageSet' -and $Evidence.AssessmentState -eq 'ReportContentNotEvaluated') ('Production decode failed: ' + $Form)
        Assert-Report (@($AllChecks | Where-Object Status -ne 'Error').Count -eq 0) 'Valid report yielded a posture verdict'
        Assert-Report (($AllChecks | ConvertTo-Json -Depth 10) -notmatch 'PRIVATE_CELL') 'Raw report cells exported'
        $Request = $script:Requests[-1]
        Assert-Report ($Request.Method -ceq 'POST' -and $Request.Uri -ceq ('https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/reports/' + $Action) -and $Request.ContentType -ceq 'application/json') 'Unexpected report transport'
        Assert-Report ($Request.Body.reportName -ceq $Contract.ReportName -and $Request.Body.top -eq 25 -and $Request.Body.skip -eq 0 -and @($Request.Body.PSObject.Properties).Count -eq 4 -and (@($Request.Body.select) -join ',') -ceq (@($Contract.Columns.Keys) -join ',')) 'Unexpected report request body'
        if ($Form -eq 'stream') { $script:ReportResponse.Dispose() }
    }
    $Report.Values = @()
    $Report.TotalRowCount = 0
    $Evidence = ConvertTo-W365ReportPageEvidence $Report $Action
    Assert-Report ($Evidence.RowsReturned -eq 0 -and $Evidence.PageCoverage -eq 'CompletePageSet') 'Empty report metadata invalid'
    foreach ($Invalid in @($null,'0',-1,1.5,$true,[long]2147483648)) {
        $Report.TotalRowCount = $Invalid
        $Rejected = $false
        try { ConvertTo-W365ReportPageEvidence $Report $Action | Out-Null } catch { $Rejected = $true }
        Assert-Report $Rejected 'Invalid count accepted'
    }
    foreach ($Defect in @('missing-schema','wrong-type','unknown-column','duplicate-column','wrong-width','nested-cell','too-many-rows','total-less-than-rows','error-object')) {
        $Bad = New-TestReport $Action
        switch ($Defect) {
            'missing-schema' { $Bad.Remove('Schema') }
            'wrong-type' { $Bad.Schema[0].PropertyType = 'Json' }
            'unknown-column' { $Bad.Schema[0].Column = 'Unreviewed' }
            'duplicate-column' { $Bad.Schema[1] = $Bad.Schema[0] }
            'wrong-width' { $Bad.Values = @(,@('too-short')) }
            'nested-cell' { $Bad.Values = @(,@(@{ secret = 'not scalar' },1,2)) }
            'too-many-rows' { $Bad.TotalRowCount = 26; $Bad.Values = @((1..26) | ForEach-Object { ,@('row',1,2) }) }
            'total-less-than-rows' { $Bad.TotalRowCount = 0 }
            'error-object' { $Bad.error = @{ code = 'failed' } }
        }
        $Rejected = $false
        try { ConvertTo-W365ReportPageEvidence $Bad $Action | Out-Null } catch { $Rejected = $true }
        Assert-Report $Rejected ('Malformed report accepted: ' + $Defect)
    }
}
$Rejected = $false
try { ConvertTo-W365ReportPageEvidence ('x' * (1MB + 1)) retrieveConnectionQualityReports | Out-Null } catch { $Rejected = $true }
Assert-Report $Rejected 'Oversized report text accepted'
foreach ($InputBody in @([byte[]]::new(1MB + 1),[IO.MemoryStream]::new([byte[]]::new(1MB + 1)),[byte[]]@(255,255),'not JSON')) {
    $Rejected = $false
    try { ConvertTo-W365ReportPageEvidence $InputBody retrieveConnectionQualityReports | Out-Null } catch { $Rejected = $true }
    finally { if ($InputBody -is [IO.Stream]) { $InputBody.Dispose() } }
    Assert-Report $Rejected 'Oversized or malformed report body accepted'
}
$script:Requests.Clear()
foreach ($Action in @('retrieveCloudPcRecommendationReports','../cloudPCs/reboot','unknownAction')) {
    $Rejected = $false
    try { Invoke-GraphReport $Action @{ top = 25 } | Out-Null } catch { $Rejected = $true }
    Assert-Report ($Rejected -and $script:Requests.Count -eq 0) 'Unsupported action reached SDK'
}
foreach ($Body in @(@{},@{ top = '25' },@{ top = 26 },@{ top = 25; filter = 'unreviewed' })) {
    $Rejected = $false
    try { Invoke-GraphReport retrieveConnectionQualityReports $Body | Out-Null } catch { $Rejected = $true }
    Assert-Report ($Rejected -and $script:Requests.Count -eq 0) 'Unsupported request options reached SDK'
}
$RecommendationStatements = @($Ast.EndBlock.Statements | Where-Object { $_.Extent.Text.Contains("RecommendationReportState") -or $_.Extent.Text.Contains("foreach (`$eid in @('W365-COST-002','W365-MON-008-REC'))") -or $_.Extent.Text.Contains("-Id 'W365-COST-001-REPORT'") })
Assert-Report ($RecommendationStatements.Count -eq 3) 'Expected explicit recommendation quarantine statements'
$AllChecks.Clear()
foreach ($Statement in $RecommendationStatements) { & ([scriptblock]::Create($Statement.Extent.Text)) }
Assert-Report ($script:Requests.Count -eq 0 -and $AllChecks.Count -eq 3 -and $Discovery.Inventory.RecommendationReportState -eq 'Unsupported') 'Retired recommendation path was not quarantined'
Assert-Report ($AllChecks[0].Category -ceq 'Cost & Optimization' -and $AllChecks[1].Category -ceq 'Monitoring & Diagnostics') 'Recommendation finding categories changed'
$script:ReportResponse = New-TestReport retrieveConnectionQualityReports 40
& $Blocks['retrieveConnectionQualityReports']
$script:ReportResponse = New-TestReport retrieveCloudPcTenantMetricsReport 1
& $Blocks['retrieveCloudPcTenantMetricsReport']
$ExportedChecks = @($AllChecks.ToArray())
$script:ReportFails = $true
foreach ($Action in $Blocks.Keys) {
    $AllChecks.Clear()
    & $Blocks[$Action]
    Assert-Report ($AllChecks.Count -gt 0 -and @($AllChecks | Where-Object Status -ne 'Error').Count -eq 0) 'SDK failure did not stay unassessed'
}
if ($MetadataPath) {
    [xml]$Metadata = Get-Content -LiteralPath $MetadataPath -Raw
    foreach ($Action in $Blocks.Keys) {
        $Definition = @($Metadata.SelectNodes("//*[local-name()='Action' and @Name='$Action' and *[local-name()='Parameter' and @Name='bindingParameter' and @Type='graph.cloudPcReports']]"))
        Assert-Report ($Definition.Count -eq 1 -and $Definition[0].ReturnType.Type -ceq 'Edm.Stream') 'Action stream contract drift'
        foreach ($Field in @('reportName','select','skip','top')) { Assert-Report ($Definition[0].SelectNodes("*[local-name()='Parameter' and @Name='$Field']").Count -eq 1) 'Action parameter contract drift' }
    }
}
if ($env:ASSAY_W365_REPORT_FIXTURE) {
    [IO.File]::WriteAllText($env:ASSAY_W365_REPORT_FIXTURE, (@{ CheckResults = $ExportedChecks } | ConvertTo-Json -Depth 10), [Text.UTF8Encoding]::new($false))
}
Write-Output 'PASS: exact read-permission report requests, retired action zero requests, bounded table decoding and metadata-only findings; all network calls mocked.'