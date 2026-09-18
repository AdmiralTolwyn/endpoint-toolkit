#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
$Tokens = $null
$ParseErrors = $null
$Ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Invoke-W365Discovery.ps1'), [ref]$Tokens, [ref]$ParseErrors)
if ($ParseErrors.Count) { throw 'Collector parse failure' }
$CheckFunction = $Ast.Find({ param($Node) $Node -is [Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -ceq 'New-CheckResult' }, $true)
. ([scriptblock]::Create($CheckFunction.Extent.Text))
function Write-Status { param($Message, $Level) }
$script:SecurityRequests = [Collections.Generic.List[string]]::new()
function Invoke-GraphPaged {
    param($Uri)
    $script:SecurityRequests.Add($Uri)
    if ($script:SecurityReadFails) { throw 'Synthetic transport failure' }
    if ($Uri -match '/managedDevices\?') { return $script:SecurityDeviceRows }
    if ($Uri -match '/deviceCompliancePolicies\?') { return $script:SecurityComplianceRows }
    if ($Uri -match '/configurationPolicies\?') { return $script:SecurityConfigurationRows }
    throw 'Unexpected security request'
}
$IntuneBase = 'https://graph.microsoft.com/beta/deviceManagement'
$Ids = @('W365-SEC-002-MDE','W365-SEC-003-BASE','W365-SEC-004-COMP')
$Blocks = @{}
foreach ($Id in $Ids) {
    $Needle = "-Id '$Id'"
    $Found = @($Ast.FindAll({ param($Node) $Node -is [Management.Automation.Language.TryStatementAst] -and $Node.Body.Extent.Text.Contains($Needle) }, $true))
    if ($Found.Count -ne 1) { throw ('Expected one security production block: ' + $Id) }
    $Blocks[$Id] = [scriptblock]::Create($Found[0].Extent.Text)
}
$script:SecurityReadFails = $false
$script:SecurityDeviceRows = @(@{ id = 'unverified-device'; complianceState = 'compliant' })
$script:SecurityComplianceRows = @(@{ id = 'unrelated-policy'; '@odata.type' = '#microsoft.graph.iosCompliancePolicy'; assignments = @(@{ id = 'unrelated-assignment' }) })
$script:SecurityConfigurationRows = @(@{ id = 'name-only'; name = 'Windows 365 security baseline'; description = 'Cloud PC'; templateReference = @{ templateFamily = 'none' } })
$AllChecks = [Collections.Generic.List[object]]::new()
foreach ($Id in $Ids) {
    & $Blocks[$Id]
    if ($AllChecks[-1].Status -ne 'Error') { throw ('Unverified security evidence produced a verdict: ' + $Id + ' = ' + $AllChecks[-1].Status) }
    if ($AllChecks[-1].Evidence.AssessmentState -ne 'InsufficientEvidence') { throw ('Context read failed instead of reaching the evidence gate: ' + $Id) }
}
$ExportedChecks = @($AllChecks.ToArray())
foreach ($State in @($null,'','unknown','noncompliant','compliant')) {
    $script:SecurityDeviceRows = @(@{ id = 'unverified-device'; complianceState = $State })
    & $Blocks['W365-SEC-002-MDE']
    if ($AllChecks[-1].Status -ne 'Error') { throw 'Compliance state substituted for Defender evidence' }
    $ExpectedUnknown = if ($State -cin @('compliant','noncompliant')) { 0 } else { 1 }
    if ($AllChecks[-1].Evidence.OtherOrMissing -ne $ExpectedUnknown) { throw 'Unknown compliance state was lost' }
}
$script:SecurityDeviceRows = @()
$script:SecurityComplianceRows = @()
$script:SecurityConfigurationRows = @()
foreach ($ReadFails in @($false,$true)) {
    $script:SecurityReadFails = $ReadFails
    foreach ($Id in $Ids) {
        & $Blocks[$Id]
        if ($AllChecks[-1].Status -ne 'Error') { throw ('Empty or failed collection produced a security verdict: ' + $Id) }
    }
}
if (@($script:SecurityRequests | Where-Object { $_ -match '\$expand' }).Count) { throw 'Unused assignment expansion still requested' }
foreach ($Path in @('deviceCompliancePolicies','configurationPolicies')) {
    if ($script:SecurityRequests -cnotcontains ($IntuneBase + '/' + $Path + '?$select=id')) { throw 'Expected metadata-only policy query' }
}
if ($env:ASSAY_W365_SECURITY_FIXTURE) {
    [IO.File]::WriteAllText($env:ASSAY_W365_SECURITY_FIXTURE, (@{ CheckResults = $ExportedChecks } | ConvertTo-Json -Depth 10), [Text.UTF8Encoding]::new($false))
}
Write-Output 'PASS: production security checks do not score names, unrelated policies or Intune compliance as Cloud PC protection; no live Graph calls.'