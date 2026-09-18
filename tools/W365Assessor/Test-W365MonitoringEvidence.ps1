#Requires -Version 5.1
[CmdletBinding()]
param([string]$MetadataPath)
$ErrorActionPreference = 'Stop'
$Tokens = $null
$ParseErrors = $null
$Ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Invoke-W365Discovery.ps1'), [ref]$Tokens, [ref]$ParseErrors)
if ($ParseErrors.Count) { throw 'Collector parse failure' }
$CheckFunction = $Ast.Find({ param($Node) $Node -is [Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -ceq 'New-CheckResult' }, $true)
. ([scriptblock]::Create($CheckFunction.Extent.Text))
function Write-Status { param($Message, $Level) }
function Invoke-GraphPaged {
    param($Uri)
    if ($Uri -cne ($IntuneBase + '/userExperienceAnalyticsDeviceScores?$top=50')) { throw 'Unexpected analytics request' }
    if ($script:ReadFails) { throw 'Synthetic analytics failure' }
    $script:AnalyticsRows
}
function Invoke-MgGraphRequest {
    param($Method, $Uri)
    if ($Method -cne 'GET' -or $Uri -cne ($IntuneBase + '/softwareUpdateStatusSummary')) { throw 'Unexpected update-summary request' }
    if ($script:ReadFails) { throw 'Synthetic update-summary failure' }
    $script:UpdateSummary
}
$script:ReadFails = $false
$ScriptRoot = $PSScriptRoot
$IntuneBase = 'https://graph.microsoft.com/beta/deviceManagement'
$Discovery = @{ Inventory = @{ CloudPCs = @(@{ Id = 'cloudpc'; ManagedDeviceId = 'managed-cloudpc' }) } }
$Ids = @('W365-MON-001-EA','W365-MON-005-UPD')
$Blocks = @{}
foreach ($Id in $Ids) {
    $Needle = "-Id '$Id'"
    $Found = @($Ast.FindAll({ param($Node) $Node -is [Management.Automation.Language.TryStatementAst] -and $Node.Body.Extent.Text.Contains($Needle) }, $true))
    if ($Found.Count -ne 1) { throw ('Expected one monitoring block: ' + $Id) }
    $Blocks[$Id] = [scriptblock]::Create($Found[0].Extent.Text)
}
$script:AnalyticsRows = @(@{ id = 'unrelated-device'; endpointAnalyticsScore = 99 })
$script:UpdateSummary = @{}
$AllChecks = [Collections.Generic.List[object]]::new()
foreach ($Id in $Ids) {
    & $Blocks[$Id]
    if ($AllChecks[-1].Status -ne 'Error') { throw ('Unscoped or missing monitoring evidence produced a verdict: ' + $Id) }
}
if ($Discovery.Inventory.EndpointAnalyticsEvidence[0].Scores.endpointAnalyticsScore.Value -ne 99) { throw 'Analytics block failed instead of decoding evidence' }
if ($Discovery.Inventory.SoftwareUpdateSummaryEvidence.FieldState -ne 'Partial' -or $Discovery.Inventory.SoftwareUpdateSummaryEvidence.UnknownFields.Count -ne 7) { throw 'Empty update summary lost missing-field state' }
foreach ($Malformed in @('-1','3 devices','1.5','not reported')) {
    $script:UpdateSummary = @{ compliantDeviceCount = $Malformed; nonCompliantDeviceCount = 0 }
    & $Blocks['W365-MON-005-UPD']
    if ($AllChecks[-1].Status -ne 'Error' -or $null -ne $AllChecks[-1].Evidence.Counts.compliantDeviceCount) { throw ('Malformed count produced evidence: ' + $Malformed) }
}
. (Join-Path $PSScriptRoot 'W365Monitoring.ps1')
function Assert-Monitoring([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
$ScoreFields = @('endpointAnalyticsScore','startupPerformanceScore','appReliabilityScore','workFromAnywhereScore','meanResourceSpikeTimeScore','batteryHealthScore')
$CountFields = @('compliantDeviceCount','nonCompliantDeviceCount','remediatedDeviceCount','errorDeviceCount','unknownDeviceCount','conflictDeviceCount','notApplicableDeviceCount')
$ScoreRow = @{ id = 'score-entry'; healthStatus = 'meetingGoals'; deviceName = 'PRIVATE_NAME'; managedDeviceId = 'UNREVIEWED_DEVICE'; secret = 'SECRET' }
foreach ($Field in $ScoreFields) { $ScoreRow[$Field] = 72.5 }
foreach ($Value in @(0,100,72.5,-1)) {
    foreach ($Field in $ScoreFields) { $ScoreRow[$Field] = $Value }
    $Observed = ConvertTo-W365AnalyticsEvidence @($ScoreRow)
    Assert-Monitoring ($Observed.FieldState -eq 'Complete') 'Documented score rejected'
    foreach ($Field in $ScoreFields) {
        if ($Value -eq -1) { Assert-Monitoring ($Observed.Scores[$Field].State -eq 'Unavailable' -and $null -eq $Observed.Scores[$Field].Value) 'Unavailable score became poor performance' }
        else { Assert-Monitoring ($Observed.Scores[$Field].State -eq 'Observed' -and $Observed.Scores[$Field].Value -eq $Value) 'Valid score lost' }
    }
}
foreach ($Field in $ScoreFields) {
    foreach ($Invalid in @($null,$true,'90',-2,101,[double]::NaN,[double]::PositiveInfinity,[double]::NegativeInfinity)) {
        $ScoreRow[$Field] = $Invalid
        $Observed = ConvertTo-W365AnalyticsEvidence @($ScoreRow)
        Assert-Monitoring ($Observed.Scores[$Field].State -eq 'Unknown' -and $null -eq $Observed.Scores[$Field].Value -and $Observed.FieldState -eq 'Partial') 'Invalid score was retained'
    }
    $ScoreRow[$Field] = 90
}
foreach ($Health in @('unknown','insufficientData','needsAttention','meetingGoals')) {
    $ScoreRow.healthStatus = $Health
    Assert-Monitoring ((ConvertTo-W365AnalyticsEvidence @($ScoreRow)).HealthStatus -ceq $Health) 'Documented health state lost'
}
$ScoreRow.healthStatus = 'unknownFutureValue'
Assert-Monitoring ($null -eq (ConvertTo-W365AnalyticsEvidence @($ScoreRow)).HealthStatus) 'Unknown health enum guessed'
$ScoreRow.healthStatus = 'insufficientData'
$Observed = ConvertTo-W365AnalyticsEvidence @($ScoreRow)
Assert-Monitoring (($Observed | ConvertTo-Json -Depth 10) -notmatch 'PRIVATE_NAME|UNREVIEWED_DEVICE|SECRET') 'Unexpected analytics metadata exported'
Assert-Monitoring ($Observed.IdentityMatch -eq 'NotEvaluated' -and $Observed.DataAge -eq 'NotAvailableInContract') 'Score entry misrepresented as fresh device evidence'

$Summary = @{ compliantUserCount = 999; displayName = 'PRIVATE_POLICY'; secret = 'SECRET' }
foreach ($Field in $CountFields) { $Summary[$Field] = 0 }
foreach ($Shape in @($Summary, @{ value = $Summary }, ($Summary | ConvertTo-Json | ConvertFrom-Json), (@{ value = $Summary } | ConvertTo-Json | ConvertFrom-Json))) {
    $Observed = ConvertTo-W365UpdateSummaryEvidence $Shape
    Assert-Monitoring ($Observed.FieldState -eq 'Complete' -and $Observed.UnknownFields.Count -eq 0) 'Valid summary shape not decoded'
    Assert-Monitoring (($Observed | ConvertTo-Json -Depth 8) -notmatch 'PRIVATE_POLICY|SECRET|compliantUserCount') 'Unsupported fields retained'
    foreach ($Field in $CountFields) { Assert-Monitoring ($Observed.Counts[$Field] -eq 0) 'Zero counter lost' }
}
foreach ($Field in $CountFields) {
    foreach ($Invalid in @($null,$true,-1,1.5,'12','-1','1,234',[long]2147483648)) {
        $Summary[$Field] = $Invalid
        $Observed = ConvertTo-W365UpdateSummaryEvidence $Summary
        Assert-Monitoring ($null -eq $Observed.Counts[$Field] -and $Observed.UnknownFields -contains $Field -and $Observed.FieldState -eq 'Partial') 'Malformed summary counter coerced'
    }
    $Summary[$Field] = [int]::MaxValue
    Assert-Monitoring ((ConvertTo-W365UpdateSummaryEvidence $Summary).Counts[$Field] -eq [int]::MaxValue) 'Int32 maximum rejected'
    $Summary.Remove($Field)
    Assert-Monitoring ($null -eq (ConvertTo-W365UpdateSummaryEvidence $Summary).Counts[$Field]) 'Absent summary counter defaulted'
    $Summary[$Field] = 0
}
foreach ($Invalid in @('invalid',@(),@{ value = @() },@{ value = $null },@{ value = @{}; compliantDeviceCount = 0 })) {
    $Rejected = $false
    try { ConvertTo-W365UpdateSummaryEvidence $Invalid | Out-Null } catch { $Rejected = $true }
    Assert-Monitoring $Rejected 'Invalid or ambiguous summary shape accepted'
}
$Summary.compliantDeviceCount = 7
$Summary.errorDeviceCount = 2
$Summary.unknownDeviceCount = 3
$Summary.conflictDeviceCount = 1
$script:AnalyticsRows = @(($ScoreRow | ConvertTo-Json | ConvertFrom-Json))
$script:UpdateSummary = @{ value = $Summary } | ConvertTo-Json | ConvertFrom-Json
$AllChecks.Clear()
foreach ($Id in $Ids) { & $Blocks[$Id] }
Assert-Monitoring ($AllChecks.Count -eq 2 -and @($AllChecks | Where-Object Status -ne 'Error').Count -eq 0) 'Production path scored tenant summary'
Assert-Monitoring ($Discovery.Inventory.SoftwareUpdateSummaryEvidence.Counts.errorDeviceCount -eq 2) 'Production path lost error counters'
$ExportedChecks = @($AllChecks.ToArray())
$script:AnalyticsRows = @()
& $Blocks['W365-MON-001-EA']
Assert-Monitoring ($AllChecks[-1].Status -eq 'Error' -and $AllChecks[-1].Evidence.ScoreEntries -eq 0) 'Empty analytics response became a verdict or failure'
$script:ReadFails = $true
foreach ($Id in $Ids) { & $Blocks[$Id]; Assert-Monitoring ($AllChecks[-1].Status -eq 'Error' -and $AllChecks[-1].Evidence.Error) 'Read failure hidden' }
Assert-Monitoring (($AllChecks | Where-Object Id -eq 'W365-MON-001-EA' | Select-Object -Last 1).Evidence.RequiredScope -eq 'DeviceManagementManagedDevices.Read.All') 'Wrong analytics permission guidance'

if ($MetadataPath) {
    [xml]$Metadata = Get-Content -LiteralPath $MetadataPath -Raw
    foreach ($Contract in @(@{ Name = 'softwareUpdateStatusSummary'; Fields = $CountFields; Type = 'Edm.Int32' }, @{ Name = 'userExperienceAnalyticsDeviceScores'; Fields = $ScoreFields; Type = 'Edm.Double' })) {
        $Types = @($Metadata.SelectNodes("//*[local-name()='EntityType' and @Name='$($Contract.Name)']"))
        Assert-Monitoring ($Types.Count -eq 1) 'Ambiguous or missing schema type'
        foreach ($Field in $Contract.Fields) {
            $Property = @($Types[0].SelectNodes("*[local-name()='Property' and @Name='$Field']"))
            Assert-Monitoring ($Property.Count -eq 1 -and $Property[0].Type -ceq $Contract.Type) 'Field/type contract drift'
        }
    }
}
if ($env:ASSAY_W365_MONITORING_FIXTURE) {
    [IO.File]::WriteAllText($env:ASSAY_W365_MONITORING_FIXTURE, (@{ CheckResults = $ExportedChecks } | ConvertTo-Json -Depth 15), [Text.UTF8Encoding]::new($false))
}
Write-Output 'PASS: typed analytics scores/sentinels and seven update counters, unknown/zero distinction, singleton shapes, production scope guards and permission guidance; no Graph calls.'