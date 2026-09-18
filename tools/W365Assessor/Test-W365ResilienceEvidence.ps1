#Requires -Version 5.1
[CmdletBinding()]
param([string]$V1MetadataPath)
$ErrorActionPreference = 'Stop'
$Tokens = $null
$ParseErrors = $null
$Ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Invoke-W365Discovery.ps1'), [ref]$Tokens, [ref]$ParseErrors)
if ($ParseErrors.Count) { throw 'Collector parse failure' }
$CheckFunction = $Ast.Find({ param($Node) $Node -is [Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -ceq 'New-CheckResult' }, $true)
. ([scriptblock]::Create($CheckFunction.Extent.Text))
$Blocks = @{}
foreach ($Id in @('W365-USER-002-','W365-PROV-005-')) {
    $Needle = '-Id "' + $Id
    $Found = @($Ast.FindAll({ param($Node) $Node -is [Management.Automation.Language.ForEachStatementAst] -and $Node.Body.Extent.Text.Contains($Needle) }, $true))
    if ($Found.Count -ne 1) { throw ('Expected one production loop: ' + $Id) }
    $Blocks[$Id] = [scriptblock]::Create($Found[0].Extent.Text)
}
$Discovery = @{ Inventory = @{
    UserSettings = @(@{ Id = 'user-setting'; DisplayName = 'Synthetic'; CrossRegionDisasterRecoverySetting = @{ disasterRecoveryType = 'unknownFutureValue' } })
    ProvisioningPolicies = @(@{ Id = 'policy'; DisplayName = 'Synthetic'; GracePeriodInHours = 24 })
} }
$AllChecks = [Collections.Generic.List[object]]::new()
foreach ($Id in $Blocks.Keys) {
    & $Blocks[$Id]
    if ($AllChecks[-1].Status -ne 'Error') { throw ('Unsupported resilience evidence produced a verdict: ' + $AllChecks[-1].Id + ' = ' + $AllChecks[-1].Status) }
}
function Assert-Resilience([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
foreach ($Payload in @($null,@{},@{ disasterRecoveryType = 'notConfigured' },@{ disasterRecoveryType = 'crossRegion'; secret = 'SECRET_DR' },@{ disasterRecoveryType = $true },@{ disasterRecoveryType = 1 })) {
    $Discovery.Inventory.UserSettings[0].CrossRegionDisasterRecoverySetting = $Payload
    & $Blocks['W365-USER-002-']
    $Result = $AllChecks[-1]
    Assert-Resilience ($Result.Status -eq 'Error' -and $Result.Evidence.AssessmentState -eq 'NotCollectedByApiContract' -and $Result.Evidence.ApiVersion -eq 'v1.0') 'DR collection gap was inferred as enablement/disablement'
    Assert-Resilience (($Result | ConvertTo-Json -Depth 10) -notmatch 'SECRET_DR') 'Unsupported DR payload copied to finding'
}
foreach ($Hours in @(0,1,24,168,169,[int]::MaxValue)) {
    $Discovery.Inventory.ProvisioningPolicies[0].GracePeriodInHours = $Hours
    & $Blocks['W365-PROV-005-']
    $Result = $AllChecks[-1]
    Assert-Resilience ($Result.Status -eq 'Error' -and $Result.Evidence.GracePeriodInHours -eq $Hours -and $Result.Evidence.FieldState -eq 'Observed' -and $Result.Evidence.ReadOnly -eq $true) 'Observed grace value was scored, changed or lost'
}
foreach ($Hours in @($null,'168','168 hours',$true,-1,1.5,[long]2147483648,[double]::NaN)) {
    $Discovery.Inventory.ProvisioningPolicies[0].GracePeriodInHours = $Hours
    & $Blocks['W365-PROV-005-']
    $Result = $AllChecks[-1]
    Assert-Resilience ($Result.Status -eq 'Error' -and $null -eq $Result.Evidence.GracePeriodInHours -and $Result.Evidence.FieldState -eq 'Unknown') 'Invalid grace value was coerced'
    Assert-Resilience ($Result.Details.Contains('gracePeriodInHours=Unknown')) 'Unknown grace observation not visible'
}

$UserLoops = @($Ast.FindAll({ param($Node) $Node -is [Management.Automation.Language.ForEachStatementAst] -and $Node.Body.Extent.Text.Contains('$Discovery.Inventory.UserSettings +=') }, $true))
$PolicyLoops = @($Ast.FindAll({ param($Node) $Node -is [Management.Automation.Language.ForEachStatementAst] -and $Node.Body.Extent.Text.Contains('$Discovery.Inventory.ProvisioningPolicies +=') }, $true))
Assert-Resilience ($UserLoops.Count -eq 1 -and $PolicyLoops.Count -eq 1) 'Expected exact production projections'
$UserProjection = [scriptblock]::Create($UserLoops[0].Extent.Text)
$PolicyProjection = [scriptblock]::Create($PolicyLoops[0].Extent.Text)
$UserSet = @(@{ id = 'user-setting'; displayName = 'Synthetic'; localAdminEnabled = $false; resetEnabled = $false; restorePointSetting = @{ frequencyInHours = 12; userRestoreEnabled = $true }; crossRegionDisasterRecoverySetting = @{ secret = 'SECRET_DR' }; notificationSetting = @{ secret = 'SECRET_NOTIFICATION' }; assignments = @() } | ConvertTo-Json -Depth 10 | ConvertFrom-Json)
$Discovery.Inventory.UserSettings = @()
& $UserProjection
$Projected = $Discovery.Inventory.UserSettings[0]
Assert-Resilience ($null -eq $Projected.PSObject.Properties['CrossRegionDisasterRecoverySetting'] -and $null -eq $Projected.PSObject.Properties['NotificationSetting']) 'Beta payload exported by v1 adapter'
Assert-Resilience ($Projected.ApiVersion -eq 'v1.0' -and $Projected.CrossRegionDisasterRecoveryState -eq 'NotCollectedByV1Contract' -and $Projected.NotificationSettingState -eq 'NotCollectedByV1Contract') 'Explicit v1 collection gaps lost'
Assert-Resilience ($Projected.LocalAdminEnabled -eq $false -and $Projected.ResetEnabled -eq $false -and $Projected.RestorePointFrequencyInHours -eq 12 -and $Projected.RestorePointUserRestoreEnabled -eq $true) 'Supported v1 properties changed'
Assert-Resilience (($Projected | ConvertTo-Json -Depth 10) -notmatch 'SECRET_DR|SECRET_NOTIFICATION') 'Unsupported nested data retained'
foreach ($Hours in @(0,24,168,$null,'168',-1,$true)) {
    $ProvPols = @(@{ id = 'policy'; displayName = 'Synthetic'; gracePeriodInHours = $Hours; assignments = @() } | ConvertTo-Json -Depth 10 | ConvertFrom-Json)
    $Discovery.Inventory.ProvisioningPolicies = @()
    & $PolicyProjection
    $Projected = $Discovery.Inventory.ProvisioningPolicies[0]
    $Valid = $Hours -is [int] -and $Hours -ge 0
    if ($Valid) { Assert-Resilience ($Projected.GracePeriodInHours -eq $Hours -and $Projected.GracePeriodFieldState -eq 'Observed') 'Valid projected grace value lost' }
    else { Assert-Resilience ($null -eq $Projected.GracePeriodInHours -and $Projected.GracePeriodFieldState -eq 'Unknown') 'Invalid raw grace value exported' }
}
$Discovery.Inventory.ProvisioningPolicies[0].GracePeriodInHours = 168
$AllChecks.Clear()
foreach ($Id in @('W365-USER-002-','W365-PROV-005-')) { & $Blocks[$Id] }
Assert-Resilience ($AllChecks.Count -eq 2 -and @($AllChecks | Where-Object Status -ne 'Error').Count -eq 0) 'Production projection/evaluation integration failed'

if ($V1MetadataPath) {
    [xml]$Metadata = Get-Content -LiteralPath $V1MetadataPath -Raw
    $UserType = @($Metadata.SelectNodes("//*[local-name()='EntityType' and @Name='cloudPcUserSetting']"))
    $PolicyType = @($Metadata.SelectNodes("//*[local-name()='EntityType' and @Name='cloudPcProvisioningPolicy']"))
    Assert-Resilience ($UserType.Count -eq 1 -and $PolicyType.Count -eq 1) 'Required v1 schema types missing or ambiguous'
    foreach ($Field in @('crossRegionDisasterRecoverySetting','notificationSetting')) {
        Assert-Resilience ($UserType[0].SelectNodes("*[local-name()='Property' and @Name='$Field']").Count -eq 0) 'v1 field contract changed; review required'
    }
    $Property = @($PolicyType[0].SelectNodes("*[local-name()='Property' and @Name='gracePeriodInHours']"))
    Assert-Resilience ($Property.Count -eq 1 -and $Property[0].Type -ceq 'Edm.Int32') 'Grace field type drift'
}
if ($env:ASSAY_W365_RESILIENCE_FIXTURE) {
    [IO.File]::WriteAllText($env:ASSAY_W365_RESILIENCE_FIXTURE, (@{ CheckResults = @($AllChecks.ToArray()) } | ConvertTo-Json -Depth 10), [Text.UTF8Encoding]::new($false))
}
Write-Output 'PASS: exact v1 projections, unsupported DR/notification exclusion, typed read-only grace observations and no unsupported resilience verdicts; no Graph calls.'