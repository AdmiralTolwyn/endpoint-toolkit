#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
$Tokens = $null
$ParseErrors = $null
$Ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Invoke-W365Discovery.ps1'), [ref]$Tokens, [ref]$ParseErrors)
if ($ParseErrors.Count) { throw 'Collector parse failure' }
$CheckFunction = $Ast.Find({ param($Node) $Node -is [Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -ceq 'New-CheckResult' }, $true)
. ([scriptblock]::Create($CheckFunction.Extent.Text))
$Branches = @($Ast.FindAll({ param($Node) $Node -is [Management.Automation.Language.TryStatementAst] -and $Node.Body.Extent.Text.Contains('$caPolicies =') }, $true))
if ($Branches.Count -ne 1) { throw 'Expected one Conditional Access production block' }
$Production = [scriptblock]::Create($Branches[0].Extent.Text)
$ScriptRoot = $PSScriptRoot
$AppIdWindowsCloudLogin = '270efc09-cd0d-444b-a71f-39af4910ec45'
$AppIdAzureVirtualDesktop = '9cdead84-a844-4324-93f2-b2e6bb768d07'
$AppIdWindows365Portal = '0af06dc6-e4b5-4f28-818e-e78e62d137a5'
$CloudPcSignInAppIds = @($AppIdWindowsCloudLogin, $AppIdAzureVirtualDesktop, $AppIdWindows365Portal)
$CaPolicyUri = 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies'
function Write-Status { param($Message, $Level) }
function Invoke-GraphPaged { param($Uri) if ($Uri -cne $CaPolicyUri) { throw 'Unexpected CA request' }; if ($script:CaReadFails) { throw 'Synthetic read failure' }; $script:TestCaPolicies }
$script:CaReadFails = $false
$script:TestCaPolicies = @(@{
    id = 'excluded'; state = 'enabled'
    conditions = @{ users = @{ includeUsers = @('All') }; applications = @{ includeApplications = @('All'); excludeApplications = $CloudPcSignInAppIds } }
    grantControls = @{ operator = 'OR'; builtInControls = @('mfa','compliantDevice') }
    sessionControls = @{ signInFrequency = @{ isEnabled = $false; frequencyInterval = 'everyTime' } }
})
$AllChecks = [Collections.Generic.List[object]]::new()
$Discovery = @{ Inventory = @{} }
& $Production
if ($AllChecks.Count -ne 4) { throw 'Expected four CA results' }
foreach ($Check in $AllChecks) {
    if ($Check.Status -ne 'Error' -or $Check.Evidence.AssessmentState -ne 'EffectiveScopeNotEvaluated') { throw ('Unsupported enforcement verdict or failed observation path: ' + $Check.Id) }
}
$Observed = $Discovery.Inventory.ConditionalAccessPolicyEvidence[0]
if ($Observed.ApplicationTargets.WindowsCloudLogin -ne 'Excluded' -or $Observed.MfaRequirement -ne 'AlternativeOrUnresolved' -or $Observed.SignInFrequencyMode -ne 'Disabled') { throw 'Documented policy intent was lost' }
$ExportedChecks = @($AllChecks.ToArray())

. (Join-Path $PSScriptRoot 'W365ConditionalAccess.ps1')
function New-TestCaPolicy {
    '{"id":"policy","displayName":"PRIVATE_NAME","state":"enabled","conditions":{"users":{"includeUsers":["All"],"excludeUsers":["PRIVATE_USER"]},"applications":{"includeApplications":["All"],"excludeApplications":[],"applicationFilter":null,"includeUserActions":[]}},"grantControls":{"builtInControls":["mfa"],"operator":"AND","customAuthenticationFactors":[],"termsOfUse":[],"authenticationStrength":null},"sessionControls":{"signInFrequency":{"isEnabled":true,"frequencyInterval":"everyTime","authenticationType":"primaryAndSecondaryAuthentication","type":null,"value":null},"secureSignInSession":{"isEnabled":true}}}' | ConvertFrom-Json
}
function Assert-Ca([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
$Policy = New-TestCaPolicy
$Observed = ConvertTo-W365ConditionalAccessEvidence @($Policy)
Assert-Ca ($Observed.ApplicationTargets.Windows365 -eq 'Included' -and $Observed.MfaRequirement -eq 'RequiredInPolicy' -and $Observed.SignInFrequencyMode -eq 'EveryTime') 'Complete typed policy intent not decoded'
Assert-Ca ($Observed.TokenProtection -eq 'NotCollectedByV1Contract') 'Beta token control interpreted on v1'
Assert-Ca (($Observed | ConvertTo-Json -Depth 8) -notmatch 'PRIVATE_NAME|PRIVATE_USER') 'Unreviewed identity/name data exported'
$Policy.grantControls.operator = 'OR'
Assert-Ca ((ConvertTo-W365ConditionalAccessEvidence @($Policy)).MfaRequirement -eq 'RequiredInPolicy') 'Single MFA OR not decoded'
$Policy.grantControls.builtInControls = @('mfa','compliantDevice')
Assert-Ca ((ConvertTo-W365ConditionalAccessEvidence @($Policy)).MfaRequirement -eq 'AlternativeOrUnresolved') 'OR alternative treated as mandatory MFA'
$Policy.grantControls.operator = 'AND'
Assert-Ca ((ConvertTo-W365ConditionalAccessEvidence @($Policy)).MfaRequirement -eq 'RequiredInPolicy') 'AND MFA requirement lost'
foreach ($Controls in @(@('mfa','unknownFutureValue'), @('mfa','invented'))) {
    $Policy.grantControls.builtInControls = $Controls
    Assert-Ca ((ConvertTo-W365ConditionalAccessEvidence @($Policy)).MfaRequirement -eq 'Unknown') 'Unknown grant was interpreted'
}
$Policy.grantControls.builtInControls = @('block','mfa')
Assert-Ca ((ConvertTo-W365ConditionalAccessEvidence @($Policy)).MfaRequirement -eq 'BlockControlPresent') 'Block policy counted as MFA enforcement'
$Policy.grantControls.builtInControls = @()
$Policy.grantControls.authenticationStrength = @{ id = 'unreviewed-strength' }
Assert-Ca ((ConvertTo-W365ConditionalAccessEvidence @($Policy)).MfaRequirement -eq 'AuthenticationStrengthNotEvaluated') 'Authentication strength was guessed'
$Policy = New-TestCaPolicy
$Policy.grantControls.operator = 'OR'
foreach ($Field in @('customAuthenticationFactors','termsOfUse')) {
    $Policy.grantControls.$Field = @('alternative')
    Assert-Ca ((ConvertTo-W365ConditionalAccessEvidence @($Policy)).MfaRequirement -eq 'AlternativeOrUnresolved') 'Additional OR grant ignored'
    $Policy.grantControls.$Field = @()
}
$Policy.conditions.applications.excludeApplications = @($AppIdWindowsCloudLogin)
$Observed = ConvertTo-W365ConditionalAccessEvidence @($Policy)
Assert-Ca ($Observed.ApplicationTargets.WindowsCloudLogin -eq 'Excluded' -and $Observed.ApplicationTargets.Windows365 -eq 'Included') 'Per-app exclusions lost'
foreach ($Alias in @('Office365','MicrosoftAdminPortals','unknownCollection')) {
    $Policy.conditions.applications.excludeApplications = @($Alias)
    Assert-Ca ((ConvertTo-W365ConditionalAccessEvidence @($Policy)).ApplicationTargets.Windows365 -eq 'Unknown') 'App-collection membership guessed'
}
$Policy.conditions.applications.excludeApplications = @()
$Policy.conditions.applications.applicationFilter = @{ mode = 'exclude'; rule = 'PRIVATE_FILTER' }
Assert-Ca ((ConvertTo-W365ConditionalAccessEvidence @($Policy)).ApplicationTargets.Windows365 -eq 'Unknown') 'Dynamic app filter ignored'
$Policy.conditions.applications.applicationFilter = $null
$Policy.conditions.applications.includeApplications = @($AppIdWindows365Portal)
$Observed = ConvertTo-W365ConditionalAccessEvidence @($Policy)
Assert-Ca ($Observed.ApplicationTargets.Windows365 -eq 'Included' -and $Observed.ApplicationTargets.WindowsCloudLogin -eq 'NotDirectlyIncluded') 'Single-app scope widened'
foreach ($Invalid in @($null,'All',@($true))) {
    $Policy.conditions.applications.includeApplications = $Invalid
    Assert-Ca ((ConvertTo-W365ConditionalAccessEvidence @($Policy)).ApplicationTargets.Windows365 -eq 'Unknown') 'Malformed application list decoded'
}
$Policy = New-TestCaPolicy
foreach ($State in @('enabled','disabled','enabledForReportingButNotEnforced','unknownFutureValue')) {
    $Policy.state = $State
    $script:TestCaPolicies = @($Policy)
    $AllChecks.Clear()
    & $Production
    Assert-Ca ($AllChecks.Count -eq 4 -and @($AllChecks | Where-Object Status -ne 'Error').Count -eq 0) 'Policy state produced an enforcement verdict'
    $Expected = if ($State -eq 'unknownFutureValue') { 'Unknown' } else { $State }
    Assert-Ca ($Discovery.Inventory.ConditionalAccessPolicyEvidence[0].State -ceq $Expected) 'Report-only/disabled state lost'
}
$Policy.sessionControls.signInFrequency.isEnabled = $false
Assert-Ca ((ConvertTo-W365ConditionalAccessEvidence @($Policy)).SignInFrequencyMode -eq 'Disabled') 'Disabled frequency treated as active'
$Policy.sessionControls.signInFrequency.isEnabled = 'true'
Assert-Ca ((ConvertTo-W365ConditionalAccessEvidence @($Policy)).SignInFrequencyMode -eq 'Unknown') 'String Boolean accepted'
$Policy.sessionControls.signInFrequency.isEnabled = $true
$Policy.sessionControls.signInFrequency.frequencyInterval = 'timeBased'
$Policy.sessionControls.signInFrequency.type = 'hours'
$Policy.sessionControls.signInFrequency.value = 12
$Observed = ConvertTo-W365ConditionalAccessEvidence @($Policy)
Assert-Ca ($Observed.SignInFrequencyMode -eq 'TimeBased' -and $Observed.SignInFrequencyValue -eq 12) 'Periodic frequency lost'
foreach ($Invalid in @($null,0,-1,1.5,'12',$true)) {
    $Policy.sessionControls.signInFrequency.value = $Invalid
    Assert-Ca ((ConvertTo-W365ConditionalAccessEvidence @($Policy)).SignInFrequencyMode -eq 'Unknown') 'Invalid frequency value accepted'
}
foreach ($Fails in @($false,$true)) {
    $script:CaReadFails = $Fails
    $script:TestCaPolicies = @()
    $AllChecks.Clear()
    & $Production
    Assert-Ca ($AllChecks.Count -eq 4 -and @($AllChecks | Where-Object Status -ne 'Error').Count -eq 0) 'Empty/failed policy read produced a verdict'
}
if ($env:ASSAY_W365_CA_FIXTURE) {
    [IO.File]::WriteAllText($env:ASSAY_W365_CA_FIXTURE, (@{ CheckResults = $ExportedChecks } | ConvertTo-Json -Depth 10), [Text.UTF8Encoding]::new($false))
}
Write-Output 'PASS: source-bound per-app exclusions, MFA grant logic, policy states, typed frequency, unsupported token protection and evidence-only production export; no live Graph calls.'