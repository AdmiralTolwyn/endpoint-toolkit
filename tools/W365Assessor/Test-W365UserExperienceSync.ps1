#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'W365UserExperienceSync.ps1')

function Assert-Sync([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
$Policy = @{ id = 'policy'; managedBy = 'windows365'; provisioningType = 'sharedByEntraGroup'; userSettingsPersistenceConfiguration = @{ userSettingsPersistenceEnabled = $true; userSettingsPersistenceStorageSizeCategory = 'fourGB' } }
Assert-Sync ((Get-W365UserExperienceSyncAssessment $Policy Enabled).Status -eq 'Pass') 'Enabled target not met'
Assert-Sync ((Get-W365UserExperienceSyncAssessment $Policy Disabled).Status -eq 'Warning') 'Disabled target mismatch lost'
Assert-Sync ((Get-W365UserExperienceSyncAssessment $Policy).Status -eq 'Error') 'Missing customer target scored'
foreach ($Storage in @('fourGB','eightGB','sixteenGB','thirtyTwoGB','sixtyFourGB')) {
    $Policy.userSettingsPersistenceConfiguration.userSettingsPersistenceStorageSizeCategory = $Storage
    Assert-Sync ((Get-W365UserExperienceSyncAssessment $Policy Enabled).Status -eq 'Pass') 'Documented storage size rejected'
}
$Policy.userSettingsPersistenceConfiguration.userSettingsPersistenceStorageSizeCategory = 'unknownFutureValue'
Assert-Sync ((Get-W365UserExperienceSyncAssessment $Policy Enabled).Status -eq 'Error') 'Unknown storage enum passed'
$Policy.userSettingsPersistenceConfiguration.userSettingsPersistenceEnabled = $false
Assert-Sync ((Get-W365UserExperienceSyncAssessment $Policy Disabled).Status -eq 'Pass') 'Explicit disabled target rejected'
foreach ($Invalid in @($null,1,'true')) {
    $Policy.userSettingsPersistenceConfiguration.userSettingsPersistenceEnabled = $Invalid
    Assert-Sync ((Get-W365UserExperienceSyncAssessment $Policy Enabled).Status -eq 'Error') 'Invalid Boolean passed'
}
foreach ($Type in @('dedicated','shared','sharedByUser','reserve')) {
    $Policy.provisioningType = $Type
    Assert-Sync ((Get-W365UserExperienceSyncAssessment $Policy Enabled).Status -eq 'N/A') 'Non-shared policy incorrectly applicable'
}
$Policy.provisioningType = 'unknownFutureValue'
Assert-Sync ((Get-W365UserExperienceSyncAssessment $Policy Enabled).Status -eq 'Error') 'Unknown policy type passed'
$Policy.managedBy = 'devBox'
Assert-Sync ((Get-W365UserExperienceSyncAssessment $Policy Enabled).Status -eq 'Error') 'Non-Windows365 owner passed'

$Results = @(Invoke-W365UserExperienceSyncRead -ExpectedState Enabled -Request {
    param($RequestUri)
    Assert-Sync ($RequestUri -like 'https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies?*') 'Wrong API path'
    @{ value = @(@{ id = 'policy'; managedBy = 'windows365'; provisioningType = 'sharedByEntraGroup'; userSettingsPersistenceConfiguration = @{ userSettingsPersistenceEnabled = $true; userSettingsPersistenceStorageSizeCategory = 'fourGB'; Unreviewed = 'SECRET' } }) }
})
Assert-Sync ($Results.Count -eq 1 -and $Results[0].Status -eq 'Pass') 'Production read/decode failed'
Assert-Sync (($Results | ConvertTo-Json -Depth 8) -notmatch 'SECRET|Unreviewed') 'Unexpected fields exported'
foreach ($NextLink in @('https://example.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies','https://graph.microsoft.com/v1.0/users','http://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies')) {
    $script:TestNextLink = $NextLink
    $Rejected = $false
    try {
        Invoke-W365UserExperienceSyncRead -Request { @{ value = @(@{ id = 'policy' }); '@odata.nextLink' = $script:TestNextLink } } | Out-Null
    } catch { $Rejected = $true }
    Assert-Sync $Rejected 'Untrusted continuation followed'
}
$Rejected = $false
try { Invoke-W365UserExperienceSyncRead -Request { throw 'Synthetic query failure' } | Out-Null } catch { $Rejected = $true }
Assert-Sync $Rejected 'Query failure became complete evidence'

$script:SyncPage = 0
$Paged = @(Invoke-W365UserExperienceSyncRead -Request {
    param($RequestUri)
    $script:SyncPage++
    if ($script:SyncPage -eq 1) {
        @{ value = @(@{ id = 'first' }); '@odata.nextLink' = $RequestUri + '&$skiptoken=second' }
    } else { @{ value = @(@{ id = 'second' }) } }
})
Assert-Sync ($Paged.Count -eq 2 -and $script:SyncPage -eq 2) 'Valid pagination failed'
foreach ($Suffix in @('&$expand=assignments','&$filter=id%20eq%20%27other%27','&$select=id')) {
    $script:SyncSuffix = $Suffix
    $Rejected = $false
    try { Invoke-W365UserExperienceSyncRead -Request { param($RequestUri) @{ value = @(@{ id = 'first' }); '@odata.nextLink' = $RequestUri + $script:SyncSuffix } } | Out-Null } catch { $Rejected = $true }
    Assert-Sync $Rejected 'Changed query accepted'
}
foreach ($Fixture in @(
    @{ value = @(@{ id = 'duplicate' },@{ id = 'duplicate' }) },
    @{ value = @(@{ id = $null }) },
    @{ value = 'invalid' },
    @{}
)) {
    $script:InvalidSyncResponse = $Fixture
    $Rejected = $false
    try { Invoke-W365UserExperienceSyncRead -Request { $script:InvalidSyncResponse } | Out-Null } catch { $Rejected = $true }
    Assert-Sync $Rejected 'Invalid response accepted'
}

$Tokens = $null
$ParseErrors = $null
$Ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Invoke-W365Discovery.ps1'), [ref]$Tokens, [ref]$ParseErrors)
Assert-Sync ($ParseErrors.Count -eq 0) 'Main collector parse failure'
$CheckFunction = $Ast.Find({ param($Node) $Node -is [Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -ceq 'New-CheckResult' }, $true)
. ([scriptblock]::Create($CheckFunction.Extent.Text))
$Branches = @($Ast.FindAll({ param($Node) $Node -is [Management.Automation.Language.IfStatementAst] -and $Node.Clauses[0].Item1.Extent.Text -ceq '$IncludeUserExperienceSync' }, $true))
Assert-Sync ($Branches.Count -eq 1) 'Expected one opt-in production branch'
$Production = [scriptblock]::Create($Branches[0].Extent.Text)
$ScriptRoot = $PSScriptRoot
$AllChecks = [Collections.Generic.List[object]]::new()
$Discovery = @{ Inventory = @{} }
$UserExperienceSyncTarget = 'Enabled'
$IncludeUserExperienceSync = $false
$script:SyncRequests = 0
function Invoke-MgGraphRequest {
    param($Method, $Uri, $Headers)
    $script:SyncRequests++
    Assert-Sync ($Method -ceq 'GET' -and $Headers.Prefer -ceq 'include-unknown-enum-members') 'Incorrect method or enum header'
    if ($script:SyncFail) { throw 'Synthetic transport failure' }
    @{ value = @(@{ id = 'policy'; managedBy = 'windows365'; provisioningType = 'sharedByEntraGroup'; userSettingsPersistenceConfiguration = @{ userSettingsPersistenceEnabled = $true; userSettingsPersistenceStorageSizeCategory = 'fourGB' } }) }
}
$script:SyncFail = $false
try {
    & $Production
    Assert-Sync ($script:SyncRequests -eq 0 -and $AllChecks.Count -eq 0) 'Disabled opt-in made a request'
    $IncludeUserExperienceSync = $true
    & $Production
    Assert-Sync ($script:SyncRequests -eq 1 -and $AllChecks[0].Id -ceq 'W365-PROV-011-policy' -and $AllChecks[0].Status -eq 'Pass') 'Main collector did not emit the comparison'
    $ExportedChecks = @($AllChecks.ToArray())
    $AllChecks.Clear()
    $script:SyncFail = $true
    & $Production
    Assert-Sync ($AllChecks.Count -eq 1 -and $AllChecks[0].Status -eq 'Error' -and $AllChecks[0].Id -ceq 'W365-PROV-011-collection') 'Main collector hid failed collection'
} finally { Remove-Item Function:\Invoke-MgGraphRequest }
if ($env:ASSAY_W365_UXSYNC_FIXTURE) {
    [IO.File]::WriteAllText($env:ASSAY_W365_UXSYNC_FIXTURE, (@{ CheckResults = $ExportedChecks } | ConvertTo-Json -Depth 10), [Text.UTF8Encoding]::new($false))
}
Write-Output 'PASS: documented UX Sync types, explicit targets, applicability, metadata projection and read failure/continuation guards; no Graph calls.'