#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
$Tokens = $null
$ParseErrors = $null
$CollectorPath = Join-Path $PSScriptRoot 'Invoke-W365Discovery.ps1'
$Ast = [Management.Automation.Language.Parser]::ParseFile($CollectorPath, [ref]$Tokens, [ref]$ParseErrors)
if ($ParseErrors.Count) { throw 'Collector parse failure' }
function Assert-Privacy([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }

foreach ($Name in @('New-CheckResult','Add-DiscoveryError','ConvertTo-W365AssignmentSummary','ConvertTo-W365UtcDate','Get-W365LastLoginDate','Get-W365UserLabel','ConvertTo-W365HealthChecks')) {
    $Definition = $Ast.Find({ param($Node) $Node -is [Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -ceq $Name }, $true)
    Assert-Privacy ($null -ne $Definition) ('Missing production function ' + $Name)
    . ([scriptblock]::Create($Definition.Extent.Text))
}
. (Join-Path $PSScriptRoot 'CollectorPrivacy.ps1')
. (Join-Path $PSScriptRoot 'W365ConditionalAccess.ps1')
function Write-Status { param($Message, $Level) }

function Get-ProductionBlock([string]$Needle, [type]$Type) {
    $Found = @($Ast.FindAll({ param($Node) $Node -is $Type -and $Node.Body.Extent.Text.Contains($Needle) }, $true))
    Assert-Privacy ($Found.Count -eq 1) ('Expected one production block: ' + $Needle)
    return [scriptblock]::Create($Found[0].Extent.Text)
}

$Initializer = @($Ast.FindAll({ param($Node) $Node -is [Management.Automation.Language.AssignmentStatementAst] -and $Node.Left.Extent.Text -ceq '$Discovery' }, $true))
Assert-Privacy ($Initializer.Count -eq 1) 'Expected one discovery initializer'
$Initializer = [scriptblock]::Create($Initializer[0].Extent.Text)
$Reads = @(
    '$CloudPCs = @(Invoke-GraphPaged', '$ProvPols = @(Invoke-GraphPaged', '$UserSet = @(Invoke-GraphPaged',
    '$ANCs = @(Invoke-GraphPaged', '$DevImgs = @(Invoke-GraphPaged', '$audits = @(Invoke-GraphPaged'
) | ForEach-Object { Get-ProductionBlock $_ ([Management.Automation.Language.TryStatementAst]) }
$Checks = @('-Id "W365-CPC-001-', '-Id "W365-COST-001-', '-Id "W365-CPC-004-', '-Id "W365-SEC-010-') |
    ForEach-Object { Get-ProductionBlock $_ ([Management.Automation.Language.ForEachStatementAst]) }
$StateLists = @($Ast.FindAll({ param($Node) $Node -is [Management.Automation.Language.AssignmentStatementAst] -and $Node.Left.Extent.Text -cin @('$failStates','$warnStates') }, $true))
Assert-Privacy ($StateLists.Count -eq 2) 'Expected production Cloud PC state lists'
$StateLists = [scriptblock]::Create((($StateLists | ForEach-Object { $_.Extent.Text }) -join "`n"))
$ExportStatements = @($Ast.EndBlock.Statements | Where-Object { $_.Extent.Text.Contains('New-CollectorPrivacyManifest') -or $_.Extent.Text.Contains('Write-CollectorExport') })
Assert-Privacy ($ExportStatements.Count -eq 2) 'Expected manifest and export statements'
$Export = [scriptblock]::Create((($ExportStatements | ForEach-Object { $_.Extent.Text }) -join "`n").Replace('exit 1', "throw 'Export failed'"))

$OldLogin = (Get-Date).ToUniversalTime().AddDays(-45).ToString('yyyy-MM-ddTHH:mm:ssZ')
function Invoke-GraphPaged {
    param($Uri)
    switch -Wildcard ($Uri) {
        '*/cloudPCs' {
            return @(
                @{ id = 'cpc-1'; displayName = 'CPC-PIISENTINEL'; status = 'failed'; userPrincipalName = 'pii.sentinel@example.test'; managedDeviceId = 'MANAGED_SENTINEL'; aadDeviceId = 'AAD_SENTINEL'; lastLoginResult = @{ time = $OldLogin }; diskEncryptionState = 'encrypting' },
                @{ id = 'cpc-2'; displayName = 'CPC-ORPHANSENTINEL'; status = 'provisioned'; userPrincipalName = $null; lastLoginResult = $null },
                @{ id = 'cpc-3'; displayName = 'CPC-IDLESENTINEL'; status = 'provisioned'; userPrincipalName = 'idle.sentinel@example.test'; lastLoginResult = @{ time = $OldLogin } },
                @{ id = 'cpc-4'; displayName = 'CPC-GRACESENTINEL'; status = 'inGracePeriod'; userPrincipalName = 'grace.sentinel@example.test'; gracePeriodEndDateTime = '2026-10-09T00:00:00Z' }
            )
        }
        '*/provisioningPolicies*' {
            return @(@{ id = 'policy-1'; displayName = 'Policy'; description = 'DESC_SENTINEL'; cloudPcGroupDisplayName = 'GROUP_SENTINEL'; assignments = @(
                @{ id = 'assignment-1'; target = @{ '@odata.type' = '#microsoft.graph.cloudPcManagementGroupAssignmentTarget'; groupId = 'ASSIGN_SENTINEL' } },
                @{ id = 'assignment-2'; target = @{ '@odata.type' = '#microsoft.graph.cloudPcManagementGroupAssignmentTarget'; groupId = 'ASSIGN_SENTINEL_2' } }
            ) })
        }
        '*/userSettings*' {
            return @(@{ id = 'settings-1'; displayName = 'Settings'; assignments = @(@{ id = 'assignment-3'; target = @{ '@odata.type' = '#microsoft.graph.cloudPcManagementGroupAssignmentTarget'; groupId = 'ASSIGN_SENTINEL_3' } }) })
        }
        '*/onPremisesConnections' {
            return @(@{ id = 'anc-1'; displayName = 'Network'; healthCheckStatus = 'failed'; adDomainUsername = 'DOMAIN_JOIN_SENTINEL'; organizationalUnit = 'OU=OU_SENTINEL,DC=example,DC=test'; healthCheckStatusDetail = @{ healthChecks = @(
                @{ displayName = 'DNS'; status = 'failed'; errorType = 'dnsCheckFqdnNotFound'; recommendedAction = 'Review DNS'; additionalDetails = 'HEALTH_SENTINEL'; correlationId = 'CORR_SENTINEL'; startDateTime = '2026-10-01T00:00:00Z'; endDateTime = '2026-10-01T00:05:00Z' }
            ) } })
        }
        '*/deviceImages' { throw 'Request for ERROR_SENTINEL@example.test failed' }
        '*/auditEvents*' {
            return @(@{ id = 'AUDIT_ID_SENTINEL'; displayName = 'Delete Cloud PC'; activityType = 'delete'; activityResult = 'success'; activityDateTime = '2026-10-01T13:45:12Z'; category = 'Cloud PC'; actor = @{ userPrincipalName = 'actor.sentinel@example.test'; applicationDisplayName = 'Portal' } })
        }
        default { throw ('Unexpected request ' + $Uri) }
    }
}

function New-TestContext([string]$Mode, [byte[]]$Key) {
    [pscustomobject]@{ Mode = $Mode; Key = $Key; KeyId = $(if ($Mode -eq 'Pseudonymous') { 'testkey' } else { $null }); KeyPath = $null; Identities = [Collections.Generic.Dictionary[string,string]]::new() }
}

function Invoke-W365PrivacyRun($PrivacyContext) {
    $script:PrivacyContext = $PrivacyContext
    $Context = @{ Account = 'operator.sentinel@example.test'; TenantId = '11111111-1111-4111-8111-111111111111' }
    $ScriptVersion = 'synthetic'
    $Assessor = 'Engagement label'
    $Script:CollectionId = 'collection-1'
    $GraphBase = 'https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint'
    $GraphBaseV1 = 'https://graph.microsoft.com/v1.0/deviceManagement/virtualEndpoint'
    . $Initializer
    $AllChecks = [Collections.ArrayList]::new()
    foreach ($Block in $Reads) { . $Block }
    $now = Get-Date
    $Today = $now.ToUniversalTime().Date
    $InactiveDays = 30
    $CpcNoLoginCount = 0
    . $StateLists
    foreach ($Block in $Checks) { . $Block }
    $Discovery.CheckResults = $AllChecks.ToArray()
    return $Discovery
}

$Sentinels = 'pii\.sentinel|idle\.sentinel|grace\.sentinel|PIISENTINEL|ORPHANSENTINEL|IDLESENTINEL|GRACESENTINEL|DESC_SENTINEL|GROUP_SENTINEL|ASSIGN_SENTINEL|DOMAIN_JOIN_SENTINEL|OU_SENTINEL|actor\.sentinel|AUDIT_ID_SENTINEL|HEALTH_SENTINEL|CORR_SENTINEL|MANAGED_SENTINEL|AAD_SENTINEL|ERROR_SENTINEL|operator\.sentinel|13:45:12'
$Key = [byte[]](1..32)
$Pseudo = Invoke-W365PrivacyRun (New-TestContext 'Pseudonymous' $Key)
$PseudoJson = $Pseudo | ConvertTo-Json -Depth 12
Assert-Privacy ($PseudoJson -notmatch $Sentinels) ('Pseudonymous export leaked: ' + [regex]::Match($PseudoJson, $Sentinels).Value)
Assert-Privacy ($null -eq $Pseudo.PSObject.Properties['AssessorId'] -and $Pseudo.Assessor -ceq 'Engagement label' -and $Pseudo.CollectionId -ceq 'collection-1') 'Assessor metadata incorrect'
$Cpcs = @{}
foreach ($Row in $Pseudo.Inventory.CloudPCs) { $Cpcs[$Row.Id] = $Row }
Assert-Privacy ($Cpcs['cpc-1'].UserKey -cmatch '^usr_[0-9a-f]{16}$' -and $Cpcs['cpc-1'].DisplayName -cmatch '^dev_[0-9a-f]{16}$' -and $Cpcs['cpc-1'].HasAssignedUser -eq $true) 'Cloud PC identity not pseudonymized'
Assert-Privacy ($Cpcs['cpc-2'].HasAssignedUser -eq $false -and $null -eq $Cpcs['cpc-2'].UserKey) 'Unassigned Cloud PC state lost'
Assert-Privacy ($null -eq $Cpcs['cpc-1'].PSObject.Properties['UserPrincipalName'] -and $null -eq $Cpcs['cpc-1'].PSObject.Properties['ManagedDeviceId'] -and $null -eq $Cpcs['cpc-1'].PSObject.Properties['LastLoginResult']) 'Identified-only Cloud PC fields exported'
Assert-Privacy ($Cpcs['cpc-3'].LastLoginDate -ceq $OldLogin.Substring(0, 10)) 'Last login date not reduced to UTC date'
$Policy = $Pseudo.Inventory.ProvisioningPolicies[0]
Assert-Privacy ($Policy.AssignmentCount -eq 2 -and $Policy.AssignmentTargetTypes.cloudPcManagementGroupAssignmentTarget -eq 2 -and $null -eq $Policy.PSObject.Properties['Assignments']) 'Provisioning assignment summary incorrect'
Assert-Privacy ($null -eq $Policy.PSObject.Properties['Description'] -and $null -eq $Policy.PSObject.Properties['CloudPcGroupDisplayName']) 'Provisioning free text exported'
Assert-Privacy ($Pseudo.Inventory.UserSettings[0].AssignmentCount -eq 1 -and $null -eq $Pseudo.Inventory.UserSettings[0].PSObject.Properties['Assignments']) 'User settings assignment summary incorrect'
$Network = $Pseudo.Inventory.AzureNetworkConnections[0]
Assert-Privacy ($Network.HasDomainJoinAccount -eq $true -and $Network.HasOrganizationalUnit -eq $true -and $null -eq $Network.PSObject.Properties['OrganizationalUnit'] -and $null -eq $Network.PSObject.Properties['HealthCheckStatusDetail']) 'Network identity fields not classified'
Assert-Privacy ((@($Network.HealthChecks[0].PSObject.Properties.Name) -join ',') -ceq 'displayName,status,errorType,recommendedAction,startDateTime,endDateTime') 'Health checks not reduced to the allowlist'
$Audit = $Pseudo.Inventory.AuditEvents[0]
Assert-Privacy ($Audit.ActivityDateTime -ceq '2026-10-01' -and $null -eq $Audit.PSObject.Properties['Id'] -and $null -eq $Audit.PSObject.Properties['ActorUpn'] -and $Audit.ActivityType -ceq 'delete') 'Audit event not minimized'
Assert-Privacy ($Pseudo.Errors.Count -eq 1 -and $Pseudo.Errors[0] -ceq '[DeviceImages] RuntimeException') 'Collection error text not sanitized'
$Results = @{}
foreach ($Result in $Pseudo.CheckResults) { $Results[$Result.Id] = $Result }
Assert-Privacy ($Results['W365-CPC-001-cpc-1'].Status -eq 'Fail' -and $Results['W365-CPC-001-cpc-1'].Evidence.User -ceq $Cpcs['cpc-1'].UserKey -and $Results['W365-CPC-001-cpc-1'].Details.Contains($Cpcs['cpc-1'].UserKey)) 'CPC-001 did not use the user pseudonym'
Assert-Privacy ($Results['W365-CPC-002-cpc-4'].Status -eq 'Warning') 'CPC-002 missing'
Assert-Privacy ($Results['W365-CPC-004-cpc-2'].Status -eq 'Fail' -and -not $Results.ContainsKey('W365-CPC-004-cpc-3')) 'CPC-004 did not use HasAssignedUser'
Assert-Privacy ($Results['W365-COST-001-cpc-3'].Evidence.IdleDays -ge 44 -and $Results['W365-COST-001-cpc-3'].Evidence.LastLogin -ceq $OldLogin.Substring(0, 10)) 'COST-001 lost the inactivity signal'
Assert-Privacy ($Results['W365-SEC-010-cpc-1'].Status -eq 'Warning') 'SEC-010 missing'

$Again = Invoke-W365PrivacyRun (New-TestContext 'Pseudonymous' $Key)
Assert-Privacy ($Again.Inventory.CloudPCs[0].UserKey -ceq $Cpcs['cpc-1'].UserKey) 'Same key changed the pseudonym'
$Other = Invoke-W365PrivacyRun (New-TestContext 'Pseudonymous' ([byte[]](2..33)))
Assert-Privacy ($Other.Inventory.CloudPCs[0].UserKey -cne $Cpcs['cpc-1'].UserKey) 'Different key produced the same pseudonym'

$Identified = Invoke-W365PrivacyRun (New-TestContext 'Identified' $null)
$IdentifiedJson = $Identified | ConvertTo-Json -Depth 12
foreach ($Expected in @('pii.sentinel@example.test','CPC-PIISENTINEL','DESC_SENTINEL','GROUP_SENTINEL','OU_SENTINEL','actor.sentinel@example.test','AUDIT_ID_SENTINEL','MANAGED_SENTINEL','ERROR_SENTINEL')) {
    Assert-Privacy ($IdentifiedJson.Contains($Expected)) ('Identified export dropped ' + $Expected)
}
Assert-Privacy ($IdentifiedJson -notmatch 'DOMAIN_JOIN_SENTINEL|ASSIGN_SENTINEL|HEALTH_SENTINEL|CORR_SENTINEL|operator\.sentinel') 'Identified export contains prohibited raw values'
Assert-Privacy ($Identified.Inventory.CloudPCs[1].HasAssignedUser -eq $false) 'Identified unassigned state lost'

$Policies = @(@{ id = 'policy-ca'; displayName = 'CA policy'; state = 'enabled'; conditions = @{
    users = @{ includeUsers = @('pii.sentinel@example.test'); excludeUsers = @('22222222-2222-4222-8222-222222222222'); includeGroups = @('GROUP_SENTINEL'); includeRoles = @('ROLE_SENTINEL') }
    applications = @{ includeApplications = @('All') }
}; grantControls = @{ operator = 'OR'; builtInControls = @('mfa') } })
$CaJson = @(ConvertTo-W365ConditionalAccessEvidence -Policies $Policies) | ConvertTo-Json -Depth 8
Assert-Privacy ($CaJson -notmatch 'pii\.sentinel|GROUP_SENTINEL|ROLE_SENTINEL|22222222-2222') 'Conditional Access evidence exported user or group targets'

$Temp = Join-Path ([IO.Path]::GetTempPath()) ('w365-privacy-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($Temp)
try {
    $script:PrivacyContext = New-CollectorPrivacyContext -Mode Pseudonymous -OutputPath (Join-Path $Temp 'w365.json')
    $Discovery = Invoke-W365PrivacyRun $script:PrivacyContext
    $OutputPath = Join-Path $Temp 'w365.json'
    $IdentityMapPath = Join-Path $Temp 'w365.identity-map.json'
    $IncludeConditionalAccess = [switch]$true
    $IncludeUserExperienceSync = [switch]$false
    . $Export
    $Written = Get-Content -LiteralPath $OutputPath -Raw | ConvertFrom-Json
    Assert-Privacy ($Written.Privacy.Mode -ceq 'Pseudonymous' -and $Written.Privacy.Classification -ceq 'Confidential' -and $Written.Privacy.PseudonymKeyId -ceq $script:PrivacyContext.KeyId -and @($Written.Privacy.OptIns) -join ',' -ceq 'IncludeConditionalAccess') 'Privacy manifest incorrect'
    Assert-Privacy ((Get-Content -LiteralPath $OutputPath -Raw) -notmatch $Sentinels) 'Written export leaked a sentinel'
    Assert-Privacy ((Get-Content -LiteralPath $IdentityMapPath -Raw).Contains('pii.sentinel@example.test')) 'Identity map missing original value'
    Assert-Privacy ((Get-Acl -LiteralPath $OutputPath).AreAccessRulesProtected) 'Export ACL inherits permissions'
    $Rejected = $false
    try { . $Export } catch { $Rejected = $true }
    Assert-Privacy $Rejected 'Existing export was overwritten'
    foreach ($Case in @(
        @{ Arguments = @{ PrivacyMode = 'Identified'; OutputPath = (Join-Path $Temp 'identified.json') }; Pattern = 'ConfirmIdentifiedExport' },
        @{ Arguments = @{ OutputPath = $OutputPath }; Pattern = 'already exists' }
    )) {
        $Arguments = $Case.Arguments
        $Message = $null
        try { & $CollectorPath @Arguments *> $null } catch { $Message = $_.Exception.Message }
        Assert-Privacy ($Message -match $Case.Pattern) ('Collector did not stop before collection: ' + $Case.Pattern)
    }
    Assert-Privacy (-not (Test-Path -LiteralPath (Join-Path $Temp 'identified.json')) -and -not (Test-Path -LiteralPath (Join-Path $Temp 'identified.pseudonym-key'))) 'Rejected identified run wrote files'
} finally {
    Remove-Item -LiteralPath $Temp -Recurse -Force -ErrorAction SilentlyContinue
}
Write-Output 'PASS: W365 pseudonymous and identified exports, assignment summaries, audit minimization, sanitized errors, CA target omission, manifest, protected CreateNew output and pre-collection gates; no Graph calls.'
