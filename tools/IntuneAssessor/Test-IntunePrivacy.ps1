#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Invoke-IntuneDiscovery.ps1') -LibraryOnly
function Assert-Privacy([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function New-TestContext([string]$Mode) {
    [pscustomobject]@{ Mode = $Mode; Key = $(if ($Mode -eq 'Pseudonymous') { [byte[]](1..32) } else { $null }); KeyId = 'testkey'; KeyPath = $null; Identities = [Collections.Generic.Dictionary[string,string]]::new() }
}
$UserId = '11111111-aaaa-4aaa-8aaa-111111111111'
$GroupId = '22222222-bbbb-4bbb-8bbb-222222222222'
$RoleId = '33333333-cccc-4ccc-8ccc-333333333333'
$Request = {
    param($Uri)
    $Values = @()
    if ($Uri.EndsWith('/managedDevices')) { $Values = @(@{ id = 'device-1'; operatingSystem = 'Windows'; deviceName = 'PII-SENTINEL-LAPTOP'; azureADDeviceId = '44444444-dddd-4ddd-8ddd-444444444444'; managementAgent = 'mdm' }) }
    if ($Uri.EndsWith('/deviceCompliancePolicies')) { $Values = @(@{ id = 'policy-1'; displayName = 'Policy' }) }
    if ($Uri.EndsWith('/policy-1/deviceStatuses')) { $Values = @(@{ id = 'report-1'; deviceDisplayName = 'PII-SENTINEL-LAPTOP'; status = 'compliant' }) }
    if ($Uri.EndsWith('/conditionalAccess/policies')) {
        $Values = @(@{ id = 'ca-1'; displayName = 'Require compliant'; state = 'enabled'; conditions = @{
            users = @{ includeUsers = @('All'); excludeUsers = @($UserId, 'GuestsOrExternalUsers'); includeGroups = @($GroupId); excludeGroups = @(); includeRoles = @($RoleId); excludeRoles = @() }
            applications = @{ includeApplications = @('All') }
        }; grantControls = @{ operator = 'AND'; builtInControls = @('compliantDevice') } })
    }
    if ($Uri.EndsWith('/policies/deviceRegistrationPolicy')) {
        return @{ StatusCode = 200; Body = @{ id = 'deviceRegistrationPolicy'; userDeviceQuota = 20; localAdminPassword = @{ isEnabled = $true }; azureADJoin = @{ isAdminConfigurable = $true; allowedToJoin = @{ '@odata.type' = '#microsoft.graph.enumeratedDeviceRegistrationMembership'; users = @($UserId); groups = @($GroupId) } } } }
    }
    return @{ StatusCode = 200; Body = @{ value = $Values } }
}
function Invoke-TestCore($Context) {
    Invoke-IntuneDiscoveryCore -SelectedTenant '55555555-eeee-4eee-8eee-555555555555' -Rbac $false -Audit $false -Entra $true -Request $Request -PrivacyContext $Context
}
$Sentinels = "PII-SENTINEL|PII_SENTINEL|$UserId|$GroupId|$RoleId|fileserver|LapsAdminSentinel"

$Pseudo = Invoke-TestCore (New-TestContext 'Pseudonymous')
$Json = $Pseudo | ConvertTo-Json -Depth 30
Assert-Privacy ($Json -notmatch $Sentinels) ('Pseudonymous Intune export leaked: ' + [regex]::Match($Json, $Sentinels).Value)
$DeviceName = $Pseudo.Inventory.ManagedDevices[0]['deviceName']
Assert-Privacy ($DeviceName -cmatch '^dev_[0-9a-f]{16}$' -and $Pseudo.Inventory.ComplianceStates[0]['deviceDisplayName'] -ceq $DeviceName) 'Device names not pseudonymized consistently'
Assert-Privacy ($Pseudo.Inventory.ManagedDevices[0]['azureADDeviceId'] -ceq '44444444-dddd-4ddd-8ddd-444444444444') 'Join key removed'
$Observation = @($Pseudo.Observations | Where-Object { $_.Module -eq 'ManagedDevices' })[0]
Assert-Privacy ($Observation.Data['deviceName'] -ceq $DeviceName) 'Observation data did not inherit the transformed row'
$Users = $Pseudo.Inventory.ConditionalAccessPolicies[0]['conditions']['users']
Assert-Privacy ((@($Users['includeUsers']) -join ',') -ceq 'All' -and $Users['includeUsersCount'] -eq 0 -and (@($Users['excludeUsers']) -join ',') -ceq 'GuestsOrExternalUsers' -and $Users['excludeUsersCount'] -eq 1) 'CA user lists not reduced to tokens and counts'
Assert-Privacy (@($Users['includeGroups']).Count -eq 0 -and $Users['includeGroupsCount'] -eq 1 -and $Users['includeRolesCount'] -eq 1 -and $Users['excludeGroupsCount'] -eq 0) 'CA group or role lists not counted'
Assert-Privacy ($Pseudo.Inventory.ConditionalAccessPolicies[0]['grantControls']['builtInControls'][0] -ceq 'compliantDevice') 'CA grant controls changed'
$Join = $Pseudo.Inventory.DeviceRegistrationPolicy[0]['azureADJoin']['allowedToJoin']
Assert-Privacy ($Join['usersCount'] -eq 1 -and $Join['groupsCount'] -eq 1 -and @($Join['users']).Count -eq 0 -and $Pseudo.Inventory.DeviceRegistrationPolicy[0]['localAdminPassword']['isEnabled'] -eq $true) 'Device registration members not counted'

$Identified = Invoke-TestCore (New-TestContext 'Identified')
$IdentifiedJson = $Identified | ConvertTo-Json -Depth 30
Assert-Privacy ($IdentifiedJson.Contains('PII-SENTINEL-LAPTOP') -and $IdentifiedJson.Contains($UserId) -and $IdentifiedJson.Contains($GroupId)) 'Identified Intune values not retained'
$Library = Invoke-TestCore $null
Assert-Privacy ($Library.Inventory.ManagedDevices[0]['deviceName'] -ceq 'PII-SENTINEL-LAPTOP') 'Library callers without a privacy context changed'

$Inventory = [ordered]@{
    SecuritySettings = @(
        [ordered]@{ id = 's1'; cspUri = 'LAPS/Policies/AdministratorAccountName'; value = 'LapsAdminSentinel'; resolution = 'Resolved' },
        [ordered]@{ id = 's2'; cspUri = 'LAPS/Policies/AdministratorAccountName'; value = 'administrator'; resolution = 'Resolved' },
        [ordered]@{ id = 's3'; cspUri = 'LAPS/Policies/AdministratorAccountName'; value = ''; resolution = 'Resolved' },
        [ordered]@{ id = 's4'; cspUri = 'Policy/Config/Defender/AttackSurfaceReductionOnlyExclusions'; value = 'C:\Users\PII_SENTINEL\tool.exe|\\fileserver\share\*.ps1|D:\Build\*'; resolution = 'Resolved' },
        [ordered]@{ id = 's5'; cspUri = 'Policy/Config/Defender/ExcludedPaths'; value = 'C:\Users\PII_SENTINEL\AppData'; resolution = 'Resolved' },
        [ordered]@{ id = 's6'; cspUri = 'Policy/Config/Defender/ExcludedExtensions'; value = '.tmp'; resolution = 'Resolved' },
        [ordered]@{ id = 's7'; cspUri = 'LAPS/Policies/PasswordLength'; value = 14; resolution = 'Resolved' }
    )
    EpmRules = @(
        [ordered]@{ id = 'r1'; filePath = 'C:\Users\PII_SENTINEL\Downloads'; fileName = 'setup.exe'; elevationType = 'Automatic' },
        [ordered]@{ id = 'r2'; filePath = '\\fileserver\apps'; fileName = '*.exe'; elevationType = 'Self' },
        [ordered]@{ id = 'r3'; filePath = '\\?\C:\Program Files\Tool'; fileName = 'tool.exe'; elevationType = 'Self' },
        [ordered]@{ id = 'r4'; filePath = ''; fileName = $null; elevationType = 'Self' }
    )
}
ConvertTo-IntunePrivacyInventory $Inventory (New-TestContext 'Pseudonymous')
$Values = @($Inventory.SecuritySettings | ForEach-Object { $_['value'] })
Assert-Privacy ($Values[0] -ceq 'Custom' -and $Values[1] -ceq 'BuiltIn' -and $Values[2] -ceq 'NotConfigured') 'LAPS account name not classified'
Assert-Privacy ($Values[3] -ceq 'C:\Users\{profile}\tool.exe|\\{host}\{share}\*.ps1|D:\Build\*' -and $Values[4] -ceq 'C:\Users\{profile}\AppData' -and $Values[5] -ceq '.tmp' -and $Values[6] -eq 14) 'Defender exclusions not masked or unrelated values changed'
Assert-Privacy ($Inventory.EpmRules[0]['filePath'] -ceq 'C:\Users\{profile}\Downloads' -and $Inventory.EpmRules[1]['filePath'] -ceq '\\{host}\{share}' -and $Inventory.EpmRules[1]['fileName'] -ceq '*.exe') 'EPM path or wildcard handling incorrect'
Assert-Privacy ($Inventory.EpmRules[2]['filePath'] -ceq '\\?\C:\Program Files\Tool' -and $Inventory.EpmRules[3]['filePath'] -ceq '' -and $null -eq $Inventory.EpmRules[3]['fileName']) 'EPM device-namespace or empty path changed'

$Temp = Join-Path ([IO.Path]::GetTempPath()) ('intune-privacy-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($Temp)
try {
    foreach ($Companion in @(
        @{ File = 'Get-IntuneEndpointEvidence.ps1'; Statements = { param($Ast) @($Ast.EndBlock.Statements) } },
        @{ File = 'Get-IntuneDefenderEvidence.ps1'; Statements = { param($Ast) @($Ast.Find({ param($Node) $Node -is [Management.Automation.Language.TryStatementAst] -and $Node.Body.Extent.Text.Contains('Write-CollectorProtectedFile') }, $true).Body.Statements) } }
    )) {
        $Tokens = $null; $Errors = $null
        $CompanionAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot $Companion.File), [ref]$Tokens, [ref]$Errors)
        $All = & $Companion.Statements $CompanionAst
        $Start = [array]::FindIndex($All, [Predicate[object]]{ param($Statement) $Statement.Extent.Text.StartsWith('$Manifest =') })
        $End = [array]::FindIndex($All, [Predicate[object]]{ param($Statement) $Statement.Extent.Text.Contains('Write-CollectorProtectedFile') })
        Assert-Privacy ($Start -ge 0 -and $End -gt $Start) ('Companion export statements not found: ' + $Companion.File)
        $Block = [scriptblock]::Create((($All[$Start..$End] | ForEach-Object { $_.Extent.Text }) -join "`n"))
        $OutputPath = Join-Path $Temp ($Companion.File + '.json')
        $Evidence = @{ SchemaVersion = '1.0'; TenantId = 't'; DeviceId = 'd'; Modules = @{} }
        $TenantId = 't'; $Started = 'now'; $Result = @{ Rows = @(); State = 'Complete' }
        . $Block
        $Written = Get-Content -LiteralPath $OutputPath -Raw | ConvertFrom-Json
        Assert-Privacy ($Written.Privacy.Classification -ceq 'Confidential' -and @($Written.Privacy.PseudonymizedFieldClasses).Count -eq 0 -and (Get-Acl -LiteralPath $OutputPath).AreAccessRulesProtected) ('Companion manifest or ACL missing: ' + $Companion.File)
        $Rejected = $false
        try { . $Block } catch { $Rejected = $true }
        Assert-Privacy $Rejected ('Companion overwrote output: ' + $Companion.File)
    }
    $Collector = Join-Path $PSScriptRoot 'Invoke-IntuneDiscovery.ps1'
    foreach ($Case in @(
        @{ Arguments = @{ TenantId = '55555555-eeee-4eee-8eee-555555555555'; PrivacyMode = 'Identified'; OutputPath = (Join-Path $Temp 'identified.json') }; Pattern = 'ConfirmIdentifiedExport' },
        @{ Arguments = @{ TenantId = '55555555-eeee-4eee-8eee-555555555555'; OutputPath = $OutputPath }; Pattern = 'already exists' }
    )) {
        $Arguments = $Case.Arguments
        $Message = $null
        try { & $Collector @Arguments *> $null } catch { $Message = $_.Exception.Message }
        Assert-Privacy ($Message -match $Case.Pattern) ('Collector did not stop before collection: ' + $Case.Pattern)
    }
    Assert-Privacy (-not (Test-Path -LiteralPath (Join-Path $Temp 'identified.json'))) 'Rejected identified run wrote output'
} finally {
    Remove-Item -LiteralPath $Temp -Recurse -Force -ErrorAction SilentlyContinue
}
Write-Output 'PASS: Intune device pseudonyms, inherited observations, CA and registration member counts, LAPS classification, exclusion/EPM path masking, companion manifests/ACLs and pre-collection gates.'
