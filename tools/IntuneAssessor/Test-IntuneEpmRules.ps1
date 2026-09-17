#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Invoke-IntuneDiscovery.ps1') -LibraryOnly
function Assert-Epm([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
$RootId = 'device_vendor_msft_policy_privilegemanagement_elevationrules_{elevationrulename}'
$Definitions = @(@{ id = $RootId; offsetUri = '/PrivilegeManagement/ElevationRules/{0}' })
foreach ($Pair in @(@('name','Name'),@('filename','FileName'),@('filepath','FilePath'),@('ruletype','RuleType'))) {
    $Definitions += @{ id = ($RootId + '_' + $Pair[0]); offsetUri = ('/PrivilegeManagement/ElevationRules/{0}/' + $Pair[1]); options = @(@{ itemId = 'opaque-option'; optionValue = @{ value = 'Automatic' } }) }
}
$First = @{ children = @(
    @{ settingDefinitionId = ($RootId + '_name'); simpleSettingValue = @{ value = 'Rule one' } },
    @{ settingDefinitionId = ($RootId + '_filename'); simpleSettingValue = @{ value = 'setup*.exe' } },
    @{ settingDefinitionId = ($RootId + '_filepath'); simpleSettingValue = @{ value = '\\server\share' } },
    @{ settingDefinitionId = ($RootId + '_ruletype'); choiceSettingValue = @{ value = 'opaque-option'; children = @(@{ settingDefinitionId = 'unknown'; simpleSettingValue = @{ value = 'SECRET' } }) } }
) }
$Second = @{ children = @(
    @{ settingDefinitionId = ($RootId + '_filename'); simpleSettingValue = @{ value = 'other.exe' } },
    @{ settingDefinitionId = ($RootId + '_ruletype'); choiceSettingValue = @{ value = 'opaque-option'; settingValueTemplateReference = @{ useTemplateDefault = $true } } }
) }
$Instance = @{ settingDefinitionId = $RootId; groupSettingCollectionValue = @($First, $Second) }
$Rules = @(ConvertTo-IntuneEpmRules ($Instance | ConvertTo-Json -Depth 20 | ConvertFrom-Json) ($Definitions | ConvertTo-Json -Depth 20 | ConvertFrom-Json) 'policy' '0')
Assert-Epm ($Rules.Count -eq 2 -and $Rules[0].id -ne $Rules[1].id) 'Rule group identity collapsed'
Assert-Epm ($Rules[0].fileName -eq 'setup*.exe' -and $Rules[0].elevationType -ceq 'Automatic') 'Exact definition/choice binding failed'
Assert-Epm ($null -eq $Rules[1].filePath -and $null -eq $Rules[1].elevationType) 'Missing/default values inherited from another rule'
Assert-Epm (-not ($Rules | ConvertTo-Json -Depth 20).Contains('SECRET')) 'Unreviewed EPM rule content leaked'
$First.settingValueTemplateReference = @{ useTemplateDefault = $true }
$DefaultRules = @(ConvertTo-IntuneEpmRules $Instance $Definitions 'policy' '0')
Assert-Epm ($null -eq $DefaultRules[0].fileName) 'Group-level template default was treated as configured'
$First.Remove('settingValueTemplateReference')
$First.children += $First.children[1]
$Rules = @(ConvertTo-IntuneEpmRules $Instance $Definitions 'policy' '0')
Assert-Epm ($null -eq $Rules[0].fileName) 'Duplicate field resolved arbitrarily'
Write-Output 'PASS: isolated EPM rule groups, definition/option identity, template defaults, duplicate fields and safe projection.'
$First.children = $First.children[0..3]
$script:EpmFixtureDefinitions = $Definitions
$script:EpmFixtureInstance = $Instance
$Document = Invoke-IntuneDiscoveryCore -SelectedTenant '22222222-2222-4222-8222-222222222222' -Configuration $true -Request {
    param($Address)
    $Rows = switch (([uri]$Address).AbsolutePath) {
        '/beta/deviceManagement/configurationPolicies' { @(@{ id = 'policy'; platforms = 'windows10'; technologies = 'endpointPrivilegeManagement'; settingCount = 1; templateReference = @{ templateFamily = 'endpointSecurityEndpointPrivilegeManagement' } }) }
        '/beta/deviceManagement/configurationPolicies/policy/settings' { @(@{ id = '0'; settingInstance = $script:EpmFixtureInstance }) }
        '/beta/deviceManagement/configurationPolicies/policy/settings/0/settingDefinitions' { $script:EpmFixtureDefinitions }
        default { @() }
    }
    @{ StatusCode = 200; Body = (@{ value = @($Rows) } | ConvertTo-Json -Depth 30 | ConvertFrom-Json) }
} -Requirements @{ ScopeConfirmed = $true; Assessor = 'Offline EPM test'; ScopeDescription = 'Selected synthetic EPM policies'; MaxCollectionAgeHours = 24; ConfigurationReview = @{ PolicyIds = @('policy') } }
Assert-Epm ($Document.Inventory.EpmRules.Count -eq 2 -and $Document.CollectionStatus.EpmRules.CompletedParentIds -contains 'policy') 'Core EPM grouped rules lost'
Assert-Epm (-not ($Document | ConvertTo-Json -Depth 30).Contains('SECRET')) 'Core EPM export leaked nested content'
if ($env:ASSAY_INTUNE_EPM_FIXTURE) { [IO.File]::WriteAllText($env:ASSAY_INTUNE_EPM_FIXTURE, ($Document | ConvertTo-Json -Depth 30), [Text.UTF8Encoding]::new($false)) }