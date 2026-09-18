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
$NestedInstance = @{ settingDefinitionId = $RootId; groupSettingCollectionValue = @(@{ children = @(
    @{ settingDefinitionId = 'parent'; choiceSettingValue = @{ value = 'chosen'; settingValueTemplateReference = @{ useTemplateDefault = $true }; children = @($First.children[1]) } },
    $First.children[0]
) }) }
$NestedRules = @(ConvertTo-IntuneEpmRules $NestedInstance $Definitions 'policy' 'nested')
Assert-Epm ($NestedRules.Count -eq 1 -and $null -eq $NestedRules[0].fileName -and $NestedRules[0].name -ceq 'Rule one') 'Defaulted EPM ancestor produced explicit filename evidence or hid a sibling'
foreach ($TemplateCase in @(
    @{ Reference = $null; Explicit = $true },
    @{ Reference = @{ useTemplateDefault = $false }; Explicit = $true },
    @{ Reference = @{ useTemplateDefault = $true }; Explicit = $false },
    @{ Reference = @{}; Explicit = $false },
    @{ Reference = @{ useTemplateDefault = $null }; Explicit = $false },
    @{ Reference = @{ useTemplateDefault = 0 }; Explicit = $false },
    @{ Reference = @{ useTemplateDefault = 'false' }; Explicit = $false },
    @{ Reference = @{ useTemplateDefault = '' }; Explicit = $false },
    @{ Reference = @{ useTemplateDefault = @($false) }; Explicit = $false },
    @{ Reference = 'SECRET_TEMPLATE'; Explicit = $false },
    @{ Reference = @(@{ useTemplateDefault = $false }); Explicit = $false },
    @{ Reference = $false; Explicit = $false }
)) {
    foreach ($Placement in @('group', 'field', 'choice-parent', 'option-parent', 'option-field')) {
        foreach ($DecodeJson in @($false, $true)) {
            $Target = @{ settingDefinitionId = ($RootId + '_filename'); simpleSettingValue = @{ value = 'setup*.exe' } }
            $ParentDefinition = @{ id = 'parent'; options = @(@{ itemId = 'chosen'; optionValue = @{ value = 1 } }) }
            $CaseDefinitions = @($Definitions) + @($ParentDefinition)
            $Group = @{ children = @($Target, $First.children[0]) }
            switch ($Placement) {
                'group' { $Group.settingValueTemplateReference = $TemplateCase.Reference }
                'field' { $Target.simpleSettingValue.settingValueTemplateReference = $TemplateCase.Reference }
                { $_ -in @('choice-parent', 'option-parent') } {
                    $Parent = @{ settingDefinitionId = 'parent'; choiceSettingValue = @{ value = 'chosen'; children = @($Target) } }
                    if ($Placement -eq 'choice-parent') { $Parent.choiceSettingValue.settingValueTemplateReference = $TemplateCase.Reference }
                    else { $ParentDefinition.options[0].optionValue.settingValueTemplateReference = $TemplateCase.Reference }
                    $Group.children = @($Parent, $First.children[0])
                }
                'option-field' {
                    $Target.Remove('simpleSettingValue')
                    $Target.choiceSettingValue = @{ value = 'filename-choice' }
                    $CaseDefinitions = @($Definitions | Where-Object { $_.id -cne ($RootId + '_filename') }) + @(@{ id = ($RootId + '_filename'); offsetUri = '/PrivilegeManagement/ElevationRules/{0}/FileName'; options = @(@{ itemId = 'filename-choice'; optionValue = @{ value = 'setup*.exe'; settingValueTemplateReference = $TemplateCase.Reference } }) })
                }
            }
            $CaseInstance = @{ settingDefinitionId = $RootId; groupSettingCollectionValue = @($Group, @{ children = @(@{ settingDefinitionId = ($RootId + '_filename'); simpleSettingValue = @{ value = 'sibling.exe' } }) }) }
            if ($DecodeJson) {
                $CaseInstance = $CaseInstance | ConvertTo-Json -Depth 25 | ConvertFrom-Json
                $CaseDefinitions = $CaseDefinitions | ConvertTo-Json -Depth 25 | ConvertFrom-Json
            }
            $CaseRules = @(ConvertTo-IntuneEpmRules $CaseInstance $CaseDefinitions 'policy' 'template-matrix')
            Assert-Epm ($CaseRules.Count -eq 2 -and ($null -ne $CaseRules[0].fileName) -eq $TemplateCase.Explicit) "EPM template state changed: $Placement; JSON=$DecodeJson"
            if ($TemplateCase.Explicit) { Assert-Epm ($CaseRules[0].fileName -ceq 'setup*.exe') 'Explicit EPM filename changed' }
            if ($Placement -ne 'group' -or $TemplateCase.Explicit) { Assert-Epm ($CaseRules[0].name -ceq 'Rule one') 'EPM template context leaked into sibling field' }
            Assert-Epm ($CaseRules[1].fileName -ceq 'sibling.exe') 'EPM template context leaked into sibling rule group'
            Assert-Epm (-not (($CaseRules | ConvertTo-Json -Depth 10) -match 'SECRET_TEMPLATE|TemplateUnresolved|settingValueTemplateReference')) 'EPM traversal internals or raw template leaked'
        }
    }
}
$Grandchild = @{ settingDefinitionId = ($RootId + '_filename'); simpleSettingValue = @{ value = 'setup*.exe'; settingValueTemplateReference = @{ useTemplateDefault = $false } } }
$Middle = @{ settingDefinitionId = 'middle'; choiceSettingValue = @{ value = 'explicit'; settingValueTemplateReference = @{ useTemplateDefault = $false }; children = @($Grandchild) } }
$Ancestor = @{ settingDefinitionId = 'ancestor'; choiceSettingValue = @{ value = 'default'; settingValueTemplateReference = @{ useTemplateDefault = $true }; children = @($Middle) } }
foreach ($AddDuplicate in @($false, $true)) {
    $NestedGroup = @{ children = @($Ancestor, $First.children[0]) }
    if ($AddDuplicate) { $NestedGroup.children += $First.children[1] }
    $TransitiveRules = @(ConvertTo-IntuneEpmRules @{ settingDefinitionId = $RootId; groupSettingCollectionValue = @($NestedGroup) } $Definitions 'policy' 'transitive')
    Assert-Epm ($null -eq $TransitiveRules[0].fileName -and $TransitiveRules[0].name -ceq 'Rule one') 'EPM inherited uncertainty or unresolved duplicate was discarded'
}
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
$EpmRequest = {
    param($Address)
    $Rows = switch (([uri]$Address).AbsolutePath) {
        '/beta/deviceManagement/configurationPolicies' { @(@{ id = 'policy'; platforms = 'windows10'; technologies = 'endpointPrivilegeManagement'; settingCount = 1; templateReference = @{ templateFamily = 'endpointSecurityEndpointPrivilegeManagement' } }) }
        '/beta/deviceManagement/configurationPolicies/policy/settings' { @(@{ id = '0'; settingInstance = $script:EpmFixtureInstance }) }
        '/beta/deviceManagement/configurationPolicies/policy/settings/0/settingDefinitions' { $script:EpmFixtureDefinitions }
        default { @() }
    }
    @{ StatusCode = 200; Body = (@{ value = @($Rows) } | ConvertTo-Json -Depth 30 | ConvertFrom-Json) }
}
$EpmRequirements = @{ ScopeConfirmed = $true; Assessor = 'Offline EPM test'; ScopeDescription = 'Selected synthetic EPM policies'; MaxCollectionAgeHours = 24; ConfigurationReview = @{ PolicyIds = @('policy') } }
$Document = Invoke-IntuneDiscoveryCore -SelectedTenant '22222222-2222-4222-8222-222222222222' -Configuration $true -Request $EpmRequest -Requirements $EpmRequirements
Assert-Epm ($Document.Inventory.EpmRules.Count -eq 2 -and $Document.CollectionStatus.EpmRules.CompletedParentIds -contains 'policy') 'Core EPM grouped rules lost'
Assert-Epm (-not ($Document | ConvertTo-Json -Depth 30).Contains('SECRET')) 'Core EPM export leaked nested content'
if ($env:ASSAY_INTUNE_EPM_FIXTURE) { [IO.File]::WriteAllText($env:ASSAY_INTUNE_EPM_FIXTURE, ($Document | ConvertTo-Json -Depth 30), [Text.UTF8Encoding]::new($false)) }
$HiddenFields = @(
    @{ settingDefinitionId = ($RootId + '_filename'); simpleSettingValue = @{ value = 'SECRET_TEMPLATE_FILE.exe' } },
    @{ settingDefinitionId = ($RootId + '_filepath'); simpleSettingValue = @{ value = 'C:\SECRET_TEMPLATE_PATH' } }
)
$script:EpmFixtureInstance = @{ settingDefinitionId = $RootId; groupSettingCollectionValue = @(
    @{ children = @(
        $First.children[0],
        $First.children[3],
        @{ settingDefinitionId = 'template-parent'; choiceSettingValue = @{ value = 'chosen'; settingValueTemplateReference = @{ useTemplateDefault = $true }; children = $HiddenFields } }
    ) },
    @{ settingValueTemplateReference = @{ useTemplateDefault = 0 }; children = @($First.children[0], $First.children[3]) }
) }
$TemplateDocument = Invoke-IntuneDiscoveryCore -SelectedTenant '22222222-2222-4222-8222-222222222222' -Configuration $true -Request $EpmRequest -Requirements $EpmRequirements
Assert-Epm ($TemplateDocument.Inventory.EpmRules.Count -eq 2 -and $TemplateDocument.CollectionStatus.EpmRules.State -ceq 'Complete' -and $TemplateDocument.CollectionStatus.EpmRules.CompletedParentIds -contains 'policy') 'Unresolved EPM groups lost collection provenance'
$TemplateRows = $TemplateDocument.Inventory.EpmRules
Assert-Epm ($TemplateRows[0].elevationType -ceq 'Automatic' -and $TemplateRows[0].name -ceq 'Rule one' -and $null -eq $TemplateRows[0].fileName -and $null -eq $TemplateRows[0].filePath) 'Ancestor template context lost during production EPM import'
foreach ($Field in @('name', 'fileName', 'filePath', 'elevationType')) { Assert-Epm ($null -eq $TemplateRows[1][$Field]) 'Malformed group template exposed EPM field' }
$TemplateJson = $TemplateDocument | ConvertTo-Json -Depth 30
Assert-Epm (-not ($TemplateJson -match 'SECRET|TemplateUnresolved|settingValueTemplateReference')) 'Unresolved EPM payload or traversal marker leaked'
if ($env:ASSAY_INTUNE_EPM_TEMPLATE_FIXTURE) { [IO.File]::WriteAllText($env:ASSAY_INTUNE_EPM_TEMPLATE_FIXTURE, $TemplateJson, [Text.UTF8Encoding]::new($false)) }