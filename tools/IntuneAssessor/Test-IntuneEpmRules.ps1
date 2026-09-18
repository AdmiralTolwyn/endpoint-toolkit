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
$MissingChoiceInstance = @{ settingDefinitionId = $RootId; groupSettingCollectionValue = @(@{ children = @(
    @{ settingDefinitionId = 'parent'; choiceSettingValue = @{ value = 'missing-option'; children = @($First.children[1]) } },
    $First.children[0]
) }) }
$MissingChoiceRules = @(ConvertTo-IntuneEpmRules $MissingChoiceInstance ($Definitions + @(@{ id = 'parent'; options = @(@{ itemId = 'known'; optionValue = @{ value = 1 } }) })) 'policy' 'missing-choice')
Assert-Epm ($MissingChoiceRules.Count -eq 1 -and $null -eq $MissingChoiceRules[0].fileName -and $MissingChoiceRules[0].name -ceq 'Rule one') 'Unresolved EPM ancestor choice produced explicit child evidence or hid a sibling'
foreach ($ChoiceCase in @('valid', 'missing-definition', 'duplicate-definition', 'missing-option', 'duplicate-option', 'numeric-choice', 'array-choice', 'null-choice', 'blank-choice', 'case-choice', 'numeric-option', 'array-option', 'missing-payload', 'scalar-payload', 'array-payload', 'mixed-values', 'template-option')) {
    foreach ($Placement in @('field', 'ancestor')) {
        foreach ($DecodeJson in @($false, $true)) {
            $ChoiceId = if ($Placement -eq 'field') { $RootId + '_filename' } else { 'choice-parent' }
            $ChoiceDefinition = @{ id = $ChoiceId; offsetUri = '/PrivilegeManagement/ElevationRules/{0}/FileName'; options = @(@{ itemId = 'chosen'; optionValue = @{ value = 'setup*.exe' } }) }
            $ChoiceNode = @{ settingDefinitionId = $ChoiceId; choiceSettingValue = @{ value = 'chosen' } }
            if ($Placement -eq 'ancestor') { $ChoiceNode.choiceSettingValue.children = @($First.children[1]) }
            switch ($ChoiceCase) {
                'missing-option' { $ChoiceNode.choiceSettingValue.value = 'missing' }
                'duplicate-option' { $ChoiceDefinition.options += @{ itemId = 'chosen'; optionValue = @{ value = 'other.exe' } } }
                'numeric-choice' { $ChoiceNode.choiceSettingValue.value = 1; $ChoiceDefinition.options[0].itemId = '1' }
                'array-choice' { $ChoiceNode.choiceSettingValue.value = @('chosen') }
                'null-choice' { $ChoiceNode.choiceSettingValue.value = $null; $ChoiceDefinition.options[0].itemId = $null }
                'blank-choice' { $ChoiceNode.choiceSettingValue.value = ''; $ChoiceDefinition.options[0].itemId = '' }
                'case-choice' { $ChoiceNode.choiceSettingValue.value = 'CHOSEN' }
                'numeric-option' { $ChoiceNode.choiceSettingValue.value = '1'; $ChoiceDefinition.options[0].itemId = 1 }
                'array-option' { $ChoiceDefinition.options[0].itemId = @('chosen') }
                'missing-payload' { $ChoiceDefinition.options[0].Remove('optionValue') }
                'scalar-payload' { $ChoiceDefinition.options[0].optionValue = 'SECRET_RAW_VALUE' }
                'array-payload' { $ChoiceDefinition.options[0].optionValue = @(@{ value = 'setup*.exe' }) }
                'mixed-values' { $ChoiceNode.simpleSettingValue = @{ value = 'SECRET_SIMPLE_VALUE' } }
                'template-option' { $ChoiceDefinition.options[0].optionValue.settingValueTemplateReference = @{ useTemplateDefault = $true } }
            }
            $ChoiceDefinitions = @($Definitions | Where-Object { $_.id -cne $ChoiceId })
            if ($ChoiceCase -ne 'missing-definition') { $ChoiceDefinitions += $ChoiceDefinition }
            if ($ChoiceCase -eq 'duplicate-definition') { $ChoiceDefinitions += $ChoiceDefinition.Clone() }
            $ChoiceInstance = @{ settingDefinitionId = $RootId; groupSettingCollectionValue = @(
                @{ children = @($ChoiceNode, $First.children[0]) },
                @{ children = @($First.children[0]) }
            ) }
            if ($DecodeJson) {
                $ChoiceInstance = $ChoiceInstance | ConvertTo-Json -Depth 20 | ConvertFrom-Json
                $ChoiceDefinitions = $ChoiceDefinitions | ConvertTo-Json -Depth 20 | ConvertFrom-Json
            }
            $ChoiceRules = @(ConvertTo-IntuneEpmRules $ChoiceInstance $ChoiceDefinitions 'policy' 'choice-matrix')
            Assert-Epm ($ChoiceRules.Count -eq 2 -and ($null -ne $ChoiceRules[0].fileName) -eq ($ChoiceCase -eq 'valid')) "EPM choice resolution changed: $ChoiceCase; $Placement; JSON=$DecodeJson"
            if ($ChoiceCase -eq 'valid') { Assert-Epm ($ChoiceRules[0].fileName -ceq 'setup*.exe') 'Valid EPM option did not decode its own value' }
            Assert-Epm ($ChoiceRules[0].name -ceq 'Rule one' -and $ChoiceRules[1].name -ceq 'Rule one') 'Choice uncertainty leaked into independent EPM evidence'
            Assert-Epm (-not (($ChoiceRules | ConvertTo-Json -Depth 10) -match 'SECRET|ChoiceUnresolved|TemplateUnresolved')) 'EPM choice traversal metadata or raw value leaked'
        }
    }
}
$ExplicitLeaf = @{ settingDefinitionId = ($RootId + '_filename'); simpleSettingValue = @{ value = 'setup*.exe' } }
$ResolvedMiddle = @{ settingDefinitionId = 'middle'; choiceSettingValue = @{ value = 'known'; children = @($ExplicitLeaf) } }
$UnknownAncestor = @{ settingDefinitionId = 'ancestor'; choiceSettingValue = @{ value = 'missing'; children = @($ResolvedMiddle) } }
$ChainDefinitions = $Definitions + @(@{ id = 'middle'; options = @(@{ itemId = 'known'; optionValue = @{ value = 1 } }) })
foreach ($DuplicateLeaf in @($false, $true)) {
    $ChainChildren = @($UnknownAncestor, $First.children[0])
    if ($DuplicateLeaf) { $ChainChildren += $ExplicitLeaf }
    $ChainRules = @(ConvertTo-IntuneEpmRules @{ settingDefinitionId = $RootId; groupSettingCollectionValue = @(@{ children = $ChainChildren }) } $ChainDefinitions 'policy' 'choice-chain')
    Assert-Epm ($null -eq $ChainRules[0].fileName -and $ChainRules[0].name -ceq 'Rule one') 'Resolved EPM descendant or explicit duplicate cleared ancestor uncertainty'
}
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
$script:EpmFixtureDefinitions = $Definitions + @(@{ id = 'template-parent'; options = @(@{ itemId = 'chosen'; optionValue = @{ value = 1 } }) })
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
$script:EpmFixtureDefinitions = $Definitions + @(@{ id = 'choice-parent'; options = @(@{ itemId = 'known'; optionValue = @{ value = 1 } }) })
$script:EpmFixtureInstance = @{ settingDefinitionId = $RootId; groupSettingCollectionValue = @(
    @{ children = @(
        $First.children[0],
        $First.children[3],
        @{ settingDefinitionId = 'choice-parent'; choiceSettingValue = @{ value = 'missing'; children = $HiddenFields } }
    ) },
    @{ children = @(
        $First.children[0],
        @{ settingDefinitionId = ($RootId + '_filename'); simpleSettingValue = @{ value = 'other.exe' } },
        @{ settingDefinitionId = ($RootId + '_ruletype'); choiceSettingValue = @{ value = 'opaque-option' }; simpleSettingValue = @{ value = 'Deny' } }
    ) }
) }
$ChoiceDocument = Invoke-IntuneDiscoveryCore -SelectedTenant '22222222-2222-4222-8222-222222222222' -Configuration $true -Request $EpmRequest -Requirements $EpmRequirements
Assert-Epm ($ChoiceDocument.Inventory.EpmRules.Count -eq 2 -and $ChoiceDocument.CollectionStatus.EpmRules.State -ceq 'Complete' -and $ChoiceDocument.CollectionStatus.EpmRules.CompletedParentIds -contains 'policy') 'Unresolved choice groups lost EPM provenance'
$ChoiceRows = $ChoiceDocument.Inventory.EpmRules
Assert-Epm ($ChoiceRows[0].name -ceq 'Rule one' -and $ChoiceRows[0].elevationType -ceq 'Automatic' -and $null -eq $ChoiceRows[0].fileName -and $null -eq $ChoiceRows[0].filePath) 'Unresolved ancestor choice exposed EPM descendant values'
Assert-Epm ($ChoiceRows[1].name -ceq 'Rule one' -and $ChoiceRows[1].fileName -ceq 'other.exe' -and $null -eq $ChoiceRows[1].elevationType) 'Mixed leaf choice exposed EPM elevation type or hid sibling evidence'
$ChoiceJson = $ChoiceDocument | ConvertTo-Json -Depth 30
Assert-Epm (-not ($ChoiceJson -match 'SECRET|ChoiceUnresolved|TemplateUnresolved|choiceSettingValue')) 'Raw EPM choice payload or traversal marker exported'
if ($env:ASSAY_INTUNE_EPM_CHOICE_FIXTURE) { [IO.File]::WriteAllText($env:ASSAY_INTUNE_EPM_CHOICE_FIXTURE, $ChoiceJson, [Text.UTF8Encoding]::new($false)) }