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
$RootValueCases = foreach ($Kind in @('choiceSettingValue', 'simpleSettingValue', 'choiceSettingCollectionValue', 'simpleSettingCollectionValue')) {
    foreach ($Shape in @('absent', 'null', 'empty-object', 'object', 'empty-array', 'array', 'false', 'zero', 'blank', 'text')) {
        @{ Kind = $Kind; Shape = $Shape }
    }
}
foreach ($RootValueCase in $RootValueCases) {
    $ExtraKind = $RootValueCase.Kind
    foreach ($DecodeJson in @($false, $true)) {
        $MixedRoot = @{ settingDefinitionId = $RootId; groupSettingCollectionValue = @($First, $Second) }
        switch ($RootValueCase.Shape) {
            'null' { $MixedRoot[$ExtraKind] = $null }
            'empty-object' { $MixedRoot[$ExtraKind] = @{} }
            'object' { $MixedRoot[$ExtraKind] = @{ value = 'SECRET_ROOT_VALUE' } }
            'empty-array' { $MixedRoot[$ExtraKind] = @() }
            'array' { $MixedRoot[$ExtraKind] = @(@{ value = 'SECRET_ROOT_VALUE' }) }
            'false' { $MixedRoot[$ExtraKind] = $false }
            'zero' { $MixedRoot[$ExtraKind] = 0 }
            'blank' { $MixedRoot[$ExtraKind] = '' }
            'text' { $MixedRoot[$ExtraKind] = 'SECRET_ROOT_VALUE' }
        }
        if ($DecodeJson) { $MixedRoot = $MixedRoot | ConvertTo-Json -Depth 20 | ConvertFrom-Json }
        $MixedRootRejected = $false
        $MixedRootRows = @()
        try { $MixedRootRows = @(ConvertTo-IntuneEpmRules $MixedRoot $Definitions 'policy' 'mixed-root-value') }
        catch { $MixedRootRejected = $true }
        $ShouldReject = $RootValueCase.Shape -notin @('absent', 'null')
        Assert-Epm ($MixedRootRejected -eq $ShouldReject) "Mixed EPM root handling changed: $ExtraKind; $($RootValueCase.Shape); JSON=$DecodeJson"
        if (-not $ShouldReject) { Assert-Epm ($MixedRootRows.Count -eq 2 -and $MixedRootRows[0].fileName -ceq 'setup*.exe' -and $MixedRootRows[0].elevationType -ceq 'Automatic') 'Absent/null alternate root kind lost valid group evidence' }
        Assert-Epm (-not (($MixedRootRows | ConvertTo-Json -Depth 10) -match 'SECRET')) 'Mixed root payload leaked'
    }
}
foreach ($NestedKind in @('groupSettingCollectionValue', 'choiceSettingCollectionValue')) {
    foreach ($DecodeJson in @($false, $true)) {
        $NestedValue = @{ children = @($First.children[1]); value = 'known' }
        $NestedNode = @{ settingDefinitionId = 'nested-container'; $NestedKind = @($NestedValue) }
        $NestedDefinitions = $Definitions + @(@{ id = 'nested-container'; options = @(@{ itemId = 'known'; optionValue = @{ value = 1 } }) })
        $NestedInstance = @{ settingDefinitionId = $RootId; groupSettingCollectionValue = @(@{ children = @($First.children[0], $First.children[1], $NestedNode) }) }
        if ($DecodeJson) {
            $NestedDefinitions = $NestedDefinitions | ConvertTo-Json -Depth 20 | ConvertFrom-Json
            $NestedInstance = $NestedInstance | ConvertTo-Json -Depth 20 | ConvertFrom-Json
        }
        $NestedRows = @(ConvertTo-IntuneEpmRules $NestedInstance $NestedDefinitions 'policy' 'nested-duplicate')
        Assert-Epm ($NestedRows.Count -eq 1 -and $null -eq $NestedRows[0].fileName -and $NestedRows[0].name -ceq 'Rule one') "Nested collection concealed an EPM duplicate field: $NestedKind; JSON=$DecodeJson"
    }
}
$CollectionCases = foreach ($Kind in @('groupSettingCollectionValue', 'choiceSettingCollectionValue')) {
    foreach ($Case in @('valid', 'duplicate', 'default', 'default-duplicate', 'malformed-template', 'mixed-simple', 'mixed-choice', 'mixed-simple-collection', 'ancestor-template', 'ancestor-choice', 'container-object', 'container-string', 'item-array', 'item-string', 'item-null', 'children-object')) {
        @{ Kind = $Kind; Case = $Case }
    }
}
foreach ($CollectionCase in $CollectionCases) {
    foreach ($DecodeJson in @($false, $true)) {
        $Kind = $CollectionCase.Kind
        $Case = $CollectionCase.Case
        $CollectionValue = @{ children = @($First.children[1]); value = 'known' }
        $CollectionNode = @{ settingDefinitionId = 'collection'; $Kind = @($CollectionValue) }
        $CollectionDefinitions = $Definitions + @(@{ id = 'collection'; options = @(@{ itemId = 'known'; optionValue = @{ value = 1 } }) })
        switch ($Case) {
            'default' { $CollectionValue.settingValueTemplateReference = @{ useTemplateDefault = $true } }
            'default-duplicate' { $CollectionValue.settingValueTemplateReference = @{ useTemplateDefault = $true } }
            'malformed-template' { $CollectionValue.settingValueTemplateReference = @{ useTemplateDefault = 0 } }
            'mixed-simple' { $CollectionNode.simpleSettingValue = @{ value = 'SECRET' } }
            'mixed-choice' { $CollectionNode.choiceSettingValue = @{ value = 'known' } }
            'mixed-simple-collection' { $CollectionNode.simpleSettingCollectionValue = @(@{ value = 'SECRET' }) }
            'container-object' { $CollectionNode[$Kind] = $CollectionValue }
            'container-string' { $CollectionNode[$Kind] = 'SECRET' }
            'item-array' { $CollectionNode[$Kind] = ,@($CollectionValue) }
            'item-string' { $CollectionNode[$Kind] = @('SECRET') }
            'item-null' { $CollectionNode[$Kind] = @($null) }
            'children-object' { $CollectionValue.children = $First.children[1] }
        }
        $OuterNode = $CollectionNode
        if ($Case -in @('ancestor-template', 'ancestor-choice')) {
            $OuterNode = @{ settingDefinitionId = 'outer'; choiceSettingValue = @{ value = 'known'; children = @($CollectionNode) } }
            if ($Case -eq 'ancestor-template') {
                $OuterNode.choiceSettingValue.settingValueTemplateReference = @{ useTemplateDefault = $true }
                $CollectionDefinitions += @{ id = 'outer'; options = @(@{ itemId = 'known'; optionValue = @{ value = 1 } }) }
            }
        }
        $CollectionChildren = @($First.children[0], $OuterNode)
        if ($Case -in @('duplicate', 'default-duplicate')) { $CollectionChildren += $First.children[1] }
        $CollectionInstance = @{ settingDefinitionId = $RootId; groupSettingCollectionValue = @(
            @{ children = $CollectionChildren },
            @{ children = @($First.children[1]) }
        ) }
        if ($DecodeJson) {
            $CollectionInstance = $CollectionInstance | ConvertTo-Json -Depth 25 | ConvertFrom-Json
            $CollectionDefinitions = $CollectionDefinitions | ConvertTo-Json -Depth 25 | ConvertFrom-Json
        }
        $CollectionRejected = $false
        $CollectionRows = @()
        try { $CollectionRows = @(ConvertTo-IntuneEpmRules $CollectionInstance $CollectionDefinitions 'policy' 'collection-matrix') }
        catch { $CollectionRejected = $true }
        $Malformed = $Case -in @('container-object', 'container-string', 'item-array', 'item-string', 'item-null', 'children-object')
        Assert-Epm ($CollectionRejected -eq $Malformed) "Nested EPM shape rejection changed: $Kind; $Case; JSON=$DecodeJson"
        if (-not $Malformed) {
            Assert-Epm ($CollectionRows.Count -eq 2 -and $CollectionRows[0].name -ceq 'Rule one' -and $CollectionRows[1].fileName -ceq 'setup*.exe') 'Nested EPM context leaked into independent evidence'
            Assert-Epm (($null -ne $CollectionRows[0].fileName) -eq ($Case -eq 'valid')) "Nested EPM context was lost: $Kind; $Case; JSON=$DecodeJson"
            if ($Case -eq 'valid') { Assert-Epm ($CollectionRows[0].fileName -ceq 'setup*.exe') 'Explicit nested EPM value changed' }
        }
        Assert-Epm (-not (($CollectionRows | ConvertTo-Json -Depth 10) -match 'SECRET|ChoiceUnresolved|TemplateUnresolved')) 'Nested EPM raw values or internal context exported'
    }
}
foreach ($ChoiceCase in @('valid', 'missing-option', 'duplicate-option', 'default-option', 'missing-definition', 'duplicate-definition')) {
    foreach ($DecodeJson in @($false, $true)) {
        $CollectionDefinition = @{ id = 'choice-collection'; options = @(
            @{ itemId = 'first'; optionValue = @{ value = 1 } },
            @{ itemId = 'second'; optionValue = @{ value = 2 } }
        ) }
        $CollectionNode = @{ settingDefinitionId = 'choice-collection'; choiceSettingCollectionValue = @(
            @{ value = 'first'; children = @($First.children[1]) },
            @{ value = 'second'; children = @($First.children[2]) }
        ) }
        switch ($ChoiceCase) {
            'missing-option' { $CollectionNode.choiceSettingCollectionValue[0].value = 'missing' }
            'duplicate-option' { $CollectionDefinition.options += @{ itemId = 'first'; optionValue = @{ value = 3 } } }
            'default-option' { $CollectionDefinition.options[0].optionValue.settingValueTemplateReference = @{ useTemplateDefault = $true } }
        }
        $CollectionDefinitions = @($Definitions)
        if ($ChoiceCase -ne 'missing-definition') { $CollectionDefinitions += $CollectionDefinition }
        if ($ChoiceCase -eq 'duplicate-definition') { $CollectionDefinitions += $CollectionDefinition.Clone() }
        $CollectionInstance = @{ settingDefinitionId = $RootId; groupSettingCollectionValue = @(@{ children = @($First.children[0], $CollectionNode) }) }
        if ($DecodeJson) {
            $CollectionDefinitions = $CollectionDefinitions | ConvertTo-Json -Depth 25 | ConvertFrom-Json
            $CollectionInstance = $CollectionInstance | ConvertTo-Json -Depth 25 | ConvertFrom-Json
        }
        $CollectionRows = @(ConvertTo-IntuneEpmRules $CollectionInstance $CollectionDefinitions 'policy' 'choice-items')
        Assert-Epm (($null -ne $CollectionRows[0].fileName) -eq ($ChoiceCase -eq 'valid')) "Choice collection option was not resolved independently: $ChoiceCase"
        Assert-Epm (($null -ne $CollectionRows[0].filePath) -eq ($ChoiceCase -notin @('missing-definition', 'duplicate-definition'))) 'First option uncertainty leaked into a valid collection item'
        Assert-Epm ($CollectionRows[0].name -ceq 'Rule one') 'Choice collection hid an independent field'
    }
}
foreach ($DecodeJson in @($false, $true)) {
    $GroupItems = @{ settingDefinitionId = 'group-items'; groupSettingCollectionValue = @(
        @{ settingValueTemplateReference = @{ useTemplateDefault = $true }; children = @($First.children[1]) },
        @{ children = @($First.children[2]) }
    ) }
    $GroupInstance = @{ settingDefinitionId = $RootId; groupSettingCollectionValue = @(@{ children = @($First.children[0], $GroupItems) }) }
    if ($DecodeJson) { $GroupInstance = $GroupInstance | ConvertTo-Json -Depth 20 | ConvertFrom-Json }
    $GroupRows = @(ConvertTo-IntuneEpmRules $GroupInstance $Definitions 'policy' 'group-items')
    Assert-Epm ($null -eq $GroupRows[0].fileName -and $GroupRows[0].filePath -ceq '\\server\share' -and $GroupRows[0].name -ceq 'Rule one') 'Group-item template uncertainty leaked into an explicit sibling item'
}
$EmptyIdentityCases = foreach ($Placement in @('root', 'definition', 'child')) {
    foreach ($EmptyId in @('', ' ', "`t", "`r`n")) {
        @{ Placement = $Placement; Id = $EmptyId }
    }
}
foreach ($EmptyIdentityCase in $EmptyIdentityCases) {
    foreach ($DecodeJson in @($false, $true)) {
        $EmptyDefinitions = @($Definitions)
        $EmptyInstance = @{ settingDefinitionId = $RootId; groupSettingCollectionValue = @(@{ children = @($First.children) }) }
        switch ($EmptyIdentityCase.Placement) {
            'root' { $EmptyInstance.settingDefinitionId = $EmptyIdentityCase.Id }
            'definition' { $EmptyDefinitions += @{ id = $EmptyIdentityCase.Id; offsetUri = '/PrivilegeManagement/ElevationRules/{0}' } }
            'child' { $EmptyInstance.groupSettingCollectionValue[0].children += @{ settingDefinitionId = $EmptyIdentityCase.Id; simpleSettingValue = @{ value = 'SECRET_EMPTY_ID' } } }
        }
        if ($DecodeJson) {
            $EmptyDefinitions = $EmptyDefinitions | ConvertTo-Json -Depth 20 | ConvertFrom-Json
            $EmptyInstance = $EmptyInstance | ConvertTo-Json -Depth 20 | ConvertFrom-Json
        }
        $EmptyRejected = $false
        try { $null = @(ConvertTo-IntuneEpmRules $EmptyInstance $EmptyDefinitions 'policy' 'empty-identity') }
        catch { $EmptyRejected = $true }
        Assert-Epm $EmptyRejected "Empty EPM identity bypassed validation: $($EmptyIdentityCase.Placement); JSON=$DecodeJson"
    }
}
$MixedBindingCases = foreach ($Placement in @('child', 'root-definition', 'field-definition')) {
    foreach ($Shape in @('array', 'multi-array', 'empty-array', 'number', 'boolean', 'object', 'null', 'missing')) {
        foreach ($MalformedFirst in @($false, $true)) { @{ Placement = $Placement; Shape = $Shape; First = $MalformedFirst } }
    }
}
foreach ($MixedCase in $MixedBindingCases) {
    $MixedBinding = $MixedCase.Placement
    foreach ($DecodeJson in @($false, $true)) {
        $MixedDefinitions = @($Definitions)
        $MixedInstance = @{ settingDefinitionId = $RootId; groupSettingCollectionValue = @(@{ children = @($First.children) }) }
        if ($MixedBinding -eq 'child') {
            $MalformedCandidate = @{ settingDefinitionId = ($RootId + '_filename'); simpleSettingValue = @{ value = 'SECRET_DUPLICATE.exe' } }
            $IdProperty = 'settingDefinitionId'
        } else {
            $DefinitionIndex = if ($MixedBinding -eq 'root-definition') { 0 } else { 2 }
            $MalformedCandidate = $Definitions[$DefinitionIndex].Clone()
            $IdProperty = 'id'
        }
        $OriginalId = $MalformedCandidate[$IdProperty]
        switch ($MixedCase.Shape) {
            'array' { $MalformedCandidate[$IdProperty] = @($OriginalId) }
            'multi-array' { $MalformedCandidate[$IdProperty] = @($OriginalId, 'SECRET_DUPLICATE_ID') }
            'empty-array' { $MalformedCandidate[$IdProperty] = @() }
            'number' { $MalformedCandidate[$IdProperty] = 1 }
            'boolean' { $MalformedCandidate[$IdProperty] = $true }
            'object' { $MalformedCandidate[$IdProperty] = @{ value = $OriginalId } }
            'null' { $MalformedCandidate[$IdProperty] = $null }
            'missing' { $MalformedCandidate.Remove($IdProperty) }
        }
        if ($MixedBinding -eq 'child') {
            if ($MixedCase.First) { $MixedInstance.groupSettingCollectionValue[0].children = @($MalformedCandidate) + $First.children }
            else { $MixedInstance.groupSettingCollectionValue[0].children += $MalformedCandidate }
        } else {
            if ($MixedCase.First) { $MixedDefinitions = @($MalformedCandidate) + $MixedDefinitions }
            else { $MixedDefinitions += $MalformedCandidate }
        }
        if ($DecodeJson) {
            $MixedDefinitions = $MixedDefinitions | ConvertTo-Json -Depth 20 | ConvertFrom-Json
            $MixedInstance = $MixedInstance | ConvertTo-Json -Depth 20 | ConvertFrom-Json
        }
        $MixedRejected = $false
        try { $null = @(ConvertTo-IntuneEpmRules $MixedInstance $MixedDefinitions 'policy' 'mixed-binding') }
        catch { $MixedRejected = $true }
        Assert-Epm $MixedRejected "Malformed duplicate was discarded before EPM uniqueness checks: $MixedBinding; $($MixedCase.Shape); First=$($MixedCase.First); JSON=$DecodeJson"
    }
}
$BindingCases = foreach ($Binding in @('root-instance', 'root-definition', 'root-offset', 'field-instance', 'field-definition', 'field-offset')) {
    foreach ($Shape in @('valid', 'array', 'multi-array', 'empty-array', 'number', 'boolean', 'object', 'null', 'missing', 'blank', 'case', 'leading-space', 'trailing-space')) {
        @{ Binding = $Binding; Shape = $Shape }
    }
}
foreach ($BindingCase in $BindingCases) {
    $Binding = $BindingCase.Binding
    $Shape = $BindingCase.Shape
    foreach ($DecodeJson in @($false, $true)) {
        $BindingDefinitions = @($Definitions | ForEach-Object { $_.Clone() })
        $BindingLeaf = $First.children[1].Clone()
        $BindingInstance = @{ settingDefinitionId = $RootId; groupSettingCollectionValue = @(@{ children = @($First.children[0], $BindingLeaf) }) }
        $Property = 'settingDefinitionId'
        switch ($Binding) {
            'root-instance' { $Target = $BindingInstance }
            'root-definition' { $Target = $BindingDefinitions[0]; $Property = 'id' }
            'root-offset' { $Target = $BindingDefinitions[0]; $Property = 'offsetUri' }
            'field-instance' { $Target = $BindingLeaf }
            'field-definition' { $Target = $BindingDefinitions[2]; $Property = 'id' }
            'field-offset' { $Target = $BindingDefinitions[2]; $Property = 'offsetUri' }
        }
        $ExactValue = $Target[$Property]
        switch ($Shape) {
            'array' { $Target[$Property] = @($ExactValue) }
            'multi-array' { $Target[$Property] = @($ExactValue, 'SECRET_BINDING') }
            'empty-array' { $Target[$Property] = @() }
            'number' { $Target[$Property] = 1 }
            'boolean' { $Target[$Property] = $true }
            'object' { $Target[$Property] = @{ value = $ExactValue } }
            'null' { $Target[$Property] = $null }
            'missing' { $Target.Remove($Property) }
            'blank' { $Target[$Property] = '' }
            'case' { $Target[$Property] = $ExactValue.ToUpperInvariant() }
            'leading-space' { $Target[$Property] = ' ' + $ExactValue }
            'trailing-space' { $Target[$Property] = $ExactValue + ' ' }
        }
        if ($DecodeJson) {
            $BindingInstance = $BindingInstance | ConvertTo-Json -Depth 20 | ConvertFrom-Json
            $BindingDefinitions = $BindingDefinitions | ConvertTo-Json -Depth 20 | ConvertFrom-Json
        }
        $BindingRejected = $false
        $BindingRules = @()
        try { $BindingRules = @(ConvertTo-IntuneEpmRules $BindingInstance $BindingDefinitions 'policy' 'binding') }
        catch { $BindingRejected = $true }
        $Context = "$Binding; $Shape; JSON=$DecodeJson"
        if ($Shape -eq 'valid') {
            Assert-Epm (-not $BindingRejected -and $BindingRules.Count -eq 1 -and $BindingRules[0].fileName -ceq 'setup*.exe' -and $BindingRules[0].name -ceq 'Rule one') "Exact EPM binding changed: $Context"
        } elseif ($Binding -eq 'root-instance') {
            if ($Shape -in @('case', 'leading-space', 'trailing-space')) {
                Assert-Epm (-not $BindingRejected -and $BindingRules.Count -eq 0) "Unsupported string root was treated as an EPM rule identity: $Context"
            } else {
                Assert-Epm $BindingRejected "Invalid EPM root identity did not invalidate coverage: $Context"
            }
        } elseif ($Binding.StartsWith('root-')) {
            Assert-Epm $BindingRejected "Malformed EPM root definition was accepted: $Context"
        } elseif ($Binding -in @('field-instance', 'field-definition') -and $Shape -in @('array', 'multi-array', 'empty-array', 'number', 'boolean', 'object', 'null', 'missing', 'blank')) {
            Assert-Epm $BindingRejected "Invalid EPM field identity did not invalidate coverage: $Context"
        } else {
            Assert-Epm (-not $BindingRejected -and $BindingRules.Count -eq 1 -and $null -eq $BindingRules[0].fileName -and $BindingRules[0].name -ceq 'Rule one') "Malformed EPM field binding produced evidence or hid a sibling: $Context"
        }
        Assert-Epm (-not (($BindingRules | ConvertTo-Json -Depth 10) -match 'SECRET')) "Malformed binding leaked: $Context"
    }
}
foreach ($DuplicateIndex in @(0, 2)) {
    foreach ($DecodeJson in @($false, $true)) {
        $DuplicateDefinitions = $Definitions + @($Definitions[$DuplicateIndex].Clone())
        $DuplicateInstance = $Instance
        if ($DecodeJson) {
            $DuplicateDefinitions = $DuplicateDefinitions | ConvertTo-Json -Depth 20 | ConvertFrom-Json
            $DuplicateInstance = $DuplicateInstance | ConvertTo-Json -Depth 20 | ConvertFrom-Json
        }
        $DuplicateRejected = $false
        $DuplicateRules = @()
        try { $DuplicateRules = @(ConvertTo-IntuneEpmRules $DuplicateInstance $DuplicateDefinitions 'policy' 'duplicate-binding') }
        catch { $DuplicateRejected = $true }
        if ($DuplicateIndex -eq 0) {
            Assert-Epm $DuplicateRejected 'Duplicate typed root definitions resolved'
        } else {
            Assert-Epm (-not $DuplicateRejected -and $DuplicateRules.Count -eq 2 -and $null -eq $DuplicateRules[0].fileName -and $DuplicateRules[0].name -ceq 'Rule one') 'Duplicate typed field definitions resolved or hid a sibling'
        }
    }
}
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
$script:EpmFixtureDefinitions = @($Definitions | ForEach-Object { $_.Clone() })
$script:EpmFixtureDefinitions[3].offsetUri = @('/PrivilegeManagement/ElevationRules/{0}/FilePath')
$script:EpmFixtureInstance = @{ settingDefinitionId = $RootId; groupSettingCollectionValue = @(
    @{ children = @(
        $First.children[0],
        $First.children[3],
        @{ settingDefinitionId = @(($RootId + '_filename')); simpleSettingValue = @{ value = 'SECRET_BINDING_FILE.exe' } },
        @{ settingDefinitionId = ($RootId + '_filepath'); simpleSettingValue = @{ value = 'C:\SECRET_BINDING_PATH' } }
    ) },
    @{ children = @(
        $First.children[0],
        @{ settingDefinitionId = ($RootId + '_filename'); simpleSettingValue = @{ value = 'other.exe' } },
        @{ settingDefinitionId = @(($RootId + '_ruletype')); simpleSettingValue = @{ value = 'Automatic' } }
    ) }
) }
$BindingDocument = Invoke-IntuneDiscoveryCore -SelectedTenant '22222222-2222-4222-8222-222222222222' -Configuration $true -Request $EpmRequest -Requirements $EpmRequirements
$BindingRows = $BindingDocument.Inventory.EpmRules
Assert-Epm ($BindingRows.Count -eq 0 -and $BindingDocument.CollectionStatus.EpmRules.State -ceq 'Partial' -and $BindingDocument.CollectionStatus.EpmRules.CompletedParentIds.Count -eq 0) 'Malformed field identity did not invalidate EPM coverage'
$BindingJson = $BindingDocument | ConvertTo-Json -Depth 30
Assert-Epm (-not ($BindingJson -match 'SECRET|offsetUri|ChoiceUnresolved|TemplateUnresolved')) 'Malformed EPM binding payload exported'
if ($env:ASSAY_INTUNE_EPM_BINDING_FIXTURE) { [IO.File]::WriteAllText($env:ASSAY_INTUNE_EPM_BINDING_FIXTURE, $BindingJson, [Text.UTF8Encoding]::new($false)) }
foreach ($RootCase in @('definition-id', 'duplicate-definition', 'offset')) {
    $script:EpmFixtureDefinitions = @($Definitions | ForEach-Object { $_.Clone() })
    $script:EpmFixtureInstance = $Instance
    switch ($RootCase) {
        'definition-id' { $script:EpmFixtureDefinitions[0].id = @($RootId) }
        'duplicate-definition' { $script:EpmFixtureDefinitions += $Definitions[0].Clone() }
        'offset' { $script:EpmFixtureDefinitions[0].offsetUri = @('/PrivilegeManagement/ElevationRules/{0}') }
    }
    $RejectedRootDocument = Invoke-IntuneDiscoveryCore -SelectedTenant '22222222-2222-4222-8222-222222222222' -Configuration $true -Request $EpmRequest -Requirements $EpmRequirements
    $ExpectedRows = if ($RootCase -eq 'duplicate-definition') { 2 } else { 0 }
    Assert-Epm ($RejectedRootDocument.Inventory.EpmRules.Count -eq $ExpectedRows -and $RejectedRootDocument.CollectionStatus.EpmRules.State -ceq 'Partial' -and $RejectedRootDocument.CollectionStatus.EpmRules.CompletedParentIds.Count -eq 0) "Malformed or duplicate EPM root retained complete coverage: $RootCase"
}
if ($env:ASSAY_INTUNE_EPM_BINDING_ROOT_FIXTURE) { [IO.File]::WriteAllText($env:ASSAY_INTUNE_EPM_BINDING_ROOT_FIXTURE, ($RejectedRootDocument | ConvertTo-Json -Depth 30), [Text.UTF8Encoding]::new($false)) }
$script:EpmFixtureDefinitions = $Definitions
$script:EpmFixtureInstance = $Instance
$script:MixedRootDefinitionRead = $false
$MixedRootRequest = {
    param($Address)
    if (([uri]$Address).AbsolutePath -eq '/beta/deviceManagement/configurationPolicies') {
        return @{ StatusCode = 200; Body = @{ value = @(@{ id = 'policy'; platforms = 'windows10'; technologies = 'endpointPrivilegeManagement'; settingCount = 2; templateReference = @{ templateFamily = 'endpointSecurityEndpointPrivilegeManagement' } }) } }
    }
    if (([uri]$Address).AbsolutePath -eq '/beta/deviceManagement/configurationPolicies/policy/settings') {
        $MalformedInstance = $script:EpmFixtureInstance.Clone()
        if ($script:MixedRootMalformed) { $MalformedInstance.settingDefinitionId = @($MalformedInstance.settingDefinitionId) }
        return @{ StatusCode = 200; Body = (@{ value = @(
            @{ id = '0'; settingInstance = $script:EpmFixtureInstance },
            @{ id = '1'; settingInstance = $MalformedInstance }
        ) } | ConvertTo-Json -Depth 30 | ConvertFrom-Json) }
    }
    if (([uri]$Address).AbsolutePath -eq '/beta/deviceManagement/configurationPolicies/policy/settings/1/settingDefinitions') {
        $script:MixedRootDefinitionRead = $true
        return @{ StatusCode = 200; Body = (@{ value = $script:EpmFixtureDefinitions } | ConvertTo-Json -Depth 30 | ConvertFrom-Json) }
    }
    & $EpmRequest $Address
}
foreach ($MalformedRoot in @($false, $true)) {
    $script:MixedRootMalformed = $MalformedRoot
    $script:MixedRootDefinitionRead = $false
    $MixedRootDocument = Invoke-IntuneDiscoveryCore -SelectedTenant '22222222-2222-4222-8222-222222222222' -Configuration $true -Request $MixedRootRequest -Requirements $EpmRequirements
    Assert-Epm $script:MixedRootDefinitionRead 'Mixed-root coverage probe did not reach the second definition response'
    if ($MalformedRoot) {
        Assert-Epm ($MixedRootDocument.Inventory.EpmRules.Count -eq 2 -and $MixedRootDocument.CollectionStatus.EpmRules.State -ceq 'Partial' -and $MixedRootDocument.CollectionStatus.EpmRules.CompletedParentIds.Count -eq 0) 'Valid EPM sibling hid malformed root identity from collection coverage'
    } else {
        Assert-Epm ($MixedRootDocument.Inventory.EpmRules.Count -eq 4 -and $MixedRootDocument.CollectionStatus.EpmRules.State -ceq 'Complete' -and $MixedRootDocument.CollectionStatus.EpmRules.CompletedParentIds -contains 'policy') 'Two explicit EPM settings lost complete coverage'
    }
}
$script:EpmDuplicateValid = @{ settingDefinitionId = $RootId; groupSettingCollectionValue = @(@{ children = @(
    $First.children[0],
    @{ settingDefinitionId = ($RootId + '_filename'); simpleSettingValue = @{ value = 'safe.exe' } },
    @{ settingDefinitionId = ($RootId + '_filepath'); simpleSettingValue = @{ value = 'C:\Program Files\Review' } },
    $First.children[3]
) }) }
$DuplicateRequest = {
    param($Address)
    $RequestPath = ([uri]$Address).AbsolutePath
    $Rows = switch ($RequestPath) {
        '/beta/deviceManagement/configurationPolicies' { @(@{ id = 'policy'; platforms = 'windows10'; technologies = 'endpointPrivilegeManagement'; settingCount = 2; templateReference = @{ templateFamily = 'endpointSecurityEndpointPrivilegeManagement' } }) }
        '/beta/deviceManagement/configurationPolicies/policy/settings' {
            @(@{ id = '0'; settingInstance = $script:EpmDuplicateValid }, @{ id = '1'; settingInstance = $script:EpmDuplicateInstance })
        }
        '/beta/deviceManagement/configurationPolicies/policy/settings/0/settingDefinitions' { $Definitions }
        '/beta/deviceManagement/configurationPolicies/policy/settings/1/settingDefinitions' {
            $script:EpmDuplicateDefinitionsRead = $true
            $script:EpmDuplicateDefinitions
        }
        default { @() }
    }
    @{ StatusCode = 200; Body = (@{ value = @($Rows) } | ConvertTo-Json -Depth 30 | ConvertFrom-Json) }
}
foreach ($DuplicateCase in @('valid', 'child', 'root-definition', 'field-definition', 'empty-root', 'whitespace-root', 'empty-child', 'whitespace-child', 'mixed-root-choice', 'mixed-root-simple', 'mixed-root-choice-collection', 'mixed-root-simple-collection')) {
    foreach ($MalformedFirst in @($false, $true)) {
        $script:EpmDuplicateInstance = @{ settingDefinitionId = $RootId; groupSettingCollectionValue = @(@{ children = @($script:EpmDuplicateValid.groupSettingCollectionValue[0].children) }) }
        $script:EpmDuplicateDefinitions = @($Definitions)
        if ($DuplicateCase.StartsWith('mixed-root-')) {
            $ExtraKind = switch ($DuplicateCase) {
                'mixed-root-choice' { 'choiceSettingValue' }
                'mixed-root-simple' { 'simpleSettingValue' }
                'mixed-root-choice-collection' { 'choiceSettingCollectionValue' }
                'mixed-root-simple-collection' { 'simpleSettingCollectionValue' }
            }
            $script:EpmDuplicateInstance[$ExtraKind] = @{ value = 'SECRET_ROOT_VALUE' }
            if ($ExtraKind.EndsWith('CollectionValue')) { $script:EpmDuplicateInstance[$ExtraKind] = @($script:EpmDuplicateInstance[$ExtraKind]) }
        } elseif ($DuplicateCase -in @('empty-root', 'whitespace-root')) {
            $script:EpmDuplicateInstance.settingDefinitionId = if ($DuplicateCase -eq 'empty-root') { '' } else { " `t" }
        } elseif ($DuplicateCase -in @('child', 'empty-child', 'whitespace-child')) {
            $MalformedChild = @{ settingDefinitionId = @(($RootId + '_filename')); simpleSettingValue = @{ value = 'SECRET_DUPLICATE*.exe' } }
            if ($DuplicateCase -eq 'empty-child') { $MalformedChild.settingDefinitionId = '' }
            if ($DuplicateCase -eq 'whitespace-child') { $MalformedChild.settingDefinitionId = " `t" }
            if ($MalformedFirst) { $script:EpmDuplicateInstance.groupSettingCollectionValue[0].children = @($MalformedChild) + $script:EpmDuplicateInstance.groupSettingCollectionValue[0].children }
            else { $script:EpmDuplicateInstance.groupSettingCollectionValue[0].children += $MalformedChild }
        } elseif ($DuplicateCase -ne 'valid') {
            $DefinitionIndex = if ($DuplicateCase -eq 'root-definition') { 0 } else { 2 }
            $MalformedDefinition = $Definitions[$DefinitionIndex].Clone()
            $MalformedDefinition.id = @($MalformedDefinition.id, 'SECRET_DUPLICATE_ID')
            if ($MalformedFirst) { $script:EpmDuplicateDefinitions = @($MalformedDefinition) + $script:EpmDuplicateDefinitions }
            else { $script:EpmDuplicateDefinitions += $MalformedDefinition }
        }
        $script:EpmDuplicateDefinitionsRead = $false
        $DuplicateDocument = Invoke-IntuneDiscoveryCore -SelectedTenant '22222222-2222-4222-8222-222222222222' -Configuration $true -Request $DuplicateRequest -Requirements $EpmRequirements
        Assert-Epm $script:EpmDuplicateDefinitionsRead 'Mixed duplicate test did not reach the second definitions response'
        if ($DuplicateCase -eq 'valid') {
            Assert-Epm ($DuplicateDocument.Inventory.EpmRules.Count -eq 2 -and $DuplicateDocument.CollectionStatus.EpmRules.State -ceq 'Complete' -and $DuplicateDocument.CollectionStatus.EpmRules.CompletedParentIds -contains 'policy') 'Wildcard-free duplicate control lost complete coverage'
        } else {
            Assert-Epm ($DuplicateDocument.Inventory.EpmRules.Count -eq 1 -and $DuplicateDocument.CollectionStatus.EpmRules.State -ceq 'Partial' -and $DuplicateDocument.CollectionStatus.EpmRules.CompletedParentIds.Count -eq 0) "Malformed duplicate retained complete coverage: $DuplicateCase; First=$MalformedFirst"
        }
        Assert-Epm ($DuplicateDocument.Inventory.EpmRules[0].fileName -ceq 'safe.exe' -and $DuplicateDocument.Inventory.EpmRules[0].elevationType -ceq 'Automatic') 'Independent valid EPM setting lost evidence'
        $DuplicateJson = $DuplicateDocument | ConvertTo-Json -Depth 30
        Assert-Epm (-not ($DuplicateJson -match 'SECRET')) 'Malformed duplicate input leaked into export'
        if ($DuplicateCase -eq 'valid' -and $env:ASSAY_INTUNE_EPM_DUPLICATE_CONTROL_FIXTURE) { [IO.File]::WriteAllText($env:ASSAY_INTUNE_EPM_DUPLICATE_CONTROL_FIXTURE, $DuplicateJson, [Text.UTF8Encoding]::new($false)) }
        if ($DuplicateCase -eq 'child' -and $env:ASSAY_INTUNE_EPM_DUPLICATE_FIXTURE) { [IO.File]::WriteAllText($env:ASSAY_INTUNE_EPM_DUPLICATE_FIXTURE, $DuplicateJson, [Text.UTF8Encoding]::new($false)) }
        if ($DuplicateCase -eq 'whitespace-root' -and $env:ASSAY_INTUNE_EPM_EMPTY_ID_FIXTURE) { [IO.File]::WriteAllText($env:ASSAY_INTUNE_EPM_EMPTY_ID_FIXTURE, $DuplicateJson, [Text.UTF8Encoding]::new($false)) }
        if ($DuplicateCase -eq 'mixed-root-choice' -and $env:ASSAY_INTUNE_EPM_ROOT_MIX_FIXTURE) { [IO.File]::WriteAllText($env:ASSAY_INTUNE_EPM_ROOT_MIX_FIXTURE, $DuplicateJson, [Text.UTF8Encoding]::new($false)) }
    }
}
$script:EpmFixtureDefinitions = $Definitions + @(@{ id = 'nested-container'; options = @(@{ itemId = 'known'; optionValue = @{ value = 1 } }) })
foreach ($NestedDuplicate in @($false, $true)) {
    $NestedGroups = @(foreach ($Kind in @('groupSettingCollectionValue', 'choiceSettingCollectionValue')) {
        $NestedFields = @(
            @{ settingDefinitionId = ($RootId + '_filename'); simpleSettingValue = @{ value = $(if ($NestedDuplicate) { 'SECRET_NESTED*.exe' } else { 'safe.exe' }) } },
            @{ settingDefinitionId = ($RootId + '_filepath'); simpleSettingValue = @{ value = $(if ($NestedDuplicate) { 'C:\SECRET_NESTED' } else { 'C:\Program Files\Review' }) } }
        )
        $NestedValue = @{ children = $NestedFields }
        if ($Kind -eq 'choiceSettingCollectionValue') { $NestedValue.value = 'known' }
        $NestedChildren = @($First.children[0], $First.children[3], @{ settingDefinitionId = 'nested-container'; $Kind = @($NestedValue) })
        if ($NestedDuplicate) { $NestedChildren += $script:EpmDuplicateValid.groupSettingCollectionValue[0].children[1..2] }
        @{ children = $NestedChildren }
    })
    $script:EpmFixtureInstance = @{ settingDefinitionId = $RootId; groupSettingCollectionValue = $NestedGroups }
    $NestedDocument = Invoke-IntuneDiscoveryCore -SelectedTenant '22222222-2222-4222-8222-222222222222' -Configuration $true -Request $EpmRequest -Requirements $EpmRequirements
    Assert-Epm ($NestedDocument.Inventory.EpmRules.Count -eq 2 -and $NestedDocument.CollectionStatus.EpmRules.State -ceq 'Complete' -and $NestedDocument.CollectionStatus.EpmRules.CompletedParentIds -contains 'policy') 'Nested EPM evidence lost rule-group or parent coverage'
    foreach ($Rule in $NestedDocument.Inventory.EpmRules) {
        Assert-Epm ($Rule.name -ceq 'Rule one' -and $Rule.elevationType -ceq 'Automatic') 'Nested duplicate hid independently resolved fields'
        if ($NestedDuplicate) { Assert-Epm ($null -eq $Rule.fileName -and $null -eq $Rule.filePath) 'Nested duplicate manufactured a unique rule field' }
        else { Assert-Epm ($Rule.fileName -ceq 'safe.exe' -and $Rule.filePath -ceq 'C:\Program Files\Review') 'Resolved nested-only fields were lost' }
    }
    $NestedJson = $NestedDocument | ConvertTo-Json -Depth 30
    Assert-Epm (-not ($NestedJson -match 'SECRET|TemplateUnresolved|ChoiceUnresolved')) 'Nested EPM raw values or traversal context exported'
    $FixturePath = if ($NestedDuplicate) { $env:ASSAY_INTUNE_EPM_COLLECTION_FIXTURE } else { $env:ASSAY_INTUNE_EPM_COLLECTION_CONTROL_FIXTURE }
    if ($FixturePath) { [IO.File]::WriteAllText($FixturePath, $NestedJson, [Text.UTF8Encoding]::new($false)) }
}