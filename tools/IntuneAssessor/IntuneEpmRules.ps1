function Get-IntuneEpmRuleChildren {
    param($Children, [int]$Depth = 0, $Definitions = @(), [bool]$InheritedTemplateUnresolved = $false, [bool]$InheritedChoiceUnresolved = $false)
    if ($Depth -gt 12) { throw 'EPM rule depth limit' }
    if ($null -ne $Children -and $Children -isnot [array]) { throw 'Invalid EPM child collection' }
    foreach ($Child in $Children) {
        $ChildId = Get-IntuneValue $Child 'settingDefinitionId'
        if ($ChildId -isnot [string] -or [string]::IsNullOrWhiteSpace($ChildId)) { throw 'Invalid EPM child setting definition identity' }
        $Choice = Get-IntuneValue $Child 'choiceSettingValue'
        $Simple = Get-IntuneValue $Child 'simpleSettingValue'
        $ValueKindCount = 0
        foreach ($ValueName in @('choiceSettingValue', 'simpleSettingValue', 'groupSettingCollectionValue', 'choiceSettingCollectionValue', 'simpleSettingCollectionValue')) {
            if ($null -ne (Get-IntuneValue $Child $ValueName)) { $ValueKindCount++ }
        }
        $TemplateUnresolved = $InheritedTemplateUnresolved -or (Test-IntuneTemplateUnresolved $Choice) -or (Test-IntuneTemplateUnresolved $Simple)
        $ChoiceUnresolved = $InheritedChoiceUnresolved -or $ValueKindCount -gt 1
        $ChildDefinitions = @()
        if ($null -ne $Choice -or $null -ne (Get-IntuneValue $Child 'choiceSettingCollectionValue')) {
            $ChildDefinitions = @(foreach ($Candidate in $Definitions) {
                $CandidateId = Get-IntuneValue $Candidate 'id'
                if ($ChildId -is [string] -and $CandidateId -is [string] -and [string]::Equals($ChildId, $CandidateId, [StringComparison]::Ordinal)) { $Candidate }
            })
        }
        if ($null -ne $Choice) {
            $Selected = Get-IntuneSelectedOptionValue $ChildDefinitions (Get-IntuneValue $Choice 'value')
            $ChoiceUnresolved = $ChoiceUnresolved -or $null -eq $Selected
            $TemplateUnresolved = $TemplateUnresolved -or (Test-IntuneTemplateUnresolved $Selected)
        }
        @{ Instance = $Child; TemplateUnresolved = $TemplateUnresolved; ChoiceUnresolved = $ChoiceUnresolved }
        Get-IntuneEpmRuleChildren (Get-IntuneValue $Choice 'children' @()) ($Depth + 1) $Definitions $TemplateUnresolved $ChoiceUnresolved
        foreach ($ValueName in @('groupSettingCollectionValue', 'choiceSettingCollectionValue')) {
            $CollectionValues = Get-IntuneValue $Child $ValueName
            if ($null -eq $CollectionValues) { continue }
            if ($CollectionValues -isnot [array]) { throw 'Invalid EPM nested setting collection' }
            foreach ($CollectionValue in $CollectionValues) {
                if ($CollectionValue -isnot [Collections.IDictionary] -and $CollectionValue -isnot [pscustomobject]) { throw 'Invalid EPM nested setting value' }
                $ChildTemplateUnresolved = $TemplateUnresolved -or (Test-IntuneTemplateUnresolved $CollectionValue)
                $ChildChoiceUnresolved = $ChoiceUnresolved
                if ($ValueName -eq 'choiceSettingCollectionValue') {
                    $CollectionOption = Get-IntuneSelectedOptionValue $ChildDefinitions (Get-IntuneValue $CollectionValue 'value')
                    $ChildChoiceUnresolved = $ChildChoiceUnresolved -or $null -eq $CollectionOption
                    $ChildTemplateUnresolved = $ChildTemplateUnresolved -or (Test-IntuneTemplateUnresolved $CollectionOption)
                }
                Get-IntuneEpmRuleChildren (Get-IntuneValue $CollectionValue 'children' @()) ($Depth + 1) $Definitions $ChildTemplateUnresolved $ChildChoiceUnresolved
            }
        }
    }
}

function ConvertTo-IntuneEpmRules {
    param($Instance, $Definitions, [string]$PolicyId, [string]$SettingId)
    $RootId = 'device_vendor_msft_policy_privilegemanagement_elevationrules_{elevationrulename}'
    $InstanceId = Get-IntuneValue $Instance 'settingDefinitionId'
    if ($InstanceId -isnot [string] -or [string]::IsNullOrWhiteSpace($InstanceId)) { throw 'Invalid EPM setting definition identity' }
    if (-not [string]::Equals($InstanceId, $RootId, [StringComparison]::Ordinal)) { return }
    foreach ($Candidate in $Definitions) {
        $CandidateId = Get-IntuneValue $Candidate 'id'
        if ($CandidateId -isnot [string] -or [string]::IsNullOrWhiteSpace($CandidateId)) { throw 'Invalid EPM definition identity' }
    }
    $Root = @($Definitions | Where-Object {
        $CandidateId = Get-IntuneValue $_ 'id'
        $CandidateId -is [string] -and [string]::Equals($CandidateId, $RootId, [StringComparison]::Ordinal)
    })
    if ($Root.Count -ne 1) { throw 'Unsupported EPM rule group definition' }
    $RootOffset = Get-IntuneValue $Root[0] 'offsetUri'
    if ($RootOffset -isnot [string] -or -not [string]::Equals($RootOffset, '/PrivilegeManagement/ElevationRules/{0}', [StringComparison]::Ordinal)) { throw 'Unsupported EPM rule group definition' }
    $Groups = Get-IntuneValue $Instance 'groupSettingCollectionValue'
    if ($Groups -isnot [array] -or $Groups.Count -gt 100) { throw 'Invalid EPM rule groups' }
    $Fields = [ordered]@{ name = 'Name'; filename = 'FileName'; filepath = 'FilePath'; ruletype = 'RuleType' }
    $OutputNames = @{ name = 'name'; filename = 'fileName'; filepath = 'filePath'; ruletype = 'elevationType' }
    $Ordinal = 0
    foreach ($Group in $Groups) {
        $Ordinal++
        $Row = [ordered]@{ id = ($SettingId + ':rule:' + $Ordinal); parentId = $PolicyId; definitionId = $RootId; bindingVersion = '1'; name = $null; fileName = $null; filePath = $null; elevationType = $null }
        if (Test-IntuneTemplateUnresolved $Group) { $Row; continue }
        $Children = @(Get-IntuneEpmRuleChildren (Get-IntuneValue $Group 'children' @()) -Definitions $Definitions)
        foreach ($Suffix in $Fields.Keys) {
            $DefinitionId = $RootId + '_' + $Suffix
            $Definition = @($Definitions | Where-Object {
                $CandidateId = Get-IntuneValue $_ 'id'
                $CandidateId -is [string] -and [string]::Equals($CandidateId, $DefinitionId, [StringComparison]::Ordinal)
            })
            $Child = @($Children | Where-Object {
                $CandidateId = Get-IntuneValue $_.Instance 'settingDefinitionId'
                $CandidateId -is [string] -and [string]::Equals($CandidateId, $DefinitionId, [StringComparison]::Ordinal)
            })
            if ($Definition.Count -ne 1 -or $Child.Count -ne 1) { continue }
            $Offset = Get-IntuneValue $Definition[0] 'offsetUri'
            if ($Offset -isnot [string] -or -not [string]::Equals($Offset, ('/PrivilegeManagement/ElevationRules/{0}/' + $Fields[$Suffix]), [StringComparison]::Ordinal)) { continue }
            if ($Child[0].TemplateUnresolved -or $Child[0].ChoiceUnresolved) { continue }
            $Choice = Get-IntuneValue $Child[0].Instance 'choiceSettingValue'
            $Selected = Get-IntuneValue $Child[0].Instance 'simpleSettingValue'
            if ($null -ne $Choice) {
                $Selected = Get-IntuneSelectedOptionValue $Definition (Get-IntuneValue $Choice 'value')
                if ($null -eq $Selected) { continue }
            }
            if ((Test-IntuneTemplateUnresolved $Selected) -or (Test-IntuneTemplateUnresolved $Choice)) { continue }
            $Value = Get-IntuneValue $Selected 'value'
            if ($Value -is [string] -and $Value.Length -le 512 -and $Value -notmatch '[\x00-\x1F]') { $Row[$OutputNames[$Suffix]] = $Value }
        }
        $Row
    }
}