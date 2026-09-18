function Get-IntuneEpmRuleChildren {
    param($Children, [int]$Depth = 0, $Definitions = @(), [bool]$InheritedTemplateUnresolved = $false)
    if ($Depth -gt 12) { throw 'EPM rule depth limit' }
    foreach ($Child in $Children) {
        $Choice = Get-IntuneValue $Child 'choiceSettingValue'
        $Simple = Get-IntuneValue $Child 'simpleSettingValue'
        $TemplateUnresolved = $InheritedTemplateUnresolved -or (Test-IntuneTemplateUnresolved $Choice) -or (Test-IntuneTemplateUnresolved $Simple)
        if ($null -ne $Choice) {
            $ChildId = Get-IntuneValue $Child 'settingDefinitionId'
            $ChildDefinitions = @(foreach ($Candidate in $Definitions) {
                $CandidateId = Get-IntuneValue $Candidate 'id'
                if ($ChildId -is [string] -and $CandidateId -is [string] -and [string]::Equals($ChildId, $CandidateId, [StringComparison]::Ordinal)) { $Candidate }
            })
            $Selected = Get-IntuneSelectedOptionValue $ChildDefinitions (Get-IntuneValue $Choice 'value')
            $TemplateUnresolved = $TemplateUnresolved -or (Test-IntuneTemplateUnresolved $Selected)
        }
        @{ Instance = $Child; TemplateUnresolved = $TemplateUnresolved }
        Get-IntuneEpmRuleChildren (Get-IntuneValue $Choice 'children' @()) ($Depth + 1) $Definitions $TemplateUnresolved
    }
}

function ConvertTo-IntuneEpmRules {
    param($Instance, $Definitions, [string]$PolicyId, [string]$SettingId)
    $RootId = 'device_vendor_msft_policy_privilegemanagement_elevationrules_{elevationrulename}'
    if ((Get-IntuneValue $Instance 'settingDefinitionId') -cne $RootId) { return }
    $Root = @($Definitions | Where-Object { (Get-IntuneValue $_ 'id') -ceq $RootId })
    if ($Root.Count -ne 1 -or (Get-IntuneValue $Root[0] 'offsetUri') -cne '/PrivilegeManagement/ElevationRules/{0}') { throw 'Unsupported EPM rule group definition' }
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
            $Definition = @($Definitions | Where-Object { (Get-IntuneValue $_ 'id') -ceq $DefinitionId })
            $Child = @($Children | Where-Object { (Get-IntuneValue $_.Instance 'settingDefinitionId') -ceq $DefinitionId })
            if ($Definition.Count -ne 1 -or $Child.Count -ne 1 -or (Get-IntuneValue $Definition[0] 'offsetUri') -cne ('/PrivilegeManagement/ElevationRules/{0}/' + $Fields[$Suffix])) { continue }
            if ($Child[0].TemplateUnresolved) { continue }
            $Choice = Get-IntuneValue $Child[0].Instance 'choiceSettingValue'
            $Selected = Get-IntuneValue $Child[0].Instance 'simpleSettingValue'
            if ($null -ne $Choice) {
                $Options = @(foreach ($Option in (Get-IntuneValue $Definition[0] 'options' @())) { if ((Get-IntuneValue $Option 'itemId') -ceq (Get-IntuneValue $Choice 'value')) { $Option } })
                if ($Options.Count -ne 1) { continue }
                $Selected = Get-IntuneValue $Options[0] 'optionValue'
            }
            if ((Test-IntuneTemplateUnresolved $Selected) -or (Test-IntuneTemplateUnresolved $Choice)) { continue }
            $Value = Get-IntuneValue $Selected 'value'
            if ($Value -is [string] -and $Value.Length -le 512 -and $Value -notmatch '[\x00-\x1F]') { $Row[$OutputNames[$Suffix]] = $Value }
        }
        $Row
    }
}