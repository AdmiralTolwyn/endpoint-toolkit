function Get-IntuneEpmRuleChildren {
    param($Children, [int]$Depth = 0)
    if ($Depth -gt 12) { throw 'EPM rule depth limit' }
    foreach ($Child in $Children) {
        $Child
        $Choice = Get-IntuneValue $Child 'choiceSettingValue'
        Get-IntuneEpmRuleChildren (Get-IntuneValue $Choice 'children' @()) ($Depth + 1)
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
        if (Get-IntuneValue (Get-IntuneValue $Group 'settingValueTemplateReference') 'useTemplateDefault' $false) { $Row; continue }
        $Children = @(Get-IntuneEpmRuleChildren (Get-IntuneValue $Group 'children' @()))
        foreach ($Suffix in $Fields.Keys) {
            $DefinitionId = $RootId + '_' + $Suffix
            $Definition = @($Definitions | Where-Object { (Get-IntuneValue $_ 'id') -ceq $DefinitionId })
            $Child = @($Children | Where-Object { (Get-IntuneValue $_ 'settingDefinitionId') -ceq $DefinitionId })
            if ($Definition.Count -ne 1 -or $Child.Count -ne 1 -or (Get-IntuneValue $Definition[0] 'offsetUri') -cne ('/PrivilegeManagement/ElevationRules/{0}/' + $Fields[$Suffix])) { continue }
            $Choice = Get-IntuneValue $Child[0] 'choiceSettingValue'
            $Selected = Get-IntuneValue $Child[0] 'simpleSettingValue'
            if ($null -ne $Choice) {
                $Options = @(foreach ($Option in (Get-IntuneValue $Definition[0] 'options' @())) { if ((Get-IntuneValue $Option 'itemId') -ceq (Get-IntuneValue $Choice 'value')) { $Option } })
                if ($Options.Count -ne 1) { continue }
                $Selected = Get-IntuneValue $Options[0] 'optionValue'
            }
            if ((Get-IntuneValue (Get-IntuneValue $Selected 'settingValueTemplateReference') 'useTemplateDefault' $false) -or (Get-IntuneValue (Get-IntuneValue $Choice 'settingValueTemplateReference') 'useTemplateDefault' $false)) { continue }
            $Value = Get-IntuneValue $Selected 'value'
            if ($Value -is [string] -and $Value.Length -le 512 -and $Value -notmatch '[\x00-\x1F]') { $Row[$OutputNames[$Suffix]] = $Value }
        }
        $Row
    }
}