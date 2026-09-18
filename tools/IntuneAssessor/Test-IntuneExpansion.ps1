#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Invoke-IntuneDiscovery.ps1') -LibraryOnly
function Assert-Expansion([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
Assert-Expansion (Test-IntuneSecurityPath 'Policy/Config/LocalSecurityAuthority/ConfigureLsaProtectedProcess') 'Published LSA protection URI blocked'
Assert-Expansion (-not (Test-IntuneSecurityPath 'Policy/Config/LSA/ConfigureLsaProtectedProcess')) 'Invented LSA URI accepted'
foreach ($Path in @('Policy/Config/MSSecurityGuide/ConfigureSMBV1ClientDriver', 'Policy/Config/MSSecurityGuide/ConfigureSMBV1Server', 'Policy/Config/WindowsPowerShell/TurnOnPowerShellScriptBlockLogging')) {
    $Unreviewed = @(ConvertTo-IntuneSettingFacts @{ settingDefinitionId = 'admx'; simpleSettingValue = @{ value = 1 } } @(@{ id = 'admx'; baseUri = './Device/Vendor/MSFT'; offsetUri = $Path }) 'policy' 'setting')
    Assert-Expansion ($Unreviewed[0].resolution -eq 'UnresolvedAdmx' -and -not $Unreviewed[0].Contains('value')) 'Unreviewed ADMX number interpreted as a CSP payload'
}
$Assignment = @{ id = 'assignment'; target = @{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = 'group'; entraObjectId = 'UNVERIFIED'; targetType = 'UNVERIFIED'; deviceAndAppManagementAssignmentFilterId = 'beta-filter'; deviceAndAppManagementAssignmentFilterType = 'include' } }
foreach ($Module in @('ComplianceAssignments', 'ConfigurationAssignments', 'ApplicationAssignments', 'MamAssignments', 'AppConfigurationAssignments', 'DeviceAppConfigurationAssignments')) {
    $Projected = ConvertTo-IntuneSafeRow -Module $Module -Row $Assignment
    Assert-Expansion ($Projected.target.groupId -eq 'group' -and -not (($Projected | ConvertTo-Json) -match 'UNVERIFIED|beta-filter')) 'V1 assignment retained unsupported target fields'
}
$Projected = ConvertTo-IntuneSafeRow -Module EnrollmentAssignments -Row $Assignment
Assert-Expansion ($Projected.target.deviceAndAppManagementAssignmentFilterId -eq 'beta-filter' -and -not (($Projected | ConvertTo-Json) -match 'UNVERIFIED')) 'Beta target field projection failed'
$App = ConvertTo-IntuneSafeRow -Module Applications -Row @{ id = 'app'; detectionRules = @(@{ scriptContent = 'SECRET' }); rules = @(@{ ruleType = 'detection'; operationType = 'notConfigured'; scriptContent = 'SECRET' }); roleScopeTagIds = @('unsupported') }
Assert-Expansion ($App.rules[0].ruleType -eq 'detection' -and -not (($App | ConvertTo-Json -Depth 10) -match 'SECRET|unsupported|detectionRules')) 'Documented Win32 rules not safely projected'
$Autopilot = ConvertTo-IntuneSafeRow -Module AutopilotProfiles -Row @{ id = 'profile'; outOfBoxExperienceSettings = @{ userType = 'administrator' }; outOfBoxExperienceSetting = @{ userType = 'standard'; privacySettingsHidden = $true } }
Assert-Expansion ($Autopilot.outOfBoxExperienceSetting.userType -eq 'standard' -and $Autopilot.outOfBoxExperienceSetting.privacySettingsHidden -and -not $Autopilot.Contains('outOfBoxExperienceSettings')) 'Deprecated Autopilot OOBE field used'
Assert-Expansion (-not (Test-IntuneUri 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies')) 'Beta became default'
Assert-Expansion (Test-IntuneUri 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies' -AllowExpansion) 'Opt-in beta blocked'
foreach ($Path in @('/beta/deviceManagement/roleScopeTags/id/settings', '/beta/deviceManagement/templates/id/assignments', '/beta/deviceManagement/intents/id/settings/id/settingDefinitions', '/v1.0/deviceManagement/managedDevices/id/roleAssignments', '/v1.0/deviceManagement/deviceConfigurations/id/deviceStatuses', '/v1.0/deviceManagement/auditEvents/id/assignments')) {
    Assert-Expansion (-not (Test-IntuneUri ('https://graph.microsoft.com' + $Path) -AllowExpansion)) 'Undocumented or unused route combination accepted'
}
foreach ($Uri in @('https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?%24expand=settings', 'https://graph.microsoft.com/v1.0/deviceAppManagement/vppTokens?$select=id,expirationDateTime,state,lastSyncDateTime,lastSyncStatus,automaticallyUpdateApps&%24select=token', 'https://graph.microsoft.com/v1.0/directory/deviceLocalCredentials?%24select=credentials')) {
    Assert-Expansion (-not (Test-IntuneUri $Uri -AllowExpansion)) 'Encoded/duplicate secret query accepted'
}
foreach ($Uri in @('https://graph.microsoft.com/v1.0/directory/deviceLocalCredentials/device-id', 'https://graph.microsoft.com/v1.0/informationProtection/bitlocker/recoveryKeys/key-id', 'https://graph.microsoft.com/v1.0/deviceAppManagement/vppTokens', 'https://graph.microsoft.com/v1.0/deviceAppManagement/vppTokens?$select=token', 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/policy/assign')) {
    Assert-Expansion (-not (Test-IntuneUri $Uri -AllowExpansion)) "Unsafe expansion endpoint: $Uri"
}
$Definition = @{ id = 'opaque'; baseUri = './Device/Vendor/MSFT/Policy/Config/Defender'; offsetUri = 'AllowCloudProtection'; version = '1'; options = @(@{ itemId = 'opaque_disabled_1'; optionValue = @{ value = 0 } }, @{ itemId = 'opaque_enabled_0'; optionValue = @{ value = 1 } }) }
$Facts = @(ConvertTo-IntuneSettingFacts @{ settingDefinitionId = 'opaque'; choiceSettingValue = @{ value = 'opaque_disabled_1' } } @($Definition) 'policy' 'setting')
Assert-Expansion ($Facts[0].value -eq 0 -and $Facts[0].resolution -eq 'Resolved') 'Choice suffix guessed instead of definition join'
foreach ($IdentityCase in @(
    @{ InstanceId = 1; DefinitionId = '1'; ChoiceId = 'option'; OptionId = 'option' },
    @{ InstanceId = '1'; DefinitionId = 1; ChoiceId = 'option'; OptionId = 'option' },
    @{ InstanceId = 'opaque'; DefinitionId = 'opaque'; ChoiceId = '1'; OptionId = 1 },
    @{ InstanceId = 'opaque'; DefinitionId = 'opaque'; ChoiceId = 1; OptionId = '1' },
    @{ InstanceId = 'opaque'; DefinitionId = 'opaque'; ChoiceId = $null; OptionId = $null },
    @{ InstanceId = 'opaque'; DefinitionId = 'opaque'; ChoiceId = $true; OptionId = 'True' },
    @{ InstanceId = @('opaque'); DefinitionId = 'opaque'; ChoiceId = 'option'; OptionId = 'option' },
    @{ InstanceId = 'opaque'; DefinitionId = 'opaque'; ChoiceId = @('option'); OptionId = 'option' },
    @{ InstanceId = 'opaque'; DefinitionId = @('opaque'); ChoiceId = 'option'; OptionId = 'option' },
    @{ InstanceId = 'opaque'; DefinitionId = 'opaque'; ChoiceId = 'option'; OptionId = @('option') },
    @{ InstanceId = $null; DefinitionId = ''; ChoiceId = 'option'; OptionId = 'option' },
    @{ InstanceId = ' '; DefinitionId = ' '; ChoiceId = 'option'; OptionId = 'option' },
    @{ InstanceId = $true; DefinitionId = 'True'; ChoiceId = 'option'; OptionId = 'option' },
    @{ InstanceId = @{ id = 'PRIVATE_VALUE' }; DefinitionId = 'opaque'; ChoiceId = 'option'; OptionId = 'option' },
    @{ InstanceId = 'opaque'; DefinitionId = 'OPAQUE'; ChoiceId = 'option'; OptionId = 'option' },
    @{ InstanceId = 'opaque'; DefinitionId = 'opaque'; ChoiceId = ''; OptionId = '' },
    @{ InstanceId = 'opaque'; DefinitionId = 'opaque'; ChoiceId = ' '; OptionId = ' ' },
    @{ InstanceId = 'opaque'; DefinitionId = 'opaque'; ChoiceId = '01'; OptionId = '1' },
    @{ InstanceId = 'opaque'; DefinitionId = 'opaque'; ChoiceId = 'OPTION'; OptionId = 'option' },
    @{ InstanceId = 'opaque'; DefinitionId = 'opaque'; ChoiceId = ' option'; OptionId = 'option' },
    @{ InstanceId = 'opaque'; DefinitionId = 'opaque'; ChoiceId = "opt`0ion"; OptionId = 'option' }
)) {
    foreach ($DecodeJson in @($false, $true)) {
        $IdentityDefinition = @{ id = $IdentityCase.DefinitionId; baseUri = './Device/Vendor/MSFT/Policy/Config/Defender'; offsetUri = 'AllowRealtimeMonitoring'; options = @(@{ itemId = $IdentityCase.OptionId; optionValue = @{ value = 1 } }) }
        $IdentityInstance = @{ settingDefinitionId = $IdentityCase.InstanceId; choiceSettingValue = @{ value = $IdentityCase.ChoiceId } }
        if ($DecodeJson) {
            $IdentityDefinition = $IdentityDefinition | ConvertTo-Json -Depth 10 | ConvertFrom-Json
            $IdentityInstance = $IdentityInstance | ConvertTo-Json -Depth 10 | ConvertFrom-Json
        }
        $IdentityFacts = @(ConvertTo-IntuneSettingFacts $IdentityInstance @($IdentityDefinition) 'policy' 'identity')
        Assert-Expansion ($IdentityFacts.Count -eq 1 -and $IdentityFacts[0].resolution -in @('UnresolvedValue', 'UnsupportedDefinition') -and -not $IdentityFacts[0].Contains('value')) 'Malformed identity was coerced into a resolved setting'
        Assert-Expansion (-not (($IdentityFacts | ConvertTo-Json -Depth 10) -match 'PRIVATE_VALUE|System.Collections')) 'Malformed identity was stringified into export metadata'
    }
}
foreach ($Token in @('0', '01', 'opaque_enabled_0')) {
    foreach ($DecodeJson in @($false, $true)) {
        $ExactDefinition = @{ id = $Token; baseUri = './Device/Vendor/MSFT/Policy/Config/Defender'; offsetUri = 'AllowRealtimeMonitoring'; options = @(@{ itemId = $Token; optionValue = @{ value = 0 } }) }
        $ExactInstance = @{ settingDefinitionId = $Token; choiceSettingValue = @{ value = $Token } }
        if ($DecodeJson) {
            $ExactDefinition = $ExactDefinition | ConvertTo-Json -Depth 10 | ConvertFrom-Json
            $ExactInstance = $ExactInstance | ConvertTo-Json -Depth 10 | ConvertFrom-Json
        }
        $ExactFacts = @(ConvertTo-IntuneSettingFacts $ExactInstance @($ExactDefinition) 'policy' 'exact')
        Assert-Expansion ($ExactFacts.Count -eq 1 -and $ExactFacts[0].resolution -ceq 'Resolved' -and $ExactFacts[0].value -eq 0) 'Exact opaque string match or zero scalar was lost'
    }
}
foreach ($DuplicateKind in @('definition', 'option')) {
    $DuplicateDefinition = @{ id = 'opaque'; baseUri = './Device/Vendor/MSFT/Policy/Config/Defender'; offsetUri = 'AllowRealtimeMonitoring'; options = @(@{ itemId = 'option'; optionValue = @{ value = 1 } }) }
    $DuplicateDefinitions = @($DuplicateDefinition)
    if ($DuplicateKind -eq 'definition') { $DuplicateDefinitions += $DuplicateDefinition.Clone() }
    else { $DuplicateDefinition.options += @{ itemId = 'option'; optionValue = @{ value = 0 } } }
    $DuplicateFacts = @(ConvertTo-IntuneSettingFacts @{ settingDefinitionId = 'opaque'; choiceSettingValue = @{ value = 'option' } } $DuplicateDefinitions 'policy' 'duplicate')
    Assert-Expansion ($DuplicateFacts.Count -eq 1 -and -not $DuplicateFacts[0].Contains('value')) 'Duplicate identity produced a resolved setting'
}
foreach ($ChoiceValue in @('missing-option', 'opaque_enabled_0')) {
    $AmbiguousFacts = @(ConvertTo-IntuneSettingFacts @{ settingDefinitionId = 'opaque'; choiceSettingValue = @{ value = $ChoiceValue }; simpleSettingValue = @{ value = 1 } } @($Definition) 'policy' 'ambiguous')
    Assert-Expansion ($AmbiguousFacts[0].resolution -eq 'UnresolvedValue' -and -not $AmbiguousFacts[0].Contains('value')) 'Mixed choice/simple payload fabricated a resolved value'
}
$ChildDefinition = @{ id = 'child'; baseUri = './Device/Vendor/MSFT/Policy/Config/Defender'; offsetUri = 'AllowBehaviorMonitoring' }
$ChildInstance = @{ settingDefinitionId = 'child'; simpleSettingValue = @{ value = 1 } }
foreach ($ChoiceDefect in @('missing-option', 'duplicate-option')) {
    $UnresolvedDefinition = $Definition.Clone()
    $UnresolvedChoice = 'missing-option'
    if ($ChoiceDefect -eq 'duplicate-option') {
        $UnresolvedChoice = 'opaque_enabled_0'
        $UnresolvedDefinition.options = @(@{ itemId = $UnresolvedChoice; optionValue = @{ value = 1 } }, @{ itemId = $UnresolvedChoice; optionValue = @{ value = 0 } })
    }
    $UnresolvedParent = @{ settingDefinitionId = 'opaque'; choiceSettingValue = @{ value = $UnresolvedChoice; children = @($ChildInstance) } }
    $UnresolvedFacts = @(ConvertTo-IntuneSettingFacts $UnresolvedParent @($UnresolvedDefinition, $ChildDefinition) 'policy' 'unresolved-choice')
    Assert-Expansion ($UnresolvedFacts.Count -eq 2 -and $UnresolvedFacts[1].resolution -ceq 'UnresolvedValue' -and -not $UnresolvedFacts[1].Contains('value')) "Unresolved $ChoiceDefect produced resolved child evidence"
}
foreach ($ParentCase in @('valid', 'valid-unknown-path', 'missing-option', 'duplicate-option', 'missing-definition', 'duplicate-definition', 'missing-option-value', 'scalar-option-value', 'array-option-value', 'template-option')) {
    foreach ($CollectionChoice in @($false, $true)) {
        foreach ($DecodeJson in @($false, $true)) {
            $ParentDefinition = @{ id = 'parent'; baseUri = './Device/Vendor/MSFT/Policy/Config/Defender'; offsetUri = 'AllowCloudProtection'; options = @(@{ itemId = 'chosen'; optionValue = @{ value = 1 } }, @{ itemId = 'sibling'; optionValue = @{ value = 0 } }) }
            switch ($ParentCase) {
                'valid-unknown-path' { $ParentDefinition.offsetUri = 'Unreviewed' }
                'missing-option' { $ParentDefinition.options = @($ParentDefinition.options[1]) }
                'duplicate-option' { $ParentDefinition.options += @{ itemId = 'chosen'; optionValue = @{ value = 0 } } }
                'missing-option-value' { $ParentDefinition.options[0].Remove('optionValue') }
                'scalar-option-value' { $ParentDefinition.options[0].optionValue = 1 }
                'array-option-value' { $ParentDefinition.options[0].optionValue = @(@{ value = 1 }) }
                'template-option' { $ParentDefinition.options[0].optionValue.settingValueTemplateReference = @{ useTemplateDefault = $true } }
            }
            $ParentDefinitions = @($ParentDefinition, $ChildDefinition)
            if ($ParentCase -eq 'missing-definition') { $ParentDefinitions = @($ChildDefinition) }
            if ($ParentCase -eq 'duplicate-definition') { $ParentDefinitions += $ParentDefinition.Clone() }
            $ParentInstance = @{ settingDefinitionId = 'parent' }
            $ChosenValue = @{ value = 'chosen'; children = @($ChildInstance) }
            if ($CollectionChoice) { $ParentInstance.choiceSettingCollectionValue = @($ChosenValue, @{ value = 'sibling'; children = @($ChildInstance) }) }
            else { $ParentInstance.choiceSettingValue = $ChosenValue }
            if ($DecodeJson) {
                $ParentInstance = $ParentInstance | ConvertTo-Json -Depth 15 | ConvertFrom-Json
                $ParentDefinitions = $ParentDefinitions | ConvertTo-Json -Depth 15 | ConvertFrom-Json
            }
            $ParentFacts = @(ConvertTo-IntuneSettingFacts $ParentInstance $ParentDefinitions 'policy' 'choice-parent')
            $ExpectedChild = if ($ParentCase -in @('valid', 'valid-unknown-path')) { 'Resolved' } elseif ($ParentCase -eq 'template-option') { 'UnresolvedTemplateDefault' } else { 'UnresolvedValue' }
            Assert-Expansion ($ParentFacts[1].resolution -ceq $ExpectedChild -and $ParentFacts[1].Contains('value') -eq ($ExpectedChild -ceq 'Resolved')) "Choice context lost: $ParentCase; collection=$CollectionChoice; JSON=$DecodeJson"
            if ($CollectionChoice) {
                $ExpectedSibling = if ($ParentCase -in @('missing-definition', 'duplicate-definition')) { 'UnresolvedValue' } else { 'Resolved' }
                Assert-Expansion ($ParentFacts.Count -eq 3 -and $ParentFacts[2].resolution -ceq $ExpectedSibling) "Choice sibling state changed: $ParentCase"
            } else { Assert-Expansion ($ParentFacts.Count -eq 2) 'Choice metadata count changed' }
        }
    }
}
foreach ($InvalidParent in @(
    @{ settingDefinitionId = @('opaque'); groupSettingCollectionValue = @(@{ children = @($ChildInstance) }) },
    @{ settingDefinitionId = 'opaque'; choiceSettingValue = @{ value = @('opaque_enabled_0'); children = @($ChildInstance) } },
    @{ settingDefinitionId = 'unknown'; choiceSettingValue = @{ value = $null; children = @($ChildInstance) } }
)) {
    foreach ($DecodeJson in @($false, $true)) {
        $ParentInput = $InvalidParent
        if ($DecodeJson) { $ParentInput = $ParentInput | ConvertTo-Json -Depth 15 | ConvertFrom-Json }
        $ParentFacts = @(ConvertTo-IntuneSettingFacts $ParentInput @($Definition, $ChildDefinition) 'policy' 'invalid-parent')
        Assert-Expansion ($ParentFacts.Count -eq 1 -and -not $ParentFacts[0].Contains('value')) 'Malformed parent identity leaked resolved descendants'
    }
}
foreach ($ParentKind in @('known', 'unknown', 'duplicate-definition', 'admx')) {
    foreach ($DecodeJson in @($false, $true)) {
        $ParentDefinition = $Definition.Clone()
        if ($ParentKind -eq 'unknown') { $ParentDefinition.offsetUri = 'Unreviewed' }
        if ($ParentKind -eq 'admx') { $ParentDefinition.baseUri = './Device/Vendor/MSFT'; $ParentDefinition.offsetUri = 'Policy/Config/InternetExplorer/DisableInternetExplorerLaunchViaCOM' }
        $Definitions = @($ParentDefinition, $ChildDefinition)
        if ($ParentKind -eq 'duplicate-definition') { $Definitions += $ParentDefinition.Clone() }
        $Instance = @{ settingDefinitionId = 'opaque'; choiceSettingValue = @{ value = 'opaque_enabled_0'; children = @($ChildInstance) }; simpleSettingValue = @{ value = '<enabled/>' } }
        if ($DecodeJson) {
            $Instance = $Instance | ConvertTo-Json -Depth 15 | ConvertFrom-Json
            $Definitions = $Definitions | ConvertTo-Json -Depth 15 | ConvertFrom-Json
        }
        $MixedFacts = @(ConvertTo-IntuneSettingFacts $Instance $Definitions 'policy' 'mixed')
        Assert-Expansion ($MixedFacts.Count -eq 1 -and -not $MixedFacts[0].Contains('value') -and -not $MixedFacts[0].Contains('admx')) "Mixed $ParentKind parent produced value or descendant evidence"
        Assert-Expansion ($MixedFacts[0].resolution -in @('UnresolvedValue', 'UnsupportedDefinition')) 'Mixed payload lost its unresolved marker'
    }
}
$GroupFacts = @(ConvertTo-IntuneSettingFacts @{ settingDefinitionId = 'group'; groupSettingCollectionValue = @(@{ children = @($ChildInstance) }) } @($ChildDefinition) 'policy' 'group')
Assert-Expansion ($GroupFacts.Count -eq 2 -and $GroupFacts[1].value -eq 1 -and $GroupFacts[1].resolution -eq 'Resolved') 'Valid nested-group child lost during ambiguity handling'
$Facts = @(ConvertTo-IntuneSettingFacts @{ settingDefinitionId = 'password'; simpleSettingValue = @{ value = 'SECRET' } } @(@{ id = 'password'; baseUri = './Device/Vendor/MSFT'; offsetUri = 'Unreviewed/Password' }) 'policy' 'setting')
Assert-Expansion (-not (($Facts | ConvertTo-Json -Depth 20).Contains('SECRET'))) 'Unknown setting secret retained'
$Facts = @(ConvertTo-IntuneSettingFacts @{ settingDefinitionId = 'opaque'; choiceSettingValue = @{ value = 'opaque_enabled_0'; settingValueTemplateReference = @{ useTemplateDefault = $true } } } @($Definition) 'policy' 'setting')
Assert-Expansion ($Facts[0].resolution -eq 'UnresolvedTemplateDefault') 'Template default guessed'
$DefaultedParent = @{ settingDefinitionId = 'opaque'; choiceSettingValue = @{ value = 'opaque_enabled_0'; settingValueTemplateReference = @{ useTemplateDefault = $true }; children = @($ChildInstance) } }
$DefaultedFacts = @(ConvertTo-IntuneSettingFacts $DefaultedParent @($Definition, $ChildDefinition) 'policy' 'defaulted-parent')
Assert-Expansion ($DefaultedFacts.Count -eq 2 -and $DefaultedFacts[1].resolution -ceq 'UnresolvedTemplateDefault' -and -not $DefaultedFacts[1].Contains('value')) 'Unresolved template parent produced resolved child evidence'
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
    @{ Reference = 'PRIVATE_TEMPLATE'; Explicit = $false },
    @{ Reference = @(@{ useTemplateDefault = $false }); Explicit = $false },
    @{ Reference = $false; Explicit = $false }
)) {
    foreach ($Kind in @('simple', 'choice', 'option', 'option-unknown-path', 'group', 'choice-collection')) {
        foreach ($DecodeJson in @($false, $true)) {
            $TemplateDefinition = $Definition.Clone()
            $ExplicitChild = @{ settingDefinitionId = 'child'; simpleSettingValue = @{ value = 1; settingValueTemplateReference = @{ useTemplateDefault = $false } } }
            $TemplateInstance = @{ settingDefinitionId = 'opaque' }
            switch ($Kind) {
                'simple' { $TemplateInstance.simpleSettingValue = @{ value = 1; settingValueTemplateReference = $TemplateCase.Reference } }
                'choice' { $TemplateInstance.choiceSettingValue = @{ value = 'opaque_enabled_0'; settingValueTemplateReference = $TemplateCase.Reference; children = @($ExplicitChild) } }
                { $_ -in @('option', 'option-unknown-path') } {
                    $TemplateInstance.choiceSettingValue = @{ value = 'opaque_enabled_0'; children = @($ExplicitChild) }
                    $TemplateDefinition.options = @(@{ itemId = 'opaque_enabled_0'; optionValue = @{ value = 1; settingValueTemplateReference = $TemplateCase.Reference } })
                    if ($Kind -eq 'option-unknown-path') { $TemplateDefinition.offsetUri = 'Unreviewed' }
                }
                { $_ -in @('group', 'choice-collection') } {
                    $TemplateInstance.settingDefinitionId = 'group'
                    if ($Kind -eq 'choice-collection') {
                        $TemplateDefinition.id = 'group'
                        $TemplateDefinition.options = @(@{ itemId = 'option'; optionValue = @{ value = 1 } }, @{ itemId = 'sibling'; optionValue = @{ value = 0 } })
                    }
                    $Member = if ($Kind -eq 'group') { 'groupSettingCollectionValue' } else { 'choiceSettingCollectionValue' }
                    $TemplateInstance[$Member] = @(
                        @{ value = 'option'; settingValueTemplateReference = $TemplateCase.Reference; children = @($ExplicitChild) },
                        @{ value = 'sibling'; settingValueTemplateReference = @{ useTemplateDefault = $false }; children = @($ExplicitChild) }
                    )
                }
            }
            $TemplateDefinitions = @($TemplateDefinition, $ChildDefinition)
            if ($DecodeJson) {
                $TemplateInstance = $TemplateInstance | ConvertTo-Json -Depth 20 | ConvertFrom-Json
                $TemplateDefinitions = $TemplateDefinitions | ConvertTo-Json -Depth 20 | ConvertFrom-Json
            }
            $TemplateFacts = @(ConvertTo-IntuneSettingFacts $TemplateInstance $TemplateDefinitions 'policy' 'template')
            $ExpectedResolution = if ($TemplateCase.Explicit) { 'Resolved' } else { 'UnresolvedTemplateDefault' }
            $TargetIndex = if ($Kind -eq 'simple') { 0 } else { 1 }
            Assert-Expansion ($TemplateFacts[$TargetIndex].resolution -ceq $ExpectedResolution) "Template context lost or guessed for $Kind; JSON=$DecodeJson; reference=$($TemplateCase.Reference | ConvertTo-Json -Depth 5 -Compress); actual=$($TemplateFacts[$TargetIndex].resolution)"
            Assert-Expansion ($TemplateFacts[$TargetIndex].Contains('value') -eq $TemplateCase.Explicit) "Template value boundary changed for $Kind"
            Assert-Expansion (-not (($TemplateFacts | ConvertTo-Json -Depth 15) -match 'PRIVATE_TEMPLATE')) 'Raw malformed template reference exported'
            if ($Kind -in @('group', 'choice-collection')) {
                Assert-Expansion ($TemplateFacts.Count -eq 3 -and $TemplateFacts[2].resolution -ceq 'Resolved' -and $TemplateFacts[2].value -eq 1) 'Default context leaked into an explicit sibling'
            }
        }
    }
}
$Grandchild = @{ settingDefinitionId = 'child'; simpleSettingValue = @{ value = 1; settingValueTemplateReference = @{ useTemplateDefault = $false } } }
$Middle = @{ settingDefinitionId = 'opaque'; choiceSettingValue = @{ value = 'opaque_enabled_0'; settingValueTemplateReference = @{ useTemplateDefault = $false }; children = @($Grandchild) } }
$TransitiveChoiceFacts = @(ConvertTo-IntuneSettingFacts @{ settingDefinitionId = 'opaque'; choiceSettingValue = @{ value = 'missing-option'; children = @($Middle) } } @($Definition, $ChildDefinition) 'policy' 'choice-transitive')
Assert-Expansion ($TransitiveChoiceFacts.Count -eq 3 -and $TransitiveChoiceFacts[1].resolution -ceq 'UnresolvedValue' -and $TransitiveChoiceFacts[2].resolution -ceq 'UnresolvedValue' -and -not $TransitiveChoiceFacts[2].Contains('value')) 'Valid child choice cleared inherited selection uncertainty'
$TransitiveFacts = @(ConvertTo-IntuneSettingFacts @{ settingDefinitionId = 'group'; groupSettingCollectionValue = @(@{ settingValueTemplateReference = @{ useTemplateDefault = $true }; children = @($Middle) }) } @($Definition, $ChildDefinition) 'policy' 'transitive')
Assert-Expansion ($TransitiveFacts.Count -eq 3 -and $TransitiveFacts[1].resolution -ceq 'UnresolvedTemplateDefault' -and $TransitiveFacts[2].resolution -ceq 'UnresolvedTemplateDefault' -and -not $TransitiveFacts[2].Contains('value')) 'Explicit child flag cleared inherited template uncertainty'
$AdmxChildDefinition = @{ id = 'admx-child'; baseUri = './Device/Vendor/MSFT'; offsetUri = 'Policy/Config/InternetExplorer/DisableInternetExplorerLaunchViaCOM' }
foreach ($Payload in @('<enabled/>', 'PRIVATE_UNREVIEWED_XML')) {
    $AdmxChild = @{ settingDefinitionId = 'admx-child'; simpleSettingValue = @{ value = $Payload } }
    $AdmxDefaultFacts = @(ConvertTo-IntuneSettingFacts @{ settingDefinitionId = 'group'; groupSettingCollectionValue = @(@{ settingValueTemplateReference = @{ useTemplateDefault = $true }; children = @($AdmxChild) }) } @($AdmxChildDefinition) 'policy' 'admx-default')
    Assert-Expansion ($AdmxDefaultFacts.Count -eq 2 -and $AdmxDefaultFacts[1].resolution -ceq 'UnresolvedTemplateDefault' -and -not $AdmxDefaultFacts[1].Contains('admx') -and -not $AdmxDefaultFacts[1].Contains('value')) 'Inherited template context decoded ADMX data'
    Assert-Expansion (-not (($AdmxDefaultFacts | ConvertTo-Json -Depth 10) -match 'PRIVATE_UNREVIEWED_XML')) 'Raw unresolved template payload escaped'
    $AdmxChoiceFacts = @(ConvertTo-IntuneSettingFacts @{ settingDefinitionId = 'opaque'; choiceSettingValue = @{ value = 'missing-option'; children = @($AdmxChild) } } @($Definition, $AdmxChildDefinition) 'policy' 'admx-choice')
    Assert-Expansion ($AdmxChoiceFacts.Count -eq 2 -and $AdmxChoiceFacts[1].resolution -ceq 'UnresolvedValue' -and -not $AdmxChoiceFacts[1].Contains('admx') -and -not $AdmxChoiceFacts[1].Contains('value')) 'Unresolved choice context decoded ADMX data'
    Assert-Expansion (-not (($AdmxChoiceFacts | ConvertTo-Json -Depth 10) -match 'PRIVATE_UNREVIEWED_XML')) 'Raw unresolved choice payload escaped'
}
$Request = {
    param($Uri)
    $Path = ([uri]$Uri).AbsolutePath
    $Rows = @()
    switch -Regex ($Path) {
        '/managedDevices$' { $Rows = @(@{ id = 'managed-sample'; deviceName = 'Synthetic device'; operatingSystem = 'Windows'; azureADDeviceId = '33333333-3333-4333-8333-333333333333'; managementAgent = 'mdm' }) }
        '/configurationPolicies$' { $Rows = @(@{ id = 'modern-1'; name = 'Renamed baseline'; platforms = 'windows10'; templateReference = @{ templateFamily = 'baseline'; templateId = 'template-1'; templateDisplayVersion = '25H2'; password = 'SECRET' }; roleScopeTagIds = @('tag-1'); isAssigned = $true }) }
        '/configurationPolicies/modern-1/settings$' { $Rows = @(@{ id = '0'; settingInstance = @{ settingDefinitionId = 'opaque'; choiceSettingValue = @{ value = 'opaque_enabled_0' } } }) }
        '/settingDefinitions$' { $Rows = @($Definition) }
        '/roleScopeTags$' { $Rows = @(@{ id = 'tag-1'; displayName = 'Example'; isBuiltIn = $false }) }
        '/deviceRegistrationPolicy$' { return @{ StatusCode = 200; Body = @{ id = 'deviceRegistrationPolicy'; localAdminPassword = @{ isEnabled = $true; password = 'SECRET' } } } }
        '/deviceLocalCredentials$' { $Rows = @(@{ id = '33333333-3333-4333-8333-333333333333'; lastBackupDateTime = '2026-09-17T08:00:00Z'; credentials = @(@{ passwordBase64 = 'SECRET' }) }) }
        '/recoveryKeys$' { $Rows = @(@{ id = 'key-metadata'; deviceId = '33333333-3333-4333-8333-333333333333'; key = 'SECRET'; volumeType = 'operatingSystemVolume' }) }
        '/applePushNotificationCertificate$' { return @{ StatusCode = 200; Body = @{ id = 'apns'; expirationDateTime = '2027-09-17T00:00:00Z'; certificate = 'SECRET' } } }
        '/vppTokens$' { $Rows = @(@{ id = 'vpp'; token = 'SECRET'; state = 'valid'; expirationDateTime = '2027-09-17T00:00:00Z'; lastSyncStatus = 'completed' }) }
    }
    return @{ StatusCode = 200; Body = (@{ value = $Rows } | ConvertTo-Json -Depth 30 | ConvertFrom-Json) }
}
$Document = Invoke-IntuneDiscoveryCore -SelectedTenant '22222222-2222-4222-8222-222222222222' -Configuration $true -Rbac $true -Entra $true -Recovery $true -Enrollment $true -Apple $true -Request $Request -Requirements @{ Assessor = 'Offline expansion fixture'; ScopeDescription = 'Synthetic configuration'; ScopeConfirmed = $true; MaxCollectionAgeHours = 24; ConfigurationReview = @{ ReferenceProfile = 'windows-25h2'; PolicyIds = @('modern-1'); DefenderPrimary = $true } }
Assert-Expansion ($Document.CollectionStatus.SecuritySettings.State -eq 'Complete') 'Settings collection incomplete'
Assert-Expansion ($Document.CollectionStatus.ModernAssignments.State -eq 'Unsupported' -and $Document.CollectionStatus.AutopilotAssignments.State -eq 'Unsupported') 'Quarantined assignment contracts looked complete'
$Quarantined = Get-IntuneCollection -Uri 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/policy/assignments' -Module ModernAssignments -AllowExpansion -Request { throw 'Must not request' }
Assert-Expansion ($Quarantined.Status.State -eq 'Unsupported' -and $Quarantined.Status.PagesRead -eq 0) 'Quarantined assignment request escaped guard'
Assert-Expansion ($Document.Inventory.SecuritySettings[0].value -eq 1) 'Decoded setting lost'
Assert-Expansion ($Document.CollectionStatus.SecuritySettings.CompletedParentIds[0] -eq 'modern-1') 'Per-policy settings completeness lost'
Assert-Expansion ($Document.Inventory.DeviceRegistrationPolicy[0].localAdminPassword.isEnabled -eq $true) 'Singleton lost'
$Json = $Document | ConvertTo-Json -Depth 30
Assert-Expansion (-not $Json.Contains('SECRET')) 'Export contains secret'
$Result = Get-IntuneCollection -Uri 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies' -Module ModernConfigurationPolicies -AllowExpansion -Request { @{ StatusCode = 200; Body = @{ value = @(); '@odata.nextLink' = 'https://graph.microsoft.com/beta/deviceManagement/intents' } } }
Assert-Expansion ($Result.Status.ErrorCode -eq 'BlockedEndpoint') 'Cross-resource paging accepted'
if ($env:ASSAY_INTUNE_EXPANSION_FIXTURE) { [IO.File]::WriteAllText($env:ASSAY_INTUNE_EXPANSION_FIXTURE, $Json, [Text.UTF8Encoding]::new($false)) }
Write-Output 'PASS: opt-in modules, metadata-only endpoints, definition choice mapping, unresolved defaults, child coverage and secret sanitization.'
. (Join-Path $PSScriptRoot 'Get-IntuneEndpointEvidence.ps1') -LibraryOnly
$Endpoint = New-IntuneEndpointEvidence -SelectedTenant 'test' -DeviceId 'test' -Read {
    param($Module)
    if ($Module -eq 'DefenderStatus') { return @{ AMRunningMode = 'Normal'; IsTamperProtected = $true; password = 'SECRET' } }
    throw 'Synthetic missing provider'
}
$SafeEndpoint = ConvertTo-IntuneSafeRow -Module EndpointEvidence -Row @{ id = 'test'; modules = $Endpoint.Modules; collectedAtUtc = $Endpoint.CollectedAtUtc }
Assert-Expansion ($SafeEndpoint.modules.DefenderStatus.Rows[0].IsTamperProtected -eq $true) 'Endpoint Boolean lost'
Assert-Expansion ($SafeEndpoint.modules.DeviceGuard.State -eq 'Error') 'Endpoint provider error hidden'
Assert-Expansion (-not (($SafeEndpoint | ConvertTo-Json -Depth 20).Contains('SECRET'))) 'Endpoint secret retained'
Write-Output 'PASS: injected endpoint provider states and field projection; no local device collection executed.'
$Admx = ConvertTo-IntuneAdmxMetadata '<enabled/><data id="UseTPMPIN" value="1"/>'
Assert-Expansion ($Admx.enabled -and $Admx.data.UseTPMPIN -eq 1 -and -not ($Admx | ConvertTo-Json).Contains('SECRET')) 'ADMX metadata parsing failed'
$DisabledAdmx = ConvertTo-IntuneAdmxMetadata '<disabled/>'
Assert-Expansion (-not $DisabledAdmx.enabled -and $DisabledAdmx.data.Count -eq 0) 'Disabled ADMX state lost'
Assert-Expansion ((ConvertTo-IntuneAdmxMetadata '<Enabled/><Data id="Value" value="1"/>').enabled) 'Documented title-case fragment rejected'
foreach ($Payload in @('<enabled/><unexpected/>', '<enabled>text</enabled>', '<enabled extra="1"/>', '<enabled/><disabled/>', '<enabled/><data id="same" value="1"/><data id="same" value="0"/>', '<enabled/><data id="value"/>', '<enabled/><data id="value" value="1"><nested/></data>', '<enabled/><data id="value" value="SECRET"/>', '<disabled/><data id="value" value="1"/>', '<enabled/>trailing text', '<enabled xmlns="urn:unreviewed"/>')) {
    $Rejected = $false
    try { ConvertTo-IntuneAdmxMetadata $Payload | Out-Null } catch { $Rejected = $true }
    Assert-Expansion $Rejected 'Unsupported ADMX fragment was partially resolved'
}
$InvalidAdmxFacts = @(ConvertTo-IntuneSettingFacts @{ settingDefinitionId = 'admx'; simpleSettingValue = @{ value = '<enabled/><unexpected/>' } } @(@{ id = 'admx'; baseUri = './Device/Vendor/MSFT'; offsetUri = 'Policy/Config/InternetExplorer/DisableInternetExplorerLaunchViaCOM' }) 'policy' 'setting')
Assert-Expansion ($InvalidAdmxFacts[0].resolution -eq 'UnresolvedAdmx' -and -not $InvalidAdmxFacts[0].Contains('admx')) 'Malformed ADMX produced resolved policy evidence'
$ScalarAdmxFacts = @(ConvertTo-IntuneSettingFacts @{ settingDefinitionId = 'admx'; simpleSettingValue = @{ value = 1 } } @(@{ id = 'admx'; baseUri = './Device/Vendor/MSFT'; offsetUri = 'Policy/Config/InternetExplorer/DisableInternetExplorerLaunchViaCOM' }) 'policy' 'setting')
Assert-Expansion ($ScalarAdmxFacts[0].resolution -ne 'Resolved' -and -not $ScalarAdmxFacts[0].Contains('value')) 'Numeric scalar bypassed ADMX parsing'
$Rejected = $false
try { Read-IntunePolicyXml '<!DOCTYPE root [<!ENTITY secret SYSTEM "file:///C:/secret">]><root>&secret;</root>' } catch { $Rejected = $true }
Assert-Expansion $Rejected 'External XML entity accepted'
$Policy = ConvertTo-IntuneAppControlMetadata '<SiPolicy xmlns="urn:schemas-microsoft-com:sipolicy" PolicyType="Base Policy"><PolicyID>{11111111-1111-4111-8111-111111111111}</PolicyID><BasePolicyID>{11111111-1111-4111-8111-111111111111}</BasePolicyID><Rules><Rule><Option>Enabled:Audit Mode</Option></Rule><Rule><Option>Enabled:Managed Installer</Option></Rule></Rules><Unexpected>SECRET</Unexpected></SiPolicy>'
Assert-Expansion ($Policy.auditMode -and $Policy.managedInstaller -and -not ($Policy | ConvertTo-Json).Contains('SECRET')) 'App Control policy metadata failed'
Write-Output 'PASS: bounded ADMX/App Control XML metadata parsing and external-entity rejection.'
$Temporary = Join-Path ([IO.Path]::GetTempPath()) ('intune-companions-' + [guid]::NewGuid().ToString())
[IO.Directory]::CreateDirectory($Temporary) | Out-Null
try {
    $Time = [datetime]::UtcNow.ToString('o')
    $Tenant = '22222222-2222-4222-8222-222222222222'
    $DefenderPath = Join-Path $Temporary 'defender.json'
    $EndpointPath = Join-Path $Temporary 'endpoint.json'
    $PolicyPath = Join-Path $Temporary 'policy.xml'
    $DefenderDocument = @{ SchemaVersion = '1.0'; TenantId = $Tenant; CollectedAtUtc = $Time; Result = @{ State = 'Complete'; PagesRead = 1; Rows = @(@{ id = 'mde'; aadDeviceId = '33333333-3333-4333-8333-333333333333'; healthStatus = 'Active'; onboardingStatus = 'onboarded'; lastSeen = $Time; lastIpAddress = 'SECRET' }) } }
    [IO.File]::WriteAllText($DefenderPath, ($DefenderDocument | ConvertTo-Json -Depth 10))
    [IO.File]::WriteAllText($EndpointPath, (@{ SchemaVersion = '1.0'; TenantId = $Tenant; DeviceId = '33333333-3333-4333-8333-333333333333'; CollectedAtUtc = $Time; Modules = $Endpoint.Modules } | ConvertTo-Json -Depth 15))
    [IO.File]::WriteAllText($PolicyPath, '<SiPolicy xmlns="urn:schemas-microsoft-com:sipolicy" PolicyType="Base Policy"><PolicyID>{11111111-1111-4111-8111-111111111111}</PolicyID><BasePolicyID>{11111111-1111-4111-8111-111111111111}</BasePolicyID><Rules><Rule><Option>Enabled:Audit Mode</Option></Rule></Rules></SiPolicy>')
    $Imported = Invoke-IntuneDiscoveryCore -SelectedTenant $Tenant -Configuration $true -Rbac $true -Entra $true -Recovery $true -Enrollment $true -Apple $true -Request $Request -DefenderPath $DefenderPath -EndpointPaths @($EndpointPath) -AppControlPaths @($PolicyPath) -Requirements @{ Assessor = 'Offline expansion fixture'; ScopeDescription = 'Synthetic configuration and companion evidence'; ScopeConfirmed = $true; MaxCollectionAgeHours = 24; ConfigurationReview = @{ ReferenceProfile = 'windows-25h2'; PolicyIds = @('modern-1'); DeviceIds = @('managed-sample'); DefenderPrimary = $true; MaxDefenderReportAgeHours = 48 } }
    Assert-Expansion ($Imported.CollectionStatus.DefenderMachines.ApiVersion -eq 'mde-v1' -and $Imported.Inventory.DefenderMachines[0].lastSeen) 'Defender import lost version/time'
    Assert-Expansion ($Imported.CollectionStatus.AppControlPolicies.ApiVersion -eq 'xml-1.0' -and $Imported.Inventory.AppControlPolicies[0].auditMode) 'App Control import lost XML state'
    Assert-Expansion ($Imported.Inventory.EndpointEvidence.Count -eq 1) 'Endpoint companion import missing'
    $Export = $Imported | ConvertTo-Json -Depth 30
    Assert-Expansion (-not $Export.Contains('SECRET')) 'Companion import leaked extra fields'
    if ($env:ASSAY_INTUNE_EXPANSION_FIXTURE) { [IO.File]::WriteAllText($env:ASSAY_INTUNE_EXPANSION_FIXTURE, $Export, [Text.UTF8Encoding]::new($false)) }
    $DefenderDocument.TenantId = 'different-tenant'
    [IO.File]::WriteAllText($DefenderPath, ($DefenderDocument | ConvertTo-Json -Depth 10))
    $Rejected = $false
    try { Invoke-IntuneDiscoveryCore -SelectedTenant $Tenant -Request $Request -DefenderPath $DefenderPath | Out-Null } catch { $Rejected = $_.Exception.Message -match 'tenant/schema mismatch' }
    Assert-Expansion $Rejected 'Cross-tenant Defender file accepted'
} finally { [IO.Directory]::Delete($Temporary, $true) }
Write-Output 'PASS: companion imports preserve identity/time/source, exclude secrets and reject tenant mismatch.'