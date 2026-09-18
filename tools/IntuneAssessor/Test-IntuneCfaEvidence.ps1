#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Invoke-IntuneDiscovery.ps1') -LibraryOnly
. (Join-Path $PSScriptRoot 'Get-IntuneEndpointEvidence.ps1') -LibraryOnly
$Tokens = $null
$ParseErrors = $null
$Ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Get-IntuneEndpointEvidence.ps1'), [ref]$Tokens, [ref]$ParseErrors)
if ($ParseErrors.Count) { throw 'Endpoint companion parse errors' }
$Assignment = @($Ast.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.AssignmentStatementAst] -and $_.Left.Extent.Text -ceq '$Evidence' })
if ($Assignment.Count -ne 1) { throw 'Expected one production endpoint call' }
$Blocks = @($Assignment[0].Right.PipelineElements[0].CommandElements | Where-Object { $_ -is [Management.Automation.Language.ScriptBlockExpressionAst] })
if ($Blocks.Count -ne 1) { throw 'Expected one production provider callback' }
$Text = $Blocks[0].ScriptBlock.Extent.Text.Trim()
$ReadProviders = [scriptblock]::Create($Text.Substring(1, $Text.Length - 2))
function Get-MpPreference {
    [CmdletBinding()]param()
    if ($script:PreferenceFails) { throw 'Synthetic preference failure' }
    [pscustomobject]@{ EnableControlledFolderAccess = $script:CfaMode; DisableRealtimeMonitoring = $false; DisableBehaviorMonitoring = $true; SignatureFallbackOrder = $script:SourceOrder; SignatureScheduleDay = $script:ScheduleDay; SignatureUpdateInterval = $script:UpdateInterval; SignatureScheduleTime = 'DO_NOT_EXPORT'; SharedSignaturesPath = 'PRIVATE_PATH'; SignatureDefinitionUpdateFileSharesSources = 'PRIVATE_PATH'; ControlledFolderAccessProtectedFolders = @('PRIVATE_PATH'); ControlledFolderAccessAllowedApplications = @('PRIVATE_PATH'); Unreviewed = 'DO_NOT_EXPORT' }
}
function Get-MpComputerStatus {
    [CmdletBinding()]param()
    if ($script:StatusFails) { throw 'Synthetic status failure' }
    [pscustomobject]@{ AMRunningMode = 'Normal'; AntivirusEnabled = $true; RealTimeProtectionEnabled = $true; BehaviorMonitorEnabled = $false; AntivirusSignatureLastUpdated = [datetimeoffset]::new(2026,9,18,10,0,0,[timespan]::FromHours(2)); Unreviewed = 'DO_NOT_EXPORT' }
}
function Get-NetFirewallProfile { [CmdletBinding()]param($PolicyStore) throw 'Synthetic unavailable provider' }
function Get-BitLockerVolume { [CmdletBinding()]param() throw 'Synthetic unavailable provider' }
function Get-CimInstance { [CmdletBinding()]param($ClassName, $Namespace) throw 'Synthetic unavailable provider' }
$Tenant = '22222222-2222-4222-8222-222222222222'
$Device = '33333333-3333-4333-8333-333333333333'
$script:PreferenceFails = $false
$script:StatusFails = $false
$script:SourceOrder = 'InternalDefinitionUpdateServer|MicrosoftUpdateServer|MMPC'
$script:ScheduleDay = 8
$script:UpdateInterval = 0
$Temporary = Join-Path ([IO.Path]::GetTempPath()) ('intune-cfa-' + [guid]::NewGuid().ToString() + '.json')
try {
    foreach ($Mode in @(0,1,2,3,4,'Disabled','Enabled','AuditMode','BlockDiskModificationOnly','AuditDiskModificationOnly',$null,'unreviewed')) {
        $script:CfaMode = $Mode
        $Sample = New-IntuneEndpointEvidence -SelectedTenant $Tenant -DeviceId $Device -Read $ReadProviders
        $Decoded = ($Sample | ConvertTo-Json -Depth 15) | ConvertFrom-Json
        $SharedProperty = $Decoded.Modules.DefenderPreferences.Rows[0].PSObject.Properties['SharedSignaturesPathState']
        if ($null -eq $SharedProperty -or $SharedProperty.Value -cne 'NonEmpty') { throw 'Shared-signature observation was not derived from the provider value' }
        $FileSharesProperty = $Decoded.Modules.DefenderPreferences.Rows[0].PSObject.Properties['SignatureFileSharesState']
        if ($null -eq $FileSharesProperty -or $FileSharesProperty.Value -cne 'NonEmpty') { throw 'File-share source observation was not derived from the provider value' }
        if ($null -ne $Decoded.Modules.DefenderPreferences.Rows[0].PSObject.Properties['SignatureDefinitionUpdateFileSharesSources']) { throw 'Raw file-share source paths exported' }
        if ($null -ne $Decoded.Modules.DefenderPreferences.Rows[0].PSObject.Properties['SharedSignaturesPath']) { throw 'Raw shared-signature path exported' }
        $Safe = ConvertTo-IntuneEndpointModules $Decoded.Modules
        if ($Safe.DefenderPreferences.Rows[0]['SharedSignaturesPathState'] -cne 'NonEmpty') { throw 'Shared-signature state lost during import' }
        if ($Safe.DefenderPreferences.Rows[0]['SignatureFileSharesState'] -cne 'NonEmpty') { throw 'File-share state lost during import' }
        if ($Safe.DefenderPreferences.State -cne 'Complete' -or $Safe.DefenderPreferences.Rows.Count -ne 1) { throw 'Provider shape changed' }
        $Value = $Safe.DefenderPreferences.Rows[0]['EnableControlledFolderAccess']
        if ($null -eq $Mode) { if ($null -ne $Value) { throw 'Missing CFA mode fabricated' } }
        elseif ($Value -cne $Mode) { throw 'CFA value changed by projection' }
        if (($Sample | ConvertTo-Json -Depth 15) -match 'PRIVATE_PATH|DO_NOT_EXPORT') { throw 'Raw provider fields escaped projection' }
    }
    foreach ($Order in @('MicrosoftUpdateServer|MMPC',' MMPC | FileShares ','ConfigMgr|MMPC','MMPC||MicrosoftUpdateServer','',$null,42)) {
        $script:SourceOrder = $Order
        $Sample = New-IntuneEndpointEvidence -SelectedTenant $Tenant -DeviceId $Device -Read $ReadProviders
        $Safe = ConvertTo-IntuneEndpointModules (($Sample | ConvertTo-Json -Depth 15 | ConvertFrom-Json).Modules)
        if ($Safe.DefenderPreferences.Rows[0]['SignatureFallbackOrder'] -cne $Order) { throw 'Source order was guessed, reordered or discarded by projection' }
        if (($Sample | ConvertTo-Json -Depth 15) -match 'PRIVATE_PATH|DO_NOT_EXPORT') { throw 'Unreviewed update paths escaped projection' }
    }
    foreach ($Case in @(
        @{ Value = ''; Expected = 'Empty' },
        @{ Value = '\\PRIVATE_PATH\share'; Expected = 'NonEmpty' },
        @{ Value = 'not-a-validated-path'; Expected = 'NonEmpty' },
        @{ Value = '|'; Expected = 'NonEmpty' },
        @{ Value = '\\PRIVATE_PATH\one | \\PRIVATE_PATH\two'; Expected = 'NonEmpty' },
        @{ Value = $null; Expected = 'Unknown' },
        @{ Value = ' '; Expected = 'Unknown' },
        @{ Value = "`t`n"; Expected = 'Unknown' },
        @{ Value = 0; Expected = 'Unknown' },
        @{ Value = $false; Expected = 'Unknown' },
        @{ Value = @('PRIVATE_PATH'); Expected = 'Unknown' },
        @{ Value = @{ Path = 'PRIVATE_PATH' }; Expected = 'Unknown' }
    )) {
        foreach ($Binding in @(
            @{ Source = 'SharedSignaturesPath'; State = 'SharedSignaturesPathState'; Other = 'SignatureFileSharesState' },
            @{ Source = 'SignatureDefinitionUpdateFileSharesSources'; State = 'SignatureFileSharesState'; Other = 'SharedSignaturesPathState' }
        )) {
            foreach ($AsObject in @($false, $true)) {
                foreach ($Missing in @($false, $true)) {
                    $script:PathRow = @{ SharedSignaturesPath = 'PRIVATE_PATH'; SignatureDefinitionUpdateFileSharesSources = 'PRIVATE_PATH'; SharedSignaturesPathState = 'Empty'; SignatureFileSharesState = 'Empty'; SignatureFallbackOrder = 'MMPC' }
                    $script:PathRow[$Binding.Source] = $Case.Value
                    if ($Missing) { $script:PathRow.Remove($Binding.Source) }
                    if ($AsObject) { $script:PathRow = [pscustomobject]$script:PathRow }
                    $Sample = New-IntuneEndpointEvidence -SelectedTenant $Tenant -DeviceId $Device -Read {
                        param($Module)
                        if ($Module -eq 'DefenderPreferences') { $script:PathRow }
                    }
                    $Serialized = $Sample | ConvertTo-Json -Depth 15
                    if ($Serialized -match 'PRIVATE_PATH|not-a-validated-path|"(SharedSignaturesPath|SignatureDefinitionUpdateFileSharesSources)"') { throw 'Raw path value survived reduction' }
                    $Safe = ConvertTo-IntuneEndpointModules (($Serialized | ConvertFrom-Json).Modules)
                    $Expected = if ($Missing) { 'Unknown' } else { $Case.Expected }
                    if ($Safe.DefenderPreferences.Rows[0][$Binding.State] -cne $Expected) { throw 'Path reduction fabricated a state' }
                    if ($Safe.DefenderPreferences.Rows[0][$Binding.Other] -cne 'NonEmpty') { throw 'Path reduction crossed independent settings' }
                    if ($Safe.DefenderPreferences.Rows[0]['SignatureFallbackOrder'] -cne 'MMPC') { throw 'Path reduction lost adjacent preference' }
                }
            }
        }
    }
    foreach ($InvalidState in @('empty', 'NonEmpty ', 'PRIVATE_PATH', '', $null, $true, 1, @('Empty'), @{ State = 'Empty' })) {
        $Safe = ConvertTo-IntuneEndpointModules @{ DefenderPreferences = @{ State = 'Complete'; Rows = @(@{ SharedSignaturesPathState = $InvalidState; SignatureFileSharesState = $InvalidState; SharedSignaturesPath = 'PRIVATE_PATH'; SignatureDefinitionUpdateFileSharesSources = 'PRIVATE_PATH'; SignatureFallbackOrder = 'MMPC' }) } }
        if ($Safe.DefenderPreferences.Rows[0].Contains('SharedSignaturesPathState') -or $Safe.DefenderPreferences.Rows[0].Contains('SignatureFileSharesState') -or ($Safe | ConvertTo-Json -Depth 15) -match 'PRIVATE_PATH') { throw 'Importer retained unsupported state or raw path' }
    }
    $script:SourceOrder = 'InternalDefinitionUpdateServer|MicrosoftUpdateServer|MMPC'
    foreach ($Cadence in @(@(8,0),@(0,0),@(8,24),@('Monday',4),@('Never',0),@($null,$null),@('unknown',25))) {
        $script:ScheduleDay = $Cadence[0]
        $script:UpdateInterval = $Cadence[1]
        $Sample = New-IntuneEndpointEvidence -SelectedTenant $Tenant -DeviceId $Device -Read $ReadProviders
        $Safe = ConvertTo-IntuneEndpointModules (($Sample | ConvertTo-Json -Depth 15 | ConvertFrom-Json).Modules)
        if ($Safe.DefenderPreferences.Rows[0]['SignatureScheduleDay'] -cne $Cadence[0] -or $Safe.DefenderPreferences.Rows[0]['SignatureUpdateInterval'] -cne $Cadence[1]) { throw 'Cadence values were changed or defaulted during projection' }
        if (($Sample | ConvertTo-Json -Depth 15) -match 'PRIVATE_PATH|DO_NOT_EXPORT') { throw 'Unreviewed schedule time or paths exported' }
    }
    $script:ScheduleDay = 8
    $script:UpdateInterval = 0
    foreach ($FailedProvider in @('DefenderPreferences','DefenderStatus')) {
        $script:PreferenceFails = $FailedProvider -eq 'DefenderPreferences'
        $script:StatusFails = $FailedProvider -eq 'DefenderStatus'
        $Sample = New-IntuneEndpointEvidence -SelectedTenant $Tenant -DeviceId $Device -Read $ReadProviders
        if ($Sample.Modules[$FailedProvider].State -cne 'Error' -or $Sample.Modules[$FailedProvider].Rows.Count -ne 0) { throw 'Provider failure hidden' }
    }
    $script:PreferenceFails = $false
    $script:StatusFails = $false
    $script:CfaMode = 3
    $Sample = New-IntuneEndpointEvidence -SelectedTenant $Tenant -DeviceId $Device -Read $ReadProviders
    [IO.File]::WriteAllText($Temporary, ($Sample | ConvertTo-Json -Depth 15), [Text.UTF8Encoding]::new($false))
    $Request = {
        param($Uri)
        $Rows = @()
        if (([uri]$Uri).AbsolutePath -eq '/v1.0/deviceManagement/managedDevices') {
            $Rows = @(@{ id = 'device'; azureADDeviceId = '33333333-3333-4333-8333-333333333333'; operatingSystem = 'Windows' })
        }
        @{ StatusCode = 200; Body = @{ value = $Rows } }
    }
    $Document = Invoke-IntuneDiscoveryCore -SelectedTenant $Tenant -EndpointPaths @($Temporary) -Request $Request -Requirements @{
        ScopeConfirmed = $true; ScopeDescription = 'Synthetic CFA endpoint sample'; Assessor = 'Test'; MaxCollectionAgeHours = 24
        ConfigurationReview = @{ DeviceIds = @('device') }
    }
    if ($Document.CollectionStatus.EndpointEvidence.State -cne 'Complete' -or $Document.Inventory.EndpointEvidence[0].modules.DefenderPreferences.Rows[0].EnableControlledFolderAccess -ne 3) { throw 'Production import dropped CFA evidence' }
    if ($Document.Inventory.EndpointEvidence[0].modules.DefenderPreferences.Rows[0].SharedSignaturesPathState -cne 'NonEmpty') { throw 'Production import lost shared-signature state' }
    if ($Document.Inventory.EndpointEvidence[0].modules.DefenderPreferences.Rows[0].SignatureFileSharesState -cne 'NonEmpty') { throw 'Production import lost file-share state' }
    if ($Document.Inventory.EndpointEvidence[0].modules.DefenderStatus.Rows[0].AntivirusSignatureLastUpdated -cne '2026-09-18T08:00:00.0000000Z') { throw 'Production file import changed signature instant' }
    if ($Document.Inventory.EndpointEvidence[0].modules.DefenderPreferences.Rows[0].SignatureFallbackOrder -cne $script:SourceOrder) { throw 'Production import dropped or reordered update sources' }
    if ($Document.Inventory.EndpointEvidence[0].modules.DefenderPreferences.Rows[0].SignatureScheduleDay -ne 8 -or $Document.Inventory.EndpointEvidence[0].modules.DefenderPreferences.Rows[0].SignatureUpdateInterval -ne 0) { throw 'Production import lost cadence values' }
    if (($Document | ConvertTo-Json -Depth 30) -match 'PRIVATE_PATH|DO_NOT_EXPORT') { throw 'Unreviewed fields retained' }
    foreach ($Target in @('Disabled','Enabled','AuditMode','BlockDiskModificationOnly','AuditDiskModificationOnly')) {
        $Targeted = Invoke-IntuneDiscoveryCore -SelectedTenant $Tenant -EndpointPaths @($Temporary) -Request $Request -Requirements @{
            ScopeConfirmed = $true; ScopeDescription = 'Synthetic explicit target'; Assessor = 'Test'; MaxCollectionAgeHours = 24
            ConfigurationReview = @{ DeviceIds = @('device'); ControlledFolderAccessTarget = $Target }
        }
        if ($Targeted.AssessmentRequirements.ConfigurationReview.ControlledFolderAccessTarget -cne $Target -or $Targeted.Inventory.EndpointEvidence[0].modules.DefenderPreferences.Rows[0].EnableControlledFolderAccess -ne 3) { throw 'Customer target or observed evidence changed during import' }
    }
    foreach ($Limit in @(0,2,87600)) {
        $Limited = Invoke-IntuneDiscoveryCore -SelectedTenant $Tenant -EndpointPaths @($Temporary) -Request $Request -Requirements @{
            ScopeConfirmed = $true; ScopeDescription = 'Synthetic signature-age limit'; Assessor = 'Test'; MaxCollectionAgeHours = 24
            ConfigurationReview = @{ DeviceIds = @('device'); MaxSignatureAgeHours = $Limit }
        }
        if ($Limited.AssessmentRequirements.ConfigurationReview.MaxSignatureAgeHours -ne $Limit -or $Limited.Inventory.EndpointEvidence[0].modules.DefenderStatus.Rows[0].AntivirusSignatureLastUpdated -cne '2026-09-18T08:00:00.0000000Z') { throw 'Signature-age requirement or evidence changed during import' }
    }
    $script:MixedCorrelation = $false
    $script:IdentityCorrelation = $false
    $CorrelationRequest = {
        param($Uri)
        $Path = ([uri]$Uri).AbsolutePath
        $Rows = @()
        switch -Regex ($Path) {
            '/managedDevices$' { $Rows = @(@{ id = 'device'; azureADDeviceId = '33333333-3333-4333-8333-333333333333'; operatingSystem = 'Windows' }) }
            '/configurationPolicies$' { $Rows = @(@{ id = 'policy-on'; name = 'Synthetic allow'; platforms = 'windows10'; isAssigned = $false }, @{ id = 'policy-off'; name = 'Synthetic disallow'; platforms = 'windows10'; isAssigned = $false }) }
            '/configurationPolicies/(policy-on|policy-off)/settings$' {
                $Choice = if ($Path -eq '/beta/deviceManagement/configurationPolicies/policy-on/settings') { 'opaque_enabled_0' } else { 'opaque_disabled_1' }
                $BehaviorChoice = if ($Choice -ceq 'opaque_enabled_0') { 'opaque_disabled_1' } else { 'opaque_enabled_0' }
                $Rows = @(@{ id = '0'; settingInstance = @{ settingDefinitionId = 'rtp'; choiceSettingValue = @{ value = $Choice } } }, @{ id = '1'; settingInstance = @{ settingDefinitionId = 'behavior'; choiceSettingValue = @{ value = $BehaviorChoice } } })
                if ($script:MixedCorrelation) {
                    $Rows[0].settingInstance.choiceSettingValue.value = 'missing-option'
                    foreach ($Row in $Rows) { $Row.settingInstance.simpleSettingValue = @{ value = 1 } }
                }
                if ($script:IdentityCorrelation) {
                    $Rows[0].settingInstance.settingDefinitionId = 1
                    $Rows[1].settingInstance.choiceSettingValue.value = '1'
                }
            }
            '/settings/0/settingDefinitions$' { $Rows = @(@{ id = 'rtp'; baseUri = './Device/Vendor/MSFT/Policy/Config/Defender'; offsetUri = 'AllowRealtimeMonitoring'; version = '1'; options = @(@{ itemId = 'opaque_enabled_0'; optionValue = @{ value = 1 } }, @{ itemId = 'opaque_disabled_1'; optionValue = @{ value = 0 } }) }) }
            '/settings/1/settingDefinitions$' { $Rows = @(@{ id = 'behavior'; baseUri = './Device/Vendor/MSFT/Policy/Config/Defender'; offsetUri = 'AllowBehaviorMonitoring'; version = '1'; options = @(@{ itemId = 'opaque_enabled_0'; optionValue = @{ value = 1 } }, @{ itemId = 'opaque_disabled_1'; optionValue = @{ value = 0 } }) }) }
        }
        if ($script:IdentityCorrelation -and $Path.EndsWith('/settings/0/settingDefinitions')) { $Rows[0].id = '1' }
        if ($script:IdentityCorrelation -and $Path.EndsWith('/settings/1/settingDefinitions')) { $Rows[0].options = @(@{ itemId = 1; optionValue = @{ value = 1 } }) }
        @{ StatusCode = 200; Body = (@{ value = $Rows } | ConvertTo-Json -Depth 20 | ConvertFrom-Json) }
    }
    $Correlation = Invoke-IntuneDiscoveryCore -SelectedTenant $Tenant -Configuration $true -EndpointPaths @($Temporary) -Request $CorrelationRequest -Requirements @{
        ScopeConfirmed = $true; ScopeDescription = 'Synthetic analyst reference, not assigned policy'; Assessor = 'Test'; MaxCollectionAgeHours = 24
        ConfigurationReview = @{ DeviceIds = @('device'); PolicyIds = @('policy-on') }
    }
    if ($Correlation.CollectionStatus.SecuritySettings.State -cne 'Complete' -or $Correlation.CollectionStatus.SecuritySettings.CompletedParentIds.Count -ne 2 -or $Correlation.Inventory.SecuritySettings.Count -ne 4) { throw 'Monitoring settings coverage lost' }
    foreach ($Fact in $Correlation.Inventory.SecuritySettings) {
        $Expected = if ($Fact.parentId -ceq 'policy-on') { 1 } else { 0 }
        if ($Fact.cspUri -ceq 'Policy/Config/Defender/AllowBehaviorMonitoring') { $Expected = 1 - $Expected }
        if ($Fact.cspUri -cnotin @('Policy/Config/Defender/AllowRealtimeMonitoring', 'Policy/Config/Defender/AllowBehaviorMonitoring') -or $Fact.resolution -cne 'Resolved' -or $Fact.value -ne $Expected) { throw 'Monitoring setting definition join changed or guessed an option suffix' }
    }
    if ($Correlation.Inventory.EndpointEvidence[0].modules.DefenderPreferences.Rows[0].DisableRealtimeMonitoring -isnot [bool] -or $Correlation.Inventory.EndpointEvidence[0].modules.DefenderPreferences.Rows[0].DisableRealtimeMonitoring -or -not $Correlation.Inventory.EndpointEvidence[0].modules.DefenderStatus.Rows[0].RealTimeProtectionEnabled) { throw 'Endpoint correlation booleans lost' }
    if ($Correlation.Inventory.EndpointEvidence[0].modules.DefenderPreferences.Rows[0].DisableBehaviorMonitoring -isnot [bool] -or -not $Correlation.Inventory.EndpointEvidence[0].modules.DefenderPreferences.Rows[0].DisableBehaviorMonitoring -or $Correlation.Inventory.EndpointEvidence[0].modules.DefenderStatus.Rows[0].BehaviorMonitorEnabled -isnot [bool] -or $Correlation.Inventory.EndpointEvidence[0].modules.DefenderStatus.Rows[0].BehaviorMonitorEnabled) { throw 'Behavior monitoring booleans lost or confused with real-time values' }
    if ($Correlation.CollectionStatus.ModernAssignments.State -cne 'Unsupported') { throw 'Synthetic correlation invented assignment coverage' }
    $CorrelationJson = $Correlation | ConvertTo-Json -Depth 30
    if ($CorrelationJson -match 'PRIVATE_PATH|DO_NOT_EXPORT') { throw 'Correlation export leaked unreviewed fields' }
    if ($env:ASSAY_INTUNE_REALTIME_FIXTURE) { [IO.File]::WriteAllText($env:ASSAY_INTUNE_REALTIME_FIXTURE, $CorrelationJson, [Text.UTF8Encoding]::new($false)) }
    $script:MixedCorrelation = $true
    try {
        $AmbiguousDocument = Invoke-IntuneDiscoveryCore -SelectedTenant $Tenant -Configuration $true -EndpointPaths @($Temporary) -Request $CorrelationRequest -Requirements @{
            ScopeConfirmed = $true; ScopeDescription = 'Synthetic ambiguous policy values'; Assessor = 'Test'; MaxCollectionAgeHours = 24
            ConfigurationReview = @{ DeviceIds = @('device'); PolicyIds = @('policy-on'); ReferenceProfile = 'windows-25h2'; DefenderPrimary = $true }
        }
    } finally { $script:MixedCorrelation = $false }
    if ($AmbiguousDocument.Inventory.SecuritySettings.Count -ne 4) { throw 'Ambiguous settings were silently dropped' }
    foreach ($Fact in $AmbiguousDocument.Inventory.SecuritySettings) {
        if ($Fact.resolution -cne 'UnresolvedValue' -or $Fact.Contains('value') -or $Fact.Contains('admx') -or $Fact.cspUri -cnotin @('Policy/Config/Defender/AllowRealtimeMonitoring', 'Policy/Config/Defender/AllowBehaviorMonitoring')) { throw 'Ambiguous production setting fabricated a decoded reference' }
    }
    $MixedJson = $AmbiguousDocument | ConvertTo-Json -Depth 30
    if ($MixedJson -match 'PRIVATE_PATH|DO_NOT_EXPORT|missing-option|simpleSettingValue') { throw 'Ambiguous raw payload leaked into export' }
    if ($env:ASSAY_INTUNE_MIXED_SETTING_FIXTURE) { [IO.File]::WriteAllText($env:ASSAY_INTUNE_MIXED_SETTING_FIXTURE, $MixedJson, [Text.UTF8Encoding]::new($false)) }
    $script:IdentityCorrelation = $true
    try {
        $IdentityDocument = Invoke-IntuneDiscoveryCore -SelectedTenant $Tenant -Configuration $true -EndpointPaths @($Temporary) -Request $CorrelationRequest -Requirements @{
            ScopeConfirmed = $true; ScopeDescription = 'Synthetic malformed setting identities'; Assessor = 'Test'; MaxCollectionAgeHours = 24
            ConfigurationReview = @{ DeviceIds = @('device'); PolicyIds = @('policy-on'); ReferenceProfile = 'windows-25h2'; DefenderPrimary = $true }
        }
    } finally { $script:IdentityCorrelation = $false }
    if ($IdentityDocument.Inventory.SecuritySettings.Count -ne 4) { throw 'Invalid-identity settings silently dropped' }
    foreach ($Fact in $IdentityDocument.Inventory.SecuritySettings) {
        if ($Fact.Contains('value') -or $Fact.Contains('admx')) { throw 'Malformed identity produced a decoded reference' }
        if ($Fact.definitionId -ceq '') {
            if ($Fact.resolution -cne 'UnsupportedDefinition' -or $Fact.Contains('cspUri')) { throw 'Numeric instance identity was preserved or resolved' }
        } elseif ($Fact.definitionId -ceq 'behavior') {
            if ($Fact.resolution -cne 'UnresolvedValue' -or $Fact.cspUri -cne 'Policy/Config/Defender/AllowBehaviorMonitoring') { throw 'Numeric option identity was resolved' }
        } else { throw 'Unexpected identity metadata' }
    }
    $IdentityJson = $IdentityDocument | ConvertTo-Json -Depth 30
    if ($IdentityJson -match 'PRIVATE_PATH|DO_NOT_EXPORT|choiceSettingValue|optionValue') { throw 'Invalid identity export retained raw payload' }
    if ($env:ASSAY_INTUNE_IDENTITY_FIXTURE) { [IO.File]::WriteAllText($env:ASSAY_INTUNE_IDENTITY_FIXTURE, $IdentityJson, [Text.UTF8Encoding]::new($false)) }
    $Sample.TenantId = '44444444-4444-4444-8444-444444444444'
    [IO.File]::WriteAllText($Temporary, ($Sample | ConvertTo-Json -Depth 15), [Text.UTF8Encoding]::new($false))
    $Rejected = $false
    try { Invoke-IntuneDiscoveryCore -SelectedTenant $Tenant -EndpointPaths @($Temporary) -Request $Request | Out-Null }
    catch { $Rejected = $_.Exception.Message -eq 'Endpoint evidence tenant/schema mismatch.' }
    if (-not $Rejected) { throw 'Wrong-tenant sample accepted' }
    if ($env:ASSAY_INTUNE_CFA_FIXTURE) {
        [IO.File]::WriteAllText($env:ASSAY_INTUNE_CFA_FIXTURE, ($Document | ConvertTo-Json -Depth 30), [Text.UTF8Encoding]::new($false))
    }
} finally {
    Remove-Item Function:\Get-MpPreference, Function:\Get-MpComputerStatus, Function:\Get-NetFirewallProfile, Function:\Get-BitLockerVolume, Function:\Get-CimInstance
    if (Test-Path -LiteralPath $Temporary) { Remove-Item -LiteralPath $Temporary }
}
Write-Output 'PASS: production CFA/source-order/cadence projection, value preservation, failure states, tenant-bound companion import and privacy; all providers and Graph calls mocked.'