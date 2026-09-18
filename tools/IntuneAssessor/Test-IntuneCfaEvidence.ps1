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
    [pscustomobject]@{ EnableControlledFolderAccess = $script:CfaMode; SignatureFallbackOrder = $script:SourceOrder; SignatureScheduleDay = $script:ScheduleDay; SignatureUpdateInterval = $script:UpdateInterval; SignatureScheduleTime = 'DO_NOT_EXPORT'; SharedSignaturesPath = 'PRIVATE_PATH'; SignatureDefinitionUpdateFileSharesSources = 'PRIVATE_PATH'; ControlledFolderAccessProtectedFolders = @('PRIVATE_PATH'); ControlledFolderAccessAllowedApplications = @('PRIVATE_PATH'); Unreviewed = 'DO_NOT_EXPORT' }
}
function Get-MpComputerStatus {
    [CmdletBinding()]param()
    if ($script:StatusFails) { throw 'Synthetic status failure' }
    [pscustomobject]@{ AMRunningMode = 'Normal'; AntivirusEnabled = $true; RealTimeProtectionEnabled = $true; Unreviewed = 'DO_NOT_EXPORT' }
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
        $Safe = ConvertTo-IntuneEndpointModules $Decoded.Modules
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