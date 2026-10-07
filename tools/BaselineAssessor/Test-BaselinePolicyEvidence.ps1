#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
$Tokens = $null
$ParseErrors = $null
$Ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Invoke-BaselineCollection.ps1'), [ref]$Tokens, [ref]$ParseErrors)
if ($ParseErrors.Count) { throw 'Collector parse errors' }
foreach ($Name in @('Read-RegistryValue', 'Read-RegistryValues', 'Read-BaselineMappedPolicies')) {
    $Definition = $Ast.Find({ param($Node) $Node -is [Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -ceq $Name }, $true)
    . ([scriptblock]::Create($Definition.Extent.Text))
}
function Assert-Policy([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
$script:RegistryReadStates = @{}
$script:Values = @{ Zero = 0; Disabled = $false; Empty = ''; Items = @('one'); Path = '%TEMP%\file' }
$script:Stores = @{}
function Get-Item {
    param($LiteralPath, $ErrorAction)
    if ($LiteralPath -like '*Denied') { throw [UnauthorizedAccessException]::new('PRIVATE_SENTINEL') }
    if ($LiteralPath -like '*Missing') { throw [Management.Automation.ItemNotFoundException]::new('PRIVATE_SENTINEL') }
    if ($LiteralPath -like '*Broken') { throw 'PRIVATE_SENTINEL' }
    $Store = $script:Values
    if ($script:Stores.Count) {
        $Store = $script:Stores[$LiteralPath.Substring('Registry::'.Length)]
        if ($null -eq $Store) { throw [Management.Automation.ItemNotFoundException]::new('PRIVATE_SENTINEL') }
    }
    $Item = [pscustomobject]@{ Store = $Store }
    $Item | Add-Member ScriptMethod GetValueNames { @($this.Store.Keys) }
    $Item | Add-Member ScriptMethod GetValue {
        param($Name, $Default, $Options)
        if ($Options -ne [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames) { throw 'Registry expansion must be disabled' }
        return ,$this.Store[$Name]
    }
    return $Item
}
foreach ($Name in $script:Values.Keys) {
    $Value = Read-RegistryValue 'HKLM\Test' $Name
    Assert-Policy (($Value | ConvertTo-Json -Compress) -ceq ($script:Values[$Name] | ConvertTo-Json -Compress)) "Value changed: $Name"
    Assert-Policy ($script:RegistryReadStates["HKLM\Test\$Name"].state -ceq 'Present') "Missing Present state: $Name"
}
$null = Read-RegistryValue 'HKLM\Test' 'Absent'
Assert-Policy ($script:RegistryReadStates['HKLM\Test\Absent'].reason -ceq 'ValueNotFound') 'Missing value not distinguished'
foreach ($Case in @(@{Path='Missing';State='Missing';Reason='KeyNotFound'}, @{Path='Denied';State='Error';Reason='AccessDenied'}, @{Path='Broken';State='Error';Reason='ReadFailed'})) {
    $Path = 'HKLM\' + $Case.Path
    Assert-Policy ($null -eq (Read-RegistryValue $Path 'Value')) 'Failed read must have no value'
    $Evidence = $script:RegistryReadStates["$Path\Value"]
    Assert-Policy ($Evidence.state -ceq $Case.State -and $Evidence.reason -ceq $Case.Reason) 'Incorrect value read state'
    Assert-Policy ((Read-RegistryValues $Path).Count -eq 0) 'Failed bulk read must not have values'
    Assert-Policy ($script:RegistryReadStates[$Path].state -ceq $Case.State) 'Incorrect bulk read state'
}
Assert-Policy ((Read-RegistryValues 'HKLM\Test').Zero -eq 0) 'Bulk read lost zero'
Assert-Policy (-not ($script:RegistryReadStates | ConvertTo-Json -Depth 5).Contains('PRIVATE_SENTINEL')) 'Error evidence leaked a raw message'
Write-Output 'PASS: real registry helpers distinguish present, missing and error; preserve zero/false/arrays and do not expand paths or leak errors.'

$Area = $Ast.Find({ param($Node) $Node -is [Management.Automation.Language.AssignmentStatementAst] -and $Node.Left.Extent.Text -ceq '$registryBaselines' }, $true)
$Block = @($Area.Right.PipelineElements[0].CommandElements | Where-Object { $_ -is [Management.Automation.Language.ScriptBlockExpressionAst] })[0].ScriptBlock.Extent.Text
function ConvertTo-CollectorMaskedPath { param($Context, $Value) return ,$Value }
$script:Values = @{}
$script:RegistryReadStates = @{ 'HKLM\SOFTWARE\Microsoft\PolicyManager\providers\PRIVATE_SENTINEL\default\Device\System\AllowTelemetry' = @{ state = 'Present' } }
$RegistryArea = & ([scriptblock]::Create($Block.Substring(1, $Block.Length - 2)))
$null = Read-RegistryValue 'HKLM\Later\Area' 'PRIVATE_SENTINEL'
$RoundTrip = ($RegistryArea | ConvertTo-Json -Depth 10) | ConvertFrom-Json
Assert-Policy ($RoundTrip._readStates.'HKLM\SOFTWARE\Policies\Microsoft\Windows\System\EnableSmartScreen'.state -ceq 'Missing') 'Production registry area omitted per-value read states'
Assert-Policy (-not ($RegistryArea | ConvertTo-Json -Depth 10).Contains('PRIVATE_SENTINEL')) 'Registry read states leaked reads from other areas'
Remove-Item Function:\ConvertTo-CollectorMaskedPath

$CurrentPath = 'HKLM\SOFTWARE\Microsoft\PolicyManager\current\device\'
$ProviderGuid = '{11111111-2222-3333-4444-555555555555}'
$ProviderPath = "HKLM\SOFTWARE\Microsoft\PolicyManager\providers\$ProviderGuid\default\Device\"
function New-PolicyStores {
    $Stores = @{}
    foreach ($Binding in @(@('SmartScreen', 'EnableSmartScreenInShell', 0), @('DeviceGuard', 'EnableVirtualizationBasedSecurity', 1), @('System', 'AllowTelemetry', 3))) {
        $Stores[$CurrentPath + $Binding[0]] = @{ ($Binding[1] + '_ProviderSet') = 1; ($Binding[1] + '_WinningProvider') = $ProviderGuid; UnknownSecret_ProviderSet = 1 }
        $Stores[$ProviderPath + $Binding[0]] = @{ $Binding[1] = $Binding[2]; UnknownSecret = 'PRIVATE_SENTINEL' }
    }
    return $Stores
}
$script:Stores = New-PolicyStores
$Mapped = Read-BaselineMappedPolicies -Enrolled $true
Assert-Policy ($Mapped.Values.SmartScreen.EnableSmartScreenInShell -eq 0 -and $Mapped.Values.DeviceGuard.EnableVirtualizationBasedSecurity -eq 1 -and $Mapped.Values.System.AllowTelemetry -eq 3) 'Winning-provider policy values lost'
$MappedJson = $Mapped | ConvertTo-Json -Depth 6
Assert-Policy (-not $MappedJson.Contains('PRIVATE_SENTINEL') -and -not $MappedJson.Contains('11111111')) 'Unbound policy or provider ID was exported'
foreach ($BadValue in @('1', $true, 2, -1, @{}, @(1, 0))) {
    $script:Stores[$ProviderPath + 'DeviceGuard'].EnableVirtualizationBasedSecurity = $BadValue
    $Mapped = Read-BaselineMappedPolicies -Enrolled $true
    Assert-Policy ($Mapped.ReadStates.DeviceGuard.EnableVirtualizationBasedSecurity.state -ceq 'Invalid' -and $Mapped.Values.DeviceGuard.Count -eq 0) 'Invalid enum/type accepted'
}
$script:Stores = New-PolicyStores
$script:Stores[$CurrentPath + 'DeviceGuard'].EnableVirtualizationBasedSecurity_ProviderSet = 0
$Mapped = Read-BaselineMappedPolicies -Enrolled $true
Assert-Policy ($Mapped.ReadStates.DeviceGuard.EnableVirtualizationBasedSecurity.state -ceq 'NotConfigured' -and $Mapped.Values.DeviceGuard.Count -eq 0) 'Default without provider marker was treated as configured'
$script:Stores.Remove($CurrentPath + 'SmartScreen')
$script:Stores[$CurrentPath + 'System'].Remove('AllowTelemetry_ProviderSet')
$Mapped = Read-BaselineMappedPolicies -Enrolled $true
Assert-Policy ($Mapped.ReadStates.SmartScreen.EnableSmartScreenInShell.state -ceq 'NotConfigured' -and $Mapped.ReadStates.System.AllowTelemetry.state -ceq 'NotConfigured') 'Unmanaged policy was not labeled NotConfigured'
$script:Stores = New-PolicyStores
$script:Stores[$CurrentPath + 'System'].AllowTelemetry_WinningProvider = '..\..\PRIVATE_SENTINEL'
$script:Stores[$ProviderPath + 'DeviceGuard'].Remove('EnableVirtualizationBasedSecurity')
$Mapped = Read-BaselineMappedPolicies -Enrolled $true
Assert-Policy ($Mapped.ReadStates.System.AllowTelemetry.state -ceq 'Invalid' -and $Mapped.Values.System.Count -eq 0) 'Malformed winning provider was used as a path'
Assert-Policy ($Mapped.ReadStates.DeviceGuard.EnableVirtualizationBasedSecurity.state -ceq 'Missing' -and $Mapped.ReadStates.DeviceGuard.EnableVirtualizationBasedSecurity.providerSet) 'Managed value missing from provider store lost'
$Mapped = Read-BaselineMappedPolicies -Enrolled $false
Assert-Policy ($Mapped.ReadStates.System.AllowTelemetry.state -ceq 'NotCollected') 'Unenrolled policy scan should be skipped'
$script:Stores = New-PolicyStores
$script:Stores['HKLM\SOFTWARE\Microsoft\PolicyManager\current\device'] = @{}
function Test-Path { param($Path) return $true }
function Get-ChildItem {
    param($Path, $ErrorAction)
    if ($Path -eq 'HKLM:\SOFTWARE\Microsoft\Enrollments') {
        $Enrollment = [pscustomobject]@{ PSPath = 'HKLM:\SyntheticEnrollment' }
        $Enrollment | Add-Member ScriptMethod GetValueNames { @('ProviderID') }
        return $Enrollment
    }
}
function Get-ItemProperty { param($Path, $Name, $ErrorAction) return [pscustomobject]@{ ProviderID = 'Synthetic'; MDMWinsOverGP = 1 } }
function Write-CollectorProgress { }
$MdmArea = $Ast.Find({ param($Node) $Node -is [Management.Automation.Language.AssignmentStatementAst] -and $Node.Left.Extent.Text -ceq '$mdmEnrollment' }, $true)
$MdmBlock = @($MdmArea.Right.PipelineElements[0].CommandElements | Where-Object { $_ -is [Management.Automation.Language.ScriptBlockExpressionAst] })[0].ScriptBlock.Extent.Text
$CollectedMdm = & ([scriptblock]::Create($MdmBlock.Substring(1, $MdmBlock.Length - 2)))
Assert-Policy ($CollectedMdm.mdmEnrolled -and $CollectedMdm.policyValues.DeviceGuard.EnableVirtualizationBasedSecurity -eq 1 -and $CollectedMdm.policyReadStates.System.AllowTelemetry.state -ceq 'Present') 'Production MDM area did not export mapped evidence'
Remove-Item Function:\Test-Path, Function:\Get-ChildItem, Function:\Get-ItemProperty, Function:\Write-CollectorProgress
if ($env:ASSAY_BASELINE_POLICY_FIXTURE) {
    $Document = @{ systemInfo = @{ hostname = 'synthetic'; osBuild = '26200'; isServer = $false }; registryBaselines = @{
        'HKLM\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard\EnableVirtualizationBasedSecurity' = $null
        'HKLM\SOFTWARE\Policies\Microsoft\Windows\System\EnableSmartScreen' = $null
        'HKLM\SOFTWARE\Policies\Microsoft\Windows\DataCollection\AllowTelemetry' = $null
    }; mdmEnrollment = $CollectedMdm }
    [IO.File]::WriteAllText($env:ASSAY_BASELINE_POLICY_FIXTURE, ($Document | ConvertTo-Json -Depth 10), [Text.UTF8Encoding]::new($false))
}
Write-Output 'PASS: mapped PolicyManager reads preserve explicit zero, reject malformed enums and omit unknown/default fields.'