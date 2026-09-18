#Requires -Version 5.1
[CmdletBinding()]
param([string]$AssayCatalogPath)
$ErrorActionPreference = 'Stop'
$Tokens = $null
$ParseErrors = $null
$script:CollectorAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Invoke-BaselineCollection.ps1'), [ref]$Tokens, [ref]$ParseErrors)
if ($ParseErrors.Count) { throw 'Collector parse errors' }

function Get-CollectorArea([string]$Variable) {
    $Assignments = @($script:CollectorAst.FindAll({ param($Node) $Node -is [Management.Automation.Language.AssignmentStatementAst] -and $Node.Left.Extent.Text -ceq $Variable }, $true))
    if ($Assignments.Count -ne 1) { throw ('Expected one collection area: ' + $Variable) }
    $Blocks = @($Assignments[0].Right.PipelineElements[0].CommandElements | Where-Object { $_ -is [Management.Automation.Language.ScriptBlockExpressionAst] })
    if ($Blocks.Count -ne 1) { throw 'Expected one collection script block' }
    $Text = $Blocks[0].ScriptBlock.Extent.Text.Trim()
    return [scriptblock]::Create($Text.Substring(1, $Text.Length - 2))
}

$Firewall = Get-CollectorArea '$firewallProfiles'
$script:RequestedStore = $null
function Get-NetFirewallProfile {
    param([string]$PolicyStore)
    $script:RequestedStore = $PolicyStore
    foreach ($Entry in @(@('Domain','True'), @('Private','False'), @('Public','NotConfigured'))) {
        [pscustomobject]@{ Name = $Entry[0]; Enabled = $Entry[1]; DefaultInboundAction = 'Block'; DefaultOutboundAction = 'Allow'; AllowLocalFirewallRules = 'NotConfigured'; LogAllowed = $Entry[1]; LogBlocked = $Entry[1]; NotifyOnListen = $Entry[1]; LogFileName = 'synthetic.log'; LogMaxSizeKilobytes = 16384 }
    }
}
try { $Result = & $Firewall } finally { Remove-Item Function:\Get-NetFirewallProfile }
if ($script:RequestedStore -cne 'ActiveStore') { throw 'Firewall did not read effective ActiveStore' }
foreach ($Field in @('Enabled','LogAllowed','LogBlocked','NotifyOnListen')) {
    if ($Result.Domain[$Field] -ne $true -or $Result.Private[$Field] -ne $false -or $null -ne $Result.Public[$Field]) { throw ('GpoBoolean evidence corrupted: ' + $Field) }
}
function Get-NetFirewallProfile { throw 'Synthetic access failure' }
try { $Result = & $Firewall } finally { Remove-Item Function:\Get-NetFirewallProfile }
if ($Result._collectionFailed -ne $true -or $Result._error -ne 'Synthetic access failure') { throw 'Firewall provider failure was hidden' }
Write-Output 'PASS: baseline collector ActiveStore, tri-state firewall evidence and explicit provider errors; no device collection executed.'

$Defender = Get-CollectorArea '$defenderConfig'
function Get-MpPreference { [pscustomobject]@{ DisableRealtimeMonitoring = $false; DisableBehaviorMonitoring = $false; DisableIOAVProtection = $false } }
function Get-MpComputerStatus { [pscustomobject]@{ RealTimeProtectionEnabled = $false; BehaviorMonitorEnabled = $false; IoavProtectionEnabled = $false; IsTamperProtected = $true; TamperProtectionSource = 'Intune'; AMRunningMode = 'Passive Mode' } }
try { $Result = & $Defender } finally { Remove-Item Function:\Get-MpComputerStatus }
if ($Result.RealTimeProtectionEnabled -ne $false -or $Result.BehaviorMonitoringEnabled -ne $false -or $Result.IoavProtectionEnabled -ne $false -or $Result.RealTimeProtectionConfigured -ne $true) { throw 'Defender intent substituted for runtime state' }
if ($Result.TamperProtectionSource -ne 'Intune' -or $Result.IsTamperProtected -ne $true) { throw 'Tamper source and protection state conflated' }
function Get-MpComputerStatus { $null }
try { $Result = & $Defender } finally { Remove-Item Function:\Get-MpPreference, Function:\Get-MpComputerStatus }
if ($Result.Contains('RealTimeProtectionEnabled') -or $Result.Contains('BehaviorMonitoringEnabled')) { throw 'Unavailable runtime fabricated from preferences' }

$DeviceGuard = Get-CollectorArea '$credentialGuard'
function Read-RegistryValue { 1 }
function Get-CimInstance { $null }
try { $Result = & $DeviceGuard } finally { Remove-Item Function:\Get-CimInstance }
foreach ($Field in @('VbsRunning','CredentialGuardIsConfigured','CredentialGuardIsRunning','MemoryIntegrityIsConfigured','MemoryIntegrityIsRunning','HypervisorEnforcedCodeIntegrityEnabled','SecurityServicesRunning','SecurityServicesConfigured')) {
    if ($null -ne $Result[$Field]) { throw ('Missing Device Guard data manufactured a value: ' + $Field) }
}
function Get-CimInstance { [pscustomobject]@{ VirtualizationBasedSecurityStatus = 2; SecurityServicesRunning = @(1,2); SecurityServicesConfigured = @(1,2); RequiredSecurityProperties = @(); AvailableSecurityProperties = @() } }
try { $Result = & $DeviceGuard } finally { Remove-Item Function:\Get-CimInstance, Function:\Read-RegistryValue }
if ($Result.VbsRunning -ne $true -or $Result.CredentialGuardIsRunning -ne $true -or $Result.MemoryIntegrityIsRunning -ne $true) { throw 'Documented Device Guard runtime state lost' }
Write-Output 'PASS: observed Defender and Device Guard runtime evidence remains distinct from policy intent and unavailable providers.'

$PowerShellArea = Get-CollectorArea '$powershellConfig'
$systemInfo = @{ isServer = $false }
function Read-RegistryValue { $null }
function Get-ExecutionPolicy { 'RemoteSigned' }
function Get-WindowsOptionalFeature {
    param([switch]$Online, [string]$FeatureName)
    if (-not $Online -or $FeatureName -cne 'MicrosoftWindowsPowerShellV2') { throw 'Unexpected feature query' }
    [pscustomobject]@{ FeatureName = $FeatureName; State = $script:FeatureState }
}
foreach ($State in @('Enabled','Disabled','DisablePending','Unknown')) {
    $script:FeatureState = $State
    $Result = & $PowerShellArea
    if ($Result.legacyEngine.collectionState -ne 'Complete' -or $Result.legacyEngine.state -cne $State) { throw 'Feature state was guessed or lost' }
    if ($State -eq 'Disabled') { $DisabledFeature = $Result }
}
Remove-Item Function:\Get-WindowsOptionalFeature
function Get-WindowsOptionalFeature { throw 'Synthetic feature query failure' }
try { $Result = & $PowerShellArea } finally { Remove-Item Function:\Get-WindowsOptionalFeature }
if ($Result.legacyEngine.collectionState -ne 'Error' -or $null -ne $Result.legacyEngine.state) { throw 'Feature query error inferred removal' }
$systemInfo.isServer = $true
function Get-WindowsFeature {
    param([string]$Name)
    if ($Name -cne 'PowerShell-V2') { throw 'Unexpected server feature query' }
    [pscustomobject]@{ Name = $Name; InstallState = 'Removed' }
}
try { $Result = & $PowerShellArea } finally { Remove-Item Function:\Get-WindowsFeature }
if ($Result.legacyEngine.provider -ne 'Get-WindowsFeature' -or $Result.legacyEngine.state -ne 'Removed') { throw 'Server feature identity lost' }
$systemInfo.isServer = $null
try { $Result = & $PowerShellArea } finally { Remove-Item Function:\Read-RegistryValue, Function:\Get-ExecutionPolicy }
if ($Result.legacyEngine.collectionState -ne 'Unsupported') { throw 'Unknown platform was guessed' }
if ($env:ASSAY_BASELINE_COLLECTOR_FIXTURE) {
    [IO.File]::WriteAllText($env:ASSAY_BASELINE_COLLECTOR_FIXTURE, (@{ systemInfo = @{ hostname = 'synthetic'; isServer = $false }; powershellConfig = $DisabledFeature } | ConvertTo-Json -Depth 10), [Text.UTF8Encoding]::new($false))
}
Write-Output 'PASS: documented client/server PowerShell 2.0 feature reads, raw pending/unknown states and explicit errors; no feature queries executed.'

if ($AssayCatalogPath) {
    $Catalog = Get-Content -LiteralPath $AssayCatalogPath -Raw | ConvertFrom-Json
    $RegistryArea = Get-CollectorArea '$registryBaselines'
    function Read-RegistryValue { 1 }
    function Read-RegistryValues {
        param([string]$Path)
        $Values = @{}
        $Prefix = 'registryBaselines.' + $Path + '\'
        foreach ($Check in $Catalog.checks) {
            foreach ($Key in $Check.collectionKeys) {
                if ($Key.StartsWith($Prefix, [StringComparison]::OrdinalIgnoreCase)) {
                    $Leaf = $Key.Substring($Prefix.Length)
                    if (-not $Leaf.Contains('\')) { $Values[$Leaf] = 1 }
                }
            }
        }
        return $Values
    }
    try { $Registry = & $RegistryArea } finally { Remove-Item Function:\Read-RegistryValue, Function:\Read-RegistryValues }
    $Missing = @(foreach ($Check in $Catalog.checks | Where-Object type -eq 'Auto') {
        foreach ($Key in $Check.collectionKeys) {
            if ($Key.StartsWith('registryBaselines.') -and -not $Registry.ContainsKey($Key.Substring('registryBaselines.'.Length))) {
                [pscustomobject]@{ Check = $Check.id; Key = $Key }
            }
        }
    })
    if ($Missing.Count) { $Missing | Format-List; throw ('Automatic registry bindings not collected: ' + $Missing.Count) }
    Write-Output 'PASS: every declared automatic registry binding is covered by the current collector; provider success and policy applicability are separate.'
}