#Requires -Version 5.1
[CmdletBinding()]
param([string]$UpstreamPath)
$ErrorActionPreference = 'Stop'
$Tokens = $null
$ParseErrors = $null
$Path = Join-Path $PSScriptRoot 'Invoke-BaselineCollection.ps1'
$Ast = [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$Tokens, [ref]$ParseErrors)
if ($ParseErrors.Count) { throw 'Collector parse errors' }
$Functions = @($Ast.FindAll({ param($Node) $Node -is [Management.Automation.Language.FunctionDefinitionAst] }, $true))
$Detector = @($Functions | Where-Object Name -ceq 'Get-SpeculationControlSettings')
$Wrapper = @($Functions | Where-Object Name -ceq 'Get-BaselineSpeculationEvidence')
if ($Detector.Count -ne 1 -or $Wrapper.Count -ne 1) { throw 'Expected one embedded detector and wrapper' }
function Get-FunctionHash([string]$Text) {
    $Algorithm = [Security.Cryptography.SHA256]::Create()
    try { [BitConverter]::ToString($Algorithm.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text.Replace("`r`n","`n").Replace("`r","`n")))).Replace('-','') }
    finally { $Algorithm.Dispose() }
}
$ExpectedHash = '6ACA20A3EAD9E45CC9E6043223502B09DBFFB915A8D0EBF44700107985C87E21'
if ((Get-FunctionHash $Detector[0].Extent.Text) -cne $ExpectedHash) { throw 'Embedded Microsoft detector drifted from reviewed source' }
if ($UpstreamPath) {
    $Upstream = [Management.Automation.Language.Parser]::ParseFile([IO.Path]::GetFullPath($UpstreamPath), [ref]$Tokens, [ref]$ParseErrors)
    $Original = @($Upstream.FindAll({ param($Node) $Node -is [Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -ceq 'Get-SpeculationControlSettings' }, $true))
    if ($ParseErrors.Count -or $Original.Count -ne 1 -or (Get-FunctionHash $Original[0].Extent.Text) -cne $ExpectedHash) { throw 'Source function mismatch' }
}
$Text = [IO.File]::ReadAllText($Path)
if (-not $Text.Contains('MIT License') -or -not $Text.Contains('Copyright (c) Microsoft Corporation. All rights reserved.') -or -not $Text.Contains('Permission is hereby granted')) { throw 'Upstream attribution/license missing' }
if (-not $Text.Contains('Get-BaselineSpeculationEvidence -Requested $IncludeSpeculationControl.IsPresent') -or -not $Text.Contains('speculationControl = $speculationControl')) { throw 'Opt-in collector wiring missing' }
. ([scriptblock]::Create($Wrapper[0].Extent.Text))
function Assert-Speculation([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
$Result = Get-BaselineSpeculationEvidence -Requested $false -Read { throw 'Must not query' }
Assert-Speculation ($Result.state -eq 'NotRequested') 'Collection not opt-in'
$Result = Get-BaselineSpeculationEvidence -Requested $true -Read { [pscustomobject]@{ BTIWindowsSupportEnabled = $false; SSBDHardwareVulnerable = $null; Unexpected = 'DO_NOT_EXPORT'; GdsStatus = 'SYSTEM_SPECULATION_CONTROL_GDS_MITIGATED' } }
Assert-Speculation ($Result.state -eq 'Complete' -and $Result.settings.BTIWindowsSupportEnabled -eq $false -and $Result.settings.Contains('SSBDHardwareVulnerable')) 'False/null provider values lost'
Assert-Speculation (-not ($Result | ConvertTo-Json -Depth 10).Contains('DO_NOT_EXPORT')) 'Unexpected provider data exported'
$Result = Get-BaselineSpeculationEvidence -Requested $true -Read { @{ BTIWindowsSupportEnabled = 'False'; BTIWindowsSupportPresent = $true } }
Assert-Speculation ($Result.state -eq 'Partial' -and -not $Result.settings.Contains('BTIWindowsSupportEnabled')) 'Non-Boolean value accepted'
$Result = Get-BaselineSpeculationEvidence -Requested $true -Read { throw 'DO_NOT_EXPORT' }
Assert-Speculation ($Result.state -eq 'Error' -and -not ($Result | ConvertTo-Json).Contains('DO_NOT_EXPORT')) 'Provider error hidden or leaked'
foreach ($Read in @({ @() }, { @(@{ BTIWindowsSupportEnabled = $true }, @{ BTIWindowsSupportEnabled = $false }) }, { @{ Unknown = $true } })) {
    $Result = Get-BaselineSpeculationEvidence -Requested $true -Read $Read
    Assert-Speculation ($Result.state -eq 'Error') 'Invalid provider row count or schema accepted'
}
$Synthetic = @{}
foreach ($Field in @('BTIHardwarePresent','BTIWindowsSupportPresent','BTIWindowsSupportEnabled','KVAShadowRequired','KVAShadowWindowsSupportPresent','KVAShadowWindowsSupportEnabled','SSBDWindowsSupportPresent','SSBDHardwareVulnerable','SSBDHardwarePresent','SSBDWindowsSupportEnabledSystemWide','L1TFHardwareVulnerable','L1TFWindowsSupportPresent','L1TFWindowsSupportEnabled','MDSWindowsSupportPresent','MDSHardwareVulnerable','MDSWindowsSupportEnabled','FBClearWindowsSupportPresent','SBDRSSDPHardwareVulnerable','FBSDPHardwareVulnerable','PSDPHardwareVulnerable','FBClearWindowsSupportEnabled')) { $Synthetic[$Field] = $true }
foreach ($Field in @('BTIDisabledBySystemPolicy','BTIDisabledByNoHardwareSupport','KVAShadowPcidEnabled','BTIKernelRetpolineEnabled','BhbEnabled')) { $Synthetic[$Field] = $false }
foreach ($Family in @(@('BranchConfusion','BRANCH_CONFUSION'),@('Gds','GDS'),@('Srso','SRSO'),@('DivideByZero','DIVIDE_BY_ZERO'),@('Rfds','RFDS'))) {
    $Synthetic[$Family[0] + 'Reported'] = $true
    $Synthetic[$Family[0] + 'Status'] = 'SYSTEM_SPECULATION_CONTROL_' + $Family[1] + '_MITIGATED'
}
$Result = Get-BaselineSpeculationEvidence -Requested $true -Read { $Synthetic }
Assert-Speculation ($Result.state -eq 'Complete') 'Synthetic mitigation families did not survive projection'
if ($env:ASSAY_SPECULATION_FIXTURE) {
    [IO.File]::WriteAllText($env:ASSAY_SPECULATION_FIXTURE, (@{systemInfo = @{hostname = 'Synthetic speculation device'; isServer = $false}; speculationControl = $Result} | ConvertTo-Json -Depth 12), [Text.UTF8Encoding]::new($false))
}
Write-Output 'PASS: embedded upstream hash/license, opt-in wiring, typed projection and failure states; no mitigation query executed.'