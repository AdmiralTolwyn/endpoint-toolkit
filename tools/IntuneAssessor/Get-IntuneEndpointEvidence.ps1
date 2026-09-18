#Requires -Version 5.1
[CmdletBinding()]
param([string]$TenantId, [string]$OutputPath, [switch]$LibraryOnly)

function New-IntuneEndpointEvidence {
    param([string]$SelectedTenant, [string]$DeviceId, [scriptblock]$Read)
    . (Join-Path $PSScriptRoot 'IntuneEndpointTimestamps.ps1')
    $PathStateFields = [ordered]@{
        SharedSignaturesPath = 'SharedSignaturesPathState'
        SignatureDefinitionUpdateFileSharesSources = 'SignatureFileSharesState'
    }
    $Modules = [ordered]@{}
    foreach ($Module in @('DefenderStatus', 'DefenderPreferences', 'FirewallProfiles', 'BitLockerVolumes', 'DeviceGuard')) {
        try {
            $Rows = @(& $Read $Module)
            if ($Module -eq 'DefenderPreferences') {
                foreach ($Row in $Rows) {
                    foreach ($PathField in $PathStateFields.Keys) {
                        $PathValue = $null
                        if ($Row -is [Collections.IDictionary]) {
                            if ($Row.Contains($PathField)) { $PathValue = $Row[$PathField] }
                        } elseif ($null -ne $Row -and $null -ne $Row.PSObject.Properties[$PathField]) {
                            $PathValue = $Row.PSObject.Properties[$PathField].Value
                        }
                        $PathState = 'Unknown'
                        if ($PathValue -is [string]) {
                            if ($PathValue.Length -eq 0) { $PathState = 'Empty' }
                            elseif (-not [string]::IsNullOrWhiteSpace($PathValue)) { $PathState = 'NonEmpty' }
                        }
                        if ($Row -is [Collections.IDictionary]) {
                            $Row.Remove($PathField) | Out-Null
                            $Row[$PathStateFields[$PathField]] = $PathState
                        } elseif ($null -ne $Row) {
                            $Row.PSObject.Properties.Remove($PathField)
                            $Row | Add-Member -NotePropertyName $PathStateFields[$PathField] -NotePropertyValue $PathState -Force
                        }
                    }
                }
            }
            if ($Module -eq 'DefenderStatus') {
                foreach ($Row in $Rows) {
                    if ($Row -is [Collections.IDictionary]) {
                        if ($Row.Contains('AntivirusSignatureLastUpdated')) {
                            $Row['AntivirusSignatureLastUpdated'] = ConvertTo-IntuneSignatureTimestamp $Row['AntivirusSignatureLastUpdated']
                        }
                    } elseif ($null -ne $Row -and $null -ne $Row.PSObject.Properties['AntivirusSignatureLastUpdated']) {
                        $Row.AntivirusSignatureLastUpdated = ConvertTo-IntuneSignatureTimestamp $Row.AntivirusSignatureLastUpdated
                    }
                }
            }
            $Modules[$Module] = @{ State = 'Complete'; Rows = $Rows }
        }
        catch { $Modules[$Module] = @{ State = 'Error'; Rows = @() } }
    }
    return @{ SchemaVersion = '1.0'; TenantId = $SelectedTenant; DeviceId = $DeviceId; CollectedAtUtc = [datetime]::UtcNow.ToString('o'); Modules = $Modules }
}

if ($LibraryOnly) { return }
$ErrorActionPreference = 'Stop'
if (-not $OutputPath -or (Test-Path -LiteralPath $OutputPath)) { throw 'Provide a new output path.' }
$ExpectedTenant = [guid]::Empty
if (-not [guid]::TryParse($TenantId, [ref]$ExpectedTenant) -or $ExpectedTenant -eq [guid]::Empty) { throw 'Provide an explicit tenant GUID.' }
$JoinStatus = (& dsregcmd.exe /status | Out-String)
if ($LASTEXITCODE -ne 0) { throw 'Device registration identity query failed.' }
$Tenants = @([regex]::Matches($JoinStatus, '(?m)^\s*TenantId\s*:\s*([0-9a-fA-F-]{36})\s*$'))
$Devices = @([regex]::Matches($JoinStatus, '(?m)^\s*DeviceId\s*:\s*([0-9a-fA-F-]{36})\s*$'))
if ($Tenants.Count -ne 1 -or $Devices.Count -ne 1 -or $Tenants[0].Groups[1].Value -ne $TenantId) { throw 'A unique device identity in the selected Entra tenant could not be established.' }
$Evidence = New-IntuneEndpointEvidence -SelectedTenant $TenantId -DeviceId $Devices[0].Groups[1].Value -Read {
    param($Module)
    switch ($Module) {
        'DefenderStatus' { Get-MpComputerStatus -ErrorAction Stop | Select-Object AMRunningMode, AMProductVersion, AMEngineVersion, AntivirusEnabled, RealTimeProtectionEnabled, BehaviorMonitorEnabled, AntivirusSignatureLastUpdated, IsTamperProtected, ControlledConfigurationState, TamperProtectionSource }
        'DefenderPreferences' { Get-MpPreference -ErrorAction Stop | Select-Object AttackSurfaceReductionRules_Ids, AttackSurfaceReductionRules_Actions, EnableNetworkProtection, PUAProtection, DisableRealtimeMonitoring, DisableBehaviorMonitoring, DisableScriptScanning, MAPSReporting, EnableControlledFolderAccess, SignatureFallbackOrder, SignatureScheduleDay, SignatureUpdateInterval, SharedSignaturesPath, SignatureDefinitionUpdateFileSharesSources }
        'FirewallProfiles' { Get-NetFirewallProfile -PolicyStore ActiveStore -ErrorAction Stop | Select-Object Name, Enabled, DefaultInboundAction, DefaultOutboundAction, LogAllowed, LogBlocked }
        'BitLockerVolumes' { Get-BitLockerVolume -ErrorAction Stop | Select-Object MountPoint, VolumeType, VolumeStatus, ProtectionStatus, EncryptionPercentage }
        'DeviceGuard' { Get-CimInstance -ClassName Win32_DeviceGuard -Namespace root\Microsoft\Windows\DeviceGuard -ErrorAction Stop | Select-Object VirtualizationBasedSecurityStatus, SecurityServicesConfigured, SecurityServicesRunning }
    }
}
$Json = $Evidence | ConvertTo-Json -Depth 15
$Bytes = [Text.UTF8Encoding]::new($false).GetBytes($Json)
$Stream = [IO.File]::Open([IO.Path]::GetFullPath($OutputPath), [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
try { $Stream.Write($Bytes, 0, $Bytes.Length) } finally { $Stream.Dispose() }
Write-Output "Exported read-only endpoint evidence to $OutputPath. Module errors remain explicit; no elevation or remediation was attempted."