#Requires -Version 5.1
[CmdletBinding()]
param([string]$EvidenceDirectory)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Invoke-IntuneDiscovery.ps1') -LibraryOnly
function Assert-Admx([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function Get-AdmxFact([string]$Path, [string]$Payload) {
    $Facts = @(ConvertTo-IntuneSettingFacts @{ settingDefinitionId = 'synthetic'; simpleSettingValue = @{ value = $Payload } } @(@{ id = 'synthetic'; baseUri = './Device/Vendor/MSFT'; offsetUri = $Path }) 'policy' 'setting')
    return $Facts[0]
}
$IePath = 'Policy/Config/InternetExplorer/DisableInternetExplorerLaunchViaCOM'
$FixtureRows = [Collections.Generic.List[object]]::new()
$FixtureRows.Add((Get-AdmxFact $IePath '<enabled/>'))
Assert-Admx ((Get-AdmxFact $IePath '<enabled/>').resolution -eq 'ResolvedAdmx') 'IE Boolean policy unresolved'
Assert-Admx ((Get-AdmxFact $IePath '<enabled/><data id="unexpected" value="1"/>').resolution -eq 'UnresolvedAdmx') 'IE accepted unreviewed numeric data'
$Startup = '<enabled/><data id="ConfigureNonTPMStartupKeyUsage_Name" value="false"/><data id="ConfigureTPMStartupKeyUsageDropDown_Name" value="0"/><data id="ConfigurePINUsageDropDown_Name" value="0"/><data id="ConfigureTPMPINKeyUsageDropDown_Name" value="0"/><data id="ConfigureTPMUsageDropDown_Name" value="2"/>'
$StartupPath = 'BitLocker/SystemDrivesRequireStartupAuthentication'
$Fact = Get-AdmxFact $StartupPath $Startup
$FixtureRows.Add($Fact)
Assert-Admx ($Fact.resolution -eq 'ResolvedAdmx' -and $Fact.admx.data.ConfigureNonTPMStartupKeyUsage_Name -is [bool] -and $Fact.admx.data.ConfigureNonTPMStartupKeyUsage_Name -eq $false) 'Documented Boolean startup value lost'
foreach ($Payload in @(
    $Startup.Replace('value="false"', 'value="0"'),
    $Startup.Replace('value="2"', 'value="3"'),
    $Startup.Replace('value="2"', 'value="true"'),
    $Startup.Replace('ConfigurePINUsageDropDown_Name','UseTPMPIN'),
    $Startup.Replace('ConfigurePINUsageDropDown_Name','configurePINUsageDropDown_Name'),
    $Startup.Replace('<data id="ConfigurePINUsageDropDown_Name" value="0"/>',''),
    $Startup.Replace('<enabled/>','<enabled/><data id="Extra" value="1"/>')
)) { Assert-Admx ((Get-AdmxFact $StartupPath $Payload).resolution -eq 'UnresolvedAdmx') 'Unreviewed startup data was resolved' }
foreach ($Entry in @(@('OS','BitLocker/SystemDrivesRecoveryOptions'), @('FDV','BitLocker/FixedDrivesRecoveryOptions'))) {
    $Prefix = $Entry[0]
    $Path = $Entry[1]
    $Recovery = '<enabled/>' + '<data id="{0}AllowDRA_Name" value="false"/><data id="{0}RecoveryPasswordUsageDropDown_Name" value="1"/><data id="{0}RecoveryKeyUsageDropDown_Name" value="0"/><data id="{0}HideRecoveryPage_Name" value="true"/><data id="{0}ActiveDirectoryBackup_Name" value="true"/><data id="{0}ActiveDirectoryBackupDropDown_Name" value="2"/><data id="{0}RequireActiveDirectoryBackup_Name" value="false"/>' -f $Prefix
    $Fact = Get-AdmxFact $Path $Recovery
    $FixtureRows.Add($Fact)
    Assert-Admx ($Fact.resolution -eq 'ResolvedAdmx' -and $Fact.admx.data.Count -eq 7 -and $Fact.admx.data[$Prefix + 'AllowDRA_Name'] -eq $false) 'Recovery payload did not retain exact typed fields'
    foreach ($Payload in @($Recovery.Replace('value="2"','value="0"'), $Recovery.Replace('value="true"','value="1"'), $Recovery.Replace($Prefix + 'RecoveryKeyUsageDropDown_Name','Unreviewed'))) {
        Assert-Admx ((Get-AdmxFact $Path $Payload).resolution -eq 'UnresolvedAdmx') 'Unreviewed recovery data was resolved'
    }
    Assert-Admx ((Get-AdmxFact $Path '<enabled/>').resolution -eq 'UnresolvedAdmx') 'Incomplete enabled recovery payload resolved'
    Assert-Admx ((Get-AdmxFact $Path '<disabled/>').resolution -eq 'ResolvedAdmx') 'Disabled recovery payload rejected'
}
if ($EvidenceDirectory) {
    $SourcePath = Join-Path $EvidenceDirectory 'csp-bitlocker-csp.md'
    $Source = Get-Content -LiteralPath $SourcePath -Raw
    $Records = @(foreach ($Path in @($StartupPath, 'BitLocker/SystemDrivesRecoveryOptions', 'BitLocker/FixedDrivesRecoveryOptions')) {
        $Node = $Path.Split('/')[-1]
        $Section = [regex]::Match($Source, '(?ms)^## ' + [regex]::Escape($Node) + '\r?$.*?(?=^## |\z)').Value
        Assert-Admx ($Section.Contains('./Device/Vendor/MSFT/' + $Path)) 'CSP source URI missing'
        $Published = [regex]::Matches($Section, '<data id="([^"<>]+)" value="(xx|yy|zz)"/>')
        $Contract = Get-IntuneAdmxDataContract $Path
        Assert-Admx ($Published.Count -eq $Contract.Count) 'Published ADMX sample changed; review required'
        foreach ($Entry in $Published) {
            $Identifier = $Entry.Groups[1].Value
            $Type = switch ($Entry.Groups[2].Value) { 'xx' { 'Boolean' }; 'yy' { 'Usage' }; 'zz' { 'Backup' } }
            Assert-Admx ($Identifier -cin @($Contract.Keys) -and $Contract[$Identifier] -ceq $Type) 'Collector type/ID drifted from published sample'
        }
        Assert-Admx ($Section.Contains('true = Explicitly allow') -and $Section.Contains('false = Policy not set') -and $Section.Contains('1 = Required') -and $Section.Contains('0 = Disallowed')) 'Published Boolean/usage semantics changed'
        if ($Path -ne $StartupPath) { Assert-Admx ($Section -match '(?m)^- 1 = Store recovery passwords and key packages\.?' -and $Section -match '(?m)^- 2 = Store recovery passwords only\.?') 'Published backup enum changed' }
        @{ Path = $Path; Source = 'https://learn.microsoft.com/en-us/windows/client-management/mdm/bitlocker-csp#' + $Node.ToLowerInvariant(); SourceSha256 = (Get-FileHash -LiteralPath $SourcePath -Algorithm SHA256).Hash; Fields = $Contract }
    })
    [IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'intune-admx-binding-audit.json'), (@{ CheckedAtUtc = [datetime]::UtcNow.ToString('o'); Contracts = $Records; Limit = 'Source sample/type/enum checks, not recommended values or device enforcement.' } | ConvertTo-Json -Depth 10), [Text.UTF8Encoding]::new($false))
}
if ($env:ASSAY_INTUNE_ADMX_FIXTURE) {
    for ($Index = 0; $Index -lt $FixtureRows.Count; $Index++) { $FixtureRows[$Index].id = 'admx-' + $Index }
    $Document = @{
        SchemaVersion = '1.0'; PackId = 'intune'; CollectionId = '11111111-1111-4111-8111-111111111111'
        Collector = @{ Name = 'Invoke-IntuneDiscovery'; Version = '0.5.1' }
        Tenant = @{ Id = '22222222-2222-4222-8222-222222222222'; Cloud = 'Global' }
        StartedAtUtc = '2026-09-18T09:00:00Z'; CompletedAtUtc = '2026-09-18T09:02:00Z'
        Scope = @{ RequestedModules = @('ModernConfigurationPolicies','SecuritySettings'); RequestedPlatforms = @('Windows'); Visibility = 'Unknown' }
        CollectionStatus = @{
            ModernConfigurationPolicies = @{ State = 'Complete'; ApiVersion = 'beta'; RowsRead = 1; PagesRead = 1 }
            SecuritySettings = @{ State = 'Complete'; ApiVersion = 'beta'; RowsRead = $FixtureRows.Count; PagesRead = 1; CompletedParentIds = @('policy') }
        }
        Inventory = @{
            ModernConfigurationPolicies = @(@{ id = 'policy'; name = 'Synthetic'; platforms = 'windows10'; templateReference = @{ templateFamily = 'baseline'; templateDisplayVersion = '25H2' } })
            SecuritySettings = @($FixtureRows.ToArray())
        }
        Observations = @()
        AssessmentRequirements = @{ ScopeConfirmed = $true; ScopeDescription = 'Synthetic ADMX fixture'; Assessor = 'Test'; ConfirmedAtUtc = '2026-09-18T10:00:00Z'; AssessmentAsOfUtc = '2026-09-18T10:00:00Z'; MaxCollectionAgeHours = 24; ConfigurationReview = @{ ReferenceProfile = 'windows-25h2'; PolicyIds = @('policy') } }
    }
    [IO.File]::WriteAllText($env:ASSAY_INTUNE_ADMX_FIXTURE, ($Document | ConvertTo-Json -Depth 20), [Text.UTF8Encoding]::new($false))
}
Write-Output 'PASS: exact IE/BitLocker data IDs, Boolean and numeric enums, required fields and unknown-value rejection; no tenant or device queries.'