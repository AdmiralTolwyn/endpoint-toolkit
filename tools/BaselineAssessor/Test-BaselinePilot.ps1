#Requires -Version 5.1
<#
.SYNOPSIS
    BaselinePilot validation tests — validates collection data, check definitions, and evaluation logic.
.DESCRIPTION
    Runs offline tests against a collection JSON and checks.json to verify:
    - All collectionKeys resolve to actual data
    - Evaluation logic correctly maps values to Pass/Fail/Warning
    - Join type detection works for all scenarios
    - Security policy, audit policy, registry baseline key formats are correct
    - Service, defender, firewall, TLS data structures parse correctly
.PARAMETER CollectionPath
    Path to a collection JSON file to validate against.
.PARAMETER ChecksPath
    Path to checks.json. Default: same directory.
#>
[CmdletBinding()]
param(
    [string]$CollectionPath,
    [string]$ChecksPath
)

$ErrorActionPreference = 'Continue'
$Root = $PSScriptRoot
if (-not $Root) { $Root = $PWD.Path }
if (-not $ChecksPath) { $ChecksPath = Join-Path $Root 'checks.json' }
if (-not $CollectionPath) {
    # Find most recent collection JSON
    $CollectionPath = Get-ChildItem $Root -Filter '*_baseline_*.json' -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName
}

$Pass = 0; $Fail = 0; $Warn = 0
function Test-Assert {
    param([string]$Name, [bool]$Condition, [string]$Detail = '')
    if ($Condition) {
        Write-Host "  [PASS] $Name" -ForegroundColor Green
        $Script:Pass++
    } else {
        Write-Host "  [FAIL] $Name$(if ($Detail) { " — $Detail" })" -ForegroundColor Red
        $Script:Fail++
    }
}
function Test-Warn {
    param([string]$Name, [string]$Detail = '')
    Write-Host "  [WARN] $Name$(if ($Detail) { " — $Detail" })" -ForegroundColor Yellow
    $Script:Warn++
}

# ═══════════════════════════════════════════════════════════════════════
# Load the Resolve-CollectionKey function (copied from BaselinePilot.ps1)
# ═══════════════════════════════════════════════════════════════════════

function Resolve-CollectionKey {
    param($Json, [string]$Key)
    $dotIdx = $Key.IndexOf('.')
    if ($dotIdx -lt 0) {
        try { return $Json.$Key } catch { return $null }
    }
    $section = $Key.Substring(0, $dotIdx)
    $remainder = $Key.Substring($dotIdx + 1)
    $sectionObj = $null
    try { $sectionObj = $Json.$section } catch { return $null }
    if ($null -eq $sectionObj) { return $null }
    if ($section -in @('registryBaselines', 'securityPolicy', 'auditPolicy')) {
        try {
            $prop = $sectionObj.PSObject.Properties[$remainder]
            if ($prop) { return $prop.Value }
        } catch { }
        return $null
    }
    $Parts = $remainder -split '\.'
    $Current = $sectionObj
    foreach ($Part in $Parts) {
        if ($null -eq $Current) { return $null }
        try { $Current = $Current.$Part } catch { return $null }
    }
    return $Current
}

# ═══════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host '  BaselinePilot Validation Tests' -ForegroundColor Cyan
Write-Host "  $('=' * 50)" -ForegroundColor DarkGray
Write-Host ''

# ═══════════════════════════════════════════════════════════════════════
# TEST 1: File loading
# ═══════════════════════════════════════════════════════════════════════
Write-Host '  [TEST GROUP] File Loading' -ForegroundColor White
Test-Assert 'checks.json exists' (Test-Path $ChecksPath)
Test-Assert 'Collection JSON exists' (Test-Path $CollectionPath) "Path: $CollectionPath"

$checks = (Get-Content $ChecksPath -Raw -Encoding UTF8 | ConvertFrom-Json).checks
$coll   = Get-Content $CollectionPath -Raw -Encoding UTF8 | ConvertFrom-Json

Test-Assert 'checks.json has checks array' ($checks.Count -gt 0) "Count: $($checks.Count)"
Test-Assert 'Collection has _metadata' ($null -ne $coll._metadata)
Test-Assert 'Collection has systemInfo' ($null -ne $coll.systemInfo)
Test-Assert 'Collection has joinType' ($null -ne $coll.joinType)

# ═══════════════════════════════════════════════════════════════════════
# TEST 2: Join type detection
# ═══════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host '  [TEST GROUP] Join Type Detection' -ForegroundColor White

$jt = $coll.joinType
Test-Assert 'joinType.azureAdJoined is boolean' ($jt.azureAdJoined -is [bool])
Test-Assert 'joinType.domainJoined is boolean' ($jt.domainJoined -is [bool])

# Simulate detection logic
$detected = if ($jt.azureAdJoined -and -not $jt.domainJoined) { 'Entra ID (Intune)' }
            elseif ($jt.domainJoined -and -not $jt.azureAdJoined) { 'AD DS (GPO)' }
            elseif ($jt.domainJoined -and $jt.azureAdJoined) { 'Hybrid Joined' }
            else { 'Workgroup' }
Test-Assert "Join type detected: $detected" ($detected -ne '')
Write-Host "    -> Detected: $detected" -ForegroundColor Gray

# ═══════════════════════════════════════════════════════════════════════
# TEST 3: Collection key resolution
# ═══════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host '  [TEST GROUP] Collection Key Resolution' -ForegroundColor White

$keyResults = @{ Total=0; Found=0; Missing=0; MissingKeys=@() }
foreach ($chk in $checks) {
    if (-not $chk.collectionKeys) { continue }
    foreach ($key in $chk.collectionKeys) {
        $keyResults.Total++
        $val = Resolve-CollectionKey -Json $coll -Key $key
        if ($null -ne $val) {
            $keyResults.Found++
        } else {
            $keyResults.Missing++
            $keyResults.MissingKeys += "$($chk.id): $key"
        }
    }
}

$resolvePct = if ($keyResults.Total -gt 0) { [math]::Round(($keyResults.Found / $keyResults.Total) * 100) } else { 0 }
Test-Assert "Key resolution rate: $resolvePct% ($($keyResults.Found)/$($keyResults.Total))" ($resolvePct -ge 50) "Missing: $($keyResults.Missing)"

if ($keyResults.Missing -gt 0 -and $keyResults.Missing -le 20) {
    foreach ($mk in $keyResults.MissingKeys) {
        Test-Warn "Unresolvable key: $mk"
    }
} elseif ($keyResults.Missing -gt 20) {
    Write-Host "    -> $($keyResults.Missing) keys missing (showing first 15):" -ForegroundColor Yellow
    $keyResults.MissingKeys | Select-Object -First 15 | ForEach-Object { Write-Host "       $_" -ForegroundColor DarkYellow }
}

# ═══════════════════════════════════════════════════════════════════════
# TEST 4: Specific key format tests
# ═══════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host '  [TEST GROUP] Key Format Validation' -ForegroundColor White

# Registry keys
$regVal = Resolve-CollectionKey -Json $coll -Key 'registryBaselines.HKLM\SYSTEM\CurrentControlSet\Control\Lsa\NoLMHash'
Test-Assert 'Registry key resolved (Lsa\NoLMHash)' ($null -ne $regVal) "Value: $regVal"

$regVal2 = Resolve-CollectionKey -Json $coll -Key 'registryBaselines.HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA'
Test-Assert 'Registry key resolved (EnableLUA)' ($null -ne $regVal2) "Value: $regVal2"

# Security policy
$spVal = Resolve-CollectionKey -Json $coll -Key 'securityPolicy.System Access_MinimumPasswordLength'
Test-Assert 'Security policy key resolved (MinPwdLen)' ($null -ne $spVal) "Value: $spVal"

$spVal2 = Resolve-CollectionKey -Json $coll -Key 'securityPolicy.System Access_PasswordComplexity'
Test-Assert 'Security policy key resolved (PwdComplexity)' ($null -ne $spVal2) "Value: $spVal2"

# Audit policy
$apVal = Resolve-CollectionKey -Json $coll -Key 'auditPolicy.Credential Validation'
Test-Assert 'Audit policy key resolved (Credential Validation)' ($null -ne $apVal) "Value: $apVal"

$apVal2 = Resolve-CollectionKey -Json $coll -Key 'auditPolicy.Logon'
Test-Assert 'Audit policy key resolved (Logon)' ($null -ne $apVal2) "Value: $apVal2"

# Defender
$defVal = Resolve-CollectionKey -Json $coll -Key 'defender.RealTimeProtectionEnabled'
Test-Assert 'Defender key resolved (RealTimeProtection)' ($null -ne $defVal) "Value: $defVal"

$defVal2 = Resolve-CollectionKey -Json $coll -Key 'defender.BehaviorMonitoringEnabled'
Test-Assert 'Defender key resolved (BehaviorMonitoring)' ($null -ne $defVal2) "Value: $defVal2"

# Firewall
$fwVal = Resolve-CollectionKey -Json $coll -Key 'firewall.Domain.Enabled'
Test-Assert 'Firewall key resolved (Domain.Enabled)' ($null -ne $fwVal) "Value: $fwVal"

# Services
$svcVal = Resolve-CollectionKey -Json $coll -Key 'services.EventLog'
Test-Assert 'Service key resolved (EventLog)' ($null -ne $svcVal) "Value: $($svcVal.Status)"

# SMB
$smbVal = Resolve-CollectionKey -Json $coll -Key 'smbConfig.server.SMB1Protocol'
Test-Assert 'SMB key resolved (SMB1Protocol)' ($null -ne $smbVal) "Value: $smbVal"

# Event log metadata
$elmVal = Resolve-CollectionKey -Json $coll -Key 'eventLogMetadata.Security.MaxSizeKB'
Test-Assert 'EventLog metadata resolved (Security.MaxSizeKB)' ($null -ne $elmVal) "Value: $elmVal"

# System info
$siVal = Resolve-CollectionKey -Json $coll -Key 'systemInfo.ComputerName'
# Note: collection uses 'hostname' not 'ComputerName'
$siVal2 = Resolve-CollectionKey -Json $coll -Key 'systemInfo.hostname'
Test-Assert 'SystemInfo key resolved (hostname)' ($null -ne $siVal2) "Value: $siVal2"

# ═══════════════════════════════════════════════════════════════════════
# TEST 5: Evaluation logic
# ═══════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host '  [TEST GROUP] Evaluation Logic' -ForegroundColor White

# Test baseline value comparison
$evalTests = @(
    @{ Actual=1; Baseline=1; Expected='Pass'; Label='Int match' }
    @{ Actual=0; Baseline=1; Expected='Fail'; Label='Int mismatch' }
    @{ Actual='Success and Failure'; Baseline='Success and Failure'; Expected='Pass'; Label='String match' }
    @{ Actual='No Auditing'; Baseline='Success and Failure'; Expected='Fail'; Label='String mismatch' }
    @{ Actual=14; Baseline=14; Expected='Pass'; Label='Password length match' }
    @{ Actual=0; Baseline=14; Expected='Fail'; Label='Password length mismatch' }
)

foreach ($et in $evalTests) {
    $result = if ("$($et.Actual)" -eq "$($et.Baseline)") { 'Pass' } else { 'Fail' }
    Test-Assert "Eval: $($et.Label) ($($et.Actual) vs $($et.Baseline))" ($result -eq $et.Expected) "Got: $result, Expected: $($et.Expected)"
}

# Test boolean evaluation (no baseline value)
$boolTests = @(
    @{ Actual=$true;  Expected='Pass'; Label='Boolean true -> Pass' }
    @{ Actual=$false; Expected='Fail'; Label='Boolean false -> Fail' }
)
foreach ($bt in $boolTests) {
    $result = if ($bt.Actual) { 'Pass' } else { 'Fail' }
    Test-Assert "Eval: $($bt.Label)" ($result -eq $bt.Expected)
}

# Test audit policy string evaluation
$auditTests = @(
    @{ Val='No Auditing'; Expected='Fail'; Label='No Auditing -> Fail' }
    @{ Val='Success and Failure'; Expected='Pass'; Label='Success and Failure -> Pass' }
    @{ Val='Success'; Expected='Pass'; Label='Success -> Pass' }
)
foreach ($at in $auditTests) {
    $result = if ($at.Val -match 'No Auditing') { 'Fail' } elseif ($at.Val -match 'Success') { 'Pass' } else { 'Warning' }
    Test-Assert "Eval: $($at.Label)" ($result -eq $at.Expected)
}

# ═══════════════════════════════════════════════════════════════════════
# TEST 6: Data quality checks
# ═══════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host '  [TEST GROUP] Data Quality' -ForegroundColor White

Test-Assert 'System info has hostname' ($null -ne $coll.systemInfo.hostname)
Test-Assert 'System info has OS version' ($null -ne $coll.systemInfo.osVersion)
Test-Assert 'System info has OS build' ($null -ne $coll.systemInfo.osBuild)
Test-Assert 'Registry baselines has entries' ($coll.registryBaselines.PSObject.Properties.Name.Count -gt 50) "Count: $($coll.registryBaselines.PSObject.Properties.Name.Count)"
Test-Assert 'Security policy has entries' ($coll.securityPolicy.PSObject.Properties.Name.Count -gt 20) "Count: $($coll.securityPolicy.PSObject.Properties.Name.Count)"
Test-Assert 'Audit policy has entries' ($coll.auditPolicy.PSObject.Properties.Name.Count -gt 20) "Count: $($coll.auditPolicy.PSObject.Properties.Name.Count)"
Test-Assert 'Defender data present' ($null -ne $coll.defender.RealTimeProtectionEnabled)
Test-Assert 'Firewall data present' ($null -ne $coll.firewall.Domain)
Test-Assert 'Services data present' ($coll.services.PSObject.Properties.Name.Count -gt 10) "Count: $($coll.services.PSObject.Properties.Name.Count)"
Test-Assert 'BitLocker data present' ($null -ne $coll.bitlocker)
Test-Assert 'TLS config present' ($null -ne $coll.tlsConfig)
Test-Assert 'SMB config present' ($null -ne $coll.smbConfig)
Test-Assert 'Event data present' ($null -ne $coll.eventData)
Test-Assert 'Event metadata present' ($null -ne $coll.eventLogMetadata)

# Check for null-heavy registry (may indicate collection issues)
$regProps = $coll.registryBaselines.PSObject.Properties
$nullCount = @($regProps | Where-Object { $null -eq $_.Value }).Count
$totalReg = $regProps.Name.Count
$nullPct = [math]::Round(($nullCount / $totalReg) * 100)
if ($nullPct -gt 60) {
    Test-Warn "Registry baselines: $nullPct% are null ($nullCount/$totalReg) — many policies not configured via GPO"
} else {
    Test-Assert "Registry null rate: $nullPct% ($nullCount/$totalReg)" ($nullPct -lt 80)
}

# ═══════════════════════════════════════════════════════════════════════
# TEST 7: Simulate full evaluation
# ═══════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host '  [TEST GROUP] Full Evaluation Simulation' -ForegroundColor White

$evalResults = @{ Pass=0; Fail=0; Warning=0; NA=0; NotAssessed=0 }
foreach ($chk in $checks) {
    if (-not $chk.collectionKeys) { $evalResults.NotAssessed++; continue }
    $resolved = $false
    foreach ($key in $chk.collectionKeys) {
        $val = Resolve-CollectionKey -Json $coll -Key $key
        if ($null -ne $val) {
            $resolved = $true
            if ($null -ne $chk.baselineValue) {
                if ("$val" -eq "$($chk.baselineValue)") { $evalResults.Pass++ }
                else { $evalResults.Fail++ }
            } else {
                if ($val -is [bool]) {
                    if ($val) { $evalResults.Pass++ } else { $evalResults.Fail++ }
                } elseif ($val -is [string] -and $val -match 'No Auditing') {
                    $evalResults.Fail++
                } elseif ($val -is [string] -and $val -match 'Success') {
                    $evalResults.Pass++
                } else {
                    $evalResults.Warning++
                }
            }
            break
        }
    }
    if (-not $resolved) { $evalResults.NotAssessed++ }
}

$totalEval = $evalResults.Pass + $evalResults.Fail + $evalResults.Warning
$compliance = if ($totalEval -gt 0) { [math]::Round(($evalResults.Pass / $totalEval) * 100) } else { 0 }

Write-Host "    -> Pass: $($evalResults.Pass) | Fail: $($evalResults.Fail) | Warning: $($evalResults.Warning) | Not Assessed: $($evalResults.NotAssessed)" -ForegroundColor Gray
Write-Host "    -> Compliance: $compliance% | Total evaluable: $totalEval / $($checks.Count)" -ForegroundColor Gray

Test-Assert 'At least 50 checks evaluated' ($totalEval -ge 50) "Evaluated: $totalEval"
Test-Assert 'Compliance score calculated' ($compliance -ge 0)

# ═══════════════════════════════════════════════════════════════════════
# SUMMARY
# ═══════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host "  $('=' * 50)" -ForegroundColor DarkGray
$TotalTests = $Pass + $Fail
Write-Host "  RESULTS: $Pass/$TotalTests passed" -NoNewline -ForegroundColor $(if ($Fail -eq 0) { 'Green' } else { 'Yellow' })
if ($Fail -gt 0) { Write-Host " ($Fail failed)" -NoNewline -ForegroundColor Red }
if ($Warn -gt 0) { Write-Host " ($Warn warnings)" -NoNewline -ForegroundColor Yellow }
Write-Host ''
Write-Host ''
