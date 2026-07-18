#Requires -Version 5.1
<#
.SYNOPSIS
    BaselinePilot validation tests — validates collection data, check definitions, and evaluation logic.
.DESCRIPTION
    Runs offline, side-effect-free tests. The evaluation-engine functions under test
    (Resolve-CollectionKey, Get-ValueEvaluation, Get-CheckEvaluation, Get-WorstMachineStatus,
    Compare-PrivilegeRights, Get-SectionFailure, ...) are extracted from the CURRENT
    BaselinePilot.ps1 via the PowerShell AST — not re-implemented — so they can never
    drift from production.

    Test groups:
    - Engine function extraction sanity
    - Synthetic engine tests (C-1 min/max, C-2 privilege rights, A-3 section failure,
      A-4 secure defaults, A-5 multi-key, C-3 summary-only, C-4 exact audit match)
    - Multi-machine worst-across-machines aggregation (A-6) with an inline synthetic store
    - Optional: collection-JSON key resolution + data quality (skipped when no sample
      collection file is present)
.PARAMETER CollectionPath
    Path to a collection JSON file to validate against (optional).
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

Write-Host ''
Write-Host '  BaselinePilot Validation Tests' -ForegroundColor Cyan
Write-Host "  $('=' * 50)" -ForegroundColor DarkGray
Write-Host ''

# ═══════════════════════════════════════════════════════════════════════
# TEST 0: Extract the CURRENT production engine functions via AST
# ═══════════════════════════════════════════════════════════════════════
# The engine functions are pulled from BaselinePilot.ps1 by parsing it and
# executing only the wanted FunctionDefinitionAst bodies (plus the
# $Global:SecureDefaultTable assignment). Nothing else in the app runs, so
# this stays side-effect-free (no WPF, no file writes).
Write-Host '  [TEST GROUP] Production Engine Extraction' -ForegroundColor White

$AppPath = Join-Path $Root 'BaselinePilot.ps1'
Test-Assert 'BaselinePilot.ps1 exists' (Test-Path $AppPath)

$WantedFunctions = @(
    'Resolve-CollectionKey'
    'Get-SectionFailure'
    'Compare-PrivilegeRights'
    'Test-IsPrivilegeRightsCheck'
    'Get-MdmPolicyValue'
    'Get-ValueEvaluation'
    'Get-SpecialCheckEvaluation'
    'Get-CheckEvaluation'
    'Get-WorstMachineStatus'
)

$extracted = @()
try {
    $tokens = $null; $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($AppPath, [ref]$tokens, [ref]$parseErrors)
    Test-Assert 'BaselinePilot.ps1 parses with 0 errors' (@($parseErrors).Count -eq 0) "$(@($parseErrors).Count) parse errors"

    $funcAsts = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
    foreach ($fname in $WantedFunctions) {
        $fAst = $funcAsts | Where-Object { $_.Name -eq $fname } | Select-Object -First 1
        if ($fAst) {
            . ([ScriptBlock]::Create($fAst.Extent.Text))
            $extracted += $fname
        }
    }

    # Extract the $Global:SecureDefaultTable data table (A-4)
    $sdAssign = $ast.FindAll({ param($n)
        $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
        $n.Left.Extent.Text -match 'SecureDefaultTable'
    }, $true) | Select-Object -First 1
    if ($sdAssign) { . ([ScriptBlock]::Create($sdAssign.Extent.Text)) }
} catch {
    Test-Warn "Extraction error: $($_.Exception.Message)"
}

Test-Assert "All engine functions extracted from production ($($extracted.Count)/$($WantedFunctions.Count))" `
    ($extracted.Count -eq $WantedFunctions.Count) "Missing: $(@($WantedFunctions | Where-Object { $_ -notin $extracted }) -join ', ')"
Test-Assert 'SecureDefaultTable extracted' ($null -ne $Global:SecureDefaultTable -and $Global:SecureDefaultTable.Count -ge 5)

# Minimal check-object factory for synthetic tests (mirrors the fields the engine reads)
function New-TestCheck {
    param([hashtable]$Props)
    $base = @{
        Id = 'TST-001'; Name = 'Test'; Type = 'Auto'
        CollectionKeys = @(); BaselineValue = $null; comparison = $null
        CspPath = $null; threshold = $null; eventIds = $null
        filterField = $null; filterValues = $null; applicableTo = $null
    }
    foreach ($k in $Props.Keys) { $base[$k] = $Props[$k] }
    [PSCustomObject]$base
}
function ConvertTo-JsonObject { param($Hash) $Hash | ConvertTo-Json -Depth 10 | ConvertFrom-Json }

# ═══════════════════════════════════════════════════════════════════════
# TEST 1: checks.json loading
# ═══════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host '  [TEST GROUP] Catalog Loading' -ForegroundColor White
Test-Assert 'checks.json exists' (Test-Path $ChecksPath)
$checksFile = Get-Content $ChecksPath -Raw -Encoding UTF8 | ConvertFrom-Json
$checks = $checksFile.checks
Test-Assert 'checks.json has checks array' (@($checks).Count -gt 0) "Count: $(@($checks).Count)"
Test-Assert 'checks.json has _metadata.version (CatalogVersion stamping source)' ($null -ne $checksFile._metadata.version)
$cmpChecks = @($checks | Where-Object { $_.comparison -in @('min','max') })
Test-Assert "Catalog has min/max comparison checks (C-1)" ($cmpChecks.Count -ge 5) "Found: $($cmpChecks.Count)"

# ═══════════════════════════════════════════════════════════════════════
# TEST 2: Synthetic engine tests — C-1 operator semantics
# ═══════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host '  [TEST GROUP] C-1 Numeric Operator Semantics' -ForegroundColor White

$minChk = New-TestCheck @{ Id='AUTH-001'; BaselineValue=14; comparison='min' }
Test-Assert 'min: 15 >= 14 -> Pass (stricter than baseline)' ((Get-ValueEvaluation -Chk $minChk -ActualValue 15).Status -eq 'Pass')
Test-Assert 'min: "15" (string) >= 14 -> Pass (type coercion)' ((Get-ValueEvaluation -Chk $minChk -ActualValue '15').Status -eq 'Pass')
Test-Assert 'min: 14 >= 14 -> Pass' ((Get-ValueEvaluation -Chk $minChk -ActualValue 14).Status -eq 'Pass')
Test-Assert 'min: 8 < 14 -> Fail' ((Get-ValueEvaluation -Chk $minChk -ActualValue 8).Status -eq 'Fail')

$maxChk = New-TestCheck @{ Id='AUTH-003'; BaselineValue=10; comparison='max' }
Test-Assert 'max: 5 <= 10 -> Pass (stricter lockout)' ((Get-ValueEvaluation -Chk $maxChk -ActualValue 5).Status -eq 'Pass')
Test-Assert 'max: 20 > 10 -> Fail' ((Get-ValueEvaluation -Chk $maxChk -ActualValue 20).Status -eq 'Fail')

$exactChk = New-TestCheck @{ BaselineValue=1 }
Test-Assert 'exact (default): 1 -eq 1 -> Pass' ((Get-ValueEvaluation -Chk $exactChk -ActualValue 1).Status -eq 'Pass')
Test-Assert 'exact (default): 0 -ne 1 -> Fail' ((Get-ValueEvaluation -Chk $exactChk -ActualValue 0).Status -eq 'Fail')

# ═══════════════════════════════════════════════════════════════════════
# TEST 3: Audit-policy values (A-1 plain strings + back-compat, C-4 exact match)
# ═══════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host '  [TEST GROUP] Audit Policy Evaluation' -ForegroundColor White

$audChk = New-TestCheck @{ BaselineValue='Success and Failure' }
Test-Assert 'Plain string "Success and Failure" -> Pass' ((Get-ValueEvaluation -Chk $audChk -ActualValue 'Success and Failure').Status -eq 'Pass')
Test-Assert 'Plain string "Success" vs "Success and Failure" -> Fail' ((Get-ValueEvaluation -Chk $audChk -ActualValue 'Success').Status -eq 'Fail')
$legacyObj = ConvertTo-JsonObject @{ name = 'Logon'; setting = 'Success and Failure' }
Test-Assert 'Legacy {name,setting} object unwrapped -> Pass (back-compat)' ((Get-ValueEvaluation -Chk $audChk -ActualValue $legacyObj).Status -eq 'Pass')

$noBaseChk = New-TestCheck @{}
Test-Assert 'No-baseline heuristic: "No Auditing" -> Fail (exact)' ((Get-ValueEvaluation -Chk $noBaseChk -ActualValue 'No Auditing').Status -eq 'Fail')
Test-Assert 'No-baseline heuristic: "Success" -> Pass (exact whitelist)' ((Get-ValueEvaluation -Chk $noBaseChk -ActualValue 'Success').Status -eq 'Pass')
$subst = Get-ValueEvaluation -Chk $noBaseChk -ActualValue 'Successfully disabled'
Test-Assert 'Substring foot-gun removed: "Successfully disabled" is NOT auto-Pass' ($subst.Status -ne 'Pass')

# ═══════════════════════════════════════════════════════════════════════
# TEST 4: A-3 collection failure -> Not Assessed (never Fail)
# ═══════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host '  [TEST GROUP] A-3 Section Failure Handling' -ForegroundColor White

$failJson = ConvertTo-JsonObject @{
    _metadata  = @{ errors = @(@{ area = 'Defender Status'; error = 'Get-MpPreference unavailable'; sections = @('defender') }) }
    systemInfo = @{ osBuild = '22631' }
    defender   = @{ _collectionFailed = $true; _error = 'Get-MpPreference unavailable' }
}
Test-Assert 'Get-SectionFailure detects _collectionFailed marker' ((Get-SectionFailure -Json $failJson -Section 'defender') -match 'Get-MpPreference')

$nullSectionJson = ConvertTo-JsonObject @{
    _metadata  = @{ errors = @(@{ area = 'Defender Status'; error = 'boom'; sections = @('defender') }) }
    systemInfo = @{ osBuild = '22631' }
}
Test-Assert 'Get-SectionFailure maps null section via _metadata.errors' ((Get-SectionFailure -Json $nullSectionJson -Section 'defender') -eq 'boom')

$defChk = New-TestCheck @{ Id='DEF-001'; CollectionKeys=@('defender.RealTimeProtectionEnabled'); BaselineValue=$true }
$r = Get-CheckEvaluation -Chk $defChk -Json $failJson -OsBuild 22631
Test-Assert 'Failed section -> Not Assessed with reason (not Fail)' ($r.Status -eq 'Not Assessed' -and $r.Details -match 'Collection error') "Got: $($r.Status) / $($r.Details)"

# ═══════════════════════════════════════════════════════════════════════
# TEST 5: A-4 OS-build-aware secure defaults
# ═══════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host '  [TEST GROUP] A-4 Secure-Default Table' -ForegroundColor White

$emptyJson = ConvertTo-JsonObject @{ _metadata = @{ errors = @() }; systemInfo = @{ osBuild = '26100' }; registryBaselines = @{ placeholder = 1 } }
$smbChk = New-TestCheck @{ Id='NET-001'; CollectionKeys=@('registryBaselines.HKLM\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters\RequireSecuritySignature'); BaselineValue=1 }
$r = Get-CheckEvaluation -Chk $smbChk -Json $emptyJson -OsBuild 26100
Test-Assert 'NET-001 absent on 24H2 (26100) -> Pass via secure default' ($r.Status -eq 'Pass' -and $r.Details -match 'secure by default') "Got: $($r.Status) / $($r.Details)"
$r = Get-CheckEvaluation -Chk $smbChk -Json $emptyJson -OsBuild 22631
Test-Assert 'NET-001 absent on 23H2 (22631) -> Fail (not-configured rule kept)' ($r.Status -eq 'Fail') "Got: $($r.Status)"

# ═══════════════════════════════════════════════════════════════════════
# TEST 6: A-5 multi-key all-must-pass
# ═══════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host '  [TEST GROUP] A-5 Multi-Key Evaluation' -ForegroundColor White

$mkJson = ConvertTo-JsonObject @{
    _metadata = @{ errors = @() }; systemInfo = @{ osBuild = '22631' }
    auditPolicy = @{ 'Logon' = 'Success and Failure'; 'Logoff' = 'No Auditing' }
}
$mkChk = New-TestCheck @{ Id='MON-003'; CollectionKeys=@('auditPolicy.Logon','auditPolicy.Logoff'); BaselineValue='Success and Failure' }
$r = Get-CheckEvaluation -Chk $mkChk -Json $mkJson -OsBuild 22631
Test-Assert 'Logon=SF + Logoff=NoAudit -> overall Fail (not first-key Pass)' ($r.Status -eq 'Fail') "Got: $($r.Status)"
Test-Assert 'Multi-key failure names the failing key' ($r.Details -match 'Logoff') "Details: $($r.Details)"

$mkJson2 = ConvertTo-JsonObject @{
    _metadata = @{ errors = @() }; systemInfo = @{ osBuild = '22631' }
    auditPolicy = @{ 'Logon' = 'Success and Failure'; 'Logoff' = 'Success and Failure' }
}
$r = Get-CheckEvaluation -Chk $mkChk -Json $mkJson2 -OsBuild 22631
Test-Assert 'Both keys pass -> overall Pass' ($r.Status -eq 'Pass') "Got: $($r.Status)"

$mkChk3 = New-TestCheck @{ Id='MON-003'; CollectionKeys=@('auditPolicy.Logon','auditPolicy.DoesNotExist'); BaselineValue='Success and Failure' }
$r = Get-CheckEvaluation -Chk $mkChk3 -Json $mkJson2 -OsBuild 22631
Test-Assert 'Dangling key among several -> noted, does not fail the check' ($r.Status -eq 'Pass' -and $r.Details -match 'not found') "Got: $($r.Status) / $($r.Details)"

# ═══════════════════════════════════════════════════════════════════════
# TEST 7: C-2 privilege rights
# ═══════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host '  [TEST GROUP] C-2 Privilege Rights' -ForegroundColor White

$r = Compare-PrivilegeRights -Expected '' -Actual $null
Test-Assert 'Empty baseline + absent secedit line -> Pass (SEC-077 case)' ($r.Status -eq 'Pass')
$r = Compare-PrivilegeRights -Expected '*S-1-5-32-544,*S-1-5-32-545' -Actual '*S-1-5-32-545,*S-1-5-32-544'
Test-Assert 'Order-insensitive SID set match -> Pass' ($r.Status -eq 'Pass')
$r = Compare-PrivilegeRights -Expected '*S-1-5-32-544' -Actual '*S-1-5-32-544,*S-1-5-32-546'
Test-Assert 'Actual beyond expected -> Fail listing offender' ($r.Status -eq 'Fail' -and $r.Details -match 'S-1-5-32-546')
$r = Compare-PrivilegeRights -Expected '*S-1-5-32-544,*S-1-5-32-545' -Actual '*S-1-5-32-544'
Test-Assert 'Actual subset of expected -> Pass' ($r.Status -eq 'Pass')

$privJson = ConvertTo-JsonObject @{
    _metadata = @{ errors = @() }; systemInfo = @{ osBuild = '22631' }
    securityPolicy = @{ 'Privilege Rights_SeCreateTokenPrivilege' = $null }
}
$privChk = New-TestCheck @{ Id='SEC-077'; CollectionKeys=@('securityPolicy.Privilege Rights_SeCreateTokenPrivilege'); BaselineValue='' }
$r = Get-CheckEvaluation -Chk $privChk -Json $privJson -OsBuild 22631
Test-Assert 'SEC-077 (expect empty, secedit omits line) CAN Pass' ($r.Status -eq 'Pass') "Got: $($r.Status) / $($r.Details)"

# ═══════════════════════════════════════════════════════════════════════
# TEST 8: C-3 event thresholds (perAccount, summary-only)
# ═══════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host '  [TEST GROUP] C-3 Event Thresholds' -ForegroundColor White

# perAccount: 6 lockouts for one account (limit 5 per account) -> Fail
$now = (Get-Date).ToString('o')
$events = @()
1..6 | ForEach-Object { $events += @{ id = 4740; TargetUserName = 'alice'; TimeCreated = $now } }
1..2 | ForEach-Object { $events += @{ id = 4740; TargetUserName = 'bob';   TimeCreated = $now } }
$paChk = New-TestCheck @{ Id='AUTH-025'; eventIds=@(4740); threshold=(ConvertTo-JsonObject @{ count=5; windowDays=7; perAccount=$true; operator='gt'; evalMode='count' }) }
$evArr = @($events | ForEach-Object { ConvertTo-JsonObject $_ })
$r = Get-ValueEvaluation -Chk $paChk -ActualValue $evArr
Test-Assert 'perAccount: alice=6 > 5 -> Fail naming alice' ($r.Status -eq 'Fail' -and $r.Details -match 'alice') "Got: $($r.Status) / $($r.Details)"

$evArrOk = @($events | Select-Object -Skip 2 | ForEach-Object { ConvertTo-JsonObject $_ })  # alice=4, bob=2
$r = Get-ValueEvaluation -Chk $paChk -ActualValue $evArrOk
Test-Assert 'perAccount: all accounts within limit -> not Fail' ($r.Status -ne 'Fail') "Got: $($r.Status)"

# windowDays: old events filtered out
$oldEvents = @(1..10 | ForEach-Object { ConvertTo-JsonObject @{ id = 4740; TargetUserName = 'carol'; TimeCreated = (Get-Date).AddDays(-30).ToString('o') } })
$r = Get-ValueEvaluation -Chk $paChk -ActualValue $oldEvents
Test-Assert 'windowDays: 30-day-old events outside 7-day window -> Pass' ($r.Status -eq 'Pass') "Got: $($r.Status)"

# Summary-only: event-filter checks must be Not Assessed
$sumChk = New-TestCheck @{ Id='DEF-028'; eventIds=@(4688); threshold=(ConvertTo-JsonObject @{ windowDays=30; count=5; operator='gt'; evalMode='count' }) }
$summary = ConvertTo-JsonObject @{ count = 9500; firstEvent = '2026-06-01T00:00:00'; lastEvent = '2026-07-01T00:00:00' }
$r = Get-ValueEvaluation -Chk $sumChk -ActualValue $summary
Test-Assert 'EventSummaryOnly + eventIds check -> Not Assessed (never raw-total Fail)' ($r.Status -eq 'Not Assessed' -and $r.Details -match 'full event collection') "Got: $($r.Status)"

# ═══════════════════════════════════════════════════════════════════════
# TEST 9: C-4 concrete evaluators
# ═══════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host '  [TEST GROUP] C-4 Concrete Evaluators' -ForegroundColor White

$renJson = ConvertTo-JsonObject @{
    _metadata = @{ errors = @() }; systemInfo = @{ osBuild = '22631' }
    securityPolicy = @{ 'System Access_NewAdministratorName' = '"LocalRoot"'; 'System Access_NewGuestName' = 'Guest' }
}
$adminChk = New-TestCheck @{ Id='SEC-041'; CollectionKeys=@('securityPolicy.System Access_NewAdministratorName') }
$guestChk = New-TestCheck @{ Id='SEC-042'; CollectionKeys=@('securityPolicy.System Access_NewGuestName') }
$r = Get-CheckEvaluation -Chk $adminChk -Json $renJson -OsBuild 22631
Test-Assert 'SEC-041: renamed Administrator -> Pass' ($r.Status -eq 'Pass') "Got: $($r.Status)"
$r = Get-CheckEvaluation -Chk $guestChk -Json $renJson -OsBuild 22631
Test-Assert 'SEC-042: default "Guest" -> Fail' ($r.Status -eq 'Fail') "Got: $($r.Status)"

$tlsJson = ConvertTo-JsonObject @{
    _metadata = @{ errors = @() }; systemInfo = @{ osBuild = '22631' }
    tlsConfig = @{ 'TLS 1.0 Server Enabled' = 1; 'TLS 1.1 Server Enabled' = 0; 'SSL 3.0 Server Enabled' = 0 }
}
$tlsChk = New-TestCheck @{ Id='NET-028'; CollectionKeys=@('tlsConfig') }
$r = Get-CheckEvaluation -Chk $tlsChk -Json $tlsJson -OsBuild 22631
Test-Assert 'NET-028: TLS 1.0 enabled -> Fail listing offender' ($r.Status -eq 'Fail' -and $r.Details -match 'TLS 1\.0') "Got: $($r.Status) / $($r.Details)"

$tlsJsonOk = ConvertTo-JsonObject @{
    _metadata = @{ errors = @() }; systemInfo = @{ osBuild = '22631' }
    tlsConfig = @{ 'TLS 1.0 Server Enabled' = 0; 'TLS 1.1 Server Enabled' = 0; 'SSL 3.0 Server Enabled' = 0; 'SSL 2.0 Server Enabled' = 0 }
}
$r = Get-CheckEvaluation -Chk $tlsChk -Json $tlsJsonOk -OsBuild 22631
Test-Assert 'NET-028: all legacy protocols disabled -> Pass' ($r.Status -eq 'Pass') "Got: $($r.Status)"

$apHash = @{}
1..12 | ForEach-Object { $apHash["Subcat$_"] = 'Success and Failure' }
1..5  | ForEach-Object { $apHash["Quiet$_"] = 'No Auditing' }
$monJson = ConvertTo-JsonObject @{ _metadata = @{ errors = @() }; systemInfo = @{ osBuild = '22631' }; auditPolicy = $apHash }
$monChk = New-TestCheck @{ Id='MON-030'; CollectionKeys=@('auditPolicy','eventData') }
$r = Get-CheckEvaluation -Chk $monChk -Json $monJson -OsBuild 22631
Test-Assert 'MON-030: 12 configured subcategories >= 10 -> Pass (no permanent Warning)' ($r.Status -eq 'Pass') "Got: $($r.Status) / $($r.Details)"

$opsChk = New-TestCheck @{ Id='OPS-011'; CollectionKeys=@('systemInfo') }
$r = Get-CheckEvaluation -Chk $opsChk -Json $monJson -OsBuild 22631
Test-Assert 'OPS-011: no data collected -> Not Assessed (manual verification)' ($r.Status -eq 'Not Assessed') "Got: $($r.Status)"

# ═══════════════════════════════════════════════════════════════════════
# TEST 10: C-5 MDM PolicyManager values
# ═══════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host '  [TEST GROUP] C-5 MDM Policy Values' -ForegroundColor White

$mdmJson = ConvertTo-JsonObject @{
    _metadata = @{ errors = @() }; systemInfo = @{ osBuild = '22631' }
    mdmEnrollment = @{
        mdmEnrolled = $true
        managedAreas = @{ Defender = @{ settingCount = 12; settings = @('AllowRealtimeMonitoring') } }
        policyValues = @{ Defender = @{ AllowRealtimeMonitoring = 1 } }
    }
}
$mdmChk = New-TestCheck @{ Id='DEF-001'; CollectionKeys=@('registryBaselines.HKLM\Missing\Path'); BaselineValue=1; CspPath='Defender/AllowRealtimeMonitoring' }
$r = Get-CheckEvaluation -Chk $mdmChk -Json $mdmJson -OsBuild 22631
Test-Assert 'MDM policyValues present -> concrete Pass (not verify-in-portal)' ($r.Status -eq 'Pass' -and $r.Details -match 'PolicyManager') "Got: $($r.Status) / $($r.Details)"

$mdmChkNoVal = New-TestCheck @{ Id='DEF-002'; CollectionKeys=@('registryBaselines.HKLM\Missing\Path2'); CspPath='Defender/SomethingUncollected' }
$r = Get-CheckEvaluation -Chk $mdmChkNoVal -Json $mdmJson -OsBuild 22631
Test-Assert 'MDM-managed area, no value -> Warning verify-in-portal fallback' ($r.Status -eq 'Warning' -and $r.Details -match 'verify') "Got: $($r.Status) / $($r.Details)"

# ═══════════════════════════════════════════════════════════════════════
# TEST 11: A-6 worst-across-machines aggregation (synthetic store)
# ═══════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host '  [TEST GROUP] A-6 Multi-Machine Aggregation' -ForegroundColor White

$results = @(
    @{ Hostname='PC-A'; Status='Pass';    ActualValue='1'; Details='' }
    @{ Hostname='PC-B'; Status='Fail';    ActualValue='0'; Details='Not configured' }
    @{ Hostname='PC-C'; Status='Warning'; ActualValue='?'; Details='review' }
)
$worst = Get-WorstMachineStatus -Results $results
Test-Assert 'Worst of Pass/Fail/Warning = Fail (from the right host)' ($worst.Status -eq 'Fail' -and $worst.Hostname -eq 'PC-B')

$results2 = @(
    @{ Hostname='PC-A'; Status='Pass';         ActualValue='1'; Details='' }
    @{ Hostname='PC-B'; Status='Not Assessed'; ActualValue=$null; Details='' }
)
$worst = Get-WorstMachineStatus -Results $results2
Test-Assert 'Not Assessed ignored when any machine assessed' ($worst.Status -eq 'Pass' -and $worst.Hostname -eq 'PC-A')

$results3 = @(
    @{ Hostname='PC-A'; Status='Warning'; ActualValue='x'; Details='' }
    @{ Hostname='PC-B'; Status='Pass';    ActualValue='1'; Details='' }
)
$worst = Get-WorstMachineStatus -Results $results3
Test-Assert 'Warning outranks Pass' ($worst.Status -eq 'Warning' -and $worst.Hostname -eq 'PC-A')

$results4 = @(
    @{ Hostname='PC-A'; Status='N/A'; ActualValue='N/A'; Details='' }
    @{ Hostname='PC-B'; Status='N/A'; ActualValue='N/A'; Details='' }
)
$worst = Get-WorstMachineStatus -Results $results4
Test-Assert 'All machines N/A -> N/A' ($worst.Status -eq 'N/A')

# End-to-end: same check evaluated against two synthetic collections, store simulated
$collA = ConvertTo-JsonObject @{
    _metadata = @{ errors = @() }; systemInfo = @{ hostname = 'PC-A'; osBuild = '22631' }
    auditPolicy = @{ 'Logon' = 'Success and Failure' }
}
$collB = ConvertTo-JsonObject @{
    _metadata = @{ errors = @() }; systemInfo = @{ hostname = 'PC-B'; osBuild = '22631' }
    auditPolicy = @{ 'Logon' = 'No Auditing' }
}
$aggChk = New-TestCheck @{ Id='MON-001'; CollectionKeys=@('auditPolicy.Logon'); BaselineValue='Success and Failure' }
$store = @{}
foreach ($c in @($collA, $collB)) {
    $res = Get-CheckEvaluation -Chk $aggChk -Json $c -OsBuild 22631
    $store[$c.systemInfo.hostname] = @{ $aggChk.Id = @{ Status=$res.Status; ActualValue="$($res.ActualValue)"; Details="$($res.Details)" } }
}
$aggResults = @()
foreach ($h in $store.Keys) { $e = $store[$h][$aggChk.Id]; $aggResults += , @{ Hostname=$h; Status=$e.Status; ActualValue=$e.ActualValue; Details=$e.Details } }
$worst = Get-WorstMachineStatus -Results $aggResults
Test-Assert 'E2E: PC-A Pass + PC-B Fail -> displayed status Fail from PC-B' ($worst.Status -eq 'Fail' -and $worst.Hostname -eq 'PC-B') "Got: $($worst.Status) @ $($worst.Hostname)"
$affected = @($aggResults | Where-Object { $_.Status -in @('Fail','Warning') } | ForEach-Object { $_.Hostname })
Test-Assert 'E2E: AffectedMachines rebuilt = only PC-B' ($affected.Count -eq 1 -and $affected[0] -eq 'PC-B')

# Re-import replacement: PC-B re-imported clean replaces its entries
$store['PC-B'] = @{ $aggChk.Id = @{ Status='Pass'; ActualValue='Success and Failure'; Details='' } }
$aggResults = @()
foreach ($h in $store.Keys) { $e = $store[$h][$aggChk.Id]; $aggResults += , @{ Hostname=$h; Status=$e.Status; ActualValue=$e.ActualValue; Details=$e.Details } }
$worst = Get-WorstMachineStatus -Results $aggResults
$affected = @($aggResults | Where-Object { $_.Status -in @('Fail','Warning') })
Test-Assert 'E2E: clean re-import of PC-B -> Pass, AffectedMachines pruned' ($worst.Status -eq 'Pass' -and $affected.Count -eq 0)

# ═══════════════════════════════════════════════════════════════════════
# TEST 12 (optional): Collection JSON validation — skipped without sample data
# ═══════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host '  [TEST GROUP] Collection JSON Validation' -ForegroundColor White

if ($CollectionPath -and (Test-Path $CollectionPath)) {
    $coll = Get-Content $CollectionPath -Raw -Encoding UTF8 | ConvertFrom-Json
    Test-Assert 'Collection has _metadata' ($null -ne $coll._metadata)
    Test-Assert 'Collection has systemInfo' ($null -ne $coll.systemInfo)
    Test-Assert 'Collection has joinType' ($null -ne $coll.joinType)

    # Key resolution against the real collection (uses production Resolve-CollectionKey)
    $keyResults = @{ Total=0; Found=0; Missing=0; MissingKeys=@() }
    foreach ($chk in $checks) {
        if (-not $chk.collectionKeys) { continue }
        foreach ($key in $chk.collectionKeys) {
            $keyResults.Total++
            $val = Resolve-CollectionKey -Json $coll -Key $key
            if ($null -ne $val) { $keyResults.Found++ }
            else {
                $keyResults.Missing++
                $keyResults.MissingKeys += "$($chk.id): $key"
            }
        }
    }
    $resolvePct = if ($keyResults.Total -gt 0) { [math]::Round(($keyResults.Found / $keyResults.Total) * 100) } else { 0 }
    Test-Assert "Key resolution rate: $resolvePct% ($($keyResults.Found)/$($keyResults.Total))" ($resolvePct -ge 50) "Missing: $($keyResults.Missing)"
    if ($keyResults.Missing -gt 0) {
        Write-Host "    -> $($keyResults.Missing) unresolved keys (first 15):" -ForegroundColor Yellow
        $keyResults.MissingKeys | Select-Object -First 15 | ForEach-Object { Write-Host "       $_" -ForegroundColor DarkYellow }
    }

    # Data quality (guarded against div-by-zero when sections are empty/missing)
    $regProps = @($coll.registryBaselines.PSObject.Properties)
    $totalReg = $regProps.Count
    if ($totalReg -gt 0) {
        $nullCount = @($regProps | Where-Object { $null -eq $_.Value }).Count
        $nullPct = [math]::Round(($nullCount / $totalReg) * 100)
        if ($nullPct -gt 60) {
            Test-Warn "Registry baselines: $nullPct% are null ($nullCount/$totalReg) — many policies not configured via GPO"
        } else {
            Test-Assert "Registry null rate: $nullPct% ($nullCount/$totalReg)" ($nullPct -lt 80)
        }
    } else {
        Test-Warn 'Registry baselines section empty or missing — null-rate check skipped'
    }

    # Full evaluation simulation via the PRODUCTION engine
    $osBuild = 0
    [void][int]::TryParse("$($coll.systemInfo.osBuild)", [ref]$osBuild)
    $evalResults = @{ Pass=0; Fail=0; Warning=0; NotAssessed=0 }
    foreach ($chk in $checks) {
        if ($chk.type -ne 'Auto' -or -not $chk.collectionKeys) { $evalResults.NotAssessed++; continue }
        $testChk = New-TestCheck @{
            Id = $chk.id; CollectionKeys = $chk.collectionKeys; BaselineValue = $chk.baselineValue
            comparison = $chk.comparison; CspPath = $chk.cspPath; threshold = $chk.threshold
            eventIds = $chk.eventIds; filterField = $chk.filterField; filterValues = $chk.filterValues
        }
        $res = Get-CheckEvaluation -Chk $testChk -Json $coll -OsBuild $osBuild
        switch ($res.Status) {
            'Pass'    { $evalResults.Pass++ }
            'Fail'    { $evalResults.Fail++ }
            'Warning' { $evalResults.Warning++ }
            default   { $evalResults.NotAssessed++ }
        }
    }
    $totalEval = $evalResults.Pass + $evalResults.Fail + $evalResults.Warning
    Write-Host "    -> Pass: $($evalResults.Pass) | Fail: $($evalResults.Fail) | Warning: $($evalResults.Warning) | Not Assessed: $($evalResults.NotAssessed)" -ForegroundColor Gray
    Test-Assert 'At least 50 checks evaluated via production engine' ($totalEval -ge 50) "Evaluated: $totalEval"
} else {
    Test-Warn 'No sample collection JSON found — collection validation group skipped (synthetic groups above still exercise the engine)'
}

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
exit $(if ($Fail -gt 0) { 1 } else { 0 })
