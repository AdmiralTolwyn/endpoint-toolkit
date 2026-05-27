#requires -Version 7.0
<#
.SYNOPSIS
    Validate scripts/DumpPilot/pattern_db.json structurally. Cheap guardrail
    against a future rule being added with a typo, an invalid regex, or a
    duplicate id.

.PARAMETER PatternDbPath
    Default: ../pattern_db.json relative to this script.

.OUTPUTS
    Non-zero exit code on any validation failure. Prints a per-rule report.
#>
[CmdletBinding()]
param([string]$PatternDbPath)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $PSCommandPath
if (-not $PatternDbPath) { $PatternDbPath = Join-Path (Split-Path -Parent $here) 'pattern_db.json' }
if (-not (Test-Path -LiteralPath $PatternDbPath)) { throw "Pattern DB not found: $PatternDbPath" }

$db = Get-Content -LiteralPath $PatternDbPath -Raw | ConvertFrom-Json

$problems = New-Object System.Collections.Generic.List[string]
$seenIds  = New-Object System.Collections.Generic.HashSet[string]
$validMatchFields = @('ProcessImage','FaultingModule','ExceptionCode','StackRegex','CommandLineRegex')
$validSeverities  = @('Critical','High','Medium','Low','Info')

if (-not $db.rules -or $db.rules.Count -eq 0) {
    $problems.Add("rules array is empty or missing") | Out-Null
}

foreach ($r in $db.rules) {
    $rid = if ($r.id) { $r.id } else { '<missing-id>' }

    if (-not $r.id)             { $problems.Add("$rid : missing 'id'") | Out-Null }
    if (-not $r.title)          { $problems.Add("$rid : missing 'title'") | Out-Null }
    if (-not $r.severity)       { $problems.Add("$rid : missing 'severity'") | Out-Null }
    elseif ($validSeverities -notcontains $r.severity) {
        $problems.Add("$rid : severity '$($r.severity)' not in [$($validSeverities -join ', ')]") | Out-Null
    }
    if (-not $r.match)          { $problems.Add("$rid : missing 'match' block") | Out-Null; continue }

    if ($r.id) {
        if (-not $seenIds.Add($r.id)) { $problems.Add("$rid : duplicate id") | Out-Null }
    }

    # match block: at least one constraint, all keys recognized, all regexes compile
    $matchKeys = @($r.match.PSObject.Properties.Name)
    if ($matchKeys.Count -eq 0) {
        $problems.Add("$rid : 'match' block is empty (rule would fire on every dump)") | Out-Null
    }
    foreach ($k in $matchKeys) {
        if ($validMatchFields -notcontains $k) {
            $problems.Add("$rid : unknown match field '$k' (valid: $($validMatchFields -join ', '))") | Out-Null
        }
        $v = $r.match.$k
        if ($v) {
            try { [void][regex]::new($v) }
            catch { $problems.Add("$rid : match.$k is not a valid regex: $($_.Exception.Message)") | Out-Null }
        }
    }
    Write-Host ("[OK ] {0,-25} sev={1} fields={2}" -f $rid, $r.severity, ($matchKeys -join ','))
}

Write-Host ''
if ($problems.Count -gt 0) {
    Write-Host ("Validation FAILED ({0} problem(s)):" -f $problems.Count) -ForegroundColor Red
    foreach ($p in $problems) { Write-Host ("  - " + $p) -ForegroundColor Red }
    exit 1
} else {
    Write-Host ("Validation OK ({0} rule(s))." -f $db.rules.Count) -ForegroundColor Green
}
