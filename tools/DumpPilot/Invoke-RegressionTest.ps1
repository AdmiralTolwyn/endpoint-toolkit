#requires -Version 5.1
<#
.SYNOPSIS
    Regression test for DumpPilot's pattern matcher. Runs every fixture in
    corpus/ through Invoke-DumpAnalysis.ps1 and asserts the expected rule
    fires (or that no rule fires).

.DESCRIPTION
    Fixtures are *.fixture.json files, each a tiny synthetic dump-summary.json
    plus an `__expect` block:

        {
          "DumpPath": "...",
          "Process": { "Name": "java.exe", "CommandLine": "..." },
          "Exception": { "Code": "c000041d", ... },
          "Faulting": { "Module": "igxelpgicd64", ..., "Stack": [...] },
          "__expect": { "rule": "java-intel-opengl-icd" }   // or { "rule": null }
        }

    Fixtures are cheap, ship in git, and don't need cdb or actual dumps.

.PARAMETER CorpusDir
    Default: ../corpus relative to this script.

.OUTPUTS
    Non-zero exit code on any failure.
#>
[CmdletBinding()]
param(
    [string]$CorpusDir
)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $PSCommandPath
if (-not $CorpusDir) { $CorpusDir = Join-Path $here 'corpus' }
if (-not (Test-Path -LiteralPath $CorpusDir)) {
    throw "Corpus folder not found: $CorpusDir"
}

$analyzer = Join-Path $here 'helpers\Invoke-DumpAnalysis.ps1'
$validator = Join-Path $here 'helpers\Test-PatternDb.ps1'

Write-Host '--- pattern_db.json validation ---'
& $validator
if (-not $?) { throw "Pattern DB validation failed; aborting regression run." }
Write-Host ''
Write-Host '--- fixture matcher regression ---'

$fixtures = Get-ChildItem -LiteralPath $CorpusDir -Filter '*.fixture.json' -File
if (-not $fixtures) { Write-Warning "No fixtures found in $CorpusDir"; return }

$tmp = Join-Path $env:TEMP ('DumpPilotRegression_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null

$pass = 0; $fail = 0
$failures = New-Object System.Collections.Generic.List[object]

try {
    foreach ($fx in $fixtures) {
        $obj = Get-Content -LiteralPath $fx.FullName -Raw | ConvertFrom-Json
        $expect = $obj.__expect
        # strip __expect from the summary we feed to the analyzer
        $obj.PSObject.Properties.Remove('__expect')
        $summaryPath = Join-Path $tmp ($fx.BaseName + '.summary.json')
        [System.IO.File]::WriteAllText($summaryPath, ($obj | ConvertTo-Json -Depth 10), [System.Text.UTF8Encoding]::new($true))

        $r = & $analyzer -SummaryPath $summaryPath | Out-Null
        $reportPath = $summaryPath -replace '\.summary\.json$', '.report.json'
        # Analyzer derives output base from the summary path basename, so:
        $base = [System.IO.Path]::GetFileNameWithoutExtension($summaryPath)
        $base = $base -replace '\.dump-summary$',''
        $reportPath = Join-Path $tmp ($base + '.report.json')
        if (-not (Test-Path -LiteralPath $reportPath)) {
            $fail++; [void]$failures.Add([pscustomobject]@{ Fixture = $fx.Name; Reason = "report not produced ($reportPath)" })
            continue
        }
        $rep = Get-Content -LiteralPath $reportPath -Raw | ConvertFrom-Json
        $top = if ($rep.PatternHits -and $rep.PatternHits.Count -gt 0) { $rep.PatternHits[0].Id } else { $null }

        if ($expect.rule -eq $top) {
            $pass++
            Write-Host ("[PASS] {0,-40} -> {1}" -f $fx.Name, ($top | ForEach-Object { if ($_) { $_ } else { '(no match)' } }))
        } else {
            $fail++
            $msg = "expected '$($expect.rule)' got '$top' (hits=$($rep.HitCount))"
            Write-Host ("[FAIL] {0,-40} {1}" -f $fx.Name, $msg) -ForegroundColor Red
            [void]$failures.Add([pscustomobject]@{ Fixture = $fx.Name; Reason = $msg })
        }
    }
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ''
$color = if ($fail -gt 0) { 'Red' } else { 'Green' }
Write-Host ("Total: {0}  Pass: {1}  Fail: {2}" -f ($pass + $fail), $pass, $fail) -ForegroundColor $color

if ($fail) { exit 1 }
