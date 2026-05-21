#requires -Version 5.1
<#
.SYNOPSIS
    DumpPilot top-level entry. Runs Export-DumpFacts -> ConvertTo-DumpSummary
    -> Invoke-DumpAnalysis on one dump and prints where the artifacts landed.

.PARAMETER DumpPath
    Path to the .dmp / .mdmp / MEMORY.DMP file.

.PARAMETER SymbolPath
    Override _NT_SYMBOL_PATH for this run.

.PARAMETER Force
    Re-run all stages, overwriting prior artifacts.

.EXAMPLE
    .\DumpPilot.ps1 -DumpPath 'C:\Dumps\java.exe.6356.dmp' -Verbose
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$DumpPath,
    [string]$SymbolPath,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $PSCommandPath

Write-Verbose 'Stage 1: Export-DumpFacts'
$exportArgs = @{ DumpPath = $DumpPath; Force = $Force }
if ($SymbolPath) { $exportArgs.SymbolPath = $SymbolPath }
$exp = & (Join-Path $here 'Export-DumpFacts.ps1') @exportArgs
Write-Verbose ("  facts root : {0}" -f $exp.FactsRoot)
Write-Verbose ("  kind       : {0}" -f $exp.DumpKind)

Write-Verbose 'Stage 2: ConvertTo-DumpSummary'
$sum = & (Join-Path $here 'ConvertTo-DumpSummary.ps1') -FactsRoot $exp.FactsRoot
Write-Verbose ("  summary    : {0}" -f $sum.OutputPath)

Write-Verbose 'Stage 3: Build facts-only report + LLM prompt'
$rep = & (Join-Path $here 'helpers\Invoke-DumpAnalysis.ps1') -SummaryPath $sum.OutputPath
Write-Verbose ("  report.md  : {0}" -f $rep.ReportMd)
Write-Verbose ("  mode       : facts-only")

Write-Verbose 'Stage 4: Export-DumpHtmlReport'
$htm = & (Join-Path $here 'helpers\Export-DumpHtmlReport.ps1') -ReportJsonPath $rep.ReportJson
Write-Verbose ("  html       : {0}" -f $htm.HtmlPath)

[pscustomobject]@{
    FactsRoot   = $exp.FactsRoot
    SummaryJson = $sum.OutputPath
    ReportJson  = $rep.ReportJson
    ReportMd    = $rep.ReportMd
    ReportHtml  = $htm.HtmlPath
    PromptMd    = $rep.PromptMd
    DumpKind    = $exp.DumpKind
    Process     = $sum.Process
    Bucket      = $sum.Bucket
    HitCount    = $rep.HitCount
    TopHit      = $rep.TopHit
}
