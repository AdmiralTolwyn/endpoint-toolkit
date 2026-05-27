#requires -Version 7.0
<#
.SYNOPSIS
    Run DumpPilot over every .dmp / .mdmp in a folder and emit a CSV index of
    results. WER LocalDumps drops files into a single folder per process; this
    is the realistic batch-triage entry point.

.PARAMETER FolderPath
    Folder containing dump files.

.PARAMETER Recurse
    Recurse into subfolders.

.PARAMETER SymbolPath
    Override _NT_SYMBOL_PATH for all runs.

.PARAMETER Force
    Re-run dumps that already have a .dump-facts folder.

.OUTPUTS
    CSV at <FolderPath>\DumpPilot-batch.csv (one row per dump):
        DumpPath, DumpKind, Process, Bucket, FaultingModule, TopHit, HitCount,
        ReportMd, Error
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$FolderPath,
    [switch]$Recurse,
    [string]$SymbolPath,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $PSCommandPath

if (-not (Test-Path -LiteralPath $FolderPath)) { throw "Folder not found: $FolderPath" }

$gciArgs = @{ LiteralPath = $FolderPath; File = $true; Include = '*.dmp','*.mdmp' }
if ($Recurse) { $gciArgs.Recurse = $true }
# -Include needs -Path (not -LiteralPath) for wildcard matching, so do it manually.
$dumps = if ($Recurse) {
    Get-ChildItem -LiteralPath $FolderPath -File -Recurse | Where-Object { $_.Extension -in '.dmp','.mdmp' }
} else {
    Get-ChildItem -LiteralPath $FolderPath -File          | Where-Object { $_.Extension -in '.dmp','.mdmp' }
}

if (-not $dumps) { Write-Warning "No .dmp/.mdmp files found under $FolderPath"; return }

Write-Host ("DumpPilot batch: {0} dump(s) under {1}" -f $dumps.Count, $FolderPath)

$rows = New-Object System.Collections.Generic.List[object]
$i = 0
foreach ($d in $dumps) {
    $i++
    Write-Host ("[{0}/{1}] {2}" -f $i, $dumps.Count, $d.Name)
    $row = [ordered]@{
        DumpPath       = $d.FullName
        DumpKind       = $null
        Process        = $null
        Bucket         = $null
        FaultingModule = $null
        TopHit         = $null
        HitCount       = 0
        ReportMd       = $null
        Error          = $null
    }
    try {
        $args = @{ DumpPath = $d.FullName; Force = $Force }
        if ($SymbolPath) { $args.SymbolPath = $SymbolPath }
        $r = & (Join-Path $here 'DumpPilot.ps1') @args
        $row.DumpKind       = $r.DumpKind
        $row.Process        = $r.Process
        $row.Bucket         = $r.Bucket
        $row.TopHit         = $r.TopHit
        $row.HitCount       = $r.HitCount
        $row.ReportMd       = $r.ReportMd
        # FaultingModule isn't on the DumpPilot output object; pull from summary.
        try {
            $sum = Get-Content -LiteralPath $r.SummaryJson -Raw | ConvertFrom-Json
            $row.FaultingModule = $sum.Faulting.Module
        } catch {}
    } catch {
        $row.Error = $_.Exception.Message
        Write-Warning ("  failed: {0}" -f $_.Exception.Message)
    }
    [void]$rows.Add([pscustomobject]$row)
}

$csvPath = Join-Path $FolderPath 'DumpPilot-batch.csv'
$rows | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
Write-Host ("Index: {0}" -f $csvPath)

# Summary by TopHit
$rows | Group-Object TopHit | Sort-Object Count -Descending |
    Select-Object @{n='TopHit';e={if ($_.Name) { $_.Name } else { '(no match)' }}}, Count |
    Format-Table -AutoSize

[pscustomobject]@{
    CsvPath = $csvPath
    Total   = $rows.Count
    Matched = ($rows | Where-Object { $_.HitCount -gt 0 }).Count
    Errors  = ($rows | Where-Object { $_.Error }).Count
}
