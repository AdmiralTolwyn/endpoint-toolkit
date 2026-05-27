#requires -Version 7.0
<#
.SYNOPSIS
    Parse a ProcMon CSV/PML and correlate with a DumpPilot summary.
    Filters events by PID and time window around the crash, groups
    registry/file failures for the faulting module's paths.

.PARAMETER ProcMonPath
    Path to a ProcMon .pml or .csv file.

.PARAMETER SummaryPath
    Path to dump-summary.json (provides PID, process name, faulting module).

.PARAMETER WindowSeconds
    How many seconds before the last event to include. Default: 60.

.OUTPUTS
    [pscustomobject] with CorrelationPath (JSON), EventCount, FailureCount.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ProcMonPath,
    [Parameter(Mandatory = $true)][string]$SummaryPath,
    [int]$WindowSeconds = 60
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $ProcMonPath)) { throw "ProcMon file not found: $ProcMonPath" }
if (-not (Test-Path -LiteralPath $SummaryPath)) { throw "Summary not found: $SummaryPath" }

$summary = Get-Content -LiteralPath $SummaryPath -Raw | ConvertFrom-Json

# --- resolve CSV path ---------------------------------------------------------

$csvPath = $ProcMonPath
if ($ProcMonPath -match '\.pml$') {
    # Try to convert PML -> CSV using ProcMon.exe
    $procmon = @(
        'C:\Tools\Procmon.exe',
        'C:\SysinternalsSuite\Procmon.exe',
        (Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\Procmon.exe')
    ) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1

    if (-not $procmon) {
        $procmon = (Get-Command Procmon.exe -ErrorAction SilentlyContinue).Source
    }

    if ($procmon) {
        $csvPath = [System.IO.Path]::ChangeExtension($ProcMonPath, '.csv')
        if (-not (Test-Path -LiteralPath $csvPath)) {
            Write-Verbose "Converting PML -> CSV via $procmon"
            & $procmon /Quiet /OpenLog $ProcMonPath /SaveAs $csvPath
            if (-not (Test-Path -LiteralPath $csvPath)) {
                throw "ProcMon conversion failed. Export to CSV manually: File > Save As > CSV."
            }
        }
    } else {
        throw "PML file provided but Procmon.exe not found. Export to CSV manually or install Sysinternals."
    }
}

# --- compile fast CSV reader (C#, ~50x faster than Import-Csv) ----------------

if (-not ([System.Management.Automation.PSTypeName]'ProcMonReader').Type) {
    Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;

public class ProcMonEvent {
    public string Time;
    public string ProcessName;
    public int PID;
    public string Operation;
    public string Path;
    public string Result;
}

public class ProcMonParseResult {
    public List<ProcMonEvent> TargetFailures;   // failures for dump PID (or process name)
    public List<ProcMonEvent> VendorFailures;   // failures for GPU vendor service PIDs
    public List<ProcMonEvent> AllFailures;      // all non-SUCCESS events (for full-trace FaultRelated scan)
    public int TotalLines;
    public int TargetEvents;
    public DateTime LastTargetTime;
}

public static class ProcMonReader {
    // Results we never want to surface: SUCCESS / REPARSE / BUFFER * are
    // not failures at all. The rest are "informational" filter-driver
    // results that the OS routinely uses as control flow (FAST IO DISALLOWED
    // is a hint to fall back to slow IO; OPLOCK HANDLE CLOSED is a normal
    // lease-release event; NO MORE FILES / END OF FILE just mean the read
    // hit the end; NO EAS ON FILE is normal on most files; RANGE NOT LOCKED
    // is a routine unlock; FILE LOCKED WITH ONLY READERS is the normal
    // state for a memory-mapped image). Including them in the correlation
    // floods the top-25 with noise and hides real ACCESS DENIED / NAME NOT
    // FOUND / PATH NOT FOUND / NAME COLLISION rows under irrelevant filter
    // driver chatter. Excluded at parse time so they never reach disk.
    static readonly HashSet<string> SkipResults = new HashSet<string>(
        StringComparer.OrdinalIgnoreCase) {
        "SUCCESS", "BUFFER OVERFLOW", "REPARSE", "BUFFER TOO SMALL", "",
        "FAST IO DISALLOWED", "OPLOCK HANDLE CLOSED", "NO MORE FILES",
        "END OF FILE", "NO EAS ON FILE", "RANGE NOT LOCKED",
        "FILE LOCKED WITH ONLY READERS", "FILE LOCKED WITH WRITERS",
        "OPLOCK NOT GRANTED", "NOT REPARSE POINT", "IS DIRECTORY",
        "NOT A DIRECTORY"
    };

    static readonly HashSet<string> GpuServices = new HashSet<string>(
        StringComparer.OrdinalIgnoreCase) {
        "RaCEF.exe", "igfxEM.exe", "igfxCUIService.exe", "igfxTray.exe",
        "nvcontainer.exe", "NvDisplay.Container.exe", "AMDRSServ.exe", "RadeonSoftware.exe"
    };

    public static ProcMonParseResult Parse(string csvPath, int targetPid, string targetProcessName, bool gpuExpand) {
        var result = new ProcMonParseResult {
            TargetFailures = new List<ProcMonEvent>(),
            VendorFailures = new List<ProcMonEvent>(),
            AllFailures = new List<ProcMonEvent>(),
            TotalLines = 0,
            TargetEvents = 0,
            LastTargetTime = DateTime.MinValue
        };

        // Detect column indices from header
        int colTime = -1, colProc = -1, colPid = -1, colOp = -1, colPath = -1, colResult = -1;

        using (var sr = new StreamReader(csvPath, System.Text.Encoding.UTF8, true, 65536)) {
            string header = sr.ReadLine();
            if (header == null) return result;
            // Remove BOM if present
            if (header.Length > 0 && header[0] == '\uFEFF') header = header.Substring(1);
            var cols = SplitCsvLine(header);
            for (int i = 0; i < cols.Length; i++) {
                var c = cols[i].Trim().Trim('"');
                if (c == "Time of Day") colTime = i;
                else if (c == "Process Name") colProc = i;
                else if (c == "PID") colPid = i;
                else if (c == "Operation") colOp = i;
                else if (c == "Path") colPath = i;
                else if (c == "Result") colResult = i;
            }
            if (colResult < 0 || colPath < 0) return result; // invalid CSV

            string line;
            while ((line = sr.ReadLine()) != null) {
                result.TotalLines++;
                var fields = SplitCsvLine(line);
                if (fields.Length <= colResult) continue;

                string res = fields[colResult].Trim().Trim('"');
                string proc = colProc >= 0 && colProc < fields.Length ? fields[colProc].Trim().Trim('"') : "";
                int pid = 0;
                if (colPid >= 0 && colPid < fields.Length) int.TryParse(fields[colPid].Trim().Trim('"'), out pid);

                bool isTarget = (targetPid > 0 && pid == targetPid) ||
                                (targetPid <= 0 && proc.Equals(targetProcessName, StringComparison.OrdinalIgnoreCase));

                if (isTarget) {
                    result.TargetEvents++;
                    if (colTime >= 0 && colTime < fields.Length) {
                        DateTime t;
                        if (DateTime.TryParse(fields[colTime].Trim().Trim('"'), out t) && t > result.LastTargetTime)
                            result.LastTargetTime = t;
                    }
                }

                if (SkipResults.Contains(res)) continue;

                var ev = new ProcMonEvent {
                    Time = colTime >= 0 && colTime < fields.Length ? fields[colTime].Trim().Trim('"') : "",
                    ProcessName = proc,
                    PID = pid,
                    Operation = colOp >= 0 && colOp < fields.Length ? fields[colOp].Trim().Trim('"') : "",
                    Path = colPath >= 0 && colPath < fields.Length ? fields[colPath].Trim().Trim('"') : "",
                    Result = res
                };

                // Collect ALL failures for full-trace FaultRelated scan
                result.AllFailures.Add(ev);

                if (isTarget) {
                    result.TargetFailures.Add(ev);
                } else if (gpuExpand && GpuServices.Contains(proc)) {
                    result.VendorFailures.Add(ev);
                }
            }
        }
        return result;
    }

    static string[] SplitCsvLine(string line) {
        var fields = new List<string>();
        bool inQuote = false;
        int start = 0;
        for (int i = 0; i < line.Length; i++) {
            if (line[i] == '"') inQuote = !inQuote;
            else if (line[i] == ',' && !inQuote) {
                fields.Add(line.Substring(start, i - start));
                start = i + 1;
            }
        }
        fields.Add(line.Substring(start));
        return fields.ToArray();
    }
}
'@
}

# --- parse CSV via compiled reader --------------------------------------------

Write-Verbose "Parsing ProcMon CSV (compiled C# reader): $csvPath"

$faultMod = if ($summary.Faulting.Module) { $summary.Faulting.Module.ToLowerInvariant() } else { '' }
$faultModVendor = ''
if ($summary.ModuleDetails) {
    $md = @($summary.ModuleDetails | Where-Object { $_.Name -eq $summary.Faulting.Module } | Select-Object -First 1)
    if ($md.Count -gt 0 -and $md[0].CompanyName) { $faultModVendor = [string]$md[0].CompanyName }
}

$dumpPid = $null
if ($summary.Thread -and $summary.Thread.Cid) {
    $dumpPid = $summary.Thread.Cid -replace '\..*$', ''
}
$pidDec = if ($dumpPid) { [Convert]::ToInt32($dumpPid, 16) } else { 0 }
$procName = if ($summary.Process) { [string]$summary.Process.Name } else { '' }

$isGpuCrash = $faultMod -match 'igxel|igc|nvoglv|nvd3d|amdxx|atiux|d3d|opengl'
if (-not $isGpuCrash -and $summary.Faulting.Module) {
    $isGpuCrash = $summary.Faulting.Module -match 'igxel|nvoglv|amdxx'
}

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$parsed = [ProcMonReader]::Parse($csvPath, $pidDec, $procName, $isGpuCrash)
$sw.Stop()

Write-Verbose ("Parsed {0:N0} lines in {1:N1}s. Target events: {2:N0}, failures: {3:N0}, vendor: {4:N0}, all failures: {5:N0}" -f `
    $parsed.TotalLines, $sw.Elapsed.TotalSeconds, $parsed.TargetEvents, `
    $parsed.TargetFailures.Count, $parsed.VendorFailures.Count, $parsed.AllFailures.Count)

if ($parsed.TargetEvents -eq 0) {
    Write-Warning "No ProcMon events found for the target process."
    return [pscustomobject]@{
        CorrelationPath = $null
        EventCount      = 0
        FailureCount    = 0
    }
}

# --- time-windowed target failures -------------------------------------------

$windowStart = if ($parsed.LastTargetTime -gt [datetime]::MinValue) {
    $parsed.LastTargetTime.AddSeconds(-$WindowSeconds)
} else { $null }

$windowedFailures = [System.Collections.Generic.List[object]]::new()
foreach ($ev in $parsed.TargetFailures) {
    if ($windowStart) {
        try {
            $t = [datetime]::Parse($ev.Time)
            if ($t -lt $windowStart) { continue }
        } catch {}
    }
    [void]$windowedFailures.Add($ev)
}
# Add vendor service failures (no time window — they may be at startup)
foreach ($ev in $parsed.VendorFailures) { [void]$windowedFailures.Add($ev) }

Write-Verbose "Windowed failures: $($windowedFailures.Count)"

# --- group windowed failures --------------------------------------------------

$groups = [System.Collections.Generic.List[object]]::new()
$seen = @{}
foreach ($ev in $windowedFailures) {
    $normPath = $ev.Path -replace '\\0\d{3}\\', '\NNNN\' -replace '\\[{][0-9a-fA-F-]+[}]', '\{GUID}'
    $key = "$($ev.Operation)|$normPath|$($ev.Result)"
    if ($seen.ContainsKey($key)) { $seen[$key].Count++; continue }
    $entry = [pscustomobject]@{
        Operation = $ev.Operation; Path = $ev.Path; NormPath = $normPath
        Result = $ev.Result; Count = 1
    }
    $seen[$key] = $entry
    [void]$groups.Add($entry)
}
$groups = @($groups | Sort-Object Count -Descending)

# --- full-trace fault-related scan -------------------------------------------
#
# Relevance regex notes:
#   The naive substring `d3d` matched anywhere causes false positives
#   inside GUIDs (hex digits include d, 3 — see PowerSetting GUID
#   02F815B5-…-D3D8 and link-state-PM GUID d3d55efd-…). Those GUIDs
#   have nothing to do with Direct3D but get tagged as graphics-related
#   and pollute FaultRelated for *every* graphics-driver crash. Replace
#   the bare `d3d` token with patterns that require either a known
#   D3D filename suffix (.dll/.sys) or a path-segment boundary on both
#   sides. Same for `amd` (would otherwise match the `amd64` arch
#   suffix in every WinSxS path).
#
$Script:GraphicsVendorRx = '\bintel\b|\bnvidia\b|\bigfx\b|\bigcl\b|\bopengl\b|\bvulkan\b|\bdxgi\b|display\\igfx|d3d\d{1,2}(?:core|warp|on12)?\.(?:dll|sys)|d3dcompiler|d3dscache|direct3d|\\d3d\d{1,2}(?:\\|$)|atikm|amdkm|radeon|\bamd(?!64)'

$faultGroups = [System.Collections.Generic.List[object]]::new()
$faultSeen = @{}
foreach ($ev in $parsed.AllFailures) {
    $pathLower = $ev.Path.ToLowerInvariant()
    $related = $false
    if ($faultMod -and $pathLower -match [regex]::Escape($faultMod)) { $related = $true }
    if ($faultModVendor -and $pathLower -match [regex]::Escape($faultModVendor.Split(' ')[0])) { $related = $true }
    if ($pathLower -match $Script:GraphicsVendorRx) { $related = $true }
    if (-not $related) { continue }
    $normPath = $ev.Path -replace '\\0\d{3}\\', '\NNNN\' -replace '\\[{][0-9a-fA-F-]+[}]', '\{GUID}'
    $key = "$($ev.Operation)|$normPath|$($ev.Result)"
    if ($faultSeen.ContainsKey($key)) { $faultSeen[$key].Count++; continue }
    $entry = [pscustomobject]@{
        Operation=$ev.Operation; Path=$ev.Path; NormPath=$normPath
        Result=$ev.Result; Count=1; RelevantToFault=$true
    }
    $faultSeen[$key] = $entry
    [void]$faultGroups.Add($entry)
}
$faultGroups = @($faultGroups | Sort-Object Count -Descending)

# Tag windowed groups for relevance (for OtherFailures classification)
foreach ($g in $groups) {
    $pathLower = $g.Path.ToLowerInvariant()
    $related = $false
    if ($faultMod -and $pathLower -match [regex]::Escape($faultMod)) { $related = $true }
    if ($faultModVendor -and $pathLower -match [regex]::Escape($faultModVendor.Split(' ')[0])) { $related = $true }
    if ($pathLower -match $Script:GraphicsVendorRx) { $related = $true }
    $g | Add-Member -NotePropertyName 'RelevantToFault' -NotePropertyValue $related -Force
}

# --- build correlation output -------------------------------------------------

$correlation = [ordered]@{
    ProcMonSource   = $ProcMonPath
    ProcessPid      = $pidDec
    WindowSeconds   = $WindowSeconds
    TotalLines      = $parsed.TotalLines
    TotalEvents     = $parsed.TargetEvents
    FailureCount    = $windowedFailures.Count
    UniqueFailures  = $groups.Count
    ParseTimeSeconds = [Math]::Round($sw.Elapsed.TotalSeconds, 2)
    FaultRelated    = @($faultGroups)
    OtherFailures   = @($groups | Where-Object { -not $_.RelevantToFault } | Select-Object -First 20)
}

$outPath = Join-Path (Split-Path -Parent $SummaryPath) 'procmon-correlation.json'
[System.IO.File]::WriteAllText(
    $outPath,
    ($correlation | ConvertTo-Json -Depth 6),
    [System.Text.UTF8Encoding]::new($true)
)

[pscustomobject]@{
    CorrelationPath = $outPath
    EventCount      = $parsed.TargetEvents
    FailureCount    = $windowedFailures.Count
    FaultRelated    = @($faultGroups).Count
}
