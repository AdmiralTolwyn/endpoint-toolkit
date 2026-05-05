<#
.SYNOPSIS
    Builds a phase timeline of a Windows in-place upgrade from setupact.log.

.DESCRIPTION
    Parses OPERATIONTRACK markers ("Start execution phase <Name>") and phase exit
    markers ("Execution phase [<Name>] exiting with HRESULT [...]") to compute how
    long each upgrade phase actually took. Uses a compiled C# helper (StreamReader
    + regex) so multi-GB setupact.log files are processed in seconds.

    By default the Downlevel phase is excluded because the user is still able to
    work during it; pass -IncludeDownlevel to keep it. If a setupact_unattendGC.log
    sits next to setupact.log, its WinDeploy/OOBE timestamps are used to extend
    the timeline through the OOBE phase.

    Idle gaps (system shut down, hibernated, or asleep) are detected by looking
    for jumps between consecutive log timestamps that exceed -IdleGapSeconds and
    are subtracted from each phase's Active duration.

.PARAMETER LogPath
    Path to setupact.log OR a folder. Recognised folder layouts:
        C:\Windows\Panther                       (post-upgrade success)
        C:\Windows\Panther\UnattendGC            (just the OOBE log)
        C:\$WINDOWS.~BT\Sources\Panther          (in-flight / rolled-back)
    Custom paths are accepted as long as they follow the same canonical
    layout: <root>\setupact.log plus optional <root>\UnattendGC\setupact.log.
    If omitted, the local machine's Panther folder is auto-discovered.

.PARAMETER IncludeDownlevel
    Include the Downlevel phase in the timeline (off by default).

.PARAMETER AsObject
    Emit the timeline as PSCustomObjects instead of writing the formatted report.

.PARAMETER IdleGapSeconds
    Maximum allowed gap (seconds) between consecutive log lines inside a phase
    before that gap is treated as the system being shut down or asleep. Default
    is 600s (10 min). Idle gaps are excluded from the per-phase Active duration.
    Lower this value with care: legitimate Setup operations such as driver
    install, dynamic update download, or BCD writes can pause logging for
    several minutes and will be misclassified as idle if the threshold is too
    aggressive.

.PARAMETER TotalActiveMinutes
    Suppress the table and emit only the rounded total active upgrade time in
    minutes. Useful for piping into CSV / CI pipelines.

.PARAMETER Csv
    Emit the timeline as CSV (semicolon-separated) instead of the formatted
    table. Implies -AsObject style data without the human-friendly rendering.

.EXAMPLE
    .\Get-SetupTimeline.ps1 -LogPath C:\temp\TK\setupact_MX-PF5041LA

.EXAMPLE
    # Auto-discover - run on the just-upgraded machine itself.
    .\Get-SetupTimeline.ps1 -Verbose

.EXAMPLE
    .\Get-SetupTimeline.ps1 -LogPath C:\Windows\Panther -IncludeDownlevel

.EXAMPLE
    $mins = .\Get-SetupTimeline.ps1 -LogPath C:\temp\setupact.log -TotalActiveMinutes

.NOTES
    Author     : Anton Romanyuk
    Version    : 1.1.0
    Requires   : Windows PowerShell 5.1 or PowerShell 7+
    License    : MIT

    DISCLAIMER
    ----------
    This script is provided AS-IS, without warranty of any kind, express or
    implied, including but not limited to the warranties of merchantability,
    fitness for a particular purpose and non-infringement. In no event shall
    the author or copyright holders be liable for any claim, damages or other
    liability arising from the use of this script. It is not a Microsoft
    product and is not supported by Microsoft. Always validate results against
    your own setup logs before drawing conclusions or making business decisions.
#>
[CmdletBinding()]
param(
    # Path to setupact.log itself, OR a folder. The folder may be:
    #   * C:\Windows\Panther                       (post-upgrade, success path)
    #   * C:\Windows\Panther\UnattendGC             (just the OOBE log)
    #   * C:\$WINDOWS.~BT\Sources\Panther           (in-flight or rolled-back upgrade)
    #   * any custom <root> that mirrors the canonical layout, i.e. contains
    #     setupact.log and optionally an UnattendGC\setupact.log subfolder
    # If omitted, the script auto-discovers logs in the standard locations on
    # the local machine (C:\Windows\Panther first, then C:\$WINDOWS.~BT\...).
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$LogPath,

    # Include Downlevel as a row (off by default - user is productive during it).
    [switch]$IncludeDownlevel,

    # Return PSCustomObjects instead of the formatted report.
    [switch]$AsObject,

    # Gap between consecutive timestamps above which the system is treated as
    # off / asleep. Default 600s comfortably exceeds legitimate Setup pauses
    # (driver install, dynamic update, BCD write).
    [int]$IdleGapSeconds = 600,

    # Emit only the rounded total active upgrade time in minutes.
    [switch]$TotalActiveMinutes,

    # Emit the timeline as semicolon-separated CSV.
    [switch]$Csv
)

$ErrorActionPreference = 'Stop'

# Bump $ScriptVersion when behaviour or output format changes.
$ScriptVersion = '1.1.0'
$ScriptAuthor  = 'Anton Romanyuk'

Write-Verbose ("Get-SetupTimeline v{0}" -f $ScriptVersion)
Write-Verbose ("PowerShell {0} on {1}" -f $PSVersionTable.PSVersion, [System.Environment]::OSVersion.VersionString)

# =============================================================================
# STEP 1 - Resolve input path to (mainLog, unattendLog).
# =============================================================================
# Canonical Panther layouts:
#   C:\Windows\Panther\setupact.log               (post-upgrade success)
#   C:\Windows\Panther\UnattendGC\setupact.log    (OOBE log)
#   C:\$WINDOWS.~BT\Sources\Panther\...           (in-flight / rolled-back)
# Custom paths must mirror this layout: <root>\setupact.log + optional
# <root>\UnattendGC\setupact.log.
# Ref: https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/add-a-custom-script-to-windows-setup
# -----------------------------------------------------------------------------

# Returns ($mainLog, $unattendLog). $unattendLog is $null if not found.
function Resolve-PantherPaths([string]$Path) {
    $abs = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path

    # File mode
    if (-not (Test-Path -LiteralPath $abs -PathType Container)) {
        $main = $abs
        $dir  = Split-Path $abs -Parent
        $leaf = Split-Path $dir -Leaf
        # If user pointed at the UnattendGC log, main log is one level up.
        if ($leaf -ieq 'UnattendGC') {
            $parent = Split-Path $dir -Parent
            $maybeMain = Join-Path $parent 'setupact.log'
            if (Test-Path -LiteralPath $maybeMain) {
                return @($maybeMain, $abs)
            }
        }
        $una = Join-Path $dir 'UnattendGC\setupact.log'
        if (-not (Test-Path -LiteralPath $una)) { $una = $null }
        return @($main, $una)
    }

    # Folder mode
    $leaf = Split-Path $abs -Leaf
    # If caller pointed at \Panther\UnattendGC, hop up one level.
    if ($leaf -ieq 'UnattendGC') {
        $parent = Split-Path $abs -Parent
        $maybeMain = Join-Path $parent 'setupact.log'
        if (Test-Path -LiteralPath $maybeMain) {
            return @($maybeMain, (Join-Path $abs 'setupact.log'))
        }
    }

    $main = Join-Path $abs 'setupact.log'
    $una  = Join-Path $abs 'UnattendGC\setupact.log'
    if (-not (Test-Path -LiteralPath $una)) { $una = $null }
    return @($main, $una)
}

# Friendly single-line error + return (avoids noisy PS stacktrace block).
function Write-FriendlyError {
    param([string]$Message, [string[]]$Hint)
    Write-Host ''
    Write-Host ('  ERROR: {0}' -f $Message) -ForegroundColor Red
    if ($Hint) {
        Write-Host ''
        foreach ($h in $Hint) { Write-Host ('         {0}' -f $h) -ForegroundColor DarkYellow }
    }
    Write-Host ''
}

if (-not $LogPath) {
    # Auto-discover: probe canonical Panther locations, first hit wins.
    $defaultRoots = @(
        "$env:SystemRoot\Panther"
        "$env:SystemDrive\`$WINDOWS.~BT\Sources\Panther"
        "$env:SystemDrive\`$Windows.~BT\Sources\Panther"
    )
    $picked = $null
    foreach ($root in $defaultRoots) {
        $candidate = Join-Path $root 'setupact.log'
        if (Test-Path -LiteralPath $candidate) {
            $picked = $root
            Write-Verbose ("LogPath not specified - using local Panther log at: {0}" -f $root)
            break
        }
    }
    if (-not $picked) {
        Write-FriendlyError -Message 'No -LogPath was specified and no setupact.log was found on this machine.' -Hint @(
            'Searched these canonical locations:'
            ($defaultRoots | ForEach-Object { '  * ' + $_ }) -join [Environment]::NewLine
            ''
            'Pass -LogPath explicitly, e.g.:'
            '  .\Get-SetupTimeline.ps1 -LogPath C:\Windows\Panther'
            '  .\Get-SetupTimeline.ps1 -LogPath C:\$WINDOWS.~BT\Sources\Panther'
            '  .\Get-SetupTimeline.ps1 -LogPath C:\path\to\setupact.log'
        )
        return
    }
    $LogPath = $picked
}

# Validate before handing off (otherwise Resolve-Path raises ItemNotFound).
if (-not (Test-Path -LiteralPath $LogPath)) {
    Write-FriendlyError -Message ("-LogPath does not exist: {0}" -f $LogPath) -Hint @(
        'Pass either a Panther folder OR a setupact.log file path:'
        '  .\Get-SetupTimeline.ps1 -LogPath C:\Windows\Panther'
        '  .\Get-SetupTimeline.ps1 -LogPath C:\path\to\setupact.log'
    )
    return
}

try {
    $pair = Resolve-PantherPaths -Path $LogPath
} catch {
    Write-FriendlyError -Message ("Could not resolve -LogPath '{0}': {1}" -f $LogPath, $_.Exception.Message)
    return
}
$mainLog     = $pair[0]
$unattendLog = $pair[1]

if (-not (Test-Path -LiteralPath $mainLog)) {
    $isFolder = Test-Path -LiteralPath $LogPath -PathType Container
    if ($isFolder) {
        $hint = @(
            ('Folder exists but contains no setupact.log: {0}' -f $LogPath)
            ''
            'Expected one of:'
            ('  {0}\setupact.log               (main log)' -f $LogPath)
            ('  {0}\UnattendGC\setupact.log   (OOBE log)'  -f $LogPath)
            ''
            'Typical Panther roots:'
            '  C:\Windows\Panther                       (post-upgrade success)'
            '  C:\$WINDOWS.~BT\Sources\Panther          (in-flight / rolled back)'
        )
    } else {
        $hint = @(
            'Pass either a Panther folder OR a setupact.log file path:'
            '  .\Get-SetupTimeline.ps1 -LogPath C:\Windows\Panther'
            '  .\Get-SetupTimeline.ps1 -LogPath C:\path\to\setupact.log'
        )
    }
    Write-FriendlyError -Message ('setupact.log not found at: {0}' -f $mainLog) -Hint $hint
    return
}

# Surface file sizes in -Verbose; helps diagnose missing-OOBE-phase reports.
$mainSize = (Get-Item -LiteralPath $mainLog).Length
Write-Verbose ("Main log     : {0} ({1:N1} MB)" -f $mainLog, ($mainSize / 1MB))
if ($unattendLog -and (Test-Path -LiteralPath $unattendLog)) {
    $unaSize = (Get-Item -LiteralPath $unattendLog).Length
    Write-Verbose ("UnattendGC   : {0} ({1:N1} MB)" -f $unattendLog, ($unaSize / 1MB))
} else {
    $unattendLog = $null
    Write-Verbose 'UnattendGC   : (not present - OOBE timeline will be skipped)'
}
Write-Verbose ("Idle gap     : > {0} seconds treated as off / standby" -f $IdleGapSeconds)
Write-Verbose ("Downlevel    : {0}" -f ($(if ($IncludeDownlevel) { 'INCLUDED' } else { 'excluded' })))

# =============================================================================
# STEP 2 - Compile the C# scanner once per AppDomain.
# =============================================================================
# A pure-PS parser is 50-100x slower on multi-hundred-MB setupact.log files.
# The 'as [type]' guard avoids re-compilation on subsequent invocations.
# -----------------------------------------------------------------------------
if (-not ('SetupTimelineV2.Scanner' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace SetupTimelineV2 {
    // One OPERATIONTRACK / phase-exit / WinDeploy event extracted from a log line.
    public class Marker {
        public DateTime Time;
        // Monotonic counter - tiebreaker because Sort-Object is NOT stable on
        // PS 5.1 and same-second markers (e.g. Downlevel-end / Pre-Finalize-start)
        // can otherwise swap order and collapse a phase to 0s.
        public int      Order;
        public string   Phase;
        public string   Kind;    // "Start" | "Exit"
        public string   HResult; // set only on Exit markers
    }

    // One detected idle window (consecutive timestamps further apart than the
    // threshold = system was off / hibernated / asleep).
    public class GapInfo {
        public DateTime From;
        public DateTime To;
        public TimeSpan Duration;
    }

    public class ScanResult {
        public List<Marker>  Markers  = new List<Marker>();
        public List<GapInfo> Gaps     = new List<GapInfo>();
        // Last timestamped line - used to close trailing open phases on
        // in-flight upgrades (no "exiting with HRESULT" written yet).
        public DateTime      LastTime = DateTime.MinValue;
        public int           LineCount;
    }

    public static class Scanner {
        // Pre-compiled regexes; per-process compilation cost paid once.
        private static readonly Regex _start = new Regex(
            @"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}),.*OPERATIONTRACK ExecuteOperations: Start execution phase (.+?)\s*$",
            RegexOptions.Compiled);
        private static readonly Regex _exit = new Regex(
            @"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}),.*Execution phase \[(.+?)\] exiting with HRESULT \[(0x[0-9A-Fa-f]+)\]",
            RegexOptions.Compiled);
        // First/last windeploy.exe lines bracket the synthetic WinDeploy/OOBE
        // phase (setupact.log proper does not record this window).
        private static readonly Regex _windeploy = new Regex(
            @"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}),.*\[windeploy\.exe\]",
            RegexOptions.Compiled);
        // Cheap prefix matcher run on every line for idle-gap detection.
        private static readonly Regex _ts = new Regex(
            @"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}),",
            RegexOptions.Compiled);

        // Single-pass scan. unattendGc=true captures only WinDeploy markers.
        public static ScanResult Scan(string path, bool unattendGc, int idleGapSeconds) {
            var result       = new ScanResult();
            var gapThreshold = TimeSpan.FromSeconds(idleGapSeconds);
            DateTime prevTs  = DateTime.MinValue;
            DateTime firstWd = DateTime.MinValue, lastWd = DateTime.MinValue;
            int order        = 0;

            // FileShare.ReadWrite so we can scan while Setup is still writing.
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 1 << 16))
            using (var sr = new StreamReader(fs)) {
                string line;
                while ((line = sr.ReadLine()) != null) {
                    if (line.Length < 20) continue;

                    var tm = _ts.Match(line);
                    if (!tm.Success) continue;

                    DateTime ts;
                    if (!DateTime.TryParse(tm.Groups[1].Value, out ts)) continue;
                    result.LineCount++;
                    result.LastTime = ts;

                    // Idle gap detection - Setup itself never pauses for more
                    // than a few seconds, so big gaps = device off/asleep.
                    if (prevTs != DateTime.MinValue) {
                        var delta = ts - prevTs;
                        if (delta > gapThreshold) {
                            result.Gaps.Add(new GapInfo { From = prevTs, To = ts, Duration = delta });
                        }
                        // delta < 0 (clock jumped back, e.g. VM time-sync) is silently ignored.
                    }
                    prevTs = ts;

                    if (!unattendGc) {
                        var m = _start.Match(line);
                        if (m.Success) {
                            result.Markers.Add(new Marker {
                                Time = ts, Order = order++, Phase = m.Groups[2].Value, Kind = "Start"
                            });
                            continue; // mutually exclusive with _exit on the same line
                        }
                        m = _exit.Match(line);
                        if (m.Success) {
                            result.Markers.Add(new Marker {
                                Time = ts, Order = order++, Phase = m.Groups[2].Value,
                                Kind = "Exit", HResult = m.Groups[3].Value
                            });
                        }
                    } else {
                        // UnattendGC: track first/last windeploy line; emit synthetic phase below.
                        var m = _windeploy.Match(line);
                        if (m.Success) {
                            if (firstWd == DateTime.MinValue) firstWd = ts;
                            lastWd = ts;
                        }
                    }
                }

                if (unattendGc && firstWd != DateTime.MinValue) {
                    result.Markers.Add(new Marker {
                        Time = firstWd, Order = order++, Phase = "WinDeploy/OOBE", Kind = "Start"
                    });
                    result.Markers.Add(new Marker {
                        Time = lastWd, Order = order++, Phase = "WinDeploy/OOBE", Kind = "Exit"
                    });
                }
            }
            return result;
        }
    }
}
'@
}

# =============================================================================
# STEP 3 - Run the scanner (main log + optional OOBE log).
# =============================================================================
$sw      = [System.Diagnostics.Stopwatch]::StartNew()
$scan    = [SetupTimelineV2.Scanner]::Scan($mainLog, $false, $IdleGapSeconds)
# Strongly-typed list lets us .Add() / [-1] index across PS 5.1 and PS 7.
$markers = [System.Collections.Generic.List[object]]::new()
foreach ($m in $scan.Markers) { [void]$markers.Add($m) }
$gaps        = $scan.Gaps
$lastLogTime = $scan.LastTime  # used to close trailing in-flight phases
$sw.Stop()
$mbPerSec = if ($sw.Elapsed.TotalSeconds -gt 0) { ($mainSize / 1MB) / $sw.Elapsed.TotalSeconds } else { 0 }
Write-Verbose ("Scanned setupact.log in {0:N1}s ({1:N1} MB/s) - {2} phase markers, {3} idle gaps" -f `
    $sw.Elapsed.TotalSeconds, $mbPerSec, $markers.Count, $gaps.Count)

# Merge OOBE log into the same marker/gap pools so segment-builder sees one stream.
if ($unattendLog -and (Test-Path -LiteralPath $unattendLog)) {
    $sw2 = [System.Diagnostics.Stopwatch]::StartNew()
    $wd  = [SetupTimelineV2.Scanner]::Scan($unattendLog, $true, $IdleGapSeconds)
    $sw2.Stop()
    foreach ($m in $wd.Markers) { [void]$markers.Add($m) }
    foreach ($g in $wd.Gaps)    { $gaps += $g }
    Write-Verbose ("Scanned {0} in {1:N1}s - {2} markers, {3} idle gaps" -f `
        (Split-Path $unattendLog -Leaf), $sw2.Elapsed.TotalSeconds, $wd.Markers.Count, $wd.Gaps.Count)
}

if ($markers.Count -eq 0) {
    Write-Warning "No phase markers found. Is this really a setupact.log from an in-place upgrade?"
    return
}


# =============================================================================
# STEP 4 - Markers -> contiguous phase segments.
# =============================================================================
# Each Start opens a segment; segment closes at matching Exit OR the next
# different-phase Start. Same-phase repeats (Downlevel retries) keep the
# earliest Start.
#
# Sort by (Time, Order): Sort-Object is NOT stable on PS 5.1, and same-second
# markers at phase handovers would otherwise collapse a phase to 0s.
# -----------------------------------------------------------------------------
$markers = $markers | Sort-Object Time, Order

$segments = New-Object System.Collections.Generic.List[object]
$open     = $null
foreach ($m in $markers) {
    if ($m.Kind -eq 'Start') {
        if ($null -eq $open) {
            $open = [pscustomobject]@{ Phase = $m.Phase; Start = $m.Time; End = $null; HResult = $null }
        } elseif ($open.Phase -ne $m.Phase) {
            # Different phase started without explicit Exit - close at transition.
            $open.End = $m.Time
            $segments.Add($open)
            $open = [pscustomobject]@{ Phase = $m.Phase; Start = $m.Time; End = $null; HResult = $null }
        }
        # else: same-phase repeat (Downlevel retry) -> keep earliest Start
    } elseif ($m.Kind -eq 'Exit') {
        if ($null -ne $open -and $open.Phase -eq $m.Phase) {
            $open.End     = $m.Time
            $open.HResult = $m.HResult
            $segments.Add($open)
            $open = $null
        }
        # stray Exit without matching Start -> ignored
    }
}
# Trailing open phase: 'End' is terminal-by-design (Setup never writes its
# Exit and keeps logging post-upgrade noise) so collapse to a zero marker.
# Any other open phase is from an in-flight/interrupted upgrade -> tag '(running)'.
if ($null -ne $open) {
    if ($open.Phase -eq 'End') {
        $open.End     = $open.Start
        $open.HResult = $null
    } else {
        $endTs = if ($lastLogTime -gt $open.Start) { $lastLogTime } else { ($markers[-1]).Time }
        $open.End     = $endTs
        $open.HResult = '(running)'
    }
    $segments.Add($open)
}

# Start markers but no Exits -> in-flight / rolled-back upgrade. Warn loudly.
$exitCount = ($markers | Where-Object { $_.Kind -eq 'Exit' } | Measure-Object).Count
if ($exitCount -eq 0) {
    Write-Warning ("Log contains {0} phase Start marker(s) but no Exit markers - this looks like an in-flight or rolled-back upgrade. The trailing phase is shown as '(running)' and ends at the log's last timestamp ({1:yyyy-MM-dd HH:mm:ss})." -f $markers.Count, $lastLogTime)
}

# Coalesce consecutive same-phase segments (Downlevel restart across reboot).
$coalesced = New-Object System.Collections.Generic.List[object]
foreach ($s in $segments) {
    if ($coalesced.Count -gt 0 -and $coalesced[-1].Phase -eq $s.Phase) {
        $coalesced[-1].End = $s.End
        if ($s.HResult) { $coalesced[-1].HResult = $s.HResult }
    } else {
        $coalesced.Add($s)
    }
}

# =============================================================================
# STEP 5 - Per-segment Active = Wall - sum(idle gaps inside it).
# =============================================================================
# Idle gaps may fall fully or partially within a phase; partial overlaps are
# clipped to the phase window. Active is clamped to >= 0.
# -----------------------------------------------------------------------------
$timeline = New-Object System.Collections.Generic.List[object]
$prevEnd  = $null
foreach ($s in $coalesced) {
    if (-not $IncludeDownlevel -and $s.Phase -eq 'Downlevel') { continue }
    $wall = $s.End - $s.Start
    $idle = [TimeSpan]::Zero
    foreach ($g in $gaps) {
        if ($g.From -ge $s.Start -and $g.To -le $s.End) {
            $idle += $g.Duration
        } elseif ($g.From -lt $s.End -and $g.To -gt $s.Start) {
            # Partial overlap - clip to the phase window.
            $clipFrom = if ($g.From -gt $s.Start) { $g.From } else { $s.Start }
            $clipTo   = if ($g.To   -lt $s.End)   { $g.To   } else { $s.End }
            $idle += ($clipTo - $clipFrom)
        }
    }
    $active = $wall - $idle
    if ($active.Ticks -lt 0) { $active = [TimeSpan]::Zero }
    # Gap = inter-phase transition (reboots, OOBE handoffs). Lives between
    # OPERATIONTRACK markers and would otherwise silently inflate the span.
    $gap = if ($null -ne $prevEnd -and $s.Start -gt $prevEnd) { $s.Start - $prevEnd } else { [TimeSpan]::Zero }
    $prevEnd = $s.End
    $timeline.Add([pscustomobject]@{
        Phase    = $s.Phase
        Start    = $s.Start
        End      = $s.End
        Duration = $active
        Idle     = $idle
        Gap      = $gap
        Wall     = $wall
        HResult  = $s.HResult
    })
}

# =============================================================================
# STEP 6 - Output dispatch: -AsObject | -Csv | -TotalActiveMinutes | table.
# =============================================================================
if ($AsObject) {
    return $timeline
}

if ($Csv) {
    $timeline |
        Select-Object Phase,
            @{n='Start';     e={ $_.Start.ToString('yyyy-MM-dd HH:mm:ss') }},
            @{n='End';       e={ $_.End.ToString('yyyy-MM-dd HH:mm:ss') }},
            @{n='ActiveSec'; e={ [int]$_.Duration.TotalSeconds }},
            @{n='IdleSec';   e={ [int]$_.Idle.TotalSeconds }},
            @{n='GapSec';    e={ [int]$_.Gap.TotalSeconds }},
            @{n='WallSec';   e={ [int]$_.Wall.TotalSeconds }},
            HResult |
        ConvertTo-Csv -NoTypeInformation -Delimiter ';'
    return
}

# ---- Render -----------------------------------------------------------------
$totalSec  = ($timeline | ForEach-Object { $_.Duration.TotalSeconds } | Measure-Object -Sum).Sum
$totalIdle = [TimeSpan]::FromSeconds((($timeline | ForEach-Object { $_.Idle.TotalSeconds } | Measure-Object -Sum).Sum))
$totalGap  = [TimeSpan]::FromSeconds((($timeline | ForEach-Object { $_.Gap.TotalSeconds  } | Measure-Object -Sum).Sum))
$totalSpan = [TimeSpan]::FromSeconds($totalSec)

if ($TotalActiveMinutes) {
    [int][math]::Round($totalSpan.TotalMinutes)
    return
}

# Compact human-friendly TimeSpan formatter.
function Format-Duration([TimeSpan]$d) {
    if ($d.TotalHours   -ge 1) { return ('{0:0}h {1:00}m {2:00}s' -f [int]$d.TotalHours, $d.Minutes, $d.Seconds) }
    if ($d.TotalMinutes -ge 1) { return ('{0:0}m {1:00}s'         -f [int]$d.TotalMinutes, $d.Seconds) }
    return ('{0}s' -f [int]$d.TotalSeconds)
}

# Build rendered rows. Date prefix only shown when it changes between adjacent rows.
$prevDay = $null
$rows = foreach ($t in $timeline) {
    $startStr = if ($t.Start.Date -ne $prevDay)       { $t.Start.ToString('MM-dd HH:mm:ss') } else { '      ' + $t.Start.ToString('HH:mm:ss') }
    $endStr   = if ($t.End.Date   -ne $t.Start.Date)  { $t.End.ToString('MM-dd HH:mm:ss')   } else { '      ' + $t.End.ToString('HH:mm:ss') }
    $prevDay  = $t.Start.Date
    [pscustomobject]@{
        Phase   = $t.Phase
        Start   = $startStr
        End     = $endStr
        Active  = Format-Duration $t.Duration
        Idle    = if ($t.Idle.TotalSeconds -ge 1) { Format-Duration $t.Idle } else { '-' }
        Gap     = if ($t.Gap.TotalSeconds  -ge 1) { Format-Duration $t.Gap  } else { '-' }
        HResult = if ($t.HResult) { $t.HResult } else { '-' }
        IsError = ($t.HResult -and $t.HResult -ne '0x00000000' -and $t.HResult -ne '0x0')
    }
}

# Column widths sized for longest phase name ("Post First Boot") + timestamps.
$fmt = "  {0,-16} {1,14} {2,14} {3,11} {4,11} {5,9}  {6,-10}"
$sep = '-' * 91

# Header
Write-Host ''
Write-Host ('  Setup Timeline')                                    -ForegroundColor Cyan
Write-Host ('  Log:       {0}' -f $mainLog)                        -ForegroundColor DarkGray
Write-Host ('  Downlevel: {0}' -f ($(if ($IncludeDownlevel) { 'INCLUDED' } else { 'excluded (user productive)' }))) -ForegroundColor DarkGray
Write-Host ('  Idle gap:  > {0}s treated as off / standby / sleep' -f $IdleGapSeconds) -ForegroundColor DarkGray
Write-Host ''
Write-Host ($fmt -f 'Phase','Start','End','Active','Idle','Gap','HRESULT') -ForegroundColor Yellow
Write-Host $sep -ForegroundColor DarkGray

# Rows (failed phases highlighted red)
foreach ($r in $rows) {
    $line = $fmt -f $r.Phase, $r.Start, $r.End, $r.Active, $r.Idle, $r.Gap, $r.HResult
    if ($r.IsError) { Write-Host $line -ForegroundColor Red }
    else            { Write-Host $line }
}

# Footer
Write-Host $sep -ForegroundColor DarkGray
Write-Host ('  Active upgrade time : {0,-14}  ({1} min)' -f (Format-Duration $totalSpan), [int][math]::Round($totalSpan.TotalMinutes)) -ForegroundColor Green
Write-Host ('  Excluded idle time  : {0}' -f (Format-Duration $totalIdle)) -ForegroundColor DarkGray
Write-Host ('  Inter-phase gaps    : {0,-14}  (reboots and phase handoffs)' -f (Format-Duration $totalGap)) -ForegroundColor DarkGray

if ($timeline.Count -gt 0) {
    $first = $timeline[0].Start
    $last  = $timeline[-1].End
    Write-Host ('  Wall-clock span     : {0,-14}  ({1:yyyy-MM-dd HH:mm} -> {2:yyyy-MM-dd HH:mm})' -f `
        (Format-Duration ($last - $first)), $first, $last) -ForegroundColor DarkGray
}
Write-Host ''
