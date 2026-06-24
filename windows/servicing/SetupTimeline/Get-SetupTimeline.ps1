<#
.SYNOPSIS
    Builds a phase timeline of a Windows in-place upgrade from setupact.log.

.DESCRIPTION
    Parses OPERATIONTRACK markers ("Start execution phase <Name>") and phase exit
    markers ("Execution phase [<Name>] exiting with HRESULT [...]") to compute how
    long each upgrade phase actually took. Uses a compiled C# helper (StreamReader
    + regex) so multi-GB setupact.log files are processed in seconds.

    By default the ONLINE phases (Pre-Downlevel, Downlevel, Pre-Finalize) are
    excluded because they run inside the still-booted source OS - the user can
    keep working during them. Pass -IncludeDownlevel to keep them. The remaining
    (OFFLINE) phases - including Finalize, which shows the full-screen
    "restarting" UI - run after the user is locked out of the desktop.
    Every row is tagged 'online' or 'offline' in the Mode column.

    A synthetic 'Pre-Downlevel' lead-in segment is added from the log's first
    timestamp to the first OPERATIONTRACK marker. setupact.log begins logging at
    SetupHost launch (compatibility scan, dynamic update / ESD download, WinRE
    servicing) well before the first formal phase marker, so anchoring here lets
    the timeline reflect the whole upgrade experience rather than just the
    phase-bracketed window. It is tagged 'online' and follows the same
    -IncludeDownlevel gating as the other online phases.

    If a setupact_unattendGC.log sits next to setupact.log, its WinDeploy/
    OOBE timestamps are used to extend the timeline through the OOBE phase.

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
    Include the online phases (Pre-Downlevel lead-in, Downlevel, Pre-Finalize)
    in the timeline. Off by default - these phases run in the source OS while
    the user is still productive, so excluding them gives the "lockout time"
    number that IT typically wants to quote. Name kept for backward compat.

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

.PARAMETER ShowDownloads
    Append a Dynamic Update / download breakdown beneath the timeline. Parses
    the DCAT transfer markers in setupact.log (Transferring file / Transfer
    progress) to report how many files were fetched, the size of the largest
    payloads (LCU, FODs, language packs), the download span, and the
    instantaneous throughput (avg / peak / min) - plus how long the transfer
    spent below 2 Mbps. The download window is almost always the bulk of the
    online 'Pre-Downlevel' lead-in, so this answers "was it the download or the
    apply that took so long?". Applies to the default report only (ignored with
    -AsObject / -Csv / -TotalActiveMinutes).

.EXAMPLE
    .\Get-SetupTimeline.ps1 -LogPath C:\temp\TK\setupact_MX-PF5041LA

.EXAMPLE
    # Auto-discover - run on the just-upgraded machine itself.
    .\Get-SetupTimeline.ps1 -Verbose

.EXAMPLE
    .\Get-SetupTimeline.ps1 -LogPath C:\Windows\Panther -IncludeDownlevel

.EXAMPLE
    # Break down the Dynamic Update download (size, throughput, slow-link time):
    .\Get-SetupTimeline.ps1 -LogPath C:\Windows\Panther -IncludeDownlevel -ShowDownloads

.EXAMPLE
    $mins = .\Get-SetupTimeline.ps1 -LogPath C:\temp\setupact.log -TotalActiveMinutes

.NOTES
    Author     : Anton Romanyuk
    Version    : 1.3.0
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

    # Include the online phases (Pre-Downlevel lead-in, Downlevel, Pre-Finalize).
    # Off by default - they run in the source OS while the user is still
    # productive. Name kept (rather than -IncludeOnlinePhases) for backward compat.
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
    [switch]$Csv,

    # Append a Dynamic Update / download breakdown (file count, payload sizes,
    # throughput, slow-link time) beneath the timeline. Parses DCAT transfer
    # markers; applies to the default report only.
    [switch]$ShowDownloads
)

$ErrorActionPreference = 'Stop'

# Bump $ScriptVersion when behaviour or output format changes.
$ScriptVersion = '1.3.0'
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
Write-Verbose ("Online phases: {0} (Pre-Downlevel, Downlevel, Pre-Finalize)" -f ($(if ($IncludeDownlevel) { 'INCLUDED' } else { 'excluded' })))

# =============================================================================
# STEP 2 - Compile the C# scanner once per AppDomain.
# =============================================================================
# A pure-PS parser is 50-100x slower on multi-hundred-MB setupact.log files.
# The 'as [type]' guard avoids re-compilation on subsequent invocations.
# NB: the namespace carries a version suffix (…V4). Bump it whenever the C#
# shape changes - Add-Type cannot replace an already-loaded type in a
# long-lived session (e.g. PS 5.1 console), so a new namespace forces the
# updated class to load instead of silently reusing the stale one.
# -----------------------------------------------------------------------------
if (-not ('SetupTimelineV4.Scanner' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace SetupTimelineV4 {
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

    // One Dynamic Update file transfer start ("Transferring file from url ... to ...").
    public class Transfer {
        public DateTime Time;
        public string   Name;   // destination leaf
        public string   Url;
    }

    // One DCAT transfer progress sample ("Transfer: [done / total] [pct%] [inst][avg]").
    public class Progress {
        public DateTime Time;
        public long     Done;
        public long     Total;
        public int      Pct;
        public long     Inst;   // instantaneous bytes/s
        public long     Avg;    // reported average bytes/s
    }

    public class ScanResult {
        public List<Marker>  Markers  = new List<Marker>();
        public List<GapInfo> Gaps     = new List<GapInfo>();
        // First timestamped line - anchors the synthetic 'Pre-Downlevel' lead-in
        // (SetupHost / compat scan / download) that precedes the first phase marker.
        public DateTime      FirstTime = DateTime.MinValue;
        // Last timestamped line - used to close trailing open phases on
        // in-flight upgrades (no "exiting with HRESULT" written yet).
        public DateTime      LastTime = DateTime.MinValue;
        public int           LineCount;
        // Dynamic Update download activity (populated only when captureDownloads=true).
        public List<Transfer>  Transfers       = new List<Transfer>();
        public List<Progress>  Progress        = new List<Progress>();
        public HashSet<string> MediaVersions   = new HashSet<string>();
        public long            DuCategoryBytes = 0;
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
        // Dynamic Update download markers (only matched when captureDownloads=true).
        private static readonly Regex _xferStart = new Regex(
            @"FCDCATHelper: Transferring file from url \[(.*?)\] to \[(.*?)\]",
            RegexOptions.Compiled);
        private static readonly Regex _xferProg = new Regex(
            @"FCDCATHelper: Transfer: \[0x([0-9A-Fa-f]+) / 0x([0-9A-Fa-f]+)\] \[(\d+)%\] \[(\d+) bytes/s\] \[(\d+) Avg bytes/s\]",
            RegexOptions.Compiled);
        private static readonly Regex _duReq = new Regex(
            @"FCAcquirerDCAT: Given DU Category requires \[(\d+)\] bytes",
            RegexOptions.Compiled);
        private static readonly Regex _media = new Regex(
            @"""MediaVersion""\s*:\s*""([^""]+)""",
            RegexOptions.Compiled);

        // Single-pass scan. unattendGc=true captures only WinDeploy markers.
        public static ScanResult Scan(string path, bool unattendGc, int idleGapSeconds, bool captureDownloads) {
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
                    if (result.FirstTime == DateTime.MinValue) result.FirstTime = ts;
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

                    // Dynamic Update download capture - cheap substring prefilter first
                    // so the per-line regex cost is only paid on DCAT / session lines.
                    if (captureDownloads &&
                        (line.IndexOf("DCAT", StringComparison.Ordinal) >= 0 ||
                         line.IndexOf("MediaVersion", StringComparison.Ordinal) >= 0)) {
                        var xpr = _xferProg.Match(line);
                        if (xpr.Success) {
                            result.Progress.Add(new Progress {
                                Time  = ts,
                                Done  = Convert.ToInt64(xpr.Groups[1].Value, 16),
                                Total = Convert.ToInt64(xpr.Groups[2].Value, 16),
                                Pct   = int.Parse(xpr.Groups[3].Value),
                                Inst  = long.Parse(xpr.Groups[4].Value),
                                Avg   = long.Parse(xpr.Groups[5].Value)
                            });
                        } else {
                            var xst = _xferStart.Match(line);
                            if (xst.Success) {
                                result.Transfers.Add(new Transfer {
                                    Time = ts, Url = xst.Groups[1].Value,
                                    Name = System.IO.Path.GetFileName(xst.Groups[2].Value)
                                });
                            } else {
                                var dq = _duReq.Match(line);
                                if (dq.Success) {
                                    long b; if (long.TryParse(dq.Groups[1].Value, out b)) result.DuCategoryBytes += b;
                                } else {
                                    var mv = _media.Match(line);
                                    if (mv.Success) result.MediaVersions.Add(mv.Groups[1].Value);
                                }
                            }
                        }
                    }

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
$scan    = [SetupTimelineV4.Scanner]::Scan($mainLog, $false, $IdleGapSeconds, $ShowDownloads.IsPresent)
# Strongly-typed list lets us .Add() / [-1] index across PS 5.1 and PS 7.
$markers = [System.Collections.Generic.List[object]]::new()
foreach ($m in $scan.Markers) { [void]$markers.Add($m) }
$gaps        = $scan.Gaps
$lastLogTime = $scan.LastTime   # used to close trailing in-flight phases
$firstLogTime = $scan.FirstTime # anchors the synthetic Pre-Downlevel lead-in
$sw.Stop()
$mbPerSec = if ($sw.Elapsed.TotalSeconds -gt 0) { ($mainSize / 1MB) / $sw.Elapsed.TotalSeconds } else { 0 }
Write-Verbose ("Scanned setupact.log in {0:N1}s ({1:N1} MB/s) - {2} phase markers, {3} idle gaps" -f `
    $sw.Elapsed.TotalSeconds, $mbPerSec, $markers.Count, $gaps.Count)
if ($ShowDownloads) {
    Write-Verbose ("Download capture: {0} file transfers, {1} progress samples" -f $scan.Transfers.Count, $scan.Progress.Count)
}

# Merge OOBE log into the same marker/gap pools so segment-builder sees one stream.
if ($unattendLog -and (Test-Path -LiteralPath $unattendLog)) {
    $sw2 = [System.Diagnostics.Stopwatch]::StartNew()
    $wd  = [SetupTimelineV4.Scanner]::Scan($unattendLog, $true, $IdleGapSeconds, $false)
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

# Synthetic 'Pre-Downlevel' lead-in: the window from the log's first timestamp
# to the first OPERATIONTRACK marker. setupact.log starts logging at SetupHost
# launch (compat scan, dynamic update / ESD download, WinRE servicing) well
# before the first formal phase marker is written. Anchoring here lets the
# report reflect the whole upgrade experience, not just the phase-bracketed
# window. Tagged 'online' (runs in the source OS, user still productive) so it
# follows the same -IncludeDownlevel gating as the other online phases; its
# internal idle gaps are clipped/subtracted by the per-segment loop below.
if ($coalesced.Count -gt 0 -and $firstLogTime -gt [datetime]::MinValue -and $firstLogTime -lt $coalesced[0].Start) {
    $coalesced.Insert(0, [pscustomobject]@{
        Phase   = 'Pre-Downlevel'
        Start   = $firstLogTime
        End     = $coalesced[0].Start
        HResult = $null
    })
}

# =============================================================================
# STEP 5 - Per-segment Active = Wall - sum(idle gaps inside it).
# =============================================================================
# Idle gaps may fall fully or partially within a phase; partial overlaps are
# clipped to the phase window. Active is clamped to >= 0.
# -----------------------------------------------------------------------------
$timeline = New-Object System.Collections.Generic.List[object]
$prevEnd  = $null
# Phases that run inside the still-booted source OS (user is productive).
# Everything else runs in WinRE / Safe OS / first boot / OOBE - user locked out.
# Finalize is offline because Setup shows its full-screen "restarting" UI
# and the user can no longer interact with the desktop.
$onlinePhases = @('Pre-Downlevel','Downlevel','Pre-Finalize')
foreach ($s in $coalesced) {
    if (-not $IncludeDownlevel -and $onlinePhases -contains $s.Phase) { continue }
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
        Mode     = if ($onlinePhases -contains $s.Phase) { 'online' } else { 'offline' }
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
            Mode,
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
# NB: use [Math]::Floor for the leading bucket, NOT [int] - PowerShell's
# [int] cast does banker's rounding (e.g. [int]4.6 = 5), which inflates
# h/m and makes the row total disagree with what the user adds up.
function Format-Duration([TimeSpan]$d) {
    if ($d.TotalHours   -ge 1) { return ('{0:0}h {1:00}m {2:00}s' -f [int][Math]::Floor($d.TotalHours),   $d.Minutes, $d.Seconds) }
    if ($d.TotalMinutes -ge 1) { return ('{0:0}m {1:00}s'         -f [int][Math]::Floor($d.TotalMinutes), $d.Seconds) }
    return ('{0}s' -f [int][Math]::Floor($d.TotalSeconds))
}

# Build rendered rows. Date prefix only shown when it changes between adjacent rows.
$prevDay = $null
$rows = foreach ($t in $timeline) {
    $startStr = if ($t.Start.Date -ne $prevDay)       { $t.Start.ToString('MM-dd HH:mm:ss') } else { '      ' + $t.Start.ToString('HH:mm:ss') }
    $endStr   = if ($t.End.Date   -ne $t.Start.Date)  { $t.End.ToString('MM-dd HH:mm:ss')   } else { '      ' + $t.End.ToString('HH:mm:ss') }
    $prevDay  = $t.Start.Date
    [pscustomobject]@{
        Phase   = $t.Phase
        Mode    = $t.Mode
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
$fmt = "  {0,-16} {1,-7} {2,14} {3,14} {4,11} {5,11} {6,9}  {7,-10}"
$sep = '-' * 100

# Header
Write-Host ''
Write-Host ('  Setup Timeline')                                    -ForegroundColor Cyan
Write-Host ('  Log:        {0}' -f $mainLog)                       -ForegroundColor DarkGray
Write-Host ('  Online:     {0}' -f ($(if ($IncludeDownlevel) { 'INCLUDED (Pre-Downlevel, Downlevel, Pre-Finalize)' } else { 'excluded (Pre-Downlevel, Downlevel, Pre-Finalize - user productive)' }))) -ForegroundColor DarkGray
Write-Host ('  Idle gap:   > {0}s treated as off / standby / sleep' -f $IdleGapSeconds) -ForegroundColor DarkGray
Write-Host ''
Write-Host ($fmt -f 'Phase','Mode','Start','End','Active','Idle','Gap','HRESULT') -ForegroundColor Yellow
Write-Host $sep -ForegroundColor DarkGray

# Rows (failed phases highlighted red)
foreach ($r in $rows) {
    $line = $fmt -f $r.Phase, $r.Mode, $r.Start, $r.End, $r.Active, $r.Idle, $r.Gap, $r.HResult
    if ($r.IsError) { Write-Host $line -ForegroundColor Red }
    else            { Write-Host $line }
}

# Footer
Write-Host $sep -ForegroundColor DarkGray
Write-Host ('  Active upgrade time : {0,-14}  ({1} min)' -f (Format-Duration $totalSpan), [int][Math]::Floor($totalSpan.TotalMinutes)) -ForegroundColor Green
Write-Host ('  Excluded idle time  : {0}' -f (Format-Duration $totalIdle)) -ForegroundColor DarkGray
Write-Host ('  Inter-phase gaps    : {0,-14}  (reboots and phase handoffs)' -f (Format-Duration $totalGap)) -ForegroundColor DarkGray

if ($timeline.Count -gt 0) {
    $first = $timeline[0].Start
    $last  = $timeline[-1].End
    Write-Host ('  Wall-clock span     : {0,-14}  ({1:yyyy-MM-dd HH:mm} -> {2:yyyy-MM-dd HH:mm})' -f `
        (Format-Duration ($last - $first)), $first, $last) -ForegroundColor DarkGray
}
Write-Host ''

# =============================================================================
# STEP 7 - Optional Dynamic Update / download breakdown (-ShowDownloads).
# =============================================================================
# The DCAT transfer markers tell us whether the (usually dominant) online
# lead-in was spent DOWNLOADING the DU payload or APPLYING it. Progress lines
# carry the payload size + instantaneous/avg byte rate; transfer-start lines
# name each file. Grouping progress samples by their Total size separates the
# big payloads (LCU, ESD) from the many tiny FOD / language-pack cabs.
# -----------------------------------------------------------------------------
if ($ShowDownloads) {
    $dlTransfers = $scan.Transfers
    $dlProgress  = $scan.Progress
    $dlMedia     = @($scan.MediaVersions)
    $dlCatBytes  = [double]$scan.DuCategoryBytes

    Write-Host ('  Dynamic Update / Downloads') -ForegroundColor Cyan
    Write-Host $sep -ForegroundColor DarkGray

    if (($dlTransfers.Count -eq 0) -and ($dlProgress.Count -eq 0)) {
        Write-Host '  No Dynamic Update download activity found.' -ForegroundColor DarkGray
        Write-Host '  (DU may be disabled / pre-staged, or this is not a downlevel setupact.log.)' -ForegroundColor DarkGray
        Write-Host ''
    } else {
        # bytes -> human; bytes/s -> Mbps (decimal megabits, the link-speed convention).
        function Format-Bytes([double]$b) {
            if ($b -ge 1GB) { return ('{0:N2} GB' -f ($b / 1GB)) }
            if ($b -ge 1MB) { return ('{0:N1} MB' -f ($b / 1MB)) }
            if ($b -ge 1KB) { return ('{0:N0} KB' -f ($b / 1KB)) }
            return ('{0:N0} B' -f $b)
        }
        # bytes/s -> link-speed string. Drops to Kbps below 1 Mbps so a stalled
        # transfer (e.g. 2.6 KB/s) shows '21 Kbps', not a misleading '0.0 Mbps'.
        function Format-Rate([double]$bps) {
            $mbps = $bps * 8 / 1e6
            if ($mbps -ge 1)    { return ('{0:N1} Mbps' -f $mbps) }
            if ($bps  -ge 125)  { return ('{0:N0} Kbps' -f ($bps * 8 / 1e3)) }
            return ('{0:N0} bps' -f ($bps * 8))
        }

        # Group progress samples into payloads: a change in Total size = new file.
        $payloads = New-Object System.Collections.Generic.List[object]
        $curTotal = -1; $grp = $null
        foreach ($p in $dlProgress) {
            if ([long]$p.Total -ne [long]$curTotal) {
                if ($grp) { $payloads.Add($grp) }
                $curTotal = [long]$p.Total
                $grp = [pscustomobject]@{ Total = [long]$p.Total; Start = $p.Time; End = $p.Time; Name = '(unknown)' }
            }
            $grp.End = $p.Time
        }
        if ($grp) { $payloads.Add($grp) }

        # Attach the nearest preceding transfer filename + true start time.
        foreach ($pl in $payloads) {
            $src = $dlTransfers | Where-Object { $_.Time -le $pl.Start } | Select-Object -Last 1
            if ($src) {
                $pl.Name = $src.Name
                if ($src.Time -lt $pl.Start) { $pl.Start = $src.Time }
            }
        }

        $allInst    = @($dlProgress | Where-Object { $_.Inst -gt 0 } | ForEach-Object { [double]$_.Inst })
        $sumPayload = ($payloads | Measure-Object Total -Sum).Sum
        $largest    = $payloads | Sort-Object Total -Descending | Select-Object -First 1

        # Download span = first transfer start -> last progress sample.
        $dlStart = if ($dlTransfers.Count -gt 0) { ($dlTransfers | Select-Object -First 1).Time }
                   elseif ($payloads.Count -gt 0) { $payloads[0].Start } else { $null }
        $dlEnd   = if ($dlProgress.Count -gt 0) { ($dlProgress | Select-Object -Last 1).Time }
                   elseif ($dlTransfers.Count -gt 0) { ($dlTransfers | Select-Object -Last 1).Time } else { $null }

        # Idle inside the download span (device asleep mid-download).
        $dlIdle = [TimeSpan]::Zero
        if ($dlStart -and $dlEnd) {
            foreach ($g in $gaps) {
                if ($g.From -lt $dlEnd -and $g.To -gt $dlStart) {
                    $cf = if ($g.From -gt $dlStart) { $g.From } else { $dlStart }
                    $ct = if ($g.To   -lt $dlEnd)   { $g.To   } else { $dlEnd }
                    $dlIdle += ($ct - $cf)
                }
            }
        }

        if ($dlMedia.Count -gt 0) {
            Write-Host ('  DU media version    : {0}' -f ($dlMedia -join ', '))
        }
        Write-Host ('  Files transferred   : {0}' -f $dlTransfers.Count)
        if ($sumPayload -gt 0) {
            Write-Host ('  Measured payload    : {0}   (largest single {1})' -f (Format-Bytes $sumPayload), (Format-Bytes $largest.Total))
        } elseif ($dlCatBytes -gt 0) {
            Write-Host ('  DU category size    : {0}' -f (Format-Bytes $dlCatBytes))
        }
        if ($dlStart -and $dlEnd) {
            $wall = $dlEnd - $dlStart
            $act  = $wall - $dlIdle; if ($act.Ticks -lt 0) { $act = [TimeSpan]::Zero }
            Write-Host ('  Download span       : {0:HH:mm:ss} -> {1:HH:mm:ss}   ({2} wall, {3} active)' -f `
                $dlStart, $dlEnd, (Format-Duration $wall), (Format-Duration $act))
        }
        if ($allInst.Count -gt 0) {
            $mn = ($allInst | Measure-Object -Minimum).Minimum
            $av = ($allInst | Measure-Object -Average).Average
            $mx = ($allInst | Measure-Object -Maximum).Maximum
            Write-Host ('  Throughput (inst)   : avg {0}   peak {1}   min {2}' -f (Format-Rate $av), (Format-Rate $mx), (Format-Rate $mn))
        }

        # Slow-link diagnostic: wall time the transfer spent below ~2 Mbps.
        $slowThresh = 250000.0  # bytes/s ~= 2 Mbps
        $slow = [TimeSpan]::Zero
        for ($i = 1; $i -lt $dlProgress.Count; $i++) {
            $a = $dlProgress[$i - 1]; $b = $dlProgress[$i]
            if ([long]$a.Total -eq [long]$b.Total -and $b.Inst -gt 0 -and $b.Inst -lt $slowThresh) {
                $d = $b.Time - $a.Time
                if ($d.TotalSeconds -ge 0 -and $d.TotalSeconds -le 600) { $slow += $d }
            }
        }
        if ($slow.TotalSeconds -ge 30) {
            Write-Host ('  Slow-link (<2 Mbps) : {0} of the transfer ran below 2 Mbps' -f (Format-Duration $slow)) -ForegroundColor Yellow
        }

        # Largest payloads table (size / duration / avg rate).
        $topN = $payloads | Sort-Object Total -Descending | Select-Object -First 6
        if ($topN.Count -gt 0) {
            Write-Host ''
            Write-Host ('  {0,-52} {1,10} {2,11} {3,12}' -f 'Largest payloads', 'Size', 'Duration', 'Avg rate') -ForegroundColor Yellow
            foreach ($pl in $topN) {
                $dur = $pl.End - $pl.Start
                $avg = if ($dur.TotalSeconds -gt 0) { $pl.Total / $dur.TotalSeconds } else { 0 }
                $nm  = if ($pl.Name.Length -gt 52) { $pl.Name.Substring(0, 49) + '...' } else { $pl.Name }
                Write-Host ('  {0,-52} {1,10} {2,11} {3,12}' -f $nm, (Format-Bytes $pl.Total), (Format-Duration $dur), (Format-Rate $avg))
            }
        }
        Write-Host ''
    }
}
