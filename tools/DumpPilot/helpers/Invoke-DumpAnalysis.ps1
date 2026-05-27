#requires -Version 7.0
<#
.SYNOPSIS
    Stage 3 of DumpPilot. Build a facts-only report + llm prompt from
    dump-summary.json (no pattern matching).

.DESCRIPTION
        This stage does not classify or match signatures. It passes deterministic
        debugger facts through to:
            - <base>.report.json
            - <base>.report.md
            - <base>.llm-prompt.md           (trimmed for paste into Gemini/ChatGPT)
            - <base>.llm-prompt.full.md      (full ProcMon/threads, for 1M-ctx models)
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$SummaryPath,
    [string]$OutputBasePath,    # default: same folder + same basename as summary
    [string]$PatternDbPath      # default: ../pattern_db.json next to this helper
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $SummaryPath)) { throw "Summary not found: $SummaryPath" }
$summary = Get-Content -LiteralPath $SummaryPath -Raw | ConvertFrom-Json

if (-not $OutputBasePath) {
    $summaryItem = Get-Item -LiteralPath $SummaryPath
    # strip .dump-summary.json / dump-summary.json
    $base = $summaryItem.BaseName -replace '\.dump-summary$',''
    $OutputBasePath = Join-Path $summaryItem.DirectoryName $base
}
$reportJson = "$OutputBasePath.report.json"
$reportMd   = "$OutputBasePath.report.md"
$promptMd   = "$OutputBasePath.llm-prompt.md"
$promptFull = "$OutputBasePath.llm-prompt.full.md"

# --- pattern DB matching -----------------------------------------------------
# Schema (per DESIGN.md §5):
#   { rules: [ { id, match:{ProcessImage,FaultingModule,ExceptionCode,StackRegex}, severity, title } ] }
# All present `match.*` regexes must hit; missing fields are wildcards.
# Reports ALL matching rules (highest severity first), not just the first.
$patternHits = New-Object System.Collections.Generic.List[object]
if (-not $PatternDbPath) {
    $PatternDbPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'pattern_db.json'
}
if (Test-Path -LiteralPath $PatternDbPath) {
    try {
        $db = Get-Content -LiteralPath $PatternDbPath -Raw | ConvertFrom-Json
        $proc  = [string]$summary.Process.Name
        $fmod  = [string]$summary.Faulting.Module
        # FaultingModule patterns in the DB include the extension (.dll/.sys/.exe);
        # the summary's Faulting.Module does not. Try the bare name AND each
        # likely-extension variant so a `.sys` driver pattern matches a kernel
        # dump whose MODULE_NAME is the bare driver name.
        $fmodVariants = New-Object System.Collections.Generic.List[string]
        if ($fmod) {
            [void]$fmodVariants.Add($fmod)
            if ($fmod -notmatch '\.(dll|sys|exe)$') {
                [void]$fmodVariants.Add($fmod + '.dll')
                [void]$fmodVariants.Add($fmod + '.sys')
                [void]$fmodVariants.Add($fmod + '.exe')
            }
        }
        $ecode = [string]$summary.Exception.Code
        $stackJoined = ''
        if ($summary.Faulting.Stack) {
            $stackJoined = (@($summary.Faulting.Stack | ForEach-Object { [string]$_.Frame }) -join "`n")
        }
        $sevRank = @{ 'Critical' = 4; 'High' = 3; 'Medium' = 2; 'Low' = 1 }
        foreach ($rule in @($db.rules)) {
            if (-not $rule -or -not $rule.match) { continue }
            $m = $rule.match
            $ok = $true
            if ($ok -and $m.ProcessImage) {
                $ok = ($proc -and [regex]::IsMatch($proc, [string]$m.ProcessImage, 'IgnoreCase'))
            }
            if ($ok -and $m.FaultingModule) {
                $rx = [string]$m.FaultingModule
                $ok = $false
                foreach ($v in $fmodVariants) {
                    if ($v -and [regex]::IsMatch($v, $rx, 'IgnoreCase')) { $ok = $true; break }
                }
            }
            if ($ok -and $m.ExceptionCode) {
                $ok = ($ecode -and [regex]::IsMatch(($ecode -replace '^0x',''), [string]$m.ExceptionCode, 'IgnoreCase'))
            }
            if ($ok -and $m.StackRegex) {
                $ok = ($stackJoined -and [regex]::IsMatch($stackJoined, [string]$m.StackRegex, 'IgnoreCase'))
            }
            if ($ok) {
                $sev = [string]$rule.severity
                $rank = 0
                if ($sevRank.ContainsKey($sev)) { $rank = [int]$sevRank[$sev] }
                [void]$patternHits.Add([pscustomobject]@{
                    Id       = [string]$rule.id
                    Title    = [string]$rule.title
                    Severity = $sev
                    SeverityRank = $rank
                })
            }
        }
        # Sort by severity desc (Critical first), preserve DB order on ties.
        if ($patternHits.Count -gt 1) {
            $sorted = $patternHits | Sort-Object -Property SeverityRank -Descending -Stable
            $patternHits = New-Object System.Collections.Generic.List[object]
            foreach ($h in $sorted) { [void]$patternHits.Add($h) }
        }
    } catch {
        Write-Verbose "pattern_db match failed: $($_.Exception.Message)"
    }
}

# --- report.json -------------------------------------------------------------

$analysisMode = if ($patternHits.Count -gt 0) { 'pattern+facts' } else { 'facts-only' }
# Materialize to a plain object[] to avoid PSEnumerableBinder edge cases
# when a `System.Collections.Generic.List[object]` is embedded directly
# inside an `[ordered]@{}` initializer alongside other typed properties.
$patternHitsArr = [object[]]@()
if ($patternHits.Count -gt 0) {
    $patternHitsArr = [object[]]@($patternHits.ToArray())
}

$report = [ordered]@{
    SchemaVersion  = '0.1'
    SummarySource  = $SummaryPath
    AnalysisMode   = $analysisMode
    GeneratedAt    = (Get-Date).ToString('o')
    Dump           = [ordered]@{
        Path     = $summary.DumpPath
        Kind     = $summary.DumpKind
        Sha256   = $summary.DumpSha256
    }
    Process        = $summary.Process
    Exception      = $summary.Exception
    Faulting       = [ordered]@{
        Module   = $summary.Faulting.Module
        Image    = $summary.Faulting.Image
        Symbol   = $summary.Faulting.Symbol
        FullStack = $summary.Faulting.Stack
    }
    Modules        = ($summary.Modules | Select-Object -First 50)
    ModuleDetails  = $summary.ModuleDetails
    ExceptionRecord = $summary.ExceptionRecord
    ContextRegisters = $summary.ContextRegisters
    Thread          = $summary.Thread
    ResourceStats   = $summary.ResourceStats
    SymbolQuality   = $summary.SymbolQuality
    Capture        = $summary.Capture
    Comparison     = $summary.Comparison
    VASummary      = $summary.VASummary
    VMStats        = $summary.VMStats
    Timing         = $summary.Timing
    LastEvent      = $summary.LastEvent
    LastError      = $summary.LastError
    ProcessInfo    = $summary.ProcessInfo
    ThreadCount    = $summary.ThreadCount
    ThreadCountReported = $summary.ThreadCountReported
    ThreadCountCaptured = $summary.ThreadCountCaptured
    ThreadCaptureTruncated = $summary.ThreadCaptureTruncated
    HandleTypes    = $summary.HandleTypes
    HeapCorrupt    = $summary.HeapCorrupt
    ThreadStacks   = $summary.ThreadStacks
    ExceptionChain = $summary.ExceptionChain
    UnresolvedHotspots = $summary.UnresolvedHotspots
    RegisterCorrelation = $summary.RegisterCorrelation
    CriticalSections = $summary.CriticalSections
    SEHChain       = $summary.SEHChain
    CLR            = $summary.CLR
    Wow64Stack     = $summary.Wow64Stack
    Token          = $summary.Token
    AppVerifier    = $summary.AppVerifier
    StackPointerDump = $summary.StackPointerDump
    Disassembly      = $summary.Disassembly
    ProcMonCorrelation = $summary.ProcMonCorrelation
    PatternHits    = $patternHitsArr
    HitCount       = $patternHits.Count
}
[System.IO.File]::WriteAllText($reportJson, ($report | ConvertTo-Json -Depth 10), [System.Text.UTF8Encoding]::new($true))

# --- report.md ---------------------------------------------------------------

$md = New-Object System.Text.StringBuilder
[void]$md.AppendLine("# DumpPilot report")
[void]$md.AppendLine()
[void]$md.AppendLine('**Dump:** `' + $summary.DumpPath + '`')
[void]$md.AppendLine('**Kind:** ' + $summary.DumpKind)
[void]$md.AppendLine('**Process:** `' + $summary.Process.Name + '`')
$codeLine = $summary.Exception.Code
if ($summary.Exception.CodeName) { $codeLine += ' (' + $summary.Exception.CodeName + ')' }
$codeLine += ' at ' + $summary.Exception.Address
[void]$md.AppendLine('**Exception:** ' + $codeLine)
[void]$md.AppendLine('**Bucket:** ' + $summary.Exception.Bucket)
[void]$md.AppendLine('**Faulting module:** ' + $summary.Faulting.Module + '  (`' + $summary.Faulting.Symbol + '`)')
[void]$md.AppendLine('**Capture commands:** ' + [string]$summary.Capture.CommandCount)
[void]$md.AppendLine()
if ($patternHits.Count -gt 0) {
    [void]$md.AppendLine('## Pattern DB hits (' + $patternHits.Count + ')')
    [void]$md.AppendLine()
    [void]$md.AppendLine('| # | Severity | Id | Title |')
    [void]$md.AppendLine('|---|---|---|---|')
    $idx = 0
    foreach ($h in $patternHits) {
        $idx++
        [void]$md.AppendLine('| ' + $idx + ' | ' + [string]$h.Severity + ' | `' + [string]$h.Id + '` | ' + [string]$h.Title + ' |')
    }
    [void]$md.AppendLine()
}
[void]$md.AppendLine('## Command line')
[void]$md.AppendLine()
[void]$md.AppendLine('```')
[void]$md.AppendLine([string]$summary.Process.CommandLine)
[void]$md.AppendLine('```')
[void]$md.AppendLine()
[void]$md.AppendLine('## Full faulting stack')
[void]$md.AppendLine()
[void]$md.AppendLine('```')
foreach ($f in $summary.Faulting.Stack) {
    [void]$md.AppendLine($f.Frame)
}
[void]$md.AppendLine('```')
[void]$md.AppendLine()
[void]$md.AppendLine('## Exception record')
[void]$md.AppendLine()
[void]$md.AppendLine('- Flags: ' + [string]$summary.ExceptionRecord.Flags)
[void]$md.AppendLine('- NumberParameters: ' + [string]$summary.ExceptionRecord.NumberParameters)
[void]$md.AppendLine('- Parameters: ' + ((@($summary.ExceptionRecord.Parameters) -join ', ')))
[void]$md.AppendLine()
[void]$md.AppendLine('## Thread/context')
[void]$md.AppendLine()
[void]$md.AppendLine('- Thread: ' + [string]$summary.Thread.Address + '  Cid: ' + [string]$summary.Thread.Cid + '  Teb: ' + [string]$summary.Thread.Teb)
[void]$md.AppendLine('- State: ' + [string]$summary.Thread.State + '  WaitReason: ' + [string]$summary.Thread.WaitReason)
[void]$md.AppendLine('- ThreadCount: ' + [string]$summary.ThreadCount)
if ($summary.ThreadCaptureTruncated) {
    [void]$md.AppendLine('- **WARNING:** thread-stack capture truncated ' +
        '(reported=' + [string]$summary.ThreadCountReported +
        ', captured=' + [string]$summary.ThreadCountCaptured + ').' +
        ' cdb output buffer overflow; re-run with `cdb -lines -srcpath` ' +
        'or split `~* kb 8` into smaller batches.')
}
[void]$md.AppendLine()
[void]$md.AppendLine('## Symbol quality')
[void]$md.AppendLine()
[void]$md.AppendLine('- Resolved stack frames: ' + [string]$summary.SymbolQuality.StackFramesResolved + ' / ' + [string]$summary.SymbolQuality.StackFramesTotal)
[void]$md.AppendLine('- Unresolved stack frames: ' + [string]$summary.SymbolQuality.StackFramesUnresolved)
[void]$md.AppendLine()
[void]$md.AppendLine('## Process context')
[void]$md.AppendLine()
if ($summary.ProcessInfo) {
    [void]$md.AppendLine('- Exe: ' + [string]$summary.ProcessInfo.ExeImagePath)
    [void]$md.AppendLine('- WorkDir: ' + [string]$summary.ProcessInfo.CurrentDirectory)
}
if ($summary.LastError -and ($summary.LastError.Code -or $summary.LastError.String)) {
    [void]$md.AppendLine('- Last error: ' + [string]$summary.LastError.Code + '  "' + [string]$summary.LastError.String + '"')
}
if ($summary.LastEvent) { [void]$md.AppendLine('- Last event: ' + [string]$summary.LastEvent) }
if ($summary.Timing) {
    [void]$md.AppendLine('- Kernel time: ' + [string]$summary.Timing.KernelTime + '  User time: ' + [string]$summary.Timing.UserTime)
}
if ($summary.HeapCorrupt) { [void]$md.AppendLine('- **HEAP CORRUPTION detected in raw output**') }
[void]$md.AppendLine()

# Exception chain
if ($summary.ExceptionChain -and @($summary.ExceptionChain).Count -gt 1) {
    [void]$md.AppendLine('## Exception chain (' + @($summary.ExceptionChain).Count + ' records)')
    [void]$md.AppendLine()
    foreach ($ec in $summary.ExceptionChain) {
        [void]$md.AppendLine('- Code: ' + [string]$ec.Code + '  Addr: ' + [string]$ec.Address + '  Flags: ' + [string]$ec.Flags)
    }
    [void]$md.AppendLine()
}

# Per-thread stacks
if ($summary.ThreadStacks -and @($summary.ThreadStacks).Count -gt 0) {
    [void]$md.AppendLine('## All thread stacks (' + @($summary.ThreadStacks).Count + ' threads)')
    [void]$md.AppendLine()
    foreach ($ts in $summary.ThreadStacks) {
        [void]$md.AppendLine('### Thread ' + [string]$ts.ThreadId)
        [void]$md.AppendLine('```')
        foreach ($fr in @($ts.Frames)) { [void]$md.AppendLine([string]$fr) }
        [void]$md.AppendLine('```')
        [void]$md.AppendLine()
    }
}

# Unresolved hotspots
if ($summary.UnresolvedHotspots -and @($summary.UnresolvedHotspots).Count -gt 0) {
    [void]$md.AppendLine('## Unresolved symbol hotspots')
    [void]$md.AppendLine()
    [void]$md.AppendLine('| Module | Unresolved | Total | % |')
    [void]$md.AppendLine('|---|---|---|---|')
    foreach ($uh in $summary.UnresolvedHotspots) {
        [void]$md.AppendLine('| ' + [string]$uh.Module + ' | ' + [string]$uh.UnresolvedCount + ' | ' + [string]$uh.TotalInStack + ' | ' + [string]$uh.Pct + '% |')
    }
    [void]$md.AppendLine()
}

# Register correlation
if ($summary.RegisterCorrelation -and @($summary.RegisterCorrelation).Count -gt 0) {
    [void]$md.AppendLine('## Register-to-module correlation')
    [void]$md.AppendLine()
    [void]$md.AppendLine('| Register | Value | Module |')
    [void]$md.AppendLine('|---|---|---|')
    foreach ($rc in $summary.RegisterCorrelation) {
        $modLabel = if ($rc.Module) { [string]$rc.Module } else { '**UNMAPPED**' }
        [void]$md.AppendLine('| ' + [string]$rc.Register + ' | ' + [string]$rc.Value + ' | ' + $modLabel + ' |')
    }
    [void]$md.AppendLine()
}

# Critical section locks
if ($summary.CriticalSections -and @($summary.CriticalSections).Count -gt 0) {
    [void]$md.AppendLine('## Critical section locks (' + @($summary.CriticalSections).Count + ')')
    [void]$md.AppendLine()
    [void]$md.AppendLine('| Lock | Address | Owner Thread | LockCount | Recursion |')
    [void]$md.AppendLine('|---|---|---|---|---|')
    foreach ($cs in $summary.CriticalSections) {
        [void]$md.AppendLine('| ' + [string]$cs.Name + ' | ' + [string]$cs.Address + ' | ' + [string]$cs.OwningThread + ' | ' + [string]$cs.LockCount + ' | ' + [string]$cs.RecursionCount + ' |')
    }
    [void]$md.AppendLine()
}

# SEH exception chain
if ($summary.SEHChain -and @($summary.SEHChain).Count -gt 0) {
    [void]$md.AppendLine('## SEH exception handler chain (' + @($summary.SEHChain).Count + ' handlers)')
    [void]$md.AppendLine()
    foreach ($seh in $summary.SEHChain) {
        [void]$md.AppendLine('- ' + [string]$seh.Address + ': ' + [string]$seh.Handler)
    }
    [void]$md.AppendLine()
}

# CLR managed stack
if ($summary.CLR) {
    $hasClr = ($summary.CLR.Runtime -or (@($summary.CLR.ManagedStack).Count -gt 0))
    if ($hasClr) {
        [void]$md.AppendLine('## CLR / .NET runtime')
        [void]$md.AppendLine()
        if ($summary.CLR.Runtime) {
            [void]$md.AppendLine('- Runtime: **' + [string]$summary.CLR.Runtime + '** v' + [string]$summary.CLR.Version)
        }
        if (@($summary.CLR.ManagedStack).Count -gt 0) {
            [void]$md.AppendLine()
            [void]$md.AppendLine('### Managed stack (!clrstack)')
            [void]$md.AppendLine('```')
            foreach ($mf in $summary.CLR.ManagedStack) { [void]$md.AppendLine([string]$mf) }
            [void]$md.AppendLine('```')
        }
        if (@($summary.CLR.Threads).Count -gt 0) {
            [void]$md.AppendLine()
            [void]$md.AppendLine('### CLR threads (' + @($summary.CLR.Threads).Count + ')')
            [void]$md.AppendLine()
            [void]$md.AppendLine('| # | OSID | Lock | State |')
            [void]$md.AppendLine('|---|---|---|---|')
            foreach ($ct in ($summary.CLR.Threads | Select-Object -First 20)) {
                [void]$md.AppendLine('| ' + [string]$ct.Index + ' | ' + [string]$ct.OSID + ' | ' + [string]$ct.LockCount + ' | ' + [string]$ct.State + ' |')
            }
        }
        [void]$md.AppendLine()
    }
}

# Wow64 x86 stack
if ($summary.Wow64Stack -and @($summary.Wow64Stack).Count -gt 0) {
    [void]$md.AppendLine('## Wow64 x86 stack (' + @($summary.Wow64Stack).Count + ' frames)')
    [void]$md.AppendLine()
    [void]$md.AppendLine('```')
    foreach ($wf in $summary.Wow64Stack) { [void]$md.AppendLine([string]$wf) }
    [void]$md.AppendLine('```')
    [void]$md.AppendLine()
}

# Security token
if ($summary.Token -and $summary.Token.User) {
    [void]$md.AppendLine('## Security context')
    [void]$md.AppendLine()
    [void]$md.AppendLine('- User: **' + [string]$summary.Token.User + '**')
    if ($summary.Token.IntegrityLevel) { [void]$md.AppendLine('- Integrity: ' + [string]$summary.Token.IntegrityLevel) }
    if (@($summary.Token.Privileges).Count -gt 0) {
        [void]$md.AppendLine('- Privileges (' + @($summary.Token.Privileges).Count + '):')
        foreach ($p in ($summary.Token.Privileges | Select-Object -First 15)) {
            [void]$md.AppendLine('  - ' + [string]$p)
        }
    }
    [void]$md.AppendLine()
}

# App Verifier
if ($summary.AppVerifier -and $summary.AppVerifier.Active) {
    [void]$md.AppendLine('## App Verifier')
    [void]$md.AppendLine()
    if ($summary.AppVerifier.StopCode) {
        [void]$md.AppendLine('- **Stop: ' + [string]$summary.AppVerifier.StopCode + '**')
        if ($summary.AppVerifier.Description) {
            [void]$md.AppendLine('- Description: ' + [string]$summary.AppVerifier.Description)
        }
    } else {
        [void]$md.AppendLine('- App Verifier is active but no stop was triggered.')
    }
    [void]$md.AppendLine()
}

[void]$md.AppendLine('## Virtual memory')
[void]$md.AppendLine()
if ($summary.VASummary) {
    [void]$md.AppendLine('| Region | Size |')
    [void]$md.AppendLine('|---|---|')
    [void]$md.AppendLine('| Free | ' + [string]$summary.VASummary.Free + ' |')
    [void]$md.AppendLine('| Heap | ' + [string]$summary.VASummary.Heap + ' |')
    [void]$md.AppendLine('| Image | ' + [string]$summary.VASummary.Image + ' |')
    [void]$md.AppendLine('| Stack | ' + [string]$summary.VASummary.Stack + ' |')
    [void]$md.AppendLine('| Unknown | ' + [string]$summary.VASummary.Unknown + ' |')
}
[void]$md.AppendLine()
[void]$md.AppendLine('## Debugger commands executed')
[void]$md.AppendLine()
if ($summary.Capture.Commands -and $summary.Capture.Commands.Count -gt 0) {
    foreach ($c in $summary.Capture.Commands) {
        [void]$md.AppendLine('- ' + [string]$c)
    }
} else {
    [void]$md.AppendLine('_No command list available in summary._')
}
[void]$md.AppendLine()
[System.IO.File]::WriteAllText($reportMd, $md.ToString(), [System.Text.UTF8Encoding]::new($true))

# --- llm-prompt.md -----------------------------------------------------------

$prompt = New-Object System.Text.StringBuilder
[void]$prompt.AppendLine('# DumpPilot LLM prompt')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('## System')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('You are a senior Windows escalation engineer performing crash triage.')
if ($patternHits.Count -gt 0) {
    [void]$prompt.AppendLine('Use the extracted dump facts in the JSON below. A pattern-DB lookup ran first and the matched rules are listed in `PatternHits` (sorted by severity).')
    [void]$prompt.AppendLine('Treat the top pattern hit as a STRONG prior, but still confirm it against the stack/registers/modules before locking in the root cause. Do not invent details that aren''t in the JSON.')
} else {
    [void]$prompt.AppendLine('Use ONLY the extracted dump facts in the JSON below. No signature database was used.')
}
[void]$prompt.AppendLine('Do not invent module names, HRESULTs, versions, or known issues.')
[void]$prompt.AppendLine('If a fact is missing, say so explicitly and explain why it matters.')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('## TriageMindset')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('- Prefer evidence over intuition.')
[void]$prompt.AppendLine('- The dump is incomplete evidence. Treat it that way.')
[void]$prompt.AppendLine('- Distinguish between **faulting location** (where execution failed) and **originating cause** (what led to the failure). Do NOT assume they are the same. Prefer upstream causes when evidence exists.')
[void]$prompt.AppendLine('- Do NOT assume system modules (ntdll, kernel32, user32, gdi32) or GPU drivers are the root cause without eliminating upstream callers first.')
[void]$prompt.AppendLine('- Consider earlier corruption vs immediate failure: a NULL deref in ntdll may be caused by a caller passing a bad pointer.')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('## ReasoningRules')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('### Corruption detection')
[void]$prompt.AppendLine('Evaluate for signs of memory corruption:')
[void]$prompt.AppendLine('- NULL or near-NULL register values used as pointers (rdx=0, rcx=0, etc.)')
[void]$prompt.AppendLine('- Duplicate or repeating stack frame sequences (loop, recursion, or capture artifact)')
[void]$prompt.AppendLine('- Large number of unresolved frames (check UnresolvedHotspots and SymbolQuality)')
[void]$prompt.AppendLine('- Execution in unexpected modules or at addresses outside loaded module ranges (see RegisterCorrelation)')
[void]$prompt.AppendLine('If corruption evidence exists, explicitly state it and downgrade confidence.')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('### Thread analysis')
[void]$prompt.AppendLine('Analyze ThreadStacks to:')
[void]$prompt.AppendLine('- Identify the crashing thread role (UI, worker, IO, GC, finalizer, etc.)')
[void]$prompt.AppendLine('- Check other threads for blocking/wait patterns, deadlocks, or abnormal clustering')
[void]$prompt.AppendLine('- Note if CriticalSections shows held locks and which threads own them')
[void]$prompt.AppendLine('Use thread context to strengthen or weaken the root cause hypothesis.')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('### Subsystem classification')
[void]$prompt.AppendLine('Identify the subsystem involved (e.g., graphics/GPU, UI/windowing, networking, JVM, file IO, .NET GC, COM/RPC).')
[void]$prompt.AppendLine('Classify the faulting module vendor: 1st-party Microsoft, 3rd-party application (Java, Chrome, etc.), or hardware vendor (Intel, NVIDIA, AMD). Treat vendor drivers as distinct from core OS modules.')
[void]$prompt.AppendLine('Use this classification to guide root cause hypotheses and next actions.')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('### Symbol name inference')
[void]$prompt.AppendLine('Use function and symbol names in the stack to infer what the code was doing when it crashed.')
[void]$prompt.AppendLine('Symbol names often reveal the operation in progress (e.g., DumpRegistryKeyDefinitions implies a registry read,')
[void]$prompt.AppendLine('SetPixelFormat implies graphics context setup, MonitorWait implies lock acquisition).')
[void]$prompt.AppendLine('If a symbol name suggests a specific operation, incorporate that into the root cause hypothesis.')
[void]$prompt.AppendLine('Do NOT treat exported symbol names as definitive — they may be the nearest export, not the actual function.')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('### Symbol quality impact')
[void]$prompt.AppendLine('Incorporate SymbolQuality and UnresolvedHotspots into confidence:')
[void]$prompt.AppendLine('- High unresolved ratio = lower confidence in stack-based reasoning')
[void]$prompt.AppendLine('- Modules with unresolved hotspots are uncertainty sources — say so')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('### Regression tracking')
[void]$prompt.AppendLine('Check the Comparison block. If StackSigChanged is false, this is a recurring crash — state that and adjust hypotheses toward systemic causes.')
[void]$prompt.AppendLine('If ModuleDelta shows added/removed/changed modules, correlate with the crash timing.')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('### Resource exhaustion')
[void]$prompt.AppendLine('Review VASummary and ResourceStats. If free virtual address space is critically low, heap committed is near reserved,')
[void]$prompt.AppendLine('or HeapCorrupt is true, explicitly state whether an OOM or heap exhaustion condition contributed to the failure.')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('### Security product interference')
[void]$prompt.AppendLine('Scan the Modules list for non-Microsoft, non-application security products (AV, EDR, DLP). Common indicators:')
[void]$prompt.AppendLine('modules from CrowdStrike, SentinelOne, Carbon Black, Symantec, McAfee, Trend Micro, Sophos, ESET, Palo Alto, Cybereason.')
[void]$prompt.AppendLine('If such modules are loaded and the crash involves a callback exception or hook-susceptible API, hypothesize whether security hooking could be intercepting or blocking the faulting operation.')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('### ProcMon correlation')
[void]$prompt.AppendLine('If ProcMonCorrelation is present in the facts, it contains registry/file/network failures from Process Monitor captured around the crash time.')
[void]$prompt.AppendLine('The FaultRelated array lists failures whose paths match the faulting module vendor or graphics/driver registry keys.')
[void]$prompt.AppendLine('Use these to confirm or refute hypotheses about missing configuration, denied access, or absent registry keys.')
[void]$prompt.AppendLine('This is external evidence — treat it as high-value corroboration. If NAME NOT FOUND or ACCESS DENIED appears for driver config keys, it likely explains a NULL pointer in the driver.')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('### Confidence grading')
[void]$prompt.AppendLine('- **High**: Direct evidence in stack + consistent register/thread context + minimal ambiguity')
[void]$prompt.AppendLine('- **Medium**: Plausible cause but alternative explanations exist or key data is missing')
[void]$prompt.AppendLine('- **Low**: Only faulting location known; origin unclear, symbols poor, or dump type limits analysis')
[void]$prompt.AppendLine('Default to Medium or Low unless strong proof exists. Do NOT inflate confidence.')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('## Constraints')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('- Do NOT equate faulting module with root cause automatically.')
[void]$prompt.AppendLine('- Do NOT invent version, OS, or driver data not present in the facts.')
[void]$prompt.AppendLine('- Do NOT claim "known issues" or cite KBs without evidence in the dump facts.')
[void]$prompt.AppendLine('- Do NOT suggest generic steps (update drivers, reinstall). Tie every action to a specific missing fact or ambiguity.')
[void]$prompt.AppendLine('- If a 3rd-party or vendor module faults and its FileVersion/Timestamp is missing, the FIRST next action must be to obtain that version. Never suggest "update the driver" without first establishing what version is running.')
[void]$prompt.AppendLine('- If a register is NULL (e.g. rdx=0) and Disassembly data is available, state whether the instruction at the faulting offset reads from, writes to, or executes via that register.')
[void]$prompt.AppendLine('- If the Disassembly array is non-empty, you MUST reference it in the Evidence section. Do NOT claim disassembly is missing when the field contains data.')
[void]$prompt.AppendLine('- Use ONLY values from the JSON. Do NOT hallucinate counts, thread numbers, or versions that differ from the facts.')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('## Facts (JSON)')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('```json')
# Snapshot the prologue (everything up to and including the opening ```json
# fence) so we can rebuild a `.llm-prompt.full.md` with the SAME prologue
# but the un-trimmed report JSON. Avoids drift between the two prompts.
$promptPrologue = $prompt.ToString()

# Build a trimmed clone of $report just for the LLM prompt. The full
# version stays in *.report.json (tooling, UI, regression tests). Aim
# is to fit Gemini / ChatGPT paste limits (~64-100 KB) without losing
# signal. ProcMon dominates raw size; everything else is small but
# we still drop empty/noise fields to keep the JSON readable.
#
# Trims:
#   1. ProcMonCorrelation.FaultRelated -> top 25 by Count, summary stub
#   2. ProcMonCorrelation.OtherFailures -> dropped entirely
#   3. ThreadStacks -> collapse identical stacks into groups
#   4. Modules -> only those that appear in faulting stack OR ModuleDetails
#   5. Empty optional blocks -> dropped (CLR/SEHChain/Token/AppVerifier/...)
#   6. Capture.Commands/CdbPath/SymbolPath -> dropped (noise for LLM)

$promptReport = [ordered]@{}
foreach ($k in $report.Keys) { $promptReport[$k] = $report[$k] }

# (1)(2) ProcMon trim
if ($promptReport.ProcMonCorrelation) {
    try {
        $pmc = $promptReport.ProcMonCorrelation
    $faultCount = if ($pmc.FaultRelated) { @($pmc.FaultRelated).Count } else { 0 }

    # (1pre) OS-chatter denylist. Some paths sneak into FaultRelated because
    # of incidental substring matches in the parser's relevance regex (most
    # notoriously `d3d` matching GUIDs like 02F815B5-…-D3D8 in the OS Power
    # Manager keys). Those rows are never the cause of a graphics-driver
    # crash and they swamp the top-25 with hundreds of NAME NOT FOUND probes
    # that mislead the LLM. Drop them up front so they don't even compete
    # for the consolidation slots.
    $denyPatterns = @(
        '\\Control\\Power\\PowerSettings\\',
        '\\Control\\Power\\User\\PowerSchemes\\',
        '\\Windows\\OneSettings\b',
        '\\Microsoft\\Windows\\WER\\Temp\b',
        '\\Microsoft\\Windows\\WER\\ReportArchive\b',
        '\\Microsoft\\Windows\\WER\\ReportQueue\b',
        '\\AppRepository\\StateRepository-',
        '\\Windows\\AppRepository\\Packages\\Microsoft\\.UI\\.Xaml',
        '\\Windows Defender\\Scans\\History\\',
        '\\Windows Defender Advanced Threat Protection\\Cache\\',
        '\\IntuneManagementExtension\\Logs\\',
        '\\WinSxS\\amd64_microsoft-windows-hotpatches_'
    )
    $denyRx = ($denyPatterns -join '|')
    $faultRelatedFiltered = @()
    $denyCount = 0
    if ($faultCount -gt 0) {
        foreach ($r in @($pmc.FaultRelated)) {
            if (([string]$r.Path) -match $denyRx) { $denyCount++; continue }
            $faultRelatedFiltered += $r
        }
    }
    # registry key. Without this, a single noisy poll loop (e.g. Intel
    # ipfsrv probes 8 sibling files in one folder, or Power-Manager
    # enumerates dozens of leaf values under one GUID-keyed parent)
    # eats the top-25. Threshold = 3 leaves: 2 siblings often carries
    # distinct signal; 3+ is virtually always the same loop.
    $folded = New-Object System.Collections.Generic.List[object]
    if ($faultRelatedFiltered.Count -gt 0) {
        $byParent = @{}
        foreach ($r in $faultRelatedFiltered) {
            $rawPath = [string]$r.NormPath
            if (-not $rawPath) { $rawPath = [string]$r.Path }
            # Split on \ or /; drop leaf segment. Skip if no separator
            # (e.g. bare HKLM root, or single-segment path) — pass through.
            $sep = [Math]::Max($rawPath.LastIndexOf('\'), $rawPath.LastIndexOf('/'))
            $parent = if ($sep -gt 0) { $rawPath.Substring(0, $sep) } else { $rawPath }
            $leaf   = if ($sep -gt 0 -and $sep -lt ($rawPath.Length - 1)) { $rawPath.Substring($sep + 1) } else { '' }
            $key = "$($r.Operation)|$parent|$($r.Result)"
            if (-not $byParent.ContainsKey($key)) {
                $byParent[$key] = [pscustomobject]@{
                    Operation = $r.Operation
                    Parent    = $parent
                    Result    = $r.Result
                    Count     = 0
                    LeafCount = 0
                    Examples  = New-Object System.Collections.Generic.List[string]
                    Rows      = New-Object System.Collections.Generic.List[object]
                }
            }
            $entry = $byParent[$key]
            $entry.Count += [int]$r.Count
            $entry.LeafCount++
            if ($leaf -and $entry.Examples.Count -lt 3 -and -not $entry.Examples.Contains($leaf)) {
                [void]$entry.Examples.Add($leaf)
            }
            [void]$entry.Rows.Add($r)
        }
        foreach ($key in $byParent.Keys) {
            $g = $byParent[$key]
            if ($g.LeafCount -ge 3) {
                # Synthetic consolidated row. No NormPath: the consolidated
                # path already contains the parent + `* (N files)` summary.
                [void]$folded.Add([pscustomobject]@{
                    Operation = $g.Operation
                    Path      = "$($g.Parent)\* ($($g.LeafCount) files)"
                    Result    = $g.Result
                    Count     = $g.Count
                    LeafCount = $g.LeafCount
                    Examples  = [object[]]$g.Examples.ToArray()
                })
            } else {
                # Keep originals as-is. Use .ToArray() to dodge the
                # PSEnumerableBinder edge case where iterating a
                # List[object] stored inside a PSCustomObject NoteProperty
                # throws "Argument types do not match".
                # Drop NormPath (redundant with Path) and RelevantToFault
                # (always true for everything in FaultRelated).
                foreach ($r in $g.Rows.ToArray()) {
                    [void]$folded.Add([pscustomobject]@{
                        Operation = $r.Operation
                        Path      = $r.Path
                        Result    = $r.Result
                        Count     = $r.Count
                    })
                }
            }
        }
    }
    $foldedAll = @($folded | Sort-Object -Property Count -Descending)
    $top = @($foldedAll | Select-Object -First 25)

    $pmcTrim = [ordered]@{
        WindowSeconds                  = $pmc.WindowSeconds
        FailureCount                   = $pmc.FailureCount
        UniqueFailures                 = $pmc.UniqueFailures
        FaultRelated                   = $top
        FaultRelatedTruncatedFrom      = $faultCount
        FaultRelatedAfterConsolidation = $foldedAll.Count
        FaultRelatedDroppedAsOsChatter = $denyCount
        OtherFailuresDropped      = if ($pmc.OtherFailures) { @($pmc.OtherFailures).Count } else { 0 }
    }
    $promptReport.ProcMonCorrelation = $pmcTrim
    } catch {
        Write-Warning "ProcMon trim failed; passing through un-trimmed payload. $($_.Exception.Message)"
    }
}

# (3) ThreadStacks collapse: group by joined-frames signature.
if ($promptReport.ThreadStacks) {
    $groups = @{}
    $order  = New-Object System.Collections.Generic.List[string]
    foreach ($ts in @($promptReport.ThreadStacks)) {
        $frames = @($ts.Frames)
        $sig = ($frames -join ' > ')
        if (-not $groups.ContainsKey($sig)) {
            $groups[$sig] = [ordered]@{
                Count     = 0
                ThreadIds = New-Object System.Collections.Generic.List[string]
                Frames    = $frames
            }
            [void]$order.Add($sig)
        }
        $groups[$sig].Count++
        [void]$groups[$sig].ThreadIds.Add([string]$ts.ThreadId)
    }
    $collapsed = New-Object System.Collections.Generic.List[object]
    foreach ($sig in $order) {
        $g = $groups[$sig]
        # Cap ThreadIds list at 8 to avoid spamming the prompt when one
        # signature has 50+ threads (jvm wait loop).
        $tids = @($g.ThreadIds)
        if ($tids.Count -gt 8) {
            $tids = (@($tids[0..7]) + @("...+$($tids.Count - 8) more"))
        }
        [void]$collapsed.Add([pscustomobject]@{
            Count     = $g.Count
            ThreadIds = $tids
            Frames    = $g.Frames
        })
    }
    # Sort by descending Count so the most common waiter sigs appear first.
    $promptReport.ThreadStacks = @($collapsed | Sort-Object -Property Count -Descending)
}

# (4) Modules trim: keep only modules that touch the stack or have details.
if ($promptReport.Modules) {
    $keep = @{}
    if ($promptReport.Faulting -and $promptReport.Faulting.FullStack) {
        foreach ($f in @($promptReport.Faulting.FullStack)) {
            if ($f.Module) { $keep[[string]$f.Module] = $true }
        }
    }
    if ($promptReport.ModuleDetails) {
        foreach ($md in @($promptReport.ModuleDetails)) {
            if ($md.Name) { $keep[[string]$md.Name] = $true }
        }
    }
    if ($keep.Count -gt 0) {
        $modsAll = @($promptReport.Modules)
        $modsKept = @($modsAll | Where-Object { $_.Name -and $keep.ContainsKey([string]$_.Name) })
        $promptReport.Modules = $modsKept
        $promptReport.ModulesDroppedCount = $modsAll.Count - $modsKept.Count
    }
}

# (5) Drop empty optional blocks (don't list as missing — just omit).
$emptyOptionals = @(
    'CLR','SEHChain','Token','Wow64Stack',
    'CriticalSections','StackPointerDump','Comparison'
)
foreach ($k in $emptyOptionals) {
    if ($promptReport.Contains($k)) {
        $v = $promptReport[$k]
        $isEmpty = $false
        if ($null -eq $v) { $isEmpty = $true }
        elseif ($v -is [System.Collections.IDictionary]) {
            $isEmpty = ($v.Count -eq 0)
            if (-not $isEmpty) {
                $hasVal = $false
                foreach ($pv in $v.Values) {
                    if ($null -ne $pv -and -not ($pv -is [array] -and $pv.Count -eq 0)) { $hasVal = $true; break }
                }
                $isEmpty = -not $hasVal
            }
        }
        elseif ($v -is [pscustomobject]) {
            $hasVal = $false
            foreach ($p in $v.PSObject.Properties) {
                if ($null -ne $p.Value -and -not ($p.Value -is [array] -and $p.Value.Count -eq 0)) { $hasVal = $true; break }
            }
            $isEmpty = -not $hasVal
        }
        elseif ($v -is [array] -or $v -is [System.Collections.IEnumerable]) {
            $isEmpty = (@($v).Count -eq 0)
        }
        if ($isEmpty) { $promptReport.Remove($k) }
    }
}

# AppVerifier: special case — drop when Active=false (the entire block is
# meaningful only when verifier was actually running).
if ($promptReport.AppVerifier -and -not $promptReport.AppVerifier.Active) {
    $promptReport.Remove('AppVerifier')
}

# (6) Strip Capture noise. Drop entirely from the prompt -- duration and
# command count carry no LLM signal; full data still lives in report.json.
if ($promptReport.Contains('Capture')) { $promptReport.Remove('Capture') | Out-Null }

# (7) PatternHits hygiene: drop the internal SeverityRank sort key.
if ($promptReport.PatternHits) {
    $cleanHits = New-Object System.Collections.Generic.List[object]
    foreach ($h in @($promptReport.PatternHits)) {
        [void]$cleanHits.Add([pscustomobject]@{
            Id       = $h.Id
            Title    = $h.Title
            Severity = $h.Severity
        })
    }
    $promptReport.PatternHits = [object[]]$cleanHits.ToArray()
}

# (8) SymbolQuality.ModuleCoverage: drop entries with Frames=0. These are
# stub rows produced by `lmv m clr / coreclr / mscorlib_ni / System_ni`
# unconditionally issued by the capture bundle even on non-.NET dumps.
if ($promptReport.SymbolQuality -and $promptReport.SymbolQuality.ModuleCoverage) {
    $coverageKept = New-Object System.Collections.Generic.List[object]
    foreach ($mc in @($promptReport.SymbolQuality.ModuleCoverage)) {
        if ([int]$mc.Frames -gt 0) { [void]$coverageKept.Add($mc) }
    }
    # Rebuild SymbolQuality so the consumer-facing JSON has the filtered list.
    $sqOld = $promptReport.SymbolQuality
    $promptReport.SymbolQuality = [ordered]@{
        StackFramesTotal      = $sqOld.StackFramesTotal
        StackFramesResolved   = $sqOld.StackFramesResolved
        StackFramesUnresolved = $sqOld.StackFramesUnresolved
        ModuleCoverage        = [object[]]$coverageKept.ToArray()
    }
}

# (9) ExceptionChain dedup. cdb's .exr -1 reports the same exception twice
# (chance-1 + chance-2) on user-mode AVs. Collapse runs of identical
# (Address, Code) and drop the field entirely if only one unique record
# remains -- the Exception block already carries it.
if ($promptReport.ExceptionChain) {
    $unique = New-Object System.Collections.Generic.List[object]
    $lastKey = $null
    foreach ($ec in @($promptReport.ExceptionChain)) {
        $key = "$($ec.Address)|$($ec.Code)|$($ec.Flags)|$($ec.NumParams)"
        if ($key -ne $lastKey) { [void]$unique.Add($ec); $lastKey = $key }
    }
    if ($unique.Count -le 1) {
        $promptReport.Remove('ExceptionChain') | Out-Null
    } else {
        $promptReport.ExceptionChain = [object[]]$unique.ToArray()
    }
}

# (10) ModuleDetails: only keep entries for modules that actually appear in
# the faulting stack. The capture bundle does `lmv m kernel32 / clr / ...`
# unconditionally; the parser then promotes them into ModuleDetails even
# when they contributed zero frames. Surfacing kernel32's version when no
# kernel32 frame faulted just gives the LLM an extra detail to over-fit to.
if ($promptReport.ModuleDetails -and $promptReport.Faulting -and $promptReport.Faulting.FullStack) {
    $stackMods = @{}
    foreach ($f in @($promptReport.Faulting.FullStack)) {
        if ($f.Module) { $stackMods[[string]$f.Module] = $true }
    }
    if ($stackMods.Count -gt 0) {
        $mdKept = @($promptReport.ModuleDetails | Where-Object {
            $_.Name -and $stackMods.ContainsKey([string]$_.Name)
        })
        $promptReport.ModuleDetails = [object[]]$mdKept
    }
}

# (11) ThreadCount cleanup. Keep only the two fields that actually mean
# something in user-mode dumps; drop the cdb-kernel-mode "Reported" field
# which is virtually always 0 here and just confuses readers.
if ($promptReport.Contains('ThreadCountReported')) { $promptReport.Remove('ThreadCountReported') | Out-Null }

# (12) Provenance / housekeeping fields the LLM never uses.
foreach ($k in @('SchemaVersion','SummarySource','GeneratedAt')) {
    if ($promptReport.Contains($k)) { $promptReport.Remove($k) | Out-Null }
}
# Drop the disk paths from the Dump block -- LLM can't use them.
if ($promptReport.Dump) {
    $promptReport.Dump = [ordered]@{
        Kind   = $promptReport.Dump.Kind
        Sha256 = $promptReport.Dump.Sha256
    }
}

[void]$prompt.AppendLine(($promptReport | ConvertTo-Json -Depth 10))
[void]$prompt.AppendLine('```')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('## Task')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('Produce a triage report in concise markdown with EXACTLY these section headers in this order.')
[void]$prompt.AppendLine('Do NOT reorganize, merge, rename, or skip any section. Do NOT add extra sections.')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('### 1. RootCause')
[void]$prompt.AppendLine('One-sentence root cause distinguishing faulting location from originating cause.')
[void]$prompt.AppendLine('State the affected subsystem (graphics, UI, networking, JVM, etc.).')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('### 2. Confidence')
[void]$prompt.AppendLine('High / Medium / Low (using the grading rules above).')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('### 3. Evidence')
[void]$prompt.AppendLine('Bullet list citing exact JSON fields, frame names, register values, and thread IDs.')
[void]$prompt.AppendLine('Include corruption indicators if any. Note symbol quality impact.')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('### 4. EvidenceQuality')
[void]$prompt.AppendLine('- Stack resolution coverage (from SymbolQuality)')
[void]$prompt.AppendLine('- Missing critical fields that limit analysis')
[void]$prompt.AppendLine('- Dump type limitation impact (minidump vs full)')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('### 5. NextActions')
[void]$prompt.AppendLine('Ranked by: (a) ability to confirm/refute root cause quickly, (b) reduction of highest uncertainty, (c) ease of collection.')
[void]$prompt.AppendLine('Each action must reference a specific missing fact or ambiguity. No generic advice.')
[void]$prompt.AppendLine()
[void]$prompt.AppendLine('### 6. Hypotheses')
[void]$prompt.AppendLine('Each hypothesis must:')
[void]$prompt.AppendLine('- Be explicitly labeled as a hypothesis')
[void]$prompt.AppendLine('- Be tied to specific evidence from the dump')
[void]$prompt.AppendLine('- Include a validation step (how to prove or disprove it)')
[System.IO.File]::WriteAllText($promptMd, $prompt.ToString(), [System.Text.UTF8Encoding]::new($true))

# .llm-prompt.full.md — same prologue + Task block, but with the un-trimmed
# report JSON. Use this when the trimmed prompt loses signal needed for a
# specific investigation (e.g. you actually want to inspect every ProcMon
# failure, or every thread's wait stack). Same path / same model context
# requirements as the trimmed one, just larger; expect >1 MB on dumps with
# heavy ProcMon correlation.
$promptSuffix = $prompt.ToString().Substring($promptPrologue.Length)
# $promptSuffix starts with the trimmed JSON; find where the JSON block ends
# (the closing ``` fence after our ConvertTo-Json output) and replace
# the trimmed JSON with the full report.
$jsonEndIdx = $promptSuffix.IndexOf('```')
if ($jsonEndIdx -ge 0) {
    $taskBlock = $promptSuffix.Substring($jsonEndIdx)   # "```\n\n## Task\n..."
    $fullPrompt = $promptPrologue + ($report | ConvertTo-Json -Depth 10) + [Environment]::NewLine + $taskBlock
    [System.IO.File]::WriteAllText($promptFull, $fullPrompt, [System.Text.UTF8Encoding]::new($true))
}

[pscustomobject]@{
    ReportJson    = $reportJson
    ReportMd      = $reportMd
    PromptMd      = $promptMd
    PromptFullMd  = $promptFull
    HitCount      = $patternHits.Count
    TopHit        = if ($patternHits.Count -gt 0) { [string]$patternHits[0].Id } else { $null }
}
