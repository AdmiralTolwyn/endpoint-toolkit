#requires -Version 5.1
<#
.SYNOPSIS
    Stage 3 of DumpPilot. Build a facts-only report + llm prompt from
    dump-summary.json (no pattern matching).

.DESCRIPTION
        This stage does not classify or match signatures. It passes deterministic
        debugger facts through to:
            - <base>.report.json
            - <base>.report.md
            - <base>.llm-prompt.md
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$SummaryPath,
    [string]$OutputBasePath  # default: same folder + same basename as summary
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

# --- report.json -------------------------------------------------------------

$report = [ordered]@{
    SchemaVersion  = '0.1'
    SummarySource  = $SummaryPath
    AnalysisMode   = 'facts-only'
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
    PatternHits    = @()
    HitCount       = 0
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
[void]$prompt.AppendLine('Use ONLY the extracted dump facts in the JSON below. No signature database was used.')
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
[void]$prompt.AppendLine(($report | ConvertTo-Json -Depth 10))
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

[pscustomobject]@{
    ReportJson = $reportJson
    ReportMd   = $reportMd
    PromptMd   = $promptMd
    HitCount   = 0
    TopHit     = $null
}
