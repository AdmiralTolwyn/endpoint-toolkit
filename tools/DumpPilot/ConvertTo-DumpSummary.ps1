#requires -Version 7.0
<#
.SYNOPSIS
    Stage 2 of DumpPilot. Parse the per-dump folder produced by
    Export-DumpFacts.ps1 into a structured dump-summary.json.

.DESCRIPTION
    Reads raw.txt + analyze.txt + stack.txt + modules.txt + peb.txt and emits
    a single JSON suitable for:
      - the classifier / pattern-DB matcher (stage 3)
      - direct consumption by an LLM prompt

    Extracted fields:
      Process.Name / Process.CommandLine
      Exception.Code / Exception.Address / Exception.Bucket
      Faulting.Module / Faulting.Symbol / Faulting.Stack[]
      Modules[]        (Base, End, Name, Version, Path)
      OsBuild
      DumpKind         (echoed from meta.json)

.PARAMETER FactsRoot
    Folder produced by Export-DumpFacts.ps1.

.PARAMETER OutputPath
    Default: <FactsRoot>\dump-summary.json
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$FactsRoot,
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $FactsRoot)) { throw "FactsRoot not found: $FactsRoot" }
if (-not $OutputPath) { $OutputPath = Join-Path $FactsRoot 'dump-summary.json' }

function _Read([string]$Name) {
    $p = Join-Path $FactsRoot $Name
    if (Test-Path -LiteralPath $p) { return Get-Content -LiteralPath $p -Raw }
    return ''
}

$raw      = _Read 'raw.txt'
$analyze  = $raw
$stackTxt = $raw
$modTxt   = $raw
$pebTxt   = $raw

$meta = $null
$metaPath = Join-Path $FactsRoot 'meta.json'
if (Test-Path -LiteralPath $metaPath) {
    try { $meta = Get-Content -LiteralPath $metaPath -Raw | ConvertFrom-Json } catch {}
}

function _Match1([string]$Text, [string]$Pattern) {
    $m = [regex]::Match($Text, $Pattern)
    if ($m.Success) { return $m.Groups[1].Value.Trim() }
    return $null
}

function _MatchAll([string]$Text, [string]$Pattern) {
    return [regex]::Matches($Text, $Pattern)
}

function New-OrderedFromHashtable {
    param([hashtable]$Map)
    $o = [ordered]@{}
    foreach ($k in ($Map.Keys | Sort-Object)) { $o[$k] = $Map[$k] }
    return [pscustomobject]$o
}

# --- exception / bucket -------------------------------------------------------

# NTSTATUS / exception code lookup (loaded from ntstatus_codes.json)
$ntstatusNames = @{}
$ntstatusPath = Join-Path (Split-Path -Parent $PSCommandPath) 'ntstatus_codes.json'
if (Test-Path -LiteralPath $ntstatusPath) {
    try {
        $ntDb = Get-Content -LiteralPath $ntstatusPath -Raw | ConvertFrom-Json
        foreach ($prop in $ntDb.codes.PSObject.Properties) {
            $ntstatusNames[$prop.Name] = [string]$prop.Value
        }
    } catch { }
}

$excCode    = _Match1 $analyze 'ExceptionCode:\s*([0-9a-fA-Fx]+)'
if (-not $excCode) { $excCode = _Match1 $analyze 'EXCEPTION_CODE_STR:\s*(\S+)' }
$excCodeName = $null
if ($excCode) {
    $codeLower = ($excCode -replace '^0x','').ToLowerInvariant()
    if ($ntstatusNames.Contains($codeLower)) { $excCodeName = $ntstatusNames[$codeLower] }
}
$excAddr    = _Match1 $analyze 'ExceptionAddress:\s*([0-9a-fA-F`]+)'
$bucket     = _Match1 $analyze 'FAILURE_BUCKET_ID:\s*(\S+)'
$processNm  = _Match1 $analyze 'PROCESS_NAME:\s*(\S+)'
$faultMod   = _Match1 $analyze 'MODULE_NAME:\s*(\S+)'
$faultImg   = _Match1 $analyze 'IMAGE_NAME:\s*(\S+)'
$faultSym   = _Match1 $analyze 'SYMBOL_NAME:\s*(\S+)'
$osBuild    = _Match1 $raw     'Windows .* Kernel Version (\d+)'

# --- exception record details (.exr -1) --------------------------------------

$excFlags = _Match1 $raw 'ExceptionFlags:\s*([0-9a-fA-Fx`]+)'
$excNumParams = _Match1 $raw 'NumberParameters:\s*(\d+)'
$excParams = New-Object System.Collections.Generic.List[string]
foreach ($m in (_MatchAll $raw 'Parameter\[(\d+)\]:\s*([0-9a-fA-Fx`]+)')) {
    [void]$excParams.Add($m.Groups[2].Value)
}

# --- register snapshot (r) ---------------------------------------------------

$registers = @{}
foreach ($m in (_MatchAll $raw '\b([er]?(?:ax|bx|cx|dx|si|di|sp|bp|ip)|r\d{1,2}|efl|cs|ss|ds|es|fs|gs)=([0-9a-fA-F`]+)\b')) {
    $registers[$m.Groups[1].Value] = $m.Groups[2].Value
}

# --- faulting stack (kb 30) ---------------------------------------------------

$stack = New-Object System.Collections.Generic.List[object]
foreach ($line in ($stackTxt -split "`r?`n")) {
    # 00000083`2b5fe9d0 00007fff`7200d209 : ... : modulename!symbol+0xNN
    $m = [regex]::Match($line, '^[0-9a-fA-F`]+\s+[0-9a-fA-F`]+\s+:.*?:\s*(.+)$')
    if ($m.Success) {
        $frame = $m.Groups[1].Value.Trim()
        if ($frame) {
            $modName = $null
            $symName = $null
            $mm = [regex]::Match($frame, '^([^!+\s]+)(?:!([^+\s]+))?')
            if ($mm.Success) {
                $modName = $mm.Groups[1].Value
                if ($mm.Groups[2].Success) { $symName = $mm.Groups[2].Value }
            }
            [void]$stack.Add([pscustomobject]@{
                Frame  = $frame
                Module = $modName
                Symbol = $symName
            })
        }
    }
    if ($stack.Count -ge 100) { break }
}

# Dedup back-to-back duplicate stack runs. The cdb command bundle issues
# BOTH `kv 60` and `kb 100` against the faulting thread; both emit the same
# frame format and our line-by-line collector picks them all up. Result: the
# crashing stack appears twice. Detect this and keep only the first copy.
# Algorithm: find the SECOND occurrence of $stack[0].Frame; if the run
# [0..N-2nd) matches [2nd..end), it's a confirmed duplicate emission.
# Compare on a normalised form (strip kv-only `(TrapFrame @ ADDR)` and
# `WARNING:` annotations) so kv and kb frames compare equal.
function _NormFrame([string]$Frame) {
    $s = $Frame
    $s = [regex]::Replace($s, '\s*\(TrapFrame @ [0-9a-fA-F`]+\)', '')
    $s = [regex]::Replace($s, '\s*\(CONTEXT @ [0-9a-fA-F`]+\)', '')
    $s = [regex]::Replace($s, '\s*WARNING:.*$', '')
    return $s.Trim()
}
if ($stack.Count -ge 4) {
    $firstFrame = _NormFrame $stack[0].Frame
    $second = -1
    for ($i = 1; $i -lt $stack.Count; $i++) {
        if ((_NormFrame $stack[$i].Frame) -eq $firstFrame) { $second = $i; break }
    }
    if ($second -gt 0 -and ($second * 2) -le $stack.Count) {
        $isDup = $true
        for ($i = 0; $i -lt $second; $i++) {
            if ((_NormFrame $stack[$i].Frame) -ne (_NormFrame $stack[$i + $second].Frame)) {
                $isDup = $false; break
            }
        }
        if ($isDup) {
            while ($stack.Count -gt $second) { $stack.RemoveAt($stack.Count - 1) }
        }
    }
}

# --- normalized stack signature (for diffing runs) ---------------------------

$sigParts = New-Object System.Collections.Generic.List[string]
foreach ($f in ($stack | Select-Object -First 12)) {
    $sig = if ($f.Symbol) { "{0}!{1}" -f $f.Module, $f.Symbol } else { [string]$f.Frame }
    $sig = ($sig -replace '\+0x[0-9a-fA-F]+','').Trim()
    if ($sig) { [void]$sigParts.Add($sig.ToLowerInvariant()) }
}
$stackSignature = ($sigParts -join ' | ')

# --- per-thread mini stacks (~* kb 8) ----------------------------------------

$threadStacks = New-Object System.Collections.Generic.List[object]
# Each thread block in ~* kb starts with a line like:  "   N  Id: HHHH.HHHH Suspend: N Teb: HHHH"
# or ".  N  Id: ..." for the current thread.
$threadBlocks = [regex]::Matches($raw, '(?ms)^\s+\d+\s+Id:\s+([0-9a-fA-F]+\.[0-9a-fA-F]+)[^\n]*\n((?:(?!\s+\d+\s+Id:).)*?)(?=\s+\d+\s+Id:|\z)')
foreach ($tb in $threadBlocks) {
    $tid = $tb.Groups[1].Value
    $tFrames = New-Object System.Collections.Generic.List[string]
    foreach ($fl in ($tb.Groups[2].Value -split "`r?`n")) {
        # kb format: "retaddr : args : module!symbol+0xNN"
        $fm = [regex]::Match($fl, ':\s+(\S+![^\s]+)\s*$')
        if (-not $fm.Success) {
            # fallback: kv format "childEBP retaddr : args : module!symbol"
            $fm = [regex]::Match($fl, '^[0-9a-fA-F`]+\s+[0-9a-fA-F`]+\s+:.*?:\s*(.+)$')
        }
        if ($fm.Success) {
            $fr = $fm.Groups[1].Value.Trim()
            if ($fr) { [void]$tFrames.Add($fr) }
        }
    }
    if ($tFrames.Count -gt 0) {
        [void]$threadStacks.Add([pscustomobject]@{
            ThreadId = $tid
            Frames   = @($tFrames)
        })
    }
}

# --- exception chain (.exr -1 nested records) --------------------------------

$excChain = New-Object System.Collections.Generic.List[object]
$excBlocks = [regex]::Matches($raw, '(?ms)ExceptionAddress:\s*([0-9a-fA-F`]+).*?ExceptionCode:\s*([0-9a-fA-Fx]+).*?(?:ExceptionFlags:\s*([0-9a-fA-Fx`]+))?.*?(?:NumberParameters:\s*(\d+))?')
foreach ($eb in $excBlocks) {
    [void]$excChain.Add([pscustomobject]@{
        Address = $eb.Groups[1].Value
        Code    = $eb.Groups[2].Value
        Flags   = if ($eb.Groups[3].Success) { $eb.Groups[3].Value } else { $null }
        NumParams = if ($eb.Groups[4].Success) { $eb.Groups[4].Value } else { $null }
    })
}

# --- unresolved-frame hotspots -----------------------------------------------
# Group unresolved frames (no symbol) by module and rank by frequency
# Only count real module names (no raw hex addresses)
$unresolvedHotspots = New-Object System.Collections.Generic.List[object]
$unresolvedByMod = @{}
foreach ($f in $stack) {
    if (-not $f.Symbol -and $f.Module -and $f.Module -match '^[A-Za-z]') {
        if (-not $unresolvedByMod.ContainsKey($f.Module)) { $unresolvedByMod[$f.Module] = 0 }
        $unresolvedByMod[$f.Module]++
    }
}
foreach ($k in ($unresolvedByMod.Keys | Sort-Object { $unresolvedByMod[$_] } -Descending)) {
    $totalInStack = @($stack | Where-Object { $_.Module -eq $k }).Count
    $pctVal = if ($totalInStack -gt 0) { [Math]::Round(100 * $unresolvedByMod[$k] / $totalInStack, 1) } else { 0 }
    [void]$unresolvedHotspots.Add([pscustomobject]@{
        Module         = $k
        UnresolvedCount = $unresolvedByMod[$k]
        TotalInStack   = $totalInStack
        Pct            = $pctVal
    })
}

# --- critical section locks (!locks / !cs -l) --------------------------------

$critLocks = New-Object System.Collections.Generic.List[object]
# !locks format: "CritSec ntdll!LdrpLoaderLock+0 at 00007ff8... OwningThread: HHHH"
foreach ($m in (_MatchAll $raw '(?m)CritSec\s+(\S+)\s+at\s+([0-9a-fA-F`]+)')) {
    $csName = $m.Groups[1].Value
    $csAddr = $m.Groups[2].Value
    # Try to find owning thread and lock count in following lines
    $csIdx = $m.Index
    $csBlock = $raw.Substring($csIdx, [Math]::Min(400, $raw.Length - $csIdx))
    $owner = _Match1 $csBlock 'OwningThread:\s*([0-9a-fA-F]+)'
    $lockCount = _Match1 $csBlock 'LockCount\s*=\s*(\S+)'
    $recCount = _Match1 $csBlock 'RecursionCount\s*=\s*(\d+)'
    [void]$critLocks.Add([pscustomobject]@{
        Name           = $csName
        Address        = $csAddr
        OwningThread   = $owner
        LockCount      = $lockCount
        RecursionCount = $recCount
    })
}

# --- SEH exception chain (!exchain) ------------------------------------------

$sehChain = New-Object System.Collections.Generic.List[object]
# !exchain format: "addr: module!handler+0xNN" or "addr: <flat>" lines
foreach ($m in (_MatchAll $raw '(?m)^([0-9a-fA-F`]+):\s+(\S+![^\s]+|<flat>)')) {
    [void]$sehChain.Add([pscustomobject]@{
        Address = $m.Groups[1].Value
        Handler = $m.Groups[2].Value
    })
}

# --- CLR managed stack (.loadby sos clr; !clrstack) --------------------------

$clrStack = New-Object System.Collections.Generic.List[string]
$clrBlockM = [regex]::Match($raw, '(?ms)OS Thread Id:\s*0x[0-9a-fA-F]+.*?(!clrstack|Child\s+SP).*?\n((?:(?:(?:[0-9a-fA-F`]+\s+[0-9a-fA-F`]+\s+)|(?:SP\s+IP\s+)).*\n)*)')
if (-not $clrBlockM.Success) {
    # fallback: grab lines after !clrstack that look like managed frames
    $clrBlockM = [regex]::Match($raw, '(?ms)!clrstack\s*\n((?:[0-9a-fA-F]+\s+[0-9a-fA-F]+\s+.+\n)*)')
}
if ($clrBlockM.Success) {
    foreach ($fl in ($clrBlockM.Value -split "`r?`n")) {
        $fl = $fl.Trim()
        if ($fl -match '^[0-9a-fA-F]' -and $fl.Length -gt 10) {
            [void]$clrStack.Add($fl)
        }
    }
}

# CLR thread list (!threads from SOS)
$clrThreads = New-Object System.Collections.Generic.List[object]
foreach ($m in (_MatchAll $raw '(?m)^\s*(\d+)\s+\d+\s+([0-9a-fA-F]+)\s+([0-9a-fA-F]+\.\d+)\s+(\d+)\s+(\S+)')) {
    [void]$clrThreads.Add([pscustomobject]@{
        Index       = $m.Groups[1].Value
        OSID        = $m.Groups[2].Value
        CLRThreadId = $m.Groups[3].Value
        LockCount   = $m.Groups[4].Value
        State       = $m.Groups[5].Value
    })
}

# --- CLR runtime version (lmv m clr / lmv m coreclr) -------------------------

$clrVersion = $null
$clrModName = $null
foreach ($cn in @('clr','coreclr')) {
    $clrBlkP = '(?ms)start\s+[0-9a-fA-F`]+\s+[0-9a-fA-F`]+\s+' + $cn + '\b.*?(?=start\s+[0-9a-fA-F`]+\s+[0-9a-fA-F`]+\s+\S+|\z)'
    $clrBlk = [regex]::Match($raw, $clrBlkP).Value
    if ($clrBlk) {
        $clrVersion = _Match1 $clrBlk 'File version:\s*(.+)'
        if (-not $clrVersion) { $clrVersion = _Match1 $clrBlk 'Product version:\s*(.+)' }
        $clrModName = $cn
        break
    }
}

# --- Wow64 x86 stack (.effmach x86; kb 60) -----------------------------------

$wow64Stack = New-Object System.Collections.Generic.List[string]
$wow64BlockM = [regex]::Match($raw, '(?ms)\.effmach x86.*?\n((?:[0-9a-fA-F]+\s+[0-9a-fA-F]+\s+.*\n)*)')
if ($wow64BlockM.Success) {
    foreach ($fl in ($wow64BlockM.Groups[1].Value -split "`r?`n")) {
        $fm = [regex]::Match($fl, ':.*?:\s*(.+)$')
        if ($fm.Success) {
            $fr = $fm.Groups[1].Value.Trim()
            if ($fr) { [void]$wow64Stack.Add($fr) }
        }
    }
}

# --- security token (!token -n) -----------------------------------------------

$tokenUser     = _Match1 $raw '(?i)User:\s+(\S+\\\S+)'
if (-not $tokenUser) { $tokenUser = _Match1 $raw 'User:\s+([^\r\n]+)' }
$tokenPrivs    = New-Object System.Collections.Generic.List[string]
foreach ($m in (_MatchAll $raw '(?m)^\s+\d+\s+0x[0-9a-fA-F]+\s+(\S+)\s+Attributes\s+-\s+(.+)$')) {
    [void]$tokenPrivs.Add('{0} ({1})' -f $m.Groups[1].Value, $m.Groups[2].Value.Trim())
}
$tokenIntegrity = _Match1 $raw '(?i)Mandatory Label.*?(Low|Medium|High|System)'

# --- App Verifier (!avrf) -----------------------------------------------------

$avrfStop   = _Match1 $raw '(?i)AVRF:\s+stop\s+(\S+)'
$avrfDesc   = _Match1 $raw '(?i)AVRF:\s+stop\s+\S+\s*:\s*(.+?)(?:\r|\n)'
$avrfActive = ($raw -match '(?i)Application verifier enabled|AVRF:\s+stop')

# --- stack pointer dump (dps @rsp) -------------------------------------------

$stackPointerDump = New-Object System.Collections.Generic.List[string]
foreach ($m in (_MatchAll $raw '(?m)^([0-9a-fA-F`]+)\s+([0-9a-fA-F`]+)\s+(\S+!\S+.*)$')) {
    [void]$stackPointerDump.Add($m.Value.Trim())
    if ($stackPointerDump.Count -ge 20) { break }
}

# --- disassembly at crash point -----------------------------------------------
# cdb prints the faulting instruction before the initial command prompt line:
#   igxelpgicd64+0x87cde:
#   00007fff`71837cde 4183be381e030000 cmp  dword ptr [r14+31E38h],0
# Also try from u @rip output if available.
#
# Be aggressive about rejecting noise:
#   * heap-summary rows ("!heap -s") have hex columns separated by many
#     spaces and end with literal "LFH" / "NoLFH" / a flag bitmap;
#   * `lm` output has hex base, hex end, module name, then `(export symbols)`
#     etc. -- the "mnemonic-shaped" field is actually a module name;
#   * thread table rows have similar shape.
# Real disassembly lines are: addr + opcode_bytes + short_mnemonic + operands.
# We tighten the regex so the mnemonic must be alphabetic, 2-10 chars (x86/x64
# mnemonics are short: mov, cmp, lea, callq, repne, etc.) and followed by
# either a comma, register, bracket, or whitespace then more text. We also
# drop duplicates (cdb emits the same `cmp` line via kv, kb, .ecxr, and u).

$disassembly = New-Object System.Collections.Generic.List[string]
$disassemblySeen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
# Pattern breakdown:
#   ^(addr) (opcode 8-32 hex chars, no spaces) (mnemonic 2-10 alphabetic chars)
#   followed by whitespace then operands OR end-of-line.
foreach ($m in (_MatchAll $raw '(?m)^([0-9a-fA-F`]+)\s+([0-9a-fA-F]{8,32})\s+([a-zA-Z]{2,10})(?:\s+(.+))?$')) {
    $line     = $m.Value.Trim()
    $mnemonic = $m.Groups[3].Value
    $tail     = if ($m.Groups[4].Success) { $m.Groups[4].Value } else { '' }
    # Reject lines whose `mnemonic` field is actually a module name (e.g.
    # `swt_awt_win32_4427` followed by `(export symbols)`).
    if ($tail -match '^\(') { continue }
    # Reject heap-summary noise (ends with LFH / NoLFH).
    if ($line -match '\b(?:No)?LFH\s*$') { continue }
    # Mnemonic length sanity check (already enforced by regex, but defensive).
    if ($mnemonic.Length -lt 2 -or $mnemonic.Length -gt 10) { continue }
    # Dedup — cdb emits the same faulting instruction repeatedly.
    if (-not $disassemblySeen.Add($line)) { continue }
    [void]$disassembly.Add($line)
    if ($disassembly.Count -ge 5) { break }
}

# --- thread context (!thread / !teb) ----------------------------------------

$threadAddr = _Match1 $raw '^THREAD\s+([0-9a-fA-F`]+)'
# Cid from !thread output or ~* thread listing
$threadCid  = _Match1 $raw 'Cid\s+([0-9a-fA-F]+\.[0-9a-fA-F]+)'
if (-not $threadCid) {
    # Fallback: parse from ~ output (user-mode). Format: ". N  Id: HHHH.HHHH"
    $threadCid = _Match1 $raw '(?m)^[.\s]+\d+\s+Id:\s+([0-9a-fA-F]+\.[0-9a-fA-F]+)'
}
$threadTeb  = _Match1 $raw 'Teb:\s*([0-9a-fA-F`]+)'
$threadState = _Match1 $raw 'THREAD\s+[0-9a-fA-F`]+\s+\S+\s+([A-Za-z]+)'
$waitReason = _Match1 $raw 'WAIT:\s*\(([^\)]+)\)'

# --- handle / heap quick stats ----------------------------------------------
# !handle 0 3 (user mode) emits a single-line total "1032 Handles" followed by
# a two-column "Type / Count" table:
#     1032 Handles
#     Type                     Count
#     None                     175
#     Event                    461
#     ...
# Older kernel-style "Handle table at <addr> with N entries" never appears in
# a user-mode dump, so don't bother matching it.

$handleCount = _Match1 $raw '(?m)^\s*(\d+)\s+Handles\s*$'

$handleTypes = @{}
$ht = [regex]::Match($raw, '(?ms)^Type\s+Count\s*\r?\n((?:[A-Za-z][\w\s\(\)\-]*?\s+\d+\s*\r?\n)+)')
if ($ht.Success) {
    foreach ($line in ($ht.Groups[1].Value -split "`r?`n")) {
        $row = [regex]::Match($line, '^\s*(.+?)\s+(\d+)\s*$')
        if ($row.Success) {
            $handleTypes[$row.Groups[1].Value.Trim()] = [int]$row.Groups[2].Value
        }
    }
}

# !heap -s (NT heap stats) is column-based in KB units:
#   Heap     Flags   Reserv  Commit  Virt   Free  List   UCR  Virt  Lock  Fast
#                     (k)     (k)    (k)     (k) length      blocks cont. heap
#   ---------------------------------------------------------------------------
#   00000237226b0000 00000002  146260 124724 145868  71830  1444    32   39    3de   LFH
# Sum the Reserv (col 3) and Commit (col 4) across all data rows. There's no
# "Reserved = N K" line in user-mode output -- that's a different command.

$heapReserved  = $null
$heapCommitted = $null
$heapRows = [regex]::Matches(
    $raw,
    '(?m)^[0-9a-fA-F`]{8,18}\s+[0-9a-fA-F]{8}\s+(\d+)\s+(\d+)\s+\d+\s+\d+\s+\d+\s+\d+\s+\d+\s+[0-9a-fA-F]+\s+(?:LFH|N/A)?'
)
if ($heapRows.Count -gt 0) {
    $rSum = 0; $cSum = 0
    foreach ($r in $heapRows) {
        $rSum += [int]$r.Groups[1].Value
        $cSum += [int]$r.Groups[2].Value
    }
    $heapReserved  = $rSum.ToString()
    $heapCommitted = $cSum.ToString()
}
$heapCorrupt   = ($raw -match 'HEAP_ENTRY corrupt|heap corruption detected')

# --- virtual address space (!address -summary) --------------------------------
# cdb format: "Free    214    7ffe`45919000 ( 127.993 TB)   99.99%"
# Capture the size token "127.993 TB" etc. per usage category
function _VasSize([string]$Text, [string]$Label) {
    $m = [regex]::Match($Text, ('(?m)^' + [regex]::Escape($Label) + '\s+\d+\s+[0-9a-fA-F`]+\s+\(\s*([\d.]+\s+[KMGT]B)\)'))
    if ($m.Success) { return $m.Groups[1].Value.Trim() }
    return $null
}
$vasFreeSize    = _VasSize $raw 'Free'
$vasHeapSize    = _VasSize $raw 'Heap'
$vasImageSize   = _VasSize $raw 'Image'
$vasStackSize   = _VasSize $raw 'Stack'
$vasUnknownSize = _VasSize $raw '<unknown>'

# --- virtual memory summary (!vm) -------------------------------------------
# cdb's !vm output varies between Windows versions. On Win11 26100 the
# labels in a USER dump are:
#   Process Working Set:    NNNN
#   Peak Working Set Size:  NNNN  ( pages | KB )
#   Page File Usage:        NNNN  ( pages | KB )
# Older cdb / kernel-mode !vm uses `Working Set:` / `PageFile Usage:`. Try
# the modern user-mode labels first, then the legacy ones, then a generic
# fallback that just grabs the number after any of these label variants.
function _VmField([string]$Text, [string[]]$Labels) {
    foreach ($lbl in $Labels) {
        $pat = '(?im)^\s*' + [regex]::Escape($lbl) + '\s*:?\s+(\d[\d,.]*\s*(?:pages|kB|KB|MB|GB|K)?)'
        $m = [regex]::Match($Text, $pat)
        if ($m.Success) { return $m.Groups[1].Value.Trim() }
    }
    return $null
}
$vmWorkingSet   = _VmField $raw @('Process Working Set','Working Set Size','Working Set')
$vmPeakWS       = _VmField $raw @('Peak Working Set Size','Peak Working Set')
$vmPagefileUsed = _VmField $raw @('Page File Usage','PageFile Usage','Pagefile Usage','Commit Charge')

# --- timing (.time) ---------------------------------------------------------

$captureKernelTime = _Match1 $raw 'Kernel time:\s*(.+?)(?:\r|\n)'
$captureUserTime   = _Match1 $raw 'User time:\s*(.+?)(?:\r|\n)'
$captureElapsed    = _Match1 $raw 'Elapsed time:\s*(.+?)(?:\r|\n)'

# --- last event (.lastevent) ------------------------------------------------

$lastEvent  = _Match1 $raw 'Last event:\s*(.+?)(?:\r|\n)'
if (-not $lastEvent) { $lastEvent = _Match1 $raw '\.lastevent[^\n]*\n([^\n]+)' }

# --- last error (!gle) ------------------------------------------------------

$gleCode    = _Match1 $raw '(?:LastErrorValue|LastStatusValue):\s*\(([^\)]+)\)'
if (-not $gleCode) { $gleCode = _Match1 $raw '(?m)^\s*Error value:\s*(.+)$' }
$gleString  = _Match1 $raw '(?m)^\s*\(interpreted[^)]*\):\s*(.+)$'
if (-not $gleString) { $gleString = _Match1 $raw '(?i)Last error text:\s*(.+?)(?:\r|\n)' }

# --- PEB extra fields -------------------------------------------------------
# !peb in this cdb version prints the Ldr module list; the first Module path is the exe
$exeImagePath   = _Match1 $pebTxt "ImagePathName:\s*'([^']*)'"
if (-not $exeImagePath) { $exeImagePath = _Match1 $pebTxt 'ImagePathName:\s*(.+?)(?:\r|\n)' }
# Fallback: first .exe path after "PEB at" within 3000 chars; cdb may line-wrap paths so join
if (-not $exeImagePath) {
    $pebIdx = $pebTxt.IndexOf('PEB at')
    if ($pebIdx -ge 0) {
        $pebBlk = $pebTxt.Substring($pebIdx, [Math]::Min(3000, $pebTxt.Length - $pebIdx))
        # Join wrapped lines: a line ending mid-word (no trailing space after last word) followed
        # by a line starting without the usual hex+space prefix means it's a continuation.
        $pebExeM = [regex]::Match($pebBlk, '([A-Za-z]:.+?\.exe)\b')
        if ($pebExeM.Success) { $exeImagePath = $pebExeM.Groups[1].Value.Trim() }
    }
}
$currentDir     = _Match1 $pebTxt "CurrentDirectory:\s*'([^']*)'"
if (-not $currentDir)  { $currentDir  = _Match1 $pebTxt 'CurrentDirectory:\s*(.+?)(?:\r|\n)' }
$dllPath        = _Match1 $pebTxt "DllPath:\s*'([^']*)'"

# --- thread count (distinct THREAD blocks in *~* kb 8 / !thread output) ----
# Two distinct sources, kept separately so consumers can detect capture
# truncation. cdb's default output buffer is small; on processes with
# 100+ threads, `~* kb 8` can stream off the end of the buffer and the
# per-thread block parser then yields fewer entries than !thread reports.

# (a) From !thread (kernel-mode pattern + user-mode !thread `THREAD <addr>`)
$threadCountReported = ([regex]::Matches($raw, '(?m)^THREAD\s+[0-9a-fA-F`]+')).Count
# (b) From `~* kb` thread-block headers (user-mode authoritative source)
$threadCountFromKb   = ([regex]::Matches($raw, '(?m)^\s+\d+\s+Id:\s+[0-9a-fA-F]+\.[0-9a-fA-F]+')).Count

# Pick the larger of (reported, kb-headers) as the canonical count.
$threadCount = [Math]::Max($threadCountReported, $threadCountFromKb)
if ($threadCount -eq 0) { $threadCount = $threadStacks.Count }

# Flag a mismatch >5 threads between what !thread reported and what
# `~* kb` actually emitted into the buffer (= the data we have stacks for).
$threadCaptureTruncated = $false
if ($threadCount -gt 0 -and ($threadCount - $threadStacks.Count) -gt 5) {
    $threadCaptureTruncated = $true
}

# --- loaded modules (lm) ------------------------------------------------------

$modules = New-Object System.Collections.Generic.List[object]
foreach ($line in ($modTxt -split "`r?`n")) {
    # 00007fff`717b0000 00007fff`72d09000   igxelpgicd64   (export symbols)   igxelpgicd64.dll
    $m = [regex]::Match($line, '^([0-9a-fA-F`]+)\s+([0-9a-fA-F`]+)\s+(\S+)\s+\([^)]+\)\s*(\S*)')
    if ($m.Success) {
        [void]$modules.Add([pscustomobject]@{
            Base = $m.Groups[1].Value
            End  = $m.Groups[2].Value
            Name = $m.Groups[3].Value
            Path = $m.Groups[4].Value
        })
    }
}

# --- register-to-VA-region correlation ---------------------------------------
# Check if key registers (rip, rsp, rax, rcx, rdx) point into known module ranges
$regCorrelation = New-Object System.Collections.Generic.List[object]
$keyRegs = @('rip','rsp','rax','rcx','rdx','r8','r9')
foreach ($rn in $keyRegs) {
    $rv = $registers[$rn]
    if (-not $rv) { continue }
    $rvClean = $rv -replace '`',''
    try { $rvNum = [uint64]("0x$rvClean") } catch { continue }
    $region = $null
    foreach ($mod in $modules) {
        $bClean = ([string]$mod.Base) -replace '`',''
        $eClean = ([string]$mod.End)  -replace '`',''
        try {
            $bNum = [uint64]("0x$bClean")
            $eNum = [uint64]("0x$eClean")
        } catch { continue }
        if ($rvNum -ge $bNum -and $rvNum -lt $eNum) { $region = [string]$mod.Name; break }
    }
    [void]$regCorrelation.Add([pscustomobject]@{
        Register = $rn
        Value    = $rv
        Module   = $region  # null means points outside any loaded module
    })
}

# --- module detail snippets (from lmv blocks when available) -----------------

$moduleDetails = New-Object System.Collections.Generic.List[object]
$topMods = @($stack | Select-Object -First 12 | ForEach-Object { $_.Module } | Where-Object { $_ } | Select-Object -Unique)
# Also include faulting module if not in top stack
if ($faultMod -and $faultMod -notin $topMods) { $topMods = @($faultMod) + $topMods }
# Also include every module that an `lmv m <name>` was issued for. The
# faulting-stack heuristic alone misses modules whose first hit in the stack
# is unresolved (e.g. ntdll!KiUserCallbackDispatcherContinue is in the stack
# as `ntdll!Ki...` but ntdll wasn't always in $topMods because of a
# resolution gap); pulling from Capture.Commands makes coverage explicit.
if ($meta -and $meta.Commands) {
    foreach ($cmd in @($meta.Commands)) {
        $cmdStr = [string]$cmd
        $mm = [regex]::Match($cmdStr, '^lmv\s+m\s+(\S+)\s*$')
        if ($mm.Success) {
            $mn = $mm.Groups[1].Value
            if ($topMods -notcontains $mn) { $topMods = @($topMods) + $mn }
        }
    }
}
foreach ($mn in $topMods) {
    # cdb's lmv emits a 1-line "Browse all global symbols ..." preamble then
    # `start  end  module ...` then the detail block. Module name match must
    # be case-insensitive (cdb normalises filenames inconsistently across
    # versions and the stack-frame module name may differ in case).
    $reMod = '(?i)' + [regex]::Escape($mn)
    # 1) standard `start ... end ... modulename` header
    $pattern = '(?ms)^start\s+[0-9a-fA-F`]+\s+[0-9a-fA-F`]+\s+' + $reMod + '\b.*?(?=^start\s+[0-9a-fA-F`]+\s+[0-9a-fA-F`]+\s+\S+|\z)'
    $blk = [regex]::Match($raw, $pattern).Value
    # 2) `Loaded symbol image file:` anchor (always present in lmv, survives
    #    even when the `start  end  name` header is mangled by a long path)
    if (-not $blk) {
        $altPattern = '(?ms)(?:Loaded symbol image file|Image path|Image name):\s*[^\n]*\b' + $reMod + '(?:\.(?:dll|sys|exe))?\b.*?(?=^start\s+|^quit:|### DumpPilot|\z)'
        $blk = [regex]::Match($raw, $altPattern).Value
    }
    # 3) Image path containing the module DLL/SYS (most permissive)
    if (-not $blk) {
        foreach ($ext in @('.dll','.sys','.exe')) {
            $imgName = $mn + $ext
            $altPattern = '(?ms)Image path:\s*[^\n]*' + [regex]::Escape($imgName) + '.*?(?=^start\s+|^quit:|### DumpPilot|\z)'
            $blk = [regex]::Match($raw, $altPattern).Value
            if ($blk) { break }
        }
    }
    if (-not $blk) { continue }
    $symStatus = $null
    if ($blk -match 'export symbols') { $symStatus = 'export-only' }
    elseif ($blk -match 'private symbols') { $symStatus = 'private' }
    elseif ($blk -match 'public symbols') { $symStatus = 'public' }
    [void]$moduleDetails.Add([pscustomobject]@{
        Name           = $mn
        ImagePath      = _Match1 $blk 'Image path:\s*(.+)'
        ImageName      = _Match1 $blk 'Image name:\s*(.+)'
        Timestamp      = _Match1 $blk 'Timestamp:\s*(.+)'
        FileVersion    = _Match1 $blk 'File version:\s*(.+)'
        ProductVersion = _Match1 $blk 'Product version:\s*(.+)'
        CompanyName    = _Match1 $blk 'CompanyName:\s*(.+)'
        SymbolStatus   = $symStatus
    })
}

# Fallback: if lmv blocks were not available, emit basic details from lm output.
if ($moduleDetails.Count -eq 0) {
    foreach ($mn in ($topMods | Select-Object -First 20)) {
        $lm = @($modules | Where-Object { [string]$_.Name -eq [string]$mn } | Select-Object -First 1)
        if (-not $lm -or $lm.Count -eq 0) { continue }
        $p = [string]$lm[0].Path
        $symStatus = $null
        if ($p -match '\.pdb$') { $symStatus = 'pdb-path' }
        elseif ($p) { $symStatus = 'image-path' }
        [void]$moduleDetails.Add([pscustomobject]@{
            Name           = [string]$lm[0].Name
            ImagePath      = $p
            ImageName      = if ($p) { [System.IO.Path]::GetFileName($p) } else { $null }
            Timestamp      = $null
            FileVersion    = $null
            ProductVersion = $null
            CompanyName    = $null
            SymbolStatus   = $symStatus
        })
    }
}

# --- symbol quality -----------------------------------------------------------

$stackTotal = $stack.Count
$stackResolved = ($stack | Where-Object { $_.Symbol } | Measure-Object).Count
$stackUnresolved = $stackTotal - $stackResolved

$moduleCoverage = New-Object System.Collections.Generic.List[object]
foreach ($mname in ($topMods | Select-Object -First 12)) {
    $frames = @($stack | Where-Object { $_.Module -eq $mname })
    $resolved = @($frames | Where-Object { $_.Symbol }).Count
    [void]$moduleCoverage.Add([pscustomobject]@{
        Module       = $mname
        Frames       = $frames.Count
        Resolved     = $resolved
        Unresolved   = ($frames.Count - $resolved)
    })
}
# Synthetic <unmapped> bucket: frames where the "module name" is actually a
# raw hex address (JIT-emitted code, Java heap pointers, etc). Without this,
# the totals don't reconcile (ModuleCoverage sums to 34 but $stack has 58)
# and the LLM blames the real modules' unresolved counts for symbol-quality
# problems that are actually just non-PE code addresses.
$unmappedFrames = @($stack | Where-Object {
    $_.Module -and ($_.Module -match '^0x[0-9a-fA-F`]+$' -or $_.Module -match '^[0-9a-fA-F]+`[0-9a-fA-F]+$')
})
if ($unmappedFrames.Count -gt 0) {
    [void]$moduleCoverage.Add([pscustomobject]@{
        Module     = '<unmapped>'
        Frames     = $unmappedFrames.Count
        Resolved   = 0
        Unresolved = $unmappedFrames.Count
    })
}

# --- command line (PEB) -------------------------------------------------------

$cmdLine = _Match1 $pebTxt "CommandLine:\s*'([^']*)'"
if (-not $cmdLine) { $cmdLine = _Match1 $pebTxt 'CommandLine:\s*(.+)' }

# --- previous-run diff --------------------------------------------------------

$prevSummary = $null
if (Test-Path -LiteralPath $OutputPath) {
    try { $prevSummary = Get-Content -LiteralPath $OutputPath -Raw | ConvertFrom-Json -Depth 16 } catch {}
}

$comparison = [ordered]@{}
if ($prevSummary) {
    $prevSig = [string]$prevSummary.Faulting.StackSignature
    $changed = New-Object System.Collections.Generic.List[string]
    $curMap = @{}
    foreach ($m in $modules) { if ($m.Name) { $curMap[[string]$m.Name] = [string]$m.Path } }
    $prevMods = @($prevSummary.Modules)
    foreach ($pm in $prevMods) {
        $n = [string]$pm.Name
        if (-not $n) { continue }
        $prevPath = [string]$pm.Path
        if (-not $curMap.ContainsKey($n)) {
            [void]$changed.Add("removed:$n")
        } elseif ($curMap[$n] -ne $prevPath) {
            [void]$changed.Add("path-changed:$n")
        }
    }
    foreach ($cn in $curMap.Keys) {
        if (-not (@($prevMods | Where-Object { [string]$_.Name -eq $cn }).Count -gt 0)) {
            [void]$changed.Add("added:$cn")
        }
    }

    $comparison = [ordered]@{
        PreviousGeneratedAt  = $prevSummary.GeneratedAt
        PreviousStackSig     = $prevSig
        StackSigChanged      = ($prevSig -ne $stackSignature)
        ModuleDelta          = @($changed | Select-Object -First 40)
    }
}

# --- emit ---------------------------------------------------------------------

$summary = [ordered]@{
    SchemaVersion = '0.1'
    DumpKind      = if ($meta) { $meta.DumpKind } else { $null }
    DumpPath      = if ($meta) { $meta.DumpPath } else { $null }
    DumpSha256    = if ($meta) { $meta.DumpSha256 } else { $null }
    OsBuild       = $osBuild
    Process       = [ordered]@{
        Name        = $processNm
        CommandLine = $cmdLine
    }
    Exception     = [ordered]@{
        Code     = $excCode
        CodeName = $excCodeName
        Address  = $excAddr
        Bucket   = $bucket
    }
    Faulting      = [ordered]@{
        Module = $faultMod
        Image  = $faultImg
        Symbol = $faultSym
        Stack  = $stack
        StackSignature = $stackSignature
    }
    Modules       = $modules
    ModuleDetails = $moduleDetails
    ExceptionRecord = [ordered]@{
        Flags            = $excFlags
        NumberParameters = $excNumParams
        Parameters       = $excParams
    }
    ContextRegisters = (New-OrderedFromHashtable -Map $registers)
    Thread         = [ordered]@{
        Address    = $threadAddr
        Cid        = $threadCid
        Teb        = $threadTeb
        State      = $threadState
        WaitReason = $waitReason
    }
    ResourceStats  = [ordered]@{
        HandleCount      = $handleCount
        HeapReservedKB   = $heapReserved
        HeapCommittedKB  = $heapCommitted
    }
    SymbolQuality  = [ordered]@{
        StackFramesTotal      = $stackTotal
        StackFramesResolved   = $stackResolved
        StackFramesUnresolved = $stackUnresolved
        ModuleCoverage        = $moduleCoverage
    }
    Counts        = [ordered]@{
        StackFrames    = $stack.Count
        LoadedModules  = $modules.Count
    }
        VASummary    = [ordered]@{
            Free    = $vasFreeSize
            Heap    = $vasHeapSize
            Image   = $vasImageSize
            Stack   = $vasStackSize
            Unknown = $vasUnknownSize
        }
        VMStats      = [ordered]@{
            WorkingSetText   = $vmWorkingSet
            PeakWorkingSet   = $vmPeakWS
            PagefileUsage    = $vmPagefileUsed
        }
        Timing       = [ordered]@{
            KernelTime  = $captureKernelTime
            UserTime    = $captureUserTime
            ElapsedTime = $captureElapsed
        }
        LastEvent    = $lastEvent
        LastError    = [ordered]@{
            Code   = $gleCode
            String = $gleString
        }
        ProcessInfo  = [ordered]@{
            ExeImagePath   = $exeImagePath
            CurrentDirectory = $currentDir
            DllPath        = $dllPath
        }
        ThreadCount  = $threadCount
        ThreadCountReported = $threadCountReported
        ThreadCountCaptured = $threadStacks.Count
        ThreadCaptureTruncated = $threadCaptureTruncated
        HandleTypes  = $handleTypes
        HeapCorrupt  = $heapCorrupt
    ThreadStacks   = $threadStacks
    ExceptionChain = $excChain
    UnresolvedHotspots = $unresolvedHotspots
    RegisterCorrelation = $regCorrelation
    CriticalSections = $critLocks
    SEHChain       = $sehChain
    CLR            = [ordered]@{
        Runtime     = $clrModName
        Version     = $clrVersion
        ManagedStack = $clrStack
        Threads     = $clrThreads
    }
    Wow64Stack     = $wow64Stack
    Token          = [ordered]@{
        User           = $tokenUser
        IntegrityLevel = $tokenIntegrity
        Privileges     = $tokenPrivs
    }
    AppVerifier    = [ordered]@{
        Active      = $avrfActive
        StopCode    = $avrfStop
        Description = $avrfDesc
    }
    StackPointerDump = $stackPointerDump
    Disassembly      = $disassembly
    Capture       = [ordered]@{
        CommandCount    = if ($meta) { $meta.CommandCount } else { $null }
        Commands        = if ($meta) { $meta.Commands } else { @() }
        DurationSeconds = if ($meta) { $meta.DurationSeconds } else { $null }
        CdbPath         = if ($meta) { $meta.CdbPath } else { $null }
        SymbolPath      = if ($meta) { $meta.SymbolPath } else { $null }
    }
    Comparison     = $comparison
    GeneratedAt   = (Get-Date).ToString('o')
}

[System.IO.File]::WriteAllText(
    $OutputPath,
    ($summary | ConvertTo-Json -Depth 8),
    [System.Text.UTF8Encoding]::new($true)
)

[pscustomobject]@{
    OutputPath    = $OutputPath
    FaultingModule = $faultMod
    Bucket         = $bucket
    Process        = $processNm
    StackFrames    = $stack.Count
    LoadedModules  = $modules.Count
}
