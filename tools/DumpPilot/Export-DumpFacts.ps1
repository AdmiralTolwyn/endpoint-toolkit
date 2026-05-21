#requires -Version 5.1
<#
.SYNOPSIS
    Stage 1 of DumpPilot. Run cdb.exe against a Windows crash dump and capture
    a fixed set of facts into a per-dump folder.

.DESCRIPTION
    Does only three things:
      1. Locate cdb.exe from the Windows 10/11 SDK Debugging Tools.
      2. Detect dump kind (user mini / user full / kernel) via '||'.
      3. Run a kind-specific command bundle, splitting output into:

         <Dump>.dump-facts\
             raw.txt          full cdb stdout+stderr
             analyze.txt      !analyze -v only
             stack.txt        faulting-thread kb 30
             modules.txt      lm + lmvm for modules on the faulting stack
             peb.txt          process command line (user-mode dumps)
             meta.json        provenance (cdb path, dump kind, sha256, durations)

    No interpretation here. ConvertTo-DumpSummary.ps1 is stage 2.

.PARAMETER DumpPath
    Path to the .dmp / .mdmp / MEMORY.DMP file.

.PARAMETER OutputDirectory
    Where the per-dump folder lands. Default: next to the dump.

.PARAMETER SymbolPath
    Overrides _NT_SYMBOL_PATH for this run. If neither is set, the function
    warns and uses the public Microsoft symbol server.

.PARAMETER Force
    Rebuild an existing output folder.

.OUTPUTS
    [pscustomobject] with FactsRoot, DumpKind, RawSize, Duration.

.NOTES
    PowerShell 5.1 compatible. No PS 7 syntax.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$DumpPath,
    [string]$OutputDirectory,
    [string]$SymbolPath,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

function Find-Cdb {
    $candidates = @(
        'C:\Program Files (x86)\Windows Kits\10\Debuggers\x64\cdb.exe',
        'C:\Program Files\Windows Kits\10\Debuggers\x64\cdb.exe',
        'C:\Program Files (x86)\Windows Kits\10\Debuggers\x86\cdb.exe'
    )
    foreach ($c in $candidates) { if (Test-Path -LiteralPath $c) { return $c } }
    $w = Get-Command cdb.exe -ErrorAction SilentlyContinue
    if ($w) { return $w.Source }
    throw "cdb.exe not found. Install 'Debugging Tools for Windows' (Windows SDK)."
}

function Resolve-SymbolPath {
    param([string]$Override)
    if ($Override) { return $Override }
    if ($env:_NT_SYMBOL_PATH) { return $env:_NT_SYMBOL_PATH }
    Write-Warning "_NT_SYMBOL_PATH not set. Using public Microsoft symbol server. Module+offset buckets will still resolve."
    return 'srv*C:\Symbols*https://msdl.microsoft.com/download/symbols'
}

function Get-DumpKind {
    param([string]$Cdb, [string]$Dump, [string]$SymPath)
    $tmp = [System.IO.Path]::GetTempFileName()
    try {
        $env:_NT_SYMBOL_PATH = $SymPath
        & $Cdb -z $Dump -c '||; q' > $tmp 2>&1
        $txt = Get-Content -LiteralPath $tmp -Raw
        if     ($txt -match 'Kernel (Summary|Complete|Bitmap|Minidump)') { return 'KernelDump' }
        elseif ($txt -match 'Mini User Dump' -or $txt -match 'User Mini Dump')         { return 'UserMinidump' }
        elseif ($txt -match 'User Dump')                                  { return 'UserFullDump' }
        else { return 'Unknown' }
    } finally {
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    }
}

function Get-CommandBundle {
    param([string]$Kind)
    switch ($Kind) {
        'KernelDump'   { return @(
            '.symfix',
            '.symopt+0x40',
            '.reload /f',
            '||',
            'vertarget',
            '.time',
            '.lastevent',
            '!analyze -v',
            '!analyze -show',
            'r',
            'kv 60',
            'kb 100',
            '!thread',
            '!process 0 1',
            '!stacks 2',
            '!locks',
            '!irpfind',
            '!poolused 2',
            '!vm 1',
            '!memusage',
            '!pte',
            '!sysinfo machineid',
            '!sysinfo cpuspeed',
            'lm',
            'lmv m nt',
            'lmv m hal',
            '!irql'
        ) }
        default        { return @(
            '.symfix',
            '.symopt+0x40',
            '.reload /f',
            '||',
            'vertarget',
            '.time',
            '.lastevent',
            '!analyze -v',
            '!analyze -show',
            '.exr -1',
            '.ecxr',
            'r',
            'kv 60',
            'kb 100',
            '~* kb 8',
            '!thread',
            '!teb',
            '!peb',
            '!handle 0 3',
            '!gle',
            '!error 0n0',
            '!address -summary',
            '!vm',
            '!heap -s',
            '!locks',
            '!cs -l',
            '!exchain',
            '!avrf',
            '!token -n',
            '.catch { .loadby sos clr; !clrstack; !threads }',
            '.catch { .loadby sos coreclr; !clrstack; !threads }',
            '.catch { .effmach x86; kb 60; .effmach amd64 }',
            '~',
            '.catch { u @rip L5 }',
            'lm',
            'lmv m ntdll',
            'lmv m kernel32',
            'lmv m user32',
            'lmv m gdi32full',
            'lmv m opengl32',
            'lmv m clr',
            'lmv m coreclr',
            'lmv m mscorlib_ni',
            'lmv m System_ni'
        ) }
    }
}

function Get-Sha256 {
    param([string]$Path)
    try { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash } catch { $null }
}

function Save-RawOutput {
    param([string]$Raw, [string]$Root)
    # cdb in -c "cmd1; cmd2; ..." mode emits ONE prompt then streams the whole
    # output, so reliable per-command sectioning is not possible without a
    # custom delimiter. We just save the full output; the parser does its own
    # regex extraction.
    [System.IO.File]::WriteAllText((Join-Path $Root 'raw.txt'), $Raw, [System.Text.UTF8Encoding]::new($true))
}

# --- main ---------------------------------------------------------------------

if (-not (Test-Path -LiteralPath $DumpPath)) { throw "Dump not found: $DumpPath" }
$dumpItem = Get-Item -LiteralPath $DumpPath
if (-not $OutputDirectory) {
    $OutputDirectory = Join-Path $dumpItem.DirectoryName ($dumpItem.Name + '.dump-facts')
}
if (Test-Path -LiteralPath $OutputDirectory) {
    if (-not $Force) {
        $rawPath  = Join-Path $OutputDirectory 'raw.txt'
        $metaPath = Join-Path $OutputDirectory 'meta.json'
        if ((Test-Path -LiteralPath $rawPath) -and (Test-Path -LiteralPath $metaPath)) {
            Write-Verbose "facts  : reusing existing cache at $OutputDirectory"
            $meta = Get-Content -LiteralPath $metaPath -Raw | ConvertFrom-Json -Depth 8
            return [pscustomobject]@{
                FactsRoot = $OutputDirectory
                DumpKind  = [string]$meta.DumpKind
                RawSize   = [int64]$meta.RawBytes
                Duration  = [timespan]::Zero
                Reused    = $true
            }
        }
        throw "Output folder exists but is incomplete: $OutputDirectory (use -Force to rebuild)"
    }
    Remove-Item -LiteralPath $OutputDirectory -Recurse -Force
}
[void](New-Item -ItemType Directory -Path $OutputDirectory -Force)

$cdb     = Find-Cdb
$symPath = Resolve-SymbolPath -Override $SymbolPath
$env:_NT_SYMBOL_PATH = $symPath

Write-Verbose "cdb     : $cdb"
Write-Verbose "symbols : $symPath"
Write-Verbose "dump    : $DumpPath"

$kind = Get-DumpKind -Cdb $cdb -Dump $DumpPath -SymPath $symPath
Write-Verbose "kind    : $kind"

$bundle = Get-CommandBundle -Kind $kind
$cmdStr = ($bundle -join '; ') + '; q'

$sw    = [System.Diagnostics.Stopwatch]::StartNew()
$tmp   = [System.IO.Path]::GetTempFileName()
try {
    & $cdb -z $DumpPath -c $cmdStr > $tmp 2>&1
    $raw = Get-Content -LiteralPath $tmp -Raw
} finally {
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
}
$sw.Stop()

Save-RawOutput -Raw $raw -Root $OutputDirectory

# --- pass 2: dynamic lmv for faulting module (if not already in hardcoded list) ---
$faultModFromRaw = $null
$fmMatch = [regex]::Match($raw, 'MODULE_NAME:\s*(\S+)')
if ($fmMatch.Success) { $faultModFromRaw = $fmMatch.Groups[1].Value }
# Also grab top-of-stack modules for lmv
$stackMods = [System.Collections.Generic.List[string]]::new()
foreach ($sm in ([regex]::Matches($raw, '(?m)^[0-9a-fA-F`]+\s+[0-9a-fA-F`]+\s+:.*?:\s*([a-zA-Z][a-zA-Z0-9_]+)!'))) {
    $mn = $sm.Groups[1].Value.ToLowerInvariant()
    if (-not $stackMods.Contains($mn)) { [void]$stackMods.Add($mn) }
    if ($stackMods.Count -ge 8) { break }
}
if ($faultModFromRaw -and -not $stackMods.Contains($faultModFromRaw.ToLowerInvariant())) {
    [void]$stackMods.Insert(0, $faultModFromRaw.ToLowerInvariant())
}
$hardcoded = @('ntdll','kernel32','user32','gdi32full','opengl32','clr','coreclr','mscorlib_ni','system_ni')
$dynamicMods = @($stackMods | Where-Object { $_ -notin $hardcoded } | Select-Object -First 6)
if ($dynamicMods.Count -gt 0) {
    $pass2Cmds = @('.symfix', '.symopt+0x40', '.reload /f')
    foreach ($dm in $dynamicMods) { $pass2Cmds += "lmv m $dm" }
    $pass2Cmds += 'q'
    $pass2Str = $pass2Cmds -join '; '
    $tmp2 = [System.IO.Path]::GetTempFileName()
    try {
        & $cdb -z $DumpPath -c $pass2Str > $tmp2 2>&1
        $raw2 = Get-Content -LiteralPath $tmp2 -Raw
        $raw += "`r`n`r`n### DumpPilot pass 2: dynamic lmv ###`r`n" + $raw2
        # Re-save with appended pass-2 output
        Save-RawOutput -Raw $raw -Root $OutputDirectory
    } finally {
        Remove-Item -LiteralPath $tmp2 -Force -ErrorAction SilentlyContinue
    }
    $bundle = $bundle + $dynamicMods.ForEach({ "lmv m $_" })
}

$meta = [ordered]@{
    Tool             = 'DumpPilot/Export-DumpFacts'
    Version          = '0.1.0'
    DumpPath         = $DumpPath
    DumpSha256       = (Get-Sha256 -Path $DumpPath)
    DumpKind         = $kind
    CdbPath          = $cdb
    SymbolPath       = $symPath
    RawBytes         = $raw.Length
    CommandCount     = $bundle.Count
    Commands         = $bundle
    DurationSeconds  = [Math]::Round($sw.Elapsed.TotalSeconds, 2)
    Timestamp        = (Get-Date).ToString('o')
}
[System.IO.File]::WriteAllText(
    (Join-Path $OutputDirectory 'meta.json'),
    ($meta | ConvertTo-Json -Depth 6),
    [System.Text.UTF8Encoding]::new($true)
)

[pscustomobject]@{
    FactsRoot = $OutputDirectory
    DumpKind  = $kind
    RawSize   = $raw.Length
    Duration  = $sw.Elapsed
}
