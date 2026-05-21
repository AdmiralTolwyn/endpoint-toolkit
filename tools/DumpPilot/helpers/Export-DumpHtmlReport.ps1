#requires -Version 5.1
<#
.SYNOPSIS
    Render a DumpPilot report.json as a rich, single-file HTML page.
    Dark/light theme with stat cards, color-coded stacks, collapsible sections.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ReportJsonPath,
    [string]$OutputPath
)
$ErrorActionPreference = 'Stop'
if (-not (Test-Path -LiteralPath $ReportJsonPath)) { throw "Report not found: $ReportJsonPath" }
$report = Get-Content -LiteralPath $ReportJsonPath -Raw | ConvertFrom-Json
if (-not $OutputPath) {
    $item = Get-Item -LiteralPath $ReportJsonPath
    $OutputPath = Join-Path $item.DirectoryName ($item.BaseName + '.html')
}

function _H([string]$s) { if ($null -eq $s) { return '' }; [System.Net.WebUtility]::HtmlEncode($s) }

# --- subsystem classifier -----------------------------------------------------
function Get-Subsystem {
    $src = if ($report.Faulting.FullStack) { $report.Faulting.FullStack } else { $report.Faulting.TopStack }
    $stackMods = @()
    if ($src) { $stackMods = @($src | ForEach-Object { [string]$_.Module } | Where-Object { $_ -and $_ -notmatch '^0x' } | Select-Object -Unique) }
    $j = ($stackMods -join '|').ToLowerInvariant()
    if ($j -match 'opengl32|d3d|dxgi|igxel|nvoglv|amdxx|vulkan|wgl') { return 'GPU / Graphics' }
    if ($j -match 'awt|swing|javafx|jfx')                            { return 'JVM / AWT UI' }
    if ($j -match 'clr|coreclr|mscorlib|system\.')                    { return '.NET / CLR' }
    if ($j -match 'ws2_32|mswsock|winhttp|wininet|nsi|dnsapi')        { return 'Networking' }
    if ($j -match 'user32|win32u|gdi32|uxtheme|dwm')                  { return 'UI / Windowing' }
    if ($j -match 'rpcrt4|combase|ole32')                              { return 'COM / RPC' }
    if ($j -match 'jvm|java_exe|java!|jni')                           { return 'JVM Runtime' }
    return 'Unknown'
}
function Get-ThreadRole([string[]]$Frames) {
    $j = ($Frames -join '|').ToLowerInvariant()
    if ($j -match 'wtoolkit_eventloop|getmessage|peekmessage|dispatchmessage') { return 'UI' }
    if ($j -match 'monitorwait')   { return 'Monitor' }
    if ($j -match 'socketread|recv|accept|mswsock') { return 'IO/Net' }
    if ($j -match 'fileinputstream|readfile') { return 'IO/File' }
    if ($j -match 'sleep')         { return 'Sleep' }
    if ($j -match 'sendmessage')   { return 'SendMsg' }
    if ($j -match 'waitforsingle|waitformultiple') { return 'Wait' }
    return 'Worker'
}

$faultMod  = if ($report.Faulting.Module) { $report.Faulting.Module.ToLowerInvariant() } else { '' }
$subsystem = Get-Subsystem
$symPct    = 0
if ($report.SymbolQuality -and $report.SymbolQuality.StackFramesTotal -gt 0) {
    $symPct = [Math]::Round(100 * $report.SymbolQuality.StackFramesResolved / $report.SymbolQuality.StackFramesTotal, 0)
}
$stackSource = if ($report.Faulting.FullStack) { $report.Faulting.FullStack } else { $report.Faulting.TopStack }
$stackCount  = if ($stackSource) { @($stackSource).Count } else { 0 }

# --- thread dedup ---
$threadGroups = [System.Collections.Generic.List[object]]::new()
if ($report.ThreadStacks) {
    $tMap = [ordered]@{}
    foreach ($ts in $report.ThreadStacks) {
        $sig = (@($ts.Frames) -join ' > ')
        if (-not $tMap.Contains($sig)) {
            $tMap[$sig] = [pscustomobject]@{
                Count = 0; Role = (Get-ThreadRole -Frames @($ts.Frames))
                TopFrame = if (@($ts.Frames).Count -gt 0) { [string]$ts.Frames[0] } else { '' }
                Ids = [System.Collections.Generic.List[string]]::new()
            }
        }
        $tMap[$sig].Count++
        [void]$tMap[$sig].Ids.Add([string]$ts.ThreadId)
    }
    foreach ($v in $tMap.Values) { [void]$threadGroups.Add($v) }
}

$html = [System.Text.StringBuilder]::new(32768)
[void]$html.Append(@"
<!DOCTYPE html><html lang="en" data-theme="dark"><head><meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>DumpPilot — $(_H $report.Process.Name) crash report</title>
<style>
:root{--bg:#0a0a0c;--bg2:#111114;--card:#18181b;--card-hover:#1e1e22;--border:rgba(255,255,255,0.08);--border-strong:rgba(255,255,255,0.15);--accent:#60cdff;--accent-dim:rgba(96,205,255,0.12);--accent-text:#60cdff;--text:#e4e4e7;--text-bright:#fafafa;--muted:#71717a;--subtle:#52525b;--green:#22c55e;--green-dim:rgba(34,197,94,0.12);--red:#ef4444;--red-dim:rgba(239,68,68,0.12);--yellow:#eab308;--yellow-dim:rgba(234,179,8,0.12);--orange:#f97316;--purple:#a855f7;--radius:10px;--radius-sm:6px;--shadow:0 1px 3px rgba(0,0,0,0.4),0 4px 12px rgba(0,0,0,0.3);--font:'Segoe UI Variable Display','Segoe UI',-apple-system,system-ui,sans-serif;--mono:'Cascadia Code','Cascadia Mono','Fira Code',Consolas,monospace;}
[data-theme="light"]{--bg:#f8f9fa;--bg2:#fff;--card:#fff;--card-hover:#f4f4f5;--border:rgba(0,0,0,0.08);--border-strong:rgba(0,0,0,0.15);--accent:#0078d4;--accent-dim:rgba(0,120,212,0.08);--accent-text:#0066b8;--text:#18181b;--text-bright:#09090b;--muted:#71717a;--subtle:#a1a1aa;--green:#16a34a;--green-dim:rgba(22,163,74,0.08);--red:#dc2626;--red-dim:rgba(220,38,38,0.08);--yellow:#ca8a04;--yellow-dim:rgba(202,138,4,0.08);--orange:#ea580c;--purple:#9333ea;--shadow:0 1px 3px rgba(0,0,0,0.06),0 4px 12px rgba(0,0,0,0.04);}
*,*::before,*::after{margin:0;padding:0;box-sizing:border-box;}html{scroll-behavior:smooth;}
body{font-family:var(--font);background:var(--bg);color:var(--text);line-height:1.6;}
.toolbar{position:sticky;top:0;z-index:100;background:var(--bg2);border-bottom:1px solid var(--border);padding:12px 32px;display:flex;align-items:center;gap:16px;backdrop-filter:blur(12px);}
.toolbar h1{font-size:16px;font-weight:600;color:var(--text-bright);white-space:nowrap;}.toolbar .spacer{flex:1;}
.badge{display:inline-block;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:600;}
.sub-badge{background:var(--accent-dim);color:var(--accent-text);}.exc-badge{background:var(--red-dim);color:var(--red);}
.btn{background:var(--card);border:1px solid var(--border);border-radius:var(--radius-sm);padding:6px 14px;color:var(--text);font-size:12px;cursor:pointer;font-family:var(--font);transition:all .15s;}
.btn:hover{background:var(--card-hover);border-color:var(--border-strong);}
.container{max-width:1320px;margin:0 auto;padding:24px 32px 64px;}
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-bottom:28px;}
.stat-card{background:var(--card);border:1px solid var(--border);border-radius:var(--radius);padding:16px 18px;transition:all .15s;}
.stat-card:hover{border-color:var(--border-strong);box-shadow:var(--shadow);}
.stat-num{font-size:26px;font-weight:700;color:var(--text-bright);line-height:1.1;}.stat-label{font-size:11px;color:var(--muted);margin-top:3px;letter-spacing:.3px;}
.stat-card.green .stat-num{color:var(--green);}.stat-card.red .stat-num{color:var(--red);}.stat-card.yellow .stat-num{color:var(--yellow);}.stat-card.accent .stat-num{color:var(--accent-text);}.stat-card.orange .stat-num{color:var(--orange);}
.meta-grid{display:grid;grid-template-columns:max-content 1fr;gap:4px 16px;background:var(--card);border:1px solid var(--border);border-radius:var(--radius);padding:16px;margin-bottom:16px;}
.meta-grid dt{font-weight:600;color:var(--muted);font-size:12px;}.meta-grid dd{margin:0;font-family:var(--mono);font-size:12px;word-break:break-all;}
.section{margin-bottom:16px;}
.section-header{background:var(--card);border:1px solid var(--border);border-radius:var(--radius);padding:14px 18px;cursor:pointer;display:flex;align-items:center;gap:12px;transition:all .15s;user-select:none;}
.section-header:hover{background:var(--card-hover);}
.section-header .title{font-size:14px;font-weight:600;color:var(--text-bright);flex:1;}
.section-header .count{background:var(--accent-dim);color:var(--accent-text);padding:2px 10px;border-radius:12px;font-size:11px;font-weight:600;}
.section-header .chevron{color:var(--subtle);transition:transform .2s;font-size:12px;}
details[open] .chevron{transform:rotate(90deg);}
.section-body{border:1px solid var(--border);border-top:none;border-radius:0 0 var(--radius) var(--radius);background:var(--bg2);overflow:hidden;}
table{width:100%;border-collapse:collapse;font-size:12px;}
th{background:var(--card);color:var(--muted);text-align:left;padding:10px 14px;font-weight:600;font-size:10.5px;text-transform:uppercase;letter-spacing:.6px;border-bottom:1px solid var(--border);}
td{padding:9px 14px;border-bottom:1px solid var(--border);vertical-align:top;}tr:hover td{background:var(--card-hover);}
td.mono{font-family:var(--mono);font-size:11px;word-break:break-all;}
.frame-fault{color:var(--red);font-weight:600;}.frame-sys{color:var(--subtle);}.frame-app{color:var(--accent-text);}.frame-gpu{color:var(--orange);}.frame-unresolved{color:var(--muted);font-style:italic;}
.badge-sm{display:inline-block;padding:1px 8px;border-radius:4px;font-size:10px;font-weight:600;}
.badge-pdb{background:var(--green-dim);color:var(--green);}.badge-export{background:var(--yellow-dim);color:var(--yellow);}.badge-none{background:var(--red-dim);color:var(--red);}.badge-role{background:var(--accent-dim);color:var(--accent-text);}
.warning-bar{background:var(--red-dim);color:var(--red);padding:12px 18px;border-radius:var(--radius);margin-bottom:16px;font-weight:600;border:1px solid rgba(239,68,68,0.2);}
pre{font-family:var(--mono);font-size:11px;background:var(--card);border:1px solid var(--border);border-radius:var(--radius-sm);padding:12px;overflow:auto;color:var(--text);white-space:pre-wrap;word-break:break-all;}
.footer{text-align:center;color:var(--subtle);font-size:11px;margin-top:48px;padding-top:16px;border-top:1px solid var(--border);}
@media(max-width:768px){.toolbar{padding:10px 16px;flex-wrap:wrap;}.container{padding:16px;}.stats{grid-template-columns:repeat(2,1fr);}}
@media print{.toolbar{position:static;}.btn{display:none;}body{background:#fff;color:#000;}.stat-card,.section-header,.section-body,th{background:#f5f5f5;border-color:#ddd;}.stat-num,.section-header .title{color:#000;}}
</style></head><body>
"@)

# --- toolbar ------------------------------------------------------------------
$threadCount = if ($report.ThreadCount) { [int]$report.ThreadCount } else { 0 }
$cmdCount    = if ($report.Capture.CommandCount) { [int]$report.Capture.CommandCount } else { 0 }

[void]$html.Append(@"
<div class="toolbar">
  <h1>&#x1F50D; DumpPilot</h1>
  <span class="badge sub-badge">$(_H $subsystem)</span>
  <span class="badge exc-badge">$(_H $report.Exception.Code)</span>
  <span class="spacer"></span>
  <button class="btn" onclick="document.documentElement.dataset.theme=document.documentElement.dataset.theme==='dark'?'light':'dark'">&#x263E; Theme</button>
  <button class="btn" onclick="window.print()">&#x1F5A8; Print</button>
</div>
<div class="container">
"@)

# --- stat cards ---------------------------------------------------------------
$symColor = if ($symPct -ge 80) { 'green' } elseif ($symPct -ge 50) { 'yellow' } else { 'red' }
$heapTag  = if ($report.HeapCorrupt) { 'red' } else { 'green' }
$heapLbl  = if ($report.HeapCorrupt) { 'CORRUPT' } else { 'Clean' }

[void]$html.Append(@"
<div class="stats">
  <div class="stat-card accent"><div class="stat-num">$(_H $report.Process.Name)</div><div class="stat-label">Process</div></div>
  <div class="stat-card red"><div class="stat-num">$(_H $report.Exception.Code)</div><div class="stat-label">Exception</div></div>
  <div class="stat-card orange"><div class="stat-num">$stackCount</div><div class="stat-label">Stack frames</div></div>
  <div class="stat-card accent"><div class="stat-num">$threadCount</div><div class="stat-label">Threads</div></div>
  <div class="stat-card $symColor"><div class="stat-num">$symPct%</div><div class="stat-label">Symbol coverage</div></div>
  <div class="stat-card $heapTag"><div class="stat-num">$heapLbl</div><div class="stat-label">Heap integrity</div></div>
  <div class="stat-card"><div class="stat-num">$cmdCount</div><div class="stat-label">cdb commands</div></div>
</div>
"@)

# --- meta grid ----------------------------------------------------------------
[void]$html.Append('<dl class="meta-grid">')
[void]$html.Append("<dt>Dump</dt><dd>$(_H $report.Dump.Path)</dd>")
[void]$html.Append("<dt>Kind</dt><dd>$(_H $report.Dump.Kind)</dd>")
[void]$html.Append("<dt>SHA-256</dt><dd>$(_H $report.Dump.Sha256)</dd>")
[void]$html.Append("<dt>Subsystem</dt><dd>$(_H $subsystem)</dd>")
[void]$html.Append("<dt>Bucket</dt><dd>$(_H $report.Exception.Bucket)</dd>")
[void]$html.Append("<dt>Faulting</dt><dd>$(_H $report.Faulting.Module) ($(_H $report.Faulting.Symbol))</dd>")
if ($report.ProcessInfo.ExeImagePath)    { [void]$html.Append("<dt>Exe</dt><dd>$(_H $report.ProcessInfo.ExeImagePath)</dd>") }
if ($report.ProcessInfo.CurrentDirectory){ [void]$html.Append("<dt>WorkDir</dt><dd>$(_H $report.ProcessInfo.CurrentDirectory)</dd>") }
if ($report.LastEvent)                   { [void]$html.Append("<dt>Last event</dt><dd>$(_H $report.LastEvent)</dd>") }
if ($report.Timing) {
    [void]$html.Append("<dt>Kernel time</dt><dd>$(_H ([string]$report.Timing.KernelTime))</dd>")
    [void]$html.Append("<dt>User time</dt><dd>$(_H ([string]$report.Timing.UserTime))</dd>")
}
[void]$html.Append("<dt>Generated</dt><dd>$(_H $report.GeneratedAt)</dd>")
[void]$html.Append('</dl>')

if ($report.HeapCorrupt) {
    [void]$html.Append('<div class="warning-bar">&#x26A0; HEAP CORRUPTION detected in raw debugger output</div>')
}

# --- helper: collapsible section ----------------------------------------------
function _SecOpen([string]$Title, [string]$Count, [bool]$Open) {
    $o = if ($Open) { ' open' } else { '' }
    [void]$html.Append("<details class=`"section`"$o><summary class=`"section-header`"><span class=`"title`">$Title</span><span class=`"count`">$Count</span><span class=`"chevron`">&#x25B6;</span></summary><div class=`"section-body`">")
}
function _SecClose { [void]$html.Append('</div></details>') }

# --- command line -------------------------------------------------------------
_SecOpen 'Command line' '' $true
[void]$html.Append('<pre>' + (_H $report.Process.CommandLine) + '</pre>')
_SecClose

# --- faulting stack (color-coded) ---------------------------------------------
_SecOpen 'Faulting stack' "$stackCount frames" $true
[void]$html.Append('<table><tr><th>#</th><th>Module</th><th>Symbol</th><th>Frame</th></tr>')
$i = 0
foreach ($f in $stackSource) {
    $mod = [string]$f.Module; $sym = [string]$f.Symbol; $cls = 'mono'
    if ($mod -and $mod -notmatch '^0x') {
        $ml = $mod.ToLowerInvariant()
        if ($ml -eq $faultMod)       { $cls = 'mono frame-fault' }
        elseif ($ml -match '^(ntdll|kernel32|kernelbase|user32|win32u|gdi32|gdi32full)$') { $cls = 'mono frame-sys' }
        elseif ($ml -match 'opengl32|igxel|d3d|dxgi') { $cls = 'mono frame-gpu' }
        elseif ($ml -match '^(jvm|java_exe|java|awt|net)$') { $cls = 'mono frame-app' }
    } elseif ($mod -match '^0x') { $cls = 'mono frame-unresolved' }
    [void]$html.Append("<tr><td>$i</td><td class=`"$cls`">$(_H $mod)</td><td class=`"$cls`">$(_H $sym)</td><td class=`"$cls`">$(_H ([string]$f.Frame))</td></tr>")
    $i++
}
[void]$html.Append('</table>')
_SecClose

# --- threads (deduplicated) ---------------------------------------------------
if ($threadGroups.Count -gt 0) {
    _SecOpen 'Threads (deduplicated)' "$threadCount threads / $($threadGroups.Count) patterns" $false
    [void]$html.Append('<table><tr><th>Group</th><th>Count</th><th>Role</th><th>Top frame</th><th>Thread IDs</th></tr>')
    foreach ($g in ($threadGroups | Sort-Object Count -Descending)) {
        $label = if ($g.Count -gt 1) { "$($g.Count)x $($g.Role)" } else { $g.Role }
        $ids = ($g.Ids -join ', ')
        if ($ids.Length -gt 100) { $ids = $ids.Substring(0,97) + '...' }
        [void]$html.Append("<tr><td><b>$(_H $label)</b></td><td>$($g.Count)</td><td><span class=`"badge-sm badge-role`">$(_H $g.Role)</span></td><td class=`"mono`">$(_H $g.TopFrame)</td><td class=`"mono`" style=`"font-size:10px;color:var(--muted)`">$(_H $ids)</td></tr>")
    }
    [void]$html.Append('</table>')
    _SecClose
}

# --- registers ----------------------------------------------------------------
if ($report.ContextRegisters -or $report.RegisterCorrelation) {
    _SecOpen 'Registers' '' $false
    if ($report.RegisterCorrelation -and @($report.RegisterCorrelation).Count -gt 0) {
        [void]$html.Append('<table><tr><th>Register</th><th>Value</th><th>Module</th></tr>')
        foreach ($rc in $report.RegisterCorrelation) {
            $ml = if ($rc.Module) { "<span style=`"color:var(--green)`">$(_H ([string]$rc.Module))</span>" } else { "<span style=`"color:var(--red)`">UNMAPPED</span>" }
            [void]$html.Append("<tr><td><b>$(_H ([string]$rc.Register))</b></td><td class=`"mono`">$(_H ([string]$rc.Value))</td><td>$ml</td></tr>")
        }
        [void]$html.Append('</table>')
    }
    if ($report.ContextRegisters) {
        [void]$html.Append('<pre style="margin-top:8px">')
        foreach ($prop in ($report.ContextRegisters.PSObject.Properties | Sort-Object Name)) {
            [void]$html.AppendLine((_H ("{0,-6} = {1}" -f $prop.Name, $prop.Value)))
        }
        [void]$html.Append('</pre>')
    }
    _SecClose
}

# --- exception record ---------------------------------------------------------
if ($report.ExceptionRecord -and ($report.ExceptionRecord.Flags -or $report.ExceptionRecord.NumberParameters)) {
    _SecOpen 'Exception record' '' $false
    [void]$html.Append("<dl class=`"meta-grid`"><dt>Flags</dt><dd>$(_H ([string]$report.ExceptionRecord.Flags))</dd><dt>Parameters</dt><dd>$(_H ([string]$report.ExceptionRecord.NumberParameters))</dd>")
    if ($report.ExceptionRecord.Parameters) { [void]$html.Append("<dt>Values</dt><dd>$(_H ((@($report.ExceptionRecord.Parameters) -join ', ')))</dd>") }
    [void]$html.Append('</dl>')
    _SecClose
}

# --- symbol quality -----------------------------------------------------------
if ($report.SymbolQuality) {
    _SecOpen "Symbol quality" "$symPct% coverage" $false
    [void]$html.Append("<dl class=`"meta-grid`"><dt>Resolved</dt><dd>$([string]$report.SymbolQuality.StackFramesResolved) / $([string]$report.SymbolQuality.StackFramesTotal)</dd><dt>Unresolved</dt><dd>$([string]$report.SymbolQuality.StackFramesUnresolved)</dd></dl>")
    if ($report.UnresolvedHotspots -and @($report.UnresolvedHotspots).Count -gt 0) {
        [void]$html.Append('<table><tr><th>Module</th><th>Unresolved</th><th>Total</th><th>%</th></tr>')
        foreach ($uh in $report.UnresolvedHotspots) { [void]$html.Append("<tr><td>$(_H ([string]$uh.Module))</td><td>$([string]$uh.UnresolvedCount)</td><td>$([string]$uh.TotalInStack)</td><td>$([string]$uh.Pct)%</td></tr>") }
        [void]$html.Append('</table>')
    }
    _SecClose
}

# --- modules ------------------------------------------------------------------
if ($report.ModuleDetails -and @($report.ModuleDetails).Count -gt 0) {
    _SecOpen 'Key modules' "$(@($report.ModuleDetails).Count)" $false
    [void]$html.Append('<table><tr><th>Module</th><th>Image</th><th>Version</th><th>Company</th><th>Symbols</th></tr>')
    foreach ($md in $report.ModuleDetails) {
        $sb = switch ([string]$md.SymbolStatus) {
            { $_ -match 'pdb|public' } { '<span class="badge-sm badge-pdb">PDB</span>' }
            { $_ -match 'export' }     { '<span class="badge-sm badge-export">export</span>' }
            default                    { '<span class="badge-sm badge-none">none</span>' }
        }
        [void]$html.Append("<tr><td><b>$(_H ([string]$md.Name))</b></td><td class=`"mono`">$(_H ([string]$md.ImageName))</td><td class=`"mono`">$(_H ([string]$md.FileVersion))</td><td>$(_H ([string]$md.CompanyName))</td><td>$sb</td></tr>")
    }
    [void]$html.Append('</table>')
    _SecClose
}

# --- VAS ----------------------------------------------------------------------
if ($report.VASummary) {
    _SecOpen 'Virtual address space' '' $false
    [void]$html.Append('<table><tr><th>Region</th><th>Size</th></tr>')
    foreach ($prop in $report.VASummary.PSObject.Properties) { [void]$html.Append("<tr><td>$(_H $prop.Name)</td><td class=`"mono`">$(_H ([string]$prop.Value))</td></tr>") }
    [void]$html.Append('</table>')
    _SecClose
}

# --- CLR / token / avrf (conditional) -----------------------------------------
if ($report.CLR -and $report.CLR.Runtime) {
    _SecOpen 'CLR / .NET' '' $false
    [void]$html.Append("<dl class=`"meta-grid`"><dt>Runtime</dt><dd>$(_H ([string]$report.CLR.Runtime)) v$(_H ([string]$report.CLR.Version))</dd></dl>")
    if (@($report.CLR.ManagedStack).Count -gt 0) {
        [void]$html.Append('<pre>'); foreach ($mf in $report.CLR.ManagedStack) { [void]$html.AppendLine((_H ([string]$mf))) }; [void]$html.Append('</pre>')
    }
    _SecClose
}
if ($report.Token -and $report.Token.User) {
    _SecOpen 'Security context' '' $false
    [void]$html.Append("<dl class=`"meta-grid`"><dt>User</dt><dd>$(_H ([string]$report.Token.User))</dd><dt>Integrity</dt><dd>$(_H ([string]$report.Token.IntegrityLevel))</dd></dl>")
    _SecClose
}
if ($report.AppVerifier -and $report.AppVerifier.Active) {
    _SecOpen 'App Verifier' '' $true
    [void]$html.Append("<p style=`"padding:14px`">Stop: <b style=`"color:var(--red)`">$(_H ([string]$report.AppVerifier.StopCode))</b> &mdash; $(_H ([string]$report.AppVerifier.Description))</p>")
    _SecClose
}

# --- debugger commands --------------------------------------------------------
_SecOpen 'Debugger commands' "$cmdCount commands / $([string]$report.Capture.DurationSeconds)s" $false
[void]$html.Append('<pre>'); foreach ($c in $report.Capture.Commands) { [void]$html.AppendLine((_H ([string]$c))) }; [void]$html.Append('</pre>')
_SecClose

# --- footer -------------------------------------------------------------------
[void]$html.Append(@"
</div>
<div class="footer">Generated by DumpPilot &middot; $(_H $report.GeneratedAt)</div>
</body></html>
"@)

[System.IO.File]::WriteAllText($OutputPath, $html.ToString(), [System.Text.UTF8Encoding]::new($true))
[pscustomobject]@{ HtmlPath = $OutputPath; Bytes = (Get-Item -LiteralPath $OutputPath).Length }
