#requires -Version 5.1
<#
.SYNOPSIS
    DumpPilot GUI (Phase 2). WPF host that wraps the Phase 1b CLI helpers.

.DESCRIPTION
    Visual fidelity matches the rest of the toolkit:
      - Frameless dark window (WindowChrome) with custom title bar
      - 48px icon NavRail (Segoe MDL2 Assets) with active-state indicator
      - DynamicResource Theme* brushes; Set-Theme swaps light/dark at runtime
      - AccentButton / GhostButton / IconButton control styles
      - Themed DataGrid + ScrollBar + ToolTip + TextBox
      - RichTextBox log console with per-level colors (INFO/SUCCESS/WARN/ERROR/DEBUG/STEP)
      - Write-DebugLog persists to %TEMP%\DumpPilot_debug.log

    Heavy work runs in a background STA runspace polled by a 250ms DispatcherTimer
    so the UI never freezes during a 30-60s cdb run.

        Dump pages, switched via NavRail:
            Analyze dump  - pick dump, run deterministic pipeline, review facts
            History       - last 200 analyses
            AI            - run/copy llm-prompt.md against local/hosted endpoint
            Settings      - runtime + symbols + LLM endpoint settings

.NOTES
    Requires STA. The script exits with a clear message if it finds itself on MTA.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 3.0

# ---------------------------------------------------------------- STA check
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Error @'
DumpPilot GUI must run on an STA thread.
Launch from Windows PowerShell 5.1 with -STA:
    powershell.exe -STA -File .\DumpPilot.ps1
'@
    exit 1
}

# ---------------------------------------------------------------- WPF
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Xaml

# ---------------------------------------------------------------- Paths
$Script:AppVersion  = '0.2.0'
$Script:Here        = Split-Path -Parent $PSCommandPath
$Script:Helpers     = Join-Path $Script:Here 'helpers'
$Script:UIPath      = Join-Path $Script:Here 'DumpPilot_UI.xaml'
$Script:HistoryPath = Join-Path $Script:Here 'analysis_history.json'
$Script:PrefsPath   = Join-Path $Script:Here 'user_prefs.json'
$Script:DbDefaults  = Join-Path $Script:Here 'pattern_db_defaults.json'
$Script:DbUser      = Join-Path $Script:Here 'pattern_db.json'
$Script:LogPath     = Join-Path $env:TEMP 'DumpPilot_debug.log'

# ---------------------------------------------------------------- Theme palettes
$Script:ThemeDark = @{
    ThemeAppBg='#FF111113'; ThemePanelBg='#FF18181B'; ThemeCardBg='#FF1E1E1E'
    ThemeInputBg='#FF141414'; ThemeDeepBg='#FF0D0D0D'
    ThemeAccent='#FF0078D4'; ThemeAccentHover='#FF1A8AD4'; ThemeAccentLight='#FF60CDFF'
    ThemeGreenAccent='#FF00C853'
    ThemeTextPrimary='#FFFFFFFF'; ThemeTextBody='#FFE0E0E0'; ThemeTextSecondary='#FFA1A1AA'
    ThemeTextMuted='#FF8B8B93'; ThemeTextDim='#FF7A7A84'
    ThemeHoverBg='#FF27272B'; ThemeSelectedBg='#FF2A2A2A'; ThemePressedBg='#FF1A1A1A'
    ThemeError='#FFFF5000'; ThemeWarning='#FFF59E0B'; ThemeSuccess='#FF00C853'
    ThemeBorder='#19FFFFFF'; ThemeBorderCard='#19FFFFFF'; ThemeBorderElevated='#FF333333'
    ThemeBorderSubtle='#19FFFFFF'
    ThemeScrollTrack='#08FFFFFF'; ThemeScrollThumb='#40FFFFFF'; ThemeOutputBg='#FF0F0F11'
    SeverityCritical='#FFFF5000'; SeverityHigh='#FFF59E0B'; SeverityMedium='#FF60CDFF'
    SeverityLow='#FFA78BFA'; SeverityInfo='#FF00C853'
}
$Script:ThemeLight = @{
    ThemeAppBg='#FFF5F5F5'; ThemePanelBg='#FFFFFFFF'; ThemeCardBg='#FFFFFFFF'
    ThemeInputBg='#FFF0F0F0'; ThemeDeepBg='#FFE8E8E8'
    ThemeAccent='#FF0078D4'; ThemeAccentHover='#FF106EBE'; ThemeAccentLight='#FF005A9E'
    ThemeGreenAccent='#FF00C853'
    ThemeTextPrimary='#FF111111'; ThemeTextBody='#FF333333'; ThemeTextSecondary='#FF6B6B73'
    ThemeTextMuted='#FF636370'; ThemeTextDim='#FF707078'
    ThemeHoverBg='#FFEBEBEB'; ThemeSelectedBg='#FFE0E0E0'; ThemePressedBg='#FFD5D5D5'
    ThemeError='#FFD32F2F'; ThemeWarning='#FFF57F17'; ThemeSuccess='#FF2E7D32'
    ThemeBorder='#19000000'; ThemeBorderCard='#19000000'; ThemeBorderElevated='#FFCCCCCC'
    ThemeBorderSubtle='#19000000'
    ThemeScrollTrack='#08000000'; ThemeScrollThumb='#40000000'; ThemeOutputBg='#FFF0F0F2'
    SeverityCritical='#FFD32F2F'; SeverityHigh='#FFF57F17'; SeverityMedium='#FF0078D4'
    SeverityLow='#FF8888AA'; SeverityInfo='#FF2E7D32'
}

# ---------------------------------------------------------------- Log palette
$Script:LogLevelColorsDark = @{
    INFO='#60CDFF'; SUCCESS='#00C853'; WARN='#F59E0B'; ERROR='#FF5000'
    DEBUG='#A78BFA'; STEP='#F472B6'; SYSTEM='#71717A'
}
$Script:LogLevelColorsLight = @{
    INFO='#0078D4'; SUCCESS='#008A2E'; WARN='#B86E00'; ERROR='#CC0000'
    DEBUG='#8888AA'; STEP='#D63384'; SYSTEM='#636370'
}
$Script:LogLevelColors = $Script:LogLevelColorsDark

# ---------------------------------------------------------------- State
$Script:State = @{
    DumpPath    = $null
    ReportPath  = $null
    Report      = $null
    Job         = $null   # @{ PS=$ps; Async=$ar; Timer=$t; OnDone=$cb; Kind='Analyze'|'Compare' }
    Prefs       = @{
        IsLightMode=$false; ShowLog=$true; Verbose=$false; Animations=$true; DebugOverlay=$false
        LlmBaseUrl='http://localhost:11434/v1'  # Ollama default; LM Studio :1234/v1, llamafile :8080/v1
        LlmModel='gemma4'                        # Gemma 4 e4b; change in Settings
        LlmApiKey=''                            # only needed for hosted endpoints
        LlmTimeout=600                          # seconds; increase for large prompts on slow models
        SymbolPath=''                           # empty = use _NT_SYMBOL_PATH or public MS symbols
    }
    PatternList = @()
    LastAiResponse = $null
}

# Load prefs (best-effort)
if (Test-Path -LiteralPath $Script:PrefsPath) {
    try {
        $p = Get-Content -LiteralPath $Script:PrefsPath -Raw | ConvertFrom-Json
        if ($p.defaults) {
            if ($null -ne $p.defaults.PSObject.Properties['isLightMode'])   { $Script:State.Prefs.IsLightMode   = [bool]$p.defaults.isLightMode }
            if ($null -ne $p.defaults.PSObject.Properties['showLog'])       { $Script:State.Prefs.ShowLog       = [bool]$p.defaults.showLog }
            if ($null -ne $p.defaults.PSObject.Properties['verbose'])       { $Script:State.Prefs.Verbose       = [bool]$p.defaults.verbose }
            if ($null -ne $p.defaults.PSObject.Properties['animations'])    { $Script:State.Prefs.Animations    = [bool]$p.defaults.animations }
            if ($null -ne $p.defaults.PSObject.Properties['debugOverlay'])  { $Script:State.Prefs.DebugOverlay  = [bool]$p.defaults.debugOverlay }
            if ($null -ne $p.defaults.PSObject.Properties['llmBaseUrl'])    { $Script:State.Prefs.LlmBaseUrl    = [string]$p.defaults.llmBaseUrl }
            if ($null -ne $p.defaults.PSObject.Properties['llmModel'])      { $Script:State.Prefs.LlmModel      = [string]$p.defaults.llmModel }
            if ($null -ne $p.defaults.PSObject.Properties['llmApiKey'])     { $Script:State.Prefs.LlmApiKey     = [string]$p.defaults.llmApiKey }
            if ($null -ne $p.defaults.PSObject.Properties['llmTimeout'])    { $Script:State.Prefs.LlmTimeout    = [int]$p.defaults.llmTimeout }
            if ($null -ne $p.defaults.PSObject.Properties['symbolPath'])    { $Script:State.Prefs.SymbolPath    = [string]$p.defaults.symbolPath }
        }
    } catch { }
}

# ---------------------------------------------------------------- XAML load
[xml]$xaml = Get-Content -LiteralPath $Script:UIPath -Raw
$reader  = [System.Xml.XmlNodeReader]::new($xaml)
$Script:Window = [Windows.Markup.XamlReader]::Load($reader)

# Auto-discover named controls into $ui
$ui = @{}
$xaml.SelectNodes("//*[@*[local-name()='Name']]") | ForEach-Object {
    $n = $_.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml')
    if (-not $n) { $n = $_.Name }
    $elem = $Script:Window.FindName($n)
    if ($elem) { $ui[$n] = $elem }
}

# Convenience alias for cleaner code below
$Window = $Script:Window

# ---------------------------------------------------------------- Theme engine
function Set-Theme {
    param([hashtable]$palette)
    foreach ($kv in $palette.GetEnumerator()) {
        $brush = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.ColorConverter]::ConvertFromString($kv.Value))
        $brush.Freeze()
        $Window.Resources[$kv.Key] = $brush
    }
}

function Apply-CurrentTheme {
    if ($Script:State.Prefs.IsLightMode) {
        Set-Theme $Script:ThemeLight
        $Script:LogLevelColors = $Script:LogLevelColorsLight
    } else {
        Set-Theme $Script:ThemeDark
        $Script:LogLevelColors = $Script:LogLevelColorsDark
    }
}

# ---------------------------------------------------------------- Logging
$Script:LogConverter = [System.Windows.Media.BrushConverter]::new()

function Write-DebugLog {
    param(
        [string]$Message,
        [ValidateSet('INFO','SUCCESS','WARN','ERROR','DEBUG','STEP','SYSTEM')]
        [string]$Level = 'INFO'
    )
    $ts = Get-Date -Format 'HH:mm:ss.fff'
    $line = "[$ts] [$Level] $Message"
    try { [System.IO.File]::AppendAllText($Script:LogPath, $line + "`r`n") } catch {}
    Write-Host $line -ForegroundColor DarkGray

    if ($Level -eq 'DEBUG' -and -not $Script:State.Prefs.Verbose) { return }
    if (-not $ui.paraLog) { return }

    $color = $Script:LogLevelColors[$Level]
    if (-not $color) { $color = $Script:LogLevelColors['SYSTEM'] }
    $tsColor = if ($Script:State.Prefs.IsLightMode) { '#8B8B93' } else { '#52525B' }

    if ($ui.paraLog.Inlines.Count -gt 0) {
        $ui.paraLog.Inlines.Add([System.Windows.Documents.LineBreak]::new())
    }
    $tsRun = [System.Windows.Documents.Run]::new("[$ts] ")
    $tsRun.Foreground = $Script:LogConverter.ConvertFromString($tsColor)
    $tsRun.FontSize = 10.5
    $ui.paraLog.Inlines.Add($tsRun)

    $msgRun = [System.Windows.Documents.Run]::new("[$Level] $Message")
    $msgRun.Foreground = $Script:LogConverter.ConvertFromString($color)
    $msgRun.FontSize = 11
    $ui.paraLog.Inlines.Add($msgRun)

    # Cap at ~1000 inlines (paraLog perf guard)
    while ($ui.paraLog.Inlines.Count -gt 2000) {
        $ui.paraLog.Inlines.Remove($ui.paraLog.Inlines.FirstInline)
    }

    # Autoscroll
    try { $ui.rtbLog.ScrollToEnd() } catch {}
}

# ---------------------------------------------------------------- Status helpers
function Set-Status {
    param([string]$Text, [ValidateSet('Ready','Working','Error','Success')] [string]$Tone = 'Ready', [string]$Detail = '')
    $ui.StatusText.Text   = $Text
    $ui.StatusDetail.Text = $Detail
    $brush = switch ($Tone) {
        'Working' { $Window.FindResource('ThemeWarning') }
        'Error'   { $Window.FindResource('ThemeError') }
        'Success' { $Window.FindResource('ThemeSuccess') }
        default   { $Window.FindResource('ThemeSuccess') }
    }
    $ui.StatusDot.Fill = $brush
}

function Set-Busy {
    param([bool]$Busy, [string]$Status = '')
    foreach ($n in @('btnAnalyze','btnCompare','btnBrowse','btnBrowseBaseline','btnBrowseCandidate')) {
        if ($ui.$n) { $ui.$n.IsEnabled = -not $Busy }
    }
    if ($Busy) {
        Set-Status -Text ($Status ? $Status : 'Working...') -Tone 'Working'
    } else {
        Set-Status -Text ($Status ? $Status : 'Ready') -Tone 'Ready'
    }
    # Shimmer progress bar shows while busy. Honours the Animations toggle.
    if ($ui.pnlGlobalProgress) {
        if ($Busy -and $Script:State.Prefs.Animations) {
            $ui.pnlGlobalProgress.Visibility = 'Visible'
        } else {
            $ui.pnlGlobalProgress.Visibility = 'Collapsed'
        }
    }
}

# ---------------------------------------------------------------- File picker
function Show-FilePicker {
    param([string]$Title = 'Select a dump', [string]$Filter = 'Dump files (*.dmp;*.mdmp)|*.dmp;*.mdmp|All files|*.*')
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = [System.Windows.Forms.OpenFileDialog]::new()
    $dlg.Title  = $Title
    $dlg.Filter = $Filter
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.FileName }
    return $null
}

# ---------------------------------------------------------------- Pattern / history
function Load-Patterns {
    $all = [System.Collections.Generic.List[object]]::new()
    $loadedFrom = [System.Collections.Generic.List[string]]::new()
    foreach ($p in @($Script:DbDefaults, $Script:DbUser)) {
        if (-not (Test-Path -LiteralPath $p)) { continue }
        try {
            $doc = Get-Content -LiteralPath $p -Raw | ConvertFrom-Json -Depth 16
            $rules = @()
            if ($doc -and $doc.PSObject.Properties['rules']) {
                $rules = @($doc.rules)
            } else {
                $rules = @($doc)
            }
            foreach ($r in $rules) {
                if (-not $r) { continue }
                $matchFields = [System.Collections.Generic.List[string]]::new()
                if ($r.match) {
                    foreach ($mp in $r.match.PSObject.Properties) {
                        if ($mp.Value) { [void]$matchFields.Add([string]$mp.Name) }
                    }
                }
                [void]$all.Add([pscustomobject]@{
                    Id            = [string]$r.id
                    Title         = [string]$r.title
                    Severity      = [string]$r.severity
                    MatchFields   = ($matchFields -join ', ')
                    SourceFile    = [System.IO.Path]::GetFileName($p)
                })
            }
            [void]$loadedFrom.Add([System.IO.Path]::GetFileName($p))
        } catch {
            Write-DebugLog "Failed to load $p`: $($_.Exception.Message)" 'WARN'
        }
    }
    $Script:State.PatternSourceFiles = @($loadedFrom)
    return @($all | Sort-Object Id)
}

function Load-History {
    if (-not (Test-Path -LiteralPath $Script:HistoryPath)) { return @() }
    try {
        $h = Get-Content -LiteralPath $Script:HistoryPath -Raw | ConvertFrom-Json -Depth 32
        if ($h.entries) { return @($h.entries | Sort-Object WhenUtc -Descending) }
    } catch {
        Write-DebugLog "Failed to load history: $($_.Exception.Message)" 'WARN'
    }
    return @()
}

function Append-History {
    param([pscustomobject]$Entry)
    $obj = if (Test-Path -LiteralPath $Script:HistoryPath) {
        try { Get-Content -LiteralPath $Script:HistoryPath -Raw | ConvertFrom-Json -Depth 32 } catch { $null }
    } else { $null }
    if (-not $obj) { $obj = [pscustomobject]@{ version = 1; entries = @() } }
    $list = [System.Collections.Generic.List[object]]::new()
    if ($obj.entries) { foreach ($e in @($obj.entries)) { [void]$list.Add($e) } }
    [void]$list.Add($Entry)
    while ($list.Count -gt 200) { $list.RemoveAt(0) }
    $obj = [pscustomobject]@{ version = 1; entries = @($list) }
    [System.IO.File]::WriteAllText(
        $Script:HistoryPath,
        ($obj | ConvertTo-Json -Depth 16),
        [System.Text.UTF8Encoding]::new($true))
}

function Save-Prefs {
    $obj = [pscustomobject]@{
        version  = 1
        defaults = [pscustomobject]@{
            isLightMode  = $Script:State.Prefs.IsLightMode
            showLog      = $Script:State.Prefs.ShowLog
            verbose      = $Script:State.Prefs.Verbose
            animations   = $Script:State.Prefs.Animations
            debugOverlay = $Script:State.Prefs.DebugOverlay
            llmBaseUrl   = $Script:State.Prefs.LlmBaseUrl
            llmModel     = $Script:State.Prefs.LlmModel
            llmApiKey    = $Script:State.Prefs.LlmApiKey
            llmTimeout   = $Script:State.Prefs.LlmTimeout
            symbolPath   = $Script:State.Prefs.SymbolPath
        }
    }
    try {
        [System.IO.File]::WriteAllText($Script:PrefsPath, ($obj | ConvertTo-Json -Depth 6), [System.Text.UTF8Encoding]::new($true))
    } catch {}
}

# ---------------------------------------------------------------- Background runspace
function Start-BackgroundJob {
    param(
        [Parameter(Mandatory)] [scriptblock] $Body,
        [Parameter(Mandatory)] [object[]]    $Arguments,
        [Parameter(Mandatory)] [scriptblock] $OnDone,
        [string] $Kind = 'Job'
    )
    if ($Script:State.Job) {
        Write-DebugLog 'Another job is already running.' 'WARN'
        return
    }

    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions = 'ReuseThread'
    $rs.Open()
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($Body)
    foreach ($a in $Arguments) { [void]$ps.AddArgument($a) }
    $ar = $ps.BeginInvoke()

    $Script:State.Job = @{
        PS     = $ps
        Async  = $ar
        Timer  = $null
        OnDone = $OnDone
        Kind   = $Kind
    }

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(250)
    $timer.Add_Tick({
        $job = $Script:State.Job
        if (-not $job)            { return }
        if (-not $job.Async.IsCompleted) { return }
        $job.Timer.Stop()
        try {
            # Drain output, verbose/warning streams BEFORE EndInvoke disposes them
            foreach ($s in @($job.PS.Streams.Verbose))     { Write-DebugLog $s.Message 'DEBUG' }
            foreach ($s in @($job.PS.Streams.Information)) { Write-DebugLog ([string]$s.MessageData) 'INFO' }
            foreach ($s in @($job.PS.Streams.Warning))     { Write-DebugLog $s.Message 'WARN' }
            $result = $null
            try { $result = $job.PS.EndInvoke($job.Async) | Select-Object -Last 1 } catch {
                Write-DebugLog "Job EndInvoke failed: $($_.Exception.Message)" 'ERROR'
            }
            try { & $job.OnDone $result } catch {
                Write-DebugLog "OnDone callback crashed: $($_.Exception.Message)" 'ERROR'
                Write-DebugLog $_.ScriptStackTrace 'ERROR'
            }
        } finally {
            try { $job.PS.Runspace.Close() } catch {}
            try { $job.PS.Dispose() } catch {}
            $Script:State.Job = $null
            Set-Busy -Busy $false
        }
    })
    $Script:State.Job.Timer = $timer
    $timer.Start()
}

# ---------------------------------------------------------------- LLM client
# Runtime-agnostic OpenAI-compatible HTTP client. Works with Ollama (default,
# port 11434), LM Studio (1234), llamafile (8080), Foundry Local (5273), and
# any hosted endpoint (OpenAI, GitHub Models, Azure OpenAI, OpenRouter).
#
# Why one path for all: every modern local + cloud LLM runtime exposes the
# same POST /v1/chat/completions schema; only base URL + (optional) Bearer
# token change. So we keep the integration in the analyzer trivial: build a
# prompt (we already emit *.llm-prompt.md), POST it, render the response.

$Script:LlmCandidates = @(
    @{ Name='Ollama';         Url='http://localhost:11434/v1' }
    @{ Name='LM Studio';      Url='http://localhost:1234/v1'  }
    @{ Name='llamafile';      Url='http://localhost:8080/v1'  }
    @{ Name='Foundry Local';  Url='http://localhost:5273/v1'  }
)

function Test-LlmEndpoint {
    # Quick ping to /models. Returns the first model id reported, or $null.
    param([string]$BaseUrl, [string]$ApiKey = '')
    if (-not $BaseUrl) { return $null }
    try {
        $headers = @{}
        if ($ApiKey) { $headers['Authorization'] = "Bearer $ApiKey" }
        $r = Invoke-RestMethod -Uri "$BaseUrl/models" -Headers $headers -TimeoutSec 2 -ErrorAction Stop
        if ($r.data -and $r.data.Count -gt 0) { return [string]$r.data[0].id }
        return ''   # endpoint live but no models advertised
    } catch { return $null }
}

function Detect-LlmRuntime {
    # Probes localhost candidates. Returns the first that responds, or $null.
    # For Ollama, also probes /api/ps to detect CPU-only vs GPU offload.
    foreach ($c in $Script:LlmCandidates) {
        $model = Test-LlmEndpoint -BaseUrl $c.Url
        if ($null -ne $model) {
            $result = [pscustomobject]@{ Name=$c.Name; Url=$c.Url; FirstModel=$model; Processor=''; SizeMB=0; CpuOnly=$false }
            # Ollama-specific: probe /api/ps for active model + processor info.
            # Also check /api/show for the configured model if /api/ps is empty
            # (model not loaded yet — it loads on first inference request).
            if ($c.Name -eq 'Ollama') {
                try {
                    $baseNoV1 = $c.Url -replace '/v1$',''
                    $ps = Invoke-RestMethod -Uri "$baseNoV1/api/ps" -TimeoutSec 2 -ErrorAction Stop
                    if ($ps.models -and $ps.models.Count -gt 0) {
                        $active = $ps.models[0]
                        # size_vram > 50% of total size = GPU offload; otherwise CPU
                        $vram = if ($active.PSObject.Properties['size_vram']) { [double]$active.size_vram } else { 0 }
                        $total = if ($active.PSObject.Properties['size']) { [double]$active.size } else { 0 }
                        $result.Processor = if ($vram -gt 0 -and $total -gt 0 -and $vram -ge ($total * 0.5)) { 'GPU' } else { 'CPU' }
                        $result.SizeMB    = if ($total -gt 0) { [int]($total / 1MB) } else { 0 }
                        $result.CpuOnly   = ($result.Processor -eq 'CPU')
                    } else {
                        # No model loaded — assume CPU-only as the safe default.
                        # Machines with a capable GPU will have fast inference anyway;
                        # the warning only matters for CPU-only users who'd wait 10 min.
                        $result.Processor = 'CPU (assumed — no model loaded yet)'
                        $result.CpuOnly   = $true
                    }
                } catch {}
            }
            return $result
        }
    }
    return $null
}

function Invoke-LlmChat {
    # Synchronous POST -- only call from inside a background runspace.
    param(
        [Parameter(Mandatory)] [string] $BaseUrl,
        [Parameter(Mandatory)] [string] $Model,
        [string] $ApiKey = '',
        [Parameter(Mandatory)] [string] $Prompt,
        [int]    $TimeoutSec = 600
    )
    $sysMsg = 'You are a Windows performance engineer. Respond in concise Markdown. Cite pattern IDs (e.g. MEM-102) when you reference findings. Do not invent numbers; if a field is missing say so. Do NOT output your thinking process, reasoning steps, internal monologue, or chain-of-thought. Start directly with the analysis — no preamble, no headers like "Defining…" or "Analyzing…".'

    # Detect Ollama by URL and use its native /api/chat endpoint so we can
    # pass options.num_ctx reliably. The /v1/chat/completions shim silently
    # drops num_ctx because it's not in the OpenAI spec.
    $isOllama = $BaseUrl -match 'localhost:11434|127\.0\.0\.1:11434'
    if ($isOllama) {
        $nativeUrl = ($BaseUrl -replace '/v1$','') + '/api/chat'
        $body = [pscustomobject]@{
            model    = $Model
            messages = @(
                @{ role='system'; content = $sysMsg },
                @{ role='user';   content = $Prompt }
            )
            stream   = $false
            options  = @{
                temperature = 0.2
                num_ctx     = 16384
            }
        }
        $headers = @{ 'Content-Type' = 'application/json' }
        $json = $body | ConvertTo-Json -Depth 8
        $resp = Invoke-RestMethod -Uri $nativeUrl -Method Post -Headers $headers -Body $json -TimeoutSec $TimeoutSec -ErrorAction Stop
        if ($resp.message -and $resp.message.content) {
            return [string]$resp.message.content
        }
        return $null
    }

    # Non-Ollama: standard OpenAI-compatible /v1/chat/completions
    $body = [pscustomobject]@{
        model       = $Model
        messages    = @(
            @{ role='system'; content = $sysMsg },
            @{ role='user';   content = $Prompt }
        )
        temperature = 0.2
        stream      = $false
    }
    $headers = @{ 'Content-Type' = 'application/json' }
    if ($ApiKey) { $headers['Authorization'] = "Bearer $ApiKey" }
    $json = $body | ConvertTo-Json -Depth 8
    $resp = Invoke-RestMethod -Uri "$BaseUrl/chat/completions" -Method Post -Headers $headers -Body $json -TimeoutSec $TimeoutSec -ErrorAction Stop
    if ($resp.choices -and $resp.choices.Count -gt 0) {
        return [string]$resp.choices[0].message.content
    }
    return $null
}

# ---------------------------------------------------------------- Markdown → FlowDocument renderer
# Lightweight Markdown renderer for the AI response pane. Handles the subset
# that LLMs actually produce: ## headers, **bold**, `code`, ```code blocks```,
# - bullet lists, | tables |, and plain paragraphs.
function Render-MarkdownToDocument {
    param([string]$Markdown, [System.Windows.Documents.FlowDocument]$Doc)
    $Doc.Blocks.Clear()

    $bodyBrush   = $Window.TryFindResource('ThemeTextBody')
    $mutedBrush  = $Window.TryFindResource('ThemeTextMuted')
    $accentBrush = $Window.TryFindResource('ThemeAccentLight')
    $codeBgBrush = $Window.TryFindResource('ThemeDeepBg')
    $monoFont    = [System.Windows.Media.FontFamily]::new('Cascadia Mono, Consolas')
    $bodyFont    = [System.Windows.Media.FontFamily]::new('Segoe UI')

    $lines = $Markdown -split "`n" | ForEach-Object { $_.TrimEnd("`r") }
    $i = 0
    while ($i -lt $lines.Count) {
        $line = $lines[$i]

        # Fenced code block
        if ($line -match '^```') {
            $codeLines = [System.Collections.Generic.List[string]]::new()
            $i++
            while ($i -lt $lines.Count -and $lines[$i] -notmatch '^```') {
                [void]$codeLines.Add($lines[$i])
                $i++
            }
            $i++ # skip closing ```
            $para = [System.Windows.Documents.Paragraph]::new()
            $para.FontFamily = $monoFont
            $para.FontSize = 11.5
            $para.Foreground = $accentBrush
            $para.Background = $codeBgBrush
            $para.Padding = [System.Windows.Thickness]::new(12, 8, 12, 8)
            $para.Margin = [System.Windows.Thickness]::new(0, 4, 0, 8)
            $para.Inlines.Add([System.Windows.Documents.Run]::new(($codeLines -join "`r`n")))
            $Doc.Blocks.Add($para)
            continue
        }

        # Headings (## or ###)
        if ($line -match '^(#{1,4})\s+(.+)$') {
            $level = $Matches[1].Length
            $text  = $Matches[2]
            $para  = [System.Windows.Documents.Paragraph]::new()
            $para.FontFamily = $bodyFont
            $para.FontWeight = 'Bold'
            $para.FontSize = switch ($level) { 1 { 18 } 2 { 15 } 3 { 13 } default { 12.5 } }
            $para.Foreground = $Window.TryFindResource('ThemeTextPrimary')
            $para.Margin = [System.Windows.Thickness]::new(0, 10, 0, 4)
            $para.Inlines.Add([System.Windows.Documents.Run]::new($text))
            $Doc.Blocks.Add($para)
            $i++; continue
        }

        # Bullet list item
        if ($line -match '^\s*[-*]\s+(.+)$') {
            $text = $Matches[1]
            $para = [System.Windows.Documents.Paragraph]::new()
            $para.FontFamily = $bodyFont
            $para.FontSize = 12.5
            $para.Foreground = $bodyBrush
            $para.Margin = [System.Windows.Thickness]::new(16, 1, 0, 1)
            $para.TextIndent = -14
            Add-InlineRuns -Paragraph $para -Text "• $text"
            $Doc.Blocks.Add($para)
            $i++; continue
        }

        # Table row (| col | col |)
        if ($line -match '^\|.+\|$') {
            # Skip separator rows (|---|---|)
            if ($line -match '^\|[\s\-:|]+\|$') { $i++; continue }
            $cells = ($line.Trim('|') -split '\|') | ForEach-Object { $_.Trim() }
            $para = [System.Windows.Documents.Paragraph]::new()
            $para.FontFamily = $monoFont
            $para.FontSize = 11.5
            $para.Foreground = $bodyBrush
            $para.Margin = [System.Windows.Thickness]::new(0, 0, 0, 0)
            $para.Inlines.Add([System.Windows.Documents.Run]::new(($cells -join '  │  ')))
            $Doc.Blocks.Add($para)
            $i++; continue
        }

        # Blank line
        if (-not $line.Trim()) { $i++; continue }

        # Regular paragraph
        $para = [System.Windows.Documents.Paragraph]::new()
        $para.FontFamily = $bodyFont
        $para.FontSize = 12.5
        $para.Foreground = $bodyBrush
        $para.Margin = [System.Windows.Thickness]::new(0, 2, 0, 4)
        Add-InlineRuns -Paragraph $para -Text $line
        $Doc.Blocks.Add($para)
        $i++
    }
}

function Add-InlineRuns {
    param([System.Windows.Documents.Paragraph]$Paragraph, [string]$Text)
    $bodyBrush  = $Window.TryFindResource('ThemeTextBody')
    $accentBrush = $Window.TryFindResource('ThemeAccentLight')
    $monoFont   = [System.Windows.Media.FontFamily]::new('Cascadia Mono, Consolas')

    # Split on **bold**, `code`, and plain text segments
    $parts = [regex]::Split($Text, '(\*\*[^*]+\*\*|`[^`]+`)')
    foreach ($part in $parts) {
        if (-not $part) { continue }
        if ($part -match '^\*\*(.+)\*\*$') {
            $run = [System.Windows.Documents.Run]::new($Matches[1])
            $run.FontWeight = 'Bold'
            $run.Foreground = $Window.TryFindResource('ThemeTextPrimary')
            $Paragraph.Inlines.Add($run)
        } elseif ($part -match '^`(.+)`$') {
            $run = [System.Windows.Documents.Run]::new($Matches[1])
            $run.FontFamily = $monoFont
            $run.Foreground = $accentBrush
            $run.FontSize = 11.5
            $Paragraph.Inlines.Add($run)
        } else {
            $run = [System.Windows.Documents.Run]::new($part)
            $run.Foreground = $bodyBrush
            $Paragraph.Inlines.Add($run)
        }
    }
}

function Set-AiResponseText {
    param([string]$Text, [switch]$Plain)
    if ($Plain -or $Text.Length -lt 20) {
        # Short messages (errors, status) — just clear and add a plain paragraph
        $ui.fdAiResponse.Blocks.Clear()
        $para = [System.Windows.Documents.Paragraph]::new()
        $para.Foreground = $Window.TryFindResource('ThemeTextMuted')
        $para.FontFamily = [System.Windows.Media.FontFamily]::new('Cascadia Mono, Consolas')
        $para.FontSize = 12
        $para.Inlines.Add([System.Windows.Documents.Run]::new($Text))
        $ui.fdAiResponse.Blocks.Add($para)
    } else {
        Render-MarkdownToDocument -Markdown $Text -Doc $ui.fdAiResponse
    }
}

function Get-CurrentPromptText {
    # The Phase 1b pipeline emits <base>.llm-prompt.md alongside the report,
    # where <base> is the ETL basename without the trailing '.report' segment
    # (e.g. trace.etl -> trace.report.json + trace.llm-prompt.md).
    if (-not $Script:State.ReportPath) { return $null }
    $dir  = [IO.Path]::GetDirectoryName($Script:State.ReportPath)
    $base = [IO.Path]::GetFileNameWithoutExtension($Script:State.ReportPath) -replace '\.report$',''
    $p = Join-Path $dir ($base + '.llm-prompt.md')
    if (Test-Path -LiteralPath $p) { return [System.IO.File]::ReadAllText($p) }
    return $null
}

# ---------------------------------------------------------------- Confidence color
function Confidence-Brush {
    param([int]$Conf)
    if ($Conf -ge 70) { return $Window.FindResource('ThemeError') }
    if ($Conf -ge 40) { return $Window.FindResource('ThemeWarning') }
    return $Window.FindResource('ThemeSuccess')
}

# ---------------------------------------------------------------- Visualizations

# Curated palette for phase + family colour assignment. Stable order across
# runs so the same name always gets the same colour.
$Script:VizPalette = @(
    '#FF0078D4','#FF60CDFF','#FF00C853','#FFF59E0B','#FFFF5000',
    '#FFA78BFA','#FF14B8A6','#FFEC4899','#FF8B5CF6','#FF22D3EE',
    '#FFFB7185','#FFFACC15','#FF34D399','#FF818CF8','#FFF472B6')

function Get-VizBrush {
    param([string]$Key, [hashtable]$Map)
    if (-not $Map.ContainsKey($Key)) {
        $idx = $Map.Count % $Script:VizPalette.Count
        $b = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.ColorConverter]::ConvertFromString($Script:VizPalette[$idx]))
        $b.Freeze()
        $Map[$Key] = $b
    }
    return $Map[$Key]
}

function Populate-BootTimeline {
    param($Report)
    $rowsPanel = $ui.bootTimelineRows
    $rowsPanel.Children.Clear()
    $ui.bootTimelineLegend.ItemsSource = $null
    $ui.lblBootTotal.Text = ''

    $tree = $null
    if ($Report.PSObject.Properties['Visualizations'] -and $Report.Visualizations -and $Report.Visualizations.BootPhases) {
        $tree = @($Report.Visualizations.BootPhases)
    }
    if (-not $tree -or $tree.Count -eq 0) { return }

    # Group by Depth for fast lookup. Compute depth from Phase if missing
    # (older summaries) so the renderer survives schema drift.
    $byDepth = @{}
    foreach ($p in $tree) {
        $d = if ($p.PSObject.Properties['Depth']) { [int]$p.Depth } else {
            ([string]$p.Phase -split '\\').Count - 1
        }
        if (-not $byDepth.ContainsKey($d)) { $byDepth[$d] = [System.Collections.Generic.List[object]]::new() }
        [void]$byDepth[$d].Add($p)
    }

    $colourMap = @{}
    $legend = [System.Collections.Generic.List[object]]::new()
    $maxTiers = 4         # cap depth so device-enumeration rows don't drown the chart
    $heights  = @(34, 26, 22, 18)
    $totalRoot = 0.0

    # Start from depth 1 (Full Boot is depth 0 and conveys nothing visual).
    # Each tier shows children of the LONGEST segment in the parent tier
    # that actually has children. If we just pick the longest segment
    # unconditionally we can hit a dead-end at tier 1 (e.g. Post Boot is
    # the largest direct child but carries no sub-phases, so the tier-2
    # row would never render). Iterate longest-first until we find one
    # with children, then descend.
    #
    # All tiers render against the same X axis (root duration) so a child
    # tier's bar sits exactly under its parent segment from the previous
    # tier. Leading/trailing padding columns hold the empty space; the
    # ThemeDeepBg row background fills the gap visually.
    $parentPath  = $null
    $parentStart = 0.0   # parent's left edge, in root-coordinate seconds
    $parentEnd   = 0.0   # parent's right edge, in root-coordinate seconds
    for ($tier = 1; $tier -le $maxTiers; $tier++) {
        $rows = $null
        if ($byDepth.ContainsKey($tier)) {
            $rows = @($byDepth[$tier] | Where-Object {
                if ($null -eq $parentPath) { $true } else { [string]$_.Parent -eq $parentPath }
            })
        }
        if (-not $rows -or $rows.Count -eq 0) { break }

        # Render rows in the order they appear in the trace (Phase path)
        # so adjacent tiers visually nest. Sorting by duration broke the
        # parent-under-child alignment.
        $rows = @($rows | Sort-Object Phase)
        $tierTotal = ($rows | Measure-Object DurationSec -Sum).Sum
        if ($tierTotal -le 0) { break }
        if ($tier -eq 1) {
            $totalRoot   = $tierTotal
            $parentStart = 0.0
            $parentEnd   = $totalRoot
            $ui.lblBootTotal.Text = ("Full Boot: {0:N1}s" -f $tierTotal)
        }

        $caption = if ($parentPath) { "$(($parentPath -split '\\')[-1]) → children" }
                   else            { "Full Boot — direct children" }
        $tb = New-Object System.Windows.Controls.TextBlock
        $tb.Text = $caption
        $tb.Foreground = $ui.lblBootTotal.Foreground
        $tb.FontSize = 10
        $tb.Opacity  = 0.7
        $tb.Margin   = '0,8,0,2'
        [void]$rowsPanel.Children.Add($tb)

        $g = New-Object System.Windows.Controls.Grid
        $g.Height = $heights[[Math]::Min($tier - 1, $heights.Count - 1)]
        # SetResourceReference keeps the background in sync if the theme is
        # toggled at runtime (.Background = brush would snapshot the value).
        $g.SetResourceReference([System.Windows.Controls.Grid]::BackgroundProperty, 'ThemeDeepBg')

        # Leading padding (transparent) so this tier's bar starts under its
        # parent segment in the previous tier.
        $leadSec  = [Math]::Max(0.0, $parentStart)
        $trailSec = [Math]::Max(0.0, $totalRoot - $parentEnd)
        if ($leadSec -gt 0) {
            $padL = New-Object System.Windows.Controls.ColumnDefinition
            $padL.Width = [System.Windows.GridLength]::new($leadSec, [System.Windows.GridUnitType]::Star)
            $g.ColumnDefinitions.Add($padL)
        }

        # Track each row's [start, end] in root-coordinate space so we can
        # locate the next tier's parent without re-walking the tree.
        $childRanges = @{}
        $cursor = $parentStart
        foreach ($p in $rows) {
            $leaf  = if ($p.PSObject.Properties['Leaf'] -and $p.Leaf) { [string]$p.Leaf } else { ([string]$p.Phase -split '\\')[-1] }
            $sec   = [double]$p.DurationSec
            $brush = Get-VizBrush -Key $leaf -Map $colourMap

            $col = New-Object System.Windows.Controls.ColumnDefinition
            $col.Width = [System.Windows.GridLength]::new($sec, [System.Windows.GridUnitType]::Star)
            $g.ColumnDefinitions.Add($col)

            $rect = New-Object System.Windows.Controls.Border
            $rect.Background = $brush
            [System.Windows.Controls.Grid]::SetColumn($rect, $g.ColumnDefinitions.Count - 1)
            $pct = if ($tierTotal -gt 0) { $sec / $tierTotal * 100 } else { 0 }
            $rect.ToolTip = ("{0}: {1:N2}s ({2:N1}% of parent)" -f $leaf, $sec, $pct)
            [void]$g.Children.Add($rect)

            $childRanges[[string]$p.Phase] = @{ Start = $cursor; End = $cursor + $sec }
            $cursor += $sec

            if ($tier -le 2) {
                # Only add legend entries for the top two tiers; deeper tiers
                # would blow out the WrapPanel and the tooltips suffice.
                [void]$legend.Add([pscustomobject]@{
                    Label   = $leaf
                    Brush   = $brush
                    DurText = ("{0:N1}s" -f $sec)
                })
            }
        }

        if ($trailSec -gt 0) {
            $padR = New-Object System.Windows.Controls.ColumnDefinition
            $padR.Width = [System.Windows.GridLength]::new($trailSec, [System.Windows.GridUnitType]::Star)
            $g.ColumnDefinitions.Add($padR)
        }

        [void]$rowsPanel.Children.Add($g)

        # Pick the longest child that HAS children of its own as the parent
        # for the next tier. Falls back to "no next tier" if none qualifies.
        $nextTier = $tier + 1
        $nextParent = $null
        if ($byDepth.ContainsKey($nextTier)) {
            $rowsByDur = @($rows | Sort-Object DurationSec -Descending)
            foreach ($cand in $rowsByDur) {
                $candPath = [string]$cand.Phase
                $hasKids = $byDepth[$nextTier] | Where-Object { [string]$_.Parent -eq $candPath } | Select-Object -First 1
                if ($hasKids) { $nextParent = $cand; break }
            }
        }
        if (-not $nextParent) { break }
        $parentPath  = [string]$nextParent.Phase
        $range       = $childRanges[$parentPath]
        $parentStart = [double]$range.Start
        $parentEnd   = [double]$range.End
    }

    $ui.bootTimelineLegend.ItemsSource = $legend
}

function Populate-SystemTiles {
    param($Report)
    $ui.systemSummaryTiles.ItemsSource = $null
    if (-not $Report) { return }

    $tiles = [System.Collections.Generic.List[object]]::new()

    $accent  = $Window.TryFindResource('ThemeTextPrimary')
    $warn    = $Window.TryFindResource('ThemeAccent')
    $danger  = $Window.TryFindResource('SeverityHigh')
    if (-not $danger) { $danger = $warn }
    if (-not $accent) { $accent = $warn }

    # ---- Trace tile ----
    $meta = $Report.Manifest.TraceMeta
    $etlName = if ($meta -and $meta.EtlPath) { [System.IO.Path]::GetFileName([string]$meta.EtlPath) } else { '—' }
    $etlMB   = if ($meta -and $meta.EtlBytes) { [Math]::Round([double]$meta.EtlBytes / 1MB, 0) } else { 0 }
    [void]$tiles.Add([pscustomobject]@{
        Caption    = 'TRACE'
        Value      = $etlName
        Sub        = if ($etlMB -gt 0) { ("{0:N0} MB" -f $etlMB) } else { '' }
        ValueBrush = $accent
        Tooltip    = if ($meta) { [string]$meta.EtlPath } else { '' }
    })

    # ---- Captured tile ----
    $cap = $null
    if ($Report.PSObject.Properties['Visualizations'] -and $Report.Visualizations -and $Report.Visualizations.PSObject.Properties['CapturedUtc']) {
        $cap = [string]$Report.Visualizations.CapturedUtc
    }
    $capText = '—'
    $capSub  = ''
    if ($cap) {
        try {
            $dt = [datetime]::Parse($cap).ToLocalTime()
            $capText = $dt.ToString('yyyy-MM-dd HH:mm')
            $capSub  = $dt.ToString('ddd zzz')
        } catch { $capText = $cap }
    }
    [void]$tiles.Add([pscustomobject]@{
        Caption    = 'CAPTURED'
        Value      = $capText
        Sub        = $capSub
        ValueBrush = $accent
        Tooltip    = $cap
    })

    # ---- Duration tile (Full Boot DurationSec) ----
    $dur = 0.0
    if ($Report.PSObject.Properties['Visualizations'] -and $Report.Visualizations.BootPhases) {
        $root = @($Report.Visualizations.BootPhases | Where-Object {
            $d = if ($_.PSObject.Properties['Depth']) { [int]$_.Depth } else { 0 }
            $d -eq 0
        }) | Select-Object -First 1
        if ($root) { $dur = [double]$root.DurationSec }
    }
    $durText = if ($dur -gt 0) { ("{0:N1} s" -f $dur) } else { '—' }
    [void]$tiles.Add([pscustomobject]@{
        Caption    = 'BOOT DURATION'
        Value      = $durText
        Sub        = if ($dur -ge 60) { ("{0:N1} min" -f ($dur / 60)) } else { '' }
        ValueBrush = $accent
        Tooltip    = "Full Boot region duration (from Regions_of_Interest)"
    })

    # ---- Boot health tile ----
    $bh = if ($Report.Manifest.BootHealth) { [string]$Report.Manifest.BootHealth } else { 'Unknown' }
    $bhBrush = if ($bh -eq 'Unhealthy') { $danger } else { $accent }
    [void]$tiles.Add([pscustomobject]@{
        Caption    = 'BOOT HEALTH'
        Value      = $bh
        Sub        = ''
        ValueBrush = $bhBrush
        Tooltip    = 'Heuristic — long boot regions or DEF/DSK pattern hits flip this to Unhealthy.'
    })

    # ---- Memory tile ----
    $totalMB = 0.0
    if ($Report.PSObject.Properties['Visualizations'] -and $Report.Visualizations.Memory -and $Report.Visualizations.Memory.PSObject.Properties['TotalPhysicalMB'] -and $Report.Visualizations.Memory.TotalPhysicalMB) {
        $totalMB = [double]$Report.Visualizations.Memory.TotalPhysicalMB
    }
    $memText = if ($totalMB -gt 0) { ("{0:N1} GB" -f ($totalMB / 1024)) } else { '—' }
    [void]$tiles.Add([pscustomobject]@{
        Caption    = 'TOTAL RAM'
        Value      = $memText
        Sub        = if ($totalMB -gt 0) { ("{0:N0} MB" -f $totalMB) } else { '' }
        ValueBrush = $accent
        Tooltip    = 'Total Physical Memory from Memory_Utilization summary.'
    })

    # ---- Findings tile ----
    $fc = @($Report.Findings).Count
    $sc = @($Report.Suppressed).Count
    $pl = if ($Report.Manifest.PatternsLoaded) { [int]$Report.Manifest.PatternsLoaded } else { 0 }
    [void]$tiles.Add([pscustomobject]@{
        Caption    = 'FINDINGS'
        Value      = "$fc"
        Sub        = "$sc suppressed / $pl patterns"
        ValueBrush = $accent
        Tooltip    = "$fc findings, $sc suppressed; $pl patterns in catalog"
    })

    # ---- System tiles (only when wpaexporter -sysconfig General captured them) ----
    $sc2 = if ($Report.PSObject.Properties['Visualizations']) { $Report.Visualizations.SystemConfig } else { $null }
    # PSCustomObject (not a hashtable or array): single sysconfig record.
    if ($sc2 -and $sc2 -isnot [System.Collections.IEnumerable]) {
        $os  = if ($sc2.PSObject.Properties['ProductName']) { [string]$sc2.ProductName } else { '' }
        $bld = if ($sc2.PSObject.Properties['OsBuild'])     { [string]$sc2.OsBuild }     else { '' }
        if ($os) {
            [void]$tiles.Add([pscustomobject]@{
                Caption    = 'OS'
                Value      = $os
                Sub        = if ($bld) { "Build $bld" } else { '' }
                ValueBrush = $accent
                Tooltip    = if ($sc2.PSObject.Properties['BuildLab']) { [string]$sc2.BuildLab } else { '' }
            })
        }
        $cpu = if ($sc2.PSObject.Properties['ProcessorName']) { [string]$sc2.ProcessorName } else { '' }
        $cpuN = if ($sc2.PSObject.Properties['ProcessorCount']) { [string]$sc2.ProcessorCount } else { '' }
        if ($cpu) {
            $cpuShort = $cpu -replace '\s+\d+-Core Processor\s*$','' -replace 'AMD\s+','' -replace 'Intel\(R\)\s+','' -replace '\(R\)','' -replace '\(TM\)',''
            [void]$tiles.Add([pscustomobject]@{
                Caption    = 'CPU'
                Value      = $cpuShort.Trim()
                Sub        = if ($cpuN) { "$cpuN logical" } else { '' }
                ValueBrush = $accent
                Tooltip    = $cpu
            })
        }
        $man = if ($sc2.PSObject.Properties['SystemManufacturer']) { [string]$sc2.SystemManufacturer } else { '' }
        $prod = if ($sc2.PSObject.Properties['SystemProduct'])     { [string]$sc2.SystemProduct }     else { '' }
        if ($man -or $prod) {
            [void]$tiles.Add([pscustomobject]@{
                Caption    = 'HARDWARE'
                Value      = if ($man) { $man } else { $prod }
                Sub        = $prod
                ValueBrush = $accent
                Tooltip    = "$man $prod"
            })
        }
    }

    $ui.systemSummaryTiles.ItemsSource = $tiles
}

function Populate-MemoryBar {
    param($Report)
    $g = $ui.memoryBarGrid
    $g.ColumnDefinitions.Clear(); $g.Children.Clear()
    $ui.memoryLegend.ItemsSource = $null
    $ui.lblMemoryTotal.Text = ''

    if (-not $Report.PSObject.Properties['Visualizations'] -or -not $Report.Visualizations.Memory) { return }
    $mem = $Report.Visualizations.Memory
    $totalMB = 0.0
    if ($mem.PSObject.Properties['TotalPhysicalMB'] -and $mem.TotalPhysicalMB) { $totalMB = [double]$mem.TotalPhysicalMB }
    if ($totalMB -le 0) { return }
    $ui.lblMemoryTotal.Text = ("Total: {0:N0} MB" -f $totalMB)

    # Categories the JSON round-trip surfaces as either Hashtable or PSCustomObject.
    $cats = @{}
    if ($mem.PSObject.Properties['Categories']) {
        $catsSrc = $mem.Categories
        if ($catsSrc -is [System.Collections.IDictionary]) {
            foreach ($k in $catsSrc.Keys) { $cats[[string]$k] = [double]$catsSrc[$k] }
        } else {
            foreach ($pp in $catsSrc.PSObject.Properties) { $cats[[string]$pp.Name] = [double]$pp.Value }
        }
    }
    $get = { param($k) if ($cats.ContainsKey($k)) { [double]$cats[$k] } else { 0.0 } }

    $active   = & $get 'Active List'
    $standby  = & $get 'Standby Lists (Total)'
    $free     = & $get 'Zero and Free Lists'
    $npp      = & $get 'Non-Paged Pool Commit'
    $paged    = & $get 'Paged Pool Commit'
    $known    = $active + $standby + $free + $npp + $paged
    $other    = [Math]::Max(0.0, $totalMB - $known)

    $segments = @(
        @{ Label='Active';     MB=$active;  Color='#FF0078D4' },
        @{ Label='Standby';    MB=$standby; Color='#FF60CDFF' },
        @{ Label='Free';       MB=$free;    Color='#FF14B8A6' },
        @{ Label='NPP';        MB=$npp;     Color='#FFF59E0B' },
        @{ Label='Paged Pool'; MB=$paged;   Color='#FFA78BFA' },
        @{ Label='Other';      MB=$other;   Color='#FF6B7280' }
    )

    $legend = [System.Collections.Generic.List[object]]::new()
    foreach ($s in $segments) {
        if ($s.MB -le 0) { continue }
        $brush = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.ColorConverter]::ConvertFromString($s.Color))
        $brush.Freeze()

        $col = New-Object System.Windows.Controls.ColumnDefinition
        $col.Width = [System.Windows.GridLength]::new([double]$s.MB, [System.Windows.GridUnitType]::Star)
        $g.ColumnDefinitions.Add($col)

        $rect = New-Object System.Windows.Controls.Border
        $rect.Background = $brush
        [System.Windows.Controls.Grid]::SetColumn($rect, $g.ColumnDefinitions.Count - 1)
        $rect.ToolTip = ("{0}: {1:N0} MB ({2:N1}%)" -f $s.Label, $s.MB, ($s.MB / $totalMB * 100))
        [void]$g.Children.Add($rect)

        [void]$legend.Add([pscustomobject]@{
            Label   = $s.Label
            Brush   = $brush
            ValText = ("{0:N0} MB ({1:N1}%)" -f $s.MB, ($s.MB / $totalMB * 100))
        })
    }
    $ui.memoryLegend.ItemsSource = $legend

    # ---- Standby sub-list breakdown (priority 0 .. 7) ----
    $sbGrid = $ui.memoryStandbyGrid
    $sbGrid.ColumnDefinitions.Clear(); $sbGrid.Children.Clear()
    $ui.memoryStandbyLegend.ItemsSource = $null
    $sbVisible = 'Collapsed'

    if ($standby -gt 0) {
        # Priority 0 = highest priority (most likely to be reused) -> 7 = lowest.
        # Colour-grade from accent to muted so the ordering is visual.
        $sbColors = @(
            '#FF60CDFF','#FF38BDF8','#FF22D3EE','#FF14B8A6',
            '#FF84CC16','#FFA3A3A3','#FF6B7280','#FF4B5563'
        )
        $sbLegend = [System.Collections.Generic.List[object]]::new()
        $sbAny = $false
        for ($i = 0; $i -lt 8; $i++) {
            $mb = & $get ("Standby Sub-List $i")
            if ($mb -le 0) { continue }
            $sbAny = $true
            $brush = [System.Windows.Media.SolidColorBrush]::new(
                [System.Windows.Media.ColorConverter]::ConvertFromString($sbColors[$i]))
            $brush.Freeze()

            $col = New-Object System.Windows.Controls.ColumnDefinition
            $col.Width = [System.Windows.GridLength]::new([double]$mb, [System.Windows.GridUnitType]::Star)
            $sbGrid.ColumnDefinitions.Add($col)

            $rect = New-Object System.Windows.Controls.Border
            $rect.Background = $brush
            [System.Windows.Controls.Grid]::SetColumn($rect, $sbGrid.ColumnDefinitions.Count - 1)
            $rect.ToolTip = ("Standby Sub-List {0}: {1:N0} MB ({2:N1}% of Standby)" -f $i, $mb, ($mb / $standby * 100))
            [void]$sbGrid.Children.Add($rect)

            [void]$sbLegend.Add([pscustomobject]@{
                Label   = ("p" + $i)
                Brush   = $brush
                ValText = ("{0:N0} MB" -f $mb)
            })
        }
        if ($sbAny) {
            $ui.memoryStandbyLegend.ItemsSource = $sbLegend
            $sbVisible = 'Visible'
        }
    }
    $ui.lblMemoryStandby.Visibility    = $sbVisible
    $ui.memoryStandbyGrid.Visibility   = $sbVisible
    $ui.memoryStandbyLegend.Visibility = $sbVisible
}

function Populate-DiskHistogram {
    param($Report)
    $ui.diskHistBars.ItemsSource = $null
    $ui.lblDiskHistTotal.Text = ''

    if (-not $Report.PSObject.Properties['Visualizations'] -or -not $Report.Visualizations.DiskLatencyHistogramMs) { return }
    $hist = $Report.Visualizations.DiskLatencyHistogramMs

    # OrderedDictionary survives JSON round-trip as PSCustomObject; iterate props.
    $order = @('0-1 ms','1-5 ms','5-25 ms','25-100 ms','>100 ms')
    $values = [ordered]@{}
    foreach ($k in $order) {
        $v = 0.0
        if ($hist -is [System.Collections.IDictionary]) {
            if ($hist.Contains($k)) { $v = [double]$hist[$k] }
        } else {
            $pp = $hist.PSObject.Properties[$k]
            if ($pp) { $v = [double]$pp.Value }
        }
        $values[$k] = $v
    }

    $total = ($values.Values | Measure-Object -Sum).Sum
    if ($total -le 0) { return }
    $ui.lblDiskHistTotal.Text = ("Total: {0:N0} IOs" -f $total)

    $max = ($values.Values | Measure-Object -Maximum).Maximum
    $maxWidthPx = 220.0
    $colors = @{
        '0-1 ms'    = '#FF14B8A6'
        '1-5 ms'    = '#FF60CDFF'
        '5-25 ms'   = '#FFF59E0B'
        '25-100 ms' = '#FFFF5000'
        '>100 ms'   = '#FFEC4899'
    }

    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($k in $order) {
        $cnt = [double]$values[$k]
        $brush = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.ColorConverter]::ConvertFromString($colors[$k]))
        $brush.Freeze()
        $pct = if ($total -gt 0) { $cnt / $total * 100 } else { 0 }
        [void]$list.Add([pscustomobject]@{
            Label     = $k
            ValueText = ("{0:N0} ({1:N1}%)" -f $cnt, $pct)
            BarWidth  = [double][Math]::Max(2, ($cnt / [Math]::Max(1, $max)) * $maxWidthPx)
            Brush     = $brush
        })
    }
    $ui.diskHistBars.ItemsSource = $list
}



function Populate-FamilyPills {
    param($Report)
    $ui.familyPills.ItemsSource = $null
    if (-not $Report.Findings) { return }

    $byFam = $Report.Findings | Group-Object Family | Sort-Object Count -Descending
    $colourMap = @{}
    $pills = [System.Collections.Generic.List[object]]::new()
    foreach ($g in $byFam) {
        $topConf = ($g.Group | Measure-Object Confidence -Maximum).Maximum
        [void]$pills.Add([pscustomobject]@{
            Family    = $g.Name
            Brush     = Get-VizBrush -Key $g.Name -Map $colourMap
            CountText = "$($g.Count) finding$(if ($g.Count -ne 1) {'s'})"
            TopConf   = "Conf $topConf"
        })
    }
    $ui.familyPills.ItemsSource = $pills
}

function Populate-TopBars {
    param($Report, $Family, $TargetItemsControl, $TargetUnitLabel, [int]$N = 10)
    $TargetItemsControl.ItemsSource = $null
    $TargetUnitLabel.Text = ''

    $rows = @($Report.Findings + $Report.Suppressed) |
            Where-Object { $_.Family -eq $Family -and $_.PSObject.Properties['Detail'] -and $_.Detail -like 'Process *' } |
            Sort-Object Magnitude -Descending |
            Select-Object -First $N
    if ($rows.Count -eq 0) { return }

    $max  = [double]($rows | Measure-Object Magnitude -Maximum).Maximum
    if ($max -le 0) { return }
    $maxWidthPx = 220.0   # Bar pane width allowance; WPF will clamp.

    $unit = if ($rows[0].PSObject.Properties['Unit']) { [string]$rows[0].Unit } else { '' }
    $TargetUnitLabel.Text = $unit

    $brush = [System.Windows.Media.SolidColorBrush]::new(
        [System.Windows.Media.ColorConverter]::ConvertFromString('#FF0078D4'))
    $brush.Freeze()

    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($r in $rows) {
        # Strip 'Process ' prefix, PID suffix, and trailing unit annotation for label
        $label = ([string]$r.Detail) -replace '^Process\s+',''
        $label = $label -replace '\s*\(\d+\)\s*',''
        $label = $label -replace ':\s+[\d.,]+\s+.*$',''
        $value = [double]$r.Magnitude
        $val   = if ($value -lt 100) { "{0:N1}" -f $value } else { "{0:N0}" -f $value }
        [void]$list.Add([pscustomobject]@{
            Label     = $label
            ValueText = $val
            BarWidth  = [double][Math]::Max(2, ($value / $max) * $maxWidthPx)
            Brush     = $brush
        })
    }
    $TargetItemsControl.ItemsSource = $list
}

# ---------------------------------------------------------------- Findings scatter (D)
# Severity → brush. Looked up via TryFindResource so theme toggles cascade.
function Get-SeverityBrush {
    param([string]$Sev)
    $key = switch -Regex ($Sev) {
        'Critical' { 'SeverityCritical' }
        'High'     { 'SeverityHigh' }
        'Medium'   { 'SeverityMedium' }
        'Low'      { 'SeverityLow' }
        default    { 'SeverityInfo' }
    }
    $b = $Window.TryFindResource($key)
    if (-not $b) { $b = $Window.TryFindResource('ThemeAccent') }
    return $b
}

function Populate-FindingsScatter {
    param($Report)
    $cv = $ui.scatterCanvas
    $cv.Children.Clear()
    $ui.scatterLegend.ItemsSource = $null
    $ui.lblScatterMeta.Text = ''

    $rows = @($Report.Findings + $Report.Suppressed) | Where-Object { $_ }
    if ($rows.Count -eq 0) { return }

    # Canvas pixel layout: padding for axes; X = ImpactScore 0..100, Y = Confidence 0..100.
    # ActualWidth/Height are 0 until layout runs, so we use a fixed plot area
    # and let the parent Grid stretch the canvas around it. The dots are
    # placed in absolute coordinates; the Grid hosting the canvas has
    # Height=220 (set in XAML).
    $padL = 36; $padR = 12; $padT = 10; $padB = 22
    $plotW = 540.0  # nominal; canvas stretches but dots use this base
    $plotH = 220.0 - $padT - $padB

    # Axis baselines + gridlines (every 25 units)
    $gridBrush = $Window.TryFindResource('ThemeBorderSubtle')
    if (-not $gridBrush) { $gridBrush = [System.Windows.Media.Brushes]::DimGray }
    for ($i = 0; $i -le 4; $i++) {
        $x = $padL + ($i / 4.0) * $plotW
        $y = $padT + (1 - ($i / 4.0)) * $plotH
        $vert = New-Object System.Windows.Shapes.Line
        $vert.X1 = $x; $vert.X2 = $x; $vert.Y1 = $padT; $vert.Y2 = $padT + $plotH
        $vert.Stroke = $gridBrush; $vert.StrokeThickness = 0.5
        [void]$cv.Children.Add($vert)
        $horiz = New-Object System.Windows.Shapes.Line
        $horiz.X1 = $padL; $horiz.X2 = $padL + $plotW; $horiz.Y1 = $y; $horiz.Y2 = $y
        $horiz.Stroke = $gridBrush; $horiz.StrokeThickness = 0.5
        [void]$cv.Children.Add($horiz)
        # Axis labels (left = Y, bottom = X)
        $lblY = New-Object System.Windows.Controls.TextBlock
        $lblY.Text = "$([int]($i * 25))"
        $lblY.FontSize = 9
        $lblY.Foreground = $gridBrush
        $lblY.FontFamily = 'Cascadia Mono, Consolas'
        [System.Windows.Controls.Canvas]::SetLeft($lblY, 4)
        [System.Windows.Controls.Canvas]::SetTop($lblY, $y - 6)
        [void]$cv.Children.Add($lblY)
        $lblX = New-Object System.Windows.Controls.TextBlock
        $lblX.Text = "$([int]($i * 25))"
        $lblX.FontSize = 9
        $lblX.Foreground = $gridBrush
        $lblX.FontFamily = 'Cascadia Mono, Consolas'
        [System.Windows.Controls.Canvas]::SetLeft($lblX, $x - 8)
        [System.Windows.Controls.Canvas]::SetTop($lblX, $padT + $plotH + 4)
        [void]$cv.Children.Add($lblX)
    }
    # Y-axis title (rotated)
    $yt = New-Object System.Windows.Controls.TextBlock
    $yt.Text = "Confidence"
    $yt.Foreground = $gridBrush
    $yt.FontSize = 10
    $yt.LayoutTransform = New-Object System.Windows.Media.RotateTransform(-90)
    [System.Windows.Controls.Canvas]::SetLeft($yt, 2)
    [System.Windows.Controls.Canvas]::SetTop($yt, $padT + $plotH / 2 + 30)
    [void]$cv.Children.Add($yt)

    # Group counts per severity for legend
    $sevBuckets = @{ 'Critical'=0; 'High'=0; 'Medium'=0; 'Low'=0; 'Info'=0 }

    # Plot dots
    foreach ($f in $rows) {
        $impact = if ($f.PSObject.Properties['ImpactScore']) { [double]$f.ImpactScore } else { 0 }
        $conf   = if ($f.PSObject.Properties['Confidence'])  { [double]$f.Confidence }  else { 0 }
        $sev    = if ($f.PSObject.Properties['Severity'])    { [string]$f.Severity }    else { 'Info' }
        $brush  = Get-SeverityBrush -Sev $sev
        if ($sevBuckets.ContainsKey($sev)) { $sevBuckets[$sev]++ } else { $sevBuckets[$sev] = 1 }

        # Magnitude → dot radius. Magnitude varies wildly (ms .. GB), so use a
        # log scale clamped to 5..12 px.
        $mag = if ($f.PSObject.Properties['Magnitude']) { [Math]::Max(1.0, [double]$f.Magnitude) } else { 1.0 }
        $r   = [Math]::Max(5.0, [Math]::Min(12.0, 4 + [Math]::Log10($mag)))

        $x = $padL + ($impact / 100.0) * $plotW - $r
        $y = $padT + (1 - ($conf / 100.0)) * $plotH - $r

        $dot = New-Object System.Windows.Shapes.Ellipse
        $dot.Width = $r * 2; $dot.Height = $r * 2
        $dot.Fill = $brush
        $dot.Opacity = if ($f -in $Report.Suppressed) { 0.30 } elseif ($sev -eq 'Info') { 0.18 } else { 0.85 }
        $dot.Stroke = $brush
        $dot.StrokeThickness = 0.5
        $tipDetail = if ($f.Detail) { [string]$f.Detail } else { '' }
        $tipPC     = if ($f.PatternCode) { " [$($f.PatternCode)]" } else { '' }
        $dot.ToolTip = ("{0}/{1} {2}{3}`nImpact={4} Conf={5} Mag={6:N1}" -f $sev, $f.Family, $tipDetail, $tipPC, [int]$impact, [int]$conf, $mag)
        [System.Windows.Controls.Canvas]::SetLeft($dot, $x)
        [System.Windows.Controls.Canvas]::SetTop($dot, $y)
        [void]$cv.Children.Add($dot)
    }

    $legend = [System.Collections.Generic.List[object]]::new()
    foreach ($k in @('Critical','High','Medium','Low','Info')) {
        $c = $sevBuckets[$k]
        if (-not $c -or $c -le 0) { continue }
        [void]$legend.Add([pscustomobject]@{
            Label     = $k
            Brush     = Get-SeverityBrush -Sev $k
            CountText = "($c)"
        })
    }
    $ui.scatterLegend.ItemsSource = $legend
    $ui.lblScatterMeta.Text = ("{0} plotted ({1} findings + {2} suppressed)" -f $rows.Count, @($Report.Findings).Count, @($Report.Suppressed).Count)
}

# ---------------------------------------------------------------- Modules-by-signer donut (G)
function Populate-SignerDonut {
    param($Report)
    $cv = $ui.signerDonut
    $cv.Children.Clear()
    $ui.signerLegend.ItemsSource = $null
    $ui.signerTopVendors.ItemsSource = $null
    $ui.lblSignerTotal.Text = ''

    if (-not $Report.PSObject.Properties['Visualizations'] -or -not $Report.Visualizations.ImagesBySigner) { return }
    $bs = $Report.Visualizations.ImagesBySigner

    # Buckets survives JSON as either OrderedDictionary or PSCustomObject.
    $bucketOrder = @('Microsoft','Third-party','Unsigned / Unknown')
    $bucketColors = @{
        'Microsoft'           = '#FF0078D4'
        'Third-party'         = '#FFF59E0B'
        'Unsigned / Unknown'  = '#FF6B7280'
    }
    $values = [ordered]@{}
    foreach ($k in $bucketOrder) {
        $v = 0
        if ($bs.Buckets -is [System.Collections.IDictionary]) {
            if ($bs.Buckets.Contains($k)) { $v = [int]$bs.Buckets[$k] }
        } else {
            $pp = $bs.Buckets.PSObject.Properties[$k]
            if ($pp) { $v = [int]$pp.Value }
        }
        $values[$k] = $v
    }
    $total = ($values.Values | Measure-Object -Sum).Sum
    if ($total -le 0) { return }
    $ui.lblSignerTotal.Text = ("Total: {0:N0} images" -f $total)

    # Donut geometry: center at (180, 90); outer radius 80, inner 48 (cuts the hole).
    $cx = 180.0; $cy = 90.0
    $rOuter = 80.0; $rInner = 48.0
    $angle = -90.0 * [Math]::PI / 180.0  # start at 12 o'clock
    $legend = [System.Collections.Generic.List[object]]::new()
    foreach ($k in $bucketOrder) {
        $count = [double]$values[$k]
        if ($count -le 0) { continue }
        $sweepRad = ($count / $total) * 2 * [Math]::PI
        $a0 = $angle
        $a1 = $angle + $sweepRad
        $angle = $a1

        $brush = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.ColorConverter]::ConvertFromString($bucketColors[$k]))
        $brush.Freeze()

        # Build a Path with an arc segment for outer arc + inner arc back.
        $p1 = New-Object System.Windows.Point ($cx + $rOuter * [Math]::Cos($a0)), ($cy + $rOuter * [Math]::Sin($a0))
        $p2 = New-Object System.Windows.Point ($cx + $rOuter * [Math]::Cos($a1)), ($cy + $rOuter * [Math]::Sin($a1))
        $p3 = New-Object System.Windows.Point ($cx + $rInner * [Math]::Cos($a1)), ($cy + $rInner * [Math]::Sin($a1))
        $p4 = New-Object System.Windows.Point ($cx + $rInner * [Math]::Cos($a0)), ($cy + $rInner * [Math]::Sin($a0))
        $isLarge = ($sweepRad -gt [Math]::PI)

        $fig = New-Object System.Windows.Media.PathFigure
        $fig.StartPoint = $p1
        $arc1 = New-Object System.Windows.Media.ArcSegment
        $arc1.Point = $p2; $arc1.Size = New-Object System.Windows.Size $rOuter, $rOuter
        $arc1.SweepDirection = [System.Windows.Media.SweepDirection]::Clockwise
        $arc1.IsLargeArc = $isLarge
        $fig.Segments.Add($arc1)
        $line1 = New-Object System.Windows.Media.LineSegment
        $line1.Point = $p3
        $fig.Segments.Add($line1)
        $arc2 = New-Object System.Windows.Media.ArcSegment
        $arc2.Point = $p4; $arc2.Size = New-Object System.Windows.Size $rInner, $rInner
        $arc2.SweepDirection = [System.Windows.Media.SweepDirection]::Counterclockwise
        $arc2.IsLargeArc = $isLarge
        $fig.Segments.Add($arc2)
        $fig.IsClosed = $true

        $geom = New-Object System.Windows.Media.PathGeometry
        $geom.Figures.Add($fig)
        $path = New-Object System.Windows.Shapes.Path
        $path.Data = $geom
        $path.Fill = $brush
        $path.Stroke = $Window.TryFindResource('ThemeAppBg')
        $path.StrokeThickness = 2
        $pct = $count / $total * 100
        $path.ToolTip = ("{0}: {1:N0} ({2:N1}%)" -f $k, $count, $pct)
        [void]$cv.Children.Add($path)

        [void]$legend.Add([pscustomobject]@{
            Label   = $k
            Brush   = $brush
            ValText = ("{0:N0} ({1:N1}%)" -f $count, $pct)
        })
    }

    # Center label (total)
    $center = New-Object System.Windows.Controls.TextBlock
    $center.Text = "$total"
    $center.FontSize = 22
    $center.FontWeight = 'Bold'
    $center.Foreground = $Window.TryFindResource('ThemeTextPrimary')
    [System.Windows.Controls.Canvas]::SetLeft($center, $cx - 22)
    [System.Windows.Controls.Canvas]::SetTop($center, $cy - 18)
    [void]$cv.Children.Add($center)
    $sub = New-Object System.Windows.Controls.TextBlock
    $sub.Text = "modules"
    $sub.FontSize = 9
    $sub.Foreground = $Window.TryFindResource('ThemeTextMuted')
    [System.Windows.Controls.Canvas]::SetLeft($sub, $cx - 18)
    [System.Windows.Controls.Canvas]::SetTop($sub, $cy + 8)
    [void]$cv.Children.Add($sub)

    $ui.signerLegend.ItemsSource = $legend

    # Top 3rd-party vendors list
    if ($bs.PSObject.Properties['TopVendors'] -and $bs.TopVendors) {
        $ui.signerTopVendors.ItemsSource = @($bs.TopVendors)
    }
}

# ---------------------------------------------------------------- Treemap (B+C)
# Squarified treemap (Bruls/Huijsen/van Wijk, 2000). Pure WPF -- one Border
# per leaf, hover-tooltip per leaf, click-to-drill for hierarchical inputs.

$Script:TreemapState = @{
    Mode    = 'Boot'   # 'Boot' | 'Cpu' | 'Disk'
    RootPath= ''       # for Boot: current drill-down phase path
    Report  = $null
}
# Stable colour memo shared across treemap modes. Initialized at script
# load so StrictMode v3 reads in _Tm-BootEntriesAt and Populate-Treemap
# don't throw 'variable has not been set'.
$Script:_TmColourMap = @{}

function _Tm-Worst {
    param([double[]]$Row, [double]$W)
    if ($Row.Count -eq 0) { return [double]::PositiveInfinity }
    $s = ($Row | Measure-Object -Sum).Sum
    $rMax = ($Row | Measure-Object -Maximum).Maximum
    $rMin = ($Row | Measure-Object -Minimum).Minimum
    if ($s -le 0 -or $rMin -le 0) { return [double]::PositiveInfinity }
    $w2 = $W * $W
    $s2 = $s * $s
    return [Math]::Max(($w2 * $rMax) / $s2, $s2 / ($w2 * $rMin))
}

function _Tm-LayoutRow {
    param(
        [array]$Items, [int]$Start, [int]$Count,
        [double]$X, [double]$Y, [double]$W, [double]$H,
        [System.Collections.Generic.List[object]]$Out,
        [bool]$AlongWidth
    )
    $values = @($Items[$Start..($Start+$Count-1)] | ForEach-Object { [double]$_.Value })
    $sum = ($values | Measure-Object -Sum).Sum
    if ($sum -le 0) { return }
    if ($AlongWidth) {
        $rowH = $sum / $W
        $cx = $X
        for ($i = 0; $i -lt $Count; $i++) {
            $v = $values[$i]
            $cw = ($v / $sum) * $W
            [void]$Out.Add(@{ Item = $Items[$Start+$i]; X = $cx; Y = $Y; W = $cw; H = $rowH })
            $cx += $cw
        }
    } else {
        $rowW = $sum / $H
        $cy = $Y
        for ($i = 0; $i -lt $Count; $i++) {
            $v = $values[$i]
            $ch = ($v / $sum) * $H
            [void]$Out.Add(@{ Item = $Items[$Start+$i]; X = $X; Y = $cy; W = $rowW; H = $ch })
            $cy += $ch
        }
    }
}

function _Tm-Squarify {
    param(
        [array]$Items,                                 # sorted desc by Value, normalized so sum == X*H
        [double]$X, [double]$Y, [double]$W, [double]$H,
        [System.Collections.Generic.List[object]]$Out
    )
    $i = 0
    while ($i -lt $Items.Count) {
        $shortSide = [Math]::Min($W, $H)
        $alongWidth = ($W -le $H)
        # Greedily extend row while aspect ratio improves.
        $rowVals = [System.Collections.Generic.List[double]]::new()
        $bestRatio = [double]::PositiveInfinity
        $rowCount = 0
        while ($i + $rowCount -lt $Items.Count) {
            $candidate = [double]$Items[$i + $rowCount].Value
            $tmp = [double[]]@($rowVals.ToArray() + $candidate)
            $thisRatio = _Tm-Worst -Row $tmp -W $shortSide
            if ($thisRatio -gt $bestRatio -and $rowVals.Count -gt 0) { break }
            $bestRatio = $thisRatio
            [void]$rowVals.Add($candidate)
            $rowCount++
        }
        _Tm-LayoutRow -Items $Items -Start $i -Count $rowCount -X $X -Y $Y -W $W -H $H -Out $Out -AlongWidth $alongWidth
        # Advance the remaining rectangle.
        $sum = ($rowVals.ToArray() | Measure-Object -Sum).Sum
        if ($alongWidth) {
            $rowH = $sum / $W
            $Y += $rowH; $H -= $rowH
        } else {
            $rowW = $sum / $H
            $X += $rowW; $W -= $rowW
        }
        $i += $rowCount
    }
}

function Render-Treemap {
    param(
        [System.Windows.Controls.Canvas]$Canvas,
        [array]$Entries,             # array of pscustomobject { Label, Value, Brush, Data, Clickable }
        [scriptblock]$OnClick = $null
    )
    $Canvas.Children.Clear()
    if (-not $Entries -or $Entries.Count -eq 0) { return }

    # Canvas might not have measured yet -- fall back to parent ActualWidth.
    $w = if ($Canvas.ActualWidth -gt 0) { [double]$Canvas.ActualWidth } else {
        $parent = [System.Windows.Media.VisualTreeHelper]::GetParent($Canvas)
        if ($parent -and $parent.ActualWidth -gt 0) { [double]$parent.ActualWidth } else { 600.0 }
    }
    $h = if ($Canvas.ActualHeight -gt 0) { [double]$Canvas.ActualHeight } else {
        $parent = [System.Windows.Media.VisualTreeHelper]::GetParent($Canvas)
        if ($parent -and $parent.ActualHeight -gt 0) { [double]$parent.ActualHeight } else { 300.0 }
    }
    if ($w -lt 50 -or $h -lt 50) { return }

    # Sort + filter zero-values; normalize so total area == w*h.
    $items = @($Entries | Where-Object { [double]$_.Value -gt 0 } | Sort-Object -Property @{ Expression={ [double]$_.Value }; Descending=$true })
    if ($items.Count -eq 0) { return }
    $total = ($items | Measure-Object -Property Value -Sum).Sum
    $scale = ($w * $h) / $total
    $scaled = @($items | ForEach-Object {
        # Clone with scaled value for layout, keep original for tooltips.
        $copy = $_.PSObject.Copy()
        $copy | Add-Member -NotePropertyName _orig -NotePropertyValue ([double]$_.Value) -Force
        $copy.Value = [double]$_.Value * $scale
        $copy
    })

    $layout = [System.Collections.Generic.List[object]]::new()
    _Tm-Squarify -Items $scaled -X 0 -Y 0 -W $w -H $h -Out $layout

    foreach ($rec in $layout) {
        $rx = [double]$rec.X; $ry = [double]$rec.Y
        $rw = [double]$rec.W; $rh = [double]$rec.H
        if ($rw -lt 1 -or $rh -lt 1) { continue }
        $it = $rec.Item

        $border = New-Object System.Windows.Controls.Border
        $border.Background = $it.Brush
        $border.BorderBrush = $Window.TryFindResource('ThemeAppBg')
        $border.BorderThickness = '1'
        $border.Width  = $rw - 1
        $border.Height = $rh - 1
        [System.Windows.Controls.Canvas]::SetLeft($border, $rx)
        [System.Windows.Controls.Canvas]::SetTop($border, $ry)
        $tipVal = if ($it.PSObject.Properties['_orig']) { [double]$it._orig } else { [double]$it.Value }
        $border.ToolTip = ("{0}`n{1:N2} {2}" -f $it.Label, $tipVal, $(if ($it.PSObject.Properties['Unit']) { [string]$it.Unit } else { '' }))

        # Inline label on tiles big enough to read
        if ($rw -gt 60 -and $rh -gt 18) {
            $lbl = New-Object System.Windows.Controls.TextBlock
            $lbl.Text = [string]$it.Label
            $lbl.Foreground = [System.Windows.Media.Brushes]::White
            $lbl.FontSize = if ($rw -gt 140 -and $rh -gt 40) { 12 } else { 10 }
            $lbl.FontWeight = 'SemiBold'
            $lbl.Margin = '6,4,6,0'
            $lbl.TextTrimming = 'CharacterEllipsis'
            $border.Child = $lbl
        }
        if ($OnClick -and $it.Clickable) {
            $border.Cursor = [System.Windows.Input.Cursors]::Hand
            $border.Tag = $it
            # No .GetNewClosure(): $Script:-scoped handler refs vanish inside a
            # closure's dynamic module (see /memories/ps51-getclosure-pitfalls.md).
            # Call the named function directly; $s.Tag carries the item.
            $border.Add_MouseLeftButtonUp({
                param($s, $e)
                Invoke-TreemapClick $s.Tag
            })
        }
        [void]$Canvas.Children.Add($border)
    }
}

function _Tm-BootEntriesAt {
    param($Tree, [string]$ParentPath)
    # Children of $ParentPath (depth = ParentPath's depth + 1).
    # If $ParentPath is '' return depth-1 (root's children).
    $targetDepth = if ($ParentPath) { (($ParentPath -split '\\').Count) } else { 1 }
    $colourMap = $Script:_TmColourMap
    $rows = @($Tree | Where-Object {
        $d = if ($_.PSObject.Properties['Depth']) { [int]$_.Depth } else { ([string]$_.Phase -split '\\').Count - 1 }
        if ($d -ne $targetDepth) { return $false }
        if ($ParentPath -and [string]$_.Parent -ne $ParentPath) { return $false }
        $true
    })
    foreach ($r in $rows) {
        $leaf = if ($r.PSObject.Properties['Leaf'] -and $r.Leaf) { [string]$r.Leaf } else { ([string]$r.Phase -split '\\')[-1] }
        # Has children?
        $rPath = [string]$r.Phase
        $hasChild = @($Tree | Where-Object {
            $d2 = if ($_.PSObject.Properties['Depth']) { [int]$_.Depth } else { ([string]$_.Phase -split '\\').Count - 1 }
            $d2 -eq $targetDepth + 1 -and [string]$_.Parent -eq $rPath
        }).Count -gt 0
        [pscustomobject]@{
            Label     = "$leaf  ({0:N1}s)" -f ([double]$r.DurationSec)
            Value     = [double]$r.DurationSec
            Unit      = 's'
            Brush     = Get-VizBrush -Key $leaf -Map $colourMap
            Clickable = $hasChild
            Data      = $r
        }
    }
}

function Populate-Treemap {
    $r = $Script:TreemapState.Report
    if (-not $r) { $ui.treemapCanvas.Children.Clear(); return }

    $entries = $null
    $hint    = 'Hover for details.'
    $crumb   = ''
    switch ($Script:TreemapState.Mode) {
        'Boot' {
            $tree = if ($r.PSObject.Properties['Visualizations'] -and $r.Visualizations.BootPhases) { @($r.Visualizations.BootPhases) } else { @() }
            if (-not $Script:_TmColourMap) { $Script:_TmColourMap = @{} }
            $entries = @(_Tm-BootEntriesAt -Tree $tree -ParentPath $Script:TreemapState.RootPath)
            $hint = 'Click a phase with children to drill down. Right-click on the canvas to go back up.'
            $crumb = if ($Script:TreemapState.RootPath) { "Path: $($Script:TreemapState.RootPath)" } else { "Path: Full Boot" }
        }
        'Cpu' {
            $list = if ($r.PSObject.Properties['Visualizations'] -and $r.Visualizations.CpuTop) { @($r.Visualizations.CpuTop) } else { @() }
            $unit = if ($r.Visualizations.PSObject.Properties['CpuUnit']) { [string]$r.Visualizations.CpuUnit } else { '%' }
            if (-not $Script:_TmColourMap) { $Script:_TmColourMap = @{} }
            $entries = @($list | ForEach-Object {
                [pscustomobject]@{
                    Label     = "$($_.Key)  ({0:N1})" -f ([double]$_.Value)
                    Value     = [double]$_.Value
                    Unit      = $unit
                    Brush     = Get-VizBrush -Key $_.Key -Map $Script:_TmColourMap
                    Clickable = $false
                }
            })
            $hint = "Top processes by $unit (no drill -- CPU summary is flat)."
            $crumb = "Showing top $($entries.Count) by $unit"
        }
        'Disk' {
            $list = if ($r.PSObject.Properties['Visualizations'] -and $r.Visualizations.DiskTop) { @($r.Visualizations.DiskTop) } else { @() }
            $unit = if ($r.Visualizations.PSObject.Properties['DiskUnit']) { [string]$r.Visualizations.DiskUnit } else { '' }
            if (-not $Script:_TmColourMap) { $Script:_TmColourMap = @{} }
            $entries = @($list | ForEach-Object {
                [pscustomobject]@{
                    Label     = "$($_.Key)  ({0:N0})" -f ([double]$_.Value)
                    Value     = [double]$_.Value
                    Unit      = $unit
                    Brush     = Get-VizBrush -Key $_.Key -Map $Script:_TmColourMap
                    Clickable = $false
                }
            })
            $hint = "Top processes by $unit (no drill -- Disk summary is flat)."
            $crumb = "Showing top $($entries.Count) by $unit"
        }
    }
    $ui.lblTreemapHint.Text       = $hint
    $ui.lblTreemapBreadcrumb.Text = $crumb
    Render-Treemap -Canvas $ui.treemapCanvas -Entries $entries -OnClick { $true }
}

# Click handler for treemap drill. Plain function so the WPF event handler
# can call it by name without going through $Script:-scoped variables (which
# are invisible inside .GetNewClosure(); see /memories notes).
function Invoke-TreemapClick {
    param($Item)
    if (-not $Item -or -not $Item.Data) { return }
    if ($Script:TreemapState.Mode -ne 'Boot') { return }
    $Script:TreemapState.RootPath = [string]$Item.Data.Phase
    Populate-Treemap
}

# ---------------------------------------------------------------- Pattern coverage donut (E)
function Populate-PatternDonut {
    param($Report)
    $cv = $ui.patternDonut
    $cv.Children.Clear()
    $ui.patternLegend.ItemsSource = $null
    $ui.patternFiredList.ItemsSource = $null
    if (-not $Report) { return }

    $loaded = if ($Report.Manifest.PatternsLoaded) { [int]$Report.Manifest.PatternsLoaded } else { 0 }
    $firedCodes = @($Report.Findings | Where-Object { $_.PatternCode } | Select-Object -ExpandProperty PatternCode -Unique)
    $supCodes   = @($Report.Suppressed | Where-Object { $_.PatternCode } | Select-Object -ExpandProperty PatternCode -Unique)
    $supOnly    = @($supCodes | Where-Object { $_ -notin $firedCodes })
    $fired = $firedCodes.Count
    $supOnlyCnt = $supOnly.Count
    $unused = [Math]::Max(0, $loaded - $fired - $supOnlyCnt)

    $segments = [ordered]@{
        'Fired'             = @{ Count = $fired;      Color='#FF00C853' }
        'Suppressed only'   = @{ Count = $supOnlyCnt; Color='#FFF59E0B' }
        'Loaded, no signal' = @{ Count = $unused;     Color='#FF6B7280' }
    }
    $total = ($segments.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
    if ($total -le 0) { return }

    $cx = 90.0; $cy = 90.0; $rOuter = 75.0; $rInner = 44.0
    $angle = -90.0 * [Math]::PI / 180.0
    $legend = [System.Collections.Generic.List[object]]::new()
    foreach ($k in $segments.Keys) {
        $cnt = [double]$segments[$k].Count
        if ($cnt -le 0) { continue }
        $sweepRad = ($cnt / $total) * 2 * [Math]::PI
        $a0 = $angle; $a1 = $angle + $sweepRad; $angle = $a1
        $brush = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.ColorConverter]::ConvertFromString($segments[$k].Color))
        $brush.Freeze()

        $p1 = New-Object System.Windows.Point ($cx + $rOuter * [Math]::Cos($a0)), ($cy + $rOuter * [Math]::Sin($a0))
        $p2 = New-Object System.Windows.Point ($cx + $rOuter * [Math]::Cos($a1)), ($cy + $rOuter * [Math]::Sin($a1))
        $p3 = New-Object System.Windows.Point ($cx + $rInner * [Math]::Cos($a1)), ($cy + $rInner * [Math]::Sin($a1))
        $p4 = New-Object System.Windows.Point ($cx + $rInner * [Math]::Cos($a0)), ($cy + $rInner * [Math]::Sin($a0))
        $isLarge = ($sweepRad -gt [Math]::PI)

        $fig = New-Object System.Windows.Media.PathFigure
        $fig.StartPoint = $p1
        $arc1 = New-Object System.Windows.Media.ArcSegment
        $arc1.Point = $p2; $arc1.Size = New-Object System.Windows.Size $rOuter,$rOuter
        $arc1.SweepDirection = [System.Windows.Media.SweepDirection]::Clockwise
        $arc1.IsLargeArc = $isLarge
        $fig.Segments.Add($arc1)
        $fig.Segments.Add((New-Object System.Windows.Media.LineSegment -Property @{ Point=$p3 }))
        $arc2 = New-Object System.Windows.Media.ArcSegment
        $arc2.Point = $p4; $arc2.Size = New-Object System.Windows.Size $rInner,$rInner
        $arc2.SweepDirection = [System.Windows.Media.SweepDirection]::Counterclockwise
        $arc2.IsLargeArc = $isLarge
        $fig.Segments.Add($arc2)
        $fig.IsClosed = $true

        $geom = New-Object System.Windows.Media.PathGeometry
        $geom.Figures.Add($fig)
        $path = New-Object System.Windows.Shapes.Path
        $path.Data = $geom; $path.Fill = $brush
        $path.Stroke = $Window.TryFindResource('ThemeAppBg'); $path.StrokeThickness = 2
        $pct = $cnt / $total * 100
        $path.ToolTip = ("{0}: {1:N0} ({2:N1}%)" -f $k, $cnt, $pct)
        [void]$cv.Children.Add($path)

        [void]$legend.Add([pscustomobject]@{
            Label   = $k
            Brush   = $brush
            ValText = ("{0:N0} ({1:N1}%)" -f $cnt, $pct)
        })
    }

    $center = New-Object System.Windows.Controls.TextBlock
    $center.Text = "$total"
    $center.FontSize = 22; $center.FontWeight = 'Bold'
    $center.Foreground = $Window.TryFindResource('ThemeTextPrimary')
    [System.Windows.Controls.Canvas]::SetLeft($center, $cx - 14)
    [System.Windows.Controls.Canvas]::SetTop($center, $cy - 18)
    [void]$cv.Children.Add($center)
    $sub = New-Object System.Windows.Controls.TextBlock
    $sub.Text = "patterns"
    $sub.FontSize = 9
    $sub.Foreground = $Window.TryFindResource('ThemeTextMuted')
    [System.Windows.Controls.Canvas]::SetLeft($sub, $cx - 18)
    [System.Windows.Controls.Canvas]::SetTop($sub, $cy + 8)
    [void]$cv.Children.Add($sub)

    # "N fired" badge just outside the donut ring at top-right
    if ($fired -gt 0) {
        $badgeBg = New-Object System.Windows.Controls.Border
        $badgeBg.Background = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.ColorConverter]::ConvertFromString('#FF00C853'))
        $badgeBg.CornerRadius = '9'
        $badgeBg.Padding = '7,2'
        $badgeLbl = New-Object System.Windows.Controls.TextBlock
        $badgeLbl.Text = "$fired fired"
        $badgeLbl.Foreground = [System.Windows.Media.Brushes]::White
        $badgeLbl.FontSize = 10; $badgeLbl.FontWeight = 'Bold'
        $badgeBg.Child = $badgeLbl
        [System.Windows.Controls.Canvas]::SetLeft($badgeBg, $cx + $rOuter - 6)
        [System.Windows.Controls.Canvas]::SetTop($badgeBg, $cy - $rOuter - 4)
        [void]$cv.Children.Add($badgeBg)
    }

    $ui.patternLegend.ItemsSource = $legend

    # Right-side: list of fired patterns (Code + Detail + Confidence)
    $firedList = @($Report.Findings |
        Where-Object { $_.PatternCode } |
        Group-Object PatternCode |
        ForEach-Object {
            $best = $_.Group | Sort-Object Confidence -Descending | Select-Object -First 1
            [pscustomobject]@{
                Code     = [string]$_.Name
                Title    = [string]$best.Detail
                ConfText = "conf $([int]$best.Confidence)"
            }
        } |
        Sort-Object Code)
    $ui.patternFiredList.ItemsSource = $firedList
}

function Populate-Overview {
    param($Report)
    $Script:TreemapState.Report   = $Report
    $Script:TreemapState.RootPath = ''
    $Script:TreemapState.Mode     = 'Boot'
    Populate-SystemTiles  -Report $Report
    Populate-BootTimeline -Report $Report
    Populate-MemoryBar    -Report $Report
    Populate-DiskHistogram -Report $Report
    Populate-FindingsScatter -Report $Report
    Populate-SignerDonut  -Report $Report
    Populate-FamilyPills  -Report $Report
    Populate-TopBars      -Report $Report -Family 'CpuHotspot' -TargetItemsControl $ui.topCpuBars  -TargetUnitLabel $ui.lblCpuUnit
    Populate-TopBars      -Report $Report -Family 'Disk'       -TargetItemsControl $ui.topDiskBars -TargetUnitLabel $ui.lblDiskUnit
    Populate-Treemap
    Populate-PatternDonut -Report $Report

    # Overview verdict banner
    $topF = $Report.Findings | Select-Object -First 1
    if ($topF) {
        $fLabel = switch ($topF.Family) {
            'Disk'          { 'DISK I/O CONTENTION' }
            'CpuHotspot'    { 'CPU HOTSPOT' }
            'BootPhase'     { 'BOOT PHASE DELAY' }
            'DpcIsr'        { 'DPC / ISR CONTENTION' }
            'ServiceStart'  { 'SERVICE STARTUP DELAY' }
            'Configuration' { 'CONFIGURATION ISSUE' }
            'DriverVersion' { 'DRIVER ISSUE' }
            default         { $topF.Family.ToUpper() }
        }
        $ui.lblOverviewVerdictFamily.Text = $fLabel
        $rootText = ([string]$topF.Detail) -replace '^Process\s+','' -replace '\s*\(\d+\)\s*',' '
        $ui.lblOverviewVerdictRoot.Text   = $rootText
        $ui.lblOverviewVerdictConf.Text   = "Conf $([int]$topF.Confidence)"
        $ui.overviewVerdictConf.Background = Confidence-Brush -Conf ([int]$topF.Confidence)
        $ui.overviewVerdictBanner.Visibility = 'Visible'
    } else {
        $ui.overviewVerdictBanner.Visibility = 'Collapsed'
    }

    $ui.lblOverviewMeta.Text = "boot=$($Report.Manifest.BootHealth) | families=$($Report.Manifest.EvidenceFamilies -join ', ') | $(@($Report.Findings).Count) findings"
    $ui.OverviewEmpty.Visibility   = 'Collapsed'
    $ui.OverviewContent.Visibility = 'Visible'
}

# ---------------------------------------------------------------- Subsystem classification
function Get-Subsystem {
    param($Report)
    $stackMods = @()
    $stackSource = if ($Report.Faulting.FullStack) { $Report.Faulting.FullStack } else { $Report.Faulting.TopStack }
    if ($stackSource) { $stackMods = @($stackSource | ForEach-Object { [string]$_.Module } | Where-Object { $_ -and $_ -notmatch '^0x' } | Select-Object -Unique) }
    $joined = ($stackMods -join '|').ToLowerInvariant()
    if ($joined -match 'opengl32|d3d|dxgi|igxel|nvoglv|amdxx|vulkan|wgl') { return 'GPU / Graphics' }
    if ($joined -match 'awt|swing|javafx|jfx')                            { return 'JVM / AWT UI' }
    if ($joined -match 'clr|coreclr|mscorlib|system\.') { return '.NET / CLR' }
    if ($joined -match 'ws2_32|mswsock|winhttp|wininet|nsi|dnsapi')        { return 'Networking' }
    if ($joined -match 'ntfs|fltmgr|ntos.*ki|fileinto')                    { return 'File I/O' }
    if ($joined -match 'user32|win32u|gdi32|uxtheme|dwm')                  { return 'UI / Windowing' }
    if ($joined -match 'rpcrt4|combase|ole32')                              { return 'COM / RPC' }
    if ($joined -match 'jvm|java_exe|java!|jni')                           { return 'JVM Runtime' }
    if ($joined -match 'ntdll|kernelbase|kernel32')                        { return 'NT Subsystem' }
    return 'Unknown'
}

function Get-ThreadRole {
    param([string[]]$Frames)
    $joined = ($Frames -join '|').ToLowerInvariant()
    if ($joined -match 'wtoolkit_eventloop|getmessage|peekmessage|dispatchmessage') { return 'UI' }
    if ($joined -match 'monitorwait|jvm_monitorwait')   { return 'Monitor' }
    if ($joined -match 'socketread|socketinput|recv|accept|winsock|mswsock') { return 'IO/Net' }
    if ($joined -match 'fileinputstream|readfile|ntreadfile') { return 'IO/File' }
    if ($joined -match 'jvm_sleep|sleep')               { return 'Sleep' }
    if ($joined -match 'gc|finali')                     { return 'GC' }
    if ($joined -match 'sendmessage|usercall')           { return 'SendMsg' }
    if ($joined -match 'waitforsingle|waitformultiple')  { return 'Wait' }
    return 'Worker'
}

function Group-ThreadStacks {
    param($ThreadStacks)
    $groups = [ordered]@{}
    foreach ($ts in $ThreadStacks) {
        $sig = (@($ts.Frames) -join ' > ')
        if (-not $groups.Contains($sig)) {
            $groups[$sig] = [pscustomobject]@{
                Count    = 0
                Role     = (Get-ThreadRole -Frames @($ts.Frames))
                TopFrame = if (@($ts.Frames).Count -gt 0) { [string]$ts.Frames[0] } else { '(empty)' }
                ThreadIds = [System.Collections.Generic.List[string]]::new()
                Frames   = @($ts.Frames)
            }
        }
        $groups[$sig].Count++
        [void]$groups[$sig].ThreadIds.Add([string]$ts.ThreadId)
    }
    return @($groups.Values | Sort-Object Count -Descending)
}

# ---------------------------------------------------------------- Render report
function Get-ConfidenceFromSeverity {
    param([string]$Severity)
    switch ([string]$Severity) {
        'Critical' { return 95 }
        'High'     { return 85 }
        'Medium'   { return 70 }
        'Low'      { return 55 }
        default    { return 40 }
    }
}

function Render-Report {
    param([string]$ReportPath)
    if (-not $ReportPath -or -not (Test-Path -LiteralPath $ReportPath)) {
        Write-DebugLog "Report not found: $ReportPath" 'ERROR'
        return
    }

    $report = Get-Content -LiteralPath $ReportPath -Raw | ConvertFrom-Json -Depth 64
    $Script:State.ReportPath = $ReportPath
    $Script:State.Report     = $report

    # --- Root-cause card (facts-only: show subsystem + bucket) ---
    $subsystem = Get-Subsystem -Report $report
    $ui.cardRoot.Visibility = 'Visible'
    $ui.lblRootFamily.Text  = $subsystem.ToUpperInvariant()
    $ui.lblRoot.Text        = [string]$report.Exception.Bucket
    $ui.lblWhy.Text         = "Process: $($report.Process.Name)`nException: $($report.Exception.Code) at $($report.Exception.Address)`nFaulting: $($report.Faulting.Module) ($($report.Faulting.Symbol))"
    $ui.lblConfidence.Text  = 'data'
    $ui.badgeConf.Background = Confidence-Brush -Conf 40
    $ui.badgePattern.Visibility = 'Collapsed'

    # --- Stack tab: full faulting stack with color type + dedup ---
    $stackRows = [System.Collections.Generic.List[object]]::new()
    $stackSource = if ($report.Faulting.FullStack) { $report.Faulting.FullStack } else { $report.Faulting.TopStack }
    $faultModLower = if ($report.Faulting.Module) { $report.Faulting.Module.ToLowerInvariant() } else { '' }
    # Detect duplicate stack segments — build signature for first half and check if it repeats
    $allFrames = @($stackSource)
    $halfLen = [Math]::Floor($allFrames.Count / 2)
    $dupStart = -1
    if ($halfLen -ge 8) {
        $firstSig = ($allFrames[0..($halfLen-1)] | ForEach-Object { [string]$_.Frame }) -join '|'
        for ($s = $halfLen; $s -le ($allFrames.Count - $halfLen); $s++) {
            $candidateSig = ($allFrames[$s..($s+$halfLen-1)] | ForEach-Object { [string]$_.Frame }) -join '|'
            if ($candidateSig -eq $firstSig) { $dupStart = $s; break }
        }
    }
    $idx = 0
    $inDup = $false
    foreach ($f in $allFrames) {
        if ($dupStart -gt 0 -and $idx -eq $dupStart -and -not $inDup) {
            # Insert a separator row for the duplicate segment
            [void]$stackRows.Add([pscustomobject]@{ Index = ''; Frame = '--- duplicate of frames 0-' + ($halfLen-1) + ' (capture artifact) ---'; Module = ''; Symbol = ''; ModuleType = 'Separator' })
            $inDup = $true
        }
        $mod = [string]$f.Module
        $modType = 'Unknown'
        if ($mod -and $mod -notmatch '^0x') {
            $ml = $mod.ToLowerInvariant()
            if ($ml -eq $faultModLower)      { $modType = 'Fault' }
            elseif ($ml -match '^(ntdll|kernel32|kernelbase|user32|win32u|gdi32|gdi32full)$') { $modType = 'System' }
            elseif ($ml -match 'opengl32|igxel|d3d|dxgi|nvoglv|amdxx') { $modType = 'GPU' }
            elseif ($ml -match '^(jvm|java_exe|java|awt|net|clr|coreclr)$') { $modType = 'App' }
            else { $modType = 'Other' }
        } elseif ($mod -match '^0x') { $modType = 'Unresolved' }
        if ($inDup) { $modType = 'Dimmed' }
        [void]$stackRows.Add([pscustomobject]@{ Index = $idx; Frame = [string]$f.Frame; Module = $mod; Symbol = [string]$f.Symbol; ModuleType = $modType })
        $idx++
    }
    $ui.dgFindings.ItemsSource = $stackRows

    # --- Threads tab: deduplicated + classified ---
    $threadRows = [System.Collections.Generic.List[object]]::new()
    if ($report.ThreadStacks) {
        $grouped = Group-ThreadStacks -ThreadStacks $report.ThreadStacks
        foreach ($g in $grouped) {
            $label = if ($g.Count -gt 1) { "$($g.Count)x $($g.Role)" } else { $g.Role }
            $ids   = ($g.ThreadIds -join ', ')
            [void]$threadRows.Add([pscustomobject]@{
                Group    = $label
                Count    = $g.Count
                Role     = $g.Role
                TopFrame = $g.TopFrame
                Threads  = $ids
            })
        }
    }
    $ui.dgSuppressed.ItemsSource = $threadRows

    # --- Facts tab: key dump facts as readable items ---
    $factItems = [System.Collections.Generic.List[object]]::new()
    # Exception
    [void]$factItems.Add([pscustomobject]@{ Category = 'Exception'; Label = 'Code';    Value = [string]$report.Exception.Code })
    [void]$factItems.Add([pscustomobject]@{ Category = 'Exception'; Label = 'Address'; Value = [string]$report.Exception.Address })
    [void]$factItems.Add([pscustomobject]@{ Category = 'Exception'; Label = 'Bucket';  Value = [string]$report.Exception.Bucket })
    # Process
    if ($report.ProcessInfo) {
        [void]$factItems.Add([pscustomobject]@{ Category = 'Process'; Label = 'Exe';     Value = [string]$report.ProcessInfo.ExeImagePath })
        [void]$factItems.Add([pscustomobject]@{ Category = 'Process'; Label = 'WorkDir'; Value = [string]$report.ProcessInfo.CurrentDirectory })
    }
    [void]$factItems.Add([pscustomobject]@{ Category = 'Process'; Label = 'Threads'; Value = [string]$report.ThreadCount })
    # Event
    if ($report.LastEvent) {
        [void]$factItems.Add([pscustomobject]@{ Category = 'Event'; Label = 'LastEvent'; Value = [string]$report.LastEvent })
    }
    if ($report.LastError -and $report.LastError.Code) {
        [void]$factItems.Add([pscustomobject]@{ Category = 'Event'; Label = 'LastError'; Value = [string]$report.LastError.Code })
    }
    # Timing
    if ($report.Timing) {
        [void]$factItems.Add([pscustomobject]@{ Category = 'Timing'; Label = 'KernelTime'; Value = [string]$report.Timing.KernelTime })
        [void]$factItems.Add([pscustomobject]@{ Category = 'Timing'; Label = 'UserTime';   Value = [string]$report.Timing.UserTime })
    }
    # Symbols
    if ($report.SymbolQuality) {
        [void]$factItems.Add([pscustomobject]@{ Category = 'Symbols'; Label = 'Resolved'; Value = "$($report.SymbolQuality.StackFramesResolved) / $($report.SymbolQuality.StackFramesTotal)" })
    }
    # VAS
    if ($report.VASummary) {
        foreach ($prop in $report.VASummary.PSObject.Properties) {
            [void]$factItems.Add([pscustomobject]@{ Category = 'VAS'; Label = $prop.Name; Value = [string]$prop.Value })
        }
    }
    # Registers (key ones)
    if ($report.RegisterCorrelation) {
        foreach ($rc in $report.RegisterCorrelation) {
            $modTag = if ($rc.Module) { " -> $($rc.Module)" } else { ' (unmapped)' }
            [void]$factItems.Add([pscustomobject]@{ Category = 'Registers'; Label = [string]$rc.Register; Value = [string]$rc.Value + $modTag })
        }
    }
    # Heap
    if ($report.HeapCorrupt) {
        [void]$factItems.Add([pscustomobject]@{ Category = 'Heap'; Label = 'Corruption'; Value = 'DETECTED' })
    }
    $ui.lstPatterns.ItemsSource = $factItems

    # --- Manifest tab: raw JSON ---
    $ui.txtManifest.Text = ($report | ConvertTo-Json -Depth 12)

    # --- Co-firing pills: show key modules ---
    $peerPills = [System.Collections.Generic.List[object]]::new()
    if ($report.ModuleDetails) {
        foreach ($md in ($report.ModuleDetails | Select-Object -First 5)) {
            [void]$peerPills.Add([pscustomobject]@{ Label = [string]$md.Name; ValueText = [string]$md.SymbolStatus })
        }
    }
    $ui.lstCoFiring.ItemsSource = $peerPills

    $symPct = if ($report.SymbolQuality -and $report.SymbolQuality.StackFramesTotal -gt 0) {
        [Math]::Round(100 * $report.SymbolQuality.StackFramesResolved / $report.SymbolQuality.StackFramesTotal, 0)
    } else { 0 }
    $ui.lblSubtitle.Text = "$subsystem | $($report.Dump.Kind) | $($report.Process.Name) | stack $($stackRows.Count) | threads $($report.ThreadCount) | symbols $symPct%"

    Append-History ([pscustomobject]@{
        WhenUtc    = (Get-Date).ToUniversalTime().ToString('o')
        WhenLocal  = (Get-Date).ToString('yyyy-MM-dd HH:mm')
        Mode       = 'Dump'
        DumpPath   = $Script:State.DumpPath
        ReportPath = $ReportPath
        DumpKind   = [string]$report.Dump.Kind
        RootCause  = [string]$report.Exception.Bucket
        Confidence = 0
    })

    if ($ui.lblAiVerdictRoot) {
        $ui.lblAiVerdictRoot.Text  = [string]$report.Exception.Bucket
        $ui.lblAiVerdictWhy.Text   = "Exception: $($report.Exception.Code) | Module: $($report.Faulting.Module) | Symbol: $($report.Faulting.Symbol)"
        $ui.lblAiVerdictBadge.Text = "facts-only / $($report.AnalysisMode)"
        $ui.lstAiVerdictPeers.ItemsSource = @()
    }

    Refresh-History
    Show-Page 'Single'
    Write-DebugLog "Report rendered. Bucket: $($report.Exception.Bucket), Stack: $($stackRows.Count) frames, Threads: $($report.ThreadCount)" 'SUCCESS'
}

function Refresh-History {
    $ui.dgHistory.ItemsSource = @(Load-History)
}

# ---------------------------------------------------------------- Page switching
$Script:Pages = @{
    Single   = $ui.PageSingle
    Overview = $ui.PageOverview
    Compare  = $ui.PageCompare
    Patterns = $ui.PagePatterns
    History  = $ui.PageHistory
    AI       = $ui.PageAI
    Settings = $ui.PageSettings
}
$Script:NavButtons = @{
    Single   = $ui.NavSingle
    Overview = $ui.NavOverview
    Compare  = $ui.NavCompare
    Patterns = $ui.NavPatterns
    History  = $ui.NavHistory
    AI       = $ui.NavAI
    Settings = $ui.NavSettings
}

# Hide trace-only and pattern pages for facts-only dump workflow.
$ui.NavOverview.Visibility  = 'Collapsed'
$ui.NavCompare.Visibility   = 'Collapsed'
$ui.NavPatterns.Visibility  = 'Collapsed'
$ui.PageOverview.Visibility = 'Collapsed'
$ui.PageCompare.Visibility  = 'Collapsed'
$ui.PagePatterns.Visibility = 'Collapsed'

function Show-Page {
    param([string]$Name)
    foreach ($p in $Script:Pages.GetEnumerator()) {
        $p.Value.Visibility = if ($p.Key -eq $Name) { 'Visible' } else { 'Collapsed' }
    }
    foreach ($b in $Script:NavButtons.GetEnumerator()) {
        if ($b.Key -eq $Name) { $b.Value.Tag = 'Active' } else { $b.Value.Tag = $null }
    }
    Write-DebugLog "Page: $Name" 'DEBUG'
}

# ---------------------------------------------------------------- Event wiring

# Title bar drag + window controls
$ui.TitleBar.Add_MouseLeftButtonDown({
    if ($_.ClickCount -eq 2) {
        $Window.WindowState = if ($Window.WindowState -eq 'Maximized') { 'Normal' } else { 'Maximized' }
    } else {
        try { $Window.DragMove() } catch {}
    }
})
$ui.btnMinimize.Add_Click({ $Window.WindowState = 'Minimized' })
$ui.btnMaximize.Add_Click({ $Window.WindowState = if ($Window.WindowState -eq 'Maximized') { 'Normal' } else { 'Maximized' } })
$ui.btnClose.Add_Click({ $Window.Close() })

# WindowStyle=None maximize fix: constrain to work area so the taskbar isn't clipped
$Window.Add_StateChanged({
    if ($Window.WindowState -eq 'Maximized') {
        # Get the work area of the screen the window is on
        $screen = [System.Windows.Forms.Screen]::FromHandle(
            (New-Object System.Windows.Interop.WindowInteropHelper($Window)).Handle)
        $dpi = [System.Windows.Media.VisualTreeHelper]::GetDpi($Window)
        $Window.MaxWidth  = $screen.WorkingArea.Width  / $dpi.DpiScaleX
        $Window.MaxHeight = $screen.WorkingArea.Height / $dpi.DpiScaleY
    } else {
        $Window.MaxWidth  = [double]::PositiveInfinity
        $Window.MaxHeight = [double]::PositiveInfinity
    }
})

# Nav rail
$ui.NavSingle.Add_Click(  { Show-Page 'Single' })
$ui.NavOverview.Add_Click({ Show-Page 'Overview' })
$ui.NavCompare.Add_Click( { Show-Page 'Compare' })
$ui.NavPatterns.Add_Click({ Show-Page 'Patterns' })
$ui.NavHistory.Add_Click( { Show-Page 'History'; Refresh-History })
$ui.NavAI.Add_Click(      { Show-Page 'AI'; Update-AiBanner })
$ui.NavSettings.Add_Click({ Show-Page 'Settings' })

# Treemap mode switches
$ui.rbTreemapBoot.Add_Checked({ $Script:TreemapState.Mode = 'Boot'; $Script:TreemapState.RootPath = ''; Populate-Treemap })
$ui.rbTreemapCpu.Add_Checked( { $Script:TreemapState.Mode = 'Cpu';  Populate-Treemap })
$ui.rbTreemapDisk.Add_Checked({ $Script:TreemapState.Mode = 'Disk'; Populate-Treemap })

# Right-click anywhere in the treemap canvas pops one drill level (boot only).
$ui.treemapCanvas.Add_MouseRightButtonUp({
    if ($Script:TreemapState.Mode -ne 'Boot') { return }
    $p = $Script:TreemapState.RootPath
    if (-not $p) { return }
    $parts = $p -split '\\'
    if ($parts.Count -le 1) {
        $Script:TreemapState.RootPath = ''
    } else {
        $Script:TreemapState.RootPath = ($parts[0..($parts.Count - 2)] -join '\')
    }
    Populate-Treemap
})

# Re-layout when the canvas resizes (initial render runs before measure).
$ui.treemapCanvas.Add_SizeChanged({ Populate-Treemap })
$ui.NavToggleConsole.Add_Click({
    $Script:State.Prefs.ShowLog = -not $Script:State.Prefs.ShowLog
    $ui.LogConsole.Visibility = if ($Script:State.Prefs.ShowLog) { 'Visible' } else { 'Collapsed' }
    Save-Prefs
})
$ui.NavTheme.Add_Click({
    $Script:State.Prefs.IsLightMode = -not $Script:State.Prefs.IsLightMode
    Apply-CurrentTheme
    $ui.chkLightMode.IsChecked = $Script:State.Prefs.IsLightMode
    Save-Prefs
})

# Browse
$ui.btnBrowse.Add_Click({
    $f = Show-FilePicker -Title 'Select dump'
    if ($f) { $ui.txtEtl.Text = $f }
})
$ui.btnBrowseBaseline.Add_Click({
    $f = Show-FilePicker -Title 'Select baseline dump'
    if ($f) { $ui.txtBaseline.Text = $f }
})
$ui.btnBrowseCandidate.Add_Click({
    $f = Show-FilePicker -Title 'Select candidate dump'
    if ($f) { $ui.txtCandidate.Text = $f }
})

# Analyze (DumpPilot pipeline)
$ui.btnAnalyze.Add_Click({
    $dmp = $ui.txtEtl.Text
    if (-not $dmp -or -not (Test-Path -LiteralPath $dmp)) {
        Write-DebugLog 'Please pick a valid .dmp file.' 'WARN'
        return
    }
    $Script:State.DumpPath = $dmp
    Write-DebugLog "Starting analysis: $dmp" 'STEP'
    Set-Busy -Busy $true -Status 'Analyzing...'
    $invokePath = Join-Path $Script:Here 'Invoke-DumpPilotPipeline.ps1'

    Start-BackgroundJob `
        -Kind 'Analyze' `
        -Body {
            param($Invoke, $Dmp, $ForceFlag, $SymPath)
            $ErrorActionPreference = 'Stop'
            try {
                $pArgs = @{ DumpPath = $Dmp; Force = $ForceFlag }
                if ($SymPath) { $pArgs['SymbolPath'] = $SymPath }
                $r = & $Invoke @pArgs -Verbose 4>&1 |
                     Where-Object { $_ -is [pscustomobject] -and $_.PSObject.Properties['ReportJson'] } |
                     Select-Object -Last 1
                if (-not $r) { throw 'Pipeline produced no result object.' }
                [pscustomobject]@{
                    Success    = $true
                    ReportPath = [string]$r.ReportJson
                    ReportMd   = [string]$r.ReportMd
                    ReportHtml = [string]$r.ReportHtml
                    PromptMd   = [string]$r.PromptMd
                    DumpKind   = [string]$r.DumpKind
                    Process    = [string]$r.Process
                    Bucket     = [string]$r.Bucket
                    TopHit     = [string]$r.TopHit
                    HitCount   = [int]$r.HitCount
                }
            } catch {
                [pscustomobject]@{ Success = $false; Error = $_.Exception.Message; Stack = $_.ScriptStackTrace }
            }
        } `
        -Arguments @($invokePath, $dmp, [bool]$ui.chkForce.IsChecked, [string]$Script:State.Prefs.SymbolPath) `
        -OnDone {
            param($r)
            if (-not $r) { Write-DebugLog 'Analysis returned no result' 'ERROR'; return }
            if ($r.Success) {
                try { Render-Report -ReportPath $r.ReportPath } catch { Write-DebugLog "Render-Report failed: $($_.Exception.Message)" 'ERROR' }
                $detail = "$($r.DumpKind) | $($r.Process) | $($r.Bucket) | top: $($r.TopHit) ($($r.HitCount) hit(s))"
                Set-Status -Text 'Done' -Tone 'Success' -Detail $detail
                Write-DebugLog "HTML: $($r.ReportHtml)" 'INFO'
                Write-DebugLog "Prompt: $($r.PromptMd)" 'INFO'
                if ($r.ReportHtml -and (Test-Path -LiteralPath $r.ReportHtml)) {
                    try { Start-Process -FilePath $r.ReportHtml } catch {}
                }
            } else {
                Write-DebugLog "Analysis failed: $($r.Error)" 'ERROR'
                if ($r.Stack) { Write-DebugLog $r.Stack 'ERROR' }
                Set-Status -Text 'Failed' -Tone 'Error' -Detail $r.Error
            }
        }
})

# Compare
$ui.btnCompare.Add_Click({
    $b = $ui.txtBaseline.Text
    $c = $ui.txtCandidate.Text
    if (-not $b -or -not (Test-Path -LiteralPath $b)) { Write-DebugLog 'Baseline not found.' 'ERROR'; return }
    if (-not $c -or -not (Test-Path -LiteralPath $c)) { Write-DebugLog 'Candidate not found.' 'ERROR'; return }
    Write-DebugLog "Comparing: $b  vs  $c" 'STEP'
    Set-Busy -Busy $true -Status 'Comparing...'
    $comparePath = Join-Path $Script:Helpers 'Compare-Traces.ps1'

    Start-BackgroundJob `
        -Kind 'Compare' `
        -Body {
            param($Cmp, $B, $C)
            $ErrorActionPreference = 'Stop'
            try {
                $out = & $Cmp -Baseline $B -Candidate $C -Verbose 4>&1
                $p = $out | Where-Object { $_ -is [string] -and $_ -like '*.comparison.report.json' } | Select-Object -Last 1
                [pscustomobject]@{ Success = $true; Path = $p }
            } catch {
                [pscustomobject]@{ Success = $false; Error = $_.Exception.Message }
            }
        } `
        -Arguments @($comparePath, $b, $c) `
        -OnDone {
            param($r)
            if ($r -and $r.Success -and $r.Path) {
                $cmp = Get-Content -LiteralPath $r.Path -Raw | ConvertFrom-Json -Depth 32
                $ui.dgDelta.ItemsSource = @($cmp.Deltas)
                Set-Status -Text 'Comparison ready' -Tone 'Success' -Detail $r.Path
            } else {
                $err = if ($r) { $r.Error } else { 'no result' }
                Write-DebugLog "Compare failed: $err" 'ERROR'
                Set-Status -Text 'Compare failed' -Tone 'Error' -Detail $err
            }
        }
})

# Root cause action buttons
$ui.btnOpenInWpa.Add_Click({
    if (-not $Script:State.DumpPath) { return }
    Start-Process explorer.exe "/select,`"$($Script:State.DumpPath)`""
})
$ui.btnOpenReportMd.Add_Click({
    if (-not $Script:State.ReportPath) { return }
    $md = [IO.Path]::ChangeExtension($Script:State.ReportPath, '.md')
    if (Test-Path -LiteralPath $md) { Start-Process -FilePath $md } else { Write-DebugLog 'Report.md not found' 'WARN' }
})
$ui.btnOpenTables.Add_Click({
    if (-not $Script:State.DumpPath) { return }
    $t = "$($Script:State.DumpPath).dump-facts"
    if (Test-Path -LiteralPath $t) { Start-Process explorer.exe $t } else { Write-DebugLog 'Tables folder not found' 'WARN' }
})

# Open HTML report in browser (generated automatically by pipeline stage 4)
$ui.btnExportHtml.Add_Click({
    if (-not $Script:State.Report) { Write-DebugLog 'No report loaded yet' 'WARN'; return }
    # Find the HTML file next to the report JSON
    $htmlPath = $null
    if ($Script:State.ReportPath) {
        $htmlPath = [System.IO.Path]::ChangeExtension($Script:State.ReportPath, '.html')
        if (-not (Test-Path -LiteralPath $htmlPath)) {
            # Try the pipeline output path
            $rptItem = Get-Item -LiteralPath $Script:State.ReportPath -ErrorAction SilentlyContinue
            if ($rptItem) {
                $htmlPath = Join-Path $rptItem.DirectoryName ($rptItem.BaseName + '.html')
            }
        }
    }
    if ($htmlPath -and (Test-Path -LiteralPath $htmlPath)) {
        Start-Process -FilePath $htmlPath
        Write-DebugLog "Opened HTML report in browser: $htmlPath" 'INFO'
    } else {
        Write-DebugLog 'HTML report not found. Re-run analysis to generate it.' 'WARN'
    }
})

# Root-cause card -> AI Review shortcut
$ui.btnAiReview.Add_Click({
    if (-not $Script:State.ReportPath) { Write-DebugLog 'No report loaded yet' 'WARN'; return }
    Show-Page 'AI'
    # Kick off the review immediately so the user sees progress without an extra click
    $ui.btnAiRun.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
})

# ---- AI Review page ----
function Update-AiEndpointLabel {
    $base = $Script:State.Prefs.LlmBaseUrl
    $mdl  = $Script:State.Prefs.LlmModel
    $ui.lblAiEndpoint.Text = "$base  /  $mdl"
}
function Update-AiBanner {
    # Contextual banner on the AI page. Probes Ollama if local; shows
    # appropriate guidance based on what we actually detect.
    $base = $Script:State.Prefs.LlmBaseUrl
    $isLocal = $base -match 'localhost|127\.0\.0\.1'
    if (-not $isLocal) {
        $ui.lblAiBannerIcon.Foreground = $Window.TryFindResource('ThemeAccentLight')
        $ui.lblAiBannerText.Text = "Hosted endpoint configured. Copy prompt pastes the structured findings; Run review sends them directly to the API."
        return
    }
    # Probe Ollama
    $hit = Detect-LlmRuntime
    if (-not $hit) {
        $ui.lblAiBannerIcon.Foreground = $Window.TryFindResource('ThemeWarning')
        $ui.lblAiBannerText.Text = "No local LLM endpoint detected. Install Ollama (winget install Ollama.Ollama) and pull a model, or use Copy prompt to paste into Copilot Chat / Claude / Gemini."
        return
    }
    if ($hit.CpuOnly) {
        $ui.lblAiBannerIcon.Foreground = $Window.TryFindResource('ThemeWarning')
        $ui.lblAiBannerText.Text = "Local model running on CPU only — inference will be slow (5–10 min). For instant results, use Copy prompt and paste into Copilot Chat / Claude / Gemini. To check GPU offload: ollama ps"
    } elseif ($hit.Processor -eq 'GPU') {
        $ui.lblAiBannerIcon.Foreground = $Window.TryFindResource('ThemeSuccess')
        $ui.lblAiBannerText.Text = "Local model detected with GPU acceleration. Run review should complete in 30–90s."
    } else {
        # Ollama is running but no model loaded yet — neutral guidance
        $ui.lblAiBannerIcon.Foreground = $Window.TryFindResource('ThemeAccentLight')
        $ui.lblAiBannerText.Text = "Ollama detected. Run review sends the prompt to the local model. If inference is slow (no GPU), use Copy prompt instead."
    }
}
$ui.btnAiCopyPrompt.Add_Click({
    if (-not $Script:State.ReportPath) {
        $ui.lblAiMeta.Text     = 'No report loaded yet — run an analysis first.'
        Set-AiResponseText 'Run an analysis first; the prompt is built from the report.' -Plain
        Set-Status -Text 'Copy prompt: no report' -Tone 'Error' -Detail ''
        return
    }
    $prompt = Get-CurrentPromptText
    if (-not $prompt) {
        # Resolve the expected path so the user can see what's missing.
        $dir  = [IO.Path]::GetDirectoryName($Script:State.ReportPath)
        $base = [IO.Path]::GetFileNameWithoutExtension($Script:State.ReportPath) -replace '\.report$',''
        $p    = Join-Path $dir ($base + '.llm-prompt.md')
        $ui.lblAiMeta.Text     = "llm-prompt.md not found"
        Set-AiResponseText "Expected: $p`r`n`r`nRe-run analysis (Analyze dump page, with `"Force re-export`" if cached) to regenerate the .llm-prompt.md artifact." -Plain
        Set-Status -Text 'Copy prompt: prompt file missing' -Tone 'Error' -Detail $p
        Write-DebugLog "Copy prompt: $p not found" 'WARN'
        return
    }
    try {
        # Clipboard.SetText can briefly fail with CLIPBRD_E_CANT_OPEN on contested
        # boards (RDP, clipboard managers). One retry + DataObject fallback covers
        # the realistic failure modes.
        try { [System.Windows.Clipboard]::SetText($prompt) } catch {
            $do = New-Object System.Windows.DataObject
            $do.SetText($prompt, [System.Windows.TextDataFormat]::UnicodeText)
            [System.Windows.Clipboard]::SetDataObject($do, $true)
        }
        $kb = [Math]::Round($prompt.Length / 1KB, 1)
        $ui.lblAiMeta.Text = "Prompt copied to clipboard ($kb KB) — paste into Copilot Chat / Claude / Gemini / GPT-4."
        Set-Status -Text 'Prompt copied to clipboard' -Tone 'Success' -Detail "$kb KB"
        Write-DebugLog "Copy prompt: $kb KB copied to clipboard" 'SUCCESS'
    } catch {
        $msg = $_.Exception.Message
        $ui.lblAiMeta.Text = "Clipboard copy failed: $msg"
        Set-Status -Text 'Clipboard copy failed' -Tone 'Error' -Detail $msg
        Write-DebugLog "Clipboard copy failed: $msg" 'ERROR'
    }
})
$ui.btnAiRun.Add_Click({
    if (-not $Script:State.ReportPath) { Write-DebugLog 'No report loaded yet' 'WARN'; return }
    $prompt = Get-CurrentPromptText
    if (-not $prompt) { Write-DebugLog 'llm-prompt.md not found next to report -- re-run analysis to regenerate' 'WARN'; return }
    $baseUrl = $Script:State.Prefs.LlmBaseUrl
    $model   = $Script:State.Prefs.LlmModel
    $apiKey  = $Script:State.Prefs.LlmApiKey
    if (-not $baseUrl -or -not $model) { Write-DebugLog 'LLM Base URL / Model not configured in Settings' 'WARN'; return }

    Update-AiEndpointLabel
    $isLocal = $baseUrl -match 'localhost|127\.0\.0\.1'
    $waitMsg = if ($isLocal) { "Sending to $model locally..." } else { "Sending to $model..." }
    $ui.lblAiMeta.Text = "Sending $([Math]::Round($prompt.Length/1KB,1)) KB to $model..."
    Set-AiResponseText $waitMsg -Plain
    $ui.btnAiRun.IsEnabled = $false
    Set-Busy -Busy $true -Status 'AI review running...'
    $Script:State.AiStarted = Get-Date
    $Script:State.AiModel   = $model

    # Re-create the function inside the runspace via string (background runspaces
    # do not see $Script:/$Global: from the GUI thread -- see /memories notes).
    $fnSrc = ${function:Invoke-LlmChat}.ToString()
    Start-BackgroundJob `
        -Kind 'AiReview' `
        -Body {
            param($FnSrc, $BaseUrl, $Model, $ApiKey, $Prompt, $Timeout)
            $ErrorActionPreference = 'Stop'
            try {
                Invoke-Expression "function Invoke-LlmChat { $FnSrc }"
                $r = Invoke-LlmChat -BaseUrl $BaseUrl -Model $Model -ApiKey $ApiKey -Prompt $Prompt -TimeoutSec $Timeout
                [pscustomobject]@{ Success = $true; Text = $r }
            } catch {
                [pscustomobject]@{ Success = $false; Error = $_.Exception.Message }
            }
        } `
        -Arguments @($fnSrc, $baseUrl, $model, $apiKey, $prompt, [int]$Script:State.Prefs.LlmTimeout) `
        -OnDone {
            param($r)
            $ui.btnAiRun.IsEnabled = $true
            $elapsed = [int]((Get-Date) - $Script:State.AiStarted).TotalSeconds
            $mdl = $Script:State.AiModel
            if ($r -and $r.Success -and $r.Text) {
                $Script:State.LastAiResponse = $r.Text
                # Format the elapsed time nicely
                $elapsedText = if ($elapsed -ge 60) {
                    "{0}m {1}s" -f [int][Math]::Floor($elapsed / 60), ($elapsed % 60)
                } else { "${elapsed}s" }
                # Append a generation footer to the markdown before rendering
                $footer = "`n`n---`n`n*Generated by **$mdl** in $elapsedText at $(Get-Date -Format 'HH:mm')*"
                Set-AiResponseText ($r.Text + $footer)
                $ui.lblAiMeta.Text     = "$mdl -> $([Math]::Round($r.Text.Length/1KB,1)) KB in $elapsedText"
                Set-Status -Text 'AI review ready' -Tone 'Success' -Detail $elapsedText
            } else {
                $err = if ($r) { $r.Error } else { 'no response' }
                Set-AiResponseText "AI review failed.`r`n`r`nError: $err`r`n`r`nChecklist:`r`n  - Is the endpoint reachable? Try the Auto-detect button in Settings.`r`n  - Is the model name correct? For Ollama, list with: ollama list`r`n  - For hosted endpoints, did you set the API key in Settings?`r`n`r`nFallback: use 'Copy prompt' and paste into any LLM chat UI." -Plain
                $ui.lblAiMeta.Text     = "failed after ${elapsed}s"
                Set-Status -Text 'AI review failed' -Tone 'Error' -Detail $err
                Write-DebugLog "AI review failed: $err" 'ERROR'
            }
        }
})

# Log console
$ui.btnClearLog.Add_Click({ $ui.paraLog.Inlines.Clear() })
$ui.btnHideLog.Add_Click({
    $ui.LogConsole.Visibility = 'Collapsed'
    $Script:State.Prefs.ShowLog = $false
    Save-Prefs
})

# Settings tab
$ui.chkLightMode.Add_Click({
    $Script:State.Prefs.IsLightMode = [bool]$ui.chkLightMode.IsChecked
    Apply-CurrentTheme
    Save-Prefs
})
$ui.chkVerbose.Add_Click({
    $Script:State.Prefs.Verbose = [bool]$ui.chkVerbose.IsChecked
    Save-Prefs
})
$ui.chkAnimations.Add_Click({
    $Script:State.Prefs.Animations = [bool]$ui.chkAnimations.IsChecked
    Save-Prefs
})
$ui.chkDebugOverlay.Add_Click({
    $Script:State.Prefs.DebugOverlay = [bool]$ui.chkDebugOverlay.IsChecked
    Save-Prefs
})

# LLM settings
$ui.txtLlmBaseUrl.Add_LostFocus({
    $v = [string]$ui.txtLlmBaseUrl.Text
    if ($v) { $Script:State.Prefs.LlmBaseUrl = $v.TrimEnd('/'); Save-Prefs; Update-AiEndpointLabel }
})

# Model ComboBox: sync on selection change and on text edit (editable ComboBox)
function Refresh-OllamaModels {
    # Populate the model dropdown from Ollama /api/tags
    $base = $Script:State.Prefs.LlmBaseUrl
    if ($base -notmatch 'localhost:11434|127\.0\.0\.1:11434') { return }
    try {
        $baseNoV1 = $base -replace '/v1$',''
        $tags = Invoke-RestMethod -Uri "$baseNoV1/api/tags" -TimeoutSec 3 -ErrorAction Stop
        if ($tags.models) {
            $names = @($tags.models | ForEach-Object { [string]$_.name } | Sort-Object)
            $ui.cboLlmModel.ItemsSource = $names
            Write-DebugLog "Loaded $($names.Count) models from Ollama" 'DEBUG'
        }
    } catch {
        Write-DebugLog "Failed to list Ollama models: $($_.Exception.Message)" 'WARN'
    }
}
$ui.cboLlmModel.Add_SelectionChanged({
    $v = [string]$ui.cboLlmModel.SelectedItem
    if ($v) { $Script:State.Prefs.LlmModel = $v; Save-Prefs; Update-AiEndpointLabel }
})
$ui.cboLlmModel.Add_LostFocus({
    $v = [string]$ui.cboLlmModel.Text
    if ($v) { $Script:State.Prefs.LlmModel = $v; Save-Prefs; Update-AiEndpointLabel }
})
$ui.btnRefreshModels.Add_Click({ Refresh-OllamaModels })
$ui.txtLlmApiKey.Add_LostFocus({
    # PasswordBox -- read .Password, not .Text
    $Script:State.Prefs.LlmApiKey = [string]$ui.txtLlmApiKey.Password
    Save-Prefs
})
if ($ui.txtSymbolPath) {
    $ui.txtSymbolPath.Add_LostFocus({
        $Script:State.Prefs.SymbolPath = [string]$ui.txtSymbolPath.Text
        Save-Prefs
    })
}
if ($ui.txtLlmTimeout) {
    $ui.txtLlmTimeout.Add_LostFocus({
        $val = [string]$ui.txtLlmTimeout.Text
        $parsed = 0
        if ([int]::TryParse($val, [ref]$parsed) -and $parsed -ge 30) {
            $Script:State.Prefs.LlmTimeout = $parsed
        } else {
            $Script:State.Prefs.LlmTimeout = 600
            $ui.txtLlmTimeout.Text = '600'
        }
        Save-Prefs
    })
}
if ($ui.btnCopyQuickStart) {
    $ui.btnCopyQuickStart.Add_Click({
        [System.Windows.Clipboard]::SetText($ui.txtQuickStart.Text)
        Write-DebugLog 'Quick-start commands copied to clipboard' 'INFO'
    })
}
$ui.btnLlmDetect.Add_Click({
    $ui.lblLlmDetect.Text = 'probing localhost...'
    $hit = Detect-LlmRuntime
    if ($hit) {
        $Script:State.Prefs.LlmBaseUrl = $hit.Url
        $ui.txtLlmBaseUrl.Text = $hit.Url
        if ($hit.FirstModel) {
            $Script:State.Prefs.LlmModel = $hit.FirstModel
            $ui.cboLlmModel.Text = $hit.FirstModel
        }
        $label = "found $($hit.Name) at $($hit.Url)"
        if ($hit.FirstModel) { $label += " (model: $($hit.FirstModel))" }
        if ($hit.Processor) {
            $label += " — running on $($hit.Processor)"
            if ($hit.CpuOnly) {
                $label += ". CPU-only inference is slow (5-10 min for a review). Recommendation: use Copy prompt and paste into Copilot Chat / Gemini / Claude for instant results."
            }
        } else {
            $label += ' (no model loaded yet — run ollama pull gemma4:e2b)'
        }
        $ui.lblLlmDetect.Text = $label
        Save-Prefs
        Update-AiEndpointLabel
        Update-AiBanner
        Refresh-OllamaModels
    } else {
        $ui.lblLlmDetect.Text = 'no local OpenAI-compatible endpoint found on 11434 / 1234 / 8080 / 5273'
    }
})

# Title-bar Help + Theme toggle
$ui.btnHelp.Add_Click({
    [System.Windows.MessageBox]::Show(
        "DumpPilot v$Script:AppVersion`r`n`r`nCrash dump triage tool.`r`nRuns cdb.exe against .dmp files, extracts structured facts,`r`nand builds an escalation-grade LLM prompt.`r`n`r`nKeyboard:`r`n  Ctrl+``  - Toggle log console`r`n  Ctrl+T  - Toggle theme`r`n  F1     - This dialog`r`n`r`nDebug log: $Script:LogPath",
        'DumpPilot', 'OK', 'Information') | Out-Null
})
$ui.btnThemeToggle.Add_Click({
    $Script:State.Prefs.IsLightMode = -not $Script:State.Prefs.IsLightMode
    Apply-CurrentTheme
    if ($ui.chkLightMode) { $ui.chkLightMode.IsChecked = $Script:State.Prefs.IsLightMode }
    Save-Prefs
})

# Findings -> patterns shortcut
$ui.dgFindings.Add_MouseDoubleClick({
    $sel = $ui.dgFindings.SelectedItem
    if ($sel -and $sel.Id) { $ui.tabsResult.SelectedIndex = 1 }
})

# ---------------------------------------------------------------- Initial state
Apply-CurrentTheme
$ui.chkLightMode.IsChecked      = $Script:State.Prefs.IsLightMode
$ui.chkVerbose.IsChecked        = $Script:State.Prefs.Verbose
$ui.chkAnimations.IsChecked     = $Script:State.Prefs.Animations
$ui.chkDebugOverlay.IsChecked   = $Script:State.Prefs.DebugOverlay
$ui.txtLlmBaseUrl.Text          = $Script:State.Prefs.LlmBaseUrl
$ui.cboLlmModel.Text            = $Script:State.Prefs.LlmModel
$ui.txtLlmApiKey.Password       = $Script:State.Prefs.LlmApiKey
if ($ui.txtLlmTimeout) { $ui.txtLlmTimeout.Text = [string]$Script:State.Prefs.LlmTimeout }
if ($ui.txtSymbolPath) { $ui.txtSymbolPath.Text = $Script:State.Prefs.SymbolPath }
Update-AiEndpointLabel
Refresh-OllamaModels
$ui.LogConsole.Visibility       = if ($Script:State.Prefs.ShowLog) { 'Visible' } else { 'Collapsed' }
$ui.lblVersion.Text             = "v$Script:AppVersion"

$Script:State.PatternList = Load-Patterns
$ui.dgPatterns.ItemsSource = @($Script:State.PatternList)
$src = if ($Script:State.PatternSourceFiles -and $Script:State.PatternSourceFiles.Count -gt 0) {
    ($Script:State.PatternSourceFiles -join ', ')
} else {
    'no pattern DB files found'
}
$ui.lblPatternSummary.Text = "$($Script:State.PatternList.Count) patterns loaded ($src)"

$adk = @(
    'C:\Program Files (x86)\Windows Kits\10\Debuggers\x64\cdb.exe',
    'C:\Program Files\Windows Kits\10\Debuggers\x64\cdb.exe',
    'C:\Program Files (x86)\Windows Kits\10\Debuggers\x86\cdb.exe'
) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1

$symPathDisplay = if ($ui.txtSymbolPath -and $ui.txtSymbolPath.Text) {
    $ui.txtSymbolPath.Text
} elseif ($env:_NT_SYMBOL_PATH) {
    $env:_NT_SYMBOL_PATH
} else {
    '(not set — will use Microsoft public symbol server)'
}

$ui.lblAdkInfo.Text = if ($adk) {
    "cdb.exe: $adk`nSymbol path: $symPathDisplay"
} else {
    'cdb.exe NOT FOUND. Install Debugging Tools for Windows from the Windows SDK.'
}

Refresh-History
Set-Status -Text 'Ready' -Tone 'Ready' -Detail 'Pick a dump file on the Analyze page.'
Write-DebugLog "DumpPilot v$Script:AppVersion ready. Debug log at $Script:LogPath" 'INFO'

# ---------------------------------------------------------------- Show
# Global keyboard shortcuts
$Window.Add_KeyDown({
    param($sender, $e)
    $ctrl = ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control) -ne 0
    if ($e.Key -eq 'F1') {
        $ui.btnHelp.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        $e.Handled = $true
    } elseif ($ctrl -and $e.Key -eq 'OemTilde') {     # Ctrl+` -- toggle log
        $ui.NavToggleConsole.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        $e.Handled = $true
    } elseif ($ctrl -and $e.Key -eq 'T') {            # Ctrl+T -- toggle theme
        $ui.btnThemeToggle.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        $e.Handled = $true
    }
})

[void]$Window.ShowDialog()
