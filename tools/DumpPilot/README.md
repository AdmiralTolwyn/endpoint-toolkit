# DumpPilot

Automated Windows crash-dump triage tool.  
Runs **cdb.exe** against `.dmp` files, extracts 40+ structured fact categories,
and builds an escalation-grade LLM prompt for AI-assisted root cause analysis.

## What it does

1. **Stage 1 — Export facts** (`Export-DumpFacts.ps1`): runs 44+ cdb commands
   in two passes (static bundle + dynamic `lmv` for faulting/stack modules).
   Outputs `raw.txt` + `meta.json`.
2. **Stage 2 — Parse** (`ConvertTo-DumpSummary.ps1`): regex-extracts 40+ field
   groups into `dump-summary.json`: exception, stack (up to 100 frames),
   per-thread stacks, registers, module details (version/vendor/symbols),
   VAS/heap/handle stats, timing, PEB, disassembly at crash point, NTSTATUS
   code decode, and more.
3. **Stage 3 — Analyze** (`helpers/Invoke-DumpAnalysis.ps1`): builds
   `report.json` + `report.md` + `llm-prompt.md` with an escalation-grade
   system prompt (v2.4) containing 10+ reasoning rules.
4. **Stage 4 — HTML** (`helpers/Export-DumpHtmlReport.ps1`): dark/light
   themed single-file report with stat cards, color-coded stack, deduplicated
   thread table, collapsible sections.

## WPF GUI

`DumpPilot.ps1` launches a frameless WPF window with:
- **Stack tab** — color-coded by module type (fault=red, GPU=orange, app=blue,
  system=gray), duplicate stack segments dimmed with separator
- **Threads tab** — deduplicated + role-classified (UI/IO/Wait/Monitor/Sleep)
- **Facts tab** — all dump facts in Category/Field/Value grid
- **Manifest tab** — raw JSON
- **AI review** — send the LLM prompt to Ollama/LM Studio/OpenAI-compatible endpoint
- **Settings** — symbol path override, LLM endpoint/model/timeout configuration

## LLM triage quality

The v2.4 prompt produces near-escalation-grade triage when used with capable models.
Tested on the same Java/Intel OpenGL crash dump:

| Model | Quality | Notes |
|-------|---------|-------|
| **Gemini 2.5 Pro** | Excellent | Used disassembly to prove `r14=0` NULL deref, identified `DumpRegistryKeyDefinitions` symbol name as registry operation, spotted concurrent UI thread, correctly rated Medium confidence |
| **Claude Opus 4** | Excellent | Similar quality to Gemini, strong causality-vs-locality distinction, good hypothesis validation steps |
| **ChatGPT (GPT-4o)** | Weak | Ignored the `Disassembly[]` field entirely (claimed "no disassembly captured"), hallucinated thread count (said 108, actual 54), produced free-form essay instead of required 6-section format, never classified module vendor |

**Recommendation:** Use Gemini or Claude for the LLM triage step. ChatGPT tends to
ignore structured data fields and doesn't follow strict output format instructions
reliably for this use case.

## Quick start

```powershell
# Full pipeline (CLI)
.\Invoke-DumpPilotPipeline.ps1 -DumpPath 'C:\Dumps\app.exe.1234.dmp' -Verbose

# GUI
.\DumpPilot.ps1
```

## Prereqs

- **cdb.exe** — part of the Windows SDK "Debugging Tools for Windows".
  Auto-located under `C:\Program Files (x86)\Windows Kits\10\Debuggers\x64\`.
- **Symbols** — if `_NT_SYMBOL_PATH` is not set, falls back to the public
  Microsoft symbol server. Override in Settings or pass `-SymbolPath`.

### Setting up the symbol path

DumpPilot needs a symbol path to resolve stack frames. You can configure it
in the GUI (Settings > Symbol path override) or set the `_NT_SYMBOL_PATH`
environment variable system-wide.

**Option 1 — Windows GUI**

1. Press `Win + R`, type `sysdm.cpl`, press Enter.
2. Go to the **Advanced** tab → **Environment Variables**.
3. Under **System variables**, click **New**.
4. Set **Variable name** to: `_NT_SYMBOL_PATH`
5. Set **Variable value** to:
   ```
   srv*C:\Symbols*https://msdl.microsoft.com/download/symbols
   ```
6. Click **OK** on all windows to save.

**Option 2 — Command Prompt (current session)**

```cmd
set _NT_SYMBOL_PATH=srv*C:\Symbols*https://msdl.microsoft.com/download/symbols
```

**Option 3 — PowerShell (persistent for current user)**

```powershell
[Environment]::SetEnvironmentVariable('_NT_SYMBOL_PATH',
    'srv*C:\Symbols*https://msdl.microsoft.com/download/symbols',
    'User')
```

> **Note:** The `C:\Symbols` folder is a local cache — symbols are downloaded
> once from Microsoft's server and reused on subsequent runs. First runs may
> take 1–2 minutes longer while the cache populates.

## Key files

| File | Purpose |
|------|---------|
| [DumpPilot.ps1](DumpPilot.ps1) | WPF GUI host |
| [DumpPilot_UI.xaml](DumpPilot_UI.xaml) | XAML layout |
| [Invoke-DumpPilotPipeline.ps1](Invoke-DumpPilotPipeline.ps1) | CLI pipeline orchestrator |
| [Export-DumpFacts.ps1](Export-DumpFacts.ps1) | Stage 1 — cdb wrapper (two-pass) |
| [ConvertTo-DumpSummary.ps1](ConvertTo-DumpSummary.ps1) | Stage 2 — parser (40+ fields) |
| [helpers/Invoke-DumpAnalysis.ps1](helpers/Invoke-DumpAnalysis.ps1) | Stage 3 — report + LLM prompt |
| [helpers/Export-DumpHtmlReport.ps1](helpers/Export-DumpHtmlReport.ps1) | Stage 4 — rich HTML report |
| [ntstatus_codes.json](ntstatus_codes.json) | 505 NTSTATUS codes from MS-ERREF |
| [pattern_db.json](pattern_db.json) | Legacy pattern DB (not used in facts-only mode) |
