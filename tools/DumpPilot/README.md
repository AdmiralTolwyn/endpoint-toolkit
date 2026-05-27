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
   `report.json` + `report.md` + `llm-prompt.md` (trimmed, ~40 KB) +
   `llm-prompt.full.md` (un-trimmed, often >500 KB) with an escalation-grade
   system prompt (v2.5) containing 10+ reasoning rules, pattern-DB matching,
   ProcMon sibling-folder consolidation, and OS-chatter denylist.
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

The v2.5 prompt produces near-escalation-grade triage when used with capable
models. Tested on a Java/Intel OpenGL crash dump:

| Model | Quality | Notes |
|-------|---------|-------|
| **Claude Opus 4.7** | Excellent | Uses `Disassembly[]` to prove the NULL deref (e.g. `cmp dword ptr [r8+31E18h],0` paired with `r8=0`), reads `DumpRegistryKeyDefinitions` symbol name as a registry-read operation, spots concurrent UI threads via the collapsed `ThreadStacks` groups, correctly rates Medium confidence when symbol coverage is incomplete |
| **Gemini 3.1 Pro** | Excellent | Equivalent quality to Claude; strong causality-vs-locality distinction, good hypothesis-validation steps, reliably follows the 6-section output format |
| **Copilot** (M365 / GitHub) | Good–Excellent | Follows the structured prompt reliably, cites JSON fields, respects the 6-section format. Quality scales with the picked backing model — strongest when "Claude Opus 4.7" or "Gemini 3.1 Pro" is selected from the M365 / GitHub Copilot model picker |
| **ChatGPT** (GPT-5 / GPT-4o) | Weak | Ignored the `Disassembly[]` field entirely (claimed "no disassembly captured"), hallucinated thread count (said 108, actual 54), produced a free-form essay instead of the required 6-section format, never classified module vendor. **Not recommended** for this workflow |

**Recommendation:** Use Claude, Gemini, or Copilot for the LLM triage step.
ChatGPT (OpenAI's consumer/Plus product) tends to ignore structured data
fields and doesn't reliably follow the strict 6-section output format
required by the prompt — this is specifically about the ChatGPT chat UI, not
Copilot, which works fine even though it can be backed by OpenAI models.

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

## ProcMon integration

DumpPilot can optionally correlate a
[Process Monitor](https://learn.microsoft.com/en-us/sysinternals/downloads/procmon)
trace with the crash dump to surface file, registry, and network failures that
occurred around the time of the crash. This is Stage 2b of the pipeline —
handled by
[helpers/Import-ProcMonCorrelation.ps1](helpers/Import-ProcMonCorrelation.ps1).

### How it works

1. **Input** — pass a `.pml` (native ProcMon binary) or `.csv` export via
   `-ProcMonPath` on the CLI, or use the **Browse** button next to the ProcMon
   field on the Analyze Dump tab in the GUI.
2. **PML → CSV** — if a `.pml` is provided and `Procmon.exe` is available,
   DumpPilot auto-converts it to CSV. If `Procmon.exe` is not on PATH, the
   Settings tab has a one-click **Install** button
   (`winget install Sysinternals.ProcessMonitor`) and a path override field.
3. **C# fast parser** — a compiled-at-runtime C# CSV reader streams the trace
   (~50× faster than `Import-Csv`), extracting Time of Day, Process Name, PID,
   Operation, Path, and Result for every event.
4. **Time-windowed filtering** — only failures within 60 s (configurable via
   `-WindowSeconds`) before the last target-process event are kept. SUCCESS,
   BUFFER OVERFLOW, REPARSE, and other benign results are excluded.
5. **GPU-crash expansion** — if the faulting module matches a GPU/graphics
   driver regex (`igxel|igc|nvoglv|nvd3d|amdxx|atiux|d3d|opengl`), vendor
   service failures (RaCEF, igfxEM, nvcontainer, AMDRSServ, …) are included
   regardless of PID.
6. **Output** — `procmon-correlation.json` is written alongside
   `dump-summary.json` and merged into it as a `ProcMonCorrelation` property
   with `FaultRelated` and `OtherFailures` arrays.

### LLM prompt trimming

Before the correlation data reaches the LLM prompt, Stage 3 applies three
reduction passes:

| Pass | What it does |
|------|--------------|
| **OS-chatter denylist** | Strips noise paths: Power settings, WER temp, AppRepository, Defender cache, IntuneManagementExtension, WinSxS hotpatches, etc. |
| **Sibling-file consolidation** | Groups failures by Operation \| Parent path \| Result. When ≥ 3 sibling files in the same folder fail the same way, they collapse into a single `\Parent\* (N files)` row with examples. |
| **Top-25 cap** | After consolidation, keeps only the 25 highest-count failure patterns. Original counts are preserved in metadata (`FaultRelatedTruncatedFrom`, `FaultRelatedDroppedAsOsChatter`). |

The **trimmed** prompt (`llm-prompt.md`, ~40 KB) contains the top-25
ProcMon failures; the **full** prompt (`llm-prompt.full.md`, often >500 KB)
retains every failure. Both variants instruct the LLM:

> *The FaultRelated array lists failures whose paths match the faulting module
> vendor or graphics/driver registry keys. Use these to confirm or refute
> hypotheses about missing configuration, denied access, or absent registry
> keys.*

### CLI usage

```powershell
.\Invoke-DumpPilotPipeline.ps1 -DumpPath 'C:\Dumps\app.dmp' `
                                -ProcMonPath 'C:\Dumps\trace.pml' -Verbose
```

### GUI usage

1. On the **Analyze Dump** tab, click **Browse** next to the ProcMon field and
   select a `.pml` or `.csv` file.
2. Click **Analyze**. The pipeline runs Stages 1 → 2a → 2b → 3 → 4
   automatically.
3. ProcMon failures appear in the HTML report (fault-related table +
   other-failures subsection) and in both LLM prompt variants.

### Installing Procmon

`Procmon.exe` is only needed when providing `.pml` files (native binary
traces). If you export to CSV from ProcMon yourself, the executable is not
required.

- **Settings tab** — click **Install** to run
  `winget install Sysinternals.ProcessMonitor`.
- **Manual** — download from the
  [Sysinternals site](https://learn.microsoft.com/en-us/sysinternals/downloads/procmon)
  and place `Procmon.exe` on PATH or set its location in Settings.
- **Auto-discovery** — DumpPilot checks `C:\Tools\Procmon.exe`,
  `C:\SysinternalsSuite\Procmon.exe`, the winget install path, and finally
  `Get-Command Procmon.exe`.

## Key files

| File | Purpose |
|------|---------|
| [DumpPilot.ps1](DumpPilot.ps1) | WPF GUI host |
| [DumpPilot_UI.xaml](DumpPilot_UI.xaml) | XAML layout |
| [Invoke-DumpPilotPipeline.ps1](Invoke-DumpPilotPipeline.ps1) | CLI pipeline orchestrator |
| [Export-DumpFacts.ps1](Export-DumpFacts.ps1) | Stage 1 — cdb wrapper (two-pass) |
| [ConvertTo-DumpSummary.ps1](ConvertTo-DumpSummary.ps1) | Stage 2a — parser (40+ fields) |
| [helpers/Import-ProcMonCorrelation.ps1](helpers/Import-ProcMonCorrelation.ps1) | Stage 2b — ProcMon correlation |
| [helpers/Invoke-DumpAnalysis.ps1](helpers/Invoke-DumpAnalysis.ps1) | Stage 3 — report + LLM prompt |
| [helpers/Export-DumpHtmlReport.ps1](helpers/Export-DumpHtmlReport.ps1) | Stage 4 — rich HTML report |
| [ntstatus_codes.json](ntstatus_codes.json) | 505 NTSTATUS codes from MS-ERREF |
| [pattern_db.json](pattern_db.json) | Legacy pattern DB (not used in facts-only mode) |
