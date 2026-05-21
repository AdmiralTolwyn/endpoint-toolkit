# DumpPilot — Design

Deterministic Windows crash-dump triage with an optional LLM rewrite stage.
Sibling of [TraceAnalyzer](../TraceAnalyzer); same pipeline shape, different
extractor and pattern DB.

---

## 1. Goals & non-goals

### Goals

- Take one `.dmp` / `.mdmp` / `MEMORY.DMP` and produce an **actionable** report:
  - Faulting module + bucket
  - Faulting stack (top N frames, symbolized when symbols available)
  - Loaded-module versions for anything that appears in the faulting stack
  - Process command line (catches JVM flags, browser switches, etc.)
  - Pattern-DB hits (known module/bucket → canned remediation)
- Reuse the TraceAnalyzer WPF host, evidence model, scoring, history,
  achievements **to the pixel**. Only the extractor and pattern DB differ.
- LLM step is opt-in and consumes the deterministic report — never the raw
  `.dmp`.

### Non-goals

- Not a replacement for WinDbg. DumpPilot triages; WinDbg drills.
- No live debugging. Post-mortem only.
- No symbol-server hosting. Respects `_NT_SYMBOL_PATH`.

---

## 2. Scope (decided)

| Dump kind          | Detected via         | Extractor bundle |
|--------------------|----------------------|------------------|
| User-mode minidump | `||` → "Mini Dump"   | `user-mini.cdb`  |
| User-mode full    | `||` → "User Dump"   | `user-full.cdb`  |
| Kernel             | `||` → "Kernel..."   | `kernel.cdb`     |

All three open in `cdb.exe`; only the command set changes.

---

## 3. Symbols policy

- Prereq check at extractor entry:
  - `_NT_SYMBOL_PATH` set? → proceed.
  - Not set? → **warn**, offer to set
    `srv*C:\Symbols*https://msdl.microsoft.com/download/symbols` for the
    current process, and **continue**. Module+offset buckets still resolve.
- Block only when a pattern-DB rule requires `StackRegex` matching against
  resolved symbol names AND resolution failed. That's a finding, not a hard
  stop.

---

## 4. Pipeline

```
DMP ──► Export-DumpFacts ──► cdb.exe ──► <dump>.dump-facts/
                                              raw.txt           (full cdb output)
                                              analyze.txt       (!analyze -v only)
                                              stack.txt         (kb 30 faulting thread)
                                              modules.txt       (lm + lmvm hits)
                                              peb.txt           (command line)
                                              meta.json         (provenance)
                                              │
                                              ▼
                                  ConvertTo-DumpSummary
                                              │
                                              ▼
                                  dump-summary.json
                                              │
                                              ▼
                                  Invoke-DumpAnalysis
                                              │
              ┌──────────────────┬────────────┴───────────┬──────────────────┐
              ▼                  ▼                        ▼                  ▼
         Enrichers          Classifiers              EvidenceManifest    Pattern DB hits
              │                  │                        │                  │
              └──────────────────┴───────────┬────────────┴──────────────────┘
                                             ▼
                                     Score-Findings
                                             │
                                             ▼
                                  dump-report.{json,md}
                                             │
                                  (opt-in) LLM rewrite
                                             ▼
                                  dump-report.llm.md
```

Stages 3+ reuse the TraceAnalyzer implementations via `..\TraceAnalyzer\helpers\`
(no copy; `dot-source` from the sibling folder for now, factor up to
`scripts\_shared\` as a follow-up).

---

## 5. Pattern DB shape

```json
{
  "id": "java-intel-opengl-icd",
  "match": {
    "ProcessImage": "^java(w)?\\.exe$",
    "FaultingModule": "^igxelp(g)?icd64\\.dll$",
    "ExceptionCode": "c0000005|c000041d",
    "StackRegex": "wglSetPixelFormat|opengl32!|sun\\.java2d|prism"
  },
  "severity": "High",
  "title": "Java OpenGL pipeline crashes in Intel iGPU ICD",
  "remediation_md": "..."
}
```

Match order: most-specific first. Multiple hits allowed; scoring picks the
highest-severity match, others become secondary evidence.

---

## 6. Reuse map (TraceAnalyzer → DumpPilot)

| TraceAnalyzer asset                                         | DumpPilot status |
|-------------------------------------------------------------|---------------------|
| `TraceAnalyzer_UI.xaml`                                     | reuse, swap title + file filter |
| `helpers\Invoke-TraceAnalysis.ps1` (scoring, classifiers)   | dot-source, register new families |
| `helpers\Export-HtmlReport.ps1`                             | reuse as-is |
| `achievements.json`, `streak.json`, `analysis_history.json` | reuse format, separate files |
| `pattern_db.json`                                           | new DB, same schema + `StackRegex` field |
| `profiles\*.wpaProfile`                                     | n/a — replaced by cdb command scripts |

---

## 7. Files in this folder

- [Export-DumpFacts.ps1](Export-DumpFacts.ps1) — cdb wrapper, stage 1.
- [ConvertTo-DumpSummary.ps1](ConvertTo-DumpSummary.ps1) — parses cdb output → JSON, stage 2.
- [pattern_db.json](pattern_db.json) — seed rules (Java/Intel iGPU is rule #1).
- [README.md](README.md) — quick-start.
- (later) `DumpPilot.ps1` — WPF host shim, mirrors `TraceAnalyzer.ps1`.

---

## 8. Open follow-ups

- Factor TraceAnalyzer `helpers\` up to `scripts\_shared\` so both tools share
  scoring/evidence code without dot-sourcing across siblings.
- Kernel-dump command bundle (`!process 0 0`, `!stacks`, `!irql`).
- LLM rewrite stage: lift the prompt builder out of TraceAnalyzer's pilot tab.
