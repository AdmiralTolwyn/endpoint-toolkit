# BaselinePilot Audit — July 2026

Full audit per the AvdAssessor program conventions (../AvdAssessor/AUDIT.md). All 308 checks'
collectionKeys resolved programmatically against the collector's modeled output shape; baseline
values web-verified against July-2026 Microsoft guidance; engine + collector + tests + hygiene
audited. `coll:` = Invoke-BaselineCollection.ps1, `app:` = BaselinePilot.ps1. Verified: 308 checks,
318 collectionKeys, 24 collector sections, 215 exact registry reads + 4 wildcards.

## A. Critical correctness — silently wrong results at scale

- **A-1 · All 44 audit-policy checks (MON-*) false-Fail on every admin collection.** Collector's
  admin path keys `auditPolicy` by subcategory **GUID** (coll:715); every MON check's collectionKey
  uses the **display name** (`auditPolicy.Credential Validation`). Since the collector requires
  admin, the GUID path always wins → lookup null → "not configured = Fail" (app:3209). Whole
  Monitoring category permanently Fail. Also values are `{name, setting}` objects — stringified
  hashtable can never equal `"Success and Failure"`. Fix: collector emits BOTH GUID and display-name
  keys with plain setting strings (or app resolves via a GUID map).
- **A-2 · 24 checks with dangling collectionKeys** → 23 permanent false-Fail + 1 permanent
  Not-Assessed (DEF-042):
  - 12 registry paths never read: SEC-015 (fAllowUnsolicited), SEC-016 (MSV1_0\RestrictSendingNTLMTraffic),
    SEC-019 (ProcessCreationIncludeCmdLine_Enabled), SEC-022 (fAllowToGetHelp), SEC-036
    (NoConnectedUser), SEC-038 (WinRS\AllowRemoteShellAccess), SEC-056 (SmartScreen
    ConfigureAppInstallControl), SEC-062+NET-015 (TCPIP6\DisabledComponents), AUTH-015 (DisableCAD),
    AUTH-017 (DisableCredManagerAutocomplete), DATA-007 (fDisableClip), **UAC-009
    (LocalAccountTokenFilterPolicy — the PtH mitigation, Critical-adjacent)**.
  - 2 Defender "Policy Manager" paths (DEF-022 network protection, DEF-023 CFA) — should key to
    `defender.EnableNetworkProtection` / `defender.EnableControlledFolderAccess` (already emitted).
  - 6 Defender properties never emitted: DEF-024 (ScheduleDay/ScheduleTime), DEF-029
    (DisableIOAVProtection — collector emits inverted `IoavProtectionEnabled`), DEF-035
    (SignatureUpdateInterval), DEF-036 (DisableScriptScanning), DEF-039 (mapped-drive scan),
    DEF-042 (ExclusionPath — never collected).
  - 3 firewall keys misnamed: NET-005/006/007 use `LogDropped`; collector emits `LogBlocked`
    (coll:980).
- **A-3 · Section failure → false-Fail wave.** `Invoke-CollectionArea` catch returns $null
  (coll:275) recording only `_metadata.errors`, which the importer NEVER reads — a broken
  Get-MpPreference produces ~40 Defender "Fails". Fix: importer maps checks whose section is null
  AND listed in `_metadata.errors` to Not Assessed with the error reason.
- **A-4 · "Not configured = Fail" is wrong for the default-secure family on 24H2+.** SMB signing
  (NET-001/SEC-045/046), TLS 1.0/1.1 disabled-by-default (NET-024/025 — absent SCHANNEL keys on
  clean 24H2 → false Fail), Credential Guard/VBS policy (SEC-061), Tamper Protection, PSv2 removed
  (SEC-065). The tool penalizes the MOST secure machines. Fix: OS-build-aware default table keyed on
  `systemInfo.osBuild` (already collected).
- **A-5 · Multi-key checks evaluate only the first resolving key** (`break`, app:3168): MON-003
  (Logon AND Logoff), DEF-024, NET-002/NET-008 ("all firewall profiles" passes on Domain alone).
  Fix: all-keys-must-pass semantics.
- **A-6 · Multi-machine aggregation is last-writer-wins fiction.** One shared check array; each
  import overwrites Status/ActualValue/Details (app:2935+); `AffectedMachines` only ever accumulates
  (never pruned on clean re-import) and — per the engine audit — `AffectedMachines`/`Machines`/
  `MachineCount` are **written but never read** by any UI/export. Manual Accepted-Risk/Deferred
  statuses are clobbered by the next import. Fix: per-machine result store keyed by hostname, real
  aggregate view, protect manual overrides, prune stale entries — or remove the multi-machine
  framing entirely.

## C. Evaluation defects

- **C-1 · Exact string equality penalizes stricter-than-baseline** (app:2941): AUTH-001 minLength
  15>14 Fails, AUTH-003 lockout 5<10 Fails, SEC-005, SEC-010, AUTH-021/022, DATA-002. Needs
  operator semantics (min/max) for numeric baselines.
- **C-2 · Privilege-rights checks (SEC-071…087) order/absence fragile**: secedit omits unassigned
  privileges entirely → SEC-077 (SeCreateTokenPrivilege, Critical, expects empty) **cannot Pass**;
  same SEC-081/086; multi-SID sets fail on ordering. Fix: absence == unassigned == empty; set-based
  order-insensitive SID comparison.
- **C-3 · threshold gaps**: `operator` field ignored (latent); `perAccount: true` (AUTH-025) not
  implemented (5-per-account applied as 5-total); `windowDays` display-only (actual window =
  collector LookbackDays); with `-EventSummaryOnly`, DEF-028 compares total event count vs 5 →
  guaranteed false Fail (eventIds filtering impossible on summary objects).
- **C-4 · Heuristic branch = permanent Warning**: NET-028 (**Critical**, whole tlsConfig object),
  MON-030, OPS-011/022/023, SEC-041/042 (admin/guest rename never evaluates). Also `-match
  'Success'`→Pass (app:3035) is a substring foot-gun.
- **C-5 · MDM fallback mostly fictional**: real MDM values live under `PolicyManager\current\device\*`
  which the collector enumerates as names only → Intune-managed compliant settings show Warning
  "verify in portal" instead of Pass; `mdmWinsOverGP` tie-break direction is correct but rarely has
  data. Fix: collect PolicyManager values for the mapped areas.
- **C-6 · Duplicates with diverging severity for the SAME data point**: SEC-018/OPS-012/OPS-024
  (EventLog service: High/Critical/Medium!), SEC-059/DATA-009 (Med/High), AUTH-013/UAC-006,
  NET-003/NET-023 (High/Critical). Same-severity duplicates: SEC-003/DATA-015, DEF-004/DEF-034,
  DEF-021/DEF-032 (same ASR GUID), AUTH-004/AUTH-020, DATA-001/DATA-020, OPS-005/OPS-027,
  MON-004/MON-008. Consolidate or align.

## E. Engine, scoring, governance

- **E-1 · Accepted Risk contradicts documented methodology**: report text says "excluded from both
  scores" (app:2508/2757); code keeps it in compliance denominator (app:1058) and scores it 0 at
  full weight in the weighted functions (`default` arm). Deferred=0 is intentional and consistent.
- **E-2 · Accept Risk has NO justification enforcement** (README:148 promises "mandatory
  justification"; handler app:1989 just sets status — one click, no dialog, Notes can be empty).
- **E-3 · Governance decisions clobber `Details`** with the literal tag string (app:1990), losing
  the evaluation context forever; **E-4 · Undo** restores Status but blanks Details and excludes
  'N/A' from undo-eligibility (app:1919 vs its own comment).
- **E-5 · Dead state/UI**: `ManualOverrides` dict write-only; `cmbBaselineVersion` and
  `cmbScoreView` fully wired via FindName but never read — controls do nothing.
- **E-6 · Compliance% vs weighted score silently disagree on Not-Assessed** (excluded vs
  0-in-denominator), shown side-by-side with no explanation; README's scoring table has NO row for
  Not Assessed. Sidebar peek score uses a third variant.
- **E-7 · Stale-check lifecycle broken in DUPLICATED code**: `Load-Assessment` (add-only merge) is
  **never called**; the live load path reimplements the same add-only logic inline
  (app:4354–4390). No catalog version stamping anywhere. Fix once, delete the dead twin.
- **E-8 · Dead concurrency scaffolding ACTIVELY polling at 20 Hz** all session (50ms
  DispatcherTimer draining always-empty queues, app:4452; `Start-BackgroundWork` has no callers).
- **E-9 · Autosave does full serialize+backup+prune every tick regardless of dirty state**
  (app:3348 — timer-only, better than AVD's, but needs an IsDirty short-circuit).
- **E-10 (clean)**: divide-by-zero guards + weight floors all correct. **E-11**: CSV formula
  injection present (app:3851); HTML + RTF escaping solid (RTF is LIVE here — clipboard button —
  unlike AVD; only non-ASCII `\u`-encoding nit). **E-12 (clean)**: 21 empty catches all reviewed as
  justified best-effort. **E-13/E-14 (clean)**: threading + closures correct.

## D. Baseline drift (web-verified July 2026)

Verified still correct: AUTH-001=14, AUTH-003=10, AUTH-004=5, SEC-005=900, SEC-009, SEC-010, UAC-003,
TLS/SSL family, CG/VBS/HVCI, DEF-010/017/027, SEC-058.

| id | issue | fix |
|---|---|---|
| AUTH-031 | legacy AdmPwd LAPS — deprecated Oct 2023; false-negative on all Windows LAPS deployments; **collector already reads the Windows LAPS keys, no check consumes them** | rekey to `LAPS\BackupDirectory >= 1` |
| SEC-062 | DisabledComponents=255 contradicts MS guidance (prefer-IPv4 = 0x20 max) | fix value or retire (also key dangling, A-2) |
| NET-034 | min SMB2 | 24H2 baseline: min SMB **3.0.0**, max 3.1.1 |
| NET-016 | Require SMB encryption True | MS baseline: workstation require-encryption **Disabled** (signing is the mandate) — over-enforced; re-label "stricter than baseline" or align |
| SEC-016 | NTLM deny-all (2) | 25H2 direction is audit (1); also key dangling (A-2) |
| DEF-018 | PSExec/WMI ASR Block (1) | 25H2 baseline: **Audit (2)** |
| SEC-029 | Spooler Stopped | CIS-style, not MS baseline — relabel origin |
| SEC-057 / SEC-065 | Spectre override / PSv2 | demote to informational (not in current baselines / feature removed in 24H2+) |
| DATA-011/012 | XTS-AES-256 | stricter than MS default (128) — relabel |

**Missing 2026 checks (add)**: **Kerberos RC4→AES enctype readiness** (DefaultDomainSupportedEncTypes
— enforcement live Apr 2026, audit-mode removal Jul 2026: the single biggest gap);
AllowAdministratorLockout=1; SMB auth rate limiter + audit-insecure-guest/signing; remote mailslots
disabled; UAC Enhanced Privilege Protection pair (24H2); MotW insecure-sources; PKINIT SHA-2;
dMSA-logon/sudo disables; Defender EDR-block-mode / warn-to-block; command-line-in-process-creation
(25H2 — also fixes SEC-019's intent); NTLM auditing; IPP/IPPS print policies.

MITRE: 25 unique IDs (metadata says 32 — count mismatch), all real/current/topically correct.

## R. Collector robustness

Good: hard admin gate, locale-independent parsing, gpresult sandboxed w/ timeout+fallbacks, per-area
try/catch. Defects: **R-1** errors vanish (see A-3); **R-2** event volume unmanaged (13×10k events on
a DC → tens of MB JSON, UI freeze on import; `[xml]$evt.ToXml()` re-parsed per-property — 70k parses
per query, hoist it); **R-3** Server/DC/Core edge cases (build map knows only Win11 client builds →
Server isSupported=null; Get-MpPreference absent → A-3 cascade; Secure Boot on BIOS → false Fail;
no-TPM VMs conflate "absent" with "unknown"; Xbox/Browser services absent on Server → false Fails);
**R-4** winrmConfig policy values only read when service running → stopped-WinRM (compliant!) →
SEC-007/008/038/039 false Fail.

## H. Pipeline, tests, hygiene

- **H-1 · Metadata provenance unpinned**: admx_metadata.json generated 2026-03-31 from the build
  machine's local PolicyDefinitions (not a pinned SCT package); csp_metadata.json has NO metadata
  stamp; neither regenerated since April. templates/ vintage unknown. Regenerate from pinned SCT
  25H2 + stamp both.
- **H-2 · Wasted collection**: appliedGPOs/deniedGPOs completely dead (gpresult is the most
  expensive+fragile step, 120s); **129 of 215 registry reads unconsumed** (incl. the Windows LAPS
  keys AUTH-031 needs, all SCHANNEL/WinRM/PowerShell/firewall-logging policy duplicates of dedicated
  sections). Trim or consume.
- **H-3 · Scoring text drift**: report claims Critical(5x)…Low(2x) multipliers; code uses catalog
  weight field.
- **Tests**: Test-BaselinePilot.ps1 tests parallel REIMPLEMENTATIONS, not production functions; its
  Resolve-CollectionKey copy has drifted (missing tlsConfig + MDM branches); unguarded
  divide-by-zero at :276 yields a spurious PASS; 16/51 pass without sample data; scoring/persistence/
  governance/multi-machine all untested.
- **Hygiene**: README promises justification (E-2) + omits Not-Assessed row + omits
  Run_Collection.bat/Test-BaselinePilot.ps1 from file structure + "36 services" (actual 37); FIVE
  version strings (header 0.1.0 / AppVersion 0.1.0-alpha / XAML v0.1.0 / checks 1.0 / collector
  1.0.0); folder BaselineAssessor vs product BaselinePilot unexplained; `.bat` files LF (pending
  `git add --renormalize` for the new .gitattributes); 3.9MB generated JSONs committed (decide:
  document or gitignore+build-step); `.gitignore` hardcodes `ANROMAN-*` machine prefix; checks.json
  has no BOM (good). Launchers are CLEAN (explicit -STA, elevation-gated collector) — better than
  pre-fix AVD. XAML orphans: colLeftPanel, lblSettingsVersion (+ lblTitleVersion found-never-bound).
  W-3: zero dangling FindNames.

---

## Fix status

_Pending — fix pass launched 2026-07-18 (collector/engine/catalog agents)._
