# W365Assessor Audit — July 2026

Full audit mirroring the AvdAssessor program (see ../AvdAssessor/AUDIT.md for conventions): all 23
Auto checks traced discovery→importer→catalog, all 128 checks drift-checked against July-2026
Microsoft guidance, engine diffed against the (now-fixed) AvdAssessor defect set, 163 reference URLs
liveness-tested. `disc:` = Invoke-W365Discovery.ps1 (v0.1.0), `app:` = W365Assessor.ps1.

Overall: the **catalog is unusually current** (April 2026: already covers Cloud Apps, UX Sync, DR
Plus, Boot/Switch, Windows App, watermarking, RDP Multipath) and mapping has **zero mis-routes**.
Damage concentrates in: one dead-on-arrival check, 4 importer orphans, silent collection failures,
the full inherited engine defect set, a pervasive product rename, and mass URL rot.

## A. Critical correctness

- **A-1 · IMG-007 (gallery image end-of-support) never runs.** disc:711 computes
  `[datetime]$gi.EndOfSupportDateTime - $now` but `$now = Get-Date` is only defined at disc:749,
  below the loop → op_Subtraction throws on $null → caught → continue, for every image. Fix: hoist
  `$now = Get-Date` above disc:708.
- **A-2 · auditEvents `$filter` datetime malformed (KNOWN — GitHub PR #1).** disc:476:
  `.ToString('o')` emits a local-offset stamp whose `+` breaks the Graph filter (BadRequest for
  UTC-positive clients) → MON-007 false-Warning, GOV-009 empty. Fix per PR #1:
  `.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')`. Verified this is the ONLY dated/special-char
  `$filter` in the collector — PR's fix is complete.
- **A-3 · Collection failure = silent vanish, and the GUI discards the diagnostics.** Every section
  try/catches into `Add-DiscoveryError` (disc:280) → `$Discovery.Errors`, which W365Assessor.ps1
  **never reads anywhere**. Missing `CloudPC.Read.All` is indistinguishable from "no Cloud PCs
  exist". Fix: emit Status 'Error' CheckResults per failed section AND surface `$Discovery.Errors`
  in the import summary.

## M. Mapping / importer (app:1330–1400)

- **M-1 · 4 of 23 Auto checks orphaned** — emitted by discovery, defined Auto in catalog, **no
  importer switch case**: W365-PROV-004 (Autopatch, disc:627), W365-PROV-005 (Grace Period,
  disc:641), W365-USER-003 (Restore Points, disc:597), W365-USER-004 (Self-Service Reset, disc:613).
  Only 19/23 Auto checks reachable; `_metadata.autoMappedCount: 23` and README's automated-checks
  table overstate by 4. Fix: add the four cases.
- **M-2 · No importer type-guard** (app:1395 sets Status/Source unconditionally — latent, all
  current targets are Auto) and **no breaks in switch -Wildcard** (last-match-wins). Port the AVD
  fix: first-match with break + `$Match.Type -eq 'Auto'` guard.
- **M-3 · Error conflated with Fail**: `$StatusRank` (app:1330) ranks Error=0=Fail and the
  aggregation rewrites Error→'Fail'. Port AVD's Error→Not-Assessed-with-reason handling (needed once
  A-3 starts emitting Error).
- Mis-routes: none. Cross-maps verified correct (PROV-002→IAM-001 SSO, PROV-003→SEC-001,
  CPC-003→COST-001, CPC-001+CPC-002→CPC-001).

## C. Check logic

- **C-1 · COST-001 uses the wrong inactivity signal**: `LastModifiedDateTime` (bumps on any config
  change). `LastLoginResult` is already collected (disc:308) and unused; the proper source is
  `reports/getInactiveCloudPcReport`.
- **C-2 · CPC-001 evaluates only failed+inGracePeriod**; its `$badStates` array (disc:769) is dead —
  description promises full unhealthy-state coverage.
- **C-3 · PROV-005 `gracePeriodInHours` property validity suspect** (grace is a fixed 7-day service
  behavior) → likely always 0/Warning; verify against live schema or reframe the check.
- **C-4 · NET-001 health-detail enrichment is dead** (disc:669 reads
  `HealthCheckStatusDetails.endDateTime`/`.failedHealthCheckItems`; real schema is `healthChecks[]`)
  — harmless, details always empty; fix or remove.

## E. Engine — inherited AvdAssessor defects, ALL PRESENT (port the AVD fix set)

(a) N/A scores 100 + Not-Assessed 0, both in denominator; README:65/288 claims both excluded
(app:1074–1189). Align to the decided semantics: N/A excluded, Not Assessed 0-in-denominator.
(b) Maturity colors: 5-tier (app:1199, strip) vs ≥80/≥50 everywhere else (~15 call sites incl.
app:1543/2822/4175/4309) — port `Get-MaturityZoneColor`.
(c) AutoSave calls Set-Dirty on itself (app:4730) + per-interaction autosave (app:2552, 2598) —
port debounce.
(d) Exclude checkbox promised (app:2170) not implemented; `exclusion_expert` unreachable — port the
checkbox. Also `speed_demon` (app:5119) is never unlocked anywhere — wire or remove.
(e) No stale-check pruning; `_metadata.version` never read — port prune + CatalogVersion stamp.
(f) CSV formula injection (app:4636) — port `'`-prefix fix.
(g) Dead `Get-ExecutiveSummaryRtf` (app:4131) + `Start-BackgroundWork` scaffolding (app:190) — remove.
(h) Version drift: header 0.1.0 / AppVersion 0.1.0-alpha / XAML v0.1.0 ×3 (`lblTitleVersion` found
but never bound, `lblSettingsVersion` never FindName'd) — consolidate + bind.
(i) 12 empty catch blocks (app:448, 488, 507, 671, 849, 3634, 4750, 4752, 4798, 5012, 5949, 5982)
— add debug logging.
HTML/RTF escaping verified solid; keep. "Effort-vs-impact matrix" in README is actually effort-only
(app:3353) — fix the wording (or add the impact axis).

## D. Guidance drift (validated July 2026)

- **D-1 · Frontline → Flex rename (Ignite Nov 2025, effective ~May 2026)** — pervasive: every
  "Frontline" in names/descriptions/`appliesTo` (INV-003, PROV-003/010/011/012/014/015, COST-003,
  USER-006 + all appliesTo arrays). Also fix the malformed token `"Frontline (dedicated)"`
  (USER-002, USER-003) → canonical enum + description caveat. Keep "(formerly Frontline)" note in
  INV-003.
- **D-2 · RD client retirement is PAST** (MSI EoS 27 Mar 2026; Store app May 2025): UX-001/UX-007
  reframe from "prepare to migrate" to "RD client retired; Windows App required"; consider High.
- **D-3 · GPU Cloud PCs GA** (Sep 2025, GPU Select SKU, supported on Flex shared): UX-005 update +
  appliesTo widen. Watermarking/screen-capture now native via settings catalog for Cloud PCs:
  SEC-007/SEC-017 update + appliesTo widen to Flex.
- **D-4 · SEC-003**: current Cloud PC security baseline is **24H1** — reference it.
- **D-5 · DR Plus GA Apr 2025** (Enterprise-only, RTO ≤30min) + CRDR GA (<4h): USER-008/USER-002/
  GOV-008 drop preview framing, add RTO/RPO figures. Boot/Switch GA: UX-009/UX-003 drop preview,
  note switch-back-to-local.
- Beta→v1.0: cloudPCs, provisioningPolicies (incl. autopatch), userSettings are GA — migrate off
  `/beta` (keep beta only for auditEvents); bump discovery version.

## F. Automate more (23 → ~40 Auto)

- **Over-consent**: `DeviceManagementConfiguration.Read.All` + `DeviceManagementManagedDevices.Read.All`
  are requested (disc:216) and drive ZERO calls. Collected-unused: ServicePlans (whole section),
  ExternalPartnerSettings (never populated), LastLoginResult.
- **P1 zero-new-calls** (data in memory): INV-004 (ServicePlans), INV-003 (provisioningType mix),
  PROV-002 (cloudPcNamingTemplate), PROV-003/PROV-010 (Flex/sharedByUser + Cloud Apps), PROV-006 +
  IAM-002 (domainJoinConfigurations type), PROV-007 (windowsSetting).
- **P2 Reports API** (`virtualEndpoint/reports/*`, scope already held): COST-001 (fix C-1) +
  COST-002 (`getCloudPcRecommendationReports`/`getInactiveCloudPcReport`), MON-002
  (`getConnectionQualityReports`), MON-010 (`getResourcePerformanceReport`), MON-008, UX-002.
- **P3 Intune** (use the consented scopes or drop them): SEC-002 (managedDevice health), SEC-004
  (compliancePolicies), SEC-003 (baseline profiles), MON-001 (Endpoint Analytics), MON-005 (update
  compliance).
- **P4 new scope Policy.Read.All** (mirror AVD): IAM-003 CA + IAM-004 MFA targeting Cloud PC /
  Windows Cloud Login app IDs.
- Partials: SEC-006 CMK via `diskEncryptionState`; PROV-011 UX Sync property.

## G. New checks

IAM-010 (Auto, High): CA targets Windows Cloud Login / Cloud PC app. IAM-011 (Auto/Manual, Med):
token protection + every-time sign-in frequency. SEC-019 (Manual, Low): Windows App/Connection
Center posture. INV-006 (Manual, Low, optional): Windows 365 Link endpoint posture. (IPv6 and
Entra-Kerberos gaps: not applicable to W365 — skip.)

## H. Metadata, URLs, docs, hygiene

- **H-1 · 60 of 163 reference URLs dead (37%), 54 of 128 checks affected** — full list in the
  audit transcript scratchpad (w365_all_dead.txt); includes core checks (IAM-001, USER-001,
  SEC-006/007/010, MON-001/003, CPC-001, IMG-007). Every dead URL needs a verified current slug.
  103 alive.
- **H-2 · Severity↔weight outliers (5)**: INV-004/MON-007/GOV-009/USER-006 Medium w2 (peers w3);
  NET-001 High w5 (peers w4). Effort/origin enums 100% clean. Genuine id gaps: W365-SEC-005,
  W365-USER-005 (document as retired). Field coverage 100% (bestPractice/impact/remediation/
  appliesTo/reference all populated).
- **H-3 · Near-duplicate**: IAM-007 vs SEC-016 (idle/lock timeout) — consolidate or cross-reference.
  SEC-006 vs SEC-012 — cross-reference note only.
- **H-4 · README**: overstates automation by 4 (M-1); "effort-vs-impact matrix" wording; -InactiveDays/
  -ImageAgeWarnDays params undocumented in Quick Start; scoring section wrong (E-a). Otherwise
  verified accurate (category/severity/origin tables exact).
- **H-5 · Launcher**: no `-STA`, `/MAX` no-op, recursive Unblock-File over output dirs, apostrophe
  quoting — port the AVD launcher fixes. XAML orphans: colLeftPanel, lblSettingsVersion.
  checks.json has NO BOM (good — keep it that way).

---

## Fix status

_Pending — fix pass launched 2026-07-18 (discovery/engine/catalog agents). PR #1 fix incorporated
into the discovery work regardless of merge timing._
