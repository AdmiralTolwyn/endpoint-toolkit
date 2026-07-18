# AvdAssessor Audit — July 2026

Full-surface audit: every Auto check's discovery/evaluation logic validated against code and current
Microsoft guidance, every Manual check assessed for automatability, all 173 catalog entries checked for
metadata consistency and reference liveness, engine (mapping/scoring/persistence/export) audited, plus
static analysis and docs verification. Findings are ID-tagged (`A-` critical correctness, `C-` check
logic, `M-` mapping, `E-` engine/scoring, `D-` guidance drift, `H-` hygiene/docs). Line references:
`disc:` = Invoke-AvdDiscovery.ps1, `app:` = AvdAssessor.ps1.

Verified counts: 173 checks (82 Auto / 91 Manual), 85 discovery ID prefixes, 85 wildcard mapping rules,
133 unique reference URLs (130 live, 3 dead).

---

## A. Critical correctness — silently wrong results at scale

- **A-1 · Cross-subscription discovery only queries the LAST subscription.** Everything after the
  per-sub loop closes (disc:1846–2238) runs on whatever Az context the loop left behind, and
  GOV-POLICY / GOV-BUDGET explicitly use `$Discovery.Subscriptions[-1]`. In multi-sub estates,
  **GOV-ORPHDISK, GOV-ORPHNIC, SEC-KEYVAULT, SEC-KVPE, GOV-POLICY, MON-ALERTS, GOV-QUOTA,
  GOV-CAPRESERV, GOV-BUDGET, GOV-RI, NET-NETWATCHER, NET-PRIVDNS** miss every subscription but one.
  Fix: move these blocks inside the per-sub loop (GOV-TAGALL is safe — it resolves by ResourceId).

- **A-2 · Missing required modules make three checks vanish.** `$RequiredModules` (disc:201–215) omits
  **Az.KeyVault** (`Get-AzKeyVault`, disc:1919/1925) and **Az.Security** (`Get-AzSecurityPricing`,
  disc:1675). With `$ErrorActionPreference='Stop'`, command-not-found is swallowed by try/catch and
  **SEC-KEYVAULT, SEC-KVPE, MON-DEFENDER emit nothing** — indistinguishable from "no Key Vaults exist".
  Related: the array literal has a `foreach { Import-Module … }` spliced *inside* `@( … )`
  (disc:209–212) — works by accident, paste artifact, move it out.

- **A-3 · IAM-004 RBAC wildcard is reversed — check can never Pass.** disc:1616 matches
  `-like '*Virtual Desktop*'`, but Azure's built-in roles are named **"Desktop Virtualization …"**.
  Always Warning regardless of reality. Also only inspects host-pool scope (assignments normally live
  on app groups) and never flags the actual anti-pattern (Owner/Contributor at AVD scopes).

- **A-4 · SEC-WM / SEC-SCP (watermarking, screen-capture protection) can never detect the feature.**
  disc:568–586 greps host-pool `CustomRdpProperty` for "watermark"/"screen capture protection", but
  these are configured via GPO/Intune registry policy (`fEnableWatermarking`,
  `fEnableScreenCaptureProtect`), never via RDP properties. Permanent false "not configured" Warning.
  Reclassify SEC-002/SH-007 to Manual or read via Intune Graph/Run Command.

- **A-5 · SEC-CLIP / SEC-PRINT: "unset = disabled = Pass" is wrong for pre-mid-2025 host pools.**
  Microsoft's secure-by-default redirection change (clipboard/printer off when unset) applies only to
  host pools created after the 2025 rollout and is not retroactive. The large installed base with no
  custom RDP properties has redirection **enabled** while the tool reports Pass (disc:473–493, 538–545).
  Return Warning for unset values with a creation-date caveat.

- **A-6 · PROF-020/022/023 false-Pass on Azure Files defaults.** `$null -match '<bad value>'` → false →
  Pass, but unset SMB settings mean the effective defaults still allow **SMB 2.1, RC4-HMAC, NTLMv2**
  (disc:1791–1821). Unhardened accounts report compliant. Treat null/unset as Warning/Fail; Pass only
  on explicit hardening. (Made urgent by D-6: Windows' Apr 2026 RC4→AES default change.)

- **A-7 · FSLogix storage detection is a name-substring heuristic.**
  `$SA.StorageAccountName -match 'fslogix|profile|avd'` (disc:1693) gates **every PROF-* check plus
  BCDR-STOR**. A profile account named `stprod01` → zero PROF results; unrelated accounts matching
  `avd` (e.g. `mavdata`) get audited as FSLogix. Classify by evidence instead (file shares present,
  Entra/AD Kerberos auth enabled, share names) with user confirm/override.

- **A-8 · SH-002 passes the exact default it exists to catch.** Pass when
  `MaxSessionLimit -le 999999` (disc:430–432); 999999 is Azure's "unlimited" sentinel and the check's
  own description calls it the failure condition. Never-configured pools report Pass.

- **A-9 · SH-025 / SH-026 (Secure Boot, vTPM) only evaluate VMs that already have a SecurityProfile**
  (disc:764–789). Standard-security VMs — the ones that should Fail — emit nothing. Emit Fail when the
  model loads but SecurityProfile/UefiSettings is null (mirror the SEC-TL else-branch).

- **A-10 · Inconsistent error handling: false Fails and silent vanishing.** When `Get-AzVM` fails,
  SEC-TL asserts "Fail — Standard security type" (fabricated conclusion, disc:753–760) while sibling
  checks silently emit nothing. 25 empty catch blocks across both scripts (11 app, 14 disc) swallow
  errors invisibly — permission/throttling failure is indistinguishable from "resource absent".
  Fix: emit `-Status 'Error'` in catches (mapping already ranks Error worst, app:1325).

## B. Always-Pass / never-Pass checks distorting the score

- **B-1** SH-017 (load balancing), APP-001 (app group config), SEC-DISK/SH-013 hardcode `Pass` with no
  evaluation (disc:605, 1063, 832) — full weight awarded, zero assurance. SEC-CAM/SEC-AUDIO are
  literal tautologies (`if … 'Pass' else 'Pass'`, disc:523/533).
- **B-2** SH-003 (VM sizing) only ever emits for B-series (Warning) — correctly-sized fleets stay
  "Not Assessed" forever; the sizing-tier validation in its description is never performed (disc:924).
- **B-3** SEC-ATTEST only ever emits Warning — SEC-019 can never show green (disc:994). GOV-DISKSKU
  (SH-022) only emits on Premium+Pooled — never Passes.
- **B-4** SEC-MDE counts the **retired MMA agent** (`MicrosoftMonitoringAgent`, retired Aug 2024) as
  Defender for Endpoint → false Pass (disc:683). Drop it; note extension-less MDE onboarding paths.

## C. Other evaluation defects

- **C-1** NET-RDP/NET-SSH (NET-004) read only singular `DestinationPortRange` — multi-port
  (`DestinationPortRanges`) and range (`3000-4000`) rules are invisible; deny-rule priority and
  NIC-level NSGs ignored (disc:1400–1458).
- **C-2** NET-HUBFW peering evidence is dead code: property is `RemoteVNet`, code reads
  `RemoteVNetId` → always null (disc:1513); pass/fail masked by the `-or ($AzFirewalls.Count -gt 0)`.
- **C-3** NET-NSG counts `GatewaySubnet`/`AzureFirewallSubnet`/`AzureBastionSubnet` in the "subnets
  without NSG" denominator — false Warning on shared VNets (disc:1264–1275).
- **C-4** BCDR-AZ passes when hosts are all pinned to a single zone — never checks spread; Warns in
  regions without AZs (disc:840–850).
- **C-5** GOV-SPACTIVE tests schedule *presence*, not `ScalingPlanEnabled` on the host-pool
  reference — a disabled plan with schedules Passes (disc:1133–1143).
- **C-6** MON-ALERTS only counts metric alerts (`Get-AzMetricAlertRuleV2`) though MON-015's definition
  says "metric or log" — add `Get-AzScheduledQueryRule` (disc:2067). MON-DIAGCAT misses
  `CategoryGroup`-based (allLogs/audit) diagnostic settings → false "categories missing" (disc:1585).
- **C-7** NET-DNS warns on default Azure DNS even for cloud-native Entra-joined estates (join type is
  already collected but unused); NET-NATGW warns when a firewall/UDR egress design exists; NET-UDR
  passes on any route table without checking a 0.0.0.0/0 NVA hop. NET-AN warns on SKUs that don't
  support accelerated networking (B-series).
- **C-8** IAM-001 mislabels most Hybrid-joined hosts as "AD DS" (hybrid = AD join + Connect sync; the
  AAD extension test is wrong signal) and reports Warning when extension data simply failed to load.
- **C-9** OPS-005 minimum agent version is hardcoded at `1.0.8431.0` (~2023) — compare against fleet
  max or a maintained constant. OPS-002 mixes local `Get-Date` with UTC heartbeat (minor). PROF-005
  TLS compare works only by lexical accident (`-ge 'TLS1_2'`).
- **C-10** SH-024 flags deallocated hosts as Warning — false positives for off-hours autoscale.

## M. Mapping layer (app:1330–1421) — verified three ways programmatically

- **M-1 · Three mis-routed wildcard rules (critical):**
  - `SEC-MDE-*` → **SH-006** (AV exclusions) but the discovery check tests MDE deployment = **SEC-006**.
  - `BCDR-MULTIREGION` → **BCDR-008** (Manual, risk-acceptance judgment) instead of **BCDR-009**
    (Auto, "Multi-Region Host Pool" — currently orphaned).
  - `BCDR-SPSCHED-*` → **BCDR-012** ("VNet for Failover Region", Manual, topically unrelated).
  Two of these silently overwrite Manual checks with automated data. **Guard the importer:** require
  `$Match.Type -eq 'Auto'` before overwriting Status/Source (app:1424–1462).
- **M-2 · 5 of 82 Auto checks are unreachable** by any mapping and stay "Not Assessed" forever,
  permanently deflating their categories: **BCDR-009, OPS-004, SEC-006, SEC-018, SH-018**
  (SEC-006/BCDR-009 fixed by M-1 repointing; SH-018 region-proximity is trivially implementable from
  already-collected locations; OPS-004 → implement or reclassify Manual; SEC-018 → route `SEC-RDP-*`
  here instead of SH-011, matching severity/weight intent).
- **M-3 · Dead rule / dropped data:** `NET-SHORTPATH-*` rule exists but discovery never emits it
  (NET-011 functionally dead). Discovery's `SH-001-<hp>` ("Host Pool Type", disc:408) has **no**
  mapping case and is silently dropped — also an ID collision with checks.json SH-001 (Trusted Launch).
- **M-4** No wildcard shadowing found (patterns include trailing `-`; verified with `-like` tests) —
  but the `switch -Wildcard` has no `break`s: last-match-wins and O(checks×patterns). Add breaks or a
  prefix hashtable.

## E. Engine, scoring, persistence, exports

- **E-1 · README contradicts code on N/A scoring (all three claims wrong).** README:41/230/516 say
  "N/A and Not Assessed are excluded from numerator and denominator". Code (app:1069–1097, verified):
  **N/A scores 100 like Pass; Not Assessed scores 0; both stay in the denominator**; only the
  `Excluded` flag removes a check. Sibling W365Assessor inherits the identical behavior; BaselinePilot
  truly excludes N/A. Decide the intended semantics once and align code+docs (recommended: exclude N/A
  from the denominator — "not applicable" shouldn't award points).
- **E-2 · The Exclude feature doesn't exist.** `Render-AssessmentChecks`' docstring promises an
  exclude checkbox; no such control exists in code or XAML, `.Excluded = $true` is never set anywhere,
  and the `exclusion_expert` achievement is unreachable. Implement or remove.
- **E-3 · Stale-check lifecycle is broken.** Load-Assessment only *adds* new definitions; checks
  removed from checks.json linger forever with frozen metadata and keep scoring.
  `checks.json _metadata.version` is never read; saved assessments carry no schema version. Prune
  retired IDs on load (with a toast) and stamp/check a schema version.
- **E-4 · Maturity color thresholds disagree within one report.** Dashboard + maturity strip use the
  5-tier scale (app:1194, 1843), but Build-HtmlReport dimension coloring (app:3362) and the email
  summary (app:4193) use `≥80/≥50` — a 78% dimension is teal "Managed" live but orange in the export.
  Single `Get-MaturityZoneColor` helper.
- **E-5 · Dirty/autosave defects.** `AutoSave-Assessment` calls `Set-Dirty` on itself (app:4612) — the
  dirty dot re-lights within 60s of every manual save. Status/notes handlers call AutoSave directly:
  full backup-rotation + directory scan per dropdown change (173+ churns per assessment). Debounce to
  the timer; remove the Set-Dirty.
- **E-6 · Exports.** HTML/email encoding is solid (all interpolations HtmlEncoded — verified).
  CSV: no neutralization of leading `=+-@` in Details/Notes → Excel formula injection; prefix with `'`.
  `Get-ExecutiveSummaryRtf` (158 lines) is dead code with an incomplete RTF escaper.
  `Start-BackgroundWork` + SyncHash/timer scaffolding (~100 lines) is never called — dead concurrency
  infrastructure.
- **E-7 · Perf.** Discovery: `Get-AzVM` (model+status) called per host, then the networking loop
  repeats both per host again — ~4–5 sequential ARM calls per session host; batch per-RG
  (`Get-AzVM -ResourceGroupName` once) or use Resource Graph. GOV-TAGALL re-fetches resources already
  in memory. App: Update-Progress recomputes Get-CategoryScore per category on every status change.

## D. Guidance drift (validated against learn.microsoft.com / TechCommunity, July 2026)

- **D-1 · Windows Cloud Login app transition (IAM-002/003/008, coordinated edit).** CA/MFA guidance now
  targets **Azure Virtual Desktop + Windows Cloud Login** (`270efc09-cd0d-444b-a71f-39af4910ec45`);
  "Microsoft Remote Desktop" is legacy. Also mention "Every time sign-in frequency" (GA Apr 2025) and
  SSO's CA dependency on Windows Cloud Login; KDC proxy → IAKerb transition for hybrid.
- **D-2 · Teams: SlimCore / VDI 2.0 (APP-009, OPS-004).** Legacy WebRTC redirector: End-of-Support
  **Oct 1 2026**, stops working **Apr 1 2027**. Rewrite APP-009 around SlimCore; replace "WebRTC
  redirector" in OPS-004's component list. Consider weight bump — this is now time-critical.
- **D-3 · App Attach (APP-002).** MSIX App Attach **retired June 1 2025**; rename check to "App
  Attach" (CIM-based, MSIX+App-V+partner formats, Server 2022/2025 support).
- **D-4 · Session Host Update GA June 2026 (SH-016/SH-020).** Drop "(preview)"; reference session host
  configuration + managed identity (no registration token); consider raising weight from Low.
- **D-5 · Trusted Launch (SH-001 remediation).** "Cannot be changed in-place" is outdated — in-place
  enablement for existing Gen2 VMs is supported; effort → "Some Effort".
- **D-6 · Kerberos RC4→AES default change, April 2026 (PROF-022).** No longer mere hardening: FSLogix
  shares not on AES-256 can lose access. Raise to High; pair with a readiness check (see additions).
- **D-7 · RDP Shortpath (NET-007/NET-011).** Managed-networks config via Intune/GPO GA Jan 2026;
  Shortpath over Private Link GA Feb 2026; TURN in 39 regions. NET-007 Low/2 understates it. Private
  Link trio NET-005/NET-014/SEC-024 should be consolidated and mention Shortpath-over-Private-Link.
- **D-8 · FSLogix (PROF-008/012/018).** Current baseline FSLogix 26.01 CU1 (Mar 2026); Cloud Cache is
  now the *recommended* multi-region HA/DR mechanism (soften "evaluate whether justified").
- **D-9 · Secure-by-default redirections for new host pools GA Jul 2025** — note in SH-011 (and see A-5).
- **Clean categories (no drift found):** Landing Zone, BCDR, Governance & Cost, Monitoring.

## F. Automate more (39 of 91 Manual checks fully automatable; prioritized)

Discovery currently collects **zero Microsoft Graph data** and no Log Analytics signals.

- **P1 — Graph: Conditional Access + MFA** (`Policy.Read.All`): converts IAM-002, IAM-003, IAM-010;
  inspect includeApplications for the three AVD app IDs, grantControls, sessionControls. The single
  highest-value addition.
- **P2 — Graph: authn methods registration** (`AuditLog.Read.All`): IAM-002 depth, IAM-009.
- **P3 — Log Analytics KQL** (workspace already discovered via diagnostic settings):
  WVDConnections/WVDCheckpoints (latency → automates NET-008; logon duration), Perf counters, FSLogix
  events → MON-002/005/008/009/010, PROF-010 with real thresholds instead of existence checks.
- **P4 — Intune Graph** (`DeviceManagementConfiguration.Read.All`): SH-028 baselines, SH-014 drift,
  SH-005 patch posture, SEC-001/003/004, PROF-007 (OneDrive KFM), PROF-019 (AV exclusions).
- **P5 — Parse scaling-plan schedules already in memory** (low effort): fixes B-1's SH-017 always-Pass
  (verify BreadthFirst ramp-up/DepthFirst peak), enriches GOV-009, personal-pool autoscale.
- **P6 — Defender**: `Get-AzSecuritySecureScore` (SEC-005), `Get-AzJitNetworkAccessPolicy` (SEC-021).
- **P7 — Azure Update Manager** assessment REST: SH-005, MON-013.
- **P8 — Run Command in-guest FSLogix registry** (highest yield, highest cost): PROF-001/008/009/010/
  013/014/015, PROF-012 versions. Requires running VMs; make opt-in.
- Trivial ARM conversions flagged Manual today: BCDR-006 (`Get-AzGalleryImageVersion` TargetRegions),
  GOV-007 (`Get-AzResourceLock`), BCDR-011 (capacity reservation in DR region), SH-020 (ARM property),
  APP-002 (`Get-AzWvdMsixPackage`), MON-003 (ServiceHealth activity alerts), NET-005/014 (ARM
  privateEndpointConnections), APP-004 (role assignments PrincipalType=Group), MON-012 (Sentinel
  onboardingStates), SEC-013 (regulatory compliance API).

## G. New check proposals (deduped across agents)

| id | Category | Sev | Type | What |
|---|---|---|---|---|
| SH-029 | Session Hosts | High | Auto | OS past end-of-support (Win10 EOS Oct 2025) without ESU — **currently uncovered** |
| SH-030 | Session Hosts | Med | Auto | GPU pools: driver extension present + GPU acceleration policy |
| SH-031 | Session Hosts | Med | Auto | Personal-pool autoscale / hibernate (GA) configured |
| SH-032 | Session Hosts | Low | Auto | Confidential VM guest attestation (parallel to SEC-019 for TL) |
| IAM-011 | Identity | High | Auto | CA policies cover **Windows Cloud Login** app (D-1 dependency) |
| IAM-012 | Identity | Med | Auto | Entra Kerberos on profile storage for cloud-native (no AD DS needed) |
| APP-010 | App Delivery | Med | Manual | SlimCore migration before WebRTC EoS Oct 2026 (D-2) |
| APP-011 | App Delivery | Low | Manual | Legacy MSIX App Attach → CIM App Attach migration |
| OPS-006 | Operations | High | Manual | Windows App client migration (RD MSI client EoS Mar 2026) |
| PROF-026 | FSLogix | High | Manual | Kerberos AES readiness before Apr 2026 default change (D-6) |
| PROF-027 | FSLogix | Low | Manual | Hibernate ↔ FSLogix/App Attach incompatibility documented |
| MON-017 | Monitoring | Low | Auto | Scaling-plan-level diagnostic settings (autoscale audit trail) |
| NET-022 | Networking | Med | Auto | FQDN allow-list for optional features behind firewall egress |
| NET-023 | Networking | Low | Auto | RDP Multipath (GA Jul 2026) enabled |
| SEC-025 | Security | Med | Manual | Purview **Endpoint DLP** RDP restrictions — replaces APP-006, which is conceptually invalid (sensitivity labels cannot attach to app-group ARM objects) |

## H. Catalog metadata & hygiene

- **H-1** 14 duplicate/overlap clusters distort composite scoring; worst: Private Link ×3
  (NET-005/NET-014/SEC-024), alerting ×3 (MON-005/007/015 — 007 Manual vs 015 Auto), version currency
  ×3 (OPS-004/OPS-005/PROF-012), Trusted Launch triple-weighted (SH-001+SH-025+SH-026 = weight 12),
  GOV-004(Manual)/GOV-016(Auto) same budget check. Consolidate or cross-reference.
- **H-2** Weight/scale mismatches: IAM-006 (Medium, weight 2 — and arguably High: tenant lockout),
  GOV-006 (Low, weight 1). Four checks use `"Significant Effort"` — not in the declared effortScale
  (GOV-017, NET-020, NET-021, SEC-024). 5 origin strings out of canonical order. ID gaps (GOV-010,
  MON-006, OPS-001, SH-012, SH-023) — document as retired.
- **H-3** Dead reference URLs (3): APP-003 `printing-on-avd`; OPS-002+SH-024 `troubleshoot-vm-connectivity`
  (used twice); SH-003 WAF `session-host-compute`. Replacements identified in agent report.
- **H-4** Description issues: BCDR-002 demands GRS but Premium Files (required by PROF-003) supports
  only LRS/ZRS — impossible combination; APP-006 conceptually invalid (see SEC-025); APP-009 stale
  "Teams 2.0" framing.
- **H-5** README: prerequisites table omits **Az.Resources, Az.Network, Az.PrivateDns** (its
  Install-Module command fails the script's own gate) and lists Az.KeyVault which isn't gated (see
  A-2); discovery version cited as 0.2.0 in 3 places (actual 0.3.0); "41 functions ~4900 lines"
  (actual 47/5876). App header says 0.1.0, `$Global:AppVersion` says 0.1.0-alpha, XAML Settings tab
  hardcodes `v0.1.0` (`lblSettingsVersion` never bound — bind like `lblSplashVersion`).
- **H-6** PSScriptAnalyzer: 685+95 warnings, 0 errors. Actionable: 25 empty catch blocks (see A-10),
  8 dead FindName variables, `$sender` shadowing, unapproved verbs (cosmetic). 622 PSAvoidGlobalVars =
  by-design architecture.
- **H-7** Environment: no `.gitattributes` (Windows tool edited on macOS; LF-only today by accident —
  pin `*.ps1 *.bat *.xaml` EOLs). Launcher: add explicit `-STA`; `/MAX` is a no-op; recursive
  Unblock-File scans assessments/reports/backups on every launch; apostrophe-in-path breaks quoting.
  checks.json BOM breaks strict JSON parsers (fine for PowerShell). 2 dead XAML names
  (`colLeftPanel`, `lblSettingsVersion`).

---

## Per-check verdict summary

82 Auto checks validated end-to-end. Clean: 34. Defective logic: 21 (A-3..A-10, B-1..B-4, C-1..C-10).
Mis-mapped: 3 (M-1). Orphaned/unreachable: 6 (M-2, M-3). Guidance-stale: 17 (D-1..D-9, overlapping).
91 Manual checks: 39 fully automatable (F), 46 partially, 6 genuinely manual; 1 conceptually invalid
(APP-006). Full per-check tables live in the agent transcripts; the fix order that pays best:
**A-1, A-2, M-1+importer type-guard, A-3, A-4/A-5 (reclassify), A-6, A-7 — then D-1/D-2/D-6 catalog
edits, then F-P1/P5 automation quick wins.**

---

## Fix status (applied 2026-07-18, same day as audit)

**Phase 1 — all A/B/C/M/E/D/H findings fixed** across Invoke-AvdDiscovery.ps1 (v0.4.0),
AvdAssessor.ps1 (v0.2.0), checks.json (v1.1), README.md, Launch_AvdAssessor.bat; repo `.gitattributes`
added. Deliberate scoring decision: N/A excluded from numerator+denominator; Not Assessed = 0 in
denominator; imported Error → Not Assessed with reason in Details. Deferred by design: H-1 duplicate-
cluster consolidation (disruptive to saved assessments — cross-references added instead).

**Phase 2 — automation expansion** (discovery v0.5.0, catalog v1.2): new Microsoft Graph identity
collection (CA/MFA/token-protection/passwordless; degrades to Error without Policy.Read.All /
AuditLog.Read.All); 16 Manual→Auto conversions (IAM-002/003/009/010, BCDR-006/011, GOV-007, SH-020,
APP-002/004, MON-003/012, NET-005, SEC-005/013/021); SH-018 orphan implemented; NET-011 → Manual;
6 new checks added (IAM-011/012, SH-029/030/031, MON-017) plus phase-1's APP-010, OPS-006,
PROF-026, PROF-027. **Final catalog: 183 checks — 99 Auto / 84 Manual** (was 173 / 82 / 91).

**Phase 3 — independent adversarial verification: SHIP-WITH-NOTES.** 0 parse errors, 0 PSSA
error-severity findings, emit→arm→catalog mapping verified complete and unambiguous in both
directions (108 emit shapes, 104 arms, 99/99 Auto reachable, 0 shadowing), empty catches 25 → 3
(all justified). Notes (report-only, low impact): `$SubShort` 8-hex suffix collision is theoretical;
SEC-JIT's RDP-exposure gate scans NSG results estate-wide rather than per-sub (over-flags,
conservative); BCDR-DRCAP "primary region" is alphabetical-first heuristic; 3 benign scalar
`-ne $null` PSSA style warnings. Not yet validated against a live tenant — new Graph/Az call shapes
are documented-behavior + try/catch-to-Error; first live discovery run should be observed.
