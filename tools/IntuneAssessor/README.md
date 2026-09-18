# Intune Discovery for Assay

Version 0.5.8. Experimental pending a live tenant pilot. Offline-tested with
Windows PowerShell 5.1 and PowerShell 7.

The [contract audit](AUDIT.md) currently verifies 49 enabled Graph GET contracts
and quarantines two assignment reads. Nested values and the live tenant pilot
remain separate gates; do not label the full collector audit complete.

Exports read-only observations for Assay's Intune pack. Assay owns control
definitions and scores: 50 source-linked controls, two bounded automatic checks,
and 48 evidence-assisted manual checks. The collector does not mutate a tenant.

## Mixed Setting Decoder Repair (0.5.8)

The setting decoder now rejects a node containing both non-null
`choiceSettingValue` and `simpleSettingValue`. Previously, a failed choice lookup
could fall back to the simple value; a valid choice could also silently ignore
the contradictory simple value. Unknown mixed parents could emit resolved children.
The repaired decoder retains unresolved metadata and stops descending through the
mixed node. No scalar, ADMX value or raw payload is exported for that node.
Other independent settings and supported nested groups continue to decode.

Microsoft documents separate [choice](https://learn.microsoft.com/en-us/graph/api/resources/intune-deviceconfigv2-devicemanagementconfigurationchoicesettinginstance?view=graph-rest-beta)
and [simple](https://learn.microsoft.com/en-us/graph/api/resources/intune-deviceconfigv2-devicemanagementconfigurationsimplesettinginstance?view=graph-rest-beta)
instance shapes. This is a bounded rejection of their unsupported coexistence,
not a full type/collection/enum audit. No live Graph payload is claimed to have
triggered the bug. Tests use injected dictionary and JSON-decoded objects.

Recollect affected policy evidence: an older export that already labels a scalar
Resolved cannot reconstruct the original discarded ambiguity. Assay rules remain
1.8.0-preview, with no count/score changes or automatic snapshot migration. New
unresolved facts cannot establish RTP-01/BM-01 references. The offline production
test can write `ASSAY_INTUNE_MIXED_SETTING_FIXTURE` for native/app regressions,
including persistence and reports. No new route, permission or provider.

## Run

```powershell
.\Invoke-IntuneDiscovery.ps1 -TenantId '<tenant-guid>' -OutputPath '.\intune-discovery.json'
```

Requires Az.Accounts for interactive/existing-connection authentication, or an
approved delegated Graph token provided as SecureString through
`-GraphAccessToken $SecureToken`. Never include token literals in scripts,
command history, logs, or chat. No automatic module installation, consent,
app registration, or role assignment.

`-UseExistingConnection` verifies the selected commercial-cloud tenant.
Authentication uses Microsoft's
[Get-AzAccessToken](https://learn.microsoft.com/en-us/powershell/module/az.accounts/get-azaccesstoken)
and [Connect-AzAccount](https://learn.microsoft.com/en-us/powershell/module/az.accounts/connect-azaccount).
An Az token may not contain the required delegated read scopes. If missing,
use an administrator-approved reader application's token; no write-permission
fallback. Decoded token metadata checks tenant/audience/expiry/scopes, not
authenticity; Graph validates the credential. Application-only/opaque tokens
and noncommercial clouds are not supported by this collector release.

| Module | Delegated read scope | Microsoft reference |
| --- | --- | --- |
| Devices | `DeviceManagementManagedDevices.Read.All` | [List managedDevices](https://learn.microsoft.com/en-us/graph/api/intune-devices-manageddevice-list?view=graph-rest-1.0) |
| Compliance definitions/assignments/reports | `DeviceManagementConfiguration.Read.All` | [Policies](https://learn.microsoft.com/en-us/graph/api/intune-deviceconfig-devicecompliancepolicy-list?view=graph-rest-1.0), [assignments](https://learn.microsoft.com/en-us/graph/api/intune-deviceconfig-devicecompliancepolicyassignment-list?view=graph-rest-1.0), [reports](https://learn.microsoft.com/en-us/graph/api/intune-deviceconfig-devicecompliancedevicestatus-list?view=graph-rest-1.0) |
| Legacy configuration | `DeviceManagementConfiguration.Read.All` | [Definitions](https://learn.microsoft.com/en-us/graph/api/intune-deviceconfig-deviceconfiguration-list?view=graph-rest-1.0), [assignments](https://learn.microsoft.com/en-us/graph/api/intune-deviceconfig-deviceconfigurationassignment-list?view=graph-rest-1.0) |
| Apps | `DeviceManagementConfiguration.Read.All` | [Apps](https://learn.microsoft.com/en-us/graph/api/intune-apps-mobileapp-list?view=graph-rest-1.0), [assignments](https://learn.microsoft.com/en-us/graph/api/intune-apps-mobileappassignment-list?view=graph-rest-1.0) |
| Optional `-IncludeRbac` | `DeviceManagementRBAC.Read.All` | [Roles](https://learn.microsoft.com/en-us/graph/api/intune-rbac-roledefinition-list?view=graph-rest-1.0), [assignments](https://learn.microsoft.com/en-us/graph/api/intune-rbac-roleassignment-list?view=graph-rest-1.0) |
| Optional `-IncludeAudit` | `DeviceManagementApps.Read.All` | [Audit](https://learn.microsoft.com/en-us/graph/api/intune-auditing-auditevent-list?view=graph-rest-1.0) |

These scopes do not establish actual RBAC visibility or licensing. Intune API
license prerequisites are stated on the linked endpoint references.

Optional assessment inputs: `-Assessor`, `-ScopeDescription`,
`-MaxCollectionAgeHours`, `-MaxPolicyReportAgeHours`, `-MaxDeviceSyncAgeDays`.
These are customer requirements, not Microsoft defaults. Alternatively set them
in Assay's **Intune evidence** dialog after importing. Audit collection accepts
`-AuditSinceUtc`; default seven-day lookback is only a collection convenience.

## Evidence boundaries

- Explicit tenant, collection ID, UTC interval, API version, module state and child-report completion.
- Full [paging](https://learn.microsoft.com/en-us/graph/paging), including empty pages with next links; bounded retries using [Retry-After](https://learn.microsoft.com/en-us/graph/throttling).
- GET-only endpoint allowlist, HTTPS, no redirects, four attempts, 120-second request timeout, 30-minute run budget, 16 MB response and 64 MB export limits. These are collector safeguards, not Microsoft service limits.
- Partial and empty states are retained; collection completeness does not establish unrestricted tenant visibility or compliance.
- Allowlisted fields exclude password/recovery values, scripts, app content, actor identifiers, and audit modified-property values. Names and IDs remain confidential organizational data.
- Existing output is not overwritten. Use a writable local folder; do not weaken endpoint protection to bypass a blocked write.
- Definitions/assignments are not effective targeting. Compliance-report IDs are entity IDs, not assumed device IDs; see the [report-row resource](https://learn.microsoft.com/en-us/graph/api/resources/intune-deviceconfig-devicecompliancedevicestatus?view=graph-rest-1.0).
- No legacy configuration-status binding: Microsoft documents deprecation starting May 2026 for that [entity](https://learn.microsoft.com/en-us/graph/api/resources/intune-deviceconfig-deviceconfigurationdevicestatus?view=graph-rest-1.0).
- No report export-job creation: its [reference](https://learn.microsoft.com/en-us/graph/api/intune-reporting-devicemanagementexportjob-create?view=graph-rest-1.0) lists write scopes.

Modern configuration, enrollment, Entra and recovery modules are opt-in as
described below, as are MAM, app configuration, threat connectors and platform
compliance. Other unsupported families remain NotRequested. Inventory is independent of Windows device counts. No group
expansion, device action, tenant mutation, targeted secret retrieval, cross-tenant
collection, or automatic token renewal after 401. External process termination
may prevent an export from being written.
Field projection limits exports, not the contents of full responses in memory;
see the audit's data-handling limits.

## Offline verification

```powershell
.\Test-IntuneDiscovery.ps1
.\Test-IntuneExpansion.ps1
.\Test-IntuneDefenderEvidence.ps1
.\Test-IntuneServices.ps1
.\Test-IntuneTunnel.ps1
.\Test-IntuneEpmRules.ps1
.\Test-IntuneMamLaunch.ps1
.\Test-IntuneCfaEvidence.ps1
.\Test-IntuneSignatureTimestamp.ps1
powershell.exe -NoProfile -File .\Test-IntuneDiscovery.ps1
powershell.exe -NoProfile -File .\Test-IntuneExpansion.ps1
powershell.exe -NoProfile -File .\Test-IntuneDefenderEvidence.ps1
powershell.exe -NoProfile -File .\Test-IntuneServices.ps1
powershell.exe -NoProfile -File .\Test-IntuneTunnel.ps1
powershell.exe -NoProfile -File .\Test-IntuneEpmRules.ps1
powershell.exe -NoProfile -File .\Test-IntuneMamLaunch.ps1
powershell.exe -NoProfile -File .\Test-IntuneCfaEvidence.ps1
powershell.exe -NoProfile -File .\Test-IntuneSignatureTimestamp.ps1
```

Tests load library-only functions and use synthetic HTTP responses, never a
tenant connection. `ASSAY_INTUNE_FIXTURE` optionally writes a synthetic export
for cross-language testing. A live pilot must reconcile authentication, API
responses and counts with portal evidence for the same visible scope before
production-validation claims.

## Configuration Expansion

### Signature Timestamp Integrity (0.5.5)

**Keep `IntuneEndpointTimestamps.ps1` beside the endpoint companion and discovery
scripts.** The companion now serializes AntivirusSignatureLastUpdated as an
explicit UTC ISO string before JSON export. UTC/offset dates and unambiguous Local
DateTime values are supported; Unspecified DateTime never inherits the importing
machine's timezone. This fixes the [.NET local-time assumption](https://learn.microsoft.com/en-us/dotnet/api/system.datetime.touniversaltime).

PowerShell 7's [automatic JSON date conversion](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/convertfrom-json?view=powershell-7.5)
can erase original syntax. The file-import adapter restores this field's raw
string using System.Text.Json before validating it; other fields are unchanged.
The reviewed Windows PowerShell 5.1 path preserves JSON strings. No 7.5-only
DateKind parameter or module installation is required.

ISO strings must include Z or a numeric offset and at most seven fractional
digits. Missing timezone, unknown-offset -00:00, legacy /Date(...), invalid values
or unreviewed types stay unavailable; no default is inferred. Adjacent Defender
fields survive. Rules 1.5.2-preview apply corresponding guards to direct native
imports. Historical lost timezone information requires recollection, not guessing.
Valid syntax is not clock accuracy, signature freshness or update-health proof.
That integrity repair added no fields, commands, permissions, findings or scores.
The separate signature-age comparison below now uses this evidence.
Test-IntuneSignatureTimestamp uses injected values only;
ASSAY_INTUNE_SIGNATURE_TIME_FIXTURE optionally writes native-test cases.

### Reported Signature-Update Age (Assay 1.6.0-preview)

Assay N-04-AGE uses the already collected `AntivirusSignatureLastUpdated` from
[Get-MpComputerStatus](https://learn.microsoft.com/en-us/powershell/module/defender/get-mpcomputerstatus?view=windowsserver2025-ps).
Age is measured at the assessment-as-of time. It defaults to Observed / Evidence;
the optional Reference input **Signature-update age limit (hours)** supplies
`AssessmentRequirements.ConfigurationReview.MaxSignatureAgeHours`. Accepted values
are integers 0-87600; absent/null means no target. These are customer/Assay inputs,
not a Microsoft freshness SLA. At or below the inclusive limit is Pass, above it
Warning, using full timestamp precision. Clearing the input restores observation.

The selected Windows device must uniquely match one fresh, complete endpoint
sample and one Complete DefenderStatus row. All four signature/sample/completion/
assessment times require supported explicit instants and valid ordering. Missing,
unknown-offset, future, stale or incomplete evidence cannot establish a clean
Pass. `MaxCollectionAgeHours` remains an independent sample-age gate. Preferences
or active antivirus are not prerequisites for a timestamp-age comparison.
This is not proof of latest security intelligence, current device state, update
delivery, active protection or platform/engine currency. Clocks/identity are not
authenticated; old timestamp provenance cannot be recovered.

This age-only increment required no collector change: existing requirement pass-through retains
the limit without a new field, command, permission or update action. Offline
production-import tests in PowerShell 5.1/7 preserve 0, 2 and 87600; native and
application tests cover boundaries, target changes, persistence and reports.

### File-Share Source State (0.5.7)

The existing preference read now also projects
`SignatureDefinitionUpdateFileSharesSources` in memory, removes the raw property
before JSON, and emits Assay-derived `SignatureFileSharesState`. Empty means a
typed empty string, NonEmpty a non-whitespace string, and Unknown all missing,
null, whitespace-only or unsupported types. It is independent of
SharedSignaturesPathState; both raw properties are removed before export.

The [PowerShell reference](https://learn.microsoft.com/en-us/powershell/module/defender/set-mppreference?view=windowsserver2025-ps#-signaturedefinitionupdatefilesharessources)
defines the exact plural property and a pipe-separated UNC string. The
[product guide](https://learn.microsoft.com/en-us/defender-endpoint/manage-protection-updates-microsoft-defender-antivirus)
documents skipping FileShares when no paths are entered. Its shorter singular
command example is not used to infer a provider alias.

Assay N-03 (rules 1.6.2-preview) describes FileShares listed with an empty reported
source string, but remains Observed / Evidence. It never claims that runtime
skipping occurred. NonEmpty is not a valid-location, access, content or delivery
check; even an unparsed pipe-only string is NonEmpty. FileShares not listed does
not create a requirement to add it. Shared-signature override context remains
separate, with no effective-source or protection conclusion.

Only exact derived states survive import; raw paths and malformed states are
dropped. Older/missing evidence stays Unknown. The existing platform, selection,
identity, freshness and complete-provider gates apply. Five provider commands
now project 38 fields: 36 direct exports and two derived states. No new provider,
route, scope, finding or score; no share/network probes or update operation.
Offline tests cover both row forms, independent reduction, malformed/missing
values, forged states, privacy, save/load and reports. Live validation remains open.

### Shared-Signature State (0.5.6)

The existing Get-MpPreference read adds `SharedSignaturesPath` in memory. Before
JSON serialization, the companion removes it and emits the Assay-derived
`SharedSignaturesPathState`: `Empty` for a typed empty string, `NonEmpty` for a
non-whitespace string, otherwise `Unknown`. Missing/null/whitespace/non-string
values never imply an empty setting. No raw path is exported or accessed; no UNC
syntax, path validity, applicability or runtime-use check is performed. The
provider object can contain the raw value in memory before reduction.

Assay N-03 (rules 1.6.1-preview) reports this state alongside the source list,
with the documented [SharedSignaturesPath override](https://learn.microsoft.com/en-us/powershell/module/defender/set-mppreference?view=windowsserver2025-ps#-sharedsignaturespath)
(revision `4a7aa7a27f92a5df33b61c78efc584abf2aad4a9`). Microsoft's
[VDI guidance](https://learn.microsoft.com/en-us/defender-endpoint/deployment-vdi-microsoft-defender-antivirus)
describes the feature and separate share/access prerequisites. These are source
references; no Set cmdlet, share access or update operation is executed.

The derived name/enum is an Assay contract, not a Microsoft provider field.
Import accepts only exact Empty/NonEmpty/Unknown values and drops raw paths.
Older or malformed state evidence stays Unknown, not a guessed empty setting.
Valid source lists remain Observed; malformed lists remain NotAssessed while
retaining supported state evidence. Existing identity/platform/freshness/provider
gates apply; no state yields a new health verdict or establishes effective source
selection, network reachability, update delivery or protection.

The 0.5.6 increment projected 37 source fields, exporting 36 directly and
one as a derived state; 0.5.7 adds the second state above. Test-IntuneEndpointContracts distinguishes read/export
names. The production pipeline test covers dictionary/object rows, missing and
malformed inputs, forged states and raw-path privacy; native/app tests cover
reassessment, save/load and reports. No new provider, permission, route, finding
or score. Tests mock providers and HTTP; live behavior remains unverified.

### Defender Update Cadence (0.5.4)

The existing endpoint preference read adds `SignatureScheduleDay` and
`SignatureUpdateInterval`. Assay N-04 (rules 1.5.1-preview) requires both fields
and records day 0-8 (or documented names Everyday/Sunday through Saturday/Never)
and integer interval 0-24 hours. Missing or invalid values stay NotAssessed.
Per Microsoft's [scheduling guide](https://learn.microsoft.com/en-us/defender-endpoint/manage-protection-update-schedule-microsoft-defender-antivirus),
day 8 plus interval 0 means no Defender-owned schedule. Interval 0 alone does not
mean all updates are disabled: a day can still be specified. Both day and interval
present is reported without guessing precedence. All valid combinations yield
Observed, never a health or compliance verdict.

Scheduled local clock time, randomization, runtime mode, actual execution and
signature freshness are not evaluated. Other update mechanisms and passive-mode
limitations need separate review. No new provider, Graph route, scope, scheduled
task or update action is introduced. Existing production pipeline tests cover
zero/missing/named values and exclude SignatureScheduleTime; native and report
tests use the generated synthetic endpoint fixture. No live endpoint query.

### Defender Update Source Evidence (0.5.3)

The existing endpoint `Get-MpPreference` projection now includes
`SignatureFallbackOrder`. Assay N-03 (rules 1.5.0-preview) retains the order of the
four documented tokens: InternalDefinitionUpdateServer, MicrosoftUpdateServer,
MMPC and FileShares. One to four unique pipe-separated tokens are supported,
including ASCII spaces around tokens; missing, malformed or unknown values stay
unassessed without discarding tokens or inventing defaults.

This is Observed/NotAssessed evidence only. It identifies MMPC's listed position
against [Microsoft's final-fallback guidance](https://learn.microsoft.com/en-us/defender-endpoint/manage-protection-updates-microsoft-defender-antivirus),
not a health verdict. File-share locations/access are not collected. The
[PowerShell reference](https://learn.microsoft.com/en-us/powershell/module/defender/set-mppreference?view=windowsserver2025-ps#-signaturefallbackorder)
also documents SharedSignaturesPath overriding fallback-order updates; 0.5.6 adds
only the path-free value state above. Never treat the list as proof of the effective source,
successful updates, WSUS approvals, signature freshness or platform/engine servicing.
No new provider, Graph request, scope, update action or path export is added.

`Test-IntuneCfaEvidence.ps1` now tests this field through the actual provider
callback and companion/discovery import with mocked providers/HTTP. Its existing
synthetic fixture includes both CFA and source-order observations. N-03 requires
the selected Windows identity, fresh sample and complete preference module, but
does not require Defender runtime status to report configuration evidence.

### Controlled Folder Access Evidence (0.5.2)

The optional endpoint companion now retains `EnableControlledFolderAccess` from
its existing `Get-MpPreference` call. No additional provider, Graph request or scope
is added. Import the companion file using `-EndpointEvidencePaths`; select managed
device IDs and confirm scope/freshness in the existing Assay evidence workflow.
The cloud collector does not remotely execute the endpoint companion.

Assay N-01 (rules 1.4.1-preview) distinguishes all five [CFA modes](https://learn.microsoft.com/en-us/defender-endpoint/controlled-folder-access-overview)
and reports active-AV/real-time prerequisites separately. The [PowerShell reference](https://learn.microsoft.com/en-us/defender-endpoint/controlled-folder-access-configure#enable-and-configure-cfa-in-powershell)
documents the exact field and integer/named values. [Defender compatibility](https://learn.microsoft.com/en-us/defender-endpoint/microsoft-defender-antivirus-compatibility)
explains why passive/EDR-block status is not active CFA protection.

By default this is unscored Observed/NotAssessed evidence. Assay's Reference tab
now offers **CFA mode target**, default Observe only. The optional exact
`ConfigurationReview.ControlledFolderAccessTarget` accepts Disabled, Enabled,
AuditMode, BlockDiskModificationOnly or AuditDiskModificationOnly. A matching
non-disabled target requires observed active AV/real-time prerequisites for Pass;
known mismatch or unmet prerequisites yields Warning. Matching Disabled does not
require active CFA, but still requires known evidence. Unknown evidence stays
unassessed; no fleet/protection verdict is implied. Clearing the target restores
observation-only behavior. Targets apply uniformly to the selected samples.

Audit and Disabled are not universal policy violations; disk-only modes do not
protect files in folders. No folder/app lists, effective
assignment or actual block event is evaluated. Missing/unknown fields and stale,
ambiguous or incomplete selected-device evidence stay unassessed. Old samples
without this field do not acquire an inferred default. Provider projection does
not mean unrelated data was never present in the provider's in-memory object.

Run `Test-IntuneCfaEvidence.ps1` in 5.1/7 for mocked production provider/import
checks. `ASSAY_INTUNE_CFA_FIXTURE` optionally writes a synthetic discovery export
for native reassessment/save/load/report tests. No provider or Graph query runs
in the test. Live provider serialization and endpoint enforcement remain unverified.

### Configuration Collection

```powershell
.\Invoke-IntuneDiscovery.ps1 -TenantId '<tenant-guid>' `
	-OutputPath '.\intune-expanded.json' `
	-IncludeConfiguration -IncludeRbac -IncludeEnrollment
```

- `-IncludeConfiguration`: beta modern policy/settings/definitions,
	legacy intent/template/assignment metadata, filters and v1 noncompliance
	schedules. Uses existing Configuration.Read.All. Legacy setting bodies and
	unrecognized modern definitions are not decoded.
	Modern assignment reads are currently quarantined as Unsupported.
- `-IncludeRbac` also collects scope tags when configuration is requested.
- `-IncludeEnrollment`: beta ESP metadata/assignments and Autopilot metadata; requires
	DeviceManagementServiceConfig.Read.All.
	Autopilot assignment reads are currently quarantined as Unsupported.
- `-IncludeEntra`: device registration and Conditional Access; requires
	Policy.Read.DeviceConfiguration and Policy.Read.All plus a supported Entra role.
- `-IncludeRecoveryMetadata`: metadata-only lists with
	DeviceLocalCredential.ReadBasic.All and BitlockerKey.ReadBasic.All. No passwords
	or recovery keys are requested or exported.
- `-IncludeApple`: APNs/VPP metadata projections with ServiceConfig.Read.All.
	Projection and response compatibility still need live validation.

The [settings API](https://learn.microsoft.com/en-us/graph/api/intune-deviceconfigv2-devicemanagementconfigurationsetting-list?view=graph-rest-beta)
and [definition API](https://learn.microsoft.com/en-us/graph/api/intune-deviceconfigv2-devicemanagementconfigurationchoicesettingdefinition-list?view=graph-rest-beta)
are interpreted by exact definition/CSP/option identities, not display names
or numeric suffixes. Unknown setting values are omitted.

Run the companion independently on an explicitly authorized endpoint:

```powershell
.\Get-IntuneEndpointEvidence.ps1 -TenantId '<tenant-guid>' -OutputPath '.\endpoint.json'
```

It reads selected Defender, ActiveStore firewall, BitLocker-volume and DeviceGuard
fields after checking local device/tenant identity. No automatic elevation or
remediation. Attach samples using `-EndpointEvidencePaths '.\endpoint.json'` in
the discovery command. The tenant collector never remotely runs the companion.

Assay exposes 104 supplementary unscored findings: by default 63 bounded comparison
entries and 41 evidence entries (64/40 with either an explicit CFA target or a
signature-age limit; 65/39 with both). All have
handlers, but several cover only part of their
feature; evidence collection is not complete automatic assessment. The original
50-control scoring catalog remains unchanged. Set reference/scenario scope in
the native evidence dialog. `ASSAY_INTUNE_EXPANSION_FIXTURE` optionally writes a
synthetic expanded snapshot, including companion imports, from the offline test.

## Selected Real-Time Policy Correlation

Assay rules 1.7.0-preview add RTP-01 without changing collector production 0.5.7.
Use the existing `-IncludeConfiguration` and `-EndpointEvidencePaths` workflow,
then select exactly one modern Windows policy in Reference `PolicyIds` and a
device cohort in `DeviceIds`. The policy is an analyst reference, not a proven
assignment. Complete supported collections and selected-parent settings coverage
must expose exactly one Resolved integer 0/1
[AllowRealtimeMonitoring](https://learn.microsoft.com/en-us/windows/client-management/mdm/policy-csp-defender#allowrealtimemonitoring).

RTP-01 compares that value with the inverse Boolean DisableRealtimeMonitoring
preference and independently with reported Boolean RealTimeProtectionEnabled.
Both fields are already collected. Each comparison is Aligned or Different;
the finding stays Observed / Evidence, never a protection or deployment verdict.
Missing/ambiguous/partial evidence, malformed values and stale/future samples stay
NotAssessed. A disabled reference can align without being approved as secure.
Assignment filters/groups, applicability, precedence, intervening changes, tamper
protection and runtime health are not resolved. Snapshots are not atomic.

No new route, provider, field, permission or action. Test-IntuneCfaEvidence can
write `ASSAY_INTUNE_REALTIME_FIXTURE` from the actual mocked provider callback,
definition decoder and companion import. It verifies opaque option-value joins,
typed fields and unavailable assignment coverage on PowerShell 5.1/7. Native
tests cover reference switching, unknown coverage, persistence and reports.

### Behavior Monitoring (Assay 1.8.0-preview)

BM-01 independently correlates the selected policy's exact integer 0/1
[AllowBehaviorMonitoring](https://learn.microsoft.com/en-us/windows/client-management/mdm/policy-csp-defender#allowbehaviormonitoring)
with the inverse Boolean DisableBehaviorMonitoring preference and reported Boolean
BehaviorMonitorEnabled. All are already collected; production remains 0.5.7.
The same single-policy, complete-settings/parent-coverage, unique Windows endpoint,
typed-provider and precise timestamp gates apply as for RTP-01. Neither finding
borrows the other's setting or endpoint fields. Missing data stays NotAssessed;
Aligned/Different stay Observed / Evidence, not security or deployment verdicts.

The existing ASSAY_INTUNE_REALTIME_FIXTURE now includes two settings per policy
and opposite real-time/behavior values to detect cross-wiring. Mocked production
definition joins and endpoint imports are tested in PowerShell 5.1/7; native
tests cover both mappings' gates, switching, persistence and HTML/PDF exports.
No new field, provider, route, permission or action. Real-time and behavior
monitoring runtime dependencies, applicability, assignment and enforcement are
not established by these independent evidence labels.

## MAM, Remote Help And Platform Services

```powershell
.\Invoke-IntuneDiscovery.ps1 -TenantId '<tenant-guid>' `
		-OutputPath '.\intune-services.json' `
		-IncludeMam -IncludeMamLaunch -IncludeRemoteHelp -IncludeConnectors `
		-IncludeAppConfiguration -IncludePlatformCompliance `
		-IncludeConfiguration -IncludeRbac -IncludeEntra -IncludeTunnel
```

- `IncludeMam`: v1 Android/iOS APP policies, assignments and apps; separate beta
	Windows APP policies. Additional `DeviceManagementApps.Read.All`. Collection
	runs even with zero enrolled devices. No MAM registration/user inventory.
- `IncludeRemoteHelp`: beta tenant singleton, state, unenrolled permission and
	chat flag. Existing Configuration.Read.All is a documented read alternative.
	No remote sessions, chat contents or device actions.
- `IncludeMamLaunch`: independent beta Android/iOS APP lists with reviewed
	root/action, PIN retry, offline/biometric timers, attestation, notification,
	clipboard, threat/priority and platform applicability fields. Apps.Read.All
	is already requested by IncludeMam; the switches can also run independently.
	No user registrations, app payloads or secrets; v1 APP evidence is not merged.
- `IncludeConnectors`: v1 mobile threat connectors, heartbeat/state, distinct
	MAM/MDM flags and privacy controls; ServiceConfig.Read.All.
- `IncludeAppConfiguration`: v1 managed-app and managed-device app configuration
	metadata/assignments. Apps.Read.All. Exact metadata-only select is required;
	arbitrary customSettings, encoded XML and secrets are excluded.
- `IncludePlatformCompliance`: separate beta Android device-owner compliance
	and modern compliance metadata (including Linux); existing Configuration read
	scope. The original v1 scoring/report module is unchanged.
- Existing `IncludeConfiguration` also decodes four exact EPM client-setting
	ID/path pairs and grouped rule name/file/path/type definitions using returned
	typed option values. Rule groups are isolated; only automatic-wildcard and
	empty/network-path checks are added. Hash/certificate/child-process safety and
	effective client settings remain unassessed.
- `IncludeTunnel`: beta sites and per-site server lists using existing
	Configuration.Read.All. Health/check-in and upgrade metadata only, with
	per-site completion. No probes, log actions, upgrades or new permissions.
	Public/probe URLs are omitted. Select Tunnel sites and a customer check-in-age
	target in Assay; a healthy captured state is not proof of current connectivity.

In Assay's Reference tab, choose a separate MAM framework level and APP policy
IDs, nullable Remote Help targets and platform security requirements. These are
bounded configuration comparisons, not proof of user/device enforcement.
`ASSAY_INTUNE_SERVICES_FIXTURE` optionally exports a synthetic service fixture
from the new offline test.

Microsoft references: [MAM framework](https://learn.microsoft.com/en-us/intune/app-management/protection/data-protection-framework),
[APP list](https://learn.microsoft.com/en-us/graph/api/intune-mam-managedapppolicy-list?view=graph-rest-1.0),
[Remote Help GET](https://learn.microsoft.com/en-us/graph/api/intune-remoteassistance-remoteassistancesettings-get?view=graph-rest-beta),
[threat connectors](https://learn.microsoft.com/en-us/graph/api/intune-onboarding-mobilethreatdefenseconnector-list?view=graph-rest-1.0),
[managed app configuration](https://learn.microsoft.com/en-us/graph/api/intune-mam-targetedmanagedappconfiguration-list?view=graph-rest-1.0),
[device app configuration](https://learn.microsoft.com/en-us/graph/api/intune-apps-manageddevicemobileappconfiguration-list?view=graph-rest-1.0),
[Android device-owner](https://learn.microsoft.com/en-us/graph/api/intune-deviceconfig-androiddeviceownercompliancepolicy-list?view=graph-rest-beta),
[modern compliance](https://learn.microsoft.com/en-us/graph/api/intune-deviceconfigv2-devicemanagementcompliancepolicy-list?view=graph-rest-beta),
[EPM settings](https://learn.microsoft.com/en-us/intune/epm/manage-elevation-settings).

Tunnel references: [sites](https://learn.microsoft.com/en-us/graph/api/intune-mstunnel-microsofttunnelsite-list?view=graph-rest-beta),
[servers](https://learn.microsoft.com/en-us/graph/api/intune-mstunnel-microsofttunnelserver-list?view=graph-rest-beta),
[health resource](https://learn.microsoft.com/en-us/graph/api/resources/intune-mstunnel-microsofttunnelserver?view=graph-rest-beta).
EPM rule recommendations: [rule creation](https://learn.microsoft.com/en-us/intune/epm/create-elevation-rules),
[planning](https://learn.microsoft.com/en-us/intune/epm/deployment-planning).
Exact grouped definition binding: [Microsoft365DSC test fixture](https://github.com/microsoft/Microsoft365DSC/blob/Dev/Tests/Unit/Microsoft365DSC/Microsoft365DSC.IntuneEpmElevationRulesPolicyWindows10.Tests.ps1),
reviewed 2026-09-17. A public fixture is not live-tenant validation.
`ASSAY_INTUNE_TUNNEL_FIXTURE` and `ASSAY_INTUNE_EPM_FIXTURE` enable synthetic
exports from their offline suites for cross-language/native testing.

## MAM Launch Review

For launch checks, select a MAM framework level and **Selected APP launch
policies** in Assay. Android/iOS IDs are qualified independently. Graph PIN
retry actions are not assumed equivalent to the UI's Reset PIN action. Timers
accept positive whole-component day/hour/minute/second forms; unsupported forms
remain Not Assessed. Platform/SDK/app enforcement still needs pilot evidence.
`ASSAY_INTUNE_MAM_LAUNCH_FIXTURE` enables the synthetic launch export in its test.

MAM launch references: [Android list](https://learn.microsoft.com/en-us/graph/api/intune-mam-androidmanagedappprotection-list?view=graph-rest-beta),
[iOS list](https://learn.microsoft.com/en-us/graph/api/intune-mam-iosmanagedappprotection-list?view=graph-rest-beta),
[Android settings](https://learn.microsoft.com/en-us/intune/app-management/protection/ref-settings-android),
[iOS settings](https://learn.microsoft.com/en-us/intune/app-management/protection/ref-settings-ios).
The previous count-based MTD conflict claim is withdrawn: current APP guidance
supports primary-partner selection with multiple connectors; effective priority
and client behavior require separate evidence.

## Defender And App Control Evidence

The Defender companion uses a separate delegated service token supplied locally
as SecureString. It performs no login, consent, tenant changes or device actions.

```powershell
.\Get-IntuneDefenderEvidence.ps1 -TenantId '<tenant-guid>' `
	-AccessToken $SecureDefenderToken -OutputPath '.\defender-evidence.json'
.\Invoke-IntuneDiscovery.ps1 -TenantId '<tenant-guid>' `
	-OutputPath '.\intune-combined.json' -IncludeConfiguration `
	-DefenderEvidencePath '.\defender-evidence.json' `
	-AppControlPolicyPaths '.\reviewed-policy.xml'
```

Obtain `$SecureDefenderToken` through an approved local authentication flow;
never paste token literals into chat or command history. The supported audience
is Defender, not Graph, with delegated `Machine.Read`. Opaque/application-only
tokens and unrecognized audience forms remain unsupported. See Microsoft's
[machine list](https://learn.microsoft.com/en-us/defender-endpoint/api/get-machines)
and [machine resource](https://learn.microsoft.com/en-us/defender-endpoint/api/machine).
Device-group visibility and retention apply. `lastSeen` is the last full report,
normally daily, not a portal UI heartbeat. Assay requires selected device IDs
and an explicit **Defender full-report age target (hours)** before comparison.

Defender reads stay on the exact HTTPS machine-list path, with redirects off,
bounded paging/retries, Retry-After support and a 30-minute budget. Non-200
responses, including 404, do not establish clean absence. Exports retain IDs,
onboarding/health state, platform metadata and timestamps, not IP/user inventory.

App Control XML files are explicit source evidence, not a live policy query.
The parser rejects DTDs/external entities and files over 1 MB, retaining only
policy/base IDs, type, selected options and Audit/Managed Installer trust flags.
No raw XML, signers or file rules are exported. Source mode does not prove policy
assignment, runtime enforcement or Managed Installer tagging. See
[Microsoft's App Control guidance](https://learn.microsoft.com/en-us/intune/device-configuration/endpoint-security/manage-app-control).

Reviewed ADMX policy payloads retain enablement and exact typed Boolean/enum elements;
unsupported/default/unresolved payloads do not become successful comparisons.
The IE COM-launch policy accepts the reviewed enable/disable-only payload. The
three BitLocker startup/recovery policies require all documented data IDs with
their exact Boolean or numeric enum types; unknown IDs or values remain unresolved.
`Test-IntuneAdmxContracts.ps1` tests these bindings without tenant calls. Add
`-EvidenceDirectory <directory>` containing the CSP source cache from
`Test-IntunePolicyContracts.ps1` to verify the published sample IDs/types/enums.
Set `ASSAY_INTUNE_ADMX_FIXTURE` to write synthetic production-decoded evidence
for native import/reassessment tests. See [AUDIT.md](AUDIT.md) for source details.
Legacy setting bodies, per-rule exclusion coverage, effective targeting, full
BitLocker/LAPS prerequisites and runtime installer pairing still need further
evidence. All verification here is synthetic, not a live-tenant validation claim.