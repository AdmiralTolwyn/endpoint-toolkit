# Intune Collector Contract Audit

Date: 2026-09-17. Status: **partial audit, not production certification**.

September 19, collector 0.5.12: repaired CSP path coercion in
ConvertTo-IntuneSettingFacts. Base/offset arrays could become valid strings;
generic binding now requires typed strings before unchanged normalization and
allowlist checks. Get-IntuneEpmPath also rejects non-string IDs/offsets and
non-string supplied bases while preserving its absent/null/string-base support.
That guard covers four sample-bound client-setting mappings, not the separate
EPM elevation-rule decoder or a new Microsoft CSP authority.

The [setting-definition contract](https://learn.microsoft.com/en-us/graph/api/resources/intune-deviceconfigv2-devicemanagementconfigurationsettingdefinition?view=graph-rest-beta)
declares baseUri/offsetUri as String, revision
`e8eace7fa9bfa9e7afacfe9f5dec1ba42ffd493f`. Tests include 24 malformed generic
path cases, four string controls, an independent child and 56 EPM mapping cases
across dictionary/JSON forms. Production mocked exports stay unsupported through
native sanitization, persistence and reports. No new provider, scope, route or
score; no live occurrence, applicability or full URI/schema validation claimed.
Recollect affected older exports whose original path types were discarded. See
[typed-path limits](README.md#typed-csp-paths-0512).

September 19, collector 0.5.11: reproduced and repaired unresolved singular or
collection choices emitting Resolved children. ConvertTo-IntuneSettingFacts now
propagates choice-resolution uncertainty separately from template defaults.
Get-IntuneSelectedOptionValue requires one exact typed definition/option match
and an object-valued optionValue. Missing/duplicate matches or absent/scalar/array
payloads withhold descendant values, retaining known metadata. A valid descendant
cannot clear inherited uncertainty. Collection siblings and selected-option
template context are evaluated independently; ordinary groups remain supported.

Sources: [choice value](https://learn.microsoft.com/en-us/graph/api/resources/intune-deviceconfigv2-devicemanagementconfigurationchoicesettingvalue?view=graph-rest-beta)
and [choice collection](https://learn.microsoft.com/en-us/graph/api/resources/intune-deviceconfigv2-devicemanagementconfigurationchoicesettingcollectioninstance?view=graph-rest-beta),
revision `4b837f772f711c890ec02678db6b05845a40b419`. They establish option identity
and child structure, not the conservative withholding rule or actual applicability.
Tests cover 40 choice matrix cases plus transitive/ADMX/template controls and
actual mocked exports. Native/app regression covers unknown results, persistence
and reports. No live occurrence, new read/scope or full dependency audit claimed.
Recollect affected older flattened evidence; see [limits](README.md#choice-context-0511).

September 19, collector 0.5.10: repaired a reproduced defaulted-parent -> Resolved
child path in ConvertTo-IntuneSettingFacts. Unresolved template context now flows
through descendants, including selected options on unsupported CSP paths and
group/choice-collection values. A false child flag cannot clear ancestor context;
explicit siblings remain independent. Child metadata is retained, values/ADMX
are withheld. Absent/null references still permit explicit values. Present
references require a Boolean useTemplateDefault; true, missing/null/untyped flags
or malformed reference objects remain unresolved, not truthiness-decoded.

[Template reference](https://learn.microsoft.com/en-us/graph/api/resources/intune-deviceconfigv2-devicemanagementconfigurationsettingvaluetemplatereference?view=graph-rest-beta)
and [group value](https://learn.microsoft.com/en-us/graph/api/resources/intune-deviceconfigv2-devicemanagementconfigurationgroupsettingvalue?view=graph-rest-beta)
contracts reviewed at revision `4b837f772f711c890ec02678db6b05845a40b419`. They
document the Boolean and structure, not the collector's conservative withholding
rule or actual child runtime behavior. Matrix tests cover 144 cases across value
kinds/representations, plus transitive/ADMX controls and actual mocked exports.
Native/app tests retain unknowns through save/load/reports; native normalization
uses UnresolvedValue for the collector's UnresolvedTemplateDefault marker.

No new source field, provider, scope, route, finding, score or native rule. Full
template/dependency/OData semantics remain open; no live occurrence is asserted.
Recollect affected older flattened policy evidence rather than guessing lost
context. See [template-context limits](README.md#template-context-0510).

September 19, collector 0.5.9: repaired type coercion in the settings decoder's
definition/option joins. Instance settingDefinitionId and choice value must be
nonblank strings; definition id and option itemId must be strings matched by
ordinal equality with exactly one match. Numeric-looking strings are retained,
never parsed; numeric/Boolean/array/null identities cannot manufacture a reference.
Malformed instance/choice nodes retain safe unresolved metadata and do not emit
descendants. Invalid candidate IDs are not matches. Duplicate/mixed-node and
valid nesting controls remain in the regression suite.

Sources: [setting definition](https://learn.microsoft.com/en-us/graph/api/resources/intune-deviceconfigv2-devicemanagementconfigurationsettingdefinition?view=graph-rest-beta)
revision `e8eace7fa9bfa9e7afacfe9f5dec1ba42ffd493f`;
[choice value](https://learn.microsoft.com/en-us/graph/api/resources/intune-deviceconfigv2-devicemanagementconfigurationchoicesettingvalue?view=graph-rest-beta)
and [option definition](https://learn.microsoft.com/en-us/graph/api/resources/intune-deviceconfigv2-devicemanagementconfigurationoptiondefinition?view=graph-rest-beta),
both revision `4b837f772f711c890ec02678db6b05845a40b419`. They establish string types;
nonblank/ordinal comparison is the adapter boundary. This is not a complete audit
of collection kinds, OData types, templates, dependencies or other decoders.
Synthetic production exports remain unknown through native import/reassessment/
save/load/reports; no live Graph occurrence or new permission/read is claimed.
Recollect affected older exports because discarded identity types are unrecoverable.
See [repair limits](README.md#typed-setting-identities-059).

Collector 0.5.8 repairs two synthetic mixed-choice/simple reproductions:
unresolved choices could borrow a simple value, and unsupported mixed parents
could still emit resolved descendants. Both non-null value members now leave the
node unresolved and block child traversal. Known unique supported definitions
retain CSP metadata/UnresolvedValue, unsupported or duplicate definitions retain
UnsupportedDefinition, and neither emits scalar/ADMX payloads. Valid nested-group
children and independent settings remain supported.

Source contracts: [choice instance](https://learn.microsoft.com/en-us/graph/api/resources/intune-deviceconfigv2-devicemanagementconfigurationchoicesettinginstance?view=graph-rest-beta)
and [simple instance](https://learn.microsoft.com/en-us/graph/api/resources/intune-deviceconfigv2-devicemanagementconfigurationsimplesettinginstance?view=graph-rest-beta),
both revision `4b837f772f711c890ec02678db6b05845a40b419`. This does not certify
all OData discriminators, collection-value kinds, nested shapes or enums. Tests
cover dictionary/JSON forms, known/unknown/duplicate/ADMX parents and valid nesting.
Actual mocked discovery output stays unresolved through native/app import,
reassessment, save/load and reports. No live-tenant occurrence is asserted.

No new read field, route, scope, provider, finding or score; native rules remain
1.8.0-preview. Recollect affected older evidence because normalized Resolved
scalars no longer carry the ambiguous source shape. See
[repair limits](README.md#mixed-setting-decoder-repair-058).

Assay rules 1.8.0-preview add behavior-monitoring correlation (BM-01) with no
production collector change. Existing AllowBehaviorMonitoring values are decoded
independently of AllowRealtimeMonitoring; preference/runtime fields are already
projected. The mocked production fixture now uses two settings per policy and
opposite values to detect cross-wiring, retaining unavailable assignment coverage.
Native strict gates and application save/load/report tests do not establish
assignment, runtime dependencies or protection. No new provider, field or scope.
See [behavior-monitoring limits](README.md#behavior-monitoring-assay-180-preview).

Assay rules 1.7.0-preview add real-time policy/endpoint alignment as Evidence.
Collector production remains 0.5.7. The existing definition join resolves exact
AllowRealtimeMonitoring integers; endpoint preference/runtime Booleans already
exist. Test-IntuneCfaEvidence now generates ASSAY_INTUNE_REALTIME_FIXTURE through
production decoding and import, with all reads/HTTP mocked, ambiguous option
suffixes and deliberately unavailable assignment coverage. Native scope/type/
time gates and save/load/report tests do not establish assignment, propagation,
applicability, precedence or protection. No new route, permission or provider.
See [correlation limits](README.md#selected-real-time-policy-correlation).

Collector 0.5.7 / Assay rules 1.6.2-preview: N-03 also observes path-free
SignatureFileSharesState from SignatureDefinitionUpdateFileSharesSources. The
exact plural property is source-backed; Empty/NonEmpty/Unknown is Assay-derived.
Both path settings reduce independently before serialization, without accessing
or exporting locations. FileShares plus Empty is linked to Microsoft's documented
skip behavior, not asserted as an observed skipped or failed update. Five commands,
38 read fields -> 36 direct + two derived exports; no new scope, route or score.
Mocked tests cover independent settings, both row forms, malformed/missing data,
raw-path exclusion and application persistence/reports. No validation of UNC
entries, share access, contents, effective precedence or delivery. See
[file-share limits](README.md#file-share-source-state-057).

Collector 0.5.6 / Assay rules 1.6.1-preview: N-03 gains a path-free observation
of SharedSignaturesPath. Microsoft documents the property/override; the exported
SharedSignaturesPathState (Empty/NonEmpty/Unknown) is explicitly Assay-derived,
not a provider enum. The companion removes the raw path before JSON and never
accesses it. The register separates 37 read fields from 36 direct + one derived
export. Missing/null/unsupported values remain Unknown, and old imported paths
do not fabricate a state. Mocked production-pipeline tests cover both row forms,
state preservation, invalid values and privacy; native workflow tests cover
save/load/reports. No new command, scope, route, score or effective-update verdict.
Path validity, share access, applicability, live serialization and runtime
override behavior remain unverified. See [state limits](README.md#shared-signature-state-056).

Assay rules 1.6.0-preview add N-04-AGE using the existing reported signature time.
The optional customer MaxSignatureAgeHours is preserved by existing collector
0.5.5 requirement handling, with 0/2/87600 tested through the production import
on PowerShell 5.1/7. No new provider, field, route, scope or update operation.
Native age-at-assessment comparisons require qualified/ordered timestamps and
fresh, complete selected evidence; no target means observation only. The input
cap/status mapping is Assay behavior, not Microsoft freshness guidance. No claim
of latest-release currency, update delivery or protection. See
[age comparison limits](README.md#reported-signature-update-age-assay-160-preview).

September 18 timestamp repair (collector 0.5.5 / Assay rules 1.5.2-preview):
AntivirusSignatureLastUpdated no longer gains an assumed timezone from an
Unspecified DateTime. The required IntuneEndpointTimestamps companion normalizes
reviewed provider instants before JSON export and preserves raw signature strings
across PowerShell 7 parsing. Direct native imports reject unqualified/invalid
forms. The regression reproduced timezone guessing and legacy JSON coercion;
tests use injected dates, not Defender reads. No freshness verdict, new field,
permission or provider. See [timestamp limits](README.md#signature-timestamp-integrity-055).

September 18 cadence follow-up (collector 0.5.4 / Assay rules 1.5.1-preview):
SignatureScheduleDay and SignatureUpdateInterval join the existing preference
projection. The [specific scheduling guide](https://learn.microsoft.com/en-us/defender-endpoint/manage-protection-update-schedule-microsoft-defender-antivirus)
defines joint 8/0 semantics; N-04 reports configuration only. Missing values never
inherit defaults, and interval zero alone is not update failure. Clock time,
effective precedence, runtime behavior and signature freshness remain unverified.
No new command, route, scope, scheduled task or update action. Synthetic provider
and native workflow tests do not establish live provider serialization.

September 18 source-order follow-up (collector 0.5.3 / Assay rules 1.5.0-preview):
one additional Get-MpPreference field, SignatureFallbackOrder. The reviewed
[Microsoft update-source guidance](https://learn.microsoft.com/en-us/defender-endpoint/manage-protection-updates-microsoft-defender-antivirus)
and [PowerShell contract](https://learn.microsoft.com/en-us/powershell/module/defender/set-mppreference?view=windowsserver2025-ps#-signaturefallbackorder)
bind four source tokens and ordered pipe syntax. N-03 reports observations only;
effective sources, SharedSignaturesPath override, share locations, approval,
delivery and freshness remain unverified. Production callback/import and privacy
tests use mocked providers/HTTP. No new command, permission, route or update action.

September 18 CFA follow-up (collector 0.5.2): the endpoint companion adds only
`Get-MpPreference.EnableControlledFolderAccess`, using the documented [read and
mode contract](https://learn.microsoft.com/en-us/defender-endpoint/controlled-folder-access-configure).
Assay N-01 retains five-mode semantics and separate active-AV/real-time prerequisite
observations; no target-based score, folder/app scope or block outcome is inferred.
Assay rules 1.4.1-preview now allow an explicit customer mode target in the Reference
tab. This adds an unscored, prerequisite-aware comparison; Observe only remains
the default. The existing collector preserves supplied requirements without new
commands, routes or permissions. Target-preservation tests cover all five choices.
No new command, API or scope. Production callback/import tests mock all providers
and HTTP, including tenant rejection and privacy. Live serialization remains open.

## Verified Slice

`GraphContracts.json` registers 51 module/route contracts. The collector requires
a registered module/route pair before collection. The public Graph CSDL and
version-pinned Microsoft Learn pages establish the top-level projected fields,
exact GET paths and named delegated read scopes for 49 enabled contracts.
The register contains 439 top-level field entries, including the two disabled
contracts. A field declared on a derived resource is not necessarily present on
every row returned by its base-resource list.

The verifier resolves CSDL namespace aliases and inheritance, checks known
positive/negative schema controls, records document/metadata SHA256 hashes, and
fails on unverified enabled contracts. A hash identifies the evidence file; it
does not authenticate an imported tenant export.

## Quarantined Reads

| Module | Reason | Behavior |
| --- | --- | --- |
| ModernAssignments | [Policy relationships](https://learn.microsoft.com/en-us/graph/api/resources/intune-deviceconfigv2-devicemanagementconfigurationpolicy?view=graph-rest-beta) and the [SDK list command](https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.beta.devicemanagement/get-mgbetadevicemanagementconfigurationpolicyassignment?view=graph-powershell-beta) exist; an exact GET permission reference was not established in this audit. | Unsupported, zero requests, no completed parents. |
| AutopilotAssignments | The [list reference](https://learn.microsoft.com/en-us/graph/api/intune-enrollment-windowsautopilotdeploymentprofileassignment-list?view=graph-rest-beta) documents a device-identity traversal, not the collector's profile-root traversal. This is a documentation gap, not proof that the route cannot work. | Unsupported, zero requests, no completed parents. |

Do not substitute a write action or grant broader permissions to bypass these
gaps. Profile metadata, modern settings, other supported assignments and the
original two automatic controls remain available.

## Corrections

- Removed generic properties not declared for the queried resource/version,
  including v1 scope tags and non-app assignment `intent`.
- [Win32 v1](https://learn.microsoft.com/en-us/graph/api/resources/intune-apps-win32lobapp?view=graph-rest-1.0)
  uses `rules`, not `detectionRules` / `requirementRules`. Export only rule type,
  operation/operator and OData type, never scripts or commands.
- [Autopilot](https://learn.microsoft.com/en-us/graph/api/resources/intune-enrollment-windowsautopilotdeploymentprofile?view=graph-rest-beta)
  uses `outOfBoxExperienceSetting`; the plural property is deprecated. Current
  collection does not fall back to the deprecated property.
- Conditional Access uses `modifiedDateTime`, not `lastModifiedDateTime`.
- Removed fabricated combinations from the route allowlist.
- Assignment target projections are now version-specific. The published v1
  `deviceAndAppManagementAssignmentTarget` has no filter properties; the beta
  type does. Both inherit `groupId` through the group target subtype. The
  collector removes unverified `entraObjectId` / `targetType` fields. The schema
  verifier checks the actual projected target keys against each version, and
  synthetic tests exercise v1 MAM/app/compliance targets and beta targets.

## Reproduce

Use a new local evidence directory. These downloads are public metadata/docs,
not authenticated tenant calls. The metadata endpoints are documented in
[Traverse Microsoft Graph](https://learn.microsoft.com/en-us/graph/traverse-the-graph).

```powershell
$EvidenceDirectory = Join-Path $env:TEMP ('intune-contracts-' + [guid]::NewGuid())
New-Item -ItemType Directory -Path $EvidenceDirectory | Out-Null
foreach ($Version in @('v1.0', 'beta')) {
    Invoke-WebRequest -UseBasicParsing -Uri ('https://graph.microsoft.com/' + $Version + '/$metadata') -OutFile (Join-Path $EvidenceDirectory ('intune-audit-' + $Version + '.xml'))
}
.\Test-IntuneGraphContracts.ps1 -EvidenceDirectory $EvidenceDirectory -CheckDocumentation
```

The generated `intune-graph-contract-audit.json` retains source URLs, source
hashes, API versions, field owners and quarantine reasons. Reuse the directory
for reproducible cached-document checks; use a fresh directory to refresh docs.
Without `-CheckDocumentation`, only schema checks run.

## Remaining Audit Gates

### Defender Companion (2026-09-18)

The [list API](https://learn.microsoft.com/en-us/defender-endpoint/api/get-machines)
documents GET `https://api.security.microsoft.com/api/machines`, delegated
`Machine.Read`, `$top` (maximum 10,000) and `$skip`. This collector deliberately
uses a page size of 1,000 and does not use filters. Its URL guard now rejects
other queries, duplicate/encoded duplicate parameters and changed page sizes.
Unexpected next links fail Partial rather than silently changing coverage;
only an explicit skip equal to the number of collected rows is followed.
The fallback advances `$skip` after a full page. This is not snapshot isolation:
concurrent tenant changes and access/retention limits can still affect coverage.

The [machine resource](https://learn.microsoft.com/en-us/defender-endpoint/api/machine)
documents the nine exported machine properties. `lastSeen` is the last full
device report, typically daily, not the portal's last-seen value. `version` is
the OS version, not the Defender sensor version. The docs spell onboarding status
with inconsistent casing; PowerShell property lookup is case-insensitive and the
export uses `onboardingStatus`. Unknown/missing values remain unknown. The API
documents 404 for no recent machines; the collector conservatively reports an
error, not proof of tenant-wide absence. Tests pass in PowerShell 7 and 5.1 with
injected responses only. No Defender authentication or live collection was run.

### Open Gates

- Nested object ownership, value types, enums and all policy/CSP bindings.
- CSP URI/format coverage is recorded below; remaining work is payload/enum
  semantics and OS/reference applicability, not the existence of those nodes.
- Full endpoint-provider enum/serialization review on representative devices.
- EPM exact setting IDs currently grounded in Microsoft365DSC sample fixtures,
  not a Microsoft Learn guarantee; do not describe them as a documented CSP.
- Recommendation semantics, applicability, deprecations and reference versions.
- Live delegated authentication, RBAC visibility, licensing, beta responses,
  metadata-only select behavior and representative device evidence.

Existing synthetic tests verify implementation behavior, not live compatibility.
GET-only means no tenant mutations, not no local side effects: authentication
may update an Az context/cache, exports write a new local file, and full API or
provider responses can contain unexported properties in process memory. Unknown
settings are discarded; raw modern setting bodies are read to decode the
reviewed subset. Never log raw bodies or treat projection as proof that no
sensitive data was retrieved.

## Endpoint Command Ledger (2026-09-18)

`EndpointContracts.json` records each of the five provider commands, all 38
projected source properties and their Microsoft Learn references. It also maps
SharedSignaturesPath and SignatureDefinitionUpdateFileSharesSources to path-free derived exports; the other 36 fields are
exported directly. Run
`Test-IntuneEndpointContracts.ps1 -EvidenceDirectory <existing-directory>
-CheckDocumentation` to compare the companion's parsed command/projection AST
with the register and download/hash source evidence. Without the switch, this
is an offline drift test. It never invokes endpoint commands. Source text naming
a field is not runtime type/enum validation or evidence that every OS supports it.

| Read | Verified scope and limits |
| --- | --- |
| `dsregcmd.exe /status` | [DeviceId and TenantId](https://learn.microsoft.com/en-us/entra/identity/devices/troubleshoot-device-dsregcmd) identify joined/hybrid-joined devices, not all registered devices. Status can perform network diagnostics. No join/leave command is used; no claim that this is strictly offline. MDM URLs alone do not prove enrollment. |
| `Get-MpComputerStatus` | Seven version/status fields are shown in the command reference. The controlled-configuration product doc explicitly adds `IsTamperProtected`, `ControlledConfigurationState`, `TamperProtectionSource`; preview/minimum version restrictions apply. Null is not Off. |
| `Get-MpPreference` | ASR IDs/actions are corresponding arrays, not interchangeable enum families. Current product docs explicitly show ASR, network protection and PUA reads. The provider class documents the three Disable booleans and MAPSReporting. A preference alone is not proof of blocking or effective enforcement. |
| `Get-NetFirewallProfile -PolicyStore ActiveStore` | The read reference explicitly distinguishes ActiveStore from the default PersistentStore. The paired Set reference documents the profile options/enums; it is a reference only and is never called. The three GpoBoolean fields can be NotConfigured and must not be coerced to ordinary booleans. No rule or Hyper-V firewall inspection is claimed. |
| `Get-BitLockerVolume` | All five retained metadata properties appear in the command's full-output example. No KeyProtector fields are exported; the provider's full object can still exist in memory before projection. EncryptionPercentage=100 alone does not prove ProtectionStatus=On or recovery escrow. |
| `Get-CimInstance ... Win32_DeviceGuard` | Exact class/namespace and all three selected fields are documented. VBS status 1 is enabled but not running; status 2 is running. SecurityServicesConfigured and SecurityServicesRunning are arrays and are not equivalent. The docs describe an elevated session; this companion never self-elevates, and access failure remains an Error. |

The endpoint modules retain raw selected observations, not generalized verdicts.
Missing fields or provider errors must not become clean findings. JSON exports
are locally authored evidence, not signed attestations; matching a tenant/device
identifier does not prove the file's authenticity.

## CSP Source Ledger (2026-09-18)

`PolicyContracts.json` covers 68 fixed nodes and eight firewall-rule template
leaves. `Test-IntunePolicyContracts.ps1 -EvidenceDirectory <existing-directory>`
downloads/caches their Microsoft CSP references and checks exact published
Device/Vendor URI strings and per-node Format declarations. The generated
`intune-policy-contract-audit.json` records the source URL, SHA256 and format
for all 76 contracts. Firewall rule-name substitution is explicit; names are
not treated as a setting-definition identity or an effective rule instance.

This check found an incorrect URI: the [LSA reference](https://learn.microsoft.com/en-us/windows/client-management/mdm/policy-csp-lsa#configurelsaprotectedprocess)
publishes `Policy/Config/LocalSecurityAuthority/ConfigureLsaProtectedProcess`,
not `Policy/Config/LSA/ConfigureLsaProtectedProcess`. Collector and Assay now
use the published URI and reject the old guessed alias. Its documented format
is integer, with 0 disabled, 1 enabled with UEFI lock, 2 enabled without UEFI lock.

The [SMBv1 nodes](https://learn.microsoft.com/en-us/windows/client-management/mdm/policy-csp-mssecurityguide)
and [PowerShell script-block logging node](https://learn.microsoft.com/en-us/windows/client-management/mdm/policy-csp-windowspowershell)
are `chr` ADMX payloads. Bare numbers for these three nodes now remain
`UnresolvedAdmx`, with no retained value, instead of being labeled Resolved.
No new ADMX binding was inferred from a numeric choice or a display name.
This guard also applies when Assay imports an older or manually authored export.

A published URI/format is not proof that every allowed enum, default, range,
ADMX data ID or reference recommendation has been verified. EPM sample IDs are
not counted among these documented CSP nodes. Unknown settings, template-default
values and missing definitions must remain unresolved. Policy-specific bindings
for the four supported ADMX paths are recorded below; effective application,
cross-setting prerequisites and App Control XML schema review remain separate
gates. XML well-formedness is not full policy validation.

### ADMX Fragment Boundary (18 September 2026)

The parser validates the reviewed Boolean/numeric subset of Microsoft's
[ADMX payload structure](https://learn.microsoft.com/en-us/windows/client-management/understanding-admx-backed-policies):
one empty enabled/disabled element, empty data elements with exactly id/value
attributes, unique bounded identifiers, Boolean literals and integral values. Documented lowercase
and title-case element spellings are accepted. Unsupported elements/namespaces,
nested content, extra/missing attributes, duplicate IDs, mixed states, unsupported
values and disabled-plus-data combinations fail closed. The latter restrictions
define this collector's supported subset; they are not claims that every other
payload is invalid in Windows. Unsupported payloads are recorded unresolved and
never emitted as partially resolved policy evidence. Comments/whitespace do not
change the state. Raw XML remains unexported.

XML-backed CSPs no longer accept numeric/boolean scalar bypasses. Assay validates
the typed ADMX object and every retained field before accepting ResolvedAdmx;
it no longer silently drops invalid fields or truncates a payload into a clean
result. ADMX objects are accepted only on the four structured paths, not scalar
CSPs. OS support, cross-setting prerequisites and App Control XSD
validation remain open gates. Synthetic malformed-input and valid-fragment
tests pass in both PowerShell runtimes and native Rust.

### Policy-Specific ADMX Bindings (0.5.1)

The [BitLocker CSP reference](https://learn.microsoft.com/en-us/windows/client-management/mdm/bitlocker-csp)
explicitly documents the IDs and types used by the three supported policies.
Production decoding now supplies the CSP path to `ConvertTo-IntuneAdmxMetadata`;
the collector and native importer both validate exact case-sensitive IDs, types,
required fields and enum values. No generic numeric field is treated as a known
policy element. Bare `<enabled/>` is insufficient for an enabled BitLocker policy;
`<disabled/>` requires no data. This is the fully specified supported subset, not
a claim about every possible provider representation or Windows policy behavior.

| Policy | Typed data contract |
| --- | --- |
| [SystemDrivesRequireStartupAuthentication](https://learn.microsoft.com/en-us/windows/client-management/mdm/bitlocker-csp#systemdrivesrequirestartupauthentication) | `ConfigureNonTPMStartupKeyUsage_Name` is Boolean; `ConfigureTPMStartupKeyUsageDropDown_Name`, `ConfigurePINUsageDropDown_Name`, `ConfigureTPMPINKeyUsageDropDown_Name`, `ConfigureTPMUsageDropDown_Name` are 0 disallowed, 1 required, 2 optional. |
| [SystemDrivesRecoveryOptions](https://learn.microsoft.com/en-us/windows/client-management/mdm/bitlocker-csp#systemdrivesrecoveryoptions) | `OSAllowDRA_Name`, `OSHideRecoveryPage_Name`, `OSActiveDirectoryBackup_Name`, `OSRequireActiveDirectoryBackup_Name` are Boolean; `OSRecoveryPasswordUsageDropDown_Name` / `OSRecoveryKeyUsageDropDown_Name` are 0 disallowed, 1 required, 2 allowed; `OSActiveDirectoryBackupDropDown_Name` is 1 passwords plus key packages or 2 passwords only. |
| [FixedDrivesRecoveryOptions](https://learn.microsoft.com/en-us/windows/client-management/mdm/bitlocker-csp#fixeddrivesrecoveryoptions) | Same seven element suffixes and types as operating-system recovery, using the documented `FDV` prefix instead of `OS`. |
| [DisableInternetExplorerLaunchViaCOM](https://learn.microsoft.com/en-us/windows/client-management/mdm/policy-csp-internetexplorer#disableinternetexplorerlaunchviacom) | Only the reviewed enable/disable-only fragment is accepted; any data elements remain unresolved. Device and user scope exist in the source; the collector's reviewed comparison uses the device path, not an inferred user/device alias. |

The CSP describes checkbox literals as `true` / `false`; do not coerce a numeric
0/1 into a Boolean or use Boolean true as enum value 1. The source wording for
checkbox false is "Policy not set", not a universal disabled-control verdict.
The complete enabled payload's individual values remain configuration intent.
Valid startup enums may still form a conflicting combination: the CSP warns
that only one additional authentication option may be required. No clean
encryption, escrow or recovery outcome is inferred from payload validation.

`Test-IntuneAdmxContracts.ps1 -EvidenceDirectory <CSP-cache>` compares all 19
BitLocker IDs and their types with the source's XML samples, checks the documented
Boolean/enum descriptions and records source hashes in
`intune-admx-binding-audit.json`. It runs synthetic positive/negative decodes;
optional `ASSAY_INTUNE_ADMX_FIXTURE` writes four decoded policies for native
reassessment testing. Microsoft sample values are not adopted as recommendations.

In Assay rules `1.3.1-preview`, S-16 reports unresolved settings as NotAssessed
rather than an empty decoded observation. S-26 requires Complete modern-policy
and SecuritySettings collection plus completed-parent coverage before Pass;
explicit disabled evidence remains Warning even with incomplete coverage.
Original scored controls, supplementary finding count and decisions are unchanged.

## Other Script Activities

The main script imports an already installed Az.Accounts module only when a
SecureString token was not supplied. It reads context, optionally connects with
process scope, requests an MSGraph token and restores the prior process context
when available. It does not install a module or grant consent. An Az token is
not guaranteed to have the required Intune scopes. Local JWT field checks are
preflight restrictions, not signature authentication; the service authenticates
the token. Opaque tokens and application-only tokens are unsupported here.

Companion JSON and App Control XML are read only from explicit user-supplied
paths. XML parsing prohibits DTDs/external resolution. Export uses CreateNew,
not overwrite, and filenames/metadata can still be confidential. Row counts,
timeouts, retry caps, depth/string limits and file-size limits are engineering
constraints, not Microsoft product recommendations. The collector does not
deploy a script, elevate privileges, apply policy, initiate an EPM elevation,
open a Remote Help session, retrieve a LAPS password or request a BitLocker key.
Full source/provider objects can exist in process memory before projection.