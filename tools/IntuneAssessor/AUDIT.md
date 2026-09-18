# Intune Collector Contract Audit

Date: 2026-09-17. Status: **partial audit, not production certification**.

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

`EndpointContracts.json` records each of the five provider commands, all 32
projected properties and their Microsoft Learn references. Run
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
values and missing definitions must remain unresolved. The existing four
reviewed ADMX payload paths and App Control XML projection still need a separate
complete binding/schema review; XML well-formedness is not full policy validation.

### ADMX Fragment Boundary (18 September 2026)

The parser now validates the reviewed numeric subset of Microsoft's
[ADMX payload structure](https://learn.microsoft.com/en-us/windows/client-management/understanding-admx-backed-policies):
one empty enabled/disabled element, empty data elements with exactly id/value
attributes, unique bounded identifiers and integral values. Documented lowercase
and title-case element spellings are accepted. Unsupported elements/namespaces,
nested content, extra/missing attributes, duplicate IDs, mixed states, nonnumeric
values and disabled-plus-data combinations fail closed. The latter restrictions
define this collector's supported subset; they are not claims that every other
payload is invalid in Windows. Unsupported payloads are recorded unresolved and
never emitted as partially resolved policy evidence. Comments/whitespace do not
change the state. Raw XML remains unexported.

XML-backed CSPs no longer accept numeric/boolean scalar bypasses. Assay validates
the typed ADMX object and every retained field before accepting ResolvedAdmx;
it no longer silently drops invalid fields or truncates a payload into a clean
result. ADMX objects are accepted only on the four structured paths, not scalar
CSPs. Full policy-specific required IDs/enums, OS support and App Control XSD
validation remain open gates. Synthetic malformed-input and valid-fragment
tests pass in both PowerShell runtimes and native Rust.

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