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

- Nested object ownership, value types, enums and all policy/CSP bindings.
- Full endpoint-provider property/enum review and Defender transport queries.
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