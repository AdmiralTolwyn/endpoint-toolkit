# Intune Discovery for Assay

`Invoke-IntuneDiscovery.ps1` 0.6.0 collects read-only Graph observations for
[Assay](https://github.com/AdmiralTolwyn/assay). The pack contains 50 controls: two bounded Auto
checks and 48 evidence-assisted Manual checks. Experimental; a live tenant pilot remains required.

## Requirements and Usage

Windows PowerShell 5.1 or PowerShell 7; `Az.Accounts` for delegated sign-in unless supplying an
approved delegated token. Keep this folder's helpers and JSON contracts together. The collector
uses allowlisted GET requests, never tenant changes or credential-field reads.

```powershell
.\Invoke-IntuneDiscovery.ps1 -TenantId '<tenant-guid>'
.\Invoke-IntuneDiscovery.ps1 -TenantId '<tenant-guid>' -IncludeConfiguration -IncludeEntra
```

Import the result into Assay's Intune pack. Service policy observations do not establish effective
endpoint enforcement or assignment. Use the Baseline pack separately for machine-level assessment.

## Arguments

### Invoke-IntuneDiscovery.ps1

| Argument | Default | Purpose |
| --- | --- | --- |
| `-TenantId` | Required for collection | Entra tenant GUID; token tenant must match. |
| `-OutputPath` | `%LOCALAPPDATA%\AssayCollections\intune\intune_<collectionId>.json` | New JSON path; never overwrites. |
| `-IncludeRbac` | Off | Read role definitions/assignments; `DeviceManagementRBAC.Read.All`. |
| `-IncludeAudit` | Off | Read audit events since `AuditSinceUtc`. |
| `-IncludeConfiguration` | Off | Settings catalog, decoded settings, intents/templates, scope tags, filters and supported assignments. |
| `-IncludeEntra` | Off | Device registration and CA; `Policy.Read.DeviceConfiguration` and `Policy.Read.All`. |
| `-IncludeRecoveryMetadata` | Off | LAPS/BitLocker metadata only; `DeviceLocalCredential.ReadBasic.All`, `BitlockerKey.ReadBasic.All`. Never passwords or keys. |
| `-IncludeEnrollment` | Off | Enrollment and Autopilot profiles; `DeviceManagementServiceConfig.Read.All`. |
| `-IncludeApple` | Off | Push-certificate/VPP metadata; `DeviceManagementServiceConfig.Read.All`. |
| `-IncludeMam` | Off | App-protection policies; `DeviceManagementApps.Read.All`. |
| `-IncludeMamLaunch` | Off | App-protection launch conditions; `DeviceManagementApps.Read.All`. |
| `-IncludeRemoteHelp` | Off | Remote Help tenant settings. |
| `-IncludeConnectors` | Off | Threat-defense connectors; `DeviceManagementServiceConfig.Read.All`. |
| `-IncludeAppConfiguration` | Off | App-configuration metadata; `DeviceManagementApps.Read.All`. |
| `-IncludePlatformCompliance` | Off | Modern settings-catalog compliance policies. |
| `-IncludeTunnel` | Off | Microsoft Tunnel sites/servers. |
| `-EndpointEvidencePaths` | Empty | Same-tenant companion endpoint exports, maximum 1 MB each. |
| `-DefenderEvidencePath` | None | Same-tenant Defender companion export, maximum 64 MB. |
| `-AppControlPolicyPaths` | Empty | WDAC policy XML to summarize, maximum 1 MB each. |
| `-AuditSinceUtc` | UTC now minus seven days | Start of the audit window. |
| `-UseExistingConnection` | Off | Reuse the Az.Accounts context instead of signing in. |
| `-GraphAccessToken` | None | Delegated Graph `SecureString` token; tenant, audience, expiry and scopes checked before requests. |
| `-Assessor` | None | Free-text operator label for AssessmentRequirements. |
| `-ScopeDescription` | None | Free-text scope; required to write AssessmentRequirements. |
| `-MaxCollectionAgeHours` | No target | Customer collection-age target, 1-87600 hours. |
| `-MaxPolicyReportAgeHours` | No target | Customer policy-report-age target, 1-87600 hours. |
| `-MaxDeviceSyncAgeDays` | No target | Customer device-sync-age target, 1-87600 days. |
| `-PrivacyMode` | `Pseudonymous` | `Pseudonymous` or `Identified`; Identified requires confirmation. |
| `-ConfirmIdentifiedExport` | Off | Explicitly permit Identified export. |
| `-PseudonymKeyPath` | Output extension replaced by `.pseudonym-key` | Load/create protected Base64 32-byte key; reuse for stable pseudonyms. |
| `-IdentityMapPath` | None | New pseudonym-to-original map; keep separate from exports. |
| `-LibraryOnly` | Off | Load functions without collecting, for tests/library callers. |

Core and optional read scopes are enforced by [GraphContracts.json](GraphContracts.json), not
granted by the script. Service roles/licensing/visibility remain additional requirements. See
[AUDIT.md](AUDIT.md) for enabled and quarantined routes; unsupported assignment reads are not absence.

### Get-IntuneEndpointEvidence.ps1

Run on the target Windows device with Defender, NetSecurity, BitLocker and Device Guard providers.
The device must have an unambiguous Entra identity in the selected tenant. Provider failures are
recorded per module; the script does not elevate itself or remediate.

| Argument | Default | Purpose |
| --- | --- | --- |
| `-TenantId` | Required | Entra tenant GUID matching the device join. |
| `-OutputPath` | Required | New endpoint evidence JSON path. |
| `-LibraryOnly` | Off | Load functions without collecting. |

```powershell
.\Get-IntuneEndpointEvidence.ps1 -TenantId '<tenant-guid>' -OutputPath .\endpoint.json
```

### Get-IntuneDefenderEvidence.ps1

Reads the Defender machines API with an approved delegated token. Device names are omitted;
device-group visibility and retention limit coverage. Do not paste tokens into scripts or reports.

| Argument | Default | Purpose |
| --- | --- | --- |
| `-TenantId` | Required | Entra tenant GUID matching the token. |
| `-AccessToken` | Required | Unexpired delegated Defender `SecureString` token with `Machine.Read`. |
| `-OutputPath` | Required | New Defender evidence JSON path. |
| `-LibraryOnly` | Off | Load functions without collecting. |

## Privacy and Interpretation

Exports are **Confidential**, not anonymous. Default mode pseudonymizes device names, counts
CA/registration members, classifies LAPS names and masks profile/UNC path segments. Join GUIDs
remain available for correlation. Observations inherit the transformed values. Keys/maps are
separate protected files; OneDrive destinations are discouraged. Existing outputs are never overwritten.

Collection completion does not certify decoded settings, applicability or enforcement. Unresolved
option/ADMX/EPM payloads remain unknown. Correlation does not prove assignment, and timestamp/metadata
checks do not prove update delivery, current release versions or runtime protection.

## Verification and References

Use `Get-Help <script> -Full`. Run `Test-Intune*.ps1` and
`../Shared/Test-CollectorDocumentation.ps1` in both runtimes. Graph/policy contract suites need
the external metadata paths documented in their parameters; offline tests do not replace live validation.
See [AUDIT.md](AUDIT.md), [Assay implementation](https://github.com/AdmiralTolwyn/assay/blob/main/docs/INTUNE_ASSESSMENT_IMPLEMENTATION.md)
and [privacy specification](https://github.com/AdmiralTolwyn/assay/blob/main/docs/COLLECTOR_PRIVACY_SPEC.md).

This script is provided "AS IS" with no warranties and confers no rights.