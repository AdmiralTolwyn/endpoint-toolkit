# Windows 365 Discovery and Assessor

`Invoke-W365Discovery.ps1` 0.4.0 exports Cloud PC inventory and bounded findings for
[Assay](https://github.com/AdmiralTolwyn/assay) or the legacy Windows/WPF assessor.
Assay has 132 checks: 26 Auto and 106 Manual. The legacy GUI catalog is independent.

## Requirements and Usage

Windows PowerShell 5.1 or PowerShell 7 and `Microsoft.Graph.Authentication` 2.0 or later.
Keep all `W365*.ps1` helpers and `CollectorPrivacy.ps1` beside the collector.
Authentication is delegated, Global-cloud only; app-only and sovereign-cloud contexts are unsupported.

```powershell
.\Invoke-W365Discovery.ps1 -TenantId '<tenant-guid>'
.\Invoke-W365Discovery.ps1 -TenantId '<tenant-guid>' -SkipLogin -IncludeConditionalAccess
```

Import into Assay's W365 pack, or use `Launch_W365Assessor.bat` for the legacy Windows GUI.

## Arguments

| Argument | Default | Purpose |
| --- | --- | --- |
| `-OutputPath` | `%LOCALAPPDATA%\AssayCollections\w365\w365_<collectionId>.json` | New JSON output; never overwrites. |
| `-TenantId` | Validated session/sign-in tenant | Explicit Entra tenant GUID, not a domain alias. |
| `-SkipLogin` | Off | Never sign in; require an existing delegated context with matching tenant and selected scopes. |
| `-IncludeConditionalAccess` | Off | Read CA observations and request `Policy.Read.All`. |
| `-InactiveDays` | `30` | Inactivity threshold in days; missing login evidence is not proof of inactivity. |
| `-ImageAgeWarnDays` | `90` | Custom-image age warning threshold in days. |
| `-IncludeUserExperienceSync` | Off | Read beta UX Sync provisioning-policy metadata with `CloudPC.Read.All`. |
| `-UserExperienceSyncTarget` | `Review` | `Review`, `Enabled`, or `Disabled`; comparisons concern policy intent, not runtime synchronization. |
| `-PrivacyMode` | `Pseudonymous` | `Pseudonymous` or `Identified`; Identified also requires confirmation. |
| `-ConfirmIdentifiedExport` | Off | Explicitly permit Identified export. |
| `-PseudonymKeyPath` | Output extension replaced by `.pseudonym-key` | Load/create a protected Base64 32-byte key; reuse for stable pseudonyms. |
| `-IdentityMapPath` | None | New pseudonym-to-original map; keep separate from shared exports. |
| `-Assessor` | None | Free-text operator label, exported as supplied. |

## Permissions and Limits

| Scope | Collection |
| --- | --- |
| `CloudPC.Read.All` | Cloud PCs, provisioning/user policies, networks, images and bounded reports |
| `DeviceManagementConfiguration.Read.All` | Intune configuration/compliance context and update summary |
| `DeviceManagementManagedDevices.Read.All` | Managed-device and Endpoint Analytics context |
| `Policy.Read.All` | Only with `-IncludeConditionalAccess` |

Role, licensing and service visibility can further restrict evidence. Context checks do not
validate token signatures or establish tenant-wide visibility. Existing sessions may carry broader grants.

- Collection does not change tenant configuration. It is not GET-only: two report POSTs use fixed
  read-scoped requests and retain validated page metadata, not raw report cells. No report pagination
  or live response validation is claimed. Retired recommendation reports make no request.
- Availability, empty reports, RTT and login metadata do not establish performance, resilience,
  successful synchronization or protection. Unsupported findings remain unassessed.
- Default exports pseudonymize UPNs/Cloud PC names, summarize assignments and minimize errors.
  Raw assignments and domain-join usernames are never exported.
- Files remain **Confidential** with restricted ACLs; OneDrive paths warn. Keep keys and identity
  maps separate. Pseudonyms are not anonymization.

## Verification and References

Run the local `Test-W365*.ps1` suites and `../Shared/Test-CollectorDocumentation.ps1`
under both runtimes. Use `Get-Help .\Invoke-W365Discovery.ps1 -Full` for full help.
See [automation review](https://github.com/AdmiralTolwyn/assay/blob/main/docs/W365_BASELINE_AUTOMATION_REVIEW_2026-09-18.md)
and [privacy specification](https://github.com/AdmiralTolwyn/assay/blob/main/docs/COLLECTOR_PRIVACY_SPEC.md).

This script is provided "AS IS" with no warranties and confers no rights.