# AVD Discovery and Assessor

`Invoke-AvdDiscovery.ps1` 0.7.0 exports Azure Virtual Desktop evidence for
[Assay](https://github.com/AdmiralTolwyn/assay) or the legacy Windows/WPF AVD Assessor.
Assay and the legacy GUI have separate catalogs/evaluators; their check counts are not interchangeable.

## Requirements and Usage

Windows PowerShell 5.1 or PowerShell 7 with Azure PowerShell (`Az`) modules and an account
authorized to read each selected subscription. Graph/Log Analytics queries additionally need
their service permissions and visibility. Missing access is not evidence that a resource is absent.
Keep `CollectorPrivacy.ps1` beside the discovery script.

```powershell
.\Invoke-AvdDiscovery.ps1 -SubscriptionId '<subscription-guid>'
.\Invoke-AvdDiscovery.ps1 -SubscriptionId @('<sub-1>', '<sub-2>') -SkipLogin
```

Include shared hub/storage subscriptions even if they have no host pools. Import the JSON into
Assay's AVD pack, or launch the legacy GUI with `Launch_AvdAssessor.bat` on Windows.

## Arguments

| Argument | Default | Purpose |
| --- | --- | --- |
| `-SubscriptionId` | Interactive selection | One or more subscription IDs; Enter at selection uses the current subscription. |
| `-OutputPath` | `%LOCALAPPDATA%\AssayCollections\avd\avd_<collectionId>.json` | New JSON output; never overwrites. |
| `-SkipLogin` | Off | Reuse the existing Az context without interactive sign-in. |
| `-IncludeGuestChecks` | Off | Execute read-only FSLogix inspection via VM Run Command on up to three running hosts per pool. Requires the VM runCommand action; not a control-plane-only read. |
| `-IncludeMdeDeviceChecks` | Off | Graph hunting with `ThreatHunting.Read.All` and Defender device-group access. Match by Azure VM/resource IDs, never hostname fallback. |
| `-PrivacyMode` | `Pseudonymous` | `Pseudonymous` or `Identified`; Identified also requires confirmation. |
| `-ConfirmIdentifiedExport` | Off | Explicitly permit Identified export. |
| `-PseudonymKeyPath` | Output extension replaced by `.pseudonym-key` | Load/create a protected Base64 32-byte key. |
| `-IdentityMapPath` | None | New pseudonym-to-original map, kept separate from shared evidence. |
| `-Assessor` | None | Free-text operator label exported as supplied; not derived from sign-in. |
| `-IncludeTagValues` | Off | Export tag values only with Identified mode; tag keys remain available by default. |

Use `Get-Help .\Invoke-AvdDiscovery.ps1 -Full` for detailed help.

## Data and Evidence Limits

- Default export keeps tag keys, removes MDE device/RBAC principal names and classifies public
  network addresses. Resource names/IDs, domain names and configuration remain linkable.
- Outputs are **Confidential**, not anonymous, with restricted ACLs. OneDrive destinations warn
  but are allowed. Never share keys or identity maps with exports.
- Guest checks are opt-in and execute inside VMs. Review authorization and sampling limits first.
- Empty or failed Graph/hunting responses do not establish absence of devices or protection.
  Manually review unsupported checks; the collector is not a compliance certification.
- Legacy GUI saves and exports use restricted ACLs. Assay adds bounded legacy redaction and
  inventory retention; the original source file is not rewritten.

## Verification

Run `Test-AvdDiscovery.ps1`, `Test-AvdPrivacy.ps1` and
`../Shared/Test-CollectorDocumentation.ps1` under both supported PowerShell runtimes.
Offline tests do not replace live Azure/Graph validation.
See [Assay's AVD data audit](https://github.com/AdmiralTolwyn/assay/blob/main/docs/AVD_DISCOVERY_DATA_AUDIT_2026-08-25.md)
and [privacy specification](https://github.com/AdmiralTolwyn/assay/blob/main/docs/COLLECTOR_PRIVACY_SPEC.md).

This script is provided "AS IS" with no warranties and confers no rights.