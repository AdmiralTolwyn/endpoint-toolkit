# Baseline Collector and BaselinePilot

`Invoke-BaselineCollection.ps1` 1.4.1 collects Windows configuration and event evidence for
[Assay](https://github.com/AdmiralTolwyn/assay). BaselinePilot is the legacy Windows/WPF assessor;
its catalog and evaluator are separate from Assay's 312 checks.

## Requirements and Usage

Windows PowerShell 5.1 or PowerShell 7, local administrator rights, and Windows built-in providers.
The collector reads configuration; it does not apply policy or remediate. Optional providers may
be unavailable on some editions/builds. No external module installation is required.

Run from an elevated PowerShell session:

```powershell
.\Invoke-BaselineCollection.ps1
.\Invoke-BaselineCollection.ps1 -IncludeSpeculationControl -OutputPath C:\Evidence\baseline.json
.\Invoke-BaselineCollection.ps1 -BusinessHours '07:00-19:00' -WorkDays 'Sun-Thu'
```

Import into Assay's Baseline pack. For the legacy GUI, use `Launch_BaselinePilot.bat`.
Use `Get-Help .\Invoke-BaselineCollection.ps1 -Full` for detailed help.

## Arguments

| Argument | Default | Purpose |
| --- | --- | --- |
| `-OutputPath` | `%LOCALAPPDATA%\AssayCollections\baseline\baseline_<collectionId>.json` | New JSON path; never overwrites an existing file. |
| `-LookbackDays` | `30` | Requested event-history window in days; retained history can be shorter. |
| `-MaxEventsPerQuery` | `2000` | Retrieval cap per group, including privilege correlation. Capped queries are incomplete. |
| `-SkipEventCollection` | Off | Skip event queries; event checks remain unassessed. |
| `-EventSummaryOnly` | Off | Export counts rather than records; Assay event checks remain unassessed. |
| `-IncludeGpoData` | Off | Run the slower gpresult/RSOP step for manual review. |
| `-IncludeSpeculationControl` | Off | Run embedded Microsoft SpeculationControl 1.0.19; no install or remediation. |
| `-AssessmentProfile` | `Generic` | `Generic`, `Windows365CloudPc`, `Windows11_25H2`, `WindowsServer2025Member`, or `WindowsServer2025DC`. Bounded applicability, not full certification. |
| `-PrivacyMode` | `Pseudonymous` | `Pseudonymous` or `Identified`; Identified requires confirmation. |
| `-ConfirmIdentifiedExport` | Off | Explicitly permit Identified export. |
| `-PseudonymKeyPath` | Output extension replaced by `.pseudonym-key` | Load/create a protected Base64 32-byte key. Reuse for stable identities. |
| `-IdentityMapPath` | None | New pseudonym-to-original map; protect and share separately. |
| `-Assessor` | None | Operator-supplied free-text label, exported as supplied. |
| `-IncludeSecurityEvents` | Off | Add named diagnostics; account/IP fields also require Identified mode. Never adds command lines, object names or raw messages. |
| `-BusinessHours` | `06:00-22:00` | Device-local, start-inclusive/end-exclusive `HH:mm-HH:mm`; start and end must differ. Overnight windows supported. |
| `-WorkDays` | `Mon-Fri` | Local calendar work days; ranges or lists such as `Sun-Thu` or `Mon,Wed,Fri`. |
| `-Quiet` | Off | Suppress progress/banner output. |

## Evidence Limits

- Area completion is not evidence completeness. Inspect query states, provider errors and missing
  values. Registry absence does not prove secure defaults or a disabled policy.
- 1.4.1 preserves empty arrays and limits crashes to `Application Error` event 1000 with validated
  named filenames. WER 1001 is excluded to avoid unrelated providers and duplicate crash counts.
- Metadata records query outcomes, caps and the oldest retained record per log. Assay withholds
  clean-window verdicts without sufficient retention and current audit prerequisites. Historical
  audit continuity and absence of logging gaps are not attested.
- `AUTH-026` is review evidence, not automatic failure. Disabled auditing cannot prove absence.
- Search resource impact, Defender platform/engine currency, service recovery and orphaned tasks
  require manual review in Assay. The legacy WPF evaluator is unchanged.
- Cloud PC applicability is bounded. SCT profiles override only two audit controls and require
  matching build/role evidence. Do not select a profile merely to reduce unknown results.

Recollect older exports with 1.4.1 and re-import into an updated Assay build. Missing provider and
retention evidence cannot be reconstructed; saved assessments are not automatically migrated.

## Data Handling

Exports remain **Confidential**, not anonymous. Files receive ACLs for the current user, SYSTEM
and Administrators. OneDrive paths warn but are allowed. Resource/domain names and configuration
remain linkable. Keep pseudonym keys and identity maps separate from shared exports.
The privacy helper is embedded, so this collector remains a single deployable script.

## Verification

```powershell
.\Test-BaselinePrivacy.ps1
.\Test-BaselineCollector.ps1
.\Test-BaselineApplicability.ps1
.\Test-SpeculationEvidence.ps1
..\Shared\Test-CollectorDocumentation.ps1
```

Run under both PowerShell runtimes. Mocked-provider tests do not replace live endpoint validation.
See [AUDIT.md](AUDIT.md), [applicability](https://github.com/AdmiralTolwyn/assay/blob/main/docs/BASELINE_APPLICABILITY.md)
and [privacy specification](https://github.com/AdmiralTolwyn/assay/blob/main/docs/COLLECTOR_PRIVACY_SPEC.md).

This script is provided "AS IS" with no warranties and confers no rights.