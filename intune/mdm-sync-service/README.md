# Intune MDM Sync Service

Simple detection and remediation for the [Microsoft-documented issue where a disabled dmwappushservice prevents Intune sync](https://learn.microsoft.com/en-us/troubleshoot/mem/intune/device-management/cannot-sync-windows-10-devices). Microsoft notes that Intune Management Extension (IME) scripts can still run in this scenario.

## Scripts

| Script | Purpose |
| --- | --- |
| [Detect-IntuneMdmSyncService.ps1](Detect-IntuneMdmSyncService.ps1) | Read the service startup mode and flag Disabled. Makes no changes. |
| [Remediate-IntuneMdmSyncService.ps1](Remediate-IntuneMdmSyncService.ps1) | Recheck the mode, change Disabled to Automatic, and verify by reading it back. |

The pair targets **Disabled only**. Manual and Automatic are unchanged, including on repeated remediation runs. A stopped Manual-start service is not the issue described by the article.

Neither script starts/restarts services, restarts IME, changes enrollment, clears caches or triggers MDM sync. Success confirms the startup configuration, not successful policy application or an updated Intune console timestamp.

## Requirements

- 64-bit Windows PowerShell 5.1 on the affected managed Windows client.
- SYSTEM or an elevated administrator for deployment, with permission to query/change the service.
- No Graph authentication, external modules, parameters or credentials in the scripts.

Both scripts run locally without depending on MDM sync. Intune delivery still needs functioning IME connectivity; use MECM, Nexthink or local execution if that path is also unavailable. Offline devices must become reachable before an agent can deliver the package.

## Intune Deployment

Create an Intune Remediations package and upload each script to its corresponding field:

| Setting | Value |
| --- | --- |
| Detection script | Detect-IntuneMdmSyncService.ps1 |
| Remediation script | Remediate-IntuneMdmSyncService.ps1 |
| Run using logged-on credentials | No (SYSTEM) |
| Run in 64-bit PowerShell | Yes |
| Enforce script signature check | Follow your organization's signing policy; the supplied scripts are unsigned. |

Assign to a pilot group of affected clients before wider rollout. Intune runs remediation when detection returns exit 1. Detection exit 2 is a query/error condition to investigate, not a request to change service configuration.

For MECM or Nexthink, stage the same scripts, run detection as SYSTEM in 64-bit Windows PowerShell 5.1, and run remediation only after detection exit 1. Capture stdout and the process exit code using the agent's supported wrapper/output mechanism.

## Local Use

From this folder, run detection in an elevated Windows PowerShell 5.1 session:

```powershell
.\Detect-IntuneMdmSyncService.ps1
$LASTEXITCODE
```

After reviewing a Disabled result, run the remediation:

```powershell
.\Remediate-IntuneMdmSyncService.ps1
$LASTEXITCODE
```

View help without executing either script:

```powershell
Get-Help .\Detect-IntuneMdmSyncService.ps1 -Full
Get-Help .\Remediate-IntuneMdmSyncService.ps1 -Full
```

## Results

Each script emits a plain-text status/error line.

| Script | Exit | Meaning |
| --- | --- | --- |
| Detection | 0 | Auto/Manual: this disabled-service condition was not found. |
| Detection | 1 | Disabled: run the paired remediation. |
| Detection | 2 | Service missing, query/access failure or unexpected startup mode. |
| Remediation | 0 | Disabled was changed to Automatic and verified, or the service was already Auto/Manual. |
| Remediation | 1 | Query, change or read-back verification failed. |

After remediation, verify actual MDM sync separately. If the service becomes disabled again, investigate the policy, optimization tool or script setting it. This package does not diagnose other enrollment, certificate, network or IME failures.

## Validation

The source pair passed 18 isolated mock-service cases under Windows PowerShell 5.1, including unchanged Auto/Manual, missing/unreadable service, failed changes, failed verification and repeated runs. Syntax, comment-based help and UTF-8 BOM checks passed. The scripts are copied without logic changes.

No live service remediation or management-agent rollout has been validated. Pilot through the intended execution agent before fleet deployment.