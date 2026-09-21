# Intune MDM Health and Sync

Local diagnostics, disabled-service repair and a one-time MDM sync request for Windows clients. The service pair addresses the [Microsoft-documented issue where a disabled dmwappushservice prevents Intune sync](https://learn.microsoft.com/en-us/troubleshoot/mem/intune/device-management/cannot-sync-windows-10-devices), even while Intune Management Extension (IME) scripts continue working. The separate diagnostic script can run through MECM, Nexthink or an elevated local session without interactive Graph authentication.

## Scripts

| Script | Purpose |
| --- | --- |
| [Detect-IntuneMdmSyncService.ps1](Detect-IntuneMdmSyncService.ps1) | Flag Disabled only; no management configuration changes. Writes a local log. |
| [Remediate-IntuneMdmSyncService.ps1](Remediate-IntuneMdmSyncService.ps1) | Recheck the service, repair Disabled to Automatic if necessary, verify, then submit and observe one enrollment-specific PushLaunch task. |
| [Invoke-IntuneClientSync.ps1](Invoke-IntuneClientSync.ps1) | Collect local join, enrollment/certificate, service, task, network and IME evidence. Diagnostic-only by default; `-Sync` opts into task submission after prerequisite checks. |

Each production script is standalone. The detection pair targets **Disabled only**. Manual and Automatic startup modes are unchanged. A stopped Manual-start service is not the issue described by the article. Direct remediation can still submit one task when the service is already enabled; one-time means once per invocation, not once for the device's lifetime.

No script explicitly starts/restarts services, restarts IME, repairs enrollment, clears caches or changes task definitions. Sync submission can cause Windows to process previously assigned management actions and policies. **A task result of zero is not proof of successful MDM communication, policy application or an updated Intune console timestamp.**

## Requirements

- 64-bit Windows PowerShell 5.1 or PowerShell 7 on Windows; built-in ScheduledTasks module for sync.
- SYSTEM or an elevated administrator with permission to query enrollment/tasks, start the selected task, and change the service when remediation is needed.
- No Graph authentication, embedded credentials, external modules or required parameters. Existing enrollment must support the requested operation.

Intune delivery still needs functioning IME connectivity; use MECM, Nexthink or local execution if that path is unavailable. Offline devices must become reachable before an agent can deliver the package. A missing agent report cannot distinguish powered-off, disconnected, retired or broken-agent devices; correlate independent last-seen evidence.

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

For MECM or Nexthink, use the same detection exit-code distinction. The diagnostic script is also available independently. Capture stdout and process exit codes using the agent's supported wrapper; generic JSON is not a native Nexthink Remote Action adapter. From a 32-bit agent, use `%windir%\Sysnative\WindowsPowerShell\v1.0\powershell.exe` to launch 64-bit Windows PowerShell. Configure an outer timeout longer than task observation plus collection time, initially about five minutes for the default 120-second wait.

## Local Use

From this folder, run detection in elevated 64-bit PowerShell:

```powershell
.\Detect-IntuneMdmSyncService.ps1
$LASTEXITCODE
```

After reviewing a Disabled result, run remediation. This also requests one sync:

```powershell
.\Remediate-IntuneMdmSyncService.ps1 -TimeoutSeconds 120
$LASTEXITCODE
```

Collect diagnostics without requesting sync:

```powershell
.\Invoke-IntuneClientSync.ps1 -OutputFormat DetailedJson
```

Collect, request sync if permitted, and inspect the native report:

```powershell
$report = .\Invoke-IntuneClientSync.ps1 -Sync -TimeoutSeconds 120 -OutputFormat Object -Verbose
$resultCode = $LASTEXITCODE
$report.Assessment | Format-List
$report.Sync | Format-List
```

The diagnostic tool supports `-Sync -WhatIf`: it collects evidence and writes logs, but submits no task. An explicitly supplied `-OutputPath` is still written. Remediation does not provide a WhatIf mode. Avoid repeated or overlapping triggers.

View script help without running an operation:

```powershell
Get-Help .\Detect-IntuneMdmSyncService.ps1 -Full
Get-Help .\Remediate-IntuneMdmSyncService.ps1 -Full
Get-Help .\Invoke-IntuneClientSync.ps1 -Full
```

## Diagnostic Output

| Parameter | Behavior |
| --- | --- |
| `Sync` | Off by default. Approve one task submission after local prerequisite checks. |
| `TimeoutSeconds` | 1-600 seconds, default 120; task observation deadline. Also available on remediation. Does not bound individual scheduler calls or the whole assessment. |
| `OutputFormat` | `SummaryJson` (default), `DetailedJson`, or `Object`. Output mode does not change collection or exit codes. |
| `OutputPath` | Optional full JSON export; creates parent directories and replaces the specified file. Independent of the automatic log. |
| `Verbose` | Diagnostic console progress on stream 4, separate from JSON/object output on stream 1. Progress is logged even when console tracing is off. |

Full reports use `SchemaVersion=3`. `ExecutionContext` records edition, bitness, elevation and SYSTEM status. `Assessment` contains Issues, Unknowns, Observations and CanSync. `EvidenceComplete` means collection completed, not enrollment health. Discovery DNS/direct TCP probes do not validate mutual TLS, proxy handling or the full management exchange. Recent MDM errors may predate the request and are not task-correlated.

`Sync` records selection/submission/observation stage, target enrollment/task, SubmissionAccepted, baseline/latest run times, NewRunObserved, latest state and decimal/hex LastTaskResult. Task codes are separate from local exception HResult. `CloudLastSyncVerified` and `PolicyApplicationVerified` remain false. Preserve the full report because console formatting and agent output limits may truncate nested evidence. WhatIf/confirmation and merged verbose/warning streams can add text; do not feed combined streams to a strict JSON parser.

## Task Safeguards

The sync path selects exactly one `MS DM Server` enrollment with an OMA-DM account. Linked `Microsoft Device Management` entries are not selected. It validates the exact enrollment's `PushLaunch` task: enabled, demand-startable, SYSTEM principal, and one action matching `%windir%\system32\deviceenroller.exe /o "<EnrollmentId>" /c /z`. Missing, disabled, ambiguous or changed targets fail without fallback. No renewal/login tasks are triggered instead.

Already Running/Queued returns immediately without submission. Otherwise, the helper captures LastRunTime, calls Start-ScheduledTask once, and observes for up to TimeoutSeconds. Completion requires a newer timestamp, Ready before and after reading task information, and a final result rather than Running/NotRun/Queued. A fast run can complete between polls. An unchanged/older timestamp never attributes an old success or failure to this request.

Timeout does not stop or retry the task. Timestamp resolution, clock changes or unavailable completion evidence can produce a conservative timeout. Task-wide history/state queries are not atomic and cannot uniquely identify this invocation under concurrent triggers. Built-in task names/actions and registry layout are Windows implementation details, not a guaranteed public sync API; validate on your supported builds. Whether PushLaunch remains alive for the entire MDM exchange is unverified.

## Broken Enrollment

Detection checks only the service and may return zero with a broken enrollment. Consequently, Intune does not automatically run the paired remediation for an enrollment-only problem when the service is enabled.

The diagnostic tool blocks Sync when primary account/certificate checks, tenant comparison or required-service checks identify an issue; ambiguous prerequisite evidence also blocks it. The smaller remediation checks enrollment/account presence and task validity but does not duplicate certificate/tenant-health checks. It retains any service repair when later enrollment/task work fails. Neither script re-enrolls, deletes certificates or repairs a missing task. Even a task returning zero cannot certify that enrollment is working end to end.

## Logs

All three scripts append UTF-8 UTC/PID-stamped entries under `%ProgramData%\Microsoft\IntuneManagementExtension\Logs`, normally `C:\ProgramData\Microsoft\IntuneManagementExtension\Logs`:

- `Detect-IntuneMdmSyncService.log`: run start, detection status/errors and exit code.
- `Remediate-IntuneMdmSyncService.log`: service status/repair, full task result, status messages and exit code.
- `Invoke-IntuneClientSync.log`: progress, full report as compressed JSON, export outcome and script errors.

The folder is created if missing. Diagnostic-only and WhatIf runs also write logs. Log failures warn once on stream 3 without changing the operation result; there is no fallback directory. Logs append without automatic rotation. Include them in retention policies and protect their tenant/device IDs, event messages and other diagnostic data.

## Results

Detection emits one status/error message; remediation emits separate service/task messages. The diagnostic tool emits the chosen JSON/object format.

| Script | Exit | Meaning |
| --- | --- | --- |
| Detection | 0 | Auto/Manual: this disabled-service condition was not found. |
| Detection | 1 | Disabled: run the paired remediation. |
| Detection | 2 | Service missing, query/access failure or unexpected startup mode. |
| Remediation | 0 | Service enabled and a newer task run finished with result zero. MDM sync remains unverified. |
| Remediation | 1 | Service/enrollment/task operation failed or a newer idle task returned a nonzero result. |
| Remediation | 2 | Already running/queued, or no qualifying completion within the observation deadline. |

| Diagnostic verdict | Exit | Meaning |
| --- | --- | --- |
| `AssessmentComplete` | 0 | Local checks completed without a blocking issue; not proof of sync success. |
| `LocalPrerequisiteIssue` | 1 | Detected issue blocks requested sync; inspect Assessment.Issues. |
| `AssessmentIncomplete` | 2 | Insufficient evidence or unsupported execution context; prerequisite Unknowns block requested sync. |
| `MdmTaskCompleted` | 0 | Newer idle task run with result zero; verify service-side check-in separately. |
| `MdmTaskFailed` | 1 | Selection/query/submission/observation failure or nonzero task result. |
| `MdmTaskAlreadyRunning` | 2 | Busy before submission; no new request. |
| `MdmTaskTimedOut` | 2 | Completion not established before the deadline; task not stopped or retried. |
| `ScriptError` | 2 | Unexpected script or report-export failure; an already submitted task may still run. |

Task verdicts take precedence over nonblocking collection errors; inspect EvidenceComplete and CollectionErrors separately. Schema-2 consumers must replace `MdmTaskSubmitted` handling with SubmissionAccepted plus the completion/failure/timeout verdicts. Verify actual check-in, intended policy outcome and console reporting independently. If the service becomes disabled again, investigate the policy, optimization tool or script setting it.

## Validation

Run the bundled offline tests from this folder:

```powershell
powershell.exe -NoProfile -NonInteractive -File .\Test-IntuneClientSync.ps1
powershell.exe -NoProfile -NonInteractive -File .\Test-IntuneMdmSyncService.ps1
pwsh.exe -NoProfile -NonInteractive -File .\Test-IntuneClientSync.ps1
pwsh.exe -NoProfile -NonInteractive -File .\Test-IntuneMdmSyncService.ps1
```

Both suites passed under Windows PowerShell 5.1 and PowerShell 7: 103 diagnostic assertions plus six detection, 44 remediation/task and 18 observation cases. OS operations are mocked, logging uses temporary ProgramData, and no real sync or service change is performed. Tests cover targeting, task states/results, stale history, timeouts, output streams, log appends/failures, help/BOM and parity of standalone helpers.

A user-run elevated PowerShell 7 pilot observed a newer PushLaunch run returning zero. SYSTEM-agent deployment, disabled-service repair on a live target, policy application and service-side check-in have not been validated by these tests. Pilot through the intended execution agent before fleet deployment.

Sources: [disabled dmwappushservice](https://learn.microsoft.com/en-us/troubleshoot/mem/intune/device-management/cannot-sync-windows-10-devices), [Start-ScheduledTask](https://learn.microsoft.com/en-us/powershell/module/scheduledtasks/start-scheduledtask), [Task Scheduler result constants](https://learn.microsoft.com/en-us/windows/win32/taskschd/task-scheduler-error-and-success-constants).