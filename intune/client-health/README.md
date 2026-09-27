# Intune Client Major-Issue Discovery

[Discover-IntuneClientMajorIssues.ps1](Discover-IntuneClientMajorIssues.ps1) is a **single, self-contained ConfigMgr Compliance Baseline discovery script**. It screens for major local Intune client faults, while keeping detailed diagnostic evidence for later investigation. No separate IME script, staged helper, Graph authentication or embedded credentials are needed.

## ConfigMgr Setup

1. Create one Windows Configuration Item with a **Script** setting, language **PowerShell**, data type **String**.
2. Import or paste the entire script into the discovery editor. No parameters or wrapper are needed.
3. Set the compliance rule to **Equals `Passed`**. Other returned strings are noncompliant; evaluation errors require investigation.
4. Run in computer/SYSTEM context and a 64-bit PowerShell process. Windows PowerShell 5.1 is the baseline; elevated 64-bit PowerShell 7 also works for local investigation. Non-SYSTEM identity alone is not rejected, but collection still requires elevation and a 64-bit process.
5. Leave remediation scripts empty and automatic remediation disabled. Respect your organization's script-signing policy.
6. Add the CI to a baseline targeted at Windows devices expected to be Intune-enrolled or co-managed. Pilot before wider deployment and verify the effective client script timeout accommodates collection. Do not target ConfigMgr-only devices that legitimately lack Intune enrollment.

## Results

| Returned value | Interpretation |
| --- | --- |
| `Passed` | Core checks completed without a detected major fault. Not proof of successful cloud sync or workload processing. |
| `IssueDetected \| reasons \| Details: log path` | Concrete faults, with up to five concise reasons. The full list remains in the log. |
| No value; script throws | A core check could not be evaluated, references/metadata are unresolved, or execution requirements were not met. The error identifies the problem and log location. |

Example noncompliant result:

```text
IssueDetected | WpnService disabled | Details: C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\Discover-IntuneClientMajorIssues.log
```

The ConfigMgr rule compares the returned string; it does not use an Intune detection-script exit-code convention. No HTML/JSON report is sent as the setting value.

## Major Faults Versus Context

Major findings include:

- Disabled startup or missing required machine services: `Schedule`, `dmwappushservice`, `DmEnrollmentSvc` and `WpnService`; disabled IME is also flagged.
- Missing primary Intune OMA-DM account, expired/not-yet-valid/missing referenced certificate or missing private-key association.
- Enrollment tenant mismatch or reported Entra device-authentication failure.
- Missing critical enrollment/PushLaunch tasks, disabled PushLaunch, or an unexpected demand-start setting, principal or enrollment-specific action.

**Stopped does not mean Disabled.** Stopped-but-enabled services, historical task results, routine MDM/WPN event errors, a certificate approaching expiry, failed direct probes, process counts and missing diagnostic log files do not fail this first-pass screen. Missing IME is retained for assignment review rather than automatically treated as a fault on every device.

Collection failures in core join, enrollment/certificate, service or task checks prevent a pass. Supporting event/log/proxy/process/probe failures alone do not. Core uncertainty is logged and thrown even when other findings exist; it is never silently converted into compliance.

The script still collects broad evidence: OS/join details, primary and linked enrollment references, exact certificates, service configuration, critical and NonCritical task definitions/history, WinHTTP proxy output, bounded discovery-host DNS/TCP results, IME log activity, OMA-DM process snapshots and recent MDM/WPN warnings/errors. A failure in one collector does not stop independent collectors. This is evidence for deeper investigation when the major-fault screen does not explain the symptom, not a reason to flag every log error.

## Log and Local Use

Run from this folder in an elevated 64-bit PowerShell session:

```powershell
.\Discover-IntuneClientMajorIssues.ps1
```

The script creates and appends to:

```text
C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\Discover-IntuneClientMajorIssues.log
```

The location follows `%ProgramData%`. Entries are UTF-8 text with UTC timestamps and process IDs. The structured report records `BaselineAssessment` (the first-pass decision), `Health` (more sensitive diagnostic findings), `Coverage`, collection errors and raw evidence. A diagnostic `ReviewRequired` or unavailable supporting channel in the log does not itself fail the baseline.

Log-write failures warn once without changing an established result. There is no fallback directory or automatic rotation; protect diagnostic identifiers and configure retention. Repeated collection can produce substantial logs.

## Safety and Verification

The script does not request sync, change services, restart IME, renew/delete certificates, recreate tasks, enable event channels or re-enroll a device. It writes its local log and makes bounded direct DNS/TCP probes without credentials.

Service state and local certificate checks do not verify WNS delivery, key access/trust, proxy/TLS behavior or cloud acceptance. Historical errors are not automatically current causes. Task and registry layout checks are implementation-based; validate against supported Windows builds. Per-user `WpnUserService_*` instances are not treated as a device-wide MDM requirement.

This package is derived from the AVD repository's generated comprehensive discovery script. Only the package help and log filename differ. The source has offline PowerShell 5.1/7 tests for major faults, noisy-but-healthy configurations, unknown core evidence, preserved diagnostics and the one-string output contract. ConfigMgr deployment, SYSTEM execution and error presentation still require a customer pilot; no deployment is performed by this package.