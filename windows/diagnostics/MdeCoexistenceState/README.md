# Get-MdeCoexistenceState

**Version:** 1.1.0
**Author:** Anton Romanyuk

> **Disclaimer:** This script is provided "as-is" without warranty of any kind, express or implied. Use at your own risk. The author assumes no liability for any damage or data loss resulting from its use. Always test in a non-production environment before deployment.

Read-only diagnostic that reports the **effective** state of Microsoft Defender Antivirus and Microsoft Defender for Endpoint on a device, detects **third-party antivirus and EDR products running alongside them**, and reviews the Defender exclusion lists for rules that are broken, inert, or wider than intended.

It answers, in one pass: *what security software is actually running here, is more than one product doing the same job, and are the exclusions doing what someone thought they were doing?*

## Problem

Endpoint security investigations usually start from an Intune or GPO **export**. An export tells you what was *intended*. It does not tell you:

- Whether the policy actually reached the device (assignment failures, filter mismatches, local admin merge, tamper protection interactions).
- Which **other** vendors' filter drivers are attached to the file-system I/O path. Two EDRs, a DLP agent, an analytics agent and a privilege-management agent can all be present without any one team knowing about the others.
- Whether exclusions written with `%USERPROFILE%` or `%APPDATA%` match anything at all. They usually do not — the Defender service runs as **LocalSystem**, so those variables resolve to the *system* profile, not the signed-in user. This is a very common cause of "we excluded the OST file and it's still being scanned."
- Whether the Defender for Endpoint sensor is actually reaching the cloud, as opposed to merely being onboarded.

Gathering this by hand means `Get-MpComputerStatus`, `Get-MpPreference`, `fltmc`, `Get-CimInstance root\SecurityCenter2`, two registry hives and the SENSE operational log — per machine, in the right order, without missing a vector.

## What it checks

| # | Area | Source |
|---|------|--------|
| 1 | Machine identity, OS build, elevation | `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion` |
| 2 | Security services (Microsoft + known third-party AV/EDR) | `Get-Service` |
| 3 | Security processes currently running | `Get-Process` |
| 4 | Defender for Endpoint onboarding state | `HKLM\SOFTWARE\Microsoft\Windows Advanced Threat Protection\Status` |
| 5 | Defender for Endpoint policy, incl. `ForceDefenderPassiveMode` | `HKLM\SOFTWARE\Policies\Microsoft\Windows Advanced Threat Protection` |
| 6 | Defender AV effective state (running mode, RTP, tamper protection) | `Get-MpComputerStatus` |
| 7 | Defender AV preferences (cloud, network protection, CPU throttling) | `Get-MpPreference` |
| 8 | Attack Surface Reduction rules actually applied | `Get-MpPreference` |
| 9 | Effective exclusions **+ automated hygiene review** | `Get-MpPreference` |
| 10 | Loaded minifilters, by altitude band, mapped to products | `fltmc.exe filters` *(needs elevation)* |
| 11 | Products registered with Windows Security Center | `root\SecurityCenter2` |
| 12 | Effective EDR configuration read back from the sensor | SENSE operational log, IDs 1803–1823 |
| 13 | Sensor cloud connectivity and recent errors | SENSE operational log, IDs 4/5/6/7/20 |
| 14 | Installed Defender platform versions | `%ProgramData%\Microsoft\Windows Defender\Platform` |
| 15 | Coexistence verdict | derived |

### Exclusion hygiene findings

| Category | Severity | Meaning |
|---|---|---|
| `EnvVarUnderLocalSystem` | High | `%USERPROFILE%`, `%APPDATA%`, `%LOCALAPPDATA%`, `%TEMP%`, `%TMP%` or `%HOMEPATH%` in an exclusion. The Defender service runs as LocalSystem, so these resolve to the system profile and the rule matches nothing. Rewrite as `C:\Users\*\…`. |
| `AutoStartLocationExcluded` | High | A user-writable auto-start location (Startup / `Start Menu\Programs`) is excluded from scanning — a common persistence path. |
| `DriveRootExcluded` | High | An entire drive is excluded. |
| `ExecutableExtensionExcluded` | High | An executable or script extension (`exe`, `dll`, `ps1`, `js`, …) is excluded wholesale — every file of that type on every drive, wherever it came from. |
| `BareNameInPathList` | Medium | A filename with no path in the **path** exclusion list. Matches nothing; probably meant to be a process exclusion. |
| `BareNameProcess` | Medium | A process exclusion by bare name. Works, but matches any file of that name from any location. Use a full path. |
| `PathAsExtension` | Medium | A path in the **extension** exclusion list. An extension exclusion must be a bare extension; this entry matches nothing. |
| `DriverAsProcess` | Low | A `.sys` file in the process list. A driver never runs as a process, so the rule is inert. Harmless if the containing folder is already path-excluded; otherwise it is a gap. |
| `NonExecutableAsProcess` | Low | A non-executable in the process list. Inert. |
| `BroadRootExcluded` | Low | A whole top-level folder is excluded. Confirm it is as narrow as it can be. |

## Parameters

| Name | Type | Default | Description |
|------|------|---------|-------------|
| `AsObject` | switch | off | Emit the structured result object instead of the console report. Useful for diffing two machines. |
| `JsonPath` | string | — | Write the full structured result to this path as JSON (UTF-8). The console report is still produced unless `-Quiet` is also supplied. |
| `Quiet` | switch | off | Suppress the console report. A single-line JSON summary is still written to STDOUT so Intune script-output harvesting and log scraping keep working. |
| `MaxEvents` | int | 400 | How many SENSE operational log records to read. Raise on a busy or long-uptime machine if the configuration events are not found. |
| `SkipEventLog` | switch | off | Skip all SENSE log queries. Use where the log is very large or slow; sections 12–13 are then reported as not collected. |

## Usage

```powershell
# Full report. Run elevated for the minifilter section.
.\Get-MdeCoexistenceState.ps1

# Console report plus a structured artifact for diffing two machines
.\Get-MdeCoexistenceState.ps1 -JsonPath C:\Temp\mde-state.json

# Capture the object and inspect the exclusion findings
$s = .\Get-MdeCoexistenceState.ps1 -AsObject
$s.Exclusions.Findings | Format-Table Severity, Category, Entry -AutoSize

# Compare two machines
$a = .\Get-MdeCoexistenceState.ps1 -AsObject
# ...on the second machine...
Compare-Object $a.SecurityFilters $b.SecurityFilters -Property Name, Altitude

# Fleet fan-out
Invoke-Command -ComputerName (Get-Content .\hosts.txt) -FilePath .\Get-MdeCoexistenceState.ps1
```

## Exit codes

| Code | Meaning |
|------|---------|
| 0 | OK — single security stack, no hygiene findings |
| 1 | WARNING — coexistence detected, or exclusion findings present |
| 2 | CRITICAL — sensor onboarded but never logged a successful server contact (event 4), or Tamper Protection off while Defender AV is active (Normal mode) |
| 3 | PARTIAL — ran unelevated, or Defender cmdlets unavailable; the result is incomplete |
| 4 | ERROR — unexpected failure caught at top level |

## Interpreting the output

**Elevation matters.** Unelevated, `fltmc.exe` is unavailable and `Get-MpPreference` returns the literal string `N/A: Must be an administrator to view exclusions` in place of every exclusion list. The script detects both cases, reports the exclusions as *not collected* rather than empty, falls back to services / processes / Security Center for the verdict, and exits `3`. **Do not quote an unelevated run as evidence that a machine has no exclusions or no third-party EDR.**

**An exclusion does not remove a filter driver from the I/O path.** Exclusions govern what a product does *after* its minifilter is called. The filter is still attached and still invoked on every file operation. Filter-stack cost is structural; it is not recoverable by exclusion tuning. Do not use this script to argue otherwise.

**Coexistence is not automatically a fault.** Defender Antivirus running in `Normal` mode alongside a third-party EDR is a legitimate design. So is Defender for Endpoint in EDR-block mode with a third-party AV in front. What is *not* legitimate is nobody knowing which of those was intended. The verdict flags the condition; the architecture decision is yours.

**Windows Security Center registration is expected to be sparse.** Most EDR products deliberately do not register as an antivirus provider, precisely so they do not flip Defender into passive mode. A third-party EDR absent from section 11 is normal, not a finding.

## Caveats and confidence

- **Vendor attribution is an inference.** Products are identified from a static lookup table of minifilter, service and process names built into the script. A driver named `CSAgent` is almost certainly CrowdStrike Falcon, but the script does **not** verify the signing certificate. Treat the *Product* column as a strong hint, not proof. Filters that are not in the table are printed with their raw name and altitude band, so nothing is silently dropped — review that list before concluding a machine is clean.
- **Altitude band names** come from Microsoft's published Filter Manager allocated altitude ranges. The band tells you what a filter *claims* to be, not what it does.
- **Microsoft in-box filters** (`UCPD`, `bfs`, `FileInfo`, `Wof`, `CldFlt`, `bindflt`, `luafv` and similar) are deliberately **excluded** from the security-filter count. Counting them inflates the apparent security stack and is the first thing a vendor will pick apart.
- **`ForceDefenderPassiveMode` is reported, not judged.** See "Coexistence is not automatically a fault" above.
- **SENSE event IDs** come from Microsoft's published event table for this log ([Review events and errors using Event Viewer](https://learn.microsoft.com/en-us/defender-endpoint/event-error-codes)): event 4 = *service contacted the server successfully*; 5/6/7/20 are connectivity and onboarding failures. The 1800-series events (1803 last connected, 1804 org ID, 1805 sensor running, 1806 onboarding state, 1807 onboarding blob, 1809 sample sharing, 1823 telemetry frequency) are logged when the MDM/CSP layer reads sensor values, so they appear only on MDM-managed devices — their absence is not a finding.
- **SENSE Error/Warning records are reported for context only.** They are noisy on healthy machines and do not on their own set the CRITICAL exit code.

## Requirements

- PowerShell 5.1+
- Windows 10 / 11 / Server with Microsoft Defender Antivirus present
- Elevation only for the minifilter and exclusion sections; everything else works unelevated
- Read-only. No changes are made to the system.
