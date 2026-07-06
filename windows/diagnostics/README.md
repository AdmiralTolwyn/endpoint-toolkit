# Get-LocationPolicyState

**Version:** 1.0
**Author:** Anton Romanyuk

> **Disclaimer:** This script is provided "as-is" without warranty of any kind, express or implied. Use at your own risk. The author assumes no liability for any damage or data loss resulting from its use. Always test in a non-production environment before deployment.

Read-only diagnostic that reports the **effective** Windows Location policy state and identifies every author that can force or lock the **Settings → Privacy & security → Location** toggle. It prints a summary table plus a plain-English verdict, so you can tell in one pass *why* location is greyed out (or won't turn on) and *which* layer is responsible.

## Problem

"Location is greyed out / won't turn on" is a common Intune and co-management support case, and it's hard to diagnose because **several independent layers** can gate the same toggle, and the most restrictive one wins. Symptoms are identical regardless of the cause:

- The toggle is disabled with *"Location has been turned off by an admin on this device"* or *"Some settings are managed by your organization."*
- It works for some users (e.g. a pre-existing local admin) but not others (fresh domain / Entra profiles).
- An Intune policy that says "let the user decide" doesn't help — because *user control* defers to whatever state the device already has, and something else is forcing it off.
- Editing one registry value (e.g. the Consent Store) has no effect, because a **different** layer is the real gate.

Manually running `reg query` against each path is slow and easy to get wrong (wrong hive, wrong machine, missing a vector).

## What it checks

The script inspects, in one pass, every known vector that controls the location master switch and per-app access:

| # | Vector | Path |
|---|---|---|
| 1 | MDM `AllowLocation` (effective / winning) | `HKLM\...\PolicyManager\current\device\System\AllowLocation` |
| 2 | MDM `AllowLocation` (per provider) | `HKLM\...\PolicyManager\providers\<GUID>\...\System\AllowLocation` |
| 2b | MDM `LetAppsAccessLocation` (effective) | `HKLM\...\PolicyManager\current\device\Privacy\LetAppsAccessLocation` |
| 3 | Legacy Location GP (+ companions) | `HKLM\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors\DisableLocation` |
| 4 | App-privacy GP (+ per-app force lists) | `HKLM\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy\LetAppsAccessLocation` |
| 5 | Capability Access Manager consent store (HKLM + HKCU, incl. key owner) | `...\CapabilityAccessManager\ConsentStore\location\Value` |
| 5b | DeviceAccess capability-broker gate (HKLM + HKCU) | `...\DeviceAccess\Global\{BFA794E4-F964-4FDB-90F6-51056BFE4B44}\Value` |
| 6 | Settings page hide | `HKLM\...\Policies\Explorer\SettingsPageVisibility` |
| 7 | Geolocation service | `lfsvc` (Start value + running state) |

For MDM `AllowLocation` the value meanings are: **0** = Force Off (greyed off), **1** = User Control (default — *does not turn it on*), **2** = Force On (greyed on). `LetAppsAccessLocation`: **0** = user in control, **1** = force allow, **2** = force deny.

The verdict specifically flags the two situations that are easy to misdiagnose:

- A **force-off policy** (GP `DisableLocation=1` or MDM `AllowLocation=0`) that wins over an Intune "user control" policy and produces the "managed by your organization" banner — and reminds you to check **SCCM/GPO** as well as Intune.
- A **DeviceAccess\Global capability gate = Deny**, which blocks location and **survives Consent Store edits** — the usual reason "I changed `ConsentStore\location\Value` and nothing happened."

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `AsObject` | `switch` | Off | Emit the result as a structured object (`Checks`, `Providers`, `Verdict`) instead of the formatted console report. Useful for piping to `Export-Csv`, `ConvertTo-Json`, or a remediation wrapper. |

## Usage

Run **as administrator** (some MDM/consent-store paths and ACLs need elevation). For the full picture on a "greyed out" case, run **in the affected user's session** so the HKCU vectors reflect that user.

### Basic — formatted report + verdict

```powershell
.\Get-LocationPolicyState.ps1
```

### Structured object (for export / automation)

```powershell
$state = .\Get-LocationPolicyState.ps1 -AsObject
$state.Verdict
$state.Checks | Export-Csv .\location-state.csv -NoTypeInformation
```

### Run on a client and send the output back

```powershell
.\Get-LocationPolicyState.ps1 *> C:\Temp\LocationState.txt
```

## Interpreting the output

- **`MDM AllowLocation (effective)` absent** → the Intune policy never applied to this device (assignment/targeting or sync), *not* a conflict. Collect an MDM report and confirm assignment.
- **`GP DisableLocation = 1`** or **`MDM AllowLocation = 0`** → something is forcing location off; the winning-provider row (or SCCM/GPO) tells you who.
- **`DeviceAccess Global ... = Deny`** → the capability broker is the gate; Consent Store edits won't fix it.
- **`lfsvc` disabled (Start=4)** → the master switch can't turn on until the service is re-enabled.
- **No forcing policy detected but still greyed** → export `ConsentStore\location` and `DeviceAccess\Global` under both HKLM and HKCU and inspect the key ACLs for a user write-deny.

## When to reach for the MDM report / boot logging

If `AllowLocation` is missing or you need to see exactly what Intune delivered:

```powershell
MdmDiagnosticsTool.exe -area DeviceProvisioning -zip C:\Temp\MDMDiag.zip
```

Because the location consent state is typically **written once** at provisioning/first-logon and only *read* afterwards, a short live Procmon capture taken later usually won't contain the write. To catch the original author, use **Procmon boot logging** or enable **registry auditing** on the key and read **Security event 4657**.

## Sources & confidence

- `AllowLocation` / `LetAppsAccessLocation` CSP value meanings are grounded in Microsoft Learn: [Policy CSP – System / AllowLocation](https://learn.microsoft.com/en-us/windows/client-management/mdm/policy-csp-system#allowlocation).
- The **DeviceAccess\Global capability-broker gate (5b)** and the **Consent Store ACL/owner check (5)** come from internal/community research, **not** a formal Learn CSP doc. They are cheap read-only checks and the leading suspects when Consent Store edits have no effect — treat their exact behaviour as "confirm on the affected client," not documented fact.

## Requirements

- Windows PowerShell **5.1** or later.
- Run **elevated** for complete coverage of MDM/consent-store paths and ACLs.
- Read-only: the script makes **no changes** to the system.
