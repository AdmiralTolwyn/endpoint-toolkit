# Remove-StaleAppxPackage

Removes outdated **per-user** registrations of packaged apps (Appx/MSIX) that
vulnerability scanners report, such as the inbox image and video extensions,
without touching users who already run a current version.

## Why this exists

Packaged apps are registered per user. When an app is updated, only users who
sign in receive the new version. Profiles of users who no longer sign in, or
who have left the company, keep the old registration, and the old package stays
under `C:\Program Files\WindowsApps`, where scanners flag it.
See [Modern apps or application packages are reported as vulnerable (KB5011324)](https://learn.microsoft.com/en-us/troubleshoot/windows-client/application-management/modern-apps-application-packages-reported-vulnerable).

Microsoft's guidance, in order of preference:

1. Let affected users sign in so the app updates for them.
2. Delete stale user profiles, e.g. with the Group Policy
   *Delete user profiles older than a specified number of days on system restart*.
3. Remove the outdated package for the affected users only - **this script**.
4. As a last resort, deprovision the app so no user can register any version.

## What it does

For each package name in the list, the script:

1. Enumerates all user registrations (`Get-AppxPackage -AllUsers`).
2. Selects registrations whose package version is **below** the configured
   minimum. Versions are compared as `[version]`; `Get-AppxPackage` returns
   them as strings, where a plain `-lt` gets e.g. `1.0.9.0` vs `1.0.10.0` wrong.
3. Removes each one for that user only (`Remove-AppxPackage -User <SID>`).
4. Re-queries afterwards and reports anything still outdated.

It does **not**:

- Remove provisioned packages. An outdated provisioned copy is logged as a
  warning, because every new profile would register it again - update the
  provisioned package separately.
- Touch well-known service accounts (SYSTEM, LOCAL SERVICE, NETWORK SERVICE).
- Remove packages that have no user registration at all.

## Requirements

- Windows PowerShell 5.1, elevated or running as SYSTEM.
- 64-bit PowerShell host. In Intune, enable *Run script in 64 bit PowerShell Host*.

## Built-in package list

| Package | Minimum version |
| --- | --- |
| Microsoft.HEIFImageExtension | 1.0.43012.0 |
| Microsoft.Microsoft3DViewer | 7.2107.7012.0 |
| Microsoft.MSPaint | 6.2203.1037.0 |
| Microsoft.RawImageExtension | 2.1.30191.0 |
| Microsoft.VP9VideoExtensions | 1.0.42791.0 |
| Microsoft.WebMediaExtensions | 1.0.42192.0 |
| Microsoft.WebpImageExtension | 1.0.42351.0 |

This list is an example, **not** a complete or current list of vulnerable app
versions. Align it with your scanner's findings before deployment. To change it
permanently, edit the ordered table at the top of the script body; to change it
for one run, pass `-MinimumVersion`.

## Parameters

| Parameter | Type | Default | Description |
| --- | --- | --- | --- |
| `-MinimumVersion` | `IDictionary` | built-in list | Package name -> lowest version considered current. Exact names, four-part versions. **Replaces** the built-in list; it is not merged. |
| `-LogPath` | `string` | `%ProgramData%\Microsoft\IntuneManagementExtension\Logs\Remove-StaleAppxPackage.log` | Log file, appended. The Intune *Collect diagnostics* action picks up this folder. |
| `-WhatIf` | `switch` | off | Full dry run: enumerates, warns and logs, removes nothing. |
| `-Confirm` | `switch` | off | Prompt before each user registration is removed. |

Without `-WhatIf` or `-Confirm` the script runs unattended and never prompts.

### Passing your own package list

`-MinimumVersion` takes a hashtable: one `'PackageName' = 'MinimumVersion'`
entry per package, separated by `;` on one line or by line breaks. Every
package you want checked must be in it - the built-in list is not used.

From a PowerShell prompt, inline:

```powershell
.\Remove-StaleAppxPackage.ps1 -WhatIf -MinimumVersion @{
    'Microsoft.VP9VideoExtensions' = '1.0.52781.0'
    'Microsoft.WebMediaExtensions' = '1.0.62192.0'
    'Microsoft.HEVCVideoExtension' = '2.1.1803.0'
}
```

Or built up in a variable first:

```powershell
$list = [ordered]@{
    'Microsoft.VP9VideoExtensions' = '1.0.52781.0'
    'Microsoft.WebMediaExtensions' = '1.0.62192.0'
    'Microsoft.HEVCVideoExtension' = '2.1.1803.0'
}
.\Remove-StaleAppxPackage.ps1 -MinimumVersion $list -WhatIf
```

Use `[ordered]@{...}` if packages must be processed in the listed order;
a plain `@{...}` processes them in arbitrary order.

From `cmd.exe`, ConfigMgr, an RMM tool or a scheduled task, launch with
**`-Command`**, not `-File`. With `-File` every argument arrives as a plain
string and the script fails with *Cannot process argument transformation on
parameter 'MinimumVersion'*.

```text
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& '.\Remove-StaleAppxPackage.ps1' -MinimumVersion @{ 'Microsoft.VP9VideoExtensions' = '1.0.52781.0'; 'Microsoft.WebMediaExtensions' = '1.0.62192.0'; 'Microsoft.HEVCVideoExtension' = '2.1.1803.0' }; exit $LASTEXITCODE"
```

`exit $LASTEXITCODE` passes the script's exit code through explicitly to the
caller.

Intune platform scripts cannot pass parameters. For Intune, edit the built-in
list in the script instead.

The package names and versions above are placeholders. Use the names reported
by your scanner or by `Get-AppxPackage -AllUsers | Select-Object Name, Version`.

## Exit codes

| Code | Meaning |
| --- | --- |
| `0` | No failures and no outdated user registrations left. A `-WhatIf` run returns 0 unless enumeration fails. |
| `1` | Not elevated, invalid version in the list, a removal failed, enumeration failed, or outdated registrations remain. |

## Output

One summary object on the pipeline:

```text
Computer  : PC0042
Mode      : EXECUTE
Removed   : 3
Failed    : 0
Skipped   : 1
WhatIf    : 0
Remaining : 0
LogPath   : C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\Remove-StaleAppxPackage.log
Details   : {...}
```

`Details` holds one object per outdated registration with `Name`,
`PackageFullName`, `Version`, `MinimumVersion`, `Sid`, `User`, `InstallState`,
`Outcome` (`Removed`, `Failed`, `Skipped`, `WhatIf`) and `Error` (HRESULT and
message on failure).

Log lines use the format:

```text
<UTC ISO 8601> [INFO|WARN|ERROR] [PID:<pid>] <message>
```

## Examples

```powershell
# Dry run, shown as a per-user table
(.\Remove-StaleAppxPackage.ps1 -WhatIf).Details |
    Format-Table Name, Version, User, InstallState, Outcome

# Unattended removal with the built-in list
.\Remove-StaleAppxPackage.ps1

# One package with a custom minimum, prompting per user
.\Remove-StaleAppxPackage.ps1 -MinimumVersion @{ 'Microsoft.VP9VideoExtensions' = '1.0.52781.0' } -Confirm

# Two packages on one line, dry run
.\Remove-StaleAppxPackage.ps1 -WhatIf -MinimumVersion @{ 'Microsoft.VP9VideoExtensions' = '1.0.52781.0'; 'Microsoft.WebMediaExtensions' = '1.0.62192.0' }
```

See [Passing your own package list](#passing-your-own-package-list) for
command-line and deployment-tool syntax.

## Deployment

Intune platform script (*Devices > Scripts and remediations > Platform scripts*):

```text
Run this script using the logged on credentials : No
Enforce script signature check                  : No (or sign the script)
Run script in 64 bit PowerShell Host            : Yes
```

ConfigMgr / RMM:

```text
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Remove-StaleAppxPackage.ps1
```

Run `-WhatIf` on a pilot group first and review the log.

## Limitations

- Files can remain in `WindowsApps` after removal; Windows cleans them up
  asynchronously. A rescan immediately afterwards may still flag the folder.
- Validate on a pilot before broad rollout, in particular removal while running
  as SYSTEM and apps installed as bundles. Failures are reported with their
  HRESULT, never silently.

## Disclaimer

Provided "AS IS" with no warranties and no rights conferred. Test in a
non-production environment first.
