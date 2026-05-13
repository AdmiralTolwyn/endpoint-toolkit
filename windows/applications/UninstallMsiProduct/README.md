# Uninstall-MsiProduct

Generic, PowerShell 5.1-compatible MSI uninstaller driven by name pattern,
publisher, version, or explicit `ProductCode`. Built for fleet use when the
.msi file is not available locally and the GUID changes with every release
(e.g. the Quest / KACE Agent — see
[Quest KB 4269674](https://support.quest.com/en-us/kb/4269674)).

## Why this exists

Most "uninstall by name" snippets on the internet either:

1. Use `Get-WmiObject Win32_Product` / `Get-CimInstance Win32_Product` — which
   triggers an MSI **self-repair pass on every installed product** as a side
   effect of being enumerated, or
2. Hard-code a single `ProductCode` GUID — which breaks the moment the vendor
   ships a new version with a new PackageCode.

This script reads the **Uninstall registry hive** directly (HKLM 64-bit,
WOW6432Node, and optionally HKCU), filters by wildcard pattern, and invokes
`msiexec.exe /x {GUID} /qn /norestart /l*v <log>` for each match — no source
.msi required, no `Win32_Product` side effects.

## Safety model

| Gate                   | Behaviour                                                                                   |
| ---------------------- | ------------------------------------------------------------------------------------------- |
| No filter supplied     | Refuses to run (exit `1605`). At least one of `-DisplayName`, `-Publisher`, `-Version`, `-ProductCode` must be set. |
| Too many matches       | If discovery returns more than `-MaxMatches` (default **10**), the run aborts (exit `1604`) unless `-Force` is supplied. |
| Interactive default    | `ConfirmImpact='High'` — every match prompts unless `-Force` (or `-Confirm:$false`) is supplied. |
| `-WhatIf` / `-ListOnly`| Discovery is printed; `msiexec` is **not** invoked.                                          |
| Per-product timeout    | `msiexec` is killed after `-TimeoutSeconds` (default 900s). The product is marked Failed (exit 1460). |

## Parameters

| Parameter              | Type        | Default      | Description |
| ---------------------- | ----------- | ------------ | ----------- |
| `-DisplayName`         | `string[]`  | —            | One or more DisplayName wildcards (case-insensitive `-like`). Multiple patterns are OR'd. |
| `-Publisher`           | `string`    | —            | Wildcard filter applied after `-DisplayName`. |
| `-Version`             | `string`    | —            | DisplayVersion wildcard (string match, not SemVer compare). |
| `-ProductCode`         | `string[]`  | —            | Explicit GUIDs (with or without braces). Bypasses all other filters. |
| `-IncludePerUser`      | `switch`    | off          | Also scan `HKCU:\…\Uninstall` of the user running the script. |
| `-Architecture`        | `Both/X64/X86` | `Both`   | Limit which HKLM views are scanned. |
| `-AdditionalArguments` | `string[]`  | `@()`        | Extra args appended to msiexec, e.g. `REBOOT=ReallySuppress`. |
| `-MaxMatches`          | `int`       | `10`         | Safety cap; exceeded → exit 1604 unless `-Force`. |
| `-Force`               | `switch`    | off          | Suppress prompts and bypass `-MaxMatches`. |
| `-TimeoutSeconds`      | `int`       | `900`        | Per-product timeout. `0` = wait forever. |
| `-LogDirectory`        | `string`    | `$env:TEMP`  | Folder for the script log + per-product MSI verbose logs. |
| `-ListOnly`            | `switch`    | off          | Discovery-only; no msiexec invoked. |

## Exit codes (script level)

| Code | Meaning                                                                       |
| ---- | ----------------------------------------------------------------------------- |
| `0`    | All matches uninstalled successfully (or `-ListOnly` completed).            |
| `1602` | User cancelled at the confirmation prompt for every target.                 |
| `1603` | At least one product failed to uninstall.                                   |
| `1604` | `-MaxMatches` safety gate tripped. Refine the filter or pass `-Force`.      |
| `1605` | No matching products found (or no filter supplied).                         |
| `3010` | Success, but a reboot is required.                                          |

Per-product MSI exit codes (0 / 1605 / 1641 / 3010 / 1460 / 1618 / …) are
captured on the returned result objects and translated into a friendly
`Action` field (`Removed`, `RebootPending`, `NotInstalled`, `Failed`,
`Skipped`, `WhatIf`, `ListOnly`).

## Output

The script emits one `PSCustomObject` per attempt on the pipeline:

```text
DisplayName     : Quest KACE Agent
DisplayVersion  : 11.2.117
Publisher       : Quest Software Inc.
ProductCode     : {12345678-90AB-CDEF-1234-567890ABCDEF}
Architecture    : X64
Scope           : Machine
ExitCode        : 0
Action          : Removed
DurationSeconds : 12.4
MsiLog          : C:\Users\…\AppData\Local\Temp\Uninstall-MsiProduct_Quest_KACE_Agent_12345678-90AB-…_20260513_141207.msi.log
```

## Examples

### Quest / KACE Agent (the original use case)

```powershell
# Dry-run preview first
.\Uninstall-MsiProduct.ps1 -DisplayName 'Quest*KACE Agent*' -Publisher '*Quest*' -ListOnly

# Non-interactive removal (Intune Win32 / scheduled task)
.\Uninstall-MsiProduct.ps1 -DisplayName 'Quest*KACE Agent*' -Publisher '*Quest*' -Force
```

### Remove by explicit ProductCode

```powershell
.\Uninstall-MsiProduct.ps1 -ProductCode '{12345678-90AB-CDEF-1234-567890ABCDEF}' -Force
```

### Bulk cleanup (raise the safety cap)

```powershell
.\Uninstall-MsiProduct.ps1 -DisplayName 'Acme *' -Version '4.*' -MaxMatches 50 -Force
```

### Pass vendor-specific MSI properties

```powershell
.\Uninstall-MsiProduct.ps1 -DisplayName 'Foglight*' -Force `
    -AdditionalArguments 'REBOOT=ReallySuppress','MSIRESTARTMANAGERCONTROL=Disable'
```

### WhatIf preview

```powershell
.\Uninstall-MsiProduct.ps1 -DisplayName '*agent*' -WhatIf
```

## Intune Win32 app packaging

```text
Install command   : powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Uninstall-MsiProduct.ps1" -DisplayName 'Quest*KACE Agent*' -Publisher '*Quest*' -Force
Uninstall command : exit 0
Detection rule    : Registry — verify the Uninstall key for the target product is ABSENT
                    (HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall and WOW6432Node)
Return codes      : 0 = Success, 3010 = Soft reboot, 1604/1605/1602 = Retry, 1603 = Failed
Run as            : System
```

## Limitations

- Vendor agents often require additional cleanup (services, drivers, certificates,
  scheduled tasks, leftover folders, EDR tamper-protection bypass) that this
  script does **NOT** perform. It only invokes the MSI's documented uninstall
  sequence. Layer a vendor-specific post-clean step on top if required.
- Per-user installs (`HKCU`) are only scanned for the user **running** the script.
  Touching other profiles would require loading each `NTUSER.DAT` — out of scope.
- `-Version` is a wildcard string compare, not a SemVer comparison. Use
  `'4.*'`, not `'<5.0'`.

## Disclaimer

Provided "AS IS" with no warranties and no rights conferred. Test in a
non-production environment first. The customer is solely responsible for
validating impact (especially for security/monitoring agents) before mass
deployment.
