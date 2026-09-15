# OneDrive Photos Shortcut Cleanup

Detection/remediation scripts that remove Start Menu / Desktop shortcuts for the **"OneDrive Photos"** app surface on managed Windows 11 devices — without touching the OneDrive client itself.

> ⚠ **Temporary mitigation.** No supported admin control suppresses this app surface today. Shortcut removal is cosmetic — the app stays installed and launchable from Search. **Do not build organisational dependencies on these scripts.**

## Background

- "OneDrive Photos" is a UI surface of `OneDrive.App.exe`, which ships as part of the Microsoft OneDrive desktop client. It is not a Store app and not independently installable.
- Launching it redirects to `onedrive.live.com` and prompts for a **personal Microsoft account (MSA)**.
- Its appearance tracks the Windows build and OneDrive client version rather than any tenant setting, and no GPO/CSP/ADMX setting currently suppresses it (including `DisablePersonalSync` — verify against your own build before assuming otherwise).

## Scripts

| Script | Role |
|--------|------|
| [Detect-OneDrivePhotosShortcut.ps1](Detect-OneDrivePhotosShortcut.ps1) | Exit `1` if any `.lnk` resolves to `OneDrive.App.exe`, else `0`. Any unexpected error also exits `0` — fail-safe, so an unknown state never triggers blind remediation. |
| [Remediate-OneDrivePhotosShortcut.ps1](Remediate-OneDrivePhotosShortcut.ps1) | Deletes exactly those shortcuts. Idempotent. Exits non-zero only if a deletion was attempted and failed. |

Both scan the all-users Start Menu, the Public Desktop, and every real user profile's Start Menu and Desktop. Profiles come from `HKLM\...\CurrentVersion\ProfileList` (no hardcoded `C:\Users`), filtered to `S-1-5-21-*` and `S-1-12-1-*` so system and service profiles are skipped.

Matching is on the shortcut's **resolved `TargetPath`**, never its display name — the name is localized and may change between client versions. A user-created `OneDrive Photos.lnk` pointing elsewhere is left alone; a differently-named shortcut pointing at `OneDrive.App.exe` is still caught.

## What these never touch

`OneDrive.exe` (sync engine, Files On-Demand, KFM), `OneDrive.App.exe` itself, anything under `C:\Program Files\Microsoft OneDrive\`, and the registry — no speculative keys, since inventing one risks conflicting with any future supported control.

## Dry run

```powershell
.\Remediate-OneDrivePhotosShortcut.ps1 -WhatIf
```

Reports each match via the standard `What if: Performing the operation "Remove shortcut" on target "<path>"` line plus a count, and exits `0` without touching the filesystem.

`-WhatIf` exists only on the remediation script. The detection script never changes anything, so running it *is* the dry run — passing it `-WhatIf` fails with `A parameter cannot be found`, which is expected.

Add `-Verbose` to either script to trace each match and skipped root. The trace goes to the verbose stream, never stdout, so the result line the management platform parses is unchanged.

## Deployment

### Intune Remediations

| Setting | Value |
|---------|-------|
| Run this script using the logged-on credentials | **No** |
| Enforce script signature check | No |
| Run script in 64-bit PowerShell | **Yes** |
| Schedule | **Daily** (not once — the shortcut returns after OneDrive client updates and for new profiles) |
| Assignment | Pilot device group first |

### ConfigMgr

- **Configuration baseline** — Configuration Item with a script setting (String): discovery script = the detect script, compliance rule = *value returned* **Begins with** `COMPLIANT`; remediation script = the remediate script, with *Run the specified remediation script when this setting is noncompliant* enabled. Evaluate daily, run as SYSTEM.
- **Package/Program or Script** — deploy the remediation script on a daily recurring schedule. It is a clean no-op (exit `0`) when nothing is found, so unconditional re-runs are safe.

### Startup script

The remediation script works as a GPO computer startup script or a scheduled task at startup, running as SYSTEM. Discovery is filesystem-based via ProfileList, so no user needs to be logged on and no `HKCU` access is required. Startup-only execution misses shortcuts re-created between reboots; pair it with a daily trigger where reboot cadence is slow.

Set *Run script in 64-bit PowerShell* = Yes. Both scripts assert their own bitness and fail loudly rather than silently scanning a WOW64-redirected path.

## Logs

Plain text, one line per event, default location:

```
%ProgramData%\Microsoft\IntuneManagementExtension\Logs\PR_OneDrivePhotosShortcut.log
```

Override with `-LogPath`. A value ending in `.log` is used as-is; anything else is treated as a folder and the default file name appended. Missing folders are created.

```powershell
.\Detect-OneDrivePhotosShortcut.ps1    -LogPath 'C:\Windows\CCM\Logs'
.\Remediate-OneDrivePhotosShortcut.ps1 -LogPath 'D:\Logs\odp.log'
```

Logging is best-effort — an unwritable path is ignored silently and never affects the exit code or the stdout verdict. There is no rotation; the file grows slowly (a few lines per run) and is safe to delete.

## Enforced blocking with WDAC / AppLocker (out of scope)

**Removing the shortcut only hides the entry.** The app stays launchable via Search, the Run dialog, and direct invocation. If your requirement is that users *cannot reach it*, layer a WDAC or AppLocker deny rule on `OneDrive.App.exe` until a supported control exists.

Authoring those policies is deliberately not part of this deliverable — a misapplied rule has a far larger blast radius than a deleted `.lnk`. Two cautions:

1. **Validate `OneDrive.exe` sync is unaffected.** `OneDrive.App.exe` and `OneDrive.exe` are separate binaries in the same directory, signed by the same publisher. A publisher rule written too broadly takes the sync engine with it — check Files On-Demand, Known Folder Move and SharePoint sync on a pilot device first.
2. **Audit before enforce.** The app is a component of a client that updates on its own cadence, so file attributes can shift under you.

## Retirement triggers

Short expected lifespan. Re-evaluate when **any** of these occurs:

- A supported policy control for this app surface becomes available — an ADMX, CSP or Settings Catalog setting is strictly better than deleting shortcuts.
- The app's role on the device changes — if it becomes a user's primary photo-viewing entry point, removing the shortcut stops being cosmetic.
- Account handling changes such that the MSA redirect that motivated this cleanup no longer applies.

## Requirements

Windows 10 / 11, PowerShell 5.1, no external modules (`WScript.Shell` COM for `.lnk` resolution). SYSTEM or elevated administrator, needed to reach every profile's folders.

## Disclaimer

Provided **AS-IS** without warranty of any kind. Test in a pilot ring before broad deployment.
