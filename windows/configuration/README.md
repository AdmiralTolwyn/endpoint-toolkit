# Windows Configuration

Registry and policy tweaks that adjust Windows behavior outside the scope of servicing, security, or feature-specific folders.

## Scripts

| Script | Purpose |
|--------|---------|
| [Set-StartupAppsDelay.ps1](Set-StartupAppsDelay.ps1) | Reverts the Windows 11 "wait-for-idle" startup-app delay (`WaitForIdleState=0`) so Outlook / Teams / Word / Excel launch promptly after sign-in on busy devices. |

---

# Set-StartupAppsDelay.ps1

Reverts the Windows 11 "wait-for-idle" startup-app delay so Run / RunOnce / Startup-folder apps launch promptly after sign-in.

## Background

On recent Windows 11 builds, Explorer no longer launches registered startup apps after a fixed delay. Instead it waits for the system to reach an "idle state" (low CPU + low disk I/O) before kicking them off.

On busy devices that never go idle straight after logon (Defender scan, OneDrive sync, GPO, Intune ESP, EDR / sensor agents, etc.) this leaves Outlook, Teams, Word and Excel apparently *"not starting"* for several minutes after sign-in. End users may perceive this as a boot / performance issue after a Windows feature update, but it is the new launch-scheduling behavior - not a performance problem.

## Registry Keys

Path (per-user): `HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Serialize`

| Value | Type | Default | Fix |
|-------|------|---------|-----|
| `WaitForIdleState` | DWORD | `1` (wait for idle) | **`0`** (legacy: launch immediately after the fixed delay) |
| `StartupDelayInMSec` | DWORD | `10000` (10 s) on legacy builds; honoured again once `WaitForIdleState=0` | **`10000`** (documented pre-22621.1555 default; only written when `-ResetStartupDelay` is supplied) |

`WaitForIdleState = 0` is what reverts the new behavior. `-ResetStartupDelay` is optional and writes `StartupDelayInMSec = 10000` to restore the documented legacy default.

## What This Script Does

1. Writes `WaitForIdleState = 0` to `HKCU\...\Explorer\Serialize` for the running user (creates the subkey if missing). **This reverts the new idle-wait behavior.**
2. **`StartupDelayInMSec` handling depends on context**:
    - **Non-elevated** (user context): not touched - any custom value an admin / GPO / user has set in HKCU is preserved.
    - **Elevated / SYSTEM**: automatically reset to `10000` (the documented pre-22621.1555 default of 10 s) across HKCU / Default User / loaded HKU\<SID>, normalising the value across the device. Opt out with `-ResetStartupDelay:$false`.
3. **Auto-adapts to the run context** (no parameters needed):
    - **Non-elevated** (user-context platform script): patches HKCU only.
    - **Elevated / SYSTEM** (SYSTEM-context platform script): also patches `C:\Users\Default\NTUSER.DAT` (so brand-new profiles inherit the fix) AND every currently loaded user hive `HKU\<SID>` (so signed-in users get fixed too). Built-in service accounts (SYSTEM / LOCAL SERVICE / NETWORK SERVICE), the `.DEFAULT` alias, and `_Classes` subkeys are skipped.
4. Idempotent: skips writes when the value is already correct. Supports `-WhatIf`.
5. `-Revert` removes the values written by a previous run, honouring the same auto-elevation scope as the apply path.

> **No reboot required.** Change applies at the user's next sign-in.

> **Why a platform script and not a Settings Catalog / ADMX / GPP policy?** `WaitForIdleState` and `StartupDelayInMSec` are **not** policy-backed values - there is no ADMX template, no Settings Catalog entry, and no Intune Configuration Profile that exposes them. Group Policy Preferences (Registry extension) could write them on AD-joined devices, but Intune-only / AAD-joined devices have no GPP equivalent. A platform script (or PowerShell-based remediation) is the supported way to set these per-user registry values at scale.

> **Why no parameters?** Intune platform scripts upload the .ps1 verbatim and don't accept arguments at runtime. The script reads the run context once at startup and decides automatically what to patch.
## Deployment

No command-line parameters required - the script auto-detects context. Pick **one** assignment shape:

### Option A - SYSTEM-context platform script (recommended for multi-user devices)

One assignment covers signed-in user(s) **and** future new profiles in the same run.

| Setting | Value |
|---------|-------|
| Run this script using the logged on credentials | **No** (SYSTEM) |
| Enforce script signature check | No |
| Run script in 64-bit PowerShell host | **Yes** |

The elevated context auto-includes:
- `C:\Users\Default\NTUSER.DAT` (new profiles inherit the fix)
- Every loaded `HKU\<SID>` (currently signed-in users)
- SYSTEM's own HKCU (harmless, ignored at sign-in)
- `StartupDelayInMSec` reset to `10000` (the documented pre-22621.1555 default) across all of the above

> **Caveat:** `HKU\<SID>` is only populated for users **loaded at the moment the script runs**. A user who signs in for the first time after the run is covered by the Default User patch. A user with an existing profile who happens to be signed out at run time is picked up the next time Intune re-runs the script after they've signed in.

### Option B - User-context platform script (single-user devices)

For classic AAD-joined laptops where each device has one primary user.

| Setting | Value |
|---------|-------|
| Run this script using the logged on credentials | **Yes** |
| Enforce script signature check | No |
| Run script in 64-bit PowerShell host | **Yes** |

Non-elevated context patches HKCU only - no Default User access (no permission). New profiles on the same device are not pre-patched.

## Examples

```powershell
# Default (no args): patches HKCU. Under elevation also patches Default User
# + every loaded HKU\<SID>. Ship verbatim as an Intune platform script.
.\Set-StartupAppsDelay.ps1

# Also reset StartupDelayInMSec to the documented legacy default (10000 ms)
# - automatic under SYSTEM / elevated; this form is only needed in user context
.\Set-StartupAppsDelay.ps1 -ResetStartupDelay

# Elevated, but skip the HKU\<SID> sweep (e.g. only seed Default User on a
# sysprep'd image where no users are loaded yet)
.\Set-StartupAppsDelay.ps1 -IncludeLoadedUsers:$false

# Back the change out (honours the same auto-elevation scope)
.\Set-StartupAppsDelay.ps1 -Revert
```

## Parameter Quick Reference

All switches are optional - the script picks safe defaults from the run context.

| Parameter | Default | Notes |
|-----------|---------|-------|
| `-ResetStartupDelay` | **auto-on under elevation**, off otherwise | Also writes `StartupDelayInMSec = 10000` (the documented pre-22621.1555 default of 10 s). Opt out of the auto-on with `:$false` to preserve a custom value. |
| `-IncludeDefaultUser` | **auto-on under elevation**, off otherwise | Patches `C:\Users\Default\NTUSER.DAT`. Opt out of the auto-on with `:$false`. |
| `-IncludeLoadedUsers` | **auto-on under elevation**, off otherwise | Enumerates `HKEY_USERS` and patches every loaded `HKU\<SID>`. Opt out with `:$false`. |
| `-Revert` | off | Removes values instead of writing them. Honours the include-* scopes. |
| `-LogDirectory` | IME logs folder if writable, else `$env:TEMP` | Override log location. |
| `-WhatIf` | off | Dry-run; logs the actions without touching the registry. |

## Exit Codes

| Code | Meaning |
|------|---------|
| `0` | Success (or no-op / `-WhatIf`) |
| `1` | Unhandled failure (see log) |

## Logs

Default location:

```
%ProgramData%\Microsoft\IntuneManagementExtension\Logs\PS_Set-StartupAppsDelay_<timestamp>.log
```

The folder is created on demand. The `PS_` prefix marks platform-script output to keep it visually distinct from IME's own logs (`IntuneManagementExtension*.log`) and Win32 app detection logs (`AgentExecutor*.log`). Under user context, if the IME logs folder isn't writable, the script transparently falls back to `%TEMP%`. Override with `-LogDirectory <path>`.

## Requirements

- Windows 10 / 11 (the registry keys are honoured on any build; the new idle-wait behavior only manifests on recent Windows 11 builds that gate startup on an idle-state check).
- PowerShell 5.1+.
- User context for the per-user fix; admin / SYSTEM additionally required for `-IncludeDefaultUser`.

## Notes on the Boot-Performance Triage

Slow-launching Office / Teams after a Windows 11 feature update is **not** a boot or CPU / memory / disk performance issue - it's the new startup scheduler waiting for an idle state that the device never reaches because of background agents. Before chasing the symptoms in a Nexthink performance report, deploy this script and re-test. If the perceived "5 minute Office / Teams launch" disappears, the underlying device performance is likely fine.

If launch latency persists after `WaitForIdleState = 0`, then a real performance investigation is warranted (top CPU / memory consumers, disk latency, post-upgrade services, page-file growth, etc.).

## References

- `HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Serialize` - per-user startup ordering / delay key (legacy `StartupDelayInMSec`, modern `WaitForIdleState`).
