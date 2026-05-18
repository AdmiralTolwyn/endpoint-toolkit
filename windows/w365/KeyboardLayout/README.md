# Set-KeyboardLayout.ps1

Adds a keyboard layout and sets it as the default input method — without changing the Windows display language.

A single script that auto-detects its execution context (SYSTEM vs. logged-on user) and runs the appropriate set of steps. Deploy the same file as one Intune platform script in SYSTEM context, another in USER context, and the script does the right thing in each.

Uses the `intl.cpl` InputPreferences XML approach (GlobalizationServices schema), which is the only reliable method on Windows 11 to change the keyboard layout independently of the preferred language list.

## How It Works

**Common to both contexts:**

1. **Add keyboard**: builds an InputPreferences XML that adds the desired keyboard layout (and optionally removes another), then applies it via `control.exe intl.cpl,,/f:"<path>"`. The XML only adds — it does not try to reorder existing keyboards (remove+re-add corrupts the language→`InputMethodTips` association on Windows 11).
2. **Reorder default keyboard**: directly rewrites `HKCU\Keyboard Layout\Preload` so the new keyboard's KLID lands at position `1` (the default input method). Existing KLIDs follow in their original order.

**SYSTEM context only:**

3. **Patch loaded user hives**: walks `HKU\<SID>` for every loaded real-user hive and applies the same Preload reorder. Critical on W365 / Autopilot where the user profile is already loaded when the device-targeted script runs.
4. **Copy to Default User / welcome screen** (`Copy-UserInternationalSettingsToSystem`, on by default). Seeds brand-new profiles. Skip with `-SkipCopyToSystem`. Requires reboot.
5. **Regional settings** (opt-in): `-Culture`, `-GeoId`, `-SystemLocale`. Admin-only cmdlets; silently skipped in user context.

**USER context only:**

6. **Live session switch**: calls `Set-WinDefaultInputMethodOverride` so the running session's text services framework picks up the new default immediately — no sign-out needed. The persisted HKCU change is also captured into the user's roaming profile container (W365 Cloud Profile / FSLogix VHDX) at sign-out, fixing the container permanently.

**Optional (any context):**

- `-ResetTaskbarInputSwitcher` clears residual values under `HKCU\Software\Microsoft\CTF\LangBar` (`ShowStatus`, `Label`, `Transparency`, `ExtraIconsOnMinimized`) so Windows reverts to the default modern indicator. **Off by default** — only opt in if a previous configuration suppressed the indicator.

The script does **not** toggle the legacy language bar — that workaround is only needed for display-language changes and corrupts the `LangBar` values mentioned above.

> **Sign out / reboot behavior:**
> - SYSTEM run: Preload writes are on disk immediately but the running session caches the default keyboard at logon, so signed-in users must sign out for the change to surface in the taskbar input switcher.
> - USER run: `Set-WinDefaultInputMethodOverride` flips the live session immediately — no sign-out needed.
> - Welcome screen / new-user profile changes from `Copy-UserInternationalSettingsToSystem` require a **reboot**.

## Deployment

The script is designed to be deployed **twice in Intune** from the same file — once per context. Pick the deployment that matches your scenario:

| Scenario | SYSTEM deployment | USER deployment |
|----------|-------------------|-----------------|
| Autopilot Device Prep / Autopilot — brand-new device | ✅ Required | Optional |
| Windows 365 Enterprise — first-time provisioning | ✅ Required | Optional |
| Windows 365 Frontline (Cloud Profiles / FSLogix) — returning users | ✅ Required (new containers) | ✅ **Required** (existing containers) |
| Already-enrolled devices, fix existing profiles | ✅ Required | ✅ **Required** |

**Why both for W365 Frontline / FSLogix?** Returning users mount their existing `NTUSER.DAT` from a roaming container that predates any SYSTEM script run. SYSTEM-context Default User edits are bypassed for those users; only the USER-context pass can patch their live HKCU (which then persists back into the container at sign-out).

### SYSTEM context (device-targeted)

Deploy as an **Intune platform script** (Devices → Scripts and remediations → Platform scripts) or as part of Autopilot Device Prep:

| Setting | Value |
|---------|-------|
| Run this script using the logged on credentials | **No** (runs as SYSTEM) |
| Enforce script signature check | No |
| Run script in 64-bit PowerShell host | **Yes** |

Assign to a **device group**. Autopilot Device Prep supports up to 10 platform scripts during OOBE.

### USER context (user-targeted)

Deploy the same file a second time:

| Setting | Value |
|---------|-------|
| Run this script using the logged on credentials | **Yes** |
| Enforce script signature check | No |
| Run script in 64-bit PowerShell host | **Yes** |

Assign to a **user group**. No special parameters needed — the script auto-detects it's running as a standard user and skips the SYSTEM-only steps (`Copy-UserInternationalSettingsToSystem`, loaded-hive patching, regional cmdlets).

## Input Locale Format

Input Locale IDs follow the format `LLLL:KKKKKKKK`:

| Part | Description |
|------|-------------|
| `LLLL` | Language Identifier (LCID low word, hex) |
| `KKKKKKKK` | Keyboard Layout Identifier (KLID, hex) |

### Common Input Locale IDs

| ID | Language — Keyboard |
|----|---------------------|
| `0807:00000807` | German (Switzerland) — Swiss German keyboard |
| `0409:00000807` | English (US) — Swiss German keyboard (app UI stays English) |
| `0407:00000407` | German (Germany) — German keyboard |
| `040C:0000040C` | French (France) — French keyboard |
| `100C:0000100C` | French (Switzerland) — Swiss French keyboard |
| `0410:00000410` | Italian (Italy) — Italian keyboard |
| `0810:00000810` | Italian (Switzerland) — Swiss Italian keyboard |
| `0409:0000040B` | English (US) — Finnish keyboard |
| `0809:00000809` | English (UK) — United Kingdom keyboard |

Full reference: [Default Input Locales for Windows Language Packs](https://learn.microsoft.com/windows-hardware/manufacture/desktop/default-input-locales-for-windows-language-packs)

### Common GeoIds

| GeoId | Region |
|-------|--------|
| 223 | Switzerland |
| 77 | Finland |
| 84 | France |
| 94 | Germany |
| 118 | Italy |
| 244 | United States |
| 242 | United Kingdom |

Full reference: [Table of Geographical Locations](https://learn.microsoft.com/windows/win32/intl/table-of-geographical-locations)

## Examples

```powershell
# Keyboard-only: adds Swiss German alongside US English (default)
.\Set-KeyboardLayout.ps1

# Swiss German keyboard + Switzerland region format
.\Set-KeyboardLayout.ps1 -GeoId 223 -Culture 'de-CH'

# Swiss German + full regional settings (including system locale)
.\Set-KeyboardLayout.ps1 -GeoId 223 -Culture 'de-CH' -SystemLocale 'de-CH'

# Replace US English with Swiss German (only Swiss German remains)
.\Set-KeyboardLayout.ps1 -RemoveInputLocale '0409:00000409'

# Swiss German keyboard under English (US) language — app UI stays English
.\Set-KeyboardLayout.ps1 -AddInputLocale '0409:00000807' -RemoveInputLocale '0409:00000409'

# French keyboard, France region (US English kept)
.\Set-KeyboardLayout.ps1 -AddInputLocale '040C:0000040C' -GeoId 84 -Culture 'fr-FR'

# Swiss French keyboard, Switzerland region
.\Set-KeyboardLayout.ps1 -AddInputLocale '100C:0000100C' -GeoId 223 -Culture 'fr-CH'
```

## Parameter Quick Reference

| Parameter | Default | Notes |
|-----------|---------|-------|
| `-AddInputLocale` | `0807:00000807` | Input Locale ID to add (`LLLL:KKKKKKKK`). |
| `-RemoveInputLocale` | *(empty)* | Input Locale ID to remove. Existing keyboards kept by default. |
| `-GeoId` | *(not set)* | Geographical location for `Set-WinHomeLocation`. Only applied when specified. |
| `-Culture` | *(not set)* | Region format for `Set-Culture` (date/time/number). Only applied when specified. |
| `-SystemLocale` | *(not set)* | System locale for non-Unicode programs. Only applied when specified. |
| `-SkipCopyToSystem` | off | SYSTEM context only. Skip copying settings to welcome screen and Default User profile. Ignored in user context (no admin). |
| `-ResetTaskbarInputSwitcher` | off | Clear residual `HKCU\Software\Microsoft\CTF\LangBar` values to restore the modern input indicator. |

## Exit Codes

| Code | Meaning |
|------|---------|
| `0` | Success |
| `1` | Failure (see log for details) |

## Logs

Log location depends on execution context:

| Context | Path |
|---------|------|
| SYSTEM | `%ProgramData%\Microsoft\IntuneManagementExtension\Logs\SetKeyboardLayout_<timestamp>.log` |
| User   | `%LocalAppData%\Microsoft\IntuneManagementExtension\Logs\SetKeyboardLayout_<timestamp>.log` |

## Requirements

- Windows 10/11 (Windows 365 Cloud PC or physical)
- PowerShell 5.1+
- No admin requirement for the script itself. SYSTEM-only steps (regional cmdlets, `Copy-UserInternationalSettingsToSystem`, loaded-hive patching) are skipped automatically when running as a standard user.

## References

- [Managing Windows 11 languages and region settings (Part 2) – Keyboard layout](https://msendpointmgr.com/2025/06/27/managing-windows-11-languages-and-region-settings-part-2/) — Sandy Zeng
- [How to automate InputPreferences during OSD](https://web.archive.org/web/20230315105902/https://vacuumbreather.com/index.php/blog/item/61-how-to-automate-inputpreferences-during-osd) — Anton Romanyuk
- [Default Input Locales for Windows Language Packs](https://learn.microsoft.com/windows-hardware/manufacture/desktop/default-input-locales-for-windows-language-packs)
- [Table of Geographical Locations](https://learn.microsoft.com/windows/win32/intl/table-of-geographical-locations)
