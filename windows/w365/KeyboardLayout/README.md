# Set-KeyboardLayout.ps1

Adds a keyboard layout and sets it as the default input method for the current user, Default User profile, and system (welcome screen) — without changing the Windows display language.

Uses the `intl.cpl` InputPreferences XML approach (GlobalizationServices schema), which is the only reliable method on Windows 11 to change the keyboard layout independently of the preferred language list.

## How It Works

1. **Add keyboard** (always): builds an InputPreferences XML that adds the desired keyboard layout (and optionally removes another), then applies it via `control.exe intl.cpl,,/f:"<path>"`. The XML only adds — it does not try to reorder existing keyboards (remove+re-add corrupts the language→`InputMethodTips` association on Windows 11).
2. **Reorder default keyboard**: directly rewrites `HKCU\Keyboard Layout\Preload` so the new keyboard's KLID lands at position `1` (the default input method). Existing KLIDs follow in their original order.
3. **Regional settings** (opt-in): sets culture, GeoId, and/or system locale only when the corresponding parameter is explicitly passed — nothing changes by default.
4. **Copies settings to welcome screen and Default User** by default (via `Copy-UserInternationalSettingsToSystem`). Use `-SkipCopyToSystem` for user-context-only deployments.
5. **Optional — reset modern taskbar input switcher** (`-ResetTaskbarInputSwitcher`): clears residual values under `HKCU\Software\Microsoft\CTF\LangBar` (`ShowStatus`, `Label`, `Transparency`, `ExtraIconsOnMinimized`) so Windows reverts to the default modern indicator. **Off by default** — only opt in if a previous configuration suppressed the indicator.

The script does **not** toggle the legacy language bar — that workaround is only needed for display-language changes and corrupts the `LangBar` values mentioned above.

> **Sign out / reboot required:** The Preload reorder is on disk immediately, but the running session caches the default keyboard at logon. Sign out and back in (or reboot) for the change to take effect in the taskbar input switcher and new apps. A reboot is also required for the welcome screen and Default User profile changes from `Copy-UserInternationalSettingsToSystem`.

## Deployment

### SYSTEM context (Autopilot Device Prep / classic Autopilot / device-targeted)

Deploy as an **Intune platform script** (Devices → Scripts and remediations → Platform scripts):

| Setting | Value |
|---------|-------|
| Run this script using the logged on credentials | **No** (runs as SYSTEM) |
| Enforce script signature check | No |
| Run script in 64-bit PowerShell host | **Yes** |

Autopilot Device Prep supports up to **10 platform scripts** during OOBE — assign the script to the device preparation device group. No Win32 packaging required.

The script automatically detects SYSTEM context and copies settings to the welcome screen and Default User profile.

### User context (already-enrolled devices)

Deploy a second copy in **user context** to change the keyboard for existing user profiles:

| Setting | Value |
|---------|-------|
| Run this script using the logged on credentials | **Yes** |
| Enforce script signature check | No |
| Run script in 64-bit PowerShell host | **Yes** |

> **Two-script pattern:** Sandy Zeng recommends deploying both a SYSTEM-context and a user-context script to cover all scenarios (pre-provisioning + already-deployed devices). The user-context version should use `-SkipCopyToSystem`.

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
| `-SkipCopyToSystem` | off | Skip copying settings to welcome screen and Default User profile. |

## Exit Codes

| Code | Meaning |
|------|---------|
| `0` | Success |
| `1` | Failure (see log for details) |

## Logs

Written to the Intune Management Extension log folder:

```
%ProgramData%\Microsoft\IntuneManagementExtension\Logs\SetKeyboardLayout_<timestamp>.log
```

## Requirements

- Windows 10/11 (Windows 365 Cloud PC or physical)
- PowerShell 5.1+
- Administrator privileges (SYSTEM via Intune, or elevated user session)

## References

- [Managing Windows 11 languages and region settings (Part 2) – Keyboard layout](https://msendpointmgr.com/2025/06/27/managing-windows-11-languages-and-region-settings-part-2/) — Sandy Zeng
- [How to automate InputPreferences during OSD](https://web.archive.org/web/20230315105902/https://vacuumbreather.com/index.php/blog/item/61-how-to-automate-inputpreferences-during-osd) — Anton Romanyuk
- [Default Input Locales for Windows Language Packs](https://learn.microsoft.com/windows-hardware/manufacture/desktop/default-input-locales-for-windows-language-packs)
- [Table of Geographical Locations](https://learn.microsoft.com/windows/win32/intl/table-of-geographical-locations)
