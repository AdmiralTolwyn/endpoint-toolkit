#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Adds a keyboard layout and sets it as the default input method for the current
    user, Default User profile, and system (welcome screen).

.DESCRIPTION
    Generates an InputPreferences XML (GlobalizationServices schema) and applies it
    via control.exe intl.cpl — the only reliable method on Windows 11 to change the
    keyboard layout without altering the display language.

    The script:
      1. Optionally sets region format (culture), home location (GeoId), and system
         locale — only when the corresponding parameter is explicitly supplied.
      2. Builds an InputPreferences XML that adds the desired keyboard (and optionally
         removes another), then applies it via 'control intl.cpl,,/f:"<path>"'.
      3. Reorders HKCU\Keyboard Layout\Preload directly so the new keyboard's KLID
         lands at position 1 (the default input method). This is more reliable than
         trying to control ordering via XML remove+re-add, which corrupts the
         language->InputMethodTips association on Windows 11.
      4. Copies settings to the Default User profile and system account (welcome
         screen) via Copy-UserInternationalSettingsToSystem unless -SkipCopyToSystem
         is specified. A reboot is required for those changes to take effect.
      5. Optional (-ResetTaskbarInputSwitcher): clears residual
         HKCU\Software\Microsoft\CTF\LangBar values that may have been written
         by previous runs or other tools (legacy bar workaround, etc.) to restore
         the modern taskbar input indicator. Off by default — only needed when an
         earlier configuration suppressed the indicator.

    The script intentionally does NOT toggle the legacy language bar. That workaround
    is only needed for display-language changes via the XML; for keyboard-only changes
    it's unnecessary and corrupts HKCU\Software\Microsoft\CTF\LangBar values, which
    suppresses the modern taskbar input indicator.

    Designed for Intune platform script deployment on Windows 365 Cloud PCs
    and physical Autopilot devices.

    Autopilot Device Prep:
      Supports up to 10 platform scripts during OOBE (System context).
      Assign the script to the device preparation device group — no Win32
      packaging required.

    Two-script pattern (Sandy Zeng):
      * SYSTEM context (device-targeted / pre-provisioning): settings are
        automatically copied to welcome screen and new-user profiles.
      * USER context (user-driven Autopilot / already-enrolled devices): only
        the current user's keyboard is changed.

.PARAMETER AddInputLocale
    Input Locale ID to add, in 'LLLL:KKKKKKKK' format (Language:Keyboard).
    Default: '0807:00000807' (German (Switzerland) — Swiss German keyboard).

    Common values:
      0807:00000807   German (Switzerland) — Swiss German keyboard
      0409:00000807   English (US) — Swiss German keyboard (keeps en-US as
                      preferred language; app UI stays English)
      0409:0000040B   English (US) — Finnish keyboard
      040C:0000040C   French (France) — French keyboard
      100C:0000100C   French (Switzerland) — Swiss French keyboard
      0407:00000407   German (Germany) — German keyboard
      0809:00000809   English (UK) — United Kingdom keyboard
      1009:00001009   English (Canada) — Canadian Multilingual keyboard

    Full reference:
      https://learn.microsoft.com/windows-hardware/manufacture/desktop/default-input-locales-for-windows-language-packs

.PARAMETER RemoveInputLocale
    Input Locale ID to remove. Default: empty (existing keyboards are kept).
    Pass '0409:00000409' to remove the US English keyboard.

.PARAMETER GeoId
    Geographical location ID for Set-WinHomeLocation. Only applied when explicitly
    specified. Example: 223 (Switzerland), 77 (Finland), 84 (France).
    Reference: https://learn.microsoft.com/windows/win32/intl/table-of-geographical-locations

.PARAMETER Culture
    Culture / region format string for Set-Culture. Only applied when explicitly
    specified. Controls date, time, number, and currency formatting.
    Example: 'de-CH', 'fr-FR', 'fi-FI'.

.PARAMETER SystemLocale
    System locale for Set-WinSystemLocale. Only applied when explicitly specified.
    Controls the language used for non-Unicode programs. Example: 'en-US', 'de-CH'.

.PARAMETER SkipCopyToSystem
    Skips copying the current user's international settings to the welcome
    screen and Default User profile. Use this for user-context-only deployments
    where the SYSTEM-context script handles the system-wide copy.

.PARAMETER ResetTaskbarInputSwitcher
    Clears residual values under HKCU\Software\Microsoft\CTF\LangBar
    (ShowStatus, Label, Transparency, ExtraIconsOnMinimized) so Windows reverts
    to its default behavior of showing the modern taskbar input indicator.
    Off by default. Only use when a previous configuration (e.g. an older
    version of this script that toggled the legacy language bar, or another
    tool) has suppressed the indicator.

.NOTES
    Version : 1.4.0
    Author  : Anton Romanyuk
    Context : Intune platform script / Autopilot Device Prep. Runs as SYSTEM or user.
    Requires: Windows 10/11, PowerShell 5.1+, admin.

    Input Locale format — 'LLLL:KKKKKKKK':
      LLLL      Language Identifier (LCID low word, hex)
      KKKKKKKK  Keyboard Layout Identifier (KLID, hex)

    References:
      https://msendpointmgr.com/2025/06/27/managing-windows-11-languages-and-region-settings-part-2/
      https://web.archive.org/web/20230315105902/https://vacuumbreather.com/index.php/blog/item/61-how-to-automate-inputpreferences-during-osd
      https://learn.microsoft.com/windows-hardware/manufacture/desktop/default-input-locales-for-windows-language-packs
      https://learn.microsoft.com/windows/win32/intl/table-of-geographical-locations

.DISCLAIMER
    This script is provided "AS IS" with no warranties and confers no rights.
    It is not supported under any Microsoft standard support program or service.
    Use of this script is entirely at your own risk. The customer is solely
    responsible for testing and validating this script in their environment
    before deploying to production. The author shall not be liable for any
    damage or data loss resulting from the use of this script.

.EXAMPLE
    .\Set-KeyboardLayout.ps1
    Adds Swiss German keyboard (0807:00000807) alongside the existing US English
    keyboard and sets it as the default input method. Settings are copied to the
    welcome screen and Default User profile (reboot required).

.EXAMPLE
    .\Set-KeyboardLayout.ps1 -GeoId 223 -Culture 'de-CH'
    Swiss German keyboard + Switzerland region format. System locale unchanged.

.EXAMPLE
    .\Set-KeyboardLayout.ps1 -RemoveInputLocale '0409:00000409'
    Adds Swiss German keyboard and removes the US English keyboard so only
    Swiss German remains.

.EXAMPLE
    .\Set-KeyboardLayout.ps1 -AddInputLocale '0409:00000807' -RemoveInputLocale '0409:00000409'
    Replaces the US English keyboard with Swiss German under the English (US)
    language. App UI stays English; only the input method changes.

.EXAMPLE
    .\Set-KeyboardLayout.ps1 -GeoId 223 -Culture 'de-CH' -SystemLocale 'de-CH'
    Swiss German keyboard + full Switzerland regional settings including system
    locale (non-Unicode programs use German (Switzerland)).

.EXAMPLE
    .\Set-KeyboardLayout.ps1 -AddInputLocale '040C:0000040C' -GeoId 84 -Culture 'fr-FR'
    French keyboard + France region format. US English keyboard kept.
#>

[CmdletBinding()]
param(
    [ValidatePattern('^[0-9a-fA-F]{4}:[0-9a-fA-F]{8}$')]
    [string]$AddInputLocale = '0807:00000807',

    [AllowEmptyString()]
    [string]$RemoveInputLocale = '',

    [int]$GeoId,

    [string]$Culture,

    [string]$SystemLocale,

    [switch]$SkipCopyToSystem,

    [switch]$ResetTaskbarInputSwitcher
)

#region --- Configuration ---
$LogFolder = Join-Path $env:ProgramData 'Microsoft\IntuneManagementExtension\Logs'
$LogFile   = Join-Path $LogFolder ("SetKeyboardLayout_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
#endregion

#region --- Helpers ---
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO'
    )
    $ts    = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$ts] [$Level] $Message"
    if (-not (Test-Path -LiteralPath $LogFolder)) {
        [void](New-Item -Path $LogFolder -ItemType Directory -Force)
    }
    Add-Content -Path $LogFile -Value $entry -ErrorAction SilentlyContinue
    $color = switch ($Level) {
        'WARN'    { 'Yellow' }
        'ERROR'   { 'Red'    }
        'SUCCESS' { 'Green'  }
        default   { 'Gray'   }
    }
    Write-Host $entry -ForegroundColor $color
}

function Test-RunningAsSystem {
    [Security.Principal.WindowsIdentity]::GetCurrent().User.Value -eq 'S-1-5-18'
}
#endregion

#region --- Main ---
$ErrorActionPreference = 'Stop'
$exitCode = 0
$xmlPath  = $null

try {
    $isSystem     = Test-RunningAsSystem
    $copySettings = -not $SkipCopyToSystem.IsPresent

    Write-Log '=== Set-KeyboardLayout v1.4.0 ===' -Level SUCCESS
    Write-Log "Running as      : $(if ($isSystem) { 'SYSTEM' } else { [Environment]::UserName })"
    Write-Log "Add keyboard    : $AddInputLocale"
    Write-Log "Remove keyboard : $(if ($RemoveInputLocale) { $RemoveInputLocale } else { '(none)' })"
    Write-Log "Copy to default user / system: $copySettings"

    # Validate RemoveInputLocale format when not empty
    if ($RemoveInputLocale -and $RemoveInputLocale -notmatch '^[0-9a-fA-F]{4}:[0-9a-fA-F]{8}$') {
        Write-Log "-RemoveInputLocale '$RemoveInputLocale' is not a valid LLLL:KKKKKKKK format" -Level ERROR
        exit 1
    }

    # --- Step 1: Region settings (opt-in — only when explicitly specified) -----------------
    if ($PSBoundParameters.ContainsKey('SystemLocale')) {
        Write-Log "Setting system locale to '$SystemLocale'"
        Set-WinSystemLocale -SystemLocale $SystemLocale
    }
    if ($PSBoundParameters.ContainsKey('Culture')) {
        Write-Log "Setting culture (region format) to '$Culture'"
        Set-Culture -CultureInfo $Culture
    }
    if ($PSBoundParameters.ContainsKey('GeoId')) {
        Write-Log "Setting home location (GeoId) to $GeoId"
        Set-WinHomeLocation -GeoId $GeoId
    }
    if (-not ($PSBoundParameters.ContainsKey('SystemLocale') -or
              $PSBoundParameters.ContainsKey('Culture') -or
              $PSBoundParameters.ContainsKey('GeoId'))) {
        Write-Log 'No regional settings specified — keyboard-only mode'
    }

    # --- Step 2: Build InputPreferences XML ------------------------------------------------
    # Only ADD the new keyboard. Earlier versions tried to remove+re-add existing
    # keyboards to control ordering, but that corrupts the language->InputMethodTips
    # association in HKCU\Control Panel\International\User Profile (Get-WinUserLanguageList
    # returns languages with empty tips after remove+re-add). Default ordering is
    # handled separately by directly rewriting HKCU\Keyboard Layout\Preload.
    $userElement  = '    <gs:User UserID="Current"/>'
    $inputActions = @("    <gs:InputLanguageID Action=`"add`" ID=`"$AddInputLocale`"/>")
    if ($RemoveInputLocale) {
        $inputActions += "    <gs:InputLanguageID Action=`"remove`" ID=`"$RemoveInputLocale`"/>"
    }

    $xmlContent = @"
<gs:GlobalizationServices xmlns:gs="urn:longhornGlobalizationUnattend">

  <!-- user list -->
  <gs:UserList>
$userElement
  </gs:UserList>

  <!-- input preferences -->
  <gs:InputPreferences>
$($inputActions -join "`n")
  </gs:InputPreferences>

</gs:GlobalizationServices>
"@

    $xmlPath = Join-Path $env:TEMP "InputPreferences_$(New-Guid).xml"
    Write-Log "Writing InputPreferences XML to '$xmlPath'"
    Set-Content -Path $xmlPath -Value $xmlContent -Encoding UTF8 -Force
    Write-Log "XML content:`n$xmlContent"

    # --- Step 3: Apply InputPreferences via intl.cpl ---------------------------------------
    # control.exe forks rundll32.exe (shell32.dll Control_RunDLL) to actually process
    # the .cpl and returns immediately. We must wait for the rundll32 child to exit,
    # otherwise the next steps run against a stale Preload list.
    Write-Log 'Applying InputPreferences via control.exe intl.cpl'
    $rundll32Before = @(Get-Process -Name 'rundll32' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    Start-Process -FilePath "$env:SystemRoot\System32\control.exe" `
                  -ArgumentList "intl.cpl,,/f:`"$xmlPath`"" `
                  -Wait -NoNewWindow

    $deadline = (Get-Date).AddSeconds(30)
    while ((Get-Date) -lt $deadline) {
        $newRundll32 = @(Get-Process -Name 'rundll32' -ErrorAction SilentlyContinue |
                            Where-Object { $rundll32Before -notcontains $_.Id })
        if ($newRundll32.Count -eq 0) { break }
        Start-Sleep -Milliseconds 500
    }
    Start-Sleep -Seconds 3  # registry flush

    # --- Step 4: Reorder HKCU\Keyboard Layout\Preload to set the new keyboard as default ---
    # The Preload subkey holds REG_SZ values "1", "2", "3", ... whose data is the KLID
    # (Keyboard Layout Identifier, e.g. "00000807"). Position "1" is the default.
    # The XML add only appends to Preload; it never reorders. So we rewrite Preload
    # so the new KLID is at position 1 and existing KLIDs follow.
    $preloadKey = 'HKCU:\Keyboard Layout\Preload'
    $newKlid    = ($AddInputLocale -split ':')[1]   # "0807:00000807" -> "00000807"
    if (Test-Path -LiteralPath $preloadKey) {
        $existing = @()
        foreach ($name in (Get-Item -LiteralPath $preloadKey).Property) {
            $existing += [pscustomobject]@{
                Name = $name
                Klid = (Get-ItemProperty -LiteralPath $preloadKey -Name $name).$name
            }
        }
        # Sort by numeric Name to preserve stable existing order
        $existing = $existing | Sort-Object { [int]$_.Name }
        Write-Log "Preload before: $((($existing | ForEach-Object { '{0}={1}' -f $_.Name, $_.Klid }) -join ', '))"

        # Build new ordered KLID list: new one first, then existing minus duplicates
        $orderedKlids = @($newKlid)
        foreach ($e in $existing) {
            if ($e.Klid -ne $newKlid -and $orderedKlids -notcontains $e.Klid) {
                $orderedKlids += $e.Klid
            }
        }

        # Wipe and rewrite Preload values
        foreach ($e in $existing) {
            Remove-ItemProperty -LiteralPath $preloadKey -Name $e.Name -ErrorAction SilentlyContinue
        }
        $i = 1
        foreach ($klid in $orderedKlids) {
            Set-ItemProperty -LiteralPath $preloadKey -Name "$i" -Value $klid -Type String
            $i++
        }
        Write-Log "Preload after : $((($orderedKlids | ForEach-Object -Begin {$j=1} -Process { $r='{0}={1}' -f $j,$_; $j++; $r }) -join ', '))"
    } else {
        Write-Log "Preload key '$preloadKey' not found — skipping reorder" -Level WARN
    }

    # --- Step 5: Copy to Default User / system (welcome screen) ----------------------------
    if ($copySettings) {
        Write-Log 'Copying international settings to Default User profile and system account'
        Copy-UserInternationalSettingsToSystem -WelcomeScreen $true -NewUser $true
    }

    # --- Step 6: Optional — reset modern taskbar input switcher --------------------------
    # Removes residual values under HKCU\Software\Microsoft\CTF\LangBar that suppress
    # the modern input indicator. These can be left behind by older versions of this
    # script (which toggled the legacy language bar), other tooling, or manual changes.
    # Off by default — only opt in when the indicator is missing on the target system.
    if ($ResetTaskbarInputSwitcher) {
        $langBarKey = 'HKCU:\Software\Microsoft\CTF\LangBar'
        if (Test-Path -LiteralPath $langBarKey) {
            Write-Log 'Clearing legacy LangBar overrides to restore modern taskbar input switcher'
            foreach ($name in 'ShowStatus','Label','Transparency','ExtraIconsOnMinimized') {
                Remove-ItemProperty -LiteralPath $langBarKey -Name $name -ErrorAction SilentlyContinue
            }
        }
    }

    # --- Step 7: Verify --------------------------------------------------------------------
    $langList = Get-WinUserLanguageList
    Write-Log 'Current language list:' -Level SUCCESS
    foreach ($lang in $langList) {
        Write-Log "  Language: $($lang.LanguageTag), InputMethodTips: $($lang.InputMethodTips -join ', ')"
    }

    Write-Log '=== Keyboard layout configuration complete ===' -Level SUCCESS
    if ($copySettings) {
        Write-Log 'A reboot is required for welcome screen and new-user profile changes to take effect.' -Level WARN
    }
}
catch {
    Write-Log "FATAL: $($_.Exception.Message)" -Level ERROR
    Write-Log $_.ScriptStackTrace -Level ERROR
    $exitCode = 1
}
finally {
    if ($xmlPath -and (Test-Path -LiteralPath $xmlPath)) {
        Remove-Item -Path $xmlPath -Force -ErrorAction SilentlyContinue
        Write-Log 'Removed temporary InputPreferences XML'
    }
}

exit $exitCode
#endregion
