<#
.SYNOPSIS
    Adds a keyboard layout and sets it as the default input method.
    Single script that handles both SYSTEM and USER deployment contexts.

.DESCRIPTION
    Generates an InputPreferences XML (GlobalizationServices schema) and applies it
    via control.exe intl.cpl - the only reliable method on Windows 11 to change the
    keyboard layout without altering the display language.

    The script auto-detects its execution context (SYSTEM vs. logged-on user) and
    runs the appropriate set of steps:

    Common steps (both contexts):
      1. Optionally sets region format (culture), home location (GeoId), and system
         locale - only when the corresponding parameter is explicitly supplied.
         (Cmdlets require admin; silently skipped in user context.)
      2. Builds an InputPreferences XML that adds the desired keyboard (and optionally
         removes another), then applies it via 'control intl.cpl,,/f:"<path>"'.
      3. Reorders HKCU\Keyboard Layout\Preload so the new keyboard's KLID lands at
         position 1 (the default input method).

    SYSTEM context (device-targeted, e.g. Intune SYSTEM platform script /
    Autopilot Device Prep):
      4S. Patches HKU\<SID>\Keyboard Layout\Preload for every loaded real-user hive.
          Critical on W365 / Autopilot where a user profile is already loaded when
          the script runs.
      5S. Copies the current user's international settings to the Default User
          profile and welcome screen via Copy-UserInternationalSettingsToSystem
          (unless -SkipCopyToSystem). Seeds brand-new user profiles created from
          Default User. Reboot required to take effect.

    USER context (logged-on user, e.g. Intune USER platform script -
    'Run this script using the logged on credentials = Yes'):
      4U. Calls Set-WinDefaultInputMethodOverride so the running session's text
          services framework picks up the change immediately (no sign-out needed
          for the active session). The persisted HKCU change is also captured in
          the user's roaming profile container (W365 Cloud Profile / FSLogix VHDX)
          at sign-out, fixing the container permanently.

    Optional (any context):
      6. -ResetTaskbarInputSwitcher clears residual HKCU\Software\Microsoft\CTF\
         LangBar values that may have been written by previous runs or other tools
         to restore the modern taskbar input indicator. Off by default.

    Why two contexts? On Windows 365 Frontline (Cloud Profiles) and any FSLogix
    environment, returning users mount their existing NTUSER.DAT from a roaming
    container that predates the SYSTEM script. The SYSTEM script's Default User
    edits are bypassed for those users; only the USER-context pass can patch their
    live HKCU (which then persists into the container at sign-out).

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
    SYSTEM context only. Skips Copy-UserInternationalSettingsToSystem (no seeding
    of Default User / welcome screen). Useful when the system-wide copy was already
    done by a previous run, or for testing. Ignored in USER context (no admin).

.PARAMETER ResetTaskbarInputSwitcher
    Clears residual values under HKCU\Software\Microsoft\CTF\LangBar
    (ShowStatus, Label, Transparency, ExtraIconsOnMinimized) so Windows reverts
    to its default behavior of showing the modern taskbar input indicator.
    Off by default. Only use when a previous configuration (e.g. an older
    version of this script that toggled the legacy language bar, or another
    tool) has suppressed the indicator.

.NOTES
    Version : 1.6.0
    Author  : Anton Romanyuk
    Context : Intune platform script (SYSTEM or USER). No admin requirement.
              SYSTEM-only steps are skipped automatically when running as a
              standard user.
    Requires: Windows 10/11, PowerShell 5.1+.

    Input Locale format — 'LLLL:KKKKKKKK':
      LLLL      Language Identifier (LCID low word, hex)
      KKKKKKKK  Keyboard Layout Identifier (KLID, hex)

    References:
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
# Log to a context-appropriate location. SYSTEM writes to the IME log folder
# (ProgramData is writable and Intune already collects from there); standard
# users write to their own LocalAppData since ProgramData isn't writable.
if ([Security.Principal.WindowsIdentity]::GetCurrent().User.Value -eq 'S-1-5-18') {
    $LogFolder = Join-Path $env:ProgramData 'Microsoft\IntuneManagementExtension\Logs'
} else {
    $LogFolder = Join-Path $env:LOCALAPPDATA 'Endpoint-Toolkit\Logs'
}
$LogFile = Join-Path $LogFolder ('SetKeyboardLayout_{0}.log' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
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

function Get-LoadedUserHive {
    <#
    .SYNOPSIS
        Returns SID + friendly name for every real-user hive loaded under HKEY_USERS.
    #>
    $builtIn = @('S-1-5-18','S-1-5-19','S-1-5-20')
    foreach ($sub in (Get-ChildItem -Path 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue)) {
        $sid = $sub.PSChildName
        if ($sid -eq '.DEFAULT')      { continue }
        if ($sid -match '_Classes$')  { continue }
        if ($builtIn -contains $sid)  { continue }
        try {
            $objSid  = [System.Security.Principal.SecurityIdentifier]::new($sid)
            $account = $objSid.Translate([System.Security.Principal.NTAccount]).Value
        } catch {
            $account = $sid
        }
        [pscustomobject]@{ SID = $sid; Account = $account }
    }
}

function Set-PreloadForHive {
    <#
    .SYNOPSIS
        Ensures a KLID is at Preload position 1 under a given registry root.
        Optionally removes another KLID.
    #>
    param(
        [string]$RootPath,
        [string]$AddKlid,
        [string]$RemoveKlid,
        [string]$Label
    )
    $preloadPath = Join-Path $RootPath 'Keyboard Layout\Preload'
    if (-not (Test-Path -LiteralPath $preloadPath)) {
        [void](New-Item -Path $preloadPath -Force)
        Set-ItemProperty -LiteralPath $preloadPath -Name '1' -Value $AddKlid -Type String
        Write-Log "[$Label] Created Preload with 1=$AddKlid"
        return
    }
    $existing = @()
    foreach ($name in (Get-Item -LiteralPath $preloadPath).Property) {
        $existing += [pscustomobject]@{
            Name = $name
            Klid = (Get-ItemProperty -LiteralPath $preloadPath -Name $name).$name
        }
    }
    $existing = $existing | Sort-Object { [int]$_.Name }
    $beforeStr = ($existing | ForEach-Object { '{0}={1}' -f $_.Name, $_.Klid }) -join ', '

    $orderedKlids = @($AddKlid)
    foreach ($e in $existing) {
        if ($e.Klid -eq $AddKlid) { continue }
        if ($RemoveKlid -and $e.Klid -eq $RemoveKlid) { continue }
        if ($orderedKlids -notcontains $e.Klid) {
            $orderedKlids += $e.Klid
        }
    }

    $alreadyCorrect = $true
    if ($existing.Count -ne $orderedKlids.Count) { $alreadyCorrect = $false }
    else {
        for ($idx = 0; $idx -lt $orderedKlids.Count; $idx++) {
            if ($existing[$idx].Klid -ne $orderedKlids[$idx]) { $alreadyCorrect = $false; break }
        }
    }
    if ($alreadyCorrect) {
        Write-Log "[$Label] Preload already correct: $beforeStr"
        return
    }

    foreach ($e in $existing) {
        Remove-ItemProperty -LiteralPath $preloadPath -Name $e.Name -ErrorAction SilentlyContinue
    }
    $i = 1
    foreach ($klid in $orderedKlids) {
        Set-ItemProperty -LiteralPath $preloadPath -Name "$i" -Value $klid -Type String
        $i++
    }
    $afterStr = ($orderedKlids | ForEach-Object -Begin {$j=1} -Process { $r = '{0}={1}' -f $j,$_; $j++; $r }) -join ', '
    Write-Log "[$Label] Preload before: $beforeStr"
    Write-Log "[$Label] Preload after : $afterStr"
}
#endregion

#region --- Main ---
$ErrorActionPreference = 'Stop'
$exitCode = 0
$xmlPath  = $null

try {
    $isSystem     = Test-RunningAsSystem
    $copySettings = -not $SkipCopyToSystem.IsPresent

    Write-Log '=== Set-KeyboardLayout v1.6.0 ===' -Level SUCCESS
    Write-Log "Running as      : $(if ($isSystem) { 'SYSTEM' } else { [Environment]::UserName })"
    Write-Log "Add keyboard    : $AddInputLocale"
    Write-Log "Remove keyboard : $(if ($RemoveInputLocale) { $RemoveInputLocale } else { '(none)' })"
    Write-Log "Copy to default user / system: $copySettings"

    # Validate RemoveInputLocale format when not empty
    if ($RemoveInputLocale -and $RemoveInputLocale -notmatch '^[0-9a-fA-F]{4}:[0-9a-fA-F]{8}$') {
        Write-Log "-RemoveInputLocale '$RemoveInputLocale' is not a valid LLLL:KKKKKKKK format" -Level ERROR
        exit 1
    }

    # --- Step 1: Region settings (opt-in -- only when explicitly specified) ----------------
    # These cmdlets require admin. In USER context they would throw, so we gate them.
    $regionalRequested = $PSBoundParameters.ContainsKey('SystemLocale') -or
                         $PSBoundParameters.ContainsKey('Culture') -or
                         $PSBoundParameters.ContainsKey('GeoId')
    if ($regionalRequested -and -not $isSystem) {
        Write-Log 'Regional settings (SystemLocale/Culture/GeoId) require admin -- skipping in user context' -Level WARN
    } elseif ($isSystem) {
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
        if (-not $regionalRequested) {
            Write-Log 'No regional settings specified -- keyboard-only mode'
        }
    } else {
        Write-Log 'User context -- keyboard-only mode'
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

    # --- Step 4: Reorder Preload to set the new keyboard as default ----------------------
    # The Preload subkey holds REG_SZ values "1", "2", "3", ... whose data is the KLID
    # (Keyboard Layout Identifier, e.g. "00000807"). Position "1" is the default.
    # The XML add only appends to Preload; it never reorders. So we rewrite Preload
    # so the new KLID is at position 1 and existing KLIDs follow.
    $newKlid    = ($AddInputLocale -split ':')[1]   # "0807:00000807" -> "00000807"
    $removeKlid = if ($RemoveInputLocale) { ($RemoveInputLocale -split ':')[1] } else { '' }

    # 4a: Current user (HKCU) - always
    Set-PreloadForHive -RootPath 'HKCU:' -AddKlid $newKlid -RemoveKlid $removeKlid -Label 'HKCU'

    # 4b: Under SYSTEM, also patch every loaded real-user hive (HKU\<SID>)
    #     This is critical on W365 / Autopilot where the user profile already exists
    #     when the SYSTEM-context platform script runs -- the Default User copy alone
    #     has no effect on existing profiles.
    if ($isSystem) {
        $loadedUsers = @(Get-LoadedUserHive)
        if ($loadedUsers.Count -gt 0) {
            Write-Log "Patching Preload for $($loadedUsers.Count) loaded user hive(s)"
            foreach ($u in $loadedUsers) {
                $hiveRoot = "Registry::HKEY_USERS\$($u.SID)"
                Set-PreloadForHive -RootPath $hiveRoot -AddKlid $newKlid -RemoveKlid $removeKlid -Label $u.Account
            }
        } else {
            Write-Log 'No loaded user hives found (no users signed in)' -Level WARN
        }
    } else {
        # 4c: In USER context, promote the new keyboard in the active session.
        # Set-WinDefaultInputMethodOverride writes to HKCU but the running text
        # services framework picks it up live, so no sign-out is needed. This
        # cmdlet works reliably in USER context (it doesn't in SYSTEM, which is
        # why the SYSTEM branch uses the direct Preload reorder above instead).
        try {
            Write-Log "Setting default input method override to '$AddInputLocale' (active session)"
            Set-WinDefaultInputMethodOverride -InputTip $AddInputLocale
        } catch {
            Write-Log "Set-WinDefaultInputMethodOverride failed: $($_.Exception.Message)" -Level WARN
        }
    }
    # --- Step 5: Copy to Default User / system (welcome screen) ----------------------------
    # SYSTEM context only -- Copy-UserInternationalSettingsToSystem requires admin.
    if ($isSystem -and $copySettings) {
        Write-Log 'Copying international settings to Default User profile and system account'
        Copy-UserInternationalSettingsToSystem -WelcomeScreen $true -NewUser $true
    } elseif (-not $isSystem -and $copySettings) {
        Write-Log 'Skipping Default User / system copy (not running as SYSTEM)'
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
    if ($isSystem -and $copySettings) {
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
