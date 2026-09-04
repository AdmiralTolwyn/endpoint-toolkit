<#
.SYNOPSIS
    Silences or restores Windows notification, ringtone and alarm sounds without
    removing the notifications themselves.

.DESCRIPTION
    Blanks the WAV binding of each system sound event so the notification still
    appears but produces no audio. Windows represents "no sound" as an empty
    default value under the event's .Current key; several stock events such as
    Maximize and Minimize ship that way already, so this is the native mechanism
    rather than a workaround.

    Silencing rather than disabling matters operationally. Turning notifications
    off removes the visual alert an agent needs; the acoustic exposure comes from
    the audio alone. This targets the audio and leaves the toast intact.

    Sound events are grouped so a targeted change can be made:

      Notifications  Notification.Default, .IM, .Mail, .Reminder, .SMS,
                     .Proximity, MailBeep, SystemNotification, SystemAsterisk,
                     SystemExclamation, SystemHand, SystemQuestion, WindowsUAC,
                     DeviceConnect, DeviceDisconnect, DeviceFail and the
                     .Default event.
      Ringtones      Notification.Looping.Call1 through Call10. These are
                     Windows ringtones. A softphone that renders its own
                     ringtone through its own audio path is NOT covered here and
                     must be handled in that application.
      Alarms         Notification.Looping.Alarm1 through Alarm10.
      Devices        DeviceConnect, DeviceDisconnect and DeviceFail only, so
                     device-arrival sounds can be silenced separately from the
                     rest of Notifications. Relevant when endpoints (a headset,
                     a dock) are repeatedly removed and recreated, such as on a
                     remoted session, which otherwise announces itself audibly
                     on every reconnect.
      All            Every event bound under every registered application.

    Every original value is written to an identity-bound backup before any
    change. Restore uses app/event-relative paths, not arbitrary registry paths
    supplied by the JSON file.

    Deploying to a non-persistent session host: these are per-user values, so
    apply them through Group Policy Preferences, a logon script, or -DefaultUser
    against the golden image. A change made to one live session does not persist.

.PARAMETER Scope
    Which group of sound events to act on. Defaults to Notifications.

.PARAMETER BackupPath
    File to write the pre-change state to, or to read from when restoring.
    Defaults to NotificationSoundBackup-<COMPUTERNAME>-<timestamp>.json in the
    current directory.

.PARAMETER Restore
    Restore the values recorded in -BackupPath instead of silencing.

.PARAMETER DefaultUser
    Operate on the default user profile hive rather than the current user, so
    that new profiles inherit the setting. Intended for golden-image builds.
    Requires elevation and an image where no other process holds the hive.

.PARAMETER DefaultUserHivePath
    Path to the default user hive. Defaults to C:\Users\Default\NTUSER.DAT.

.PARAMETER LogPath
    Optional path to a run log. When supplied, the script appends a line for
    the run start and scope, the backup path, each binding changed or
    restored, the restore verification result, and any terminating error
    before it is rethrown. Omitted by default; no log is written. This is
    independent of -WhatIf/-Confirm, which govern the actual registry writes.

.EXAMPLE
    .\Set-NotificationSoundState.ps1 -WhatIf

    Show which sound events would be silenced, and their current WAV bindings,
    without changing anything.

.EXAMPLE
    .\Set-NotificationSoundState.ps1 -Scope Notifications -BackupPath C:\temp\sounds.json

    Silence notification sounds for the current user, recording the previous
    state so the change can be reversed exactly.

.EXAMPLE
    .\Set-NotificationSoundState.ps1 -Restore -BackupPath C:\temp\sounds.json

    Put every recorded sound binding back.

.EXAMPLE
    .\Set-NotificationSoundState.ps1 -Scope All -DefaultUser -BackupPath C:\temp\image.json

    Silence all sound events in the default user hive during an image build so
    every new profile starts silent.

.NOTES
    Author:   Anton Romanyuk
    Version:  1.0.0
    Requires: PowerShell 5.1. -DefaultUser requires elevation.

    Takes effect for subsequent sounds; no logoff or restart is needed, because
    the binding is read when the sound is played.

    This script deliberately does not touch toast delivery. Suppressing toasts
    is a Group Policy decision, not a scripted registry edit: use User
    Configuration, Start Menu and Taskbar, Notifications. Letting Group Policy
    write those values keeps them managed and reversible, and avoids hand-writing
    values whose exact names differ across builds.

    AppEvents\Schemes is an observed, documented registry mechanism for binding
    a WAV to a UI event; it is not asserted to be the only path Windows can use
    to produce a given sound. Silencing an event here is a mitigation and a
    diagnostic pathway test: if the symptom stops, the audio path is
    confirmed, but that is not by itself root-cause proof of why the sound was
    audible in the first place.
#>

#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateSet('Notifications', 'Ringtones', 'Alarms', 'Devices', 'All')]
    [string] $Scope = 'Notifications',

    [string] $BackupPath,

    [switch] $Restore,

    [switch] $DefaultUser,

    [string] $DefaultUserHivePath = 'C:\Users\Default\NTUSER.DAT',

    [string] $LogPath
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$TEMP_HIVE_NAME = 'AAH_DefaultUser'
$SILENT_VALUE = ''

# Event name patterns per scope. Anchored so that, for example, the ringtone
# pattern cannot also capture the alarm events.
$SCOPE_PATTERN = @{
    Notifications = '^(\.Default|MailBeep|SystemAsterisk|SystemExclamation|SystemNotification|SystemHand|SystemQuestion|WindowsUAC|DeviceConnect|DeviceDisconnect|DeviceFail|Notification\.(Default|IM|Mail|Reminder|SMS|Proximity))$'
    Ringtones     = '^Notification\.Looping\.Call\d*$'
    Alarms        = '^Notification\.Looping\.Alarm\d*$'
    Devices       = '^(DeviceConnect|DeviceDisconnect|DeviceFail)$'
    All           = '.'
}

<#
.SYNOPSIS
    Confirms the current process is running elevated.

.OUTPUTS
    System.Boolean
#>
function Test-Elevated {
    [CmdletBinding()]
    param()

    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    return ([Security.Principal.WindowsPrincipal] $identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

<#
.SYNOPSIS
    Appends a timestamped line to -LogPath, if one was supplied.

.PARAMETER Message
    Text to record.
#>
function Write-RunLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Message
    )

    if ([string]::IsNullOrWhiteSpace($Script:LogPath)) { return }

    $line = '{0:yyyy-MM-dd HH:mm:ss.fff}  {1}' -f (Get-Date), $Message
    Add-Content -LiteralPath $Script:LogPath -Value $line -Encoding UTF8
}

<#
.SYNOPSIS
    Maps a KeyPath built against $appsRoot to a (hive, subkey) pair usable
    with Microsoft.Win32.Registry.

.DESCRIPTION
    Get-ItemProperty/Set-ItemProperty coerce REG_EXPAND_SZ to REG_SZ and
    expand embedded environment variables, which loses the original value
    kind and unexpanded data. Direct .NET registry access preserves both, so
    every read/write of a sound event binding is routed through the hive and
    relative subkey this function derives from the same KeyPath used for
    -WhatIf messages and display.

.PARAMETER KeyPath
    Registry path either built against $appsRoot ('HKCU:\...' or
    'Registry::HKEY_USERS\<TEMP_HIVE_NAME>\...'), or the fully-qualified
    provider path PowerShell returns as .PSPath from Get-ChildItem/Get-Item
    ('Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\...' or
    '...HKEY_USERS\<TEMP_HIVE_NAME>\...'). Both forms occur in this script.

.OUTPUTS
    System.Management.Automation.PSCustomObject with Hive
    (Microsoft.Win32.RegistryKey) and SubKey (string) properties. The
    returned Hive is a well-known root (CurrentUser/Users) and is not meant
    to be disposed by the caller.
#>
function Get-RegistryKeyRef {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $KeyPath
    )

    # Normalise the 'HKCU:\' drive form to 'HKEY_CURRENT_USER\' so both that
    # form and the fully-qualified provider path PSPath returns
    # (...Registry::HKEY_CURRENT_USER\...) match with one pattern.
    $normalized = $KeyPath -replace '^HKCU:\\', 'HKEY_CURRENT_USER\'

    if ($normalized -match 'HKEY_CURRENT_USER\\(.+)$') {
        return [pscustomobject] @{
            Hive   = [Microsoft.Win32.Registry]::CurrentUser
            SubKey = $Matches[1]
        }
    }

    if ($normalized -match 'HKEY_USERS\\(.+)$') {
        return [pscustomobject] @{
            Hive   = [Microsoft.Win32.Registry]::Users
            SubKey = $Matches[1]
        }
    }

    throw "Unrecognised registry path form: $KeyPath"
}

<#
.SYNOPSIS
    Enumerates every sound event binding beneath an AppEvents root.

.DESCRIPTION
    Returns one record per application and event, carrying the current WAV path.
    An empty value means the event is already silent, which is how Windows
    itself represents stock silent events.

.PARAMETER AppsRoot
    Registry path of the AppEvents Schemes Apps key.

.PARAMETER Pattern
    Regular expression matched against the event name.

.OUTPUTS
    System.Management.Automation.PSCustomObject[]
#>
function Get-SoundEventBinding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $AppsRoot,
        [Parameter(Mandatory)] [string] $Pattern
    )

    $results = New-Object System.Collections.Generic.List[object]

    if (-not (Test-Path -LiteralPath $AppsRoot)) {
        return $results.ToArray()
    }

    foreach ($app in @(Get-ChildItem -LiteralPath $AppsRoot -ErrorAction SilentlyContinue)) {
        if ($Scope -ne 'All' -and $app.PSChildName -ne '.Default') { continue }

        foreach ($soundEvent in @(Get-ChildItem -LiteralPath $app.PSPath -ErrorAction SilentlyContinue)) {
            if ($soundEvent.PSChildName -notmatch $Pattern) { continue }

            $currentKey = Join-Path $soundEvent.PSPath '.Current'
            if (-not (Test-Path -LiteralPath $currentKey)) { continue }

            $value = $null
            $valueExists = $false
            $valueKind = 'String'
            $subKey = $null
            try {
                $ref = Get-RegistryKeyRef -KeyPath $currentKey
                $subKey = $ref.Hive.OpenSubKey($ref.SubKey)
                if ($null -ne $subKey) {
                    # GetValueKind throws when the default value itself is
                    # absent (the key exists but carries no (default)), which
                    # is how ValueExists is determined here.
                    $value = $subKey.GetValue('', $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                    $valueKind = [string] $subKey.GetValueKind('')
                    $valueExists = $true
                }
            } catch {
                Write-Verbose "No default value under $currentKey"
            } finally {
                if ($null -ne $subKey) { $subKey.Close() }
            }

            [void] $results.Add([pscustomobject] @{
                App          = $app.PSChildName
                Event        = $soundEvent.PSChildName
                KeyPath      = $currentKey
                CurrentValue = [string] $value
                ValueExists  = $valueExists
                ValueKind    = $valueKind
                IsSilent     = [string]::IsNullOrEmpty([string] $value)
            })
        }
    }

    return $results.ToArray()
}

<#
.SYNOPSIS
    Writes a sound event binding.

.PARAMETER KeyPath
    Registry path of the event's .Current key.

.PARAMETER Value
    WAV path to bind, or an empty string to silence the event.

.PARAMETER ValueKind
    Registry value kind to write the default value as: 'String' or
    'ExpandString'. Any other recorded kind falls back to 'String'. Defaults
    to 'String'.
#>
function Set-SoundEventBinding {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory)] [string] $KeyPath,
        [Parameter(Mandatory)] [AllowEmptyString()] [string] $Value,
        [bool] $ValueExists = $true,
        [string] $ValueKind = 'String'
    )

    if (-not (Test-Path -LiteralPath $KeyPath)) {
        Write-Warning "Key no longer present, skipped: $KeyPath"
        return
    }

    if (-not $ValueExists) {
        if ($PSCmdlet.ShouldProcess($KeyPath, 'Remove default value')) {
            Remove-ItemProperty -LiteralPath $KeyPath -Name '(default)' -ErrorAction SilentlyContinue
        }
    } else {
        $regType = if ($ValueKind -eq 'ExpandString') { 'ExpandString' } else { 'String' }
        if ($PSCmdlet.ShouldProcess($KeyPath, "Set default value to '$Value' ($regType)")) {
            Set-ItemProperty -LiteralPath $KeyPath -Name '(default)' -Value $Value -Type $regType -ErrorAction Stop
        }
    }
}

$hiveLoaded = $false

Write-RunLog -Message ("Run started. Scope={0} Restore={1} DefaultUser={2}" -f $Scope, [bool] $Restore, [bool] $DefaultUser)

try {
    if ($DefaultUser) {
        if (-not (Test-Elevated)) {
            throw 'Elevation is required to load the default user hive.'
        }

        if (-not (Test-Path -LiteralPath $DefaultUserHivePath -PathType Leaf)) {
            throw "Default user hive not found: $DefaultUserHivePath"
        }

        if ($WhatIfPreference) {
            Write-Host "What if: Would load and inspect default user hive '$DefaultUserHivePath'."
            Write-Host 'No hive was loaded and no values were changed.'
            return
        }

        Write-Verbose "Loading $DefaultUserHivePath as HKU\$TEMP_HIVE_NAME"
        $loadOutput = & reg.exe load "HKU\$TEMP_HIVE_NAME" $DefaultUserHivePath 2>&1
        if ($LASTEXITCODE -ne 0) {
            throw "Could not load the default user hive: $loadOutput"
        }

        $hiveLoaded = $true
        $appsRoot = "Registry::HKEY_USERS\$TEMP_HIVE_NAME\AppEvents\Schemes\Apps"
        if (-not (Test-Path -LiteralPath $appsRoot)) {
            throw "Loaded hive does not contain AppEvents: $DefaultUserHivePath"
        }
    } else {
        $appsRoot = 'HKCU:\AppEvents\Schemes\Apps'
    }

    if ($Restore) {
        if ([string]::IsNullOrWhiteSpace($BackupPath)) {
            throw '-Restore requires -BackupPath.'
        }

        if (-not (Test-Path -LiteralPath $BackupPath -PathType Leaf)) {
            throw "Backup file not found: $BackupPath"
        }

        # ConvertFrom-Json emits a JSON array as ONE object down the pipeline, so
        # @( ... | ConvertFrom-Json ) yields a single element containing the whole
        # array. Parse via -InputObject, then wrap, to get real records.
        $parsed = ConvertFrom-Json -InputObject (Get-Content -LiteralPath $BackupPath -Raw)
        if ($parsed.SchemaVersion -ne '2.1' -and $parsed.SchemaVersion -ne '2.0') {
            throw "Unsupported backup schema '$($parsed.SchemaVersion)'."
        }
        if ($parsed.SchemaVersion -eq '2.0') {
            Write-Warning "Backup schema '2.0' has no recorded ValueKind; every value will be restored as 'String'."
        }

        $expectedTarget = if ($DefaultUser) { [System.IO.Path]::GetFullPath($DefaultUserHivePath) } else { [Security.Principal.WindowsIdentity]::GetCurrent().User.Value }
        if ($parsed.Target -ne $expectedTarget) {
            throw "Backup target '$($parsed.Target)' does not match '$expectedTarget'."
        }

        $records = @($parsed.Records)

        # Validated up front, before any write, because App/Event feed a
        # Join-Path used to build the registry path to change.
        $namePattern = '^[^\\/]+$'
        foreach ($record in $records) {
            $appName = [string] $record.App
            $eventName = [string] $record.Event
            if ([string]::IsNullOrWhiteSpace($appName) -or $appName -notmatch $namePattern -or $appName -eq '.' -or $appName -eq '..') {
                throw "Backup contains an invalid App name: '$appName'"
            }
            if ([string]::IsNullOrWhiteSpace($eventName) -or $eventName -notmatch $namePattern -or $eventName -eq '.' -or $eventName -eq '..') {
                throw "Backup contains an invalid Event name: '$eventName'"
            }
        }

        Write-Host ("Restoring {0} sound event binding(s) from {1}" -f $records.Count, $BackupPath)
        Write-RunLog -Message ("Restoring {0} sound event binding(s) from {1}" -f $records.Count, $BackupPath)

        $restored = 0
        foreach ($record in $records) {
            $keyPath = Join-Path (Join-Path (Join-Path $appsRoot ([string] $record.App)) ([string] $record.Event)) '.Current'
            $recordValueKind = 'String'
            if ($record.PSObject.Properties.Match('ValueKind').Count -and -not [string]::IsNullOrWhiteSpace([string] $record.ValueKind)) {
                $recordValueKind = [string] $record.ValueKind
            }
            Set-SoundEventBinding -KeyPath $keyPath -Value ([string] $record.CurrentValue) -ValueExists ([bool] $record.ValueExists) -ValueKind $recordValueKind
            Write-RunLog -Message ("Restored {0}\{1} -> '{2}' ({3})" -f $record.App, $record.Event, [string] $record.CurrentValue, $recordValueKind)
            $restored++
        }

        if ($WhatIfPreference) {
            Write-Host ("Would restore {0} binding(s). Nothing was changed." -f $restored)
            return
        }

        # A restore that silently does nothing is the worst possible outcome, so
        # every value is read back and confirmed before reporting success.
        $mismatched = New-Object System.Collections.Generic.List[string]
        foreach ($record in $records) {
            $keyPath = Join-Path (Join-Path (Join-Path $appsRoot ([string] $record.App)) ([string] $record.Event)) '.Current'
            $recordValueKind = 'String'
            if ($record.PSObject.Properties.Match('ValueKind').Count -and -not [string]::IsNullOrWhiteSpace([string] $record.ValueKind)) {
                $recordValueKind = [string] $record.ValueKind
            }

            $actual = $null
            $actualExists = $false
            $actualKind = $null
            $subKey = $null
            try {
                $ref = Get-RegistryKeyRef -KeyPath $keyPath
                $subKey = $ref.Hive.OpenSubKey($ref.SubKey)
                if ($null -ne $subKey) {
                    $actual = $subKey.GetValue('', $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                    $actualKind = [string] $subKey.GetValueKind('')
                    $actualExists = $true
                }
            } catch {
                Write-Verbose "Could not read back $keyPath"
            } finally {
                if ($null -ne $subKey) { $subKey.Close() }
            }

            if ($actualExists -ne [bool] $record.ValueExists -or [string] $actual -ne [string] $record.CurrentValue -or ($actualExists -and $actualKind -ne $recordValueKind)) {
                [void] $mismatched.Add('{0}\{1}' -f $record.App, $record.Event)
            }
        }

        if ($mismatched.Count -gt 0) {
            Write-Warning ("{0} binding(s) did not restore: {1}" -f $mismatched.Count, ($mismatched -join ', '))
            Write-RunLog -Message ("Restore verification: {0} binding(s) did not restore: {1}" -f $mismatched.Count, ($mismatched -join ', '))
        } else {
            Write-Host ("Restored and verified {0} binding(s)." -f $restored)
            Write-RunLog -Message ("Restore verification: {0} binding(s) restored and verified." -f $restored)
        }

        return
    }

    $pattern = $SCOPE_PATTERN[$Scope]
    $bindings = @(Get-SoundEventBinding -AppsRoot $appsRoot -Pattern $pattern)

    if ($bindings.Count -eq 0) {
        Write-Warning "No sound events matched scope '$Scope' under $appsRoot."
        return
    }

    $audible = @($bindings | Where-Object { -not $_.IsSilent })

    Write-Host ''
    Write-Host ("Scope        : {0}" -f $Scope)
    Write-Host ("Target hive  : {0}" -f $appsRoot)
    Write-Host ("Matched      : {0} event(s), {1} currently audible" -f $bindings.Count, $audible.Count)
    Write-Host ''

    $bindings | Sort-Object App, Event | Format-Table -AutoSize App, Event, IsSilent, CurrentValue | Out-Host

    if ($audible.Count -eq 0) {
        Write-Host 'Every matched event is already silent. Nothing to do.'
        return
    }

    if ([string]::IsNullOrWhiteSpace($BackupPath)) {
        $BackupPath = Join-Path (Get-Location).Path ('NotificationSoundBackup-{0}-{1}.json' -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMdd-HHmmss'))
    }

    # The backup covers every matched event, not just the audible ones, so a
    # restore reproduces the original state exactly rather than approximately.
    $backupWritten = $false
    if ($PSCmdlet.ShouldProcess($BackupPath, 'Write pre-change backup')) {
        $targetIdentity = if ($DefaultUser) { [System.IO.Path]::GetFullPath($DefaultUserHivePath) } else { [Security.Principal.WindowsIdentity]::GetCurrent().User.Value }
        $backup = [pscustomobject] @{
            SchemaVersion = '2.1'
            CreatedUtc    = (Get-Date).ToUniversalTime().ToString('o')
            Target        = $targetIdentity
            Scope         = $Scope
            Records       = @($bindings | Select-Object App, Event, CurrentValue, ValueExists, ValueKind)
        }
        $json = $backup | ConvertTo-Json -Depth 5
        [System.IO.File]::WriteAllText($BackupPath, $json, [System.Text.UTF8Encoding]::new($true))

        # Fail closed: nothing is silenced unless the backup on disk reads
        # back as the same backup that was just written.
        $verify = ConvertFrom-Json -InputObject (Get-Content -LiteralPath $BackupPath -Raw)
        $verifyRecordCount = @($verify.Records).Count
        if ($verify.SchemaVersion -ne $backup.SchemaVersion -or $verify.Target -ne $backup.Target -or $verifyRecordCount -ne $bindings.Count) {
            throw "Backup verification failed for $BackupPath before any change was made."
        }

        $backupWritten = $true
        Write-Host "Backup written to $BackupPath"
        Write-RunLog -Message "Backup written and verified: $BackupPath ($verifyRecordCount record(s))"
    }

    if (-not $backupWritten -and -not $WhatIfPreference) {
        throw 'No settings were changed because the backup was not written.'
    }

    $changed = 0
    foreach ($binding in $audible) {
        Set-SoundEventBinding -KeyPath $binding.KeyPath -Value $SILENT_VALUE -ValueKind $binding.ValueKind
        Write-RunLog -Message ("Silenced {0}\{1} (was '{2}', {3})" -f $binding.App, $binding.Event, $binding.CurrentValue, $binding.ValueKind)
        $changed++
    }

    Write-Host ''

    if ($WhatIfPreference) {
        Write-Host ("Would silence {0} sound event(s). Nothing was changed." -f $changed)
        return
    }

    Write-Host ("Silenced {0} sound event(s)." -f $changed)
    Write-Host ("Reverse with: .\Set-NotificationSoundState.ps1 -Restore -BackupPath `"{0}`"" -f $BackupPath)
    Write-RunLog -Message "Silenced $changed sound event(s)."
} catch {
    Write-RunLog -Message "ERROR: $($_.Exception.Message)"
    throw
} finally {
    if ($hiveLoaded) {
        # The hive cannot unload while PowerShell still holds a handle from the
        # enumeration above, so collect before releasing it.
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        $unloadOutput = & reg.exe unload "HKU\$TEMP_HIVE_NAME" 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "Default user hive did not unload cleanly: $unloadOutput"
        } else {
            Write-Verbose 'Default user hive unloaded.'
        }
    }
}
