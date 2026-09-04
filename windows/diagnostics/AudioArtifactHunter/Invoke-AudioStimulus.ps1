<#
.SYNOPSIS
    Fires controlled Path 1 (Windows audio on the VDA/client, through HDX) audio
    stimuli so an intermittent loud transient can be provoked and correlated
    with the endpoint-state monitor, the loopback recorder and client-side
    logs.

.DESCRIPTION
    This is Path 1 only: Windows audio rendered on the machine this script runs
    on. It cannot fire a Teams chat/call ding (endpoint HdxRtcEngine) or a Webex
    ringtone (Cisco Media Engine); those are Path 2/Path 3 events that live in a
    different process's audio stack and cannot be scripted from here. Use
    -IncludeManualCue with 'ManualCue' in -Sequence to log an operator-driven
    window for those instead - the script prints a prompt and waits briefly for
    Enter, it does not send the Teams/Webex event itself.

    Six stimulus types are supported, each implemented as a small function
    that returns a result object carrying Result/HResult/Error rather than
    throwing, so one failing stimulus never aborts the run:

      ToastDefaultSound  WinRT toast with the default notification sound
                         (ms-winsoundevent:Notification.Default).
      ToastSilent        Same toast, with <audio silent="true"/>: exercises the
                         visual/WNS machinery without a sound.
      DirectWav          Plays a WAV synchronously via System.Media.SoundPlayer,
                         bypassing WinRT toasts entirely.
      SystemSound        [System.Media.SystemSounds]::Asterisk.Play(), the
                         classic MessageBeep path.
      MailBeep           winmm PlaySound of the AppEvents "MailBeep" alias, the
                         Windows "New Mail Notification" event that classic
                         Outlook's "Play a sound" on message arrival uses.
                         Fires that binding without Outlook or a mailbox.
                         SND_NODEFAULT is set, so a blanked binding plays
                         nothing and the row reads NotPlayed rather than
                         falling back to the default beep.
      ManualCue          Fires nothing. Prints a prompt so an operator can
                         trigger an external (Teams/Webex) stimulus by hand and
                         logs a Cue row. Only used when -IncludeManualCue is
                         given.

    A ToastDefaultSound failure is expected and informative, not a bug, when
    the Windows Push Notification service (WpnService) or the per-user
    WpnUserService instance is disabled on this host - toasts (and their Path 1
    sound) cannot fire without it. This script never changes that service, or
    any other setting; it only records the service state read-only in
    session.json so a run can be interpreted against the host state it ran
    under. It changes no volume, mute, default device, notification-sound
    binding or registry value, and downloads nothing.

    Each cycle walks -Sequence in order. Before every stimulus the script
    sleeps -IdleSeconds (plus up to -JitterSeconds of random extra idle, since
    transients of this kind are typically reported after idle periods),
    polling the stop file every 500 ms so a run can always be ended
    promptly. Immediately before and ~1 second after firing, the script reads
    the default render endpoint's identity, device state, master volume and
    mute for both the Console and Multimedia roles
    ([AudioArtifactHunter.EndpointStateReader]::ReadDefaultRenderEndpoint), so
    a change coincident with the stimulus is visible in the row itself. This is
    read-only endpoint state, the same mechanism used by
    Start-AudioEndpointStateMonitor.ps1; it does not confirm acoustic output
    reached a headset.

    With -Interactive, the script also polls the keyboard during each idle wait
    (console host only; silently skipped everywhere else, e.g. a scheduled
    task) and logs a standalone OperatorHeardBang row, timestamped in UTC, the
    moment the operator presses B - a way to mark a real, unprompted artefact
    without waiting for the next scripted stimulus.

    Absence of a bang during a run is not proof of absence: an intermittent
    artefact is intermittent, and a short run at a given cadence samples
    only a fraction of the conditions that could produce it.

    Correlating a row with other collectors is a matter of matching
    TimestampUtc/LocalTime (UTC recommended) against:
      - levels.csv / triggers.csv from Start-AudioLoopbackRecorder.ps1 (did the
        WASAPI loopback tap see a peak at this time?),
      - endpoint-volume.csv from the same recorder, or the CSV produced by
        Start-AudioEndpointStateMonitor.ps1 (did master volume, mute, device
        state or endpoint identity change at this time?),
      - client-side logs (Citrix/Teams/Webex/Windows Event Log) captured by
        other means for the same UTC window.
    -RecorderOutputDirectory, when given, also drops a MARK-INCIDENT.txt marker
    in the recorder's output directory a few seconds after each stimulus so the
    recorder itself preserves its rolling window for that moment; see that
    parameter's help for the disk cost of doing this on every stimulus.

    HEARING SAFETY: this script does not, and by design cannot, lower system
    volume, so a stimulus plays at whatever level is currently configured.
    -AcknowledgeHearingSafety is required before anything is played; run the
    first pass with the headset off the ear (or not worn) until the operator
    knows what level to expect.

    Designed for non-interactive use (a scheduled task, or a console session an
    operator walks away from): -OutputDirectory and -AcknowledgeHearingSafety
    are not marked Mandatory, so PowerShell never blocks on a missing-parameter
    prompt; a missing required value throws immediately instead. Create the
    file named by -StopFileName (default STOP-STIMULUS.txt) inside
    -OutputDirectory to end a run cleanly - the file is deleted and the run
    stops at the next poll, at most 1 second later plus whatever is left of an
    in-flight stimulus.

.PARAMETER OutputDirectory
    Root directory for stimulus-log.csv, stimulus.log and session.json.
    Created if absent. Required (an empty value throws); not marked Mandatory
    so a scheduled task or -ListPlan run never blocks on a prompt.

.PARAMETER Sequence
    Ordered list of stimulus type names to run once per cycle. Defaults to
    @('ToastDefaultSound','MailBeep','DirectWav','SystemSound','ToastSilent').
    Each value must be one of ToastDefaultSound, ToastSilent, DirectWav,
    SystemSound, MailBeep, ManualCue. Including 'ManualCue' also requires
    -IncludeManualCue. ToastDefaultSound plus MailBeep together reproduce what
    a classic Outlook desktop alert with sound does on Path 1.

.PARAMETER Cycles
    Number of times to repeat -Sequence. 1 to 1000. Defaults to 3.

.PARAMETER IdleSeconds
    Idle time to sleep immediately before each stimulus, because transients
    of this kind are typically reported after an idle period. 0 to 3600.
    Defaults to 120. Combine a small value with -JitterSeconds to sweep a
    range of idle lengths; IdleSecondsBefore in each row then lets a bang
    rate be tabulated against idle length.

.PARAMETER JitterSeconds
    Additional random idle, 0 to this many seconds, added on top of
    -IdleSeconds before each stimulus so consecutive stimuli do not land at an
    exactly predictable cadence. 0 to 600. Defaults to 0 (no jitter).

.PARAMETER WavPath
    WAV file for the DirectWav stimulus. Defaults to the WAV currently bound to
    HKCU AppEvents .Default\Notification.Default\.Current, expanded, falling
    back to %SystemRoot%\media\Windows Notify System Generic.wav. Whichever
    path is used, it must exist at fire time or the stimulus is recorded as
    Skipped rather than attempted.

.PARAMETER ToastAppId
    AUMID used to create the toast notifier for ToastDefaultSound and
    ToastSilent. Defaults to
    '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
    (the well-known Windows PowerShell AUMID).

.PARAMETER IncludeManualCue
    Required in addition to including 'ManualCue' in -Sequence. Without it,
    'ManualCue' in -Sequence throws immediately rather than silently doing
    nothing.

.PARAMETER RecorderOutputDirectory
    When given, -MarkerDelaySeconds after each stimulus fires, a
    MARK-INCIDENT.txt marker is written into this directory so a concurrently
    running Start-AudioLoopbackRecorder.ps1 preserves its rolling window for
    that moment (see that script's -DESCRIPTION for how the marker is
    consumed). Disk cost: each marker copies the recorder's entire rolling
    window (roughly 11 MB per minute of window at that script's default
    format) into its preserved\ directory, so firing many stimuli with a short
    -IdleSeconds against a large rolling window can consume disk quickly;
    watch that script's -MaxPreservedMegabytes retention limit. If this
    directory does not exist, the marker is skipped for that stimulus and a
    warning is logged rather than the run failing.

.PARAMETER MarkerDelaySeconds
    Delay after firing before writing the -RecorderOutputDirectory marker, so
    the recorder's rolling window still contains the stimulus once the marker
    triggers preservation. Defaults to 5.

.PARAMETER Interactive
    During each idle wait, poll the keyboard for the B key (console host only;
    silently skipped everywhere else) and log a standalone OperatorHeardBang
    row, timestamped in UTC, the moment it is pressed.

.PARAMETER StopFileName
    Name of a marker file that, when created directly in -OutputDirectory,
    ends the run cleanly at the next poll (at most 500 ms during an idle wait).
    The file is deleted once detected. Defaults to STOP-STIMULUS.txt.

.PARAMETER ListPlan
    Print the planned cycle/sequence schedule and exit without playing, WAV
    testing, endpoint reads or writing any file.

.PARAMETER AcknowledgeHearingSafety
    Required before anything is played (not required for -ListPlan). This
    script never lowers system, endpoint or application volume; a stimulus
    plays at whatever level is already configured. Run the first pass with the
    headset off the ear, or not worn, until the operator knows what level to
    expect. Missing this throws immediately rather than prompting.

.EXAMPLE
    .\Invoke-AudioStimulus.ps1 -ListPlan -OutputDirectory C:\Temp\Stim -Cycles 2 -IdleSeconds 30

    Preview the schedule for two cycles of the default sequence with no
    stimuli fired and nothing written to disk.

.EXAMPLE
    .\Invoke-AudioStimulus.ps1 -OutputDirectory C:\ProgramData\AudioStimulus -AcknowledgeHearingSafety

    Run the default sequence for 3 cycles with a 120 second idle before each
    stimulus.

.EXAMPLE
    .\Invoke-AudioStimulus.ps1 -OutputDirectory C:\ProgramData\AudioStimulus -Sequence ToastDefaultSound,DirectWav,SystemSound,ToastSilent,ManualCue -IncludeManualCue -RecorderOutputDirectory C:\ProgramData\AudioEvidence -Interactive -AcknowledgeHearingSafety

    Run the full sequence including an operator-cued Teams/Webex window,
    marking the loopback recorder's rolling window after each stimulus, while
    also watching for an operator-reported bang during idle periods.

.EXAMPLE
    .\Invoke-AudioStimulus.ps1 -OutputDirectory C:\ProgramData\AudioStimulus -Cycles 200 -IdleSeconds 300 -JitterSeconds 120 -AcknowledgeHearingSafety

    Long unattended run (scheduled task) with widely jittered idle periods.
    Create C:\ProgramData\AudioStimulus\STOP-STIMULUS.txt to stop it cleanly.

.NOTES
    Author:   Anton Romanyuk
    Version:  1.0.0
    Requires: PowerShell 5.1 and AudioLoopbackCapture.cs alongside this script
              (for [AudioArtifactHunter.EndpointStateReader]).

    ToastDefaultSound and ToastSilent use the Windows Runtime toast APIs
    ([Windows.UI.Notifications.ToastNotificationManager] /
    [Windows.Data.Xml.Dom.XmlDocument]) via Windows PowerShell's WinRT
    projection. This is Windows PowerShell 5.1 only; it is not expected to work
    under pwsh (PowerShell 7+), and a failure there is caught and recorded like
    any other ToastDefaultSound/ToastSilent failure rather than terminating
    the run.

    DirectWav and SystemSound calls are synchronous
    (System.Media.SoundPlayer.PlaySync / SystemSounds.*.Play); the ~1 second
    settle before the "after" endpoint snapshot is a fixed wait, not a measured
    playback duration.

    This script never disables, enables, starts or stops WpnService or any
    WpnUserService* instance; it only reads and records their Status/StartType
    in session.json, once, at run start.
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [string] $OutputDirectory,

    [ValidateSet('ToastDefaultSound', 'ToastSilent', 'DirectWav', 'SystemSound', 'MailBeep', 'ManualCue')]
    [string[]] $Sequence = @('ToastDefaultSound', 'MailBeep', 'DirectWav', 'SystemSound', 'ToastSilent'),

    [ValidateRange(1, 1000)]
    [int] $Cycles = 3,

    [ValidateRange(0, 3600)]
    [int] $IdleSeconds = 120,

    [ValidateRange(0, 600)]
    [int] $JitterSeconds = 0,

    [string] $WavPath,

    [string] $ToastAppId = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe',

    [switch] $IncludeManualCue,

    [string] $RecorderOutputDirectory,

    [ValidateRange(0, 3600)]
    [int] $MarkerDelaySeconds = 5,

    [switch] $Interactive,

    [string] $StopFileName = 'STOP-STIMULUS.txt',

    [switch] $ListPlan,

    [switch] $AcknowledgeHearingSafety
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$MARKER_NAME = 'MARK-INCIDENT.txt'
$CSV_HEADER = 'TimestampUtc,LocalTime,ComputerName,UserSid,SessionId,RunId,Cycle,Sequence,Stimulus,Asset,Result,HResult,Error,IdleSecondsBefore,EndpointIdBefore,VolumeBefore,MutedBefore,EndpointIdAfter,VolumeAfter,MutedAfter,DeviceStateBefore,DeviceStateAfter,EndpointIdMultimediaBefore,EndpointIdMultimediaAfter'

$script:consoleAvailable = $true

<#
.SYNOPSIS
    Appends one timestamped line to the stimulus run log and echoes it to the
    console.

.DESCRIPTION
    Every status line is routed through here so a run with no attached console
    (a scheduled task) still leaves a record on disk. A lost log line must
    never stop the run, so a write failure is swallowed.
#>
function Write-StimulusLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $Level,
        [AllowEmptyString()] [string] $Message
    )

    $line = '{0} {1} {2}' -f (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ'), $Level, $Message

    try {
        [System.IO.File]::AppendAllText($Path, $line + "`r`n", [System.Text.Encoding]::UTF8)
    } catch {
        # A lost log line must never stop the run.
    }

    Write-Host $Message
}

<#
.SYNOPSIS
    Reads an exception's HResult as an '0x{X8}' string, falling back to the
    inner exception when the outer one carries no useful value.
#>
function Get-ExceptionHResultText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [System.Exception] $ExceptionObject
    )

    $hresultValue = 0
    try { $hresultValue = $ExceptionObject.HResult } catch { $hresultValue = 0 }

    if ($hresultValue -eq 0 -and $null -ne $ExceptionObject.InnerException) {
        try { $hresultValue = $ExceptionObject.InnerException.HResult } catch { }
    }

    return ('0x{0:X8}' -f $hresultValue)
}

<#
.SYNOPSIS
    Resolves the WAV bound to the current user's default notification sound.

.DESCRIPTION
    Reads HKCU\AppEvents\Schemes\Apps\.Default\Notification.Default\.Current's
    default value with DoNotExpandEnvironmentNames, then expands it. Falls
    back to %SystemRoot%\media\Windows Notify System Generic.wav when the
    binding is absent, unreadable, or does not point at a file that exists.
#>
function Resolve-DefaultNotificationWav {
    [CmdletBinding()]
    param()

    $expandedPath = $null
    $subKey = $null
    try {
        $subKey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('AppEvents\Schemes\Apps\.Default\Notification.Default\.Current')
        if ($null -ne $subKey) {
            $rawValue = $subKey.GetValue('', $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
            if (-not [string]::IsNullOrWhiteSpace($rawValue)) {
                $expandedPath = [System.Environment]::ExpandEnvironmentVariables($rawValue)
            }
        }
    } catch {
        Write-Verbose "Could not read default notification WAV binding: $($_.Exception.Message)"
    } finally {
        if ($null -ne $subKey) { $subKey.Close() }
    }

    if (-not [string]::IsNullOrWhiteSpace($expandedPath) -and (Test-Path -LiteralPath $expandedPath -PathType Leaf)) {
        return $expandedPath
    }

    return [System.Environment]::ExpandEnvironmentVariables('%SystemRoot%\media\Windows Notify System Generic.wav')
}

<#
.SYNOPSIS
    Fires a WinRT toast carrying the Path 1 default notification sound, or a
    silent variant, via Windows PowerShell's WinRT projection.

.DESCRIPTION
    Never throws: any failure (WinRT unavailable, WNS/push service disabled,
    AUMID rejected, etc.) is caught and returned as a Failed result carrying
    the HResult and message, because a disabled push service producing this
    failure is itself an expected, informative outcome that documents host
    state - not a defect in this script.

.OUTPUTS
    PSCustomObject with Result ('Played' or 'Failed'), HResult, Error.
#>
function Send-ToastStimulus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $AppId,
        [Parameter(Mandatory)] [string] $SequenceLabel,
        [Parameter(Mandatory)] [datetime] $NowUtc,
        [switch] $Silent
    )

    $audioXml = '<audio src="ms-winsoundevent:Notification.Default"/>'
    if ($Silent) { $audioXml = '<audio silent="true"/>' }

    $toastXml = '<toast><visual><binding template="ToastGeneric"><text>AudioArtifactHunter stimulus</text><text>{0} {1}</text></binding></visual>{2}</toast>' -f $SequenceLabel, $NowUtc.ToString('o'), $audioXml

    try {
        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
        [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom, ContentType = WindowsRuntime] | Out-Null

        $xmlDocument = New-Object Windows.Data.Xml.Dom.XmlDocument
        $xmlDocument.LoadXml($toastXml)
        $toast = New-Object Windows.UI.Notifications.ToastNotification $xmlDocument
        $notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($AppId)
        $notifier.Show($toast)

        return [pscustomobject] @{
            Result  = 'Played'
            HResult = '0x00000000'
            Error   = ''
        }
    } catch {
        return [pscustomobject] @{
            Result  = 'Failed'
            HResult = (Get-ExceptionHResultText -ExceptionObject $_.Exception)
            Error   = $_.Exception.Message
        }
    }
}

<#
.SYNOPSIS
    Plays a WAV file synchronously via System.Media.SoundPlayer.

.OUTPUTS
    PSCustomObject with Result ('Played', 'Failed' or 'Skipped'), HResult, Error.
#>
function Invoke-DirectWavStimulus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $WavFilePath
    )

    if (-not (Test-Path -LiteralPath $WavFilePath -PathType Leaf)) {
        return [pscustomobject] @{
            Result  = 'Skipped'
            HResult = ''
            Error   = "WAV file not found: $WavFilePath"
        }
    }

    $player = $null
    try {
        $player = New-Object System.Media.SoundPlayer($WavFilePath)
        $player.PlaySync()

        return [pscustomobject] @{
            Result  = 'Played'
            HResult = '0x00000000'
            Error   = ''
        }
    } catch {
        return [pscustomobject] @{
            Result  = 'Failed'
            HResult = (Get-ExceptionHResultText -ExceptionObject $_.Exception)
            Error   = $_.Exception.Message
        }
    } finally {
        if ($null -ne $player) { $player.Dispose() }
    }
}

<#
.SYNOPSIS
    Plays the Asterisk system sound (the MessageBeep path).

.OUTPUTS
    PSCustomObject with Result ('Played' or 'Failed'), HResult, Error.
#>
function Invoke-SystemSoundStimulus {
    [CmdletBinding()]
    param()

    try {
        [System.Media.SystemSounds]::Asterisk.Play()

        return [pscustomobject] @{
            Result  = 'Played'
            HResult = '0x00000000'
            Error   = ''
        }
    } catch {
        return [pscustomobject] @{
            Result  = 'Failed'
            HResult = (Get-ExceptionHResultText -ExceptionObject $_.Exception)
            Error   = $_.Exception.Message
        }
    }
}

<#
.SYNOPSIS
    Plays the AppEvents "MailBeep" alias through winmm PlaySound.

.DESCRIPTION
    MailBeep is the Windows "New Mail Notification" sound event, the binding
    classic Outlook's "Play a sound" on message arrival uses. Playing the
    alias synchronously reproduces that Path 1 producer without Outlook.
    SND_ALIAS | SND_SYNC | SND_NODEFAULT: a blanked or missing binding plays
    nothing and returns false instead of falling back to the default beep,
    which makes a GPP-blanked host visible in the log as NotPlayed.

.OUTPUTS
    PSCustomObject with Result ('Played', 'NotPlayed' or 'Failed'), HResult, Error.
#>
function Invoke-MailBeepStimulus {
    [CmdletBinding()]
    param()

    try {
        if (-not ('AudioArtifactHunter.WinMm' -as [type])) {
            Add-Type -Namespace 'AudioArtifactHunter' -Name 'WinMm' -MemberDefinition @'
[DllImport("winmm.dll", CharSet = CharSet.Unicode, SetLastError = true)]
public static extern bool PlaySound(string pszSound, IntPtr hmod, uint fdwSound);
'@ -ErrorAction Stop
        }

        # SND_SYNC 0x0000 | SND_NODEFAULT 0x0002 | SND_ALIAS 0x00010000
        $played = [AudioArtifactHunter.WinMm]::PlaySound('MailBeep', [IntPtr]::Zero, 0x00010002)
        if ($played) {
            return [pscustomobject] @{ Result = 'Played'; HResult = '0x00000000'; Error = '' }
        }

        return [pscustomobject] @{
            Result  = 'NotPlayed'
            HResult = '0x00000000'
            Error   = 'PlaySound returned false: MailBeep binding blank, file missing, or no render device.'
        }
    } catch {
        return [pscustomobject] @{
            Result  = 'Failed'
            HResult = (Get-ExceptionHResultText -ExceptionObject $_.Exception)
            Error   = $_.Exception.Message
        }
    }
}

<#
.SYNOPSIS
    Prompts an operator to trigger an external (Teams/Webex) stimulus by hand
    and waits briefly for Enter, or a bounded timeout, before returning.

.DESCRIPTION
    Fires nothing itself - Teams engine dings and Webex ringtones cannot be
    scripted from this host. Console key detection is best-effort: when the
    console is unavailable (no attached console, redirected input, a scheduled
    task) this silently falls back to just waiting out the fixed timeout.

.OUTPUTS
    PSCustomObject with Result ('Cue'), HResult (''), Error ('').
#>
function Invoke-ManualCueStimulus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $StopFilePath
    )

    Write-Host 'Now trigger the external stimulus (Teams chat message / Webex call) and press Enter or wait'

    $waitSeconds = 10
    $deadlineUtc = (Get-Date).ToUniversalTime().AddSeconds($waitSeconds)

    while ((Get-Date).ToUniversalTime() -lt $deadlineUtc) {
        if (Test-Path -LiteralPath $StopFilePath -PathType Leaf) { break }

        if ($script:consoleAvailable) {
            try {
                if ([Console]::KeyAvailable) {
                    $keyInfo = [Console]::ReadKey($true)
                    if ($keyInfo.Key -eq [ConsoleKey]::Enter) { break }
                }
            } catch {
                $script:consoleAvailable = $false
            }
        }

        Start-Sleep -Milliseconds 250
    }

    return [pscustomobject] @{
        Result  = 'Cue'
        HResult = ''
        Error   = ''
    }
}

<#
.SYNOPSIS
    Reads the default render endpoint's Console and Multimedia role state.

.DESCRIPTION
    Wraps two [AudioArtifactHunter.EndpointStateReader]::ReadDefaultRenderEndpoint
    calls (role 0 Console, role 1 Multimedia) into one snapshot used for both
    the "before" and "after" columns of a stimulus row.

.OUTPUTS
    PSCustomObject with EndpointId, DeviceState, Volume (formatted scalar
    string), Muted, EndpointIdMultimedia.
#>
function Get-EndpointSnapshotPair {
    [CmdletBinding()]
    param()

    $consoleSnapshot = [AudioArtifactHunter.EndpointStateReader]::ReadDefaultRenderEndpoint(0)
    $multimediaSnapshot = [AudioArtifactHunter.EndpointStateReader]::ReadDefaultRenderEndpoint(1)

    return [pscustomobject] @{
        EndpointId           = $consoleSnapshot.EndpointId
        DeviceState          = $consoleSnapshot.DeviceState
        Volume               = $consoleSnapshot.MasterVolumeScalar.ToString('F6', [Globalization.CultureInfo]::InvariantCulture)
        Muted                = $consoleSnapshot.Muted
        EndpointIdMultimedia = $multimediaSnapshot.EndpointId
    }
}

<#
.SYNOPSIS
    Builds one stimulus-log.csv row.
#>
function New-StimulusRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [int] $Cycle,
        [Parameter(Mandatory)] [int] $SequenceIndex,
        [Parameter(Mandatory)] [AllowEmptyString()] [string] $Stimulus,
        [Parameter(Mandatory)] [AllowEmptyString()] [string] $Asset,
        [Parameter(Mandatory)] [string] $Result,
        [Parameter(Mandatory)] [AllowEmptyString()] [string] $HResult,
        [Parameter(Mandatory)] [AllowEmptyString()] [string] $Err,
        [Parameter(Mandatory)] [int] $IdleSecondsBefore,
        [Parameter(Mandatory)] $Before,
        [Parameter(Mandatory)] $After,
        [Parameter(Mandatory)] [datetime] $NowUtc
    )

    return [pscustomobject] [ordered] @{
        TimestampUtc               = $NowUtc.ToString('o')
        LocalTime                  = $NowUtc.ToLocalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffzzz')
        ComputerName               = $env:COMPUTERNAME
        UserSid                    = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        SessionId                  = (Get-Process -Id $PID).SessionId
        RunId                      = $script:runId
        Cycle                      = $Cycle
        Sequence                   = $SequenceIndex
        Stimulus                   = $Stimulus
        Asset                      = $Asset
        Result                     = $Result
        HResult                    = $HResult
        Error                      = $Err
        IdleSecondsBefore          = $IdleSecondsBefore
        EndpointIdBefore           = $Before.EndpointId
        VolumeBefore               = $Before.Volume
        MutedBefore                = $Before.Muted
        EndpointIdAfter            = $After.EndpointId
        VolumeAfter                = $After.Volume
        MutedAfter                 = $After.Muted
        DeviceStateBefore          = $Before.DeviceState
        DeviceStateAfter           = $After.DeviceState
        EndpointIdMultimediaBefore = $Before.EndpointIdMultimedia
        EndpointIdMultimediaAfter  = $After.EndpointIdMultimedia
    }
}

<#
.SYNOPSIS
    Sleeps up to $Seconds, polling the stop file every 500 ms and, with
    -Interactive, the keyboard for an operator-reported bang (key B).

.DESCRIPTION
    A detected bang is logged immediately as a standalone OperatorHeardBang
    row and does not interrupt the idle wait itself. Console key detection is
    best-effort and permanently disabled for the rest of the run the first
    time it throws (no console attached, redirected input, etc.).

.OUTPUTS
    $true if the full idle period elapsed; $false if the stop file appeared.
#>
function Wait-IdlePeriod {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [int] $Seconds,
        [Parameter(Mandatory)] [string] $StopFilePath,
        [switch] $Interactive,
        [Parameter(Mandatory)] [string] $CsvPath,
        [Parameter(Mandatory)] [string] $LogPath,
        [Parameter(Mandatory)] [int] $Cycle,
        [Parameter(Mandatory)] [int] $SequenceIndex,
        [AllowEmptyString()] [string] $NextStimulus = ''
    )

    $deadlineUtc = (Get-Date).ToUniversalTime().AddSeconds($Seconds)
    $activity = 'AudioArtifactHunter stimulus: idle wait before {0} (cycle {1}, seq {2})' -f $NextStimulus, $Cycle, $SequenceIndex
    Write-Verbose ("Idle wait {0}s before {1}; fires at about {2:HH:mm:ss} local. Create the stop file to end the run." -f $Seconds, $NextStimulus, $deadlineUtc.ToLocalTime())
    $lastVerboseUtc = (Get-Date).ToUniversalTime()

    while ((Get-Date).ToUniversalTime() -lt $deadlineUtc) {
        if (Test-Path -LiteralPath $StopFilePath -PathType Leaf) {
            Write-Progress -Activity $activity -Completed
            return $false
        }

        $nowLoopUtc = (Get-Date).ToUniversalTime()
        $remaining = [int] [Math]::Ceiling(($deadlineUtc - $nowLoopUtc).TotalSeconds)
        if ($Seconds -gt 0) {
            $percent = [int] [Math]::Min(100, [Math]::Max(0, (($Seconds - $remaining) * 100 / $Seconds)))
            Write-Progress -Activity $activity -Status ("{0}s remaining" -f $remaining) -PercentComplete $percent
        }

        if (($nowLoopUtc - $lastVerboseUtc).TotalSeconds -ge 30) {
            Write-Verbose ("Idle: {0}s remaining before {1}" -f $remaining, $NextStimulus)
            $lastVerboseUtc = $nowLoopUtc
        }

        if ($Interactive -and $script:consoleAvailable) {
            try {
                if ([Console]::KeyAvailable) {
                    $keyInfo = [Console]::ReadKey($true)
                    if ($keyInfo.Key -eq [ConsoleKey]::B) {
                        $bangUtc = (Get-Date).ToUniversalTime()
                        $bangSnapshot = Get-EndpointSnapshotPair
                        $bangRow = New-StimulusRow -Cycle $Cycle -SequenceIndex $SequenceIndex -Stimulus '' -Asset '' -Result 'OperatorHeardBang' -HResult '' -Err '' -IdleSecondsBefore $Seconds -Before $bangSnapshot -After $bangSnapshot -NowUtc $bangUtc
                        $bangRow | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8 -Append
                        Write-StimulusLog -Path $LogPath -Level 'INFO' -Message ("Operator reported hearing the artefact at {0}" -f $bangUtc.ToString('o'))
                    }
                }
            } catch {
                $script:consoleAvailable = $false
            }
        }

        Start-Sleep -Milliseconds 500
    }

    Write-Progress -Activity $activity -Completed
    return $true
}

# ---------------------------------------------------------------------------
# Upfront validation. These throw before anything is created or written, so a
# missing requirement always fails fast with a clear message and no prompt.
# ---------------------------------------------------------------------------

if ($Sequence -contains 'ManualCue' -and -not $IncludeManualCue) {
    throw "'ManualCue' is in -Sequence but -IncludeManualCue was not given. ManualCue only prompts for an operator-driven Teams/Webex trigger and is never run without this explicit switch."
}

if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    throw '-OutputDirectory is required.'
}

if ($ListPlan) {
    Write-Host ''
    Write-Host 'Planned Invoke-AudioStimulus schedule (nothing is played, tested or written):'
    Write-Host ("  Cycles           : {0}" -f $Cycles)
    Write-Host ("  Sequence         : {0}" -f ($Sequence -join ', '))
    Write-Host ("  IdleSeconds      : {0} (+ up to {1}s jitter before each stimulus)" -f $IdleSeconds, $JitterSeconds)
    Write-Host ("  IncludeManualCue : {0}" -f [bool] $IncludeManualCue)
    Write-Host ''

    $planIndex = 0
    for ($planCycle = 1; $planCycle -le $Cycles; $planCycle++) {
        for ($planSeqIndex = 0; $planSeqIndex -lt $Sequence.Count; $planSeqIndex++) {
            $planIndex++
            Write-Host ("  #{0,-4} cycle={1} seq={2} stimulus={3}" -f $planIndex, $planCycle, ($planSeqIndex + 1), $Sequence[$planSeqIndex])
        }
    }

    $estimatedMinSeconds = $Cycles * $Sequence.Count * ($IdleSeconds + 1)
    $estimatedMaxSeconds = $Cycles * $Sequence.Count * ($IdleSeconds + $JitterSeconds + 1)
    Write-Host ''
    Write-Host ("Estimated minimum duration: {0:N0}s; maximum with jitter: {1:N0}s." -f $estimatedMinSeconds, $estimatedMaxSeconds)
    Write-Host 'Excludes any -RecorderOutputDirectory marker delay and ManualCue operator wait.'
    return
}

if (-not $AcknowledgeHearingSafety) {
    throw '-AcknowledgeHearingSafety is required before any stimulus is played. This script never lowers system, endpoint or application volume. Run the first pass with the headset off the ear, or not worn, until you know what level to expect.'
}

# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------

if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
    [void] (New-Item -Path $OutputDirectory -ItemType Directory -Force)
}

$resolvedOutput = (Resolve-Path -LiteralPath $OutputDirectory).ProviderPath
$csvPath = Join-Path $resolvedOutput 'stimulus-log.csv'
$logPath = Join-Path $resolvedOutput 'stimulus.log'
$sessionPath = Join-Path $resolvedOutput 'session.json'
$stopFilePath = Join-Path $resolvedOutput $StopFileName

try {
    $sourcePath = Join-Path $PSScriptRoot 'AudioLoopbackCapture.cs'
    $coreType = 'AudioArtifactHunter.EndpointStateReader' -as [type]
    if ($null -eq $coreType) {
        Add-Type -Path $sourcePath -ErrorAction Stop
    } else {
        # .NET Framework cannot unload an assembly; an older build compiled
        # earlier in this session stays loaded until the process exits.
        $versionProperty = $coreType.GetProperty('CoreVersion')
        $loadedVersion = '0.0.0'
        if ($null -ne $versionProperty) { $loadedVersion = [string] $versionProperty.GetValue($null, $null) }
        if ([version] $loadedVersion -lt [version] '1.3.0') {
            throw ("An older AudioLoopbackCapture.cs build ({0}) is already loaded in this PowerShell session; 1.3.0 or later is required and .NET cannot unload it. Start a new PowerShell process, for example: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Invoke-AudioStimulus.ps1 ..." -f $loadedVersion)
        }
    }

    $resolvedWavPath = $WavPath
    if ([string]::IsNullOrWhiteSpace($resolvedWavPath)) {
        $resolvedWavPath = Resolve-DefaultNotificationWav
    }

    # Record the raw Notification.Default binding read-only. A blank binding
    # (local silencing or Group Policy Preferences) makes ToastDefaultSound
    # silent while DirectWav still plays, which is itself a useful check that
    # the policy took effect on this host.
    $notificationBinding = $null
    $bindingKey = $null
    try {
        $bindingKey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('AppEvents\Schemes\Apps\.Default\Notification.Default\.Current')
        if ($null -ne $bindingKey) {
            $notificationBinding = [string] $bindingKey.GetValue('', $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
        }
    } catch {
        Write-Verbose "Could not read Notification.Default binding: $($_.Exception.Message)"
    } finally {
        if ($null -ne $bindingKey) { $bindingKey.Close() }
    }

    # MailBeep is the "New Mail Notification" event classic Outlook's
    # "Play a sound" uses; it is a separate binding from Notification.Default,
    # so a GPP that blanks one need not blank the other.
    $mailBeepBinding = $null
    $mailBeepKey = $null
    try {
        $mailBeepKey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('AppEvents\Schemes\Apps\.Default\MailBeep\.Current')
        if ($null -ne $mailBeepKey) {
            $mailBeepBinding = [string] $mailBeepKey.GetValue('', $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
        }
    } catch {
        Write-Verbose "Could not read MailBeep binding: $($_.Exception.Message)"
    } finally {
        if ($null -ne $mailBeepKey) { $mailBeepKey.Close() }
    }
    if ([string]::IsNullOrWhiteSpace($mailBeepBinding)) { $mailBeepBinding = 'AppEvents MailBeep (blank)' }

    try { [void] [Console]::KeyAvailable } catch { $script:consoleAvailable = $false }

    $wpnServiceInfo = $null
    try {
        $wpnService = Get-Service -Name 'WpnService' -ErrorAction Stop
        $wpnServiceInfo = [pscustomobject] @{ Status = $wpnService.Status.ToString(); StartType = $wpnService.StartType.ToString() }
    } catch {
        Write-Verbose "WpnService state unavailable: $($_.Exception.Message)"
    }

    $wpnUserServiceInfo = @()
    try {
        $wpnUserServiceInfo = @(Get-Service -Name 'WpnUserService*' -ErrorAction Stop | ForEach-Object {
            [pscustomobject] @{ Name = $_.Name; Status = $_.Status.ToString(); StartType = $_.StartType.ToString() }
        })
    } catch {
        Write-Verbose "WpnUserService state unavailable: $($_.Exception.Message)"
    }

    # SessionId is the Windows logon session and stays constant across runs on
    # one machine, so each process gets its own RunId and the Cycle counter is
    # only meaningful within one RunId.
    $script:runId = '{0:yyyyMMdd-HHmmss}-{1}' -f (Get-Date).ToUniversalTime(), $PID

    if (Test-Path -LiteralPath $csvPath -PathType Leaf) {
        $existingHeader = Get-Content -LiteralPath $csvPath -TotalCount 1
        if ($existingHeader -ne $CSV_HEADER) {
            $archived = Join-Path $resolvedOutput ('stimulus-log-schema-{0:yyyyMMdd-HHmmss}.csv' -f (Get-Date).ToUniversalTime())
            Move-Item -LiteralPath $csvPath -Destination $archived -Force
            Write-StimulusLog -Path $logPath -Level 'INFO' -Message ("Existing stimulus-log.csv had a different column layout and was moved to {0}" -f $archived)
        }
    }
    if (-not (Test-Path -LiteralPath $csvPath -PathType Leaf)) {
        $CSV_HEADER | Set-Content -LiteralPath $csvPath -Encoding UTF8
    }

    Write-StimulusLog -Path $logPath -Level 'INFO' -Message ''
    Write-StimulusLog -Path $logPath -Level 'INFO' -Message 'Audio stimulus run started.'
    Write-StimulusLog -Path $logPath -Level 'INFO' -Message ("  Output      : {0}" -f $resolvedOutput)
    Write-StimulusLog -Path $logPath -Level 'INFO' -Message ("  Sequence    : {0}" -f ($Sequence -join ', '))
    Write-StimulusLog -Path $logPath -Level 'INFO' -Message ("  Cycles      : {0}" -f $Cycles)
    Write-StimulusLog -Path $logPath -Level 'INFO' -Message ("  IdleSeconds : {0} (+ up to {1}s jitter)" -f $IdleSeconds, $JitterSeconds)
    Write-StimulusLog -Path $logPath -Level 'INFO' -Message ("  WavPath     : {0}" -f $resolvedWavPath)
    if ($RecorderOutputDirectory) {
        Write-StimulusLog -Path $logPath -Level 'INFO' -Message ("  RecorderOutputDirectory: {0} (marker after {1}s per stimulus)" -f $RecorderOutputDirectory, $MarkerDelaySeconds)
    }
    Write-StimulusLog -Path $logPath -Level 'INFO' -Message ("To stop cleanly, create: {0}" -f $stopFilePath)
    Write-StimulusLog -Path $logPath -Level 'INFO' -Message ''

    $startedUtc = (Get-Date).ToUniversalTime()
    $sessionMetadata = [pscustomobject] [ordered] @{
        SchemaVersion           = '1.0'
        StartedUtc              = $startedUtc.ToString('o')
        ComputerName            = $env:COMPUTERNAME
        UserSid                 = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        OutputDirectory         = $resolvedOutput
        Sequence                = $Sequence
        Cycles                  = $Cycles
        IdleSeconds             = $IdleSeconds
        JitterSeconds           = $JitterSeconds
        WavPath                 = $resolvedWavPath
        ToastAppId              = $ToastAppId
        IncludeManualCue        = [bool] $IncludeManualCue
        RecorderOutputDirectory = $RecorderOutputDirectory
        MarkerDelaySeconds      = $MarkerDelaySeconds
        Interactive             = [bool] $Interactive
        StopFileName            = $StopFileName
        NotificationDefaultBinding = [pscustomobject] @{
            RawValue = $notificationBinding
            IsBlank  = [string]::IsNullOrWhiteSpace($notificationBinding)
            Note     = 'HKCU AppEvents .Default\Notification.Default\.Current read-only at run start. Blank means notification sounds are silenced for this user (locally or by Group Policy Preferences); ToastDefaultSound is then silent while DirectWav still plays.'
        }
        MailBeepBinding         = [pscustomobject] @{
            RawValue = $mailBeepBinding
            IsBlank  = ($mailBeepBinding -eq 'AppEvents MailBeep (blank)')
            Note     = 'HKCU AppEvents .Default\MailBeep\.Current, the Windows "New Mail Notification" event classic Outlook uses for its new-mail sound. Read-only at run start. Blank means the MailBeep stimulus logs NotPlayed.'
        }
        PushServiceState        = [pscustomobject] @{
            Note            = 'Recorded read-only at run start. A Disabled/Stopped WpnService or WpnUserService here documents host state for this run; it is not itself evidence of what happened during the run, and this script never changes it.'
            WpnService      = $wpnServiceInfo
            WpnUserServices = $wpnUserServiceInfo
        }
    }
    [System.IO.File]::WriteAllText($sessionPath, ($sessionMetadata | ConvertTo-Json -Depth 4), [System.Text.UTF8Encoding]::new($true))

    $summary = @{}
    $stopReason = $null

    try {
        for ($cycle = 1; $cycle -le $Cycles; $cycle++) {
            for ($seqIndex = 0; $seqIndex -lt $Sequence.Count; $seqIndex++) {
                $stimulusName = $Sequence[$seqIndex]
                $sequenceNumber = $seqIndex + 1

                $jitter = 0
                if ($JitterSeconds -gt 0) { $jitter = Get-Random -Minimum 0 -Maximum ($JitterSeconds + 1) }
                $idleTotal = $IdleSeconds + $jitter

                if ($cycle -eq 1 -and $seqIndex -eq 0) {
                    Write-StimulusLog -Path $logPath -Level 'INFO' -Message ("First stimulus ({0}) fires at about {1:HH:mm:ss} local after {2}s idle. Nothing is played before then." -f $stimulusName, (Get-Date).AddSeconds($idleTotal), $idleTotal)
                }

                $idleCompleted = Wait-IdlePeriod -Seconds $idleTotal -StopFilePath $stopFilePath -Interactive:$Interactive -CsvPath $csvPath -LogPath $logPath -Cycle $cycle -SequenceIndex $sequenceNumber -NextStimulus $stimulusName

                if (-not $idleCompleted) {
                    $stopReason = 'Stop file detected during idle wait'
                    break
                }

                $beforeSnapshot = Get-EndpointSnapshotPair
                $fireUtc = (Get-Date).ToUniversalTime()
                Write-Verbose ("Firing {0}: console endpoint {1} state={2} volume={3} muted={4}; multimedia endpoint {5}" -f $stimulusName, $beforeSnapshot.EndpointId, $beforeSnapshot.DeviceState, $beforeSnapshot.Volume, $beforeSnapshot.Muted, $beforeSnapshot.EndpointIdMultimedia)

                $stimulusResult = switch ($stimulusName) {
                    'ToastDefaultSound' { Send-ToastStimulus -AppId $ToastAppId -SequenceLabel $stimulusName -NowUtc $fireUtc }
                    'ToastSilent'       { Send-ToastStimulus -AppId $ToastAppId -SequenceLabel $stimulusName -NowUtc $fireUtc -Silent }
                    'DirectWav'         { Invoke-DirectWavStimulus -WavFilePath $resolvedWavPath }
                    'SystemSound'       { Invoke-SystemSoundStimulus }
                    'MailBeep'          { Invoke-MailBeepStimulus }
                    'ManualCue'         { Invoke-ManualCueStimulus -StopFilePath $stopFilePath }
                }

                Start-Sleep -Seconds 1
                $afterSnapshot = Get-EndpointSnapshotPair

                $asset = switch ($stimulusName) {
                    'ToastDefaultSound' { $ToastAppId }
                    'ToastSilent'       { $ToastAppId }
                    'DirectWav'         { $resolvedWavPath }
                    'SystemSound'       { 'SystemSounds.Asterisk' }
                    'MailBeep'          { $mailBeepBinding }
                    'ManualCue'         { 'external' }
                }

                $row = New-StimulusRow -Cycle $cycle -SequenceIndex $sequenceNumber -Stimulus $stimulusName -Asset $asset -Result $stimulusResult.Result -HResult $stimulusResult.HResult -Err $stimulusResult.Error -IdleSecondsBefore $idleTotal -Before $beforeSnapshot -After $afterSnapshot -NowUtc $fireUtc
                $row | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8 -Append

                if (-not $summary.ContainsKey($stimulusName)) { $summary[$stimulusName] = @{} }
                if (-not $summary[$stimulusName].ContainsKey($stimulusResult.Result)) { $summary[$stimulusName][$stimulusResult.Result] = 0 }
                $summary[$stimulusName][$stimulusResult.Result] = $summary[$stimulusName][$stimulusResult.Result] + 1

                $errorSuffix = ''
                if (-not [string]::IsNullOrWhiteSpace($stimulusResult.Error)) { $errorSuffix = " ({0})" -f $stimulusResult.Error }
                Write-StimulusLog -Path $logPath -Level 'INFO' -Message ("[{0:HH:mm:ss}] cycle={1}/{2} seq={3} stimulus={4} result={5}{6}" -f (Get-Date), $cycle, $Cycles, $sequenceNumber, $stimulusName, $stimulusResult.Result, $errorSuffix)

                if ($RecorderOutputDirectory) {
                    $markerDeadlineUtc = (Get-Date).ToUniversalTime().AddSeconds($MarkerDelaySeconds)
                    while ((Get-Date).ToUniversalTime() -lt $markerDeadlineUtc) {
                        if (Test-Path -LiteralPath $stopFilePath -PathType Leaf) { break }
                        Start-Sleep -Milliseconds 500
                    }

                    try {
                        if (-not (Test-Path -LiteralPath $RecorderOutputDirectory -PathType Container)) {
                            Write-StimulusLog -Path $logPath -Level 'WARN' -Message ("Recorder output directory not found, marker skipped: {0}" -f $RecorderOutputDirectory)
                        } else {
                            $markerPath = Join-Path $RecorderOutputDirectory $MARKER_NAME
                            $markerContent = 'AudioArtifactHunter stimulus marker {0} cycle={1} seq={2} stimulus={3}' -f $fireUtc.ToString('o'), $cycle, $sequenceNumber, $stimulusName
                            Set-Content -LiteralPath $markerPath -Value $markerContent -Encoding UTF8
                        }
                    } catch {
                        Write-StimulusLog -Path $logPath -Level 'WARN' -Message ("Could not write recorder marker: {0}" -f $_.Exception.Message)
                    }
                }

                if (Test-Path -LiteralPath $stopFilePath -PathType Leaf) {
                    $stopReason = 'Stop file detected after stimulus'
                    break
                }
            }

            if ($stopReason) { break }
        }

        if (-not $stopReason) { $stopReason = 'All cycles completed' }
    } finally {
        if (Test-Path -LiteralPath $stopFilePath -PathType Leaf) {
            Remove-Item -LiteralPath $stopFilePath -Force -ErrorAction SilentlyContinue
        }

        Write-StimulusLog -Path $logPath -Level 'INFO' -Message ''
        Write-StimulusLog -Path $logPath -Level 'INFO' -Message ("Stopped: {0}" -f $stopReason)
        Write-StimulusLog -Path $logPath -Level 'INFO' -Message 'Summary (played/failed/skipped/cue per stimulus type):'

        foreach ($stimulusKey in ($summary.Keys | Sort-Object)) {
            $resultCounts = $summary[$stimulusKey]
            $parts = New-Object System.Collections.Generic.List[string]
            foreach ($resultKey in ($resultCounts.Keys | Sort-Object)) {
                [void] $parts.Add(('{0}={1}' -f $resultKey, $resultCounts[$resultKey]))
            }
            Write-StimulusLog -Path $logPath -Level 'INFO' -Message ("  {0}: {1}" -f $stimulusKey, ($parts -join ', '))
        }

        Write-StimulusLog -Path $logPath -Level 'INFO' -Message ("Stimulus log  : {0}" -f $csvPath)
        Write-StimulusLog -Path $logPath -Level 'INFO' -Message ("Session file  : {0}" -f $sessionPath)
    }
} catch {
    Write-StimulusLog -Path $logPath -Level 'ERROR' -Message ("Stimulus run failed: {0}" -f $_.Exception.Message)
    throw
}
