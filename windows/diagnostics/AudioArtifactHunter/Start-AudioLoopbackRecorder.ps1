<#
.SYNOPSIS
    Runs a continuous WASAPI loopback recorder with rolling retention, level
    logging and both automatic and operator-triggered preservation.

.DESCRIPTION
    Captures the WASAPI loopback stream exposed by a render endpoint so that a
    rare audio artefact can be bracketed relative to that capture point.

    Designed to run unattended, including from a scheduled task: no parameter
    prompts for interactive-only confirmation (missing -AcknowledgeAudioCapture
    or -OutputDirectory throws immediately instead), a -StopFileName marker
    gives an external caller a clean way to end the run, and -LogPath records
    every status line to disk so nothing is lost when there is no console.

    Outputs are produced under -OutputDirectory:

      rolling\      Rolling WAV segments. Only the most recent -RetainSegments
                    are kept, so disk usage is bounded.
      preserved\    Segments kept permanently because a trigger fired.
      levels.csv    Peak and RMS level in dBFS, one row per second, for the
                    entire monitoring period.
      triggers.csv  One row per automatic trigger. OnsetSegment names the
                    rolling segment that was active when the trigger opened -
                    the segment containing the onset itself - which may differ
                    from PreservedSegment, the segment preceding it.
                    PreOnsetPeakDbfs is the peak of the quiet window before
                    the event (blank with -AbsoluteTrigger); ClosedBy is
                    Cooldown, or Shutdown when the recorder stopped while the
                    event was still open. Rolling segments named in
                    OnsetSegment are purged at stop unless -KeepRollingOnStop.
    capture-events.csv
              WASAPI discontinuities, timestamp errors and HRESULTs.
    endpoint-volume.csv
              Read-only endpoint master scalar and mute state sampled four
              times per second. This is endpoint state, not per-session
              application volume.
    session-volume.csv
              Read-only per-application audio session state (the Volume Mixer
              rows) sampled once per second: one row per session, written on
              change or every 30 second heartbeat, plus a SessionGone row when
              a previously seen session instance disappears. This is the
              per-application layer that endpoint-volume.csv cannot see; the
              level actually heard is session volume x endpoint master volume
              x device.
    endpoint-generations.csv
              Recorder starts/stops and endpoint IDs when supervision is
              enabled across endpoint replacement.
    recorder.log
              Timestamped operational log (see -LogPath). Every status line
              also printed to the console is recorded here, so unattended or
              scheduled-task runs keep a record even with no console attached.
    session.json  Machine, user, endpoint, format and capture parameters.
    evidence-hashes.csv
              SHA-256 hashes of preserved files. Status is Hashed for a normal
              entry or DeletedByRetention for a tombstone left when retention
              removed an already-hashed file.

    The level log is the output that has value even when no artefact occurs. It
    establishes the normal operating level of the endpoint, which is required
    before any reported level can be called abnormal.

    Two preservation paths exist:

      Automatic  - a packet whose peak reaches -TriggerThresholdDbfs while the
                   preceding -OnsetQuietSeconds stayed at or below
                   -OnsetQuietDbfs (a transient out of an idle sink). Sustained
                   loud content such as a call never opens an event, so the
                   preserved directory is not filled by ordinary audio; the
                   number of crossings the gate rejected is reported in the
                   status line. -AbsoluteTrigger restores the plain threshold.
      Operator   - creation of the marker file. A desktop shortcut that creates
                   that file gives a person a one-click way to say "it just
                   happened", which is the only workable trigger for an event
                   that cannot be predicted. The marker is consumed and deleted,
                   so the shortcut can be used repeatedly.

        Interpreting the result requires the endpoint vendor to identify where its
        loopback tap sits. Presence proves the artefact existed no later than that
        tap. Absence only proves it was not present at that tap; it does not assign
        component or vendor ownership.

        This records the endpoint mix and may capture calls, media, notification
        sounds and other user audio. Do not run it without the required customer
        privacy, legal and employee approvals.

    Peak levels are measured from the raw float samples before conversion to the
    16-bit output format, so a level above full scale is reported as a positive
    dBFS value rather than being silently clipped to zero.

.PARAMETER OutputDirectory
    Root directory for all recorder output. Created when absent. Required in
    Record mode; not marked Mandatory on the parameter itself (so a scheduled
    task never blocks on a prompt), but an empty value throws immediately.

.PARAMETER ListDevices
    List render endpoints and exit without recording.

.PARAMETER DeviceId
    Endpoint ID to capture. Defaults to the default render endpoint, which is
    the correct choice on a host with a single virtual render endpoint.

.PARAMETER SegmentSeconds
    Length of each rolling WAV segment. Defaults to 30.

.PARAMETER RetainSegments
    Number of rolling segments kept on disk. Defaults to 20, giving a ten minute
    rolling window at the default segment length.

.PARAMETER TriggerThresholdDbfs
    Peak level at or above which segments are preserved automatically. Defaults
    to -3. Check levels.csv to choose a value from measured data rather than
    guessing.

.PARAMETER OnsetQuietDbfs
    The automatic trigger opens only when the peak of the preceding
    -OnsetQuietSeconds was at or below this level. Defaults to -40. A first
    pilot on a laptop in a Teams call showed a plain -6 dBFS threshold firing
    six times in two minutes, which would fill the preserved directory and
    evict real events under retention.

.PARAMETER OnsetQuietSeconds
    Length of the quiet window the onset gate inspects. Defaults to 2.

.PARAMETER AbsoluteTrigger
    Disables the onset gate so any packet at or above -TriggerThresholdDbfs
    preserves segments, whatever preceded it.

.PARAMETER DurationHours
    Stop automatically after this many hours. Zero, the default, runs until the
    script is stopped.

.PARAMETER StatusIntervalSeconds
    Interval between status lines. Defaults to 60.

.PARAMETER MaxPreservedMegabytes
    Maximum retained size of the preserved directory. The recorder removes the
    oldest preserved files after a marker or segment rotation. Defaults to 2048.

.PARAMETER AcknowledgeAudioCapture
    Required acknowledgment that the operator is authorized to capture and
    retain the endpoint mix in this environment. Not marked Mandatory on the
    parameter itself, so PowerShell never prompts for it interactively; an
    explicit throw enforces the requirement instead, which is what makes the
    missing-acknowledgment case fail cleanly under a scheduled task.

.PARAMETER RestartOnEndpointChange
    When the current endpoint is invalidated, reopen the default render endpoint
    and continue in a new, explicitly logged generation. Valid only when
    -DeviceId is omitted; an explicit endpoint ID must never silently switch.
    A start failure after a restart is retried with capped exponential backoff
    (2s, 4s, 8s, ... up to 30s) until it succeeds or -EndpointWaitMinutes
    elapses, at which point the recorder gives up and throws.

.PARAMETER MaxEndpointRestarts
    Maximum endpoint generations after the initial capture. Defaults to 50.

.PARAMETER EndpointWaitMinutes
    Maximum time to keep retrying a failed endpoint generation start (see
    -RestartOnEndpointChange) before giving up and throwing. Defaults to 30.

.PARAMETER StopFileName
    Name of a marker file that, when it appears directly in -OutputDirectory,
    ends the run cleanly (the file is deleted and a 'StopRequested' generation
    event is logged). Defaults to 'STOP-RECORDER.txt'. Gives an external
    caller - a scheduled task, another script - a way to stop the recorder
    without sending Ctrl+C to a console that may not exist.

.PARAMETER LogPath
    Path to the operational log file that mirrors every status and summary
    line normally written to the console. Defaults to recorder.log inside
    -OutputDirectory. Useful when the recorder runs unattended and no console
    output would otherwise be captured.

.PARAMETER KeepRollingOnStop
    Leave the rolling\ segments on disk when the recorder makes its final stop
    instead of deleting them. By default they are purged at final stop because
    the rolling directory holds the last several minutes of endpoint audio and
    should not linger once the run has ended; this switch opts out of that for
    cases where the rolling window itself is needed afterward.

.EXAMPLE
    .\Start-AudioLoopbackRecorder.ps1 -ListDevices

    Show the render endpoints on this machine.

.EXAMPLE
    .\Start-AudioLoopbackRecorder.ps1 -OutputDirectory D:\AudioCapture

    Record the default render endpoint until stopped, preserving anything at or
    above -3 dBFS.

.EXAMPLE
    .\Start-AudioLoopbackRecorder.ps1 -OutputDirectory D:\AudioCapture -TriggerThresholdDbfs -6 -DurationHours 12

    Record a twelve hour shift with a more sensitive automatic trigger.

.NOTES
    Author:   Anton Romanyuk
    Version:  1.0.0
    Requires: PowerShell 5.1 and AudioLoopbackCapture.cs alongside this script.

    The capture core is C# because PowerShell 5.1 cannot call WASAPI directly.
    It is compiled at run time by the in-box .NET Framework compiler, so no SDK,
    toolchain or third-party package is needed on the target machine.

    Validate on one machine before fleet deployment. If any COM interface ID is
    rejected by the platform the recorder fails immediately with the HRESULT
    rather than recording silence.

    Disk use at 48 kHz stereo 16-bit is roughly 11 MB per minute of rolling
    window. Defaults hold about 110 MB plus anything preserved.
#>

#Requires -Version 5.1

[CmdletBinding(DefaultParameterSetName = 'Record')]
param(
    [Parameter(ParameterSetName = 'Record')]
    [string] $OutputDirectory,

    [Parameter(ParameterSetName = 'List', Mandatory)]
    [switch] $ListDevices,

    [Parameter(ParameterSetName = 'Record')]
    [string] $DeviceId,

    [Parameter(ParameterSetName = 'Record')]
    [ValidateRange(5, 600)]
    [int] $SegmentSeconds = 30,

    [Parameter(ParameterSetName = 'Record')]
    [ValidateRange(2, 500)]
    [int] $RetainSegments = 20,

    [Parameter(ParameterSetName = 'Record')]
    [ValidateRange(-60, 12)]
    [double] $TriggerThresholdDbfs = -3,

    [Parameter(ParameterSetName = 'Record')]
    [ValidateRange(-120, -1)]
    [double] $OnsetQuietDbfs = -40,

    [Parameter(ParameterSetName = 'Record')]
    [ValidateRange(0.5, 60)]
    [double] $OnsetQuietSeconds = 2,

    [Parameter(ParameterSetName = 'Record')]
    [switch] $AbsoluteTrigger,

    [Parameter(ParameterSetName = 'Record')]
    [ValidateRange(0, 168)]
    [double] $DurationHours = 0,

    [Parameter(ParameterSetName = 'Record')]
    [ValidateRange(5, 3600)]
    [int] $StatusIntervalSeconds = 60,

    [Parameter(ParameterSetName = 'Record')]
    [ValidateRange(128, 102400)]
    [int] $MaxPreservedMegabytes = 2048,

    [Parameter(ParameterSetName = 'Record')]
    [switch] $AcknowledgeAudioCapture,

    [Parameter(ParameterSetName = 'Record')]
    [switch] $RestartOnEndpointChange,

    [Parameter(ParameterSetName = 'Record')]
    [ValidateRange(1, 1000)]
    [int] $MaxEndpointRestarts = 50,

    [Parameter(ParameterSetName = 'Record')]
    [ValidateRange(1, 1440)]
    [int] $EndpointWaitMinutes = 30,

    [Parameter(ParameterSetName = 'Record')]
    [string] $StopFileName = 'STOP-RECORDER.txt',

    [Parameter(ParameterSetName = 'Record')]
    [string] $LogPath,

    [Parameter(ParameterSetName = 'Record')]
    [switch] $KeepRollingOnStop
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$MMDEVICES_RENDER = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render'
$RENDER_ID_PREFIX = '{0.0.0.00000000}.'
$GUID_PATTERN = '\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}'
$MARKER_NAME = 'MARK-INCIDENT.txt'
$DEVICE_STATE_ACTIVE = 1

function Limit-PreservedData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [long] $MaximumBytes,
        [Parameter(Mandatory)] [string] $ManifestPath,
        [Parameter(Mandatory)] [string] $GenerationLogPath,
        [Parameter(Mandatory)] [int] $Generation,
        [AllowEmptyString()] [string] $EndpointId
    )

    $hashedRelativePaths = @{}
    if (Test-Path -LiteralPath $ManifestPath -PathType Leaf) {
        foreach ($row in @(Import-Csv -LiteralPath $ManifestPath)) {
            # A manifest written before the Status column existed has no such
            # property at all; PSObject.Properties.Match avoids a strict-mode
            # error for a property that is simply absent on an older row.
            if ($row.PSObject.Properties.Match('Status').Count -and $row.Status -eq 'Hashed') { $hashedRelativePaths[$row.RelativePath] = $true }
        }
    }

    $files = @(Get-ChildItem -LiteralPath $Path -Recurse -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTimeUtc)
    $total = [long] 0
    foreach ($file in $files) { $total += $file.Length }

    $deletedRelativePaths = New-Object System.Collections.Generic.List[string]

    foreach ($file in $files) {
        if ($total -le $MaximumBytes) { break }

        $relativePath = $file.FullName.Substring($Path.Length).TrimStart('\')

        # A file that has not been hashed yet must never be deleted by
        # retention: doing so would destroy evidence with no record it ever
        # existed.
        if (-not $hashedRelativePaths.ContainsKey($relativePath)) {
            Write-EndpointGenerationEvent -Path $GenerationLogPath -Generation $Generation -EventName 'RetentionSkippedUnhashed' -EndpointId $EndpointId -Details $relativePath
            continue
        }

        $length = $file.Length
        Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
        $total -= $length
        [void] $deletedRelativePaths.Add($relativePath)
    }

    return $deletedRelativePaths.ToArray()
}

function Write-RetentionTombstones {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowNull()] [AllowEmptyCollection()] [string[]] $DeletedRelativePaths,
        [Parameter(Mandatory)] [string] $ManifestPath,
        [Parameter(Mandatory)] [string] $GenerationLogPath,
        [Parameter(Mandatory)] [int] $Generation,
        [AllowEmptyString()] [string] $EndpointId
    )

    # An empty array returned from Limit-PreservedData unrolls to nothing in
    # PowerShell, so the caller wraps the call in @() and this guard also
    # tolerates $null; either way there is nothing to tombstone.
    if ($null -eq $DeletedRelativePaths -or $DeletedRelativePaths.Count -eq 0) { return }

    $manifestRows = @{}
    if (Test-Path -LiteralPath $ManifestPath -PathType Leaf) {
        foreach ($row in @(Import-Csv -LiteralPath $ManifestPath)) {
            if ($row.PSObject.Properties.Match('Status').Count -and $row.Status -eq 'Hashed') { $manifestRows[$row.RelativePath] = $row }
        }
    }

    $tombstoneRows = New-Object System.Collections.Generic.List[object]

    foreach ($relativePath in $DeletedRelativePaths) {
        $priorRow = $manifestRows[$relativePath]
        $sha256 = ''
        $length = ''
        $lastWriteTimeUtc = ''

        if ($null -ne $priorRow) {
            $sha256 = $priorRow.Sha256
            $length = $priorRow.Length
            $lastWriteTimeUtc = $priorRow.LastWriteTimeUtc
        }

        [void] $tombstoneRows.Add([pscustomobject] @{
            RelativePath     = $relativePath
            Length           = $length
            LastWriteTimeUtc = $lastWriteTimeUtc
            Sha256           = $sha256
            Status           = 'DeletedByRetention'
        })

        Write-EndpointGenerationEvent -Path $GenerationLogPath -Generation $Generation -EventName 'RetentionDeleted' -EndpointId $EndpointId -Details $relativePath
    }

    $tombstoneRows | Export-Csv -LiteralPath $ManifestPath -NoTypeInformation -Encoding UTF8 -Append
}

function Protect-CaptureDirectory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path
    )

    $acl = Get-Acl -LiteralPath $Path -ErrorAction Stop
    $acl.SetAccessRuleProtection($true, $false)

    foreach ($sidValue in @([Security.Principal.WindowsIdentity]::GetCurrent().User.Value, 'S-1-5-18', 'S-1-5-32-544')) {
        $sid = New-Object Security.Principal.SecurityIdentifier($sidValue)
        $rule = New-Object Security.AccessControl.FileSystemAccessRule($sid, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
        [void] $acl.AddAccessRule($rule)
    }

    Set-Acl -LiteralPath $Path -AclObject $acl -ErrorAction Stop
}

function Update-EvidenceHashes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $PreservedPath,
        [Parameter(Mandatory)] [string] $ManifestPath,
        [Parameter(Mandatory)] [string] $GenerationLogPath,
        [Parameter(Mandatory)] [int] $Generation,
        [AllowEmptyString()] [string] $EndpointId,
        [switch] $IncludeRecent
    )

    $known = @{}
    if (Test-Path -LiteralPath $ManifestPath -PathType Leaf) {
        foreach ($row in @(Import-Csv -LiteralPath $ManifestPath)) { $known[$row.RelativePath] = $true }
    }

    # A file still being copied into place must not be hashed yet: reading it
    # mid-copy would record the wrong hash. Anything written in the last five
    # seconds is deferred to the next pass instead. The final pass after the
    # recorder has stopped passes -IncludeRecent because every file is closed
    # by then and there is no next pass.
    $cutoffUtc = (Get-Date).ToUniversalTime().AddSeconds(-5)

    $newRows = New-Object System.Collections.Generic.List[object]
    foreach ($file in @(Get-ChildItem -LiteralPath $PreservedPath -Recurse -File -ErrorAction SilentlyContinue)) {
        $relativePath = $file.FullName.Substring($PreservedPath.Length).TrimStart('\')
        if ($known.ContainsKey($relativePath)) { continue }
        if (-not $IncludeRecent -and $file.LastWriteTimeUtc -gt $cutoffUtc) { continue }

        try {
            $hash = Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256 -ErrorAction Stop
        } catch {
            # A sharing violation while a copy finishes must never terminate
            # the recorder; retry on the next pass instead.
            Write-EndpointGenerationEvent -Path $GenerationLogPath -Generation $Generation -EventName 'HashDeferred' -EndpointId $EndpointId -Details $_.Exception.Message
            continue
        }

        [void] $newRows.Add([pscustomobject] @{
            RelativePath     = $relativePath
            Length           = $file.Length
            LastWriteTimeUtc = $file.LastWriteTimeUtc.ToString('o')
            Sha256           = $hash.Hash
            Status           = 'Hashed'
        })
    }

    if ($newRows.Count -gt 0) {
        $newRows | Export-Csv -LiteralPath $ManifestPath -NoTypeInformation -Encoding UTF8 -Append
    }
}

function New-CaptureRecorder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path,
        [AllowNull()] [string] $EndpointId,
        [Parameter(Mandatory)] [int] $SegmentLength,
        [Parameter(Mandatory)] [int] $SegmentsToRetain,
        [Parameter(Mandatory)] [double] $ThresholdDbfs,
        [Parameter(Mandatory)] [double] $QuietDbfs,
        [Parameter(Mandatory)] [double] $QuietSeconds
    )

    return New-Object 'AudioArtifactHunter.LoopbackRecorder' -ArgumentList @(
        $Path,
        $EndpointId,
        $SegmentLength,
        $SegmentsToRetain,
        $ThresholdDbfs,
        $QuietDbfs,
        $QuietSeconds
    )
}

function Write-EndpointGenerationEvent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [int] $Generation,
        [Parameter(Mandatory)] [string] $EventName,
        [AllowEmptyString()] [string] $EndpointId,
        [AllowEmptyString()] [string] $Details
    )

    [pscustomobject] @{
        TimestampUtc = (Get-Date).ToUniversalTime().ToString('o')
        Generation   = $Generation
        Event        = $EventName
        EndpointId   = $EndpointId
        Details      = $Details
    } | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8 -Append
}

<#
.SYNOPSIS
    Writes one timestamped line to the recorder log and to the console.

.DESCRIPTION
    Every status and summary line the recorder would otherwise only print with
    Write-Host is routed through here so a run with no attached console (a
    scheduled task) still leaves a record on disk. Losing a log line must
    never stop the recorder, so a write failure is swallowed.
#>
function Write-RecorderLog {
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
        # A lost log line must never stop the capture.
    }

    Write-Host $Message
}

<#
.SYNOPSIS
    Compiles the WASAPI capture core into the current session.

.DESCRIPTION
    Loads AudioLoopbackCapture.cs from the script directory. The type is only
    added once per session; a second call is a no-op.
#>
function Import-CaptureCore {
    [CmdletBinding()]
    param()

    if ('AudioArtifactHunter.LoopbackRecorder' -as [type]) {
        # .NET Framework cannot unload an assembly; an older build compiled
        # earlier in this session stays loaded until the process exits, so
        # check its version instead of assuming it matches this script.
        $readerType = 'AudioArtifactHunter.EndpointStateReader' -as [type]
        $loadedVersion = '0.0.0'
        if ($null -ne $readerType) {
            $versionProperty = $readerType.GetProperty('CoreVersion')
            if ($null -ne $versionProperty) { $loadedVersion = [string] $versionProperty.GetValue($null, $null) }
        }
        if ([version] $loadedVersion -lt [version] '1.3.0') {
            throw ("An older AudioLoopbackCapture.cs build ({0}) is already loaded in this PowerShell session; 1.3.0 or later is required and .NET cannot unload it. Start a new PowerShell process, for example: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Start-AudioLoopbackRecorder.ps1 ..." -f $loadedVersion)
        }
        return
    }

    $sourcePath = Join-Path $PSScriptRoot 'AudioLoopbackCapture.cs'
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw "Capture core not found: $sourcePath"
    }

    Add-Type -Path $sourcePath -ErrorAction Stop
}

<#
.SYNOPSIS
    Lists render endpoints with their names and device state.

.DESCRIPTION
    Names are read from the PnP subsystem. The endpoint ID is reconstructed from
    the registry key name and the standard render prefix; the recorder prints the
    authoritative ID it resolved when it starts, which is the value to use if the
    reconstructed one is rejected.

.OUTPUTS
    System.Management.Automation.PSCustomObject[]
#>
function Get-RenderEndpoint {
    [CmdletBinding()]
    param()

    $names = @{}
    try {
        foreach ($device in @(Get-PnpDevice -Class 'AudioEndpoint' -ErrorAction Stop)) {
            if ([string]::IsNullOrWhiteSpace($device.InstanceId)) { continue }
            $match = [regex]::Match($device.InstanceId, $GUID_PATTERN)
            if ($match.Success) { $names[$match.Value.ToLowerInvariant()] = $device.FriendlyName }
        }
    } catch {
        Write-Verbose "PnP enumeration unavailable: $($_.Exception.Message)"
    }

    $results = New-Object System.Collections.Generic.List[object]

    if (-not (Test-Path -LiteralPath $MMDEVICES_RENDER)) {
        return $results.ToArray()
    }

    foreach ($key in @(Get-ChildItem -LiteralPath $MMDEVICES_RENDER -ErrorAction SilentlyContinue)) {
        $guid = $key.PSChildName

        $state = $null
        try {
            $item = Get-ItemProperty -LiteralPath $key.PSPath -Name 'DeviceState' -ErrorAction Stop
            if ($item.PSObject.Properties.Match('DeviceState').Count) { $state = $item.DeviceState }
        } catch {
            Write-Verbose "DeviceState unavailable for $guid"
        }

        $name = $null
        if ($names.ContainsKey($guid.ToLowerInvariant())) { $name = $names[$guid.ToLowerInvariant()] }

        [void] $results.Add([pscustomobject] @{
            FriendlyName = $name
            DeviceState  = $state
            IsActive     = ($state -eq $DEVICE_STATE_ACTIVE)
            EndpointId   = $RENDER_ID_PREFIX + $guid
        })
    }

    return $results.ToArray()
}

if ($PSCmdlet.ParameterSetName -eq 'List') {
    Get-RenderEndpoint | Sort-Object -Property IsActive -Descending | Format-Table -AutoSize FriendlyName, DeviceState, IsActive, EndpointId
    return
}

if (-not $AcknowledgeAudioCapture) {
    throw '-AcknowledgeAudioCapture is required because loopback recording may capture calls and other user audio.'
}

if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    throw '-OutputDirectory is required in Record mode.'
}

if ($RestartOnEndpointChange -and -not [string]::IsNullOrWhiteSpace($DeviceId)) {
    throw '-RestartOnEndpointChange requires the default endpoint; omit -DeviceId.'
}

Import-CaptureCore

if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
    [void] (New-Item -Path $OutputDirectory -ItemType Directory -Force)
}

$resolvedOutput = (Resolve-Path -LiteralPath $OutputDirectory).ProviderPath
Protect-CaptureDirectory -Path $resolvedOutput
$markerPath = Join-Path $resolvedOutput $MARKER_NAME
$stopFilePath = Join-Path $resolvedOutput $StopFileName
$rollingDirectory = Join-Path $resolvedOutput 'rolling'
$preservedDirectory = Join-Path $resolvedOutput 'preserved'
$evidenceManifestPath = Join-Path $resolvedOutput 'evidence-hashes.csv'
$generationLogPath = Join-Path $resolvedOutput 'endpoint-generations.csv'

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Join-Path $resolvedOutput 'recorder.log'
}

$gateQuietDbfs = if ($AbsoluteTrigger) { 0 } else { $OnsetQuietDbfs }
$gateQuietSeconds = if ($AbsoluteTrigger) { 0 } else { $OnsetQuietSeconds }
$recorder = New-CaptureRecorder -Path $resolvedOutput -EndpointId $DeviceId -SegmentLength $SegmentSeconds -SegmentsToRetain $RetainSegments -ThresholdDbfs $TriggerThresholdDbfs -QuietDbfs $gateQuietDbfs -QuietSeconds $gateQuietSeconds
$generation = 1
$recorder.Generation = $generation
$completedTriggerCount = [long] 0

try {
try {
    $recorder.Start()
    $startedDetails = '{0} Hz, {1} ch; attempts=1' -f $recorder.SampleRate, $recorder.Channels
    Write-EndpointGenerationEvent -Path $generationLogPath -Generation $generation -EventName 'Started' -EndpointId $recorder.DeviceId -Details $startedDetails

    Write-RecorderLog -Path $LogPath -Level 'INFO' -Message ''
    Write-RecorderLog -Path $LogPath -Level 'INFO' -Message 'Loopback recorder started.'
    Write-RecorderLog -Path $LogPath -Level 'INFO' -Message ("  Endpoint    : {0}" -f $recorder.DeviceId)
    Write-RecorderLog -Path $LogPath -Level 'INFO' -Message ("  Format      : {0} Hz, {1} channel(s)" -f $recorder.SampleRate, $recorder.Channels)
    Write-RecorderLog -Path $LogPath -Level 'INFO' -Message ("  Output      : {0}" -f $resolvedOutput)
    Write-RecorderLog -Path $LogPath -Level 'INFO' -Message ("  Rolling     : {0} x {1}s segments" -f $RetainSegments, $SegmentSeconds)
    if ($AbsoluteTrigger) {
        Write-RecorderLog -Path $LogPath -Level 'INFO' -Message ("  Auto-trigger: {0} dBFS peak, absolute (no onset gate)" -f $TriggerThresholdDbfs)
    } else {
        Write-RecorderLog -Path $LogPath -Level 'INFO' -Message ("  Auto-trigger: {0} dBFS peak after {1}s at or below {2} dBFS (onset gate)" -f $TriggerThresholdDbfs, $OnsetQuietSeconds, $OnsetQuietDbfs)
    }
    Write-RecorderLog -Path $LogPath -Level 'INFO' -Message ("  Preservation: maximum {0} MB" -f $MaxPreservedMegabytes)
    Write-RecorderLog -Path $LogPath -Level 'INFO' -Message ''

    # Read-only preview of what session-volume.csv will log, so the operator
    # can see the current Volume Mixer rows before an artefact happens. This
    # always previews the Console-role default endpoint; when -DeviceId
    # captures a different endpoint the sessions actually logged may differ.
    try {
        $startupSessions = [AudioArtifactHunter.EndpointStateReader]::ReadDefaultRenderSessions(0)
        foreach ($startupSession in $startupSessions) {
            if ($startupSession.HResult -ne 0) { continue }
            Write-Verbose ("Session at start: Process={0} DisplayName={1} State={2} Volume={3:F2} Muted={4} IsSystemSounds={5}" -f $startupSession.ProcessName, $startupSession.DisplayName, $startupSession.StateName, $startupSession.Volume, $startupSession.Muted, $startupSession.IsSystemSounds)
        }
    } catch {
        Write-Verbose "Could not list startup sessions: $($_.Exception.Message)"
    }
    Write-RecorderLog -Path $LogPath -Level 'INFO' -Message 'To preserve the current audio after hearing an artefact, create this file:'
    Write-RecorderLog -Path $LogPath -Level 'INFO' -Message ("  {0}" -f $markerPath)
    Write-RecorderLog -Path $LogPath -Level 'INFO' -Message 'A desktop shortcut to the following gives an operator a one-click trigger:'
    Write-RecorderLog -Path $LogPath -Level 'INFO' -Message ("  cmd.exe /c echo incident > `"{0}`"" -f $markerPath)
    Write-RecorderLog -Path $LogPath -Level 'INFO' -Message ("To stop cleanly without Ctrl+C, create: {0}" -f $stopFilePath)
    Write-RecorderLog -Path $LogPath -Level 'INFO' -Message ''
    Write-RecorderLog -Path $LogPath -Level 'INFO' -Message 'Press Ctrl+C to stop.'
    Write-RecorderLog -Path $LogPath -Level 'INFO' -Message ''

    $startedUtc = (Get-Date).ToUniversalTime()
    $sessionMetadata = [pscustomobject] @{
        SchemaVersion        = '1.0'
        StartedUtc           = $startedUtc.ToString('o')
        ComputerName         = $env:COMPUTERNAME
        UserSid              = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        EndpointId           = $recorder.DeviceId
        SampleRate           = $recorder.SampleRate
        Channels             = $recorder.Channels
        SegmentSeconds       = $SegmentSeconds
        RetainSegments       = $RetainSegments
        TriggerThresholdDbfs = $TriggerThresholdDbfs
        OnsetGateEnabled     = -not [bool] $AbsoluteTrigger
        OnsetQuietDbfs       = $gateQuietDbfs
        OnsetQuietSeconds    = $gateQuietSeconds
        MaxPreservedMegabytes = $MaxPreservedMegabytes
        RestartOnEndpointChange = [bool] $RestartOnEndpointChange
        MaxEndpointRestarts    = $MaxEndpointRestarts
        EndpointWaitMinutes    = $EndpointWaitMinutes
        StopFileName           = $StopFileName
        LogPath                = $LogPath
        KeepRollingOnStop      = [bool] $KeepRollingOnStop
    }
    [System.IO.File]::WriteAllText((Join-Path $resolvedOutput 'session.json'), ($sessionMetadata | ConvertTo-Json -Depth 3), [System.Text.UTF8Encoding]::new($true))

    $lastStatusUtc = $startedUtc
    $manualMarkCount = 0
    $markerStuck = $false

    while ($true) {
        Start-Sleep -Milliseconds 500

        if (Test-Path -LiteralPath $stopFilePath -PathType Leaf) {
            Remove-Item -LiteralPath $stopFilePath -Force -ErrorAction SilentlyContinue
            Write-EndpointGenerationEvent -Path $generationLogPath -Generation $generation -EventName 'StopRequested' -EndpointId $recorder.DeviceId -Details 'Stop file detected; exiting cleanly.'
            Write-RecorderLog -Path $LogPath -Level 'INFO' -Message 'Stop file detected. Exiting cleanly.'
            break
        }

        if (-not $recorder.IsRunning) {
            $failure = $recorder.LastError
            if ([string]::IsNullOrWhiteSpace($failure)) { $failure = 'the capture loop exited unexpectedly' }
            Write-EndpointGenerationEvent -Path $generationLogPath -Generation $generation -EventName 'Stopped' -EndpointId $recorder.DeviceId -Details $failure

            if (-not $RestartOnEndpointChange -or $generation -gt $MaxEndpointRestarts) {
                throw "Recorder stopped: $failure"
            }

            $completedTriggerCount += $recorder.TriggerCount
            $recorder.Dispose()
            $generation++

            # Bounded retry with capped exponential backoff: an endpoint that
            # just vanished (device unplugged, driver reset) often needs a few
            # seconds before the default render endpoint can be reopened.
            # $recorder deliberately keeps pointing at the prior (already
            # Stop()/Dispose()d) instance until a candidate actually starts,
            # so a permanent failure below still leaves the finally block a
            # safe, non-null object to call Stop()/Dispose() on.
            $waitDeadlineUtc = (Get-Date).ToUniversalTime().AddMinutes($EndpointWaitMinutes)
            $backoffSeconds = 2
            $attempts = 0

            while ($true) {
                $attempts++
                try {
                    $candidate = New-CaptureRecorder -Path $resolvedOutput -EndpointId $null -SegmentLength $SegmentSeconds -SegmentsToRetain $RetainSegments -ThresholdDbfs $TriggerThresholdDbfs -QuietDbfs $gateQuietDbfs -QuietSeconds $gateQuietSeconds
                    $candidate.Generation = $generation
                    $candidate.Start()
                    $recorder = $candidate
                    break
                } catch {
                    $startFailure = $_.Exception.Message
                    Write-EndpointGenerationEvent -Path $generationLogPath -Generation $generation -EventName 'StartFailed' -EndpointId '' -Details $startFailure

                    $nowUtcRetry = (Get-Date).ToUniversalTime()
                    if ($nowUtcRetry -ge $waitDeadlineUtc) {
                        throw "Endpoint generation $generation failed to start after $attempts attempt(s) within $EndpointWaitMinutes minute(s): $startFailure"
                    }

                    Write-RecorderLog -Path $LogPath -Level 'WARN' -Message ("Endpoint generation {0} start attempt {1} failed: {2}. Retrying in {3}s." -f $generation, $attempts, $startFailure, $backoffSeconds)
                    Start-Sleep -Seconds $backoffSeconds
                    $backoffSeconds = [Math]::Min($backoffSeconds * 2, 30)
                }
            }

            $restartDetails = '{0} Hz, {1} ch; attempts={2}' -f $recorder.SampleRate, $recorder.Channels, $attempts
            Write-EndpointGenerationEvent -Path $generationLogPath -Generation $generation -EventName 'Started' -EndpointId $recorder.DeviceId -Details $restartDetails
            Write-RecorderLog -Path $LogPath -Level 'INFO' -Message ("[{0:HH:mm:ss}] Endpoint generation {1} started: {2} ({3})" -f (Get-Date), $generation, $recorder.DeviceId, $restartDetails)
            continue
        }

        if (-not $markerStuck -and (Test-Path -LiteralPath $markerPath -PathType Leaf)) {
            # Preserve the whole rolling window rather than guessing how far back
            # the artefact was; a person reacts seconds after the sound.
            $stamp = (Get-Date).ToUniversalTime().ToString('yyyyMMdd-HHmmss-fff')
            $markerConsumed = $false

            try {
                Remove-Item -LiteralPath $markerPath -Force -ErrorAction Stop
                $markerConsumed = $true
            } catch {
                try {
                    Rename-Item -LiteralPath $markerPath -NewName ("$MARKER_NAME.consumed-$stamp") -ErrorAction Stop
                    $markerConsumed = $true
                } catch {
                    Write-EndpointGenerationEvent -Path $generationLogPath -Generation $generation -EventName 'MarkerStuck' -EndpointId $recorder.DeviceId -Details "Could not remove or rename marker: $($_.Exception.Message)"
                    $markerStuck = $true
                }
            }

            if ($markerConsumed) {
                $targetDirectory = Join-Path $preservedDirectory ("manual-$stamp")
                $copied = $recorder.PreserveRollingWindow($targetDirectory)
                $deleted = @(Limit-PreservedData -Path $preservedDirectory -MaximumBytes ([long] $MaxPreservedMegabytes * 1MB) -ManifestPath $evidenceManifestPath -GenerationLogPath $generationLogPath -Generation $generation -EndpointId $recorder.DeviceId)
                Write-RetentionTombstones -DeletedRelativePaths $deleted -ManifestPath $evidenceManifestPath -GenerationLogPath $generationLogPath -Generation $generation -EndpointId $recorder.DeviceId
                Update-EvidenceHashes -PreservedPath $preservedDirectory -ManifestPath $evidenceManifestPath -GenerationLogPath $generationLogPath -Generation $generation -EndpointId $recorder.DeviceId

                $manualMarkCount++
                Write-RecorderLog -Path $LogPath -Level 'INFO' -Message ("[{0:HH:mm:ss}] Operator marker: {1} segment(s) preserved to {2}" -f (Get-Date), $copied, $targetDirectory)
            }
        }

        $nowUtc = (Get-Date).ToUniversalTime()

        if (($nowUtc - $lastStatusUtc).TotalSeconds -ge $StatusIntervalSeconds) {
            $lastStatusUtc = $nowUtc
            $deleted = @(Limit-PreservedData -Path $preservedDirectory -MaximumBytes ([long] $MaxPreservedMegabytes * 1MB) -ManifestPath $evidenceManifestPath -GenerationLogPath $generationLogPath -Generation $generation -EndpointId $recorder.DeviceId)
            Write-RetentionTombstones -DeletedRelativePaths $deleted -ManifestPath $evidenceManifestPath -GenerationLogPath $generationLogPath -Generation $generation -EndpointId $recorder.DeviceId
            Update-EvidenceHashes -PreservedPath $preservedDirectory -ManifestPath $evidenceManifestPath -GenerationLogPath $generationLogPath -Generation $generation -EndpointId $recorder.DeviceId
            Write-RecorderLog -Path $LogPath -Level 'INFO' -Message ("[{0:HH:mm:ss}] generation {1} | running {2:F1}h | peak {3:F1} dBFS | endpoint volume {4:P0} | muted {5} | triggers {6} | gated crossings {7} | marks {8}" -f `
                (Get-Date), $generation, ($nowUtc - $startedUtc).TotalHours, $recorder.LastPeakDbfs, $recorder.LastEndpointVolumeScalar, $recorder.LastEndpointMuted, ($completedTriggerCount + $recorder.TriggerCount), $recorder.GatedCrossingCount, $manualMarkCount)
        }

        if ($DurationHours -gt 0 -and ($nowUtc - $startedUtc).TotalHours -ge $DurationHours) {
            Write-RecorderLog -Path $LogPath -Level 'INFO' -Message 'Requested duration reached.'
            break
        }
    }
} finally {
    Write-RecorderLog -Path $LogPath -Level 'INFO' -Message 'Stopping recorder.'
    if ($null -ne $recorder) {
        Write-EndpointGenerationEvent -Path $generationLogPath -Generation $generation -EventName 'FinalStop' -EndpointId $recorder.DeviceId -Details 'Recorder wrapper stopped.'
    }
    $recorder.Stop()
    $recorder.Dispose()
    $deleted = @(Limit-PreservedData -Path $preservedDirectory -MaximumBytes ([long] $MaxPreservedMegabytes * 1MB) -ManifestPath $evidenceManifestPath -GenerationLogPath $generationLogPath -Generation $generation -EndpointId $recorder.DeviceId)
    Write-RetentionTombstones -DeletedRelativePaths $deleted -ManifestPath $evidenceManifestPath -GenerationLogPath $generationLogPath -Generation $generation -EndpointId $recorder.DeviceId
    Update-EvidenceHashes -PreservedPath $preservedDirectory -ManifestPath $evidenceManifestPath -GenerationLogPath $generationLogPath -Generation $generation -EndpointId $recorder.DeviceId -IncludeRecent

    if ($KeepRollingOnStop) {
        Write-EndpointGenerationEvent -Path $generationLogPath -Generation $generation -EventName 'RollingRetained' -EndpointId $recorder.DeviceId -Details "Rolling segments retained at $rollingDirectory."
    } else {
        $rollingFiles = @(Get-ChildItem -LiteralPath $rollingDirectory -Filter 'segment-*.wav' -File -ErrorAction SilentlyContinue)
        foreach ($rollingFile in $rollingFiles) {
            try { Remove-Item -LiteralPath $rollingFile.FullName -Force -ErrorAction Stop } catch { }
        }
        Write-EndpointGenerationEvent -Path $generationLogPath -Generation $generation -EventName 'RollingPurged' -EndpointId $recorder.DeviceId -Details "$($rollingFiles.Count) segment(s) purged."
    }

    Write-RecorderLog -Path $LogPath -Level 'INFO' -Message ''
    Write-RecorderLog -Path $LogPath -Level 'INFO' -Message ("Auto-triggers : {0}" -f ($completedTriggerCount + $recorder.TriggerCount))
    Write-RecorderLog -Path $LogPath -Level 'INFO' -Message ("Level log     : {0}" -f (Join-Path $resolvedOutput 'levels.csv'))
    Write-RecorderLog -Path $LogPath -Level 'INFO' -Message ("Preserved     : {0}" -f $preservedDirectory)
}
} catch {
    Write-RecorderLog -Path $LogPath -Level 'ERROR' -Message ("Recorder failed: {0}" -f $_.Exception.Message)
    throw
}
