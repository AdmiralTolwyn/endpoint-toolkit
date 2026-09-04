<#
.SYNOPSIS
    Monitors default render endpoint identity, state, master volume and mute,
    plus per-application audio session state, without recording audio.

.DESCRIPTION
    Samples the default Windows render endpoint for the Console, Multimedia and
    Communications roles. A row is written when endpoint identity, device state,
    master volume, mute state or read status changes, plus a periodic heartbeat.

    This collector is intended for wider deployment than loopback recording
    because it does not capture audio content. It records endpoint-level state
    exposed by IAudioEndpointVolume, and (unless -ExcludeSessions is used)
    per-application session state exposed by IAudioSessionControl2 and
    ISimpleAudioVolume. It does not record headset firmware gain, Citrix
    client-side volume, or acoustic output.

    Sessions are the Volume Mixer layer: each running application that has
    played or is playing audio on the endpoint gets its own session with its
    own volume and mute state, independent of the endpoint's master volume.
    The level a listener actually hears is the product of the session volume,
    the endpoint master volume, and any hardware/device-level gain: session x
    endpoint master x device. A session muted or turned down in the Volume
    Mixer explains a "quiet app, loud endpoint" report that endpoint-only state
    cannot. The System Sounds session is a special, always-present session that
    hosts Windows notification sounds rather than any one process's audio.
    Windows also applies "ducking": when a communications-role stream (a call)
    becomes active, other sessions are attenuated automatically; a session that
    reads quieter only during a call may be ducking, not a user or app change.
    Path 2/3 sound (an endpoint-side Teams/Webex optimization engine rendering
    audio outside the VDA) has no corresponding VDA-side session at all, so an
    audio problem on those paths will not show up in session-volume.csv here -
    only in endpoint-side evidence.

    Run the same collector on a Windows VDA and a Windows BYOD endpoint to compare
    adjacent Windows control planes. Dell/Citrix must provide the equivalent
    endpoint-side evidence for ThinOS.

    -OutputPath accepts either a directory or a file. The path is treated as a
    directory when it already exists as a directory, or when it has no file
    extension; in that case the directory is created if absent and a file named
    endpoint-state-<COMPUTERNAME>-<yyyyMMdd-HHmmss>.csv is created inside it.
    Any other path is treated as a file and its parent directory is created if
    absent. The resolved CSV path is always printed to the console and recorded
    in the log file.

    Unless -ExcludeSessions is used, a second CSV is written alongside the
    endpoint-state CSV, named <same base>-sessions.csv, with one row per audio
    session on the Console and Communications default render endpoints
    (Multimedia is not sampled because it duplicates Console's default endpoint
    in practice on current Windows). A small <same base>-session-context.json
    is written once at start recording the observed ducking-preference registry
    value and the measured cost of a session probe.

    A log file, endpoint-state.log, is written alongside the CSV (append,
    UTF-8, ISO 8601 UTC timestamps) recording start, the resolved output path,
    the average per-role probe cost, the stop reason, and any terminating
    error, so a run driven by a scheduled task can be audited afterwards.

    For non-interactive or scheduled-task use, Ctrl+C is not available. Create
    the file named by -StopFileName (default STOP-MONITOR.txt) in the output
    directory to stop the monitor cleanly; the file is deleted and a final
    'Stopped' row is written for each role. Use -Quiet to suppress the
    Write-Host status lines (heartbeat progress is still available via
    -Verbose) when output is redirected to a scheduled-task log.

.PARAMETER OutputPath
    Directory or CSV file to create or append. See the description for the
    directory-or-file resolution rule.

.PARAMETER SampleIntervalMilliseconds
    Polling interval. Defaults to 250 milliseconds.

.PARAMETER HeartbeatSeconds
    Write an unchanged row at this interval to prove the collector remained
    active. Defaults to 30 seconds.

.PARAMETER DurationHours
    Stop after this many hours. Zero runs until interrupted.

.PARAMETER ExcludeSessions
    Skip per-application audio session sampling (the Volume Mixer layer).
    Sessions are captured by default; use this switch to fall back to
    endpoint-only sampling, for example on a build where the extra COM
    activation per poll is undesirable.

.PARAMETER StopFileName
    Name of a marker file that, when created in the output directory, stops
    the monitor cleanly. Intended for non-interactive and scheduled-task use
    where Ctrl+C cannot be sent. Defaults to STOP-MONITOR.txt.

.PARAMETER Quiet
    Suppress the Write-Host status lines (startup banner and heartbeat status).
    Intended for scheduled-task runs; use -Verbose to still capture status in
    the verbose stream.

.EXAMPLE
    .\Start-AudioEndpointStateMonitor.ps1 -OutputPath C:\ProgramData\AudioEvidence

.EXAMPLE
    .\Start-AudioEndpointStateMonitor.ps1 -OutputPath C:\Temp\endpoint-state.csv -DurationHours 8

.EXAMPLE
    .\Start-AudioEndpointStateMonitor.ps1 -OutputPath C:\ProgramData\AudioEvidence -Quiet -StopFileName STOP-MONITOR.txt

    Scheduled-task style run. Create C:\ProgramData\AudioEvidence\STOP-MONITOR.txt
    to stop it cleanly.

.NOTES
    Author:   Anton Romanyuk
    Version:  1.3.0
    Requires: PowerShell 5.1 and AudioLoopbackCapture.cs alongside this script.

    IAudioEndpointVolume returns a normalized, audio-tapered scalar from 0.0 to
    1.0. The percentage in this output is a control position, not a linear signal
    amplitude and not an acoustic dB measurement.

    Each poll performs three COM activations, one per role (Console,
    Multimedia, Communications), through
    EndpointStateReader.ReadDefaultRenderEndpoint. Moving to a callback-based
    model (IMMNotificationClient / IAudioEndpointVolumeCallback) instead of
    polling is a possible future change to reduce that overhead; it is not
    implemented here.

    Unless -ExcludeSessions is used, each poll performs two further COM
    activations (Console and Communications) through
    EndpointStateReader.ReadDefaultRenderSessions, each of which enumerates
    every session on that endpoint. The measured average cost of one such call
    is recorded in <same base>-session-context.json at start.

    The ducking-preference value recorded at start,
    HKCU:\Software\Microsoft\Multimedia\Audio\UserDuckingPreference, is an
    observed implementation location seen on current Windows, not a documented
    policy contract; treat it as informational, not authoritative.
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $OutputPath,

    [ValidateRange(100, 60000)]
    [int] $SampleIntervalMilliseconds = 250,

    [ValidateRange(1, 3600)]
    [int] $HeartbeatSeconds = 30,

    [ValidateRange(0, 168)]
    [double] $DurationHours = 0,

    [switch] $ExcludeSessions,

    [string] $StopFileName = 'STOP-MONITOR.txt',

    [switch] $Quiet
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$roleNames = @{
    0 = 'Console'
    1 = 'Multimedia'
    2 = 'Communications'
}

# Minimum AudioLoopbackCapture.cs build this script needs (session capture).
$REQUIRED_CORE_VERSION = '1.3.0'

<#
.SYNOPSIS
    Appends one timestamped line to the monitor log file.

.DESCRIPTION
    Logging failures must never stop monitoring, so write errors are swallowed.

.PARAMETER Path
    Full path to the log file.

.PARAMETER Message
    Message text to log.
#>
function Write-MonitorLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $Message
    )

    $line = '{0:o} {1}' -f (Get-Date).ToUniversalTime(), $Message

    try {
        Add-Content -LiteralPath $Path -Value $line -Encoding UTF8
    } catch {
        # Logging must never crash monitoring.
    }
}

<#
.SYNOPSIS
    Locks down a capture output directory to the current user, SYSTEM and
    Administrators, removing inherited ACEs.

.DESCRIPTION
    Copied verbatim from Start-AudioLoopbackRecorder.ps1's Protect-CaptureDirectory
    so both collectors apply the same restriction to output that may contain
    evidence. See that script for the owning implementation. Writes the DACL
    only (DirectoryInfo.SetAccessControl), skips the write when the DACL is
    already in the intended state, and reports failure instead of throwing.

.PARAMETER Path
    Directory to protect.

.OUTPUTS
    PSCustomObject with Protected (bool), Changed (bool) and Message.
#>
function Protect-CaptureDirectory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path
    )

    $requiredSids = @([Security.Principal.WindowsIdentity]::GetCurrent().User.Value, 'S-1-5-18', 'S-1-5-32-544')
    $fullControl = [System.Security.AccessControl.FileSystemRights]::FullControl

    try {
        $directory = New-Object System.IO.DirectoryInfo($Path)
        $acl = $directory.GetAccessControl([System.Security.AccessControl.AccessControlSections]::Access)
    } catch {
        return [pscustomobject] @{ Protected = $false; Changed = $false; Message = "Could not read the DACL: $($_.Exception.Message)" }
    }

    $explicitRules = @($acl.GetAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier]))
    $satisfied = $acl.AreAccessRulesProtected
    foreach ($sidValue in $requiredSids) {
        $match = @($explicitRules | Where-Object {
            $_.IdentityReference.Value -eq $sidValue -and
            $_.AccessControlType -eq [System.Security.AccessControl.AccessControlType]::Allow -and
            (($_.FileSystemRights -band $fullControl) -eq $fullControl)
        })
        if ($match.Count -eq 0) { $satisfied = $false }
    }
    foreach ($rule in $explicitRules) {
        if ($requiredSids -notcontains $rule.IdentityReference.Value) { $satisfied = $false }
    }
    if ($satisfied) {
        return [pscustomobject] @{ Protected = $true; Changed = $false; Message = 'DACL already restricted to the current user, SYSTEM and Administrators.' }
    }

    try {
        $acl.SetAccessRuleProtection($true, $false)
        foreach ($rule in $explicitRules) {
            if ($requiredSids -notcontains $rule.IdentityReference.Value) { [void] $acl.RemoveAccessRule($rule) }
        }
        foreach ($sidValue in $requiredSids) {
            $sid = New-Object Security.Principal.SecurityIdentifier($sidValue)
            $rule = New-Object Security.AccessControl.FileSystemAccessRule($sid, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
            [void] $acl.AddAccessRule($rule)
        }
        $directory.SetAccessControl($acl)
        return [pscustomobject] @{ Protected = $true; Changed = $true; Message = 'DACL restricted to the current user, SYSTEM and Administrators.' }
    } catch {
        return [pscustomobject] @{ Protected = $false; Changed = $false; Message = "Could not restrict the DACL (continuing without it): $($_.Exception.Message)" }
    }
}

<#
.SYNOPSIS
    Builds one output row for an endpoint state snapshot.

.PARAMETER Role
    Friendly role name (Console, Multimedia, Communications).

.PARAMETER Snapshot
    EndpointStateSnapshot returned by ReadDefaultRenderEndpoint.

.PARAMETER NowUtc
    Timestamp to record for this row.

.PARAMETER Reason
    Changed, Heartbeat or Stopped.

.OUTPUTS
    System.Management.Automation.PSCustomObject
#>
function New-EndpointStateRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Role,
        [Parameter(Mandatory)] $Snapshot,
        [Parameter(Mandatory)] [datetime] $NowUtc,
        [Parameter(Mandatory)] [string] $Reason
    )

    return [pscustomobject] @{
        TimestampUtc  = $NowUtc.ToString('o')
        ComputerName  = $env:COMPUTERNAME
        UserSid       = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        Role          = $Role
        EndpointId    = $Snapshot.EndpointId
        DeviceState   = $Snapshot.DeviceState
        MasterScalar  = $Snapshot.MasterVolumeScalar.ToString('F6', [Globalization.CultureInfo]::InvariantCulture)
        MasterPercent = ($Snapshot.MasterVolumeScalar * 100.0).ToString('F2', [Globalization.CultureInfo]::InvariantCulture)
        Muted         = $Snapshot.Muted
        HResult       = ('0x{0:X8}' -f $Snapshot.HResult)
        Error         = $Snapshot.Error
        Reason        = $Reason
    }
}

<#
.SYNOPSIS
    Builds one output row for an audio session snapshot (a Volume Mixer row).

.PARAMETER Role
    Friendly role name (Console or Communications) of the endpoint the
    session was read from.

.PARAMETER Session
    AudioSessionSnapshot returned by ReadDefaultRenderSessions.

.PARAMETER NowUtc
    Timestamp to record for this row.

.PARAMETER Reason
    Changed, Heartbeat, Gone or Stopped. For Gone and Stopped the StateName
    column is forced to 'Gone' so a reader does not mistake a stale last-known
    state for a current one.

.OUTPUTS
    System.Management.Automation.PSCustomObject
#>
function New-SessionStateRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Role,
        [Parameter(Mandatory)] $Session,
        [Parameter(Mandatory)] [datetime] $NowUtc,
        [Parameter(Mandatory)] [string] $Reason
    )

    $stateName = $Session.StateName
    if ($Reason -eq 'Gone') { $stateName = 'Gone' }

    return [pscustomobject] @{
        TimestampUtc              = $NowUtc.ToString('o')
        ComputerName              = $env:COMPUTERNAME
        UserSid                   = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        Role                      = $Role
        EndpointId                = $Session.EndpointId
        SessionInstanceIdentifier = $Session.SessionInstanceIdentifier
        SessionIdentifier         = $Session.SessionIdentifier
        ProcessId                 = $Session.ProcessId
        ProcessName               = $Session.ProcessName
        DisplayName               = $Session.DisplayName
        IsSystemSounds            = $Session.IsSystemSounds
        StateName                 = $stateName
        Volume                    = $Session.Volume.ToString('F6', [Globalization.CultureInfo]::InvariantCulture)
        Muted                     = $Session.Muted
        HResult                   = ('0x{0:X8}' -f $Session.HResult)
        Error                     = $Session.Error
        Reason                    = $Reason
    }
}

# Resolve -OutputPath to a directory (for the log and stop file) and a CSV
# file path. A path is treated as a directory when it already exists as a
# directory, or when it carries no file extension; otherwise it is a file
# and its parent directory is what gets created and protected.
$isDirectory = $false
if (Test-Path -LiteralPath $OutputPath -PathType Container) {
    $isDirectory = $true
} elseif ([string]::IsNullOrEmpty([System.IO.Path]::GetExtension($OutputPath))) {
    $isDirectory = $true
}

if ($isDirectory) {
    if (-not (Test-Path -LiteralPath $OutputPath -PathType Container)) {
        [void] (New-Item -Path $OutputPath -ItemType Directory -Force)
    }

    $outputDirectory = (Resolve-Path -LiteralPath $OutputPath).ProviderPath
    $fileName = 'endpoint-state-{0}-{1}.csv' -f $env:COMPUTERNAME, (Get-Date).ToString('yyyyMMdd-HHmmss')
    $resolvedOutputPath = Join-Path $outputDirectory $fileName
} else {
    $parentDirectory = Split-Path -Parent $OutputPath

    if ([string]::IsNullOrWhiteSpace($parentDirectory)) {
        $outputDirectory = (Get-Location).ProviderPath
    } else {
        if (-not (Test-Path -LiteralPath $parentDirectory -PathType Container)) {
            [void] (New-Item -Path $parentDirectory -ItemType Directory -Force)
        }

        $outputDirectory = (Resolve-Path -LiteralPath $parentDirectory).ProviderPath
    }

    $resolvedOutputPath = Join-Path $outputDirectory (Split-Path -Leaf $OutputPath)
}

$logPath = Join-Path $outputDirectory 'endpoint-state.log'
$stopFilePath = Join-Path $outputDirectory $StopFileName
$resolvedOutputBaseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedOutputPath)
$resolvedSessionsPath = Join-Path $outputDirectory ($resolvedOutputBaseName + '-sessions.csv')
$sessionContextPath = Join-Path $outputDirectory ($resolvedOutputBaseName + '-session-context.json')

try {
    $sourcePath = Join-Path $PSScriptRoot 'AudioLoopbackCapture.cs'
    $coreType = 'AudioArtifactHunter.EndpointStateReader' -as [type]
    if ($null -eq $coreType) {
        Add-Type -Path $sourcePath -ErrorAction Stop
    } else {
        # .NET Framework cannot unload an assembly. If this session compiled an
        # older AudioLoopbackCapture.cs earlier, that build stays until the
        # process exits, so fail here with a clear message rather than later
        # with a missing-method error.
        $versionProperty = $coreType.GetProperty('CoreVersion')
        $loadedVersion = '0.0.0'
        if ($null -ne $versionProperty) { $loadedVersion = [string] $versionProperty.GetValue($null, $null) }
        if ([version] $loadedVersion -lt [version] $REQUIRED_CORE_VERSION) {
            throw ("An older AudioLoopbackCapture.cs build ({0}) is already loaded in this PowerShell session; {1} or later is required and .NET cannot unload it. Start a new PowerShell process, for example: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Start-AudioEndpointStateMonitor.ps1 ..." -f $loadedVersion, $REQUIRED_CORE_VERSION)
        }
    }

    $aclResult = Protect-CaptureDirectory -Path $outputDirectory
    if ($aclResult.Protected) {
        Write-MonitorLog -Path $logPath -Message ("Output DACL: {0}" -f $aclResult.Message)
    } else {
        Write-Warning ("Output DACL: {0}" -f $aclResult.Message)
        Write-MonitorLog -Path $logPath -Message ("WARNING: Output DACL: {0}" -f $aclResult.Message)
    }

    if (-not (Test-Path -LiteralPath $resolvedOutputPath -PathType Leaf)) {
        'TimestampUtc,ComputerName,UserSid,Role,EndpointId,DeviceState,MasterScalar,MasterPercent,Muted,HResult,Error,Reason' |
            Set-Content -LiteralPath $resolvedOutputPath -Encoding UTF8
    }

    if ((-not $ExcludeSessions) -and -not (Test-Path -LiteralPath $resolvedSessionsPath -PathType Leaf)) {
        'TimestampUtc,ComputerName,UserSid,Role,EndpointId,SessionInstanceIdentifier,SessionIdentifier,ProcessId,ProcessName,DisplayName,IsSystemSounds,StateName,Volume,Muted,HResult,Error,Reason' |
            Set-Content -LiteralPath $resolvedSessionsPath -Encoding UTF8
    }

    Write-Host "Monitoring endpoint state: $resolvedOutputPath"
    if (-not $ExcludeSessions) { Write-Host "Monitoring audio sessions: $resolvedSessionsPath" }
    Write-Host 'No audio content is recorded. Press Ctrl+C to stop, or create the stop file to stop non-interactively:'
    Write-Host "  $stopFilePath"

    Write-MonitorLog -Path $logPath -Message ("Monitor started. Interval={0}ms Heartbeat={1}s Duration={2}h StopFile={3} ExcludeSessions={4}" -f $SampleIntervalMilliseconds, $HeartbeatSeconds, $DurationHours, $stopFilePath, [bool] $ExcludeSessions)
    Write-MonitorLog -Path $logPath -Message ("Resolved output path: {0}" -f $resolvedOutputPath)

    # Measure the per-role activation cost once at start so overhead can be
    # judged from the log without changing what gets sampled during the run.
    foreach ($role in 0, 1, 2) {
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        for ($sampleIndex = 0; $sampleIndex -lt 20; $sampleIndex++) {
            [void] [AudioArtifactHunter.EndpointStateReader]::ReadDefaultRenderEndpoint($role)
        }
        $stopwatch.Stop()

        $averageMs = $stopwatch.Elapsed.TotalMilliseconds / 20.0
        Write-MonitorLog -Path $logPath -Message ("Probe cost role={0} avgMs={1:F3}" -f $roleNames[$role], $averageMs)
    }

    # Read-only, once at start: the observed ducking-preference registry value
    # (see .NOTES - this is an observed implementation location, not a
    # documented policy contract) and the measured cost of one session probe,
    # so both can be judged from the session-context file without changing
    # what gets sampled during the run.
    $userDuckingPreference = $null
    try {
        $duckingItem = Get-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Multimedia\Audio' -Name 'UserDuckingPreference' -ErrorAction Stop
        if ($duckingItem.PSObject.Properties.Match('UserDuckingPreference').Count) { $userDuckingPreference = $duckingItem.UserDuckingPreference }
    } catch {
        Write-Verbose "UserDuckingPreference unavailable: $($_.Exception.Message)"
    }

    $sessionStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    for ($sampleIndex = 0; $sampleIndex -lt 20; $sampleIndex++) {
        [void] [AudioArtifactHunter.EndpointStateReader]::ReadDefaultRenderSessions(0)
    }
    $sessionStopwatch.Stop()
    $sessionProbeAverageMs = $sessionStopwatch.Elapsed.TotalMilliseconds / 20.0
    Write-MonitorLog -Path $logPath -Message ("Session probe cost avgMs={0:F3}" -f $sessionProbeAverageMs)
    Write-MonitorLog -Path $logPath -Message ("UserDuckingPreference={0}" -f $userDuckingPreference)

    $sessionContext = [pscustomobject] @{
        SchemaVersion             = '1.0'
        RecordedUtc               = (Get-Date).ToUniversalTime().ToString('o')
        UserDuckingPreference     = $userDuckingPreference
        UserDuckingPreferenceNote = 'HKCU:\Software\Microsoft\Multimedia\Audio\UserDuckingPreference is an observed implementation location seen on current Windows, not a documented policy contract.'
        SessionProbeAverageMs     = $sessionProbeAverageMs
        SessionsEnabled           = (-not [bool] $ExcludeSessions)
        SessionRoles              = @('Console', 'Communications')
    }
    [System.IO.File]::WriteAllText($sessionContextPath, ($sessionContext | ConvertTo-Json -Depth 3), [System.Text.UTF8Encoding]::new($true))
    Write-MonitorLog -Path $logPath -Message ("Session context written: {0}" -f $sessionContextPath)

    $lastState = @{}
    $lastSnapshot = @{}
    $lastSessionState = @{}
    $lastSessionSnapshot = @{}
    $lastHeartbeatUtc = [datetime]::MinValue
    $startedUtc = (Get-Date).ToUniversalTime()
    $rowCount = 0
    $changeCount = 0
    $sessionRowCount = 0
    $sessionChangeCount = 0
    $stopReason = $null

    while ($true) {
        $nowUtc = (Get-Date).ToUniversalTime()
        $heartbeatDue = (($nowUtc - $lastHeartbeatUtc).TotalSeconds -ge $HeartbeatSeconds)

        foreach ($role in 0, 1, 2) {
            $snapshot = [AudioArtifactHunter.EndpointStateReader]::ReadDefaultRenderEndpoint($role)
            $stateKey = '{0}|{1}|{2:F6}|{3}|{4}|{5}' -f $snapshot.EndpointId, $snapshot.DeviceState, $snapshot.MasterVolumeScalar, $snapshot.Muted, $snapshot.HResult, $snapshot.Error
            $changed = (-not $lastState.ContainsKey($role) -or $lastState[$role] -ne $stateKey)

            if ($changed -or $heartbeatDue) {
                $reason = if ($changed) { 'Changed' } else { 'Heartbeat' }
                $row = New-EndpointStateRow -Role $roleNames[$role] -Snapshot $snapshot -NowUtc $nowUtc -Reason $reason

                $row | Export-Csv -LiteralPath $resolvedOutputPath -NoTypeInformation -Encoding UTF8 -Append
                $rowCount++
                if ($changed) { $changeCount++ }

                Write-Verbose ("Role={0} EndpointId={1} State={2} Percent={3}% Muted={4} Reason={5}" -f $row.Role, $row.EndpointId, $row.DeviceState, $row.MasterPercent, $row.Muted, $row.Reason)
            }

            $lastState[$role] = $stateKey
            $lastSnapshot[$role] = $snapshot
        }

        if (-not $ExcludeSessions) {
            # Multimedia is not sampled here: it duplicates Console's default
            # render endpoint in practice on current Windows, so it would only
            # write the same sessions twice.
            $currentSessionKeys = @{}

            foreach ($role in 0, 2) {
                $sessions = [AudioArtifactHunter.EndpointStateReader]::ReadDefaultRenderSessions($role)

                foreach ($session in $sessions) {
                    if ([string]::IsNullOrEmpty($session.SessionInstanceIdentifier)) { continue }

                    $sessionKey = '{0}|{1}' -f $role, $session.SessionInstanceIdentifier
                    $currentSessionKeys[$sessionKey] = $true
                    $sessionStateKey = '{0}|{1:F6}|{2}|{3}' -f $session.StateName, $session.Volume, $session.Muted, $session.HResult
                    $sessionChanged = (-not $lastSessionState.ContainsKey($sessionKey) -or $lastSessionState[$sessionKey] -ne $sessionStateKey)

                    if ($sessionChanged -or $heartbeatDue) {
                        $sessionReason = if ($sessionChanged) { 'Changed' } else { 'Heartbeat' }
                        $sessionRow = New-SessionStateRow -Role $roleNames[$role] -Session $session -NowUtc $nowUtc -Reason $sessionReason

                        $sessionRow | Export-Csv -LiteralPath $resolvedSessionsPath -NoTypeInformation -Encoding UTF8 -Append
                        $sessionRowCount++
                        if ($sessionChanged) { $sessionChangeCount++ }

                        Write-Verbose ("Role={0} Session={1} Process={2} State={3} Volume={4} Muted={5} Reason={6}" -f $sessionRow.Role, $sessionRow.DisplayName, $sessionRow.ProcessName, $sessionRow.StateName, $sessionRow.Volume, $sessionRow.Muted, $sessionRow.Reason)
                    }

                    $lastSessionState[$sessionKey] = $sessionStateKey
                    $lastSessionSnapshot[$sessionKey] = $session
                }
            }

            # A previously tracked session not seen in this pass has gone away
            # (process exited, stream torn down) rather than merely changed.
            $goneSessionKeys = @($lastSessionState.Keys | Where-Object { -not $currentSessionKeys.ContainsKey($_) })

            foreach ($goneSessionKey in $goneSessionKeys) {
                $goneRole = [int] ($goneSessionKey.Split('|')[0])
                $goneRow = New-SessionStateRow -Role $roleNames[$goneRole] -Session $lastSessionSnapshot[$goneSessionKey] -NowUtc $nowUtc -Reason 'Gone'

                $goneRow | Export-Csv -LiteralPath $resolvedSessionsPath -NoTypeInformation -Encoding UTF8 -Append
                $sessionRowCount++

                Write-Verbose ("Role={0} Session={1} Process={2} Reason=Gone" -f $goneRow.Role, $goneRow.DisplayName, $goneRow.ProcessName)

                $lastSessionState.Remove($goneSessionKey)
                $lastSessionSnapshot.Remove($goneSessionKey)
            }
        }

        if ($heartbeatDue) {
            $lastHeartbeatUtc = $nowUtc
            $statusLine = '[{0:HH:mm:ss}] rows={1} changes={2} sessionRows={3} sessionChanges={4} heartbeat ok' -f (Get-Date), $rowCount, $changeCount, $sessionRowCount, $sessionChangeCount
            if (-not $Quiet) { Write-Host $statusLine }
            Write-Verbose $statusLine
        }

        if ($DurationHours -gt 0 -and ($nowUtc - $startedUtc).TotalHours -ge $DurationHours) {
            $stopReason = 'Duration expired'
            break
        }

        if (Test-Path -LiteralPath $stopFilePath -PathType Leaf) {
            Remove-Item -LiteralPath $stopFilePath -Force -ErrorAction SilentlyContinue
            $stopReason = 'Stop file detected'
            break
        }

        Start-Sleep -Milliseconds $SampleIntervalMilliseconds
    }

    $stopUtc = (Get-Date).ToUniversalTime()
    foreach ($role in 0, 1, 2) {
        if ($lastSnapshot.ContainsKey($role)) {
            $row = New-EndpointStateRow -Role $roleNames[$role] -Snapshot $lastSnapshot[$role] -NowUtc $stopUtc -Reason 'Stopped'
            $row | Export-Csv -LiteralPath $resolvedOutputPath -NoTypeInformation -Encoding UTF8 -Append
            $rowCount++

            Write-Verbose ("Role={0} EndpointId={1} State={2} Percent={3}% Muted={4} Reason={5}" -f $row.Role, $row.EndpointId, $row.DeviceState, $row.MasterPercent, $row.Muted, $row.Reason)
        }
    }

    if (-not $ExcludeSessions) {
        foreach ($sessionKey in @($lastSessionSnapshot.Keys)) {
            $stoppedRole = [int] ($sessionKey.Split('|')[0])
            $sessionRow = New-SessionStateRow -Role $roleNames[$stoppedRole] -Session $lastSessionSnapshot[$sessionKey] -NowUtc $stopUtc -Reason 'Stopped'
            $sessionRow | Export-Csv -LiteralPath $resolvedSessionsPath -NoTypeInformation -Encoding UTF8 -Append
            $sessionRowCount++

            Write-Verbose ("Role={0} Session={1} Process={2} Reason=Stopped" -f $sessionRow.Role, $sessionRow.DisplayName, $sessionRow.ProcessName)
        }
    }

    Write-MonitorLog -Path $logPath -Message ("Monitor stopped. Reason={0} Rows={1} Changes={2} SessionRows={3} SessionChanges={4}" -f $stopReason, $rowCount, $changeCount, $sessionRowCount, $sessionChangeCount)
    if (-not $Quiet) { Write-Host ("Stopped: {0}" -f $stopReason) }
} catch {
    Write-MonitorLog -Path $logPath -Message ("ERROR: {0}" -f $_.Exception.Message)
    throw
}
