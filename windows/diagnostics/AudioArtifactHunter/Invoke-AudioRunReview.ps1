<#
.SYNOPSIS
    Reviews one AudioArtifactHunter output directory and lists everything
    that could be an incident, so an unattended run answers "did it happen?"
    without a person listening.

.DESCRIPTION
    Joins the outputs of the recorder, the stimulus script and the state
    monitor by UTC time and writes review.md and review.json into the
    directory. Read-only apart from those two files.

    Candidate events, in order of strength:

      1. Recorder triggers not explained by a stimulus: a trigger with no
         stimulus in the preceding window, or one whose peak exceeds what the
         stimulus asset can produce at unity gain (asset peak from the
         inventory or measured from the WAV) by more than -MarginDb.
      2. Operator markers (preserved\manual-*) and OperatorHeardBang rows.
      3. Endpoint state changes: endpoint ID change, mute change, master
         volume jump of -VolumeJumpPoints or more, device-state change.
      4. Session (Volume Mixer) changes per process.
      5. Recorder lifecycle: endpoint lost, start failures, hash deferrals,
         retention deletions, WASAPI discontinuities after the first packet.

    Per stimulus type it reports how many fired, what the loopback saw in the
    window after each (peak), and whether that is silent (blanked binding or
    not delivered), as expected, or louder than the asset allows.

    What this cannot see: a transient produced after the VDA (Citrix channel,
    client sink, headset). The loopback is clean for those. A second recorder
    instance on the capture endpoint (-CaptureEndpoint, the headset
    microphone) or a listener with the marker is required to detect them.

.PARAMETER Path
    Output directory of a run. Required.

.PARAMETER InventoryPath
    Inventory JSON with SoundScheme peaks. Defaults to the newest
    AudioStackInventory-*.json in -Path, if any.

.PARAMETER WindowSeconds
    Seconds after a stimulus in which a level event is attributed to it.
    Defaults to 4.

.PARAMETER MarginDb
    Decibels above the asset peak from which a playback counts as louder
    than the asset allows. Defaults to 3.

.PARAMETER VolumeJumpPoints
    Master volume change, in percentage points, reported as a jump.
    Defaults to 10.

.EXAMPLE
    .\Invoke-AudioRunReview.ps1 -Path C:\temp\AudioArtifactHunter

.NOTES
    Author:   Anton Romanyuk
    Version:  1.0.0
    Requires: PowerShell 5.1.
#>
[CmdletBinding()]
param(
    [string] $Path,
    [string] $InventoryPath,
    [ValidateRange(1, 60)] [int] $WindowSeconds = 4,
    [ValidateRange(0, 30)] [double] $MarginDb = 3,
    [ValidateRange(1, 100)] [int] $VolumeJumpPoints = 10
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($Path)) { throw '-Path is required.' }
if (-not (Test-Path -LiteralPath $Path -PathType Container)) { throw "Directory not found: $Path" }
$root = (Resolve-Path -LiteralPath $Path).ProviderPath

function ConvertTo-Utc {
    param([AllowEmptyString()] [string] $Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    try { return [datetime]::Parse($Text, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime() } catch { return $null }
}

function Import-CsvIfPresent {
    param([string] $File)
    $full = Join-Path $root $File
    if (-not (Test-Path -LiteralPath $full -PathType Leaf)) { return @() }
    try { return @(Import-Csv -LiteralPath $full) } catch { return @() }
}

function Measure-WavPeakDbfs {
    param([string] $WavPath)
    if ([string]::IsNullOrWhiteSpace($WavPath) -or -not (Test-Path -LiteralPath $WavPath -PathType Leaf)) { return $null }
    try { $bytes = [System.IO.File]::ReadAllBytes($WavPath) } catch { return $null }
    if ($bytes.Length -lt 44 -or [Text.Encoding]::ASCII.GetString($bytes, 0, 4) -ne 'RIFF') { return $null }
    $offset = 12; $bits = 0; $tag = 0; $dataOffset = -1; $dataLength = 0
    while ($offset + 8 -le $bytes.Length) {
        $id = [Text.Encoding]::ASCII.GetString($bytes, $offset, 4)
        $size = [BitConverter]::ToUInt32($bytes, $offset + 4)
        if ($id -eq 'fmt ') { $tag = [BitConverter]::ToUInt16($bytes, $offset + 8); $bits = [BitConverter]::ToUInt16($bytes, $offset + 22) }
        elseif ($id -eq 'data') { $dataOffset = $offset + 8; $dataLength = [Math]::Min([long] $size, [long] ($bytes.Length - $dataOffset)); break }
        $offset += 8 + $size + ($size % 2)
    }
    if ($dataOffset -lt 0 -or $tag -ne 1 -or $bits -ne 16) { return $null }
    $peak = 0
    for ($i = $dataOffset; $i + 1 -lt $dataOffset + $dataLength; $i += 2) {
        $s = [Math]::Abs([int] [BitConverter]::ToInt16($bytes, $i)); if ($s -gt $peak) { $peak = $s }
    }
    if ($peak -eq 0) { return -144.0 }
    return [Math]::Round(20 * [Math]::Log10($peak / 32768.0), 2)
}

# ---------------------------------------------------------------- load
$stimuli = @(Import-CsvIfPresent 'stimulus-log.csv' | ForEach-Object { $_ | Add-Member -NotePropertyName Utc -NotePropertyValue (ConvertTo-Utc $_.TimestampUtc) -PassThru })
$triggers = @(Import-CsvIfPresent 'triggers.csv' | ForEach-Object { $_ | Add-Member -NotePropertyName Utc -NotePropertyValue (ConvertTo-Utc $_.StartUtc) -PassThru })
$levels = @(Import-CsvIfPresent 'levels.csv' | ForEach-Object { $_ | Add-Member -NotePropertyName Utc -NotePropertyValue (ConvertTo-Utc $_.TimestampUtc) -PassThru })
$captureEvents = @(Import-CsvIfPresent 'capture-events.csv')
$generations = @(Import-CsvIfPresent 'endpoint-generations.csv')
$endpointVolume = @(Import-CsvIfPresent 'endpoint-volume.csv')
$sessionVolume = @(Import-CsvIfPresent 'session-volume.csv')
$hashes = @(Import-CsvIfPresent 'evidence-hashes.csv')
$stateFiles = @(Get-ChildItem -LiteralPath $root -Filter 'endpoint-state-*.csv' -File | Where-Object { $_.Name -notmatch '-sessions\.csv$' })
$sessionFiles = @(Get-ChildItem -LiteralPath $root -Filter 'endpoint-state-*-sessions.csv' -File)
$stateRows = @($stateFiles | ForEach-Object { Import-Csv -LiteralPath $_.FullName } | ForEach-Object { $_ | Add-Member -NotePropertyName Utc -NotePropertyValue (ConvertTo-Utc $_.TimestampUtc) -PassThru })
$sessionRows = @($sessionFiles | ForEach-Object { Import-Csv -LiteralPath $_.FullName })
$manualFolders = @(Get-ChildItem -LiteralPath (Join-Path $root 'preserved') -Directory -Filter 'manual-*' -ErrorAction SilentlyContinue)
$preservedFiles = @(Get-ChildItem -LiteralPath (Join-Path $root 'preserved') -Recurse -File -ErrorAction SilentlyContinue)

if ([string]::IsNullOrWhiteSpace($InventoryPath)) {
    $candidate = Get-ChildItem -LiteralPath $root -Filter 'AudioStackInventory-*.json' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime | Select-Object -Last 1
    if ($null -ne $candidate) { $InventoryPath = $candidate.FullName }
}
$soundScheme = @{}
if (-not [string]::IsNullOrWhiteSpace($InventoryPath) -and (Test-Path -LiteralPath $InventoryPath -PathType Leaf)) {
    try {
        $inventory = Get-Content -LiteralPath $InventoryPath -Raw | ConvertFrom-Json
        foreach ($sound in @($inventory.SoundScheme.Sounds)) {
            if ($sound.App -eq '.Default') { $soundScheme[$sound.Event] = $sound }
        }
    } catch { Write-Warning "Inventory could not be read: $($_.Exception.Message)" }
}

function Get-ExpectedAssetPeak {
    param([string] $Stimulus, [string] $Asset)
    $eventName = switch ($Stimulus) {
        'ToastDefaultSound' { 'Notification.Default' }
        'MailBeep'          { 'MailBeep' }
        'SystemSound'       { 'SystemAsterisk' }
        default             { $null }
    }
    $file = $null
    if ($Stimulus -eq 'DirectWav') { $file = $Asset }
    elseif ($null -ne $eventName -and $soundScheme.ContainsKey($eventName)) {
        $entry = $soundScheme[$eventName]
        if ($null -ne $entry.PeakDbfs -and "$($entry.PeakDbfs)" -ne '') { return [pscustomobject] @{ Peak = [double] $entry.PeakDbfs; Source = "inventory $eventName" } }
        $file = $entry.File
    }
    if ($null -eq $file -and $null -ne $eventName) {
        try {
            $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey("AppEvents\Schemes\Apps\.Default\$eventName\.Current")
            if ($null -ne $key) { $file = [Environment]::ExpandEnvironmentVariables([string] $key.GetValue('')); $key.Close() }
        } catch { }
    }
    if ([string]::IsNullOrWhiteSpace($file)) { return $null }
    $peak = Measure-WavPeakDbfs -WavPath $file
    if ($null -eq $peak) { return $null }
    return [pscustomobject] @{ Peak = $peak; Source = "measured $file" }
}

function Get-WindowPeak {
    param([datetime] $Utc)
    if ($null -eq $Utc -or $levels.Count -eq 0) { return $null }
    $from = $Utc.AddSeconds(-1); $to = $Utc.AddSeconds($WindowSeconds)
    $rows = @($levels | Where-Object { $null -ne $_.Utc -and $_.Utc -ge $from -and $_.Utc -le $to })
    if ($rows.Count -eq 0) { return $null }
    return [double] (($rows | ForEach-Object { [double] $_.PeakDbfs } | Measure-Object -Maximum).Maximum)
}

function Get-NearestStimulus {
    param([datetime] $Utc, [int] $BeforeSeconds, [int] $AfterSeconds)
    if ($null -eq $Utc) { return $null }
    $best = $null; $bestDelta = $null
    foreach ($s in $stimuli) {
        if ($null -eq $s.Utc) { continue }
        $delta = ($Utc - $s.Utc).TotalSeconds
        if ($delta -ge -$AfterSeconds -and $delta -le $BeforeSeconds) {
            if ($null -eq $bestDelta -or [Math]::Abs($delta) -lt [Math]::Abs($bestDelta)) { $best = $s; $bestDelta = $delta }
        }
    }
    if ($null -eq $best) { return $null }
    return [pscustomobject] @{ Row = $best; DeltaSeconds = [Math]::Round($bestDelta, 2) }
}

$candidates = New-Object System.Collections.Generic.List[object]
function Add-Candidate {
    param([string] $Kind, [string] $Strength, $Utc, [string] $Detail)
    [void] $candidates.Add([pscustomobject] @{ Kind = $Kind; Strength = $Strength; TimestampUtc = $(if ($null -ne $Utc) { $Utc.ToString('o') } else { '' }); Detail = $Detail })
}

# ---------------------------------------------------------------- 1. triggers
$triggerRows = New-Object System.Collections.Generic.List[object]
foreach ($t in $triggers) {
    $near = Get-NearestStimulus -Utc $t.Utc -BeforeSeconds $WindowSeconds -AfterSeconds 1
    $classification = 'Unexplained'
    $expectedText = ''
    if ($null -ne $near) {
        $expected = Get-ExpectedAssetPeak -Stimulus $near.Row.Stimulus -Asset $near.Row.Asset
        if ($null -eq $expected) { $classification = 'StimulusPlayback (asset peak unknown)' }
        elseif (([double] $t.PeakDbfs) -gt ($expected.Peak + $MarginDb)) { $classification = 'LouderThanAsset'; $expectedText = ('asset {0} dBFS ({1})' -f $expected.Peak, $expected.Source) }
        else { $classification = 'StimulusPlayback'; $expectedText = ('asset {0} dBFS' -f $expected.Peak) }
    }
    $row = [pscustomobject] @{
        StartUtc = $t.StartUtc; PeakDbfs = $t.PeakDbfs; PreOnsetPeakDbfs = $t.PreOnsetPeakDbfs; DurationSeconds = $t.DurationSeconds
        Stimulus = $(if ($null -ne $near) { '{0} ({1:+0.0;-0.0}s)' -f $near.Row.Stimulus, $near.DeltaSeconds } else { 'none' })
        Classification = $classification; Expected = $expectedText; PreservedSegment = $t.PreservedSegment
    }
    [void] $triggerRows.Add($row)
    if ($classification -eq 'Unexplained') { Add-Candidate -Kind 'RecorderTrigger' -Strength 'Strong' -Utc $t.Utc -Detail ("peak {0} dBFS after {1} dBFS background, no stimulus within {2}s; segment {3}" -f $t.PeakDbfs, $t.PreOnsetPeakDbfs, $WindowSeconds, $t.PreservedSegment) }
    elseif ($classification -eq 'LouderThanAsset') { Add-Candidate -Kind 'RecorderTrigger' -Strength 'Strong' -Utc $t.Utc -Detail ("peak {0} dBFS during {1}, {2}; segment {3}" -f $t.PeakDbfs, $row.Stimulus, $expectedText, $t.PreservedSegment) }
}

# ---------------------------------------------------------------- 2. markers
foreach ($folder in $manualFolders) {
    $utc = $null
    if ($folder.Name -match '^manual-(\d{8})-(\d{6})-(\d{3})$') {
        $utc = ConvertTo-Utc ('{0}-{1}-{2}T{3}:{4}:{5}.{6}Z' -f $Matches[1].Substring(0,4), $Matches[1].Substring(4,2), $Matches[1].Substring(6,2), $Matches[2].Substring(0,2), $Matches[2].Substring(2,2), $Matches[2].Substring(4,2), $Matches[3])
    }
    $near = Get-NearestStimulus -Utc $utc -BeforeSeconds 60 -AfterSeconds 5
    $peak = Get-WindowPeak -Utc $(if ($null -ne $utc) { $utc.AddSeconds(-30) } else { $null })
    Add-Candidate -Kind 'OperatorMarker' -Strength 'Strong' -Utc $utc -Detail ("{0}; {1} file(s); nearest stimulus {2}; loopback peak in the 30s before the marker {3}" -f $folder.Name, @(Get-ChildItem -LiteralPath $folder.FullName -File).Count, $(if ($null -ne $near) { '{0} {1}s earlier' -f $near.Row.Stimulus, $near.DeltaSeconds } else { 'none within 60s' }), $(if ($null -ne $peak) { "$peak dBFS" } else { 'no level data' }))
}
foreach ($s in @($stimuli | Where-Object { $_.Result -eq 'OperatorHeardBang' })) {
    Add-Candidate -Kind 'OperatorHeardBang' -Strength 'Strong' -Utc $s.Utc -Detail ("key pressed during cycle {0}; loopback peak in window {1}" -f $s.Cycle, $(Get-WindowPeak -Utc $s.Utc.AddSeconds(-$WindowSeconds)))
}

# ---------------------------------------------------------------- 3. endpoint state
$roles = @($stateRows | Select-Object -ExpandProperty Role -Unique)
foreach ($role in $roles) {
    $rows = @($stateRows | Where-Object { $_.Role -eq $role } | Sort-Object Utc)
    $previous = $null
    foreach ($r in $rows) {
        if ($r.Reason -ne 'Changed' -or $null -eq $previous) { if ($r.Reason -ne 'Stopped') { $previous = $r }; continue }
        $notes = @()
        if ($r.EndpointId -ne $previous.EndpointId) { $notes += ("endpoint {0} -> {1}" -f $previous.EndpointId, $r.EndpointId) }
        if ($r.Muted -ne $previous.Muted) { $notes += ("muted {0} -> {1}" -f $previous.Muted, $r.Muted) }
        if ($r.DeviceState -ne $previous.DeviceState) { $notes += ("device state {0} -> {1}" -f $previous.DeviceState, $r.DeviceState) }
        $d = 0.0; try { $d = [double] $r.MasterPercent - [double] $previous.MasterPercent } catch { }
        if ([Math]::Abs($d) -ge $VolumeJumpPoints) { $notes += ("master {0}% -> {1}%" -f $previous.MasterPercent, $r.MasterPercent) }
        if ($r.HResult -ne '0x00000000' -and -not [string]::IsNullOrWhiteSpace($r.HResult)) { $notes += ("HResult {0} {1}" -f $r.HResult, $r.Error) }
        if ($notes.Count -gt 0) {
            $near = Get-NearestStimulus -Utc $r.Utc -BeforeSeconds $WindowSeconds -AfterSeconds 1
            Add-Candidate -Kind 'EndpointStateChange' -Strength $(if ($notes -match 'endpoint |HResult') { 'Strong' } else { 'Medium' }) -Utc $r.Utc -Detail ("{0}: {1}{2}" -f $role, ($notes -join '; '), $(if ($null -ne $near) { ' (during ' + $near.Row.Stimulus + ')' } else { '' }))
        }
        $previous = $r
    }
}
$prevVol = @{}
foreach ($v in @($endpointVolume | Where-Object { $_.Changed -eq 'True' })) {
    $key = $v.Generation
    if ($prevVol.ContainsKey($key)) {
        $p = $prevVol[$key]
        $d = 0.0; try { $d = [double] $v.MasterPercent - [double] $p.MasterPercent } catch { }
        if ([Math]::Abs($d) -ge $VolumeJumpPoints -or $v.Muted -ne $p.Muted) {
            Add-Candidate -Kind 'RecorderEndpointVolume' -Strength 'Medium' -Utc (ConvertTo-Utc $v.TimestampUtc) -Detail ("master {0}% -> {1}%, muted {2} -> {3} (generation {4})" -f $p.MasterPercent, $v.MasterPercent, $p.Muted, $v.Muted, $key)
        }
    }
    $prevVol[$key] = $v
}

# ---------------------------------------------------------------- 4. sessions
$sessionChanges = New-Object System.Collections.Generic.List[object]
foreach ($s in @($sessionVolume | Where-Object { $_.Changed -eq 'True' })) {
    [void] $sessionChanges.Add([pscustomobject] @{ Source = 'recorder'; TimestampUtc = $s.TimestampUtc; Process = $(if ($s.IsSystemSounds -eq 'True') { 'System Sounds' } else { $s.ProcessName }); State = $s.StateName; Volume = $s.Volume; Muted = $s.Muted })
}
foreach ($s in @($sessionRows | Where-Object { $_.Reason -eq 'Changed' -or $_.Reason -eq 'Gone' })) {
    [void] $sessionChanges.Add([pscustomobject] @{ Source = 'monitor ' + $s.Role; TimestampUtc = $s.TimestampUtc; Process = $(if ($s.IsSystemSounds -eq 'True') { 'System Sounds' } else { $s.ProcessName }); State = $s.StateName; Volume = $s.Volume; Muted = $s.Muted })
}
foreach ($s in @($sessionChanges | Where-Object { $_.Muted -eq 'True' -or ($_.Volume -ne '' -and [double] $_.Volume -lt 0.999) })) {
    Add-Candidate -Kind 'SessionVolume' -Strength 'Weak' -Utc (ConvertTo-Utc $s.TimestampUtc) -Detail ("{0} [{1}] volume {2} muted {3} state {4}" -f $s.Process, $s.Source, $s.Volume, $s.Muted, $s.State)
}

# ---------------------------------------------------------------- 5. lifecycle
foreach ($g in @($generations | Where-Object { $_.Event -in @('Stopped', 'StartFailed', 'HashDeferred', 'RetentionDeleted', 'MarkerStuck', 'RollingRetained') })) {
    Add-Candidate -Kind 'RecorderLifecycle' -Strength $(if ($g.Event -in @('Stopped', 'StartFailed')) { 'Strong' } else { 'Weak' }) -Utc (ConvertTo-Utc $g.TimestampUtc) -Detail ("{0} generation {1}: {2}" -f $g.Event, $g.Generation, $g.Details)
}
$seenFirst = @{}
foreach ($c in $captureEvents) {
    if ($c.Event -eq 'DataDiscontinuity' -and $c.Details -like 'Expected discontinuity*') { continue }
    Add-Candidate -Kind 'CaptureEvent' -Strength 'Medium' -Utc (ConvertTo-Utc $c.TimestampUtc) -Detail ("{0} hr={1} flags={2} frames={3}: {4}" -f $c.Event, $c.HResult, $c.Flags, $c.Frames, $c.Details)
}

# ---------------------------------------------------------------- stimulus verification
$stimulusSummary = New-Object System.Collections.Generic.List[object]
foreach ($group in @($stimuli | Where-Object { $_.Stimulus -ne '' } | Group-Object Stimulus)) {
    $rows = @($group.Group)
    $peaks = @($rows | ForEach-Object { Get-WindowPeak -Utc $_.Utc } | Where-Object { $null -ne $_ })
    $expected = Get-ExpectedAssetPeak -Stimulus $group.Name -Asset $rows[0].Asset
    $maxPeak = $(if ($peaks.Count -gt 0) { ($peaks | Measure-Object -Maximum).Maximum } else { $null })
    $medPeak = $(if ($peaks.Count -gt 0) { [Math]::Round((($peaks | Sort-Object)[[int][Math]::Floor($peaks.Count / 2)]), 2) } else { $null })
    $verdict = 'no level data'
    if ($null -ne $maxPeak) {
        if ($group.Name -in @('ToastSilent', 'ManualCue', 'OutlookSelfMail')) { $verdict = 'not expected to play at fire time' }
        elseif ($null -eq $expected) { $verdict = 'asset peak unknown' }
        elseif ($maxPeak -gt $expected.Peak + $MarginDb) { $verdict = 'LOUDER than the asset allows' }
        elseif ($medPeak -lt $expected.Peak - 20) { $verdict = 'silent or far below asset (blanked binding, low session volume, or not delivered)' }
        else { $verdict = 'as expected' }
    }
    [void] $stimulusSummary.Add([pscustomobject] @{
        Stimulus = $group.Name; Count = $rows.Count
        Results = (($rows | Group-Object Result | ForEach-Object { '{0}={1}' -f $_.Name, $_.Count }) -join ' ')
        LoopbackMedianPeakDbfs = $medPeak; LoopbackMaxPeakDbfs = $maxPeak
        AssetPeakDbfs = $(if ($null -ne $expected) { $expected.Peak } else { $null })
        Verdict = $verdict
    })
    if ($verdict -like 'LOUDER*') { Add-Candidate -Kind 'StimulusLouderThanAsset' -Strength 'Strong' -Utc $null -Detail ("{0}: max loopback peak {1} dBFS vs asset {2} dBFS" -f $group.Name, $maxPeak, $expected.Peak) }
}

# ---------------------------------------------------------------- integrity and coverage
$integrity = New-Object System.Collections.Generic.List[string]
$levelUtcs = @($levels | Where-Object { $null -ne $_.Utc } | ForEach-Object { $_.Utc } | Sort-Object)
if ($stimuli.Count -gt 0) {
    if ($levelUtcs.Count -eq 0) {
        [void] $integrity.Add(("COVERAGE: {0} stimuli fired but the recorder produced no level data; nothing could be detected on the render side." -f $stimuli.Count))
    } else {
        $uncovered = @($stimuli | Where-Object { $null -ne $_.Utc -and ($_.Utc -lt $levelUtcs[0].AddSeconds(-1) -or $_.Utc -gt $levelUtcs[-1].AddSeconds(5)) }).Count
        if ($uncovered -gt 0) { [void] $integrity.Add(("COVERAGE: {0} of {1} stimuli fired outside the recorder's level span ({2:o} to {3:o}); the recorder was not running for them." -f $uncovered, $stimuli.Count, $levelUtcs[0], $levelUtcs[-1])) }
    }
    if ($stateRows.Count -eq 0) { [void] $integrity.Add("COVERAGE: no endpoint-state monitor output in this directory; endpoint and session changes could not be checked.") }
}
$hashed = @{}; foreach ($h in $hashes) { $hashed[$h.RelativePath] = $h }
$preservedRoot = Join-Path $root 'preserved'
foreach ($f in $preservedFiles) {
    $rel = $f.FullName.Substring($preservedRoot.Length).TrimStart('\')
    if (-not $hashed.ContainsKey($rel)) { [void] $integrity.Add("preserved file without hash: $rel") }
    elseif ([long] $hashed[$rel].Length -ne $f.Length) { [void] $integrity.Add("length differs from manifest: $rel") }
}
$sortedLevels = @($levels | Where-Object { $null -ne $_.Utc } | Sort-Object Utc)
for ($i = 1; $i -lt $sortedLevels.Count; $i++) {
    $gap = ($sortedLevels[$i].Utc - $sortedLevels[$i - 1].Utc).TotalSeconds
    if ($gap -gt 5) { [void] $integrity.Add(("level log gap of {0:F0}s from {1}" -f $gap, $sortedLevels[$i - 1].TimestampUtc)) }
}
$rollingLeft = @(Get-ChildItem -LiteralPath (Join-Path $root 'rolling') -Filter '*.wav' -File -ErrorAction SilentlyContinue).Count
if ($rollingLeft -gt 0) { [void] $integrity.Add("$rollingLeft rolling segment(s) still on disk (recorder not stopped cleanly or -KeepRollingOnStop)") }

# ---------------------------------------------------------------- report
function Get-Span { param($Rows, [string] $Field) if ($Rows.Count -eq 0) { return 'none' }; $u = @($Rows | ForEach-Object { ConvertTo-Utc $_.$Field } | Where-Object { $null -ne $_ } | Sort-Object); if ($u.Count -eq 0) { return 'none' }; return ('{0:o} to {1:o} ({2} rows)' -f $u[0], $u[-1], $Rows.Count) }

$strong = @($candidates | Where-Object { $_.Strength -eq 'Strong' })
$coverageIssues = @($integrity | Where-Object { $_ -like 'COVERAGE:*' })
$verdict = $(if ($strong.Count -gt 0) { "{0} strong candidate event(s): open the preserved segments and the ledger for these times." -f $strong.Count } elseif ($coverageIssues.Count -gt 0) { "No candidate event, but coverage is incomplete: " + ($coverageIssues -join ' ') } elseif ($candidates.Count -gt 0) { "No strong candidate; {0} weaker signal(s) listed." -f $candidates.Count } else { 'No candidate event in the VDA-side data. Downstream transients would not appear here.' })

$md = New-Object System.Text.StringBuilder
[void] $md.AppendLine("# Run review: $root")
[void] $md.AppendLine("")
[void] $md.AppendLine("Generated $((Get-Date).ToUniversalTime().ToString('o')). Window after stimulus: ${WindowSeconds}s; margin: ${MarginDb} dB; volume jump: ${VolumeJumpPoints} points.")
[void] $md.AppendLine("")
[void] $md.AppendLine("**Verdict: $verdict**")
[void] $md.AppendLine("")
[void] $md.AppendLine("## Coverage")
[void] $md.AppendLine("")
[void] $md.AppendLine("- Recorder levels: $(Get-Span $levels 'TimestampUtc')")
[void] $md.AppendLine("- Recorder triggers: $($triggers.Count); operator markers: $($manualFolders.Count); preserved files: $($preservedFiles.Count)")
[void] $md.AppendLine("- Stimuli: $(Get-Span $stimuli 'TimestampUtc')")
[void] $md.AppendLine("- Endpoint state rows: $($stateRows.Count) from $($stateFiles.Count) file(s); session rows: $($sessionRows.Count)")
[void] $md.AppendLine("- Inventory used for asset peaks: $(if ($soundScheme.Count -gt 0) { $InventoryPath } else { 'none (WAV files measured directly where reachable)' })")
[void] $md.AppendLine("")
[void] $md.AppendLine("## Candidate events")
[void] $md.AppendLine("")
if ($candidates.Count -eq 0) { [void] $md.AppendLine("None.") } else {
    [void] $md.AppendLine("| Strength | Kind | UTC | Detail |"); [void] $md.AppendLine("|---|---|---|---|")
    foreach ($c in ($candidates | Sort-Object @{ Expression = { switch ($_.Strength) { 'Strong' { 0 } 'Medium' { 1 } default { 2 } } } }, TimestampUtc)) {
        [void] $md.AppendLine(("| {0} | {1} | {2} | {3} |" -f $c.Strength, $c.Kind, $c.TimestampUtc, ($c.Detail -replace '\|', '/')))
    }
}
[void] $md.AppendLine("")
[void] $md.AppendLine("## Stimulus verification")
[void] $md.AppendLine("")
if ($stimulusSummary.Count -eq 0) { [void] $md.AppendLine("No stimulus log.") } else {
    [void] $md.AppendLine("| Stimulus | Count | Results | Loopback median peak | Loopback max peak | Asset peak | Verdict |"); [void] $md.AppendLine("|---|---|---|---|---|---|---|")
    foreach ($s in $stimulusSummary) { [void] $md.AppendLine(("| {0} | {1} | {2} | {3} | {4} | {5} | {6} |" -f $s.Stimulus, $s.Count, $s.Results, $s.LoopbackMedianPeakDbfs, $s.LoopbackMaxPeakDbfs, $s.AssetPeakDbfs, $s.Verdict)) }
}
[void] $md.AppendLine("")
[void] $md.AppendLine("## Recorder triggers")
[void] $md.AppendLine("")
if ($triggerRows.Count -eq 0) { [void] $md.AppendLine("None.") } else {
    [void] $md.AppendLine("| Start UTC | Peak dBFS | Background dBFS | Duration s | Stimulus | Classification | Expected | Preserved |"); [void] $md.AppendLine("|---|---|---|---|---|---|---|---|")
    foreach ($t in $triggerRows) { [void] $md.AppendLine(("| {0} | {1} | {2} | {3} | {4} | {5} | {6} | {7} |" -f $t.StartUtc, $t.PeakDbfs, $t.PreOnsetPeakDbfs, $t.DurationSeconds, $t.Stimulus, $t.Classification, $t.Expected, $t.PreservedSegment)) }
}
[void] $md.AppendLine("")
[void] $md.AppendLine("## Session changes")
[void] $md.AppendLine("")
if ($sessionChanges.Count -eq 0) { [void] $md.AppendLine("None.") } else {
    foreach ($g in ($sessionChanges | Group-Object Process)) { [void] $md.AppendLine(("- {0}: {1} change row(s); volumes {2}" -f $g.Name, $g.Count, ((@($g.Group | Select-Object -ExpandProperty Volume -Unique) | Sort-Object) -join ', '))) }
}
[void] $md.AppendLine("")
[void] $md.AppendLine("## Integrity")
[void] $md.AppendLine("")
if ($integrity.Count -eq 0) { [void] $md.AppendLine("No issues.") } else { foreach ($i in $integrity) { [void] $md.AppendLine("- $i") } }
[void] $md.AppendLine("")
[void] $md.AppendLine("## Limits")
[void] $md.AppendLine("")
[void] $md.AppendLine("The loopback sees the VDA render side only. A transient produced after it (Citrix channel, client sink, headset) is invisible here; it needs the operator marker, the capture-endpoint recorder on the headset microphone, or client-side capture. No candidate is not proof of absence.")

$mdPath = Join-Path $root 'review.md'
$jsonPath = Join-Path $root 'review.json'
[System.IO.File]::WriteAllText($mdPath, $md.ToString(), [System.Text.UTF8Encoding]::new($false))
[pscustomobject] @{
    SchemaVersion = '1.0'; GeneratedUtc = (Get-Date).ToUniversalTime().ToString('o'); Path = $root; Verdict = $verdict
    Candidates = $candidates.ToArray(); StimulusSummary = $stimulusSummary.ToArray(); Triggers = $triggerRows.ToArray(); Integrity = $integrity.ToArray()
} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $jsonPath -Encoding UTF8

Write-Host ''
Write-Host $verdict
Write-Host ("Candidates: {0} strong, {1} medium, {2} weak" -f $strong.Count, @($candidates | Where-Object { $_.Strength -eq 'Medium' }).Count, @($candidates | Where-Object { $_.Strength -eq 'Weak' }).Count)
foreach ($s in $stimulusSummary) { Write-Host ("  {0,-18} n={1,-4} loopback max {2,6} dBFS  asset {3,6} dBFS  {4}" -f $s.Stimulus, $s.Count, $s.LoopbackMaxPeakDbfs, $s.AssetPeakDbfs, $s.Verdict) }
Write-Host ("Report: {0}" -f $mdPath)
