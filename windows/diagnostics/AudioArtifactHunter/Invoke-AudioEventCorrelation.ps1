<#
.SYNOPSIS
    Builds a cross-provider event timeline around an audio incident and profiles
    event-log onset dates across an archived event collection.

.DESCRIPTION
    Consumes archived .evtx files (loose, in a folder, or inside .zip packages
    as collected by support tooling) and answers three questions that manual
    Event Viewer inspection does not:

     1. RETAINED RANGE. In -Scan mode, profiles every machine, channel, provider
         and event ID present: total count, earliest and latest retained record
         (UTC), and level breakdown. The earliest record is not necessarily the
         event's true onset because logs roll over and exports can be
         incomplete.

    2. TIMELINE. With -IncidentTime, emits every event from every provider
       within a window around a reported incident, merged into one chronological
       sequence. Cross-provider ordering is what identifies the sequence leading
       to an artefact; per-log inspection cannot show it. The timeline CSV keeps
       both the local TimeCreated and the UTC TimeCreatedUtc for every event.

     3. DENSITY. Groups events into one-second buckets, keyed on the UTC
         timestamp, and reports any bucket exceeding -BurstThreshold occurrences
         of the same machine, channel, provider and ID. This is a navigation aid
         only: event density does not prove simultaneous audio playback or
         causation.

    Events whose data does not match their publisher's manifest cannot be
    rendered by Event Viewer, which then falls back to an unrelated message
    string. Those events are detected, counted separately, and emitted with
    their raw XML so the payload is preserved for the publisher to decode.
    A rendered description on such an event is meaningless and must not be
    read as a finding. In -Scan mode this detection only inspects events whose
    message failed to render, because reading the raw XML for that check is
    the same operation that is otherwise skipped for performance; an event
    that renders successfully is not XML-scanned in -Scan mode. -IncidentTime
    (timeline) mode captures the raw XML for every event in the window
    regardless, so the check is complete there.

    An event whose TimeCreated could not be read (a malformed or truncated
    record) is not dropped: it is kept with TimeCreated/TimeCreatedUtc set to
    [datetime]::MinValue and TimeMissing set to true, so StrictMode does not
    abort the run and the record is still visible for triage, just not
    orderable against real timestamps.

.PARAMETER Path
    File, folder, or .zip package to analyse. Folders are searched recursively
    for .evtx files. Zip packages are expanded to a temporary directory that is
    removed on completion; any .zip found inside an expanded package is also
    expanded, recursively, up to three levels deep, and the same archive file
    is never expanded twice.

.PARAMETER IncidentTime
    ISO 8601 timestamp including a UTC offset, for example
    2026-09-01T16:59:00+02:00. Timestamps without an explicit offset are
    rejected. The value is converted to the analyst machine's local time only
    for the Get-WinEvent query; UTC is retained in output.

.PARAMETER WindowMinutes
    Half-width of the timeline window in minutes. Defaults to 10, giving a
    20 minute span centred on -IncidentTime.

.PARAMETER ProviderFilter
    Regular expression matched against provider names, case-insensitively.
    Defaults to audio, remoting, device, notification and service-control
    providers. Ignored when -AllProviders is supplied.

.PARAMETER AllProviders
    Include every provider rather than applying -ProviderFilter.

.PARAMETER BurstThreshold
    Minimum number of identical provider/ID events within one UTC second before
    the second is reported as a burst. Defaults to 5.

.PARAMETER OutputPath
    Directory to write CSV output to. The timeline, onset profile, burst list
    and unrendered-event XML are written as separate files.

.PARAMETER LogPath
    Optional run log file. When supplied, start, the files read, event counts,
    output paths written and any terminating error are appended (UTF-8,
    timestamped ISO 8601 UTC), so a scheduled or unattended run can be audited
    afterwards.

.PARAMETER PassThru
    Emit the timeline records on the pipeline. Off by default so that a normal
    run prints the summaries and writes the CSVs without flooding the console
    with thousands of objects.

.EXAMPLE
    .\Invoke-AudioEventCorrelation.ps1 -Path .\HOST_Full_Events.zip -Scan

    Profile every provider and event ID in the package, with first and last
    occurrence, to establish when each error stream began.

.EXAMPLE
    .\Invoke-AudioEventCorrelation.ps1 -Path .\HOST_Full_Events.zip -IncidentTime '2026-09-01 16:59' -OutputPath C:\temp\out

    Emit the merged cross-provider timeline for the twenty minutes around a
    reported incident.

.EXAMPLE
    .\Invoke-AudioEventCorrelation.ps1 -Path .\logs -Scan -AllProviders -BurstThreshold 3

    Profile everything and report any second containing three or more identical
    events.

.NOTES
    Author:   Anton Romanyuk
    Version:  1.1.0
    Requires: PowerShell 5.1.

    No event IDs are interpreted. The script reports provider, ID, level, time
    and payload only. Publisher-specific IDs must be decoded by their publisher.

    Zip expansion uses the in-box Microsoft.PowerShell.Archive Expand-Archive
    cmdlet available in this PowerShell 5.1 release; path-traversal and
    overwrite safety for archive entries relies entirely on that module's
    implementation, not on anything added here. Packages collected from
    untrusted sources should be expanded under an isolated, low-privilege
    account rather than an analyst's primary workstation account.
#>

#Requires -Version 5.1

[CmdletBinding(DefaultParameterSetName = 'Scan')]
param(
    [Parameter(Mandatory, Position = 0)]
    [string] $Path,

    [Parameter(ParameterSetName = 'Scan')]
    [switch] $Scan,

    [Parameter(ParameterSetName = 'Timeline', Mandatory)]
    [ValidatePattern('^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})$')]
    [string] $IncidentTime,

    [Parameter(ParameterSetName = 'Timeline')]
    [ValidateRange(1, 1440)]
    [int] $WindowMinutes = 10,

    [string] $ProviderFilter = 'audio|citrix|pnp|notification|wns|push|terminalservices|remotedesktop|service control manager|application error|windows error reporting|usb|kernel-power',

    [switch] $AllProviders,

    [ValidateRange(2, 1000)]
    [int] $BurstThreshold = 5,

    [string] $OutputPath,

    [string] $LogPath,

    [switch] $PassThru
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$LEVEL_NAMES = @{ 0 = 'LogAlways'; 1 = 'Critical'; 2 = 'Error'; 3 = 'Warning'; 4 = 'Information'; 5 = 'Verbose' }
$MAX_ZIP_DEPTH = 3

<#
.SYNOPSIS
    Appends one timestamped line to the run log file.

.DESCRIPTION
    Logging failures must never stop a run, so write errors are swallowed.

.PARAMETER Path
    Full path to the log file.

.PARAMETER Message
    Message text to log.
#>
function Write-CorrelationLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $Message
    )

    $line = '{0:o} {1}' -f (Get-Date).ToUniversalTime(), $Message

    try {
        Add-Content -LiteralPath $Path -Value $line -Encoding UTF8
    } catch {
        # Logging must never crash the run.
    }
}

<#
.SYNOPSIS
    Expands one zip file and recursively expands any zip found inside it.

.DESCRIPTION
    Guards against expanding the same archive twice (by resolved full path)
    and against unbounded recursion via -MaxDepth. A failure to expand one
    archive is a warning, not a terminating error, so the rest of the package
    can still be analysed.

.PARAMETER ZipPath
    Zip file to expand.

.PARAMETER TempRoot
    Root of the temporary tree that all expansions are written under.

.PARAMETER Depth
    Current recursion depth. The top-level call uses 1.

.PARAMETER MaxDepth
    Maximum recursion depth. Nested zips beyond this depth are left unexpanded.

.PARAMETER ExpandedZips
    Set of resolved, lower-invariant zip paths already expanded in this run.
#>
function Expand-ZipRecursive {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $ZipPath,
        [Parameter(Mandatory)] [string] $TempRoot,
        [Parameter(Mandatory)] [int] $Depth,
        [Parameter(Mandatory)] [int] $MaxDepth,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [System.Collections.Generic.HashSet[string]] $ExpandedZips
    )

    $fullZipPath = (Resolve-Path -LiteralPath $ZipPath).ProviderPath

    if (-not $ExpandedZips.Add($fullZipPath.ToLowerInvariant())) {
        Write-Verbose "Already expanded, skipping: $fullZipPath"
        return
    }

    $target = Join-Path $TempRoot ([guid]::NewGuid().ToString('N'))
    Write-Verbose "Expanding $fullZipPath (depth $Depth)"

    try {
        [void] (New-Item -Path $target -ItemType Directory -Force)
        Expand-Archive -LiteralPath $fullZipPath -DestinationPath $target -Force -ErrorAction Stop
    } catch {
        Write-Warning "Could not expand $fullZipPath : $($_.Exception.Message)"
        return
    }

    if ($Depth -ge $MaxDepth) {
        return
    }

    foreach ($nested in @(Get-ChildItem -LiteralPath $target -Recurse -File -Filter '*.zip' -ErrorAction SilentlyContinue)) {
        Expand-ZipRecursive -ZipPath $nested.FullName -TempRoot $TempRoot -Depth ($Depth + 1) -MaxDepth $MaxDepth -ExpandedZips $ExpandedZips
    }
}

<#
.SYNOPSIS
    Resolves the input path to a list of .evtx files, expanding archives.

.PARAMETER InputPath
    File, folder or archive supplied by the caller.

.PARAMETER TempRoot
    Directory that archives are expanded into.

.OUTPUTS
    System.String[]. Full paths of the .evtx files to read.
#>
function Resolve-EvtxFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $InputPath,
        [Parameter(Mandatory)] [string] $TempRoot
    )

    if (-not (Test-Path -LiteralPath $InputPath)) {
        throw "Path not found: $InputPath"
    }

    $files = New-Object System.Collections.Generic.List[string]
    $candidates = New-Object System.Collections.Generic.List[string]

    if (Test-Path -LiteralPath $InputPath -PathType Leaf) {
        [void] $candidates.Add((Resolve-Path -LiteralPath $InputPath).ProviderPath)
    } else {
        foreach ($item in @(Get-ChildItem -LiteralPath $InputPath -Recurse -File -Include '*.evtx', '*.zip' -ErrorAction SilentlyContinue)) {
            [void] $candidates.Add($item.FullName)
        }
    }

    $expandedZips = New-Object 'System.Collections.Generic.HashSet[string]'

    foreach ($candidate in $candidates) {
        if ([System.IO.Path]::GetExtension($candidate) -ieq '.zip') {
            Expand-ZipRecursive -ZipPath $candidate -TempRoot $TempRoot -Depth 1 -MaxDepth $MAX_ZIP_DEPTH -ExpandedZips $expandedZips
        } elseif ([System.IO.Path]::GetExtension($candidate) -ieq '.evtx') {
            [void] $files.Add($candidate)
        }
    }

    # A single pass over the whole temp tree picks up .evtx from every
    # top-level and nested archive expanded above, regardless of how deep it
    # was nested.
    foreach ($item in @(Get-ChildItem -LiteralPath $TempRoot -Recurse -File -Filter '*.evtx' -ErrorAction SilentlyContinue)) {
        [void] $files.Add($item.FullName)
    }

    return $files.ToArray()
}

<#
.SYNOPSIS
    Reads events from one archived event log into flat records.

.DESCRIPTION
    Message rendering is attempted but never required. An event collected from
    a different machine frequently cannot be rendered locally; such events are
    marked Unrendered and carry their raw XML instead of a description. The
    raw XML is only read when it is actually needed: when -IncludeXml is set,
    or when the event failed to render (unrendered events are also the only
    ones tested for a publisher template mismatch). An event whose Message
    renders successfully is never XML-scanned unless -IncludeXml is set,
    which keeps a -Scan pass over a large archive fast.

.PARAMETER File
    Full path to the .evtx file.

.PARAMETER StartTime
    Optional lower time bound.

.PARAMETER EndTime
    Optional upper time bound.

.PARAMETER IncludeXml
    Capture the raw XML for every event rather than only for unrendered ones.

.PARAMETER SourceSha256
    SHA-256 hash of the source EVTX, retained with every output record.

.OUTPUTS
    System.Management.Automation.PSCustomObject[]
#>
function Read-EventRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $File,
        [datetime] $StartTime,
        [datetime] $EndTime,
        [switch] $IncludeXml,
        [Parameter(Mandatory)] [string] $SourceSha256
    )

    $filter = @{ Path = $File }
    if ($PSBoundParameters.ContainsKey('StartTime')) { $filter['StartTime'] = $StartTime }
    if ($PSBoundParameters.ContainsKey('EndTime')) { $filter['EndTime'] = $EndTime }

    $events = @()
    try {
        $events = @(Get-WinEvent -FilterHashtable $filter -ErrorAction Stop)
    } catch {
        Write-Verbose "No matching events in $File : $($_.Exception.Message)"
        return @()
    }

    $logName = [System.IO.Path]::GetFileNameWithoutExtension($File)
    $records = New-Object System.Collections.Generic.List[object]

    foreach ($eventRecord in $events) {
        $message = $null
        try { $message = $eventRecord.Message } catch { $message = $null }

        $unrendered = [string]::IsNullOrWhiteSpace($message)

        # ToXml() is a comparatively expensive call. It is only made when the
        # result is actually needed: an explicit -IncludeXml request, or an
        # event that failed to render (which is also the only case tested for
        # a publisher template mismatch, below).
        $xml = $null
        if ($IncludeXml -or $unrendered) {
            try { $xml = $eventRecord.ToXml() } catch { $xml = $null }
        }

        # A publisher template mismatch makes Event Viewer fall back to an
        # unrelated message table; the description on such an event is noise.
        # Detecting it needs the raw XML fetched above, so an event that
        # rendered successfully and was not captured via -IncludeXml is never
        # tested here.
        $templateMismatch = $false
        $hasProcessingError = $false
        $processingErrorCode = $null
        if ($null -ne $xml -and $xml -match '<ProcessingErrorData>') {
            $hasProcessingError = $true
            $codeMatch = [regex]::Match($xml, '<ErrorCode>(\d+)</ErrorCode>')
            if ($codeMatch.Success) { $processingErrorCode = [int] $codeMatch.Groups[1].Value }
            if ($processingErrorCode -eq 15005) { $templateMismatch = $true }
        }

        $level = 4
        try { $level = [int] $eventRecord.Level } catch { $level = 4 }

        $levelName = 'Unknown'
        if ($LEVEL_NAMES.ContainsKey($level)) { $levelName = $LEVEL_NAMES[$level] }

        $summary = $null
        if (-not $unrendered -and -not $templateMismatch) {
            $summary = (($message -split "`r?`n") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 2) -join ' '
            if ($summary.Length -gt 400) { $summary = $summary.Substring(0, 400) }
        }

        # A malformed or truncated record can carry a null TimeCreated. Under
        # StrictMode, calling a method on that null throws, so it is guarded
        # here rather than dropping the record: it is kept with a sentinel
        # timestamp and TimeMissing = $true so it stays visible for triage.
        $timeMissing = ($null -eq $eventRecord.TimeCreated)
        if ($timeMissing) {
            $timeCreatedLocal = [datetime]::MinValue
            $timeCreatedUtc = [datetime]::SpecifyKind([datetime]::MinValue, [System.DateTimeKind]::Utc)
        } else {
            $timeCreatedLocal = $eventRecord.TimeCreated
            $timeCreatedUtc = $eventRecord.TimeCreated.ToUniversalTime()
        }

        [void] $records.Add([pscustomobject] @{
            TimeCreated      = $timeCreatedLocal
            TimeCreatedUtc   = $timeCreatedUtc
            TimeMissing      = $timeMissing
            EventRecordId    = $eventRecord.RecordId
            LogFile          = $logName
            SourceSha256     = $SourceSha256
            Channel          = $eventRecord.LogName
            Provider         = $eventRecord.ProviderName
            Id               = $eventRecord.Id
            Level            = $level
            LevelName        = $levelName
            MachineName      = $eventRecord.MachineName
            ProcessId        = $eventRecord.ProcessId
            Unrendered       = $unrendered
            ProcessingError  = $hasProcessingError
            TemplateMismatch = $templateMismatch
            ProcessingErrorCode = $processingErrorCode
            Summary          = $summary
            Xml              = if ($IncludeXml -or $unrendered -or $templateMismatch) { $xml } else { $null }
        })
    }

    return $records.ToArray()
}

<#
.SYNOPSIS
    Summarises each provider and event ID with counts and onset dates.

.PARAMETER Record
    Event records to profile.

.OUTPUTS
    System.Management.Automation.PSCustomObject[]
#>
function Get-OnsetProfile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $Record
    )

    $groups = $Record | Group-Object -Property MachineName, Channel, Provider, Id
    $retainedRanges = New-Object System.Collections.Generic.List[object]

    foreach ($group in $groups) {
        $sorted = $group.Group | Sort-Object -Property TimeCreatedUtc
        $first = $sorted | Select-Object -First 1
        $last = $sorted | Select-Object -Last 1

        $errorCount = @($group.Group | Where-Object { $_.Level -le 2 }).Count
        $mismatchCount = @($group.Group | Where-Object { $_.TemplateMismatch }).Count

        [void] $retainedRanges.Add([pscustomobject] @{
            MachineName        = $first.MachineName
            Channel            = $first.Channel
            Provider           = $first.Provider
            Id                 = $first.Id
            Count              = $group.Count
            ErrorOrCritical    = $errorCount
            TemplateMismatch   = $mismatchCount
            EarliestRetainedUtc = $first.TimeCreatedUtc
            LatestRetainedUtc   = $last.TimeCreatedUtc
            SpanDays           = [Math]::Round(($last.TimeCreatedUtc - $first.TimeCreatedUtc).TotalDays, 2)
            Sample             = $first.Summary
        })
    }

    return @($retainedRanges | Sort-Object -Property EarliestRetainedUtc)
}

<#
.SYNOPSIS
    Finds one-second buckets containing repeated identical events.

.DESCRIPTION
    Reports high event density for navigation. It does not identify which audio
    asset played and does not establish simultaneous playback or causation.
    Buckets are keyed on the UTC timestamp so density is not affected by the
    analyst machine's time zone or daylight-saving transitions.

.PARAMETER Record
    Event records to search.

.PARAMETER Threshold
    Minimum occurrences within one second to report.

.OUTPUTS
    System.Management.Automation.PSCustomObject[]
#>
function Get-EventBurst {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $Record,
        [Parameter(Mandatory)] [int] $Threshold
    )

    $buckets = @{}

    foreach ($item in $Record) {
        $secondUtc = $item.TimeCreatedUtc.ToString('yyyy-MM-dd HH:mm:ss')
        $key = '{0}|{1}|{2}|{3}|{4}' -f $secondUtc, $item.MachineName, $item.Channel, $item.Provider, $item.Id

        if (-not $buckets.ContainsKey($key)) {
            $buckets[$key] = [pscustomobject] @{
                SecondUtc = $item.TimeCreatedUtc
                Machine   = $item.MachineName
                Channel   = $item.Channel
                Provider  = $item.Provider
                Id        = $item.Id
                Count     = 0
                Sample    = $item.Summary
            }
        }

        $buckets[$key].Count++
    }

    return @($buckets.Values | Where-Object { $_.Count -ge $Threshold } | Sort-Object -Property Count -Descending)
}

$hasLog = -not [string]::IsNullOrWhiteSpace($LogPath)
if ($hasLog) {
    $logDirectory = Split-Path -Parent $LogPath
    if (-not [string]::IsNullOrWhiteSpace($logDirectory) -and -not (Test-Path -LiteralPath $logDirectory -PathType Container)) {
        [void] (New-Item -Path $logDirectory -ItemType Directory -Force)
    }

    Write-CorrelationLog -Path $LogPath -Message ("Correlation started. Path={0} Mode={1} AllProviders={2} BurstThreshold={3}" -f $Path, $PSCmdlet.ParameterSetName, [bool] $AllProviders, $BurstThreshold)
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('AudioEventCorrelation-' + [guid]::NewGuid().ToString('N'))
[void] (New-Item -Path $tempRoot -ItemType Directory -Force)

try {
    $files = @(Resolve-EvtxFile -InputPath $Path -TempRoot $tempRoot)
    if ($files.Count -eq 0) { throw "No .evtx files found under: $Path" }

    Write-Verbose "Reading $($files.Count) event log file(s)."
    if ($hasLog) { Write-CorrelationLog -Path $LogPath -Message ("Files resolved: {0}" -f $files.Count) }

    $isTimeline = ($PSCmdlet.ParameterSetName -eq 'Timeline')
    $windowStart = [datetime]::MinValue
    $windowEnd = [datetime]::MaxValue

    if ($isTimeline) {
        $incidentOffset = [datetimeoffset]::Parse($IncidentTime, [Globalization.CultureInfo]::InvariantCulture)
        $windowStart = $incidentOffset.AddMinutes(-$WindowMinutes).LocalDateTime
        $windowEnd = $incidentOffset.AddMinutes($WindowMinutes).LocalDateTime
        Write-Verbose ("Timeline window {0:yyyy-MM-dd HH:mm:ss} to {1:yyyy-MM-dd HH:mm:ss}" -f $windowStart, $windowEnd)
    }

    $all = New-Object System.Collections.Generic.List[object]
    $fileIndex = 0

    foreach ($file in $files) {
        $fileIndex++
        Write-Progress -Activity 'Reading event logs' -Status ([System.IO.Path]::GetFileName($file)) -PercentComplete (($fileIndex / $files.Count) * 100)
        $sourceSha256 = (Get-FileHash -LiteralPath $file -Algorithm SHA256 -ErrorAction Stop).Hash

        if ($isTimeline) {
            $records = Read-EventRecord -File $file -StartTime $windowStart -EndTime $windowEnd -IncludeXml -SourceSha256 $sourceSha256
        } else {
            $records = Read-EventRecord -File $file -SourceSha256 $sourceSha256
        }

        foreach ($record in $records) { [void] $all.Add($record) }
    }

    Write-Progress -Activity 'Reading event logs' -Completed
    if ($hasLog) { Write-CorrelationLog -Path $LogPath -Message ("Events read: {0} from {1} file(s)" -f $all.Count, $files.Count) }

    $deduplicated = New-Object System.Collections.Generic.List[object]
    $seenEvents = New-Object 'System.Collections.Generic.HashSet[string]'
    $duplicateCount = 0

    foreach ($item in $all) {
        $identity = '{0}|{1}|{2}|{3}|{4}|{5:o}' -f $item.MachineName, $item.Channel, $item.EventRecordId, $item.Provider, $item.Id, $item.TimeCreatedUtc
        if ($seenEvents.Add($identity)) {
            [void] $deduplicated.Add($item)
        } else {
            $duplicateCount++
        }
    }

    $filtered = $deduplicated.ToArray()
    if (-not $AllProviders) {
        $filtered = @($filtered | Where-Object { $_.Provider -match $ProviderFilter })
    }

    if ($hasLog) { Write-CorrelationLog -Path $LogPath -Message ("Duplicates removed: {0}; events after filter: {1}" -f $duplicateCount, $filtered.Count) }

    if ($filtered.Count -eq 0) {
        Write-Warning 'No events matched. Retry with -AllProviders or widen -ProviderFilter.'
        if ($hasLog) { Write-CorrelationLog -Path $LogPath -Message 'No events matched filter; run ended without output.' }
        return
    }

    $timeline = @($filtered | Sort-Object -Property TimeCreated)

    # PowerShell unwraps a single-element array on return, so each call is
    # re-wrapped; under StrictMode a bare .Count on the unwrapped object throws.
    $onset = @(Get-OnsetProfile -Record $filtered)
    $bursts = @(Get-EventBurst -Record $filtered -Threshold $BurstThreshold)
    $mismatched = @($filtered | Where-Object { $_.TemplateMismatch })

    Write-Host ''
    Write-Host ('Events read      : {0} ({1} duplicate records removed; {2} after provider filter)' -f $all.Count, $duplicateCount, $filtered.Count)
    Write-Host ('Distinct streams : {0} provider/ID combinations' -f $onset.Count)
    Write-Host ('Bursts           : {0} second(s) with >= {1} identical events' -f $bursts.Count, $BurstThreshold)
    Write-Host ('Unrenderable     : {0} event(s) whose payload does not match their publisher manifest' -f $mismatched.Count)
    Write-Host ''

    if ($mismatched.Count -gt 0) {
        Write-Host 'Unrenderable events, by provider and ID. Any description shown for these in'
        Write-Host 'Event Viewer is a fallback string and carries no meaning. Send the XML to the publisher.'
        $mismatched | Group-Object -Property Provider, Id | Sort-Object -Property Count -Descending |
            Select-Object -First 20 @{ N = 'Provider'; E = { $_.Group[0].Provider } },
                                     @{ N = 'Id'; E = { $_.Group[0].Id } },
                                     @{ N = 'Count'; E = { $_.Count } },
                                     @{ N = 'FirstSeen'; E = { ($_.Group | Sort-Object TimeCreated | Select-Object -First 1).TimeCreated } } |
            Format-Table -AutoSize | Out-Host
    }

    if ($bursts.Count -gt 0) {
        Write-Host 'Top bursts:'
        $bursts | Select-Object -First 15 | Format-Table -AutoSize SecondUtc, Machine, Channel, Provider, Id, Count | Out-Host
    }

    if (-not $isTimeline) {
        Write-Host 'Retained event ranges, ordered by earliest retained record (UTC):'
        $onset | Format-Table -AutoSize MachineName, Channel, Provider, Id, Count, ErrorOrCritical, TemplateMismatch, EarliestRetainedUtc, LatestRetainedUtc | Out-Host
    }

    if ($PSBoundParameters.ContainsKey('OutputPath') -and -not [string]::IsNullOrWhiteSpace($OutputPath)) {
        if (-not (Test-Path -LiteralPath $OutputPath -PathType Container)) {
            [void] (New-Item -Path $OutputPath -ItemType Directory -Force)
        }

        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'

        $timelinePath = Join-Path $OutputPath "Timeline-$stamp.csv"
        $timeline |
            Select-Object TimeCreated, TimeCreatedUtc, MachineName, Channel, EventRecordId, Provider, Id, LevelName, LogFile, SourceSha256, ProcessId, Unrendered, ProcessingError, ProcessingErrorCode, TemplateMismatch, TimeMissing, Summary |
            Export-Csv -LiteralPath $timelinePath -NoTypeInformation -Encoding UTF8

        $onsetPath = Join-Path $OutputPath "OnsetProfile-$stamp.csv"
        $onset | Export-Csv -LiteralPath $onsetPath -NoTypeInformation -Encoding UTF8

        if ($hasLog) { Write-CorrelationLog -Path $LogPath -Message ("Output written: {0}" -f $timelinePath) }
        if ($hasLog) { Write-CorrelationLog -Path $LogPath -Message ("Output written: {0}" -f $onsetPath) }

        if ($bursts.Count -gt 0) {
            $burstsPath = Join-Path $OutputPath "Bursts-$stamp.csv"
            $bursts | Export-Csv -LiteralPath $burstsPath -NoTypeInformation -Encoding UTF8
            if ($hasLog) { Write-CorrelationLog -Path $LogPath -Message ("Output written: {0}" -f $burstsPath) }
        }

        if ($mismatched.Count -gt 0) {
            $xmlPath = Join-Path $OutputPath "UnrenderedEvents-$stamp.xml"
            $builder = New-Object System.Text.StringBuilder
            [void] $builder.AppendLine('<UnrenderedEvents>')
            foreach ($item in $mismatched) {
                if ($null -ne $item.Xml) { [void] $builder.AppendLine($item.Xml) }
            }
            [void] $builder.AppendLine('</UnrenderedEvents>')
            [System.IO.File]::WriteAllText($xmlPath, $builder.ToString(), [System.Text.UTF8Encoding]::new($true))
            if ($hasLog) { Write-CorrelationLog -Path $LogPath -Message ("Output written: {0}" -f $xmlPath) }
        }

        Write-Host "Output written to $OutputPath"
    }

    if ($hasLog) { Write-CorrelationLog -Path $LogPath -Message 'Correlation completed.' }

    if ($PassThru) {
        $timeline
    }
} catch {
    if ($hasLog) {
        Write-CorrelationLog -Path $LogPath -Message ("ERROR: {0}" -f $_.Exception.Message)
    }
    throw
} finally {
    if (Test-Path -LiteralPath $tempRoot) {
        [System.IO.Directory]::Delete($tempRoot, $true)
    }
}
