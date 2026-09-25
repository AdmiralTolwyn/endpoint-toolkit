#Requires -Version 5.1

<#
.SYNOPSIS
    Read-only NTLM evidence detection with the Ivanti four-line output contract.
.DESCRIPTION
    Examines local NTLM Operational and Security records. Positive evidence,
    audit gaps, and no observed evidence are distinct assessments. Does not
    change authentication, auditing, services, registry values, or event logs.
    Appends diagnostic stage/count/error logging to a local file. Default host
    output remains the Ivanti contract; -Detailed selects an interactive report.
.PARAMETER Days
    Lookback in days. Defaults to 14. Retention may cover less than this window.
.PARAMETER MaxEventsPerSource
    Maximum matching records processed per source. Defaults to 10000.
.PARAMETER MaxEvidence
    Maximum recent normalized records retained for JSON and console samples.
    Command-line output displays at most three of these samples.
.PARAMETER ReportPath
    Optional JSON report path. Parent directory must exist. No JSON by default.
    Overwrites the specified file; must differ from LogPath.
.PARAMETER Detailed
    Replace the four-line Ivanti output with a compact host-only console report.
    Alias: CmdLine. Shows counts, the first collection diagnostic and three samples.
    Full diagnostics and additional evidence remain available in JSON and logs.
    Does not change collection, evidence limits or assessment semantics.
.PARAMETER LogPath
    Append-only UTF-8 diagnostic log. Defaults to
    C:\Windows\Temp\Get-NtlmUsageDetection-Ivanti.log (under the Windows directory).
    Parent directory must exist. Supply an empty string to disable file logging.
    Logging errors are reported without changing the evidence verdict.
.INPUTS
    None. Pipeline input is not supported.
.OUTPUTS
    Host/information-stream text only. Default: exactly four Ivanti lines.
    With -Detailed: a human-readable report. No success-stream objects.
.EXAMPLE
    .\Get-NtlmUsageDetection-Ivanti.ps1

    Collect evidence, append diagnostic logging, and emit the four Ivanti lines.
.EXAMPLE
    .\Get-NtlmUsageDetection-Ivanti.ps1 -Days 30 -ReportPath C:\Temp\NtlmEvidence.json

    Collect a 30-day window and write the optional detailed JSON report.
.EXAMPLE
    .\Get-NtlmUsageDetection-Ivanti.ps1 -Detailed -Days 7

    Display the interactive report while retaining normal diagnostic logging.
.EXAMPLE
    .\Get-NtlmUsageDetection-Ivanti.ps1 -CmdLine -LogPath ''

    Display the compact console report without creating or appending a log file.
.NOTES
    Run locally in 64-bit Windows PowerShell 5.1 as SYSTEM or administrator.
    Reports observed event records, current settings and collection errors.
    Default: four Write-Host lines: detected, reason, expected, found.
    detected=true means matching NTLM or derived-credential records were found.
    Collection diagnostics and file errors do not set detected. Full Assessment
    is available with -CmdLine or JSON.
    Handled findings/errors do not use Intune exit 1.
    No network requests or remediation. Default diagnostic log contains stage
    counts, configuration, errors and assessment, not raw event/account samples.
    Protect log and report paths; there is no automatic log rotation or deletion.
    ReportPath overwrites the specified file; protect account and host details.
    Parameter/host failures before entry may prevent both logging and output.
    Version: 1.2. Date: 2026-09-22. Windows PowerShell 5.1 baseline.
    Documentation: README.md in this directory.
    Offline tests: Test-NtlmUsageDetection.ps1 in this directory.
.LINK
    https://techcommunity.microsoft.com/blog/windows-itpro-blog/retiring-ntlm-frequently-asked-questions/4550522
.LINK
    https://support.microsoft.com/en-us/servicing/os/windows/2025/07/overview-of-ntlm-auditing-enhancements-in-windows-11-version-24h2-and-windows-server-2025
#>
[CmdletBinding()]
param(
    [ValidateRange(1, 365)]
    [int]$Days = 14,
    [ValidateRange(1, 100000)]
    [int]$MaxEventsPerSource = 10000,
    [ValidateRange(0, 10000)]
    [int]$MaxEvidence = 200,
    [string]$ReportPath,
    [Alias('CmdLine')]
    [switch]$Detailed,
    [AllowEmptyString()]
    [string]$LogPath = "$env:windir\Temp\Get-NtlmUsageDetection-Ivanti.log"
)

function Write-NtlmLog {
    <#
    .SYNOPSIS
        Appends one diagnostic line without emitting console or pipeline output.
    .DESCRIPTION
        Uses the current script-scoped NtlmLogState initialized by the entry point.
        Each UTF-8 line contains UTC time, run ID, PID and severity. Sanitizes
        control characters; stops further writes after the first failure and
        retains that error in Logging. Missing/disabled state is a no-op.
    .PARAMETER Message
        Stage, aggregate, policy or error text. Do not pass raw event samples.
    .PARAMETER Level
        INFO, WARN or ERROR. Defaults to INFO.
    .OUTPUTS
        None. Updates EntriesWritten or Error on the shared logging state.
    .NOTES
        Never creates parent directories, rotates files, or changes the detection
        assessment. Logging failures are available through -CmdLine and JSON.
    #>
    param([string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO')
    if ($null -eq $script:NtlmLogState -or -not $script:NtlmLogState.Enabled -or $script:NtlmLogState.Error) { return }
    try {
        $cleanMessage = $Message -replace '[\x00-\x1F\x7F]', ' '
        $line = '{0} [{1}] [Run:{2}] [PID:{3}] {4}{5}' -f [datetime]::UtcNow.ToString('o'), $Level, $script:NtlmLogState.RunId, $PID, $cleanMessage, [Environment]::NewLine
        [IO.File]::AppendAllText($script:NtlmLogState.Path, $line, (New-Object System.Text.UTF8Encoding($true)))
        $script:NtlmLogState.EntriesWritten++
    }
    catch { $script:NtlmLogState.Error = $_.Exception.Message }
}

function Get-NtlmField {
    <#
    .SYNOPSIS
        Returns the first populated named event field.
    .DESCRIPTION
        Checks alternate schema field names in priority order without interpreting
        their values. Missing fields remain null rather than becoming zero.
    .PARAMETER Data
        Case-insensitive dictionary of raw named EventData values.
    .PARAMETER Names
        Candidate field names, most specific first.
    .OUTPUTS
        System.String, or null when no candidate has a nonblank value.
    #>
    param([System.Collections.IDictionary]$Data, [string[]]$Names)
    foreach ($name in $Names) {
        if ($Data.Contains($name) -and -not [string]::IsNullOrWhiteSpace([string]$Data[$name])) {
            return [string]$Data[$name]
        }
    }
    return $null
}

function ConvertFrom-NtlmEvent {
    <#
    .SYNOPSIS
        Normalizes one supported NTLM event from named XML fields.
    .DESCRIPTION
        Validates provider, channel and ID before classifying evidence. Preserves
        raw fields and separates attempts, blocks, status results and derived
        credentials. Warning level is not used to infer NTLMv1. Anonymous logon
        version labels are preserved only in raw data. Does not render messages.
    .PARAMETER Xml
        EventRecord.ToXml() output. External XML resolution is disabled.
    .OUTPUTS
        PSCustomObject for recognized evidence; no output for unrelated events.
    .NOTES
        Malformed XML or missing required system fields throws to the collector,
        which records a parsing gap rather than manufacturing negative evidence.
    #>
    param([Parameter(Mandatory = $true)][string]$Xml)
    $document = New-Object System.Xml.XmlDocument
    $document.XmlResolver = $null
    $document.LoadXml($Xml)
    $eventNode = $document.DocumentElement
    $system = $eventNode.SelectSingleNode("*[local-name()='System']")
    $eventId = [int]$system.SelectSingleNode("*[local-name()='EventID']").InnerText
    $provider = $system.SelectSingleNode("*[local-name()='Provider']").GetAttribute('Name')
    $channel = $system.SelectSingleNode("*[local-name()='Channel']").InnerText
    $data = [ordered]@{}
    foreach ($node in $eventNode.SelectNodes("*[local-name()='EventData']/*[local-name()='Data']")) {
        $name = $node.GetAttribute('Name')
        if ($name) { $data[$name] = $node.InnerText }
    }

    $kind = $null
    $direction = 'Unknown'
    $outcome = 'Attempt'
    $version = 'Unknown'
    $isWarning = $false
    if ($channel -eq 'Microsoft-Windows-NTLM/Operational') {
        if ($provider -eq 'Microsoft-Windows-NTLM' -and $eventId -in @(4020, 4021, 4022, 4023)) {
            $kind = 'EnhancedNtlm'
            $direction = 'Outgoing'
            if ($eventId -in @(4022, 4023)) { $direction = 'Incoming' }
        }
        elseif ($provider -eq 'Microsoft-Windows-Security-Netlogon' -and $eventId -in @(4030, 4031, 4032, 4033)) {
            $kind = 'EnhancedNtlm'
            $direction = 'DomainValidation'
        }
        elseif ($provider -eq 'Microsoft-Windows-NTLM' -and $eventId -in @(4001, 4002, 4003, 8001, 8002, 8003)) {
            $kind = 'LegacyNtlm'
            $direction = 'Incoming'
            if ($eventId -in @(4001, 8001)) { $direction = 'Outgoing' }
            if ($eventId -in @(4003, 8003)) { $direction = 'DomainServer' }
            if ($eventId -lt 8000) { $outcome = 'Blocked' }
        }
        elseif ($provider -eq 'Microsoft-Windows-Security-Netlogon' -and $eventId -in @(4004, 8004)) {
            $kind = 'LegacyNtlm'
            $direction = 'DomainValidation'
            if ($eventId -eq 4004) { $outcome = 'Blocked' }
        }
        elseif ($provider -eq 'Microsoft-Windows-NTLM' -and $eventId -in @(4024, 4025)) {
            $kind = 'NtlmV1DerivedCredentials'
            $direction = 'Outgoing'
            if ($eventId -eq 4025) { $outcome = 'Blocked' }
        }
        elseif ($provider -eq 'Microsoft-Windows-NTLM' -and $eventId -in @(4026, 4027)) {
            $kind = 'EnhancedNtlmPolicy'
            $direction = 'Outgoing'
            if ($eventId -eq 4027) { $outcome = 'Blocked' }
        }
        else { return }
        $isWarning = $eventId -in @(4021, 4023, 4031, 4033)
        if ($kind -eq 'EnhancedNtlm' -and $direction -ne 'Outgoing') {
            $status = Get-NtlmField $data @('Status')
            $outcome = 'Unknown'
            if ($status -match '^(0x)?0+$') { $outcome = 'Success' }
            elseif ($status -match '^(0x[0-9a-fA-F]+|[0-9]+)$') { $outcome = 'NonSuccessStatus' }
        }
        $versionText = Get-NtlmField $data @('NtlmVersion')
        if ($versionText -match '^NTLM\s*V?1$') { $version = 'NTLMv1' }
        elseif ($versionText -match '^NTLM\s*V?2$') { $version = 'NTLMv2' }
    }
    elseif ($channel -eq 'Security' -and $provider -eq 'Microsoft-Windows-Security-Auditing') {
        if ($eventId -in @(4624, 4625)) {
            if ((Get-NtlmField $data @('AuthenticationPackageName')) -ne 'NTLM') { return }
            $kind = 'SecurityLogon'
            $direction = 'LocalLogon'
            $outcome = 'Failure'
            if ($eventId -eq 4624) { $outcome = 'Success' }
            $versionText = Get-NtlmField $data @('LmPackageName')
            if ($versionText -match '^NTLM\s*V?1$') { $version = 'NTLMv1' }
            elseif ($versionText -match '^NTLM\s*V?2$') { $version = 'NTLMv2' }
            elseif ($versionText -eq 'LM') { $version = 'LM' }
            if ((Get-NtlmField $data @('TargetUserSid')) -eq 'S-1-5-7' -or
                (Get-NtlmField $data @('TargetUserName')) -eq 'ANONYMOUS LOGON') { $version = 'Unknown' }
        }
        elseif ($eventId -eq 4776 -and (Get-NtlmField $data @('PackageName')) -eq 'MICROSOFT_AUTHENTICATION_PACKAGE_V1_0') {
            $kind = 'CredentialValidation'
            $direction = 'CredentialValidation'
            $status = Get-NtlmField $data @('Status')
            $outcome = 'Unknown'
            if ($status -match '^(0x)?0+$') { $outcome = 'Success' }
            elseif ($status -match '^(0x[0-9a-fA-F]+|[0-9]+)$') { $outcome = 'Failure' }
        }
        else { return }
    }
    else { return }

    [pscustomobject][ordered]@{
        TimeUtc = ([datetimeoffset]::Parse($system.SelectSingleNode("*[local-name()='TimeCreated']").GetAttribute('SystemTime'))).UtcDateTime.ToString('o')
        RecordId = [long]$system.SelectSingleNode("*[local-name()='EventRecordID']").InnerText
        Computer = $system.SelectSingleNode("*[local-name()='Computer']").InnerText
        Channel = $channel
        Provider = $provider
        EventId = $eventId
        Kind = $kind
        Direction = $direction
        Outcome = $outcome
        NtlmVersion = $version
        SecurityWarning = $isWarning
        Account = Get-NtlmField $data @('Username', 'TargetUserName', 'AccountName')
        Domain = Get-NtlmField $data @('DomainName', 'TargetDomainName', 'AccountDomain')
        Process = Get-NtlmField $data @('ProcessName')
        ProcessId = Get-NtlmField $data @('ProcessPID', 'ProcessId', 'CallerPID')
        Target = Get-NtlmField $data @('TargetService', 'TargetName', 'ServerName', 'TargetMachine', 'SChannelName')
        TargetIp = Get-NtlmField $data @('TargetIP', 'ServerIP')
        Client = Get-NtlmField $data @('RemoteClientMachine', 'WorkstationName', 'Workstation', 'Hostname', 'AccountMachine')
        ClientIp = Get-NtlmField $data @('ClientIP', 'IpAddress')
        LogonType = Get-NtlmField $data @('LogonType')
        UsageId = Get-NtlmField $data @('NtlmUsageId')
        UsageReason = Get-NtlmField $data @('NtlmUsageReason')
        Status = Get-NtlmField $data @('Status')
        ChannelBinding = Get-NtlmField $data @('ChannelBindingStatus')
        MicStatus = Get-NtlmField $data @('Mic Status', 'MicStatus')
        Data = $data
    }
}

function Get-NtlmEventBatch {
    <#
    .SYNOPSIS
        Reads a bounded event batch while preserving partial results.
    .DESCRIPTION
        Requests one additional record to distinguish a full batch from actual
        truncation. NoMatchingEventsFound is an empty result, not an access error.
    .PARAMETER LogName
        Local Windows event channel to read.
    .PARAMETER XPath
        Windows Event Log XPath filter; defaults to every record.
    .PARAMETER Limit
        Maximum returned records, excluding the truncation probe.
    .PARAMETER Oldest
        Read oldest first, used to inspect retained history.
    .OUTPUTS
        PSCustomObject with Events, Truncated and Error, including partial reads.
    .NOTES
        Disposes the extra probe record. The caller must dispose returned records.
        Bounds record count, not elapsed execution time.
    #>
    param([string]$LogName, [string]$XPath = '*', [int]$Limit, [switch]$Oldest)
    $records = New-Object 'System.Collections.Generic.List[object]'
    $errorText = $null
    try {
        Get-WinEvent -LogName $LogName -FilterXPath $XPath -MaxEvents ($Limit + 1) -Oldest:$Oldest -ErrorAction Stop |
            ForEach-Object { [void]$records.Add($_) }
    }
    catch {
        if ($_.FullyQualifiedErrorId -notlike 'NoMatchingEventsFound*') { $errorText = $_.Exception.Message }
    }
    $truncated = $records.Count -gt $Limit
    if ($truncated) {
        $records[$Limit].Dispose()
        $records.RemoveAt($Limit)
    }
    [pscustomobject]@{ Events = $records.ToArray(); Truncated = $truncated; Error = $errorText }
}

function Get-NtlmLogState {
    <#
    .SYNOPSIS
        Inspects channel configuration and its oldest retained timestamp.
    .DESCRIPTION
        Reads even a disabled channel's retained data. Missing metadata and access
        errors remain explicit. Retention is not proof of continuous auditing.
    .PARAMETER Name
        Local Windows event channel name.
    .PARAMETER StartUtc
        Requested collection-window start in UTC.
    .OUTPUTS
        PSCustomObject containing configuration, retention and read-error details.
    #>
    param([string]$Name, [datetime]$StartUtc)
    $result = [ordered]@{
        Name = $Name; Enabled = $null; RecordCount = $null; MaximumSizeBytes = $null
        LogMode = $null; OldestRetainedUtc = $null; RetentionCoversWindow = $false; Error = $null
    }
    try {
        $configuration = Get-WinEvent -ListLog $Name -ErrorAction Stop
        $result.Enabled = $configuration.IsEnabled
        $result.RecordCount = $configuration.RecordCount
        $result.MaximumSizeBytes = $configuration.MaximumSizeInBytes
        $result.LogMode = [string]$configuration.LogMode
        $oldestBatch = Get-NtlmEventBatch -LogName $Name -Limit 1 -Oldest
        if ($oldestBatch.Error) { $result.Error = $oldestBatch.Error }
        foreach ($record in $oldestBatch.Events) {
            try {
                if ($null -ne $record.TimeCreated) {
                    $oldestUtc = $record.TimeCreated.ToUniversalTime()
                    $result.OldestRetainedUtc = $oldestUtc.ToString('o')
                    $result.RetentionCoversWindow = $oldestUtc -le $StartUtc
                }
            }
            finally { $record.Dispose() }
        }
    }
    catch { $result.Error = $_.Exception.Message }
    [pscustomobject]$result
}

function Get-NtlmRegistryValue {
    <#
    .SYNOPSIS
        Reads one policy value from the 64-bit HKLM view.
    .DESCRIPTION
        Opens keys read-only and distinguishes absent configuration from denied
        access. Registry handles are disposed on success and failure.
    .PARAMETER Path
        Subkey relative to HKEY_LOCAL_MACHINE, without a hive prefix.
    .PARAMETER Name
        Registry value name.
    .OUTPUTS
        PSCustomObject with Path, Name, State, Value and Error.
    #>
    param([string]$Path, [string]$Name)
    $baseKey = $null
    $key = $null
    $result = [ordered]@{ Path = 'HKLM\' + $Path; Name = $Name; State = 'NotConfigured'; Value = $null; Error = $null }
    try {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
        $key = $baseKey.OpenSubKey($Path, $false)
        if ($null -ne $key -and $key.GetValueNames() -contains $Name) {
            $result.Value = $key.GetValue($Name)
            $result.State = 'Configured'
        }
    }
    catch { $result.State = 'Unreadable'; $result.Error = $_.Exception.Message }
    finally {
        if ($null -ne $key) { $key.Dispose() }
        if ($null -ne $baseKey) { $baseKey.Dispose() }
    }
    [pscustomobject]$result
}

function Get-NtlmPolicySnapshot {
    <#
    .SYNOPSIS
        Collects current enhanced and legacy NTLM policy values.
    .DESCRIPTION
        Includes audit prerequisites and restriction context. Policy state alone
        does not establish actual usage or historical audit coverage.
    .OUTPUTS
        One registry-state PSCustomObject per inspected value.
    #>
    $definitions = @(
        @('SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\NTLM\Parameters', 'LogEnhancedAuditEvents'),
        @('SOFTWARE\Policies\Microsoft\Netlogon\Parameters', 'EnableEnhancedDomainNtlmLogs'),
        @('SYSTEM\CurrentControlSet\Control\Lsa\MSV1_0', 'AuditReceivingNTLMTraffic'),
        @('SYSTEM\CurrentControlSet\Control\Lsa\MSV1_0', 'RestrictSendingNTLMTraffic'),
        @('SYSTEM\CurrentControlSet\Control\Lsa\MSV1_0', 'RestrictReceivingNTLMTraffic'),
        @('SYSTEM\CurrentControlSet\Services\Netlogon\Parameters', 'AuditNTLMInDomain'),
        @('SYSTEM\CurrentControlSet\Services\Netlogon\Parameters', 'RestrictNTLMInDomain'),
        @('SYSTEM\CurrentControlSet\Control\Lsa', 'LmCompatibilityLevel'),
        @('SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\NTLM\Parameters', 'EnhancedNtlmBlocks')
    )
    foreach ($definition in $definitions) { Get-NtlmRegistryValue -Path $definition[0] -Name $definition[1] }
}

function Get-NtlmAuditPolicy {
    <#
    .SYNOPSIS
        Queries effective logon and credential-validation audit settings.
    .DESCRIPTION
        Calls auditpol /get /r with stable subcategory GUIDs. Parses CSV numeric
        settings without localized column names. Unknown formats and nonzero
        native exit codes return an explicit error, never an assumed default.
    .OUTPUTS
        PSCustomObject with State, Logon, CredentialValidation and Error.
    .NOTES
        Numeric flags: 0 none, 1 success, 2 failure, 3 success and failure.
        Requires permission to query audit policy; performs no policy changes.
    #>
    $subcategoryIds = @('{0cce9215-69ae-11d9-bed3-505054503030}', '{0cce923f-69ae-11d9-bed3-505054503030}')
    $result = [ordered]@{ State = 'Unknown'; Logon = $null; CredentialValidation = $null; Error = $null }
    try {
        $text = @(& "$env:windir\System32\auditpol.exe" /get (('/subcategory:' + ($subcategoryIds -join ','))) /r 2>&1)
        if ($LASTEXITCODE -ne 0) { throw "auditpol exit $LASTEXITCODE : $($text -join ' ')" }
        $rows = @($text | ConvertFrom-Csv -ErrorAction Stop)
        $settings = @{}
        foreach ($row in $rows) {
            $values = @($row.PSObject.Properties | ForEach-Object { [string]$_.Value })
            foreach ($subcategoryId in $subcategoryIds) {
                if ($values -contains $subcategoryId) {
                    $numericValue = 0
                    if (-not [int]::TryParse($values[-1], [ref]$numericValue) -or $numericValue -lt 0 -or $numericValue -gt 3) {
                        throw 'auditpol CSV did not contain a recognized numeric setting value.'
                    }
                    $settings[$subcategoryId] = $numericValue
                }
            }
        }
        if ($settings.Count -ne 2) { throw 'auditpol CSV did not return both required subcategory GUIDs.' }
        $result.Logon = $settings[$subcategoryIds[0]]
        $result.CredentialValidation = $settings[$subcategoryIds[1]]
        $result.State = 'Read'
    }
    catch { $result.Error = $_.Exception.Message }
    [pscustomobject]$result
}

function Get-NtlmProviderState {
    <#
    .SYNOPSIS
        Reads installed enhanced NTLM and Netlogon event schemas.
    .DESCRIPTION
        Inspects provider metadata only. Registered IDs do not prove controlled
        feature rollout activation or that enhanced events have been emitted.
    .OUTPUTS
        PSCustomObject with client/server IDs, domain IDs and metadata errors.
    #>
    $result = [ordered]@{ ClientServerEventIds = @(); DomainEventIds = @(); Errors = @() }
    foreach ($providerName in @('Microsoft-Windows-NTLM', 'Microsoft-Windows-Security-Netlogon')) {
        try {
            $metadata = Get-WinEvent -ListProvider $providerName -ErrorAction Stop
            $ids = @($metadata.Events | Where-Object { $_.Id -in @(4020, 4021, 4022, 4023, 4030, 4031, 4032, 4033) } |
                Select-Object -ExpandProperty Id -Unique)
            if ($providerName -eq 'Microsoft-Windows-NTLM') { $result.ClientServerEventIds = $ids }
            else { $result.DomainEventIds = $ids }
        }
        catch { $result.Errors += "$providerName : $($_.Exception.Message)" }
    }
    [pscustomobject]$result
}

function Get-NtlmQueryDefinitions {
    <#
    .SYNOPSIS
        Builds the five time-bounded local event queries.
    .DESCRIPTION
        Filters Security logons by the NTLM authentication package in the event
        service. Separates evidence sources from audit-change context queries.
    .PARAMETER StartUtc
        Inclusive start of the requested window, supplied as UTC.
    .PARAMETER EndUtc
        Inclusive end of the requested window, supplied as UTC.
    .OUTPUTS
        Query PSCustomObjects containing name, channel, XPath and evidence flag.
    #>
    param([datetime]$StartUtc, [datetime]$EndUtc)
    $timeFilter = "TimeCreated[@SystemTime >= '$($StartUtc.ToString('yyyy-MM-ddTHH:mm:ss.fffffffZ'))' and @SystemTime <= '$($EndUtc.ToString('yyyy-MM-ddTHH:mm:ss.fffffffZ'))']"
    $ntlmIds = @(4001, 4002, 4003, 4004, 4020, 4021, 4022, 4023, 4024, 4025, 4026, 4027, 4030, 4031, 4032, 4033, 8001, 8002, 8003, 8004)
    $idFilter = ($ntlmIds | ForEach-Object { "EventID=$_" }) -join ' or '
    [pscustomobject]@{ Name = 'NtlmOperational'; LogName = 'Microsoft-Windows-NTLM/Operational'; XPath = "*[System[($idFilter) and $timeFilter]]"; Evidence = $true }
    [pscustomobject]@{ Name = 'SecurityLogon'; LogName = 'Security'; XPath = "*[System[Provider[@Name='Microsoft-Windows-Security-Auditing'] and (EventID=4624 or EventID=4625) and $timeFilter] and EventData[Data[@Name='AuthenticationPackageName']='NTLM']]"; Evidence = $true }
    [pscustomobject]@{ Name = 'CredentialValidation'; LogName = 'Security'; XPath = "*[System[Provider[@Name='Microsoft-Windows-Security-Auditing'] and EventID=4776 and $timeFilter]]"; Evidence = $true }
    [pscustomobject]@{ Name = 'SecurityAuditChanges'; LogName = 'Security'; XPath = "*[System[(EventID=1102 or EventID=4719) and $timeFilter]]"; Evidence = $false }
    [pscustomobject]@{ Name = 'LogClears'; LogName = 'System'; XPath = "*[System[Provider[@Name='Microsoft-Windows-Eventlog'] and EventID=104 and $timeFilter]]"; Evidence = $false }
}

function Get-NtlmCollection {
    <#
    .SYNOPSIS
        Collects and normalizes bounded evidence and audit-change records.
    .DESCRIPTION
        Keeps source errors, unclassified records and parse failures separate
        from event evidence. Disposes every record and preserves positive evidence
        from successful sources when another source fails.
    .PARAMETER StartUtc
        Inclusive collection-window start in UTC.
    .PARAMETER EndUtc
        Inclusive collection-window end in UTC.
    .PARAMETER Limit
        Maximum processed records for each independent query.
    .OUTPUTS
        PSCustomObject with Events, Sources and AuditMarkers arrays.
    #>
    param([datetime]$StartUtc, [datetime]$EndUtc, [int]$Limit)
    $events = New-Object 'System.Collections.Generic.List[object]'
    $sources = New-Object 'System.Collections.Generic.List[object]'
    $markers = New-Object 'System.Collections.Generic.List[object]'
    foreach ($query in Get-NtlmQueryDefinitions -StartUtc $StartUtc -EndUtc $EndUtc) {
        Write-NtlmLog ("Query start: Source=$($query.Name); Channel=$($query.LogName); Limit=$Limit")
        $batch = Get-NtlmEventBatch -LogName $query.LogName -XPath $query.XPath -Limit $Limit
        $parseErrors = 0
        $unclassified = 0
        $firstParseError = $null
        foreach ($record in $batch.Events) {
            try {
                if ($query.Evidence) {
                    $evidence = ConvertFrom-NtlmEvent -Xml $record.ToXml()
                    if ($null -ne $evidence) { [void]$events.Add($evidence) }
                    else { $unclassified++ }
                }
                else {
                    $xml = $record.ToXml()
                    if ($query.Name -eq 'LogClears') {
                        $markerDocument = New-Object System.Xml.XmlDocument
                        $markerDocument.XmlResolver = $null
                        $markerDocument.LoadXml($xml)
                        $clearedChannel = $markerDocument.SelectSingleNode("/*/*[local-name()='UserData']//*[local-name()='Channel']")
                        if ($null -ne $clearedChannel -and $clearedChannel.InnerText -notin @('Security', 'Microsoft-Windows-NTLM/Operational')) { continue }
                    }
                    [void]$markers.Add([pscustomobject]@{
                        Channel = $query.LogName; EventId = $record.Id; RecordId = $record.RecordId
                        TimeUtc = $record.TimeCreated.ToUniversalTime().ToString('o'); Xml = $xml
                    })
                }
            }
            catch {
                $parseErrors++
                if (-not $firstParseError) { $firstParseError = $_.Exception.Message }
            }
            finally { $record.Dispose() }
        }
        [void]$sources.Add([pscustomobject]@{
            Name = $query.Name; LogName = $query.LogName; RecordsRead = $batch.Events.Count
            Truncated = $batch.Truncated; Error = $batch.Error; ParseErrors = $parseErrors
            FirstParseError = $firstParseError; UnclassifiedRecords = $unclassified
        })
        $logLevel = 'INFO'
        if ($batch.Error -or $batch.Truncated -or $parseErrors -gt 0 -or $unclassified -gt 0) { $logLevel = 'WARN' }
        Write-NtlmLog -Level $logLevel -Message ('Query complete: ' + ($sources[$sources.Count - 1] | ConvertTo-Json -Compress))
    }
    [pscustomobject]@{ Events = $events.ToArray(); Sources = $sources.ToArray(); AuditMarkers = $markers.ToArray() }
}

function Get-NtlmMachineState {
    <#
    .SYNOPSIS
        Collects the local execution context and Windows role.
    .DESCRIPTION
        Reports elevation without attempting it. Uses one bounded CIM metadata
        request; ProductType distinguishes workstation, server and DC roles.
    .OUTPUTS
        PSCustomObject with host, PowerShell, bitness, elevation and OS metadata.
    #>
    $principal = New-Object Security.Principal.WindowsPrincipal ([Security.Principal.WindowsIdentity]::GetCurrent())
    $result = [ordered]@{
        Computer = $env:COMPUTERNAME; PowerShellVersion = $PSVersionTable.PSVersion.ToString()
        Is64BitProcess = [Environment]::Is64BitProcess
        Elevated = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        Caption = $null; Version = $null; Build = $null; ProductType = $null; Error = $null
    }
    try {
        $operatingSystem = Get-CimInstance Win32_OperatingSystem -OperationTimeoutSec 10 -ErrorAction Stop
        $result.Caption = $operatingSystem.Caption
        $result.Version = $operatingSystem.Version
        $result.Build = $operatingSystem.BuildNumber
        $result.ProductType = $operatingSystem.ProductType
    }
    catch { $result.Error = $_.Exception.Message }
    [pscustomobject]$result
}

function Get-NtlmTopValues {
    <#
    .SYNOPSIS
        Counts populated field values across all processed evidence.
    .DESCRIPTION
        Excludes empty and placeholder values and sorts by descending count then
        name. These counts are event records, not distinct authentications.
    .PARAMETER Events
        All normalized records, independent of the evidence-sample limit.
    .PARAMETER Property
        Normalized property to group by.
    .PARAMETER Limit
        Maximum groups returned, default 10.
    .OUTPUTS
        PSCustomObjects with Value and Records; no output for an empty group set.
    #>
    param([object[]]$Events, [string]$Property, [int]$Limit = 10)
    @($Events | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.$Property) -and [string]$_.$Property -ne '-' } |
        Group-Object -Property $Property | Sort-Object -Property @{ Expression = 'Count'; Descending = $true }, Name |
        Select-Object -First $Limit | ForEach-Object { [pscustomobject]@{ Value = $_.Name; Records = $_.Count } })
}

function Get-NtlmAssessment {
    <#
    .SYNOPSIS
        Evaluates observed evidence separately from audit-coverage gaps.
    .DESCRIPTION
        Positive NTLM evidence takes precedence over gaps; derived-credential-only
        evidence remains separate. NoEvidenceObserved requires zero matching
        records and no recognized collection or audit-configuration findings.
    .PARAMETER Collection
        Normalized Events, per-source status and AuditMarkers.
    .PARAMETER Logs
        Channel configuration and retention observations.
    .PARAMETER Policies
        Current 64-bit HKLM policy snapshot.
    .PARAMETER AuditPolicy
        Effective Security audit subcategory settings or an explicit read error.
    .PARAMETER Providers
        Registered enhanced event IDs and provider-read errors.
    .PARAMETER Machine
        OS role and execution context used to apply DC-specific prerequisites.
    .OUTPUTS
        PSCustomObject containing assessment, evidence flag, reason, gaps and
        counters. Detected is true only for matching NTLM or derived credentials.
    #>
    param([object]$Collection, [object[]]$Logs, [object[]]$Policies, [object]$AuditPolicy, [object]$Providers, [object]$Machine)
    $gaps = New-Object 'System.Collections.Generic.List[string]'
    if ($Machine.Error) { [void]$gaps.Add('OS metadata unavailable: ' + $Machine.Error) }
    foreach ($log in $Logs) {
        if ($log.Error) { [void]$gaps.Add($log.Name + ' metadata/read error: ' + $log.Error) }
        if ($log.Enabled -ne $true) { [void]$gaps.Add($log.Name + ' is disabled or its state is unknown.') }
        if ($log.Name -ne 'System' -and -not $log.RetentionCoversWindow) { [void]$gaps.Add("$($log.Name) retention: RetentionCoversWindow=False; OldestRetainedUtc=$($log.OldestRetainedUtc).") }
    }
    foreach ($source in $Collection.Sources) {
        if ($source.Error) { [void]$gaps.Add($source.Name + ' query error: ' + $source.Error) }
        if ($source.Truncated) { [void]$gaps.Add($source.Name + ' reached the record limit; counts are lower bounds.') }
        if ($source.ParseErrors -gt 0) { [void]$gaps.Add($source.Name + ' has unparsed records: ' + $source.ParseErrors) }
        if ($source.UnclassifiedRecords -gt 0) { [void]$gaps.Add($source.Name + ' has unclassified candidate records: ' + $source.UnclassifiedRecords) }
    }
    $settings = @{}
    foreach ($policy in $Policies) {
        $settings[$policy.Name] = $policy
        if ($policy.State -eq 'Unreadable') { [void]$gaps.Add('Unreadable policy: ' + $policy.Name) }
    }
    $hasClientSchemas = @(4020, 4021, 4022, 4023 | Where-Object { $Providers.ClientServerEventIds -notcontains $_ }).Count -eq 0
    $clientPolicy = $settings['LogEnhancedAuditEvents']
    $enhancedObserved = @($Collection.Events | Where-Object { $_.EventId -in @(4020, 4021, 4022, 4023) }).Count -gt 0
    $enhancedReady = $hasClientSchemas -and $clientPolicy.State -ne 'Unreadable' -and
        (($clientPolicy.State -eq 'Configured' -and $clientPolicy.Value -eq 1) -or
        ($clientPolicy.State -eq 'NotConfigured' -and $enhancedObserved))
    $legacyOutgoing = $settings['RestrictSendingNTLMTraffic']
    $legacyIncoming = $settings['AuditReceivingNTLMTraffic']
    if (-not $enhancedReady -and -not ($legacyOutgoing.State -eq 'Configured' -and $legacyOutgoing.Value -eq 1)) {
        [void]$gaps.Add("Outgoing audit settings: EnhancedSchemas=$hasClientSchemas; LogEnhancedAuditEvents=$($clientPolicy.State)/$($clientPolicy.Value); EnhancedEventsObserved=$enhancedObserved; RestrictSendingNTLMTraffic=$($legacyOutgoing.State)/$($legacyOutgoing.Value).")
    }
    if (-not $enhancedReady -and -not ($legacyIncoming.State -eq 'Configured' -and $legacyIncoming.Value -eq 2)) {
        [void]$gaps.Add("Incoming audit settings: EnhancedSchemas=$hasClientSchemas; LogEnhancedAuditEvents=$($clientPolicy.State)/$($clientPolicy.Value); EnhancedEventsObserved=$enhancedObserved; AuditReceivingNTLMTraffic=$($legacyIncoming.State)/$($legacyIncoming.Value).")
    }
    if ($Machine.ProductType -eq 2) {
        $domainPolicy = $settings['EnableEnhancedDomainNtlmLogs']
        $domainSchemas = @(4030, 4031, 4032, 4033 | Where-Object { $Providers.DomainEventIds -notcontains $_ }).Count -eq 0
        $domainObserved = @($Collection.Events | Where-Object { $_.EventId -in @(4030, 4031, 4032, 4033) }).Count -gt 0
        $domainReady = $domainSchemas -and (($domainPolicy.State -eq 'Configured' -and $domainPolicy.Value -eq 1) -or
            ($domainPolicy.State -eq 'NotConfigured' -and $domainObserved))
        $domainLegacy = $settings['AuditNTLMInDomain']
        if (-not $domainReady -and -not ($domainLegacy.State -eq 'Configured' -and $domainLegacy.Value -eq 7)) {
            [void]$gaps.Add("Domain audit settings: EnhancedSchemas=$domainSchemas; EnableEnhancedDomainNtlmLogs=$($domainPolicy.State)/$($domainPolicy.Value); EnhancedEventsObserved=$domainObserved; AuditNTLMInDomain=$($domainLegacy.State)/$($domainLegacy.Value).")
        }
    }
    if ($AuditPolicy.State -ne 'Read') { [void]$gaps.Add('Security audit policy could not be read: ' + $AuditPolicy.Error) }
    elseif ($AuditPolicy.Logon -ne 3 -or $AuditPolicy.CredentialValidation -ne 3) {
        [void]$gaps.Add('Security logon and credential-validation auditing are not both Success and Failure.')
    }
    if ($Collection.AuditMarkers.Count -gt 0) {
        [void]$gaps.Add("Log-clear or audit-policy-change records in the window: $($Collection.AuditMarkers.Count).")
    }
    $ntlm = @($Collection.Events | Where-Object { $_.Kind -ne 'NtlmV1DerivedCredentials' })
    $derived = @($Collection.Events | Where-Object { $_.Kind -eq 'NtlmV1DerivedCredentials' })
    $assessment = 'NoEvidenceObserved'
    $detected = $false
    $reason = 'No matching NTLM records found.'
    if ($ntlm.Count -gt 0) {
        $assessment = 'NtlmEvidenceObserved'
        $detected = $true
        $reason = "Matching NTLM records: $($ntlm.Count)."
    }
    elseif ($derived.Count -gt 0) {
        $assessment = 'DerivedCredentialEvidenceOnly'
        $detected = $true
        $reason = "NTLMv1-derived SSO credential records: $($derived.Count)."
    }
    elseif ($gaps.Count -gt 0) {
        $assessment = 'InsufficientEvidence'
        $reason = "No matching NTLM records found. Collection diagnostics: $($gaps.Count)."
    }
    [pscustomobject]@{
        Assessment = $assessment; Detected = $detected; Reason = $reason; Gaps = $gaps.ToArray()
        NtlmRecords = $ntlm.Count; DerivedCredentialRecords = $derived.Count
        NtlmV1Records = @($ntlm | Where-Object { $_.NtlmVersion -eq 'NTLMv1' }).Count
        LmRecords = @($ntlm | Where-Object { $_.NtlmVersion -eq 'LM' }).Count
        EnhancedWarnings = @($ntlm | Where-Object { $_.SecurityWarning }).Count
        SuccessfulLogonRecords = @($ntlm | Where-Object { $_.EventId -eq 4624 }).Count
        SuccessfulValidationRecords = @($ntlm | Where-Object { $_.Kind -eq 'CredentialValidation' -and $_.Outcome -eq 'Success' }).Count
        EnhancedSuccessRecords = @($ntlm | Where-Object { $_.Kind -eq 'EnhancedNtlm' -and $_.Outcome -eq 'Success' }).Count
        BlockedRecords = @($ntlm | Where-Object { $_.Outcome -eq 'Blocked' }).Count
        AttemptRecords = @($ntlm | Where-Object { $_.Outcome -eq 'Attempt' }).Count
        FailureOrNonSuccessRecords = @($ntlm | Where-Object { $_.Outcome -in @('Failure', 'NonSuccessStatus') }).Count
    }
}

function ConvertTo-NtlmSingleLine {
    <#
    .SYNOPSIS
        Sanitizes bounded text for the management-agent output contract.
    .DESCRIPTION
        Removes control characters and field delimiters, collapses whitespace,
        then truncates with an ellipsis. Does not alter raw report evidence.
    .PARAMETER Value
        Text or scalar value; null becomes an empty string.
    .PARAMETER MaximumLength
        Maximum output characters including ellipsis; must be at least three.
    .OUTPUTS
        System.String containing no line breaks or pipe delimiters.
    #>
    param([AllowNull()][object]$Value, [int]$MaximumLength = 600)
    $text = ([string]$Value -replace '[\x00-\x1F\x7F|]', ' ' -replace '\s+', ' ').Trim()
    if ($text.Length -gt $MaximumLength) { return $text.Substring(0, $MaximumLength - 3) + '...' }
    return $text
}

function Write-NtlmIvantiResult {
    <#
    .SYNOPSIS
        Emits the existing four-line Ivanti detection contract.
    .DESCRIPTION
        Writes detected, reason, expected and found with Write-Host, matching
        the other Ivanti detectors. Expected describes zero matching records;
        found also includes the informational collection-diagnostics count.
        Detailed counters, error messages and event samples are not printed here.
    .PARAMETER Report
        Completed or partially collected report with a Result assessment.
    .OUTPUTS
        Four host/information-stream lines; no success-stream objects.
    #>
    param([object]$Report)
    $assessment = $Report.Result
    $ntlmCount = 'Unknown'
    $derivedCount = 'Unknown'
    if ($null -ne $assessment.NtlmRecords) { $ntlmCount = [string]$assessment.NtlmRecords }
    if ($null -ne $assessment.DerivedCredentialRecords) { $derivedCount = [string]$assessment.DerivedCredentialRecords }
    $detectedString = $assessment.Detected.ToString().ToLowerInvariant()
    $reasonString = ConvertTo-NtlmSingleLine $assessment.Reason
    $expectedString = 'NTLM records: 0 | NTLMv1-derived SSO records: 0'
    $foundString = "NTLM records: $ntlmCount | NTLMv1-derived SSO records: $derivedCount | Collection diagnostics: $($assessment.Gaps.Count)"

    Write-Host "detected = $detectedString"
    Write-Host "reason = $reasonString"
    Write-Host "expected = $expectedString"
    Write-Host "found = $foundString"
}

function Write-NtlmConsoleTable {
    <#
    .SYNOPSIS
        Renders a labeled table on the host without leaking formatting objects.
    .DESCRIPTION
        Sanitizes data cells, wraps long values and uses a bounded layout width.
        Empty input prints an explicit none marker. Raw JSON data is unchanged.
    .PARAMETER Title
        Section heading supplied by the report renderer.
    .PARAMETER Rows
        Objects whose selected properties form the table columns.
    .OUTPUTS
        Host/information-stream text only; no success-stream formatting records.
    #>
    param([string]$Title, [object[]]$Rows)
    Write-Host ("`n" + $Title) -ForegroundColor Cyan
    if (@($Rows).Count -eq 0) { Write-Host '  (none)'; return }
    $displayRows = foreach ($row in $Rows) {
        $cells = [ordered]@{}
        foreach ($property in $row.PSObject.Properties) {
            if ($null -eq $property.Value) { $cells[$property.Name] = '(unknown)' }
            else { $cells[$property.Name] = ConvertTo-NtlmSingleLine ($property.Value -join ', ') 400 }
        }
        [pscustomobject]$cells
    }
    Write-Host (($displayRows | Format-Table -AutoSize -Wrap | Out-String -Width 140).TrimEnd())
}

function Write-NtlmDetailedResult {
    <#
    .SYNOPSIS
        Displays the compact opt-in command-line NTLM report.
    .DESCRIPTION
        Shows the result, key counts, first collection diagnostic and at most three
        recent samples. Omits policy/schema/retention tables, top lists and audit
        markers; these remain in JSON. Logging and JSON-write errors are displayed
        briefly. Rendering does not change evidence or detection semantics.
    .PARAMETER Report
        Same report consumed by the Ivanti renderer and optional JSON export.
    .OUTPUTS
        Host/information-stream text only. Never mutates the report or verdict.
    #>
    param([object]$Report)
    Write-Host "`nNTLM Usage Detection" -ForegroundColor Cyan
    $color = 'Yellow'
    if ($Report.Result.Assessment -eq 'NoEvidenceObserved') { $color = 'Green' }
    Write-Host (ConvertTo-NtlmSingleLine ("Computer: $($Report.Machine.Computer); Lookback: $($Report.Days) days; Detected: $($Report.Result.Detected); Status: $($Report.Result.Assessment)") 200) -ForegroundColor $color
    Write-Host (ConvertTo-NtlmSingleLine $Report.Result.Reason 240)
    $metrics = @('NtlmRecords', 'NtlmV1Records', 'DerivedCredentialRecords')
    Write-NtlmConsoleTable -Title 'Evidence Counts' -Rows @($metrics | ForEach-Object { [pscustomobject]@{ Metric = $_; Records = $Report.Result.$_ } })
    Write-Host ("Collection diagnostics: $($Report.Result.Gaps.Count)")
    if ($Report.Result.Gaps.Count -gt 0) {
        Write-Host ('  First: ' + (ConvertTo-NtlmSingleLine $Report.Result.Gaps[0] 200)) -ForegroundColor Yellow
    }
    $samples = @($Report.Evidence | Select-Object -First 3)
    if ($samples.Count -gt 0) {
        Write-Host ("`nRecent Evidence: $($samples.Count) of $($Report.Evidence.Count) retained samples") -ForegroundColor Cyan
    }
    elseif ($Report.Result.NtlmRecords -gt 0 -or $Report.Result.DerivedCredentialRecords -gt 0) {
        Write-Host 'No samples retained.'
    }
    foreach ($sample in $samples) {
        Write-Host (ConvertTo-NtlmSingleLine ("$($sample.TimeUtc) Event=$($sample.EventId) Record=$($sample.RecordId) $($sample.Direction) $($sample.Outcome) $($sample.NtlmVersion)") 200)
        Write-Host ('  Process: ' + (ConvertTo-NtlmSingleLine $sample.Process 90) + '; Target: ' + (ConvertTo-NtlmSingleLine $sample.Target 90))
        if ($sample.UsageId -or $sample.UsageReason) {
            Write-Host ('  Usage: ' + (ConvertTo-NtlmSingleLine ("$($sample.UsageId) $($sample.UsageReason)") 200))
        }
    }
    if ($Report.Logging.Error) { Write-Host ('Log write failed: ' + (ConvertTo-NtlmSingleLine $Report.Logging.Error 200)) -ForegroundColor Yellow }
    elseif ($Report.Logging.Enabled) { Write-Host ('Log: ' + (ConvertTo-NtlmSingleLine $Report.Logging.Path 200)) }
    if ($Report.ReportFileError) { Write-Host ('JSON write failed: ' + (ConvertTo-NtlmSingleLine $Report.ReportFileError 200)) -ForegroundColor Yellow }
    elseif ($Report.JsonReportPath) { Write-Host ('JSON: ' + (ConvertTo-NtlmSingleLine $Report.JsonReportPath 200)) }
    else { Write-Host 'Full diagnostics: use -ReportPath <file.json>.' }
}

$ErrorActionPreference = 'Stop'
$timer = [Diagnostics.Stopwatch]::StartNew()
$script:NtlmLogState = [pscustomobject]@{ Enabled = -not [string]::IsNullOrWhiteSpace($LogPath); Path = $LogPath; RunId = [guid]::NewGuid().ToString(); EntriesWritten = 0; Error = $null }
if ($script:NtlmLogState.Enabled) {
    try {
        $script:NtlmLogState.Path = [IO.Path]::GetFullPath($LogPath)
        if ($ReportPath -and [IO.Path]::GetFullPath($ReportPath) -eq $script:NtlmLogState.Path) { throw 'LogPath and ReportPath must name different files.' }
    }
    catch { $script:NtlmLogState.Error = $_.Exception.Message }
}
$report = [pscustomobject][ordered]@{
    SchemaVersion = '1.2'; GeneratedUtc = [datetime]::UtcNow.ToString('o'); Days = $Days
    Logging = $script:NtlmLogState; DurationMilliseconds = $null; JsonReportPath = $ReportPath
    StartUtc = $null; EndUtc = $null; Machine = $null; Result = $null
    Logs = @(); Sources = @(); Policies = @(); AuditPolicy = $null; Providers = $null
    TopProcesses = @(); TopTargets = @(); TopAccounts = @(); TopClients = @(); TopUsageReasons = @()
    ByEventId = @(); ByDirection = @(); ByVersion = @(); ByLogonType = @()
    Evidence = @(); EvidenceOmitted = 0; AuditMarkers = @(); ReportFileError = $null
    Limitations = @(
        'Collection scope: local NTLM Operational, Security and System channels.'
        'Queries are limited by Days and MaxEventsPerSource.'
        'Retained evidence samples are limited by MaxEvidence.'
        'Authentication, auditing and registry settings are not changed.'
    )
}
Write-NtlmLog ("Run start: Computer=$env:COMPUTERNAME; PowerShell=$($PSVersionTable.PSVersion); Days=$Days; MaxEventsPerSource=$MaxEventsPerSource; MaxEvidence=$MaxEvidence; Detailed=$Detailed")
try {
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT -or -not [Environment]::Is64BitProcess) {
        throw 'Run in 64-bit Windows PowerShell 5.1 or later on Windows.'
    }
    $endUtc = [datetime]::UtcNow
    $startUtc = $endUtc.AddDays(-$Days)
    $report.StartUtc = $startUtc.ToString('o')
    $report.EndUtc = $endUtc.ToString('o')
    $report.Machine = Get-NtlmMachineState
    Write-NtlmLog ('Execution context: ' + ($report.Machine | ConvertTo-Json -Compress))
    $report.Policies = @(Get-NtlmPolicySnapshot)
    Write-NtlmLog ('Policy snapshot: ' + (ConvertTo-Json -InputObject $report.Policies -Compress -Depth 4))
    $report.AuditPolicy = Get-NtlmAuditPolicy
    Write-NtlmLog ('Audit policy: ' + ($report.AuditPolicy | ConvertTo-Json -Compress))
    $report.Providers = Get-NtlmProviderState
    Write-NtlmLog ('Provider schemas: ' + ($report.Providers | ConvertTo-Json -Compress -Depth 4))
    $report.Logs = @(foreach ($logName in @('Microsoft-Windows-NTLM/Operational', 'Security', 'System')) {
        Get-NtlmLogState -Name $logName -StartUtc $startUtc
    })
    Write-NtlmLog ('Log metadata: ' + (ConvertTo-Json -InputObject $report.Logs -Compress -Depth 4))
    $collection = Get-NtlmCollection -StartUtc $startUtc -EndUtc $endUtc -Limit $MaxEventsPerSource
    $report.Sources = $collection.Sources
    $report.AuditMarkers = $collection.AuditMarkers
    $report.Result = Get-NtlmAssessment -Collection $collection -Logs $report.Logs -Policies $report.Policies -AuditPolicy $report.AuditPolicy -Providers $report.Providers -Machine $report.Machine
    $report.TopProcesses = @(Get-NtlmTopValues $collection.Events 'Process')
    $report.TopTargets = @(Get-NtlmTopValues $collection.Events 'Target')
    $report.TopAccounts = @(Get-NtlmTopValues $collection.Events 'Account')
    $report.TopClients = @(Get-NtlmTopValues $collection.Events 'Client')
    $report.TopUsageReasons = @(Get-NtlmTopValues $collection.Events 'UsageId')
    $report.ByEventId = @(Get-NtlmTopValues $collection.Events 'EventId' 32)
    $report.ByDirection = @(Get-NtlmTopValues $collection.Events 'Direction')
    $report.ByVersion = @(Get-NtlmTopValues $collection.Events 'NtlmVersion')
    $report.ByLogonType = @(Get-NtlmTopValues $collection.Events 'LogonType')
    $report.Evidence = @($collection.Events | Sort-Object TimeUtc, RecordId -Descending | Select-Object -First $MaxEvidence)
    $report.EvidenceOmitted = $collection.Events.Count - $report.Evidence.Count
}
catch {
    Write-NtlmLog -Level ERROR -Message ('Collection/report error: ' + $_.Exception.Message)
    if ($null -ne $report.Result) {
        $report.Result.Gaps = @($report.Result.Gaps) + @('Report assembly failed: ' + $_.Exception.Message)
        $report.Result.Reason += ' Report assembly is incomplete.'
        if ($report.Result.Assessment -eq 'NoEvidenceObserved') { $report.Result.Assessment = 'CollectionError' }
    }
    else {
        $report.Result = [pscustomobject]@{
            Assessment = 'CollectionError'; Detected = $false; Reason = 'Collection failed. Error details are listed in the collection diagnostics.'
            Gaps = @($_.Exception.Message); NtlmRecords = $null; DerivedCredentialRecords = $null; NtlmV1Records = $null; LmRecords = $null
            EnhancedWarnings = $null; SuccessfulLogonRecords = $null; SuccessfulValidationRecords = $null; EnhancedSuccessRecords = $null
            BlockedRecords = $null; AttemptRecords = $null; FailureOrNonSuccessRecords = $null
        }
    }
}
$timer.Stop()
$report.DurationMilliseconds = $timer.ElapsedMilliseconds
foreach ($gap in $report.Result.Gaps) { Write-NtlmLog -Level WARN -Message ('Coverage gap: ' + $gap) }
Write-NtlmLog ('Assessment: ' + ($report.Result | ConvertTo-Json -Compress -Depth 4))
Write-NtlmLog ("Collection complete: DurationMs=$($report.DurationMilliseconds); Samples=$($report.Evidence.Count); SamplesOmitted=$($report.EvidenceOmitted)")
if ($ReportPath) {
    try {
        $fullPath = [IO.Path]::GetFullPath($ReportPath)
        if ($LogPath -and $fullPath -eq [IO.Path]::GetFullPath($LogPath)) { throw 'LogPath and ReportPath must name different files.' }
        $report.JsonReportPath = $fullPath
        Write-NtlmLog ('Writing JSON report: ' + $fullPath)
        [IO.File]::WriteAllText($fullPath, ($report | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding($true)))
        Write-NtlmLog 'JSON report written successfully.'
    }
    catch {
        $report.ReportFileError = $_.Exception.Message
        $report.Result.Gaps = @($report.Result.Gaps) + @('Requested report could not be written: ' + $_.Exception.Message)
        if ($report.Result.Assessment -eq 'NoEvidenceObserved') {
            $report.Result.Assessment = 'InsufficientEvidence'
            $report.Result.Reason = 'No matching NTLM records found. JSON report write failed.'
        }
        else { $report.Result.Reason += ' JSON report write failed.' }
        Write-NtlmLog -Level ERROR -Message ('JSON report failed: ' + $_.Exception.Message)
    }
}
    Write-NtlmLog ("Run end: Assessment=$($report.Result.Assessment); Detected=$($report.Result.Detected)")
    if ($Detailed) { Write-NtlmDetailedResult -Report $report }
    else { Write-NtlmIvantiResult -Report $report }