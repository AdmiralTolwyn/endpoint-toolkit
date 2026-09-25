#Requires -Version 5.1

<#
.SYNOPSIS
    Runs offline regression tests for the Ivanti NTLM detector.
.DESCRIPTION
    Loads production functions and the entry point from their AST, then replaces
    Windows reads with deterministic fixtures. Uses only uniquely named temporary
    log/JSON files for export tests and removes those files in finally blocks.
    Does not read real authentication records or change audit/registry policy.
.EXAMPLE
    .\Test-NtlmUsageDetection.ps1
.OUTPUTS
    One assertion-count summary on success; a terminating exception on failure.
.NOTES
    Run in 64-bit Windows PowerShell 5.1 or PowerShell 7 on Windows. Tests include
    comment-based help, output-stream isolation, logging failures and rich output.
#>
param()

$ErrorActionPreference = 'Stop'
$sourcePath = Join-Path $PSScriptRoot 'Get-NtlmUsageDetection-Ivanti.ps1'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($sourcePath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ($parseErrors | Out-String) }
$bytes = [IO.File]::ReadAllBytes($sourcePath)
if ([BitConverter]::ToString($bytes, 0, 4) -ne 'EF-BB-BF-23') { throw 'Missing or duplicated UTF-8 BOM.' }
foreach ($definition in $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)) {
    . ([scriptblock]::Create($definition.Extent.Text))
}
$entryText = ($ast.ParamBlock.Attributes.Extent.Text -join [Environment]::NewLine) + [Environment]::NewLine + $ast.ParamBlock.Extent.Text + [Environment]::NewLine
$entryText += ($ast.EndBlock.Statements | Where-Object { $_ -isnot [System.Management.Automation.Language.FunctionDefinitionAst] } |
    ForEach-Object { $_.Extent.Text }) -join [Environment]::NewLine
$entryPoint = [scriptblock]::Create($entryText + [Environment]::NewLine + '$report')
$script:assertions = 0

function Assert-Ntlm {
    <#
    .SYNOPSIS
        Counts an assertion and stops the suite when its condition is false.
    .DESCRIPTION
        Maintains the script-scoped assertion counter used by the final summary.
    .PARAMETER Condition
        Expected true condition.
    .PARAMETER Message
        Diagnostic label identifying the invariant being tested.
    .OUTPUTS
        None. Throws on failure.
    #>
    param([bool]$Condition, [string]$Message)
    $script:assertions++
    if (-not $Condition) { throw "Assertion failed: $Message" }
}

function New-TestEvent {
    <#
    .SYNOPSIS
        Creates a disposable synthetic event record with XML-backed data.
    .DESCRIPTION
        Uses XML APIs to escape field values. ToXml and Dispose mimic the small
        EventRecord surface required by production collection code.
    .PARAMETER Id
        Synthetic event ID.
    .PARAMETER Data
        Named EventData values.
    .PARAMETER Provider
        Provider identity; defaults to Microsoft-Windows-NTLM.
    .PARAMETER Channel
        Local channel identity used by the query mock.
    .PARAMETER Time
        Event timestamp used for lookback and retention tests.
    .PARAMETER RecordId
        Synthetic record number for stable ordering and references.
    .OUTPUTS
        PSCustomObject with XML, record metadata and mocked lifecycle methods.
    #>
    param([int]$Id, [hashtable]$Data = @{}, [string]$Provider = 'Microsoft-Windows-NTLM',
        [string]$Channel = 'Microsoft-Windows-NTLM/Operational', [datetime]$Time = ([datetime]::UtcNow.AddHours(-1)), [long]$RecordId = 20)
    $document = New-Object System.Xml.XmlDocument
    $document.LoadXml('<Event><System/><EventData/></Event>')
    $system = $document.DocumentElement.SelectSingleNode('System')
    $providerNode = $document.CreateElement('Provider')
    $providerNode.SetAttribute('Name', $Provider)
    [void]$system.AppendChild($providerNode)
    foreach ($pair in @(@('EventID', $Id), @('Channel', $Channel), @('EventRecordID', $RecordId), @('Computer', 'TEST-PC'))) {
        $node = $document.CreateElement($pair[0])
        $node.InnerText = [string]$pair[1]
        [void]$system.AppendChild($node)
    }
    $timeNode = $document.CreateElement('TimeCreated')
    $timeNode.SetAttribute('SystemTime', $Time.ToUniversalTime().ToString('o'))
    [void]$system.AppendChild($timeNode)
    foreach ($name in $Data.Keys) {
        $node = $document.CreateElement('Data')
        $node.SetAttribute('Name', $name)
        $node.InnerText = [string]$Data[$name]
        [void]$document.DocumentElement.SelectSingleNode('EventData').AppendChild($node)
    }
    $record = [pscustomobject]@{ Id = $Id; RecordId = $RecordId; TimeCreated = $Time; LogName = $Channel; Xml = $document.OuterXml; Disposed = $false }
    Add-Member -InputObject $record -MemberType ScriptMethod -Name ToXml -Value { $this.Xml }
    Add-Member -InputObject $record -MemberType ScriptMethod -Name Dispose -Value { $this.Disposed = $true }
    return $record
}

function Reset-NtlmFixture {
    <#
    .SYNOPSIS
        Resets Windows-read fixtures to an adequate-auditing baseline.
    .DESCRIPTION
        Seeds old, unrelated channel records to establish retained history without
        adding NTLM evidence. Individual tests then change only their own inputs.
    .OUTPUTS
        None. Replaces script-scoped fixture state.
    #>
    $script:fixture = @{
        Records = @(); DeniedLog = ''; DisabledLog = ''; ThrowAfterRecord = $false; CorruptRecord = $false
        PolicyValues = @{ LogEnhancedAuditEvents = 1; EnableEnhancedDomainNtlmLogs = 1 }
        AuditPolicy = [pscustomobject]@{ State = 'Read'; Logon = 3; CredentialValidation = 3; Error = $null }
        ClientIds = @(4020, 4021, 4022, 4023); DomainIds = @(4030, 4031, 4032, 4033)
        ProductType = 1; RecentRetention = $false
    }
    foreach ($channel in @('Microsoft-Windows-NTLM/Operational', 'Security', 'System')) {
        $script:fixture.Records += New-TestEvent -Id 1 -Channel $channel -Time ([datetime]::UtcNow.AddDays(-60))
    }
}

function Get-WinEvent {
    <#
    .SYNOPSIS
        Mocks local event queries, provider schemas and log metadata.
    .DESCRIPTION
        Applies structural XPath with .NET and time bounds separately, since
        Windows Event Log timestamp comparisons are not standard XPath semantics.
        Supports injected access failures, truncation, partial reads and bad XML.
    .PARAMETER LogName
        Fixture channel to query.
    .PARAMETER FilterXPath
        Production Windows Event Log predicate.
    .PARAMETER MaxEvents
        Maximum synthetic records to emit.
    .PARAMETER Oldest
        Sort oldest first for retention probes.
    .PARAMETER ListLog
        Return mock channel configuration instead of events.
    .PARAMETER ListProvider
        Return mock installed event IDs instead of events.
    .OUTPUTS
        Metadata objects or synthetic event records; throws configured errors.
    #>
    [CmdletBinding()]
    param([string]$LogName, [string]$FilterXPath, [int]$MaxEvents, [switch]$Oldest, [string]$ListLog, [string]$ListProvider)
    if ($ListProvider) {
        $ids = $script:fixture.ClientIds
        if ($ListProvider -eq 'Microsoft-Windows-Security-Netlogon') { $ids = $script:fixture.DomainIds }
        return [pscustomobject]@{ Events = @($ids | ForEach-Object { [pscustomobject]@{ Id = $_ } }) }
    }
    $requestedLog = $LogName
    if ($ListLog) { $requestedLog = $ListLog }
    if ($script:fixture.DeniedLog -eq $requestedLog) { throw 'Access denied by fixture.' }
    if ($ListLog) {
        return [pscustomobject]@{ IsEnabled = ($script:fixture.DisabledLog -ne $ListLog); RecordCount = 10; MaximumSizeInBytes = 1048576; LogMode = 'Circular' }
    }
    $timePredicate = [regex]::Match($FilterXPath, "TimeCreated\[@SystemTime >= '([^']+)' and @SystemTime <= '([^']+)'\]")
    $structuralFilter = $FilterXPath -replace 'TimeCreated\[[^\]]+\]', 'true()'
    $selected = @($script:fixture.Records | Where-Object {
        if ($_.LogName -ne $LogName) { return $false }
        if ($timePredicate.Success -and ($_.TimeCreated -lt [datetime]$timePredicate.Groups[1].Value -or $_.TimeCreated -gt [datetime]$timePredicate.Groups[2].Value)) { return $false }
        $document = [xml]$_.Xml
        return $document.SelectNodes($structuralFilter).Count -gt 0
    } | Sort-Object TimeCreated -Descending:(-not $Oldest) | Select-Object -First $MaxEvents)
    if ($selected.Count -eq 0) {
        $exception = New-Object System.Exception 'No matching records.'
        $errorRecord = New-Object System.Management.Automation.ErrorRecord $exception, 'NoMatchingEventsFound', ([System.Management.Automation.ErrorCategory]::ObjectNotFound), $null
        $PSCmdlet.ThrowTerminatingError($errorRecord)
    }
    foreach ($record in $selected) {
        if ($script:fixture.RecentRetention -and $Oldest) { $record.TimeCreated = [datetime]::UtcNow.AddHours(-2) }
        if ($script:fixture.CorruptRecord -and -not $Oldest -and $record.Id -eq 4020) {
            $record.Xml = '<broken'
        }
        $record
        if ($script:fixture.ThrowAfterRecord -and -not $Oldest) { throw 'Read interrupted after partial results.' }
    }
}

function Get-NtlmRegistryValue {
    <#
    .SYNOPSIS
        Returns a fixture policy value without touching the registry.
    .DESCRIPTION
        Mirrors configured versus absent value states for policy snapshot tests.
    .PARAMETER Path
        Requested HKLM-relative path, retained for assertions.
    .PARAMETER Name
        Value name looked up in the fixture dictionary.
    .OUTPUTS
        PSCustomObject matching the production registry-reader contract.
    #>
    param([string]$Path, [string]$Name)
    $state = 'NotConfigured'
    $value = $null
    if ($script:fixture.PolicyValues.ContainsKey($Name)) { $state = 'Configured'; $value = $script:fixture.PolicyValues[$Name] }
    [pscustomobject]@{ Path = $Path; Name = $Name; Value = $value; State = $state; Error = $null }
}

function Get-NtlmAuditPolicy {
    <#
    .SYNOPSIS
        Supplies the fixture's effective audit policy.
    .DESCRIPTION
        Prevents native auditpol execution in entry-point tests.
    .OUTPUTS
        Fixture audit-policy PSCustomObject.
    #>
    $script:fixture.AuditPolicy
}

function Get-NtlmMachineState {
    <#
    .SYNOPSIS
        Supplies synthetic machine-role and execution metadata.
    .DESCRIPTION
        Avoids CIM calls and real identity inspection in entry-point tests.
    .OUTPUTS
        PSCustomObject with fixture computer, elevation, role and read status.
    #>
    [pscustomobject]@{ Computer = 'TEST-PC'; Elevated = $true; ProductType = $script:fixture.ProductType; Error = $null }
}

function Invoke-TestDetection {
    <#
    .SYNOPSIS
        Runs the production entry point and verifies default Ivanti output.
    .DESCRIPTION
        Disables file logging unless a test explicitly supplies LogPath. Captures
        four host lines and the test-only report; detects output contamination.
    .PARAMETER Arguments
        Script parameters to splat into the extracted production entry point.
    .OUTPUTS
        The completed report object after contract assertions pass.
    #>
    param([hashtable]$Arguments = @{})
    if (-not $Arguments.ContainsKey('LogPath')) { $Arguments['LogPath'] = '' }
    $results = @(& $entryPoint @Arguments 6>&1)
    Assert-Ntlm ($results.Count -eq 5) 'Four information-stream lines plus one test-only report.'
    $lines = @($results[0..3] | ForEach-Object { [string]$_ })
    foreach ($index in 0..3) {
        $key = @('detected', 'reason', 'expected', 'found')[$index]
        Assert-Ntlm ($lines[$index] -match ('^' + $key + ' = ') -and $lines[$index] -notmatch '[\r\n]') "Single-line key $key."
    }
    $result = $results[4].Result
    $expectedLines = @(
        'detected = ' + $result.Detected.ToString().ToLowerInvariant()
        'reason = ' + (ConvertTo-NtlmSingleLine $result.Reason)
        'expected = NTLM records: 0 | NTLMv1-derived SSO records: 0'
        "found = NTLM records: $($result.NtlmRecords) | NTLMv1-derived SSO records: $($result.DerivedCredentialRecords) | Collection diagnostics: $($result.Gaps.Count)"
    )
    Assert-Ntlm (($lines -join "`n") -ceq ($expectedLines -join "`n")) 'Default output exactly matches the concise Ivanti contract.'
    Assert-Ntlm ($result.Detected -eq ($result.NtlmRecords -gt 0 -or $result.DerivedCredentialRecords -gt 0)) 'Only matching evidence sets detected, never collection or file errors.'
    Assert-Ntlm ($results[4].Result.Assessment -ne 'CollectionError') ('No unexpected collection error: ' + ($results[4].Result.Gaps -join '; '))
    return $results[4]
}

foreach ($eventId in @(4020, 4021)) {
    $record = New-TestEvent $eventId @{ NtlmVersion = 'NTLMv2'; NtlmUsageId = '7'; TargetService = 'cifs/192.0.2.10' }
    $event = ConvertFrom-NtlmEvent $record.ToXml()
    Assert-Ntlm ($event.Outcome -eq 'Attempt' -and $event.NtlmVersion -eq 'NTLMv2') 'Client event is attempt, warning is not v1.'
    Assert-Ntlm ($event.SecurityWarning -eq ($eventId -eq 4021)) 'Warning based on documented paired ID.'
}
foreach ($status in @('0', '0x00000000', '0xc000006d', '0x00000103', 'unknown', '')) {
    $event = ConvertFrom-NtlmEvent (New-TestEvent 4023 @{ Status = $status; NtlmVersion = 'NTLMv1'; 'Mic Status' = 'Unprotected' }).ToXml()
    $expected = 'Unknown'
    if ($status -in @('0', '0x00000000')) { $expected = 'Success' }
    elseif ($status -like '0x*') { $expected = 'NonSuccessStatus' }
    Assert-Ntlm ($event.Outcome -eq $expected -and $event.MicStatus -eq 'Unprotected') 'Server status and exact spaced field.'
}
foreach ($eventId in @(4030, 4031, 4032, 4033)) {
    $event = ConvertFrom-NtlmEvent (New-TestEvent $eventId @{ AccountName = 'example'; AccountDomain = 'TEST'; AccountMachine = 'CLIENT'; ServerName = 'SERVER'; Status = '0x0' } 'Microsoft-Windows-Security-Netlogon').ToXml()
    Assert-Ntlm ($event.Direction -eq 'DomainValidation' -and $event.Account -eq 'example' -and $event.Client -eq 'CLIENT' -and $event.Target -eq 'SERVER') 'Verified DC schema maps who and where.'
}
foreach ($eventId in @(4001, 4002, 4003, 8001, 8002, 8003, 4024, 4025, 4026, 4027)) {
    $event = ConvertFrom-NtlmEvent (New-TestEvent $eventId).ToXml()
    $expected = 'Attempt'
    if ($eventId -in @(4001, 4002, 4003, 4025, 4027)) { $expected = 'Blocked' }
    Assert-Ntlm ($event.Outcome -eq $expected) "Legacy/derived/block outcome $eventId."
}
foreach ($eventId in @(4004, 8004)) {
    $event = ConvertFrom-NtlmEvent (New-TestEvent $eventId @{} 'Microsoft-Windows-Security-Netlogon').ToXml()
    Assert-Ntlm ($event.Direction -eq 'DomainValidation') 'Legacy DC provider accepted.'
}
$securityProvider = 'Microsoft-Windows-Security-Auditing'
foreach ($package in @('NTLM', 'Kerberos', 'Negotiate')) {
    $event = ConvertFrom-NtlmEvent (New-TestEvent 4624 @{ AuthenticationPackageName = $package; LmPackageName = 'NTLM V1'; TargetUserSid = 'S-1-5-21-1-2-3-1000' } $securityProvider 'Security').ToXml()
    Assert-Ntlm (($null -ne $event) -eq ($package -eq 'NTLM')) 'Do not treat Negotiate or Kerberos as proof of NTLM.'
}
$anonymous = ConvertFrom-NtlmEvent (New-TestEvent 4624 @{ AuthenticationPackageName = 'NTLM'; LmPackageName = 'NTLM V1'; TargetUserSid = 'S-1-5-7' } $securityProvider 'Security').ToXml()
Assert-Ntlm ($anonymous.NtlmVersion -eq 'Unknown' -and $anonymous.Data['LmPackageName'] -eq 'NTLM V1') 'Anonymous version label preserved but not counted as v1.'
$validation = ConvertFrom-NtlmEvent (New-TestEvent 4776 @{ PackageName = 'MICROSOFT_AUTHENTICATION_PACKAGE_V1_0'; Status = '0x0' } $securityProvider 'Security').ToXml()
Assert-Ntlm ($validation.NtlmVersion -eq 'Unknown' -and $validation.Outcome -eq 'Success' -and -not $validation.Target) '4776 V1_0 is not NTLMv1 and has no destination.'
$wrongProvider = ConvertFrom-NtlmEvent (New-TestEvent 4020 @{} 'Other-Provider').ToXml()
Assert-Ntlm ($null -eq $wrongProvider) 'Event ID alone does not establish evidence.'
$namespaced = (New-TestEvent 4020 @{ ProcessName = 'app.exe' }).ToXml().Replace('<Event>', '<Event xmlns="http://schemas.microsoft.com/win/2004/08/events/event">')
Assert-Ntlm ((ConvertFrom-NtlmEvent $namespaced).Process -eq 'app.exe') 'Default XML namespace works.'

Reset-NtlmFixture
$report = Invoke-TestDetection
Assert-Ntlm ($report.Result.Assessment -eq 'NoEvidenceObserved' -and -not $report.Result.Detected) 'No evidence with current prerequisites is not a positive finding.'
Assert-Ntlm ($report.Result.Reason -eq 'No matching NTLM records found.') 'Zero-match summary reports the observed result only.'
Assert-Ntlm ($report.Policies.Count -eq 9 -and $report.Policies[0].Name -eq 'LogEnhancedAuditEvents') 'Registry path/name pairs must not flatten.'
foreach ($source in $report.Sources) { Assert-Ntlm (-not $source.Error -and -not $source.Truncated) 'No-match query is not a read failure.' }

Reset-NtlmFixture
$script:fixture.RecentRetention = $true
$script:fixture.PolicyValues.LogEnhancedAuditEvents = 0
$script:fixture.AuditPolicy.Logon = 0
$report = Invoke-TestDetection
Assert-Ntlm ($report.Result.Gaps.Count -eq 5 -and $report.Result.NtlmRecords -eq 0 -and $report.Result.DerivedCredentialRecords -eq 0 -and -not $report.Result.Detected) 'Regression: five collection diagnostics with zero matching records must not set detected.'

Reset-NtlmFixture
$script:fixture.DeniedLog = 'Security'
$report = Invoke-TestDetection
Assert-Ntlm ($report.Result.Assessment -eq 'InsufficientEvidence' -and -not $report.Result.Detected) 'Unreadable Security log is a collection finding, not detected NTLM.'
Assert-Ntlm ($report.Result.Reason -eq "No matching NTLM records found. Collection diagnostics: $($report.Result.Gaps.Count).") 'Collection-diagnostic summary reports the count without a reliability verdict.'
$script:fixture.Records += New-TestEvent 4020 @{ ProcessName = "app`r`nfound = forged"; NtlmUsageId = 1 }
$report = Invoke-TestDetection
Assert-Ntlm ($report.Result.Assessment -eq 'NtlmEvidenceObserved' -and $report.Result.Gaps.Count -gt 0) 'Positive evidence survives another source access failure.'
Assert-Ntlm ($report.Result.Reason -eq "Matching NTLM records: $($report.Result.NtlmRecords).") 'Positive summary reports matching record count.'

foreach ($gapMode in @('Disabled', 'Retention', 'Policy', 'AuditPolicy', 'Schemas', 'Domain')) {
    Reset-NtlmFixture
    switch ($gapMode) {
        'Disabled' { $script:fixture.DisabledLog = 'Microsoft-Windows-NTLM/Operational' }
        'Retention' { $script:fixture.RecentRetention = $true }
        'Policy' { $script:fixture.PolicyValues.LogEnhancedAuditEvents = 0 }
        'AuditPolicy' { $script:fixture.AuditPolicy = [pscustomobject]@{ State = 'Unknown'; Logon = $null; CredentialValidation = $null; Error = 'Denied' } }
        'Schemas' { $script:fixture.ClientIds = @() }
        'Domain' { $script:fixture.ProductType = 2; $script:fixture.PolicyValues.EnableEnhancedDomainNtlmLogs = 0 }
    }
    $report = Invoke-TestDetection
    Assert-Ntlm ($report.Result.Assessment -eq 'InsufficientEvidence') "Coverage gap: $gapMode."
}

Reset-NtlmFixture
$script:fixture.PolicyValues = @{ RestrictSendingNTLMTraffic = 1; AuditReceivingNTLMTraffic = 2 }
$script:fixture.ClientIds = @()
$report = Invoke-TestDetection
Assert-Ntlm ($report.Result.Assessment -eq 'NoEvidenceObserved') 'Legacy Audit all can establish current client/server readiness.'

Reset-NtlmFixture
$script:fixture.Records += New-TestEvent 4025
$report = Invoke-TestDetection
Assert-Ntlm ($report.Result.Assessment -eq 'DerivedCredentialEvidenceOnly' -and $report.Result.NtlmRecords -eq 0 -and $report.Result.NtlmV1Records -eq 0) 'Derived credentials are separate from NTLM network evidence.'

Reset-NtlmFixture
foreach ($recordId in 1..3) { $script:fixture.Records += New-TestEvent 4020 @{ ProcessName = 'sample.exe'; NtlmVersion = 'NTLMv2' } -RecordId $recordId }
$report = Invoke-TestDetection @{ MaxEventsPerSource = 2; MaxEvidence = 1 }
Assert-Ntlm ($report.Result.NtlmRecords -eq 2 -and $report.Evidence.Count -eq 1 -and $report.EvidenceOmitted -eq 1) 'Query and sample caps are independent.'
Assert-Ntlm ($report.Sources[0].Truncated -and $report.TopProcesses[0].Records -eq 2) 'Aggregates include all processed records, not only the sample.'
$report = Invoke-TestDetection @{ MaxEventsPerSource = 3; MaxEvidence = 0 }
Assert-Ntlm (-not $report.Sources[0].Truncated -and $report.Evidence.Count -eq 0) 'Exactly at cap is not falsely truncated; zero samples supported.'
$script:fixture.ThrowAfterRecord = $true
$report = Invoke-TestDetection
Assert-Ntlm ($report.Result.NtlmRecords -eq 1 -and $report.Result.Gaps.Count -gt 0) 'Partial query errors retain positive evidence.'

Reset-NtlmFixture
$script:fixture.Records += New-TestEvent 4020
$script:fixture.CorruptRecord = $true
$report = Invoke-TestDetection
Assert-Ntlm ($report.Result.Assessment -eq 'InsufficientEvidence' -and $report.Sources[0].ParseErrors -eq 1) 'Malformed XML is a coverage gap.'

Reset-NtlmFixture
$script:fixture.Records += New-TestEvent 4719 @{} $securityProvider 'Security'
$report = Invoke-TestDetection
Assert-Ntlm ($report.AuditMarkers.Count -eq 1 -and $report.Result.Assessment -eq 'InsufficientEvidence') 'Audit change marker prevents an unqualified negative.'

Reset-NtlmFixture
$badPath = Join-Path $env:TEMP (([guid]::NewGuid().ToString()) + '\not-created\report.json')
$report = Invoke-TestDetection @{ ReportPath = $badPath }
Assert-Ntlm (-not $report.Result.Detected -and $report.ReportFileError) 'Requested report failure is visible without setting detected.'
Assert-Ntlm (-not (Test-Path -LiteralPath (Split-Path -Parent $badPath))) 'No report directory is created implicitly.'
$script:fixture.Records += New-TestEvent 4020
$report = Invoke-TestDetection @{ ReportPath = $badPath }
Assert-Ntlm ($report.Result.Detected -and $report.Result.NtlmRecords -eq 1 -and $report.ReportFileError) 'JSON-write failure preserves matching NTLM evidence.'

Reset-NtlmFixture
$script:fixture.Records += New-TestEvent 4020 @{ NtlmUsageId = '99'; NtlmUsageReason = 'Future reason'; ProcessName = 'test.exe' }
$exportPath = Join-Path $env:TEMP ('NtlmDetectionTest-' + [guid]::NewGuid().ToString() + '.json')
try {
    $report = Invoke-TestDetection @{ ReportPath = $exportPath }
    $export = [IO.File]::ReadAllText($exportPath) | ConvertFrom-Json
    Assert-Ntlm ($export.Result.NtlmRecords -eq 1 -and $export.Evidence[0].Data.NtlmUsageId -eq '99') 'JSON export retains unknown reason IDs and named data.'
    $exportBytes = [IO.File]::ReadAllBytes($exportPath)
    Assert-Ntlm ([BitConverter]::ToString($exportBytes, 0, 3) -eq 'EF-BB-BF') 'JSON is UTF-8 with BOM.'
}
finally { if ([IO.File]::Exists($exportPath)) { [IO.File]::Delete($exportPath) } }

$auditDefinition = $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-NtlmAuditPolicy' }, $false)[0]
$auditCommands = @($auditDefinition.Body.FindAll({ param($node) $node -is [System.Management.Automation.Language.CommandAst] -and $node.Extent.Text -like '*auditpol.exe*' }, $true))
Assert-Ntlm ($auditCommands.Count -eq 1) 'Locate exactly one native auditpol invocation for the parser probe.'
$auditCommand = $auditCommands[0]
$auditProbe = [scriptblock]::Create($auditDefinition.Extent.Text.Replace($auditCommand.Extent.Text, 'Invoke-TestAuditpol') + [Environment]::NewLine + 'Get-NtlmAuditPolicy')
function Invoke-TestAuditpol {
    <#
    .SYNOPSIS
        Supplies native-command output for the production audit CSV parser.
    .DESCRIPTION
        Sets the caller's LASTEXITCODE exactly as an external command would.
    .OUTPUTS
        Fixture CSV or error text lines.
    #>
    Set-Variable -Name LASTEXITCODE -Value $script:auditFixture.ExitCode -Scope 1
    $script:auditFixture.Lines
}
$header = 'Computer,Policy,Subcategory,Guid,Inclusion,Exclusion,Setting'
$logonRow = 'TEST,System,Localized logon,{0cce9215-69ae-11d9-bed3-505054503030},Localized success and failure,,3'
$validationRow = 'TEST,System,Localized validation,{0cce923f-69ae-11d9-bed3-505054503030},Localized success and failure,,3'
foreach ($headerText in @($header, 'Ordinateur,Strategie,SousCategorie,Identifiant,Inclusion,Exclusion,Valeur')) {
    $script:auditFixture = @{ ExitCode = 0; Lines = @($headerText, $logonRow, $validationRow) }
    $policyResult = & $auditProbe
    Assert-Ntlm ($policyResult.State -eq 'Read' -and $policyResult.Logon -eq 3 -and $policyResult.CredentialValidation -eq 3) 'Audit parser uses GUIDs and numeric flags, not localized headers.'
}
foreach ($invalidAudit in @(
    @{ ExitCode = 5; Lines = @('Access denied') },
    @{ ExitCode = 0; Lines = @($header, $logonRow) },
    @{ ExitCode = 0; Lines = @($header, $logonRow, ($validationRow -replace ',3$', ',unexpected')) }
)) {
    $script:auditFixture = $invalidAudit
    $policyResult = & $auditProbe
    Assert-Ntlm ($policyResult.State -eq 'Unknown' -and $policyResult.Error) 'Native failure or ambiguous audit CSV is not silently accepted.'
}

Reset-NtlmFixture
$script:fixture.Records += New-TestEvent 4020
$realAggregation = ${function:Get-NtlmTopValues}
try {
    function Get-NtlmTopValues {
        <#
        .SYNOPSIS
            Injects an aggregation failure after positive evidence exists.
        .DESCRIPTION
            Verifies report-assembly errors cannot erase a positive assessment.
        .OUTPUTS
            None. Always throws the synthetic failure.
        #>
        throw 'Synthetic aggregation failure.'
    }
    $report = Invoke-TestDetection
    Assert-Ntlm ($report.Result.Assessment -eq 'NtlmEvidenceObserved' -and $report.Result.NtlmRecords -eq 1 -and $report.Result.Gaps.Count -gt 0) 'Positive evidence survives report-assembly failure.'
    Reset-NtlmFixture
    $noEvidenceOutput = @(& $entryPoint -LogPath '' 6>&1)
    $noEvidenceReport = $noEvidenceOutput[-1]
    Assert-Ntlm ($noEvidenceOutput.Count -eq 5 -and [string]$noEvidenceOutput[0] -eq 'detected = false' -and $noEvidenceReport.Result.Assessment -eq 'CollectionError') 'Report-assembly failure with no matching records does not set detected.'
}
finally { Set-Item -Path function:Get-NtlmTopValues -Value $realAggregation }

Reset-NtlmFixture
$realMachineState = ${function:Get-NtlmMachineState}
try {
    function Get-NtlmMachineState {
        <#
        .SYNOPSIS
            Injects a collection failure before any records are read.
        .DESCRIPTION
            Verifies the fatal error path reports unknown counts, not detection.
        .OUTPUTS
            None. Always throws the synthetic failure.
        #>
        throw 'Synthetic machine-context failure.'
    }
    $fatalOutput = @(& $entryPoint -LogPath '' 6>&1)
    $fatalReport = $fatalOutput[-1]
    Assert-Ntlm ($fatalOutput.Count -eq 5 -and [string]$fatalOutput[0] -eq 'detected = false') 'Fatal collection still emits exactly four lines and does not set detected.'
    Assert-Ntlm ($fatalReport.Result.Assessment -eq 'CollectionError' -and $null -eq $fatalReport.Result.NtlmRecords -and $null -eq $fatalReport.Result.DerivedCredentialRecords) 'Fatal collection reports unknown evidence counts.'
    Assert-Ntlm ([string]$fatalOutput[3] -ceq 'found = NTLM records: Unknown | NTLMv1-derived SSO records: Unknown | Collection diagnostics: 1') 'Fatal collection preserves explicit Unknown counters.'
}
finally { Set-Item -Path function:Get-NtlmMachineState -Value $realMachineState }

foreach ($invalid in @(@{ Days = 0 }, @{ MaxEventsPerSource = 0 }, @{ MaxEvidence = -1 })) {
    $caught = $false
    try { & $entryPoint @invalid 6>$null | Out-Null } catch { $caught = $true }
    Assert-Ntlm $caught 'Reject invalid bounds before collection.'
}

foreach ($helpPath in @($sourcePath, $PSCommandPath)) {
    $helpTokens = $null
    $helpErrors = $null
    $helpAst = [System.Management.Automation.Language.Parser]::ParseFile($helpPath, [ref]$helpTokens, [ref]$helpErrors)
    Assert-Ntlm ($helpErrors.Count -eq 0) "Parse help-bearing file: $helpPath."
    $helpTargets = @($helpAst) + @($helpAst.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true))
    foreach ($helpTarget in $helpTargets) {
        $help = $helpTarget.GetHelpContent()
        $label = [IO.Path]::GetFileName($helpPath)
        $parameterBlock = $helpTarget.ParamBlock
        if ($helpTarget -is [System.Management.Automation.Language.FunctionDefinitionAst]) {
            $label = $helpTarget.Name
            $parameterBlock = $helpTarget.Body.ParamBlock
        }
        Assert-Ntlm ($null -ne $help -and -not [string]::IsNullOrWhiteSpace($help.Synopsis)) "Synopsis exists: $label."
        Assert-Ntlm (-not [string]::IsNullOrWhiteSpace($help.Description)) "Description exists: $label."
        Assert-Ntlm (@($help.Outputs).Count -gt 0) "Output contract exists: $label."
        foreach ($parameter in $parameterBlock.Parameters) {
            Assert-Ntlm (@($help.Parameters.Keys) -contains $parameter.Name.VariablePath.UserPath) "Parameter help: $label / $($parameter.Name.VariablePath.UserPath)."
        }
    }
}

Reset-NtlmFixture
$logTestPath = Join-Path $env:TEMP ('NtlmLogTest-' + [guid]::NewGuid().ToString() + '.log')
try {
    $report = Invoke-TestDetection @{ LogPath = $logTestPath }
    Assert-Ntlm ($report.Logging.Enabled -and -not $report.Logging.Error -and $report.Logging.EntriesWritten -gt 10) 'File logging is active without extra host lines.'
    $firstRunId = $report.Logging.RunId
    $firstLength = (Get-Item -LiteralPath $logTestPath).Length
    $logText = [IO.File]::ReadAllText($logTestPath)
    Assert-Ntlm ($logText -match 'Run start:' -and $logText -match 'Query complete:' -and $logText -match 'Assessment:' -and $logText -match 'Run end:') 'Log includes run, collection and result stages.'
    Assert-Ntlm ($logText -match '\d{4}-\d{2}-\d{2}T.*Z \[INFO\] \[Run:' -and $logText.Contains($firstRunId)) 'Log lines include UTC, severity and run ID.'
    $script:fixture.Records += New-TestEvent 4020 @{ Username = 'private-account-sample'; ProcessName = 'private-process-sample.exe' }
    $report = Invoke-TestDetection @{ LogPath = $logTestPath }
    $logText = [IO.File]::ReadAllText($logTestPath)
    Assert-Ntlm ((Get-Item -LiteralPath $logTestPath).Length -gt $firstLength -and $report.Logging.RunId -ne $firstRunId -and $logText.Contains($firstRunId)) 'Log appends and separates runs without discarding previous entries.'
    Assert-Ntlm (-not $logText.Contains('private-account-sample') -and -not $logText.Contains('private-process-sample.exe')) 'Diagnostic log omits raw identity and event samples.'
    Assert-Ntlm ([BitConverter]::ToString([IO.File]::ReadAllBytes($logTestPath), 0, 3) -eq 'EF-BB-BF') 'Diagnostic log is UTF-8 with BOM.'
    Reset-NtlmFixture
    $conflict = Invoke-TestDetection @{ LogPath = $logTestPath; ReportPath = $logTestPath }
    Assert-Ntlm ($conflict.Logging.Error -and $conflict.ReportFileError -and $conflict.Logging.EntriesWritten -eq 0) 'Reject shared log/JSON path before either writer changes it.'
    Assert-Ntlm ([IO.File]::ReadAllText($logTestPath) -ceq $logText) 'Path collision preserves the existing file.'
}
finally { if ([IO.File]::Exists($logTestPath)) { [IO.File]::Delete($logTestPath) } }

Reset-NtlmFixture
$unwritableLog = Join-Path $env:TEMP (([guid]::NewGuid().ToString()) + '\missing\trace.log')
$report = Invoke-TestDetection @{ LogPath = $unwritableLog }
Assert-Ntlm ($report.Logging.Error -and $report.Logging.EntriesWritten -eq 0 -and -not $report.Result.Detected) 'Log failure is visible but does not manufacture NTLM evidence or a coverage gap.'
$logFailureLines = @(Write-NtlmIvantiResult $report 6>&1)
Assert-Ntlm ($logFailureLines.Count -eq 4 -and [string]$logFailureLines[3] -notmatch 'LogError:|Access denied|trace.log') 'Log errors stay out of the default Ivanti summary.'
$disabledLogPath = Join-Path $env:TEMP ('NtlmDisabled-' + [guid]::NewGuid().ToString() + '.log')
$script:NtlmLogState = [pscustomobject]@{ Enabled = $false; Path = $disabledLogPath; Error = $null; EntriesWritten = 0; RunId = 'test' }
$silent = @(Write-NtlmLog -Message 'must not create a file' *>&1)
Assert-Ntlm ($silent.Count -eq 0 -and -not (Test-Path -LiteralPath $disabledLogPath)) 'Disabled logger has no file or stream side effects.'

Reset-NtlmFixture
$script:fixture.DeniedLog = 'Security'
$script:fixture.Records += New-TestEvent 4021 @{ NtlmVersion = 'NTLMv2'; ProcessName = 'detailed-app.exe'; TargetService = 'cifs/host.example'; NtlmUsageId = '7'; NtlmUsageReason = 'IP target' }
$report = Invoke-TestDetection
$beforeRendering = $report | ConvertTo-Json -Depth 12 -Compress
$rendered = @(Write-NtlmDetailedResult $report 6>&1)
$renderText = ($rendered | ForEach-Object { [string]$_ }) -join [Environment]::NewLine
foreach ($heading in @('NTLM Usage Detection', 'Evidence Counts', 'Collection diagnostics:', 'Recent Evidence')) {
    Assert-Ntlm ($renderText.Contains($heading)) "Detailed section: $heading."
}
foreach ($removedSection in @('Run Context', 'Event Sources', 'Log Retention', 'Current Registry Policies', 'Security Audit Policy', 'Registered Enhanced Provider Schemas', 'TopProcesses', 'Audit Change / Clear Markers')) {
    Assert-Ntlm (-not $renderText.Contains($removedSection)) "No diagnostic table in compact output: $removedSection."
}
Assert-Ntlm (@($renderText -split '\r?\n').Count -le 25) 'Single-evidence command-line report stays within 25 lines.'
Assert-Ntlm ($renderText -notmatch 'attestation|reliable negative assessment|unique authentications|Interpretation Limits|remains unproven|Not Proof Of Activation') 'Detailed output excludes general interpretation commentary.'
$ivantiText = (@(Write-NtlmIvantiResult $report 6>&1) | ForEach-Object { [string]$_ }) -join [Environment]::NewLine
Assert-Ntlm ($ivantiText.Contains('expected = NTLM records: 0 | NTLMv1-derived SSO records: 0') -and $ivantiText -notmatch 'Assessment:|WindowDays:|FirstGap:|Latest:|detailed-app.exe|cifs/host.example') 'Default output contains only the concise Ivanti summary, not diagnostic detail.'
$failedReport = [pscustomobject]@{
    Result = [pscustomobject]@{ Detected = $false; Reason = 'Collection failed.'; NtlmRecords = $null; DerivedCredentialRecords = $null; Gaps = @('Synthetic read failure') }
}
$failedLines = @(Write-NtlmIvantiResult $failedReport 6>&1)
Assert-Ntlm ($failedLines.Count -eq 4 -and [string]$failedLines[3] -ceq 'found = NTLM records: Unknown | NTLMv1-derived SSO records: Unknown | Collection diagnostics: 1') 'Failed collection retains the four-line format without inventing zero counts.'
Assert-Ntlm ($renderText.Contains('detailed-app.exe') -and $renderText.Contains('cifs/host.example') -and $renderText.Contains('IP target')) 'Rich output includes caller, target and reason.'
Assert-Ntlm (@($rendered | Where-Object { $_ -isnot [System.Management.Automation.InformationRecord] }).Count -eq 0) 'Rich renderer leaks no objects or formatting records to the success stream.'
Assert-Ntlm (($report | ConvertTo-Json -Depth 12 -Compress) -ceq $beforeRendering) 'Rich rendering does not mutate the assessment or report.'
foreach ($switchName in @('Detailed', 'CmdLine')) {
    $detailedArguments = @{ LogPath = ''; MaxEvidence = 0; $switchName = $true }
    $detailedOutput = @(& $entryPoint @detailedArguments 6>&1)
    $detailedReport = $detailedOutput[-1]
    $detailedText = ($detailedOutput[0..($detailedOutput.Count - 2)] | ForEach-Object { [string]$_ }) -join [Environment]::NewLine
    Assert-Ntlm ($detailedText.Contains('NTLM Usage Detection') -and $detailedText -notmatch '(?m)^detected = ') "Opt-in renderer selected by $switchName."
    Assert-Ntlm ($detailedReport.Result.Assessment -eq $report.Result.Assessment -and $detailedReport.Result.NtlmRecords -eq $report.Result.NtlmRecords) 'Presentation switch does not change detection.'
    Assert-Ntlm ($detailedText.Contains('No samples retained') -and $detailedReport.Evidence.Count -eq 0) 'Zero samples is not described as zero NTLM evidence.'
}
Reset-NtlmFixture
foreach ($recordId in 1..10) {
    $script:fixture.Records += New-TestEvent 4020 @{ ProcessName = "app-$recordId.exe"; NtlmUsageReason = 'Fixture reason' } -RecordId $recordId
}
$manySamples = Invoke-TestDetection
$manyText = (@(Write-NtlmDetailedResult $manySamples 6>&1) | ForEach-Object { [string]$_ }) -join [Environment]::NewLine
Assert-Ntlm ([regex]::Matches($manyText, 'Event=4020 Record=').Count -eq 3 -and $manySamples.Evidence.Count -eq 10) 'Console limits samples to three without discarding JSON evidence.'
Assert-Ntlm (@($manyText -split '\r?\n').Count -le 32) 'Console line count stays bounded for a larger evidence set.'
Write-Output "PASS: $script:assertions assertions; PowerShell $($PSVersionTable.PSVersion); mocked Windows reads only."