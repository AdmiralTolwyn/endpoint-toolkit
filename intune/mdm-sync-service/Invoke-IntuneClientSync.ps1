<#
.SYNOPSIS
    Collects local Intune sync evidence and optionally requests an MDM check-in.

.DESCRIPTION
    Diagnostic by default. Collects in 64-bit Windows PowerShell 5.1 or PowerShell
    7 on Windows, under SYSTEM or an elevated administrator, including MECM and
    Nexthink execution. Submits the existing enrollment-specific PushLaunch task
    directly through ScheduledTasks. No WinRT or child-process bridge is used.
    Uses existing enrollment credentials; no Graph authentication is required.

    Collects OS and Entra join details, primary and linked MDM enrollments,
    referenced certificates, services, task activity, proxy configuration,
    discovery-host connectivity, IME log activity and recent MDM errors.
    Missing evidence is reported separately from an observed prerequisite issue.

    With -Sync, submits at most one task when the local prerequisite assessment
    permits it and ShouldProcess approves. Missing, ambiguous, disabled or changed
    task definitions fail without guessing. Already-running tasks are not restarted.
    Observes a newer task run for up to TimeoutSeconds; an old scheduler result
    cannot establish completion. Task completion is not proof of MDM sync success.

    Does not repair enrollment, change services, clear caches or restart IME.
    Diagnostic collection does make DNS/direct TCP probes. Requesting sync can
    cause Windows to process previously assigned management actions and policies.
    Use the common -Verbose switch to see collection progress, skip/failure
    reasons and sync outcomes on the verbose stream, separate from report output.
    Always appends progress and the full report to its own IME log file, including
    diagnostic-only and WhatIf runs. OutputPath remains an optional JSON export.

.PARAMETER Sync
    Submit one enrollment-specific PushLaunch task. Omitted by default.
    -WhatIf suppresses task submission, but still collects evidence and writes
    OutputPath when supplied. Confirmation applies to the sync request only.

.PARAMETER TimeoutSeconds
    Task observation deadline, 1-600 seconds, default 120. No task is stopped or
    retried on timeout. Does not bound individual scheduler calls or the entire
    assessment. Ignored in diagnostic-only mode.

.PARAMETER OutputPath
    Optional full JSON report path, independent of the selected OutputFormat.
    Relative paths use the current working directory. Parent directories are
    created and an existing report is overwritten. The file uses UTF-8 with BOM.
    Report-write failure returns ScriptError (exit 2), even if sync already ran.

.PARAMETER OutputFormat
    SummaryJson (default): compact JSON summary for management agents.
    DetailedJson: indented JSON with all collected evidence and full errors.
    Object: the full PSCustomObject for interactive inspection and pipelines.
    Does not change collection, prerequisites, sync behavior or exit codes.

.INPUTS
    None. Pipeline input is not supported.

.OUTPUTS
    System.String
    SummaryJson emits one compressed summary; DetailedJson emits the full report
    as indented JSON. Summary errors are limited to three messages, 200 characters
    each. DetailedJson and OutputPath preserve full errors and diagnostic evidence.

    System.Management.Automation.PSCustomObject
    Object emits the full report without formatting or serialization. Capture it
    when invoking this script in the current PowerShell process. A separate
    powershell.exe process returns text, not a live PowerShell object.
    -WhatIf and interactive confirmation can add host text before the JSON.

.EXAMPLE
    .\Invoke-IntuneClientSync.ps1

    Collect diagnostic evidence and emit a JSON summary without requesting sync.

.EXAMPLE
    .\Invoke-IntuneClientSync.ps1 -OutputPath 'C:\ProgramData\IntuneSync\health.json'

    Collect evidence and retain the full report for investigation.

.EXAMPLE
    .\Invoke-IntuneClientSync.ps1 -Sync -TimeoutSeconds 120 -OutputPath 'C:\ProgramData\IntuneSync\sync.json'

    Collect evidence, request one sync if permitted, and report its observed result.

.EXAMPLE
    .\Invoke-IntuneClientSync.ps1 -Sync -WhatIf

    Collect evidence and show the proposed request without submitting a task.

.EXAMPLE
    .\Invoke-IntuneClientSync.ps1 -OutputFormat DetailedJson

    Display the full report as readable JSON, including exact execution-context
    failures when collection cannot proceed.

.EXAMPLE
    $report = .\Invoke-IntuneClientSync.ps1 -OutputFormat Object
    $report.ExecutionContext | Format-List
    $report.Services | Format-Table -AutoSize

    Capture the full report and inspect its nested properties interactively.

.EXAMPLE
    .\Invoke-IntuneClientSync.ps1 -OutputFormat Object -Verbose

    Show diagnostic progress while returning the full report object. Does not sync.

.NOTES
    Collection requires 64-bit PowerShell on Windows as SYSTEM or an elevated
    administrator. Edition is reported, not rejected. Requires ScheduledTasks.
    The selected task runs under its existing SYSTEM principal. No task is created
    or modified. WhatIf/declined confirmation and failed prerequisites submit none.
    Execute as a script; the entry point emits a result and exits with a code.
    It is not intended to be dot-sourced as a function library.
    ExecutionContext reports actual edition, bitness and administrator/SYSTEM
    status; PowerShellVersion reports the version. Failed requirements are listed
    separately in CollectionErrors. Formatting options do not bypass these checks.

    Task submission is not proof that sync completed, policies applied or the Intune
    console updated. Task run times and IME log writes are activity, not success.
    CloudLastSyncVerified and PolicyApplicationVerified always remain false.
    EvidenceComplete describes collection completeness, not device health.
    Log: C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\Invoke-IntuneClientSync.log.
    Creates the folder if missing. Log-write failures warn without changing exit codes.
    Logs contain the full diagnostic evidence; protect them like exported reports.

    Exit codes:
    0 - Assessment complete or a newer PushLaunch run finished with result zero.
    1 - Local prerequisite issue, task operation failure or nonzero task result.
    2 - Incomplete evidence, unsupported context, busy task, timeout or script error.

    Offline tests: scripts\Test-IntuneClientSync.ps1.
    Deployment, result interpretation and limits: docs\INTUNE-CLIENT-SYNC.md.
    Live SYSTEM-agent sync and Intune console propagation still require a pilot.

.LINK
    https://learn.microsoft.com/en-us/powershell/module/scheduledtasks/start-scheduledtask

.LINK
    https://learn.microsoft.com/en-us/troubleshoot/mem/intune/device-management/cannot-sync-windows-10-devices
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$Sync,
    [ValidateRange(1, 600)]
    [int]$TimeoutSeconds = 120,
    [string]$OutputPath,
    [ValidateSet('SummaryJson', 'DetailedJson', 'Object')]
    [string]$OutputFormat = 'SummaryJson'
)

#region Helpers

function Write-ClientSyncLog {
    <#
    .SYNOPSIS
        Appends a UTC-stamped message without changing the success output stream.
    .PARAMETER Message
        Text or a report object to append; objects use compressed JSON.
    .PARAMETER OutputMessage
        Also emits the original message on the success stream for service scripts.
    .NOTES
        A write failure warns once per run; it does not change the operation result.
    #>
    param([object]$Message, [switch]$OutputMessage)
    try {
        [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($script:ClientSyncLogPath))
        $text = $Message
        if ($Message -isnot [string]) { $text = $Message | ConvertTo-Json -Depth 10 -Compress }
        $line = '{0} [PID:{1}] {2}{3}' -f [DateTime]::UtcNow.ToString('o'), $PID, $text, [Environment]::NewLine
        [IO.File]::AppendAllText($script:ClientSyncLogPath, $line, (New-Object System.Text.UTF8Encoding($false)))
    }
    catch {
        if (-not $script:ClientSyncLogWarningWritten) {
            $script:ClientSyncLogWarningWritten = $true
            Write-Warning "Cannot write log '$script:ClientSyncLogPath': $($_.Exception.Message)" -WarningAction Continue
        }
    }
    if ($OutputMessage) { Write-Output $Message }
}

function Invoke-ClientMdmTask {
    <#
    .SYNOPSIS
        Validates, submits and observes the primary Intune enrollment's PushLaunch task.
    .DESCRIPTION
        Caller must validate prerequisites and obtain any ShouldProcess approval.
        Selects MS DM Server with an OMA-DM account; never guesses among multiple
        candidates or starts linked-provider tasks. Requires an enabled, demand-
        startable SYSTEM task whose sole action matches the enrollment GUID and
        the observed deviceenroller PushLaunch command. Never modifies tasks.
    .PARAMETER Enrollments
        Collected enrollment objects with EnrollmentId, ProviderId, TenantId and
        OmaDmAccountExists. A missing or ambiguous primary target is a failure.
    .PARAMETER WaitSeconds
        Observation deadline after submission, default 120 seconds. Does not stop
        the task or bound individual scheduler provider calls.
    .OUTPUTS
        PSCustomObject: Status (Completed, Failed, AlreadyRunning or TimedOut),
        target, stage, submission acceptance, baseline/observed run times and
        last task result. TaskState is pre-submission; ObservedTaskState is latest.
        Completion requires a newer run and an idle task, not an old result.
        Scheduler history is task-wide, not an instance ID or proof of MDM sync.
    #>
    param([object[]]$Enrollments, [ValidateRange(1, 600)][int]$WaitSeconds = 120)

    $result = [ordered]@{
        Status = 'Failed'
        Stage = 'SelectEnrollment'
        EnrollmentId = $null
        TenantId = $null
        TaskPath = $null
        TaskName = 'PushLaunch'
        TaskState = $null
        SubmissionAccepted = $false
        BaselineLastRunTimeUtc = $null
        ObservedLastRunTimeUtc = $null
        ObservedTaskState = $null
        NewRunObserved = $false
        LastTaskResult = $null
        LastTaskResultHex = $null
        HResult = $null
        Error = $null
        RequestedUtc = $null
        FinishedUtc = $null
    }
    try {
        $primary = @($Enrollments | Where-Object { $_.ProviderId -eq 'MS DM Server' -and $_.OmaDmAccountExists })
        if ($primary.Count -ne 1) { throw "Expected one primary Intune enrollment with an OMA-DM account; found $($primary.Count). No task submitted." }
        $enrollment = $primary[0]
        $id = [string]$enrollment.EnrollmentId
        if ($id -notmatch '^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$') { throw 'Invalid enrollment ID. No task submitted.' }
        $result.EnrollmentId = $id
        $result.TenantId = $enrollment.TenantId
        $result.TaskPath = "\Microsoft\Windows\EnterpriseMgmt\$id\"
        $result.Stage = 'ValidateTask'
        Write-Verbose "Validating $($result.TaskPath)$($result.TaskName) for tenant $($result.TenantId)."
        $tasks = @(Get-ScheduledTask -TaskPath $result.TaskPath -TaskName $result.TaskName -ErrorAction Stop)
        if ($tasks.Count -ne 1) { throw 'Expected exactly one enrollment PushLaunch task.' }
        $task = $tasks[0]
        if ($task.TaskPath -ne $result.TaskPath -or $task.TaskName -ne $result.TaskName) { throw 'Returned task does not match the selected enrollment.' }
        $result.TaskState = [string]$task.State
        if (-not $task.Settings.Enabled -or $result.TaskState -eq 'Disabled') { throw 'PushLaunch is disabled. It was not enabled or started.' }
        if (-not $task.Settings.AllowDemandStart) { throw 'PushLaunch does not allow on-demand starts.' }
        if ($task.Principal.UserId -notin @('SYSTEM', 'NT AUTHORITY\SYSTEM', 'S-1-5-18')) { throw 'PushLaunch does not run as SYSTEM.' }
        $actions = @($task.Actions)
        if ($actions.Count -ne 1) { throw 'Unexpected PushLaunch action count.' }
        $executable = [Environment]::ExpandEnvironmentVariables([string]$actions[0].Execute)
        $expectedExecutable = Join-Path $env:SystemRoot 'System32\deviceenroller.exe'
        $argumentPattern = '^\s*/o\s+(?:"{0}"|{0})\s+/c\s+/z\s*$' -f [regex]::Escape($id)
        if ($executable -ine $expectedExecutable -or $actions[0].Arguments -notmatch $argumentPattern) { throw 'PushLaunch action does not match the expected enrollment-specific command.' }
        if ($result.TaskState -in @('Running', 'Queued')) {
            $result.Status = 'AlreadyRunning'
            Write-Verbose 'PushLaunch is already running/queued; no new instance submitted.'
        }
        else {
            if ($result.TaskState -ne 'Ready') { throw "PushLaunch is not ready (state: $($result.TaskState))." }
            $result.Stage = 'ReadTaskBaseline'
            $baseline = Get-ScheduledTaskInfo -InputObject $task -ErrorAction Stop
            if ($null -eq $baseline -or $baseline.LastRunTime -isnot [DateTime]) { throw 'PushLaunch baseline run time is unavailable. No task submitted.' }
            $baselineTime = $baseline.LastRunTime
            $result.BaselineLastRunTimeUtc = $baselineTime.ToUniversalTime().ToString('o')
            $result.Stage = 'SubmitTask'
            $result.RequestedUtc = [DateTime]::UtcNow.ToString('o')
            Start-ScheduledTask -InputObject $task -ErrorAction Stop
            $result.SubmissionAccepted = $true
            $result.Stage = 'ObserveTask'
            $result.Status = 'TimedOut'
            $clock = [Diagnostics.Stopwatch]::StartNew()
            Write-Verbose "PushLaunch submitted once; observing a newer task run for up to $WaitSeconds seconds."
            do {
                $before = @(Get-ScheduledTask -TaskPath $result.TaskPath -TaskName $result.TaskName -ErrorAction Stop)
                if ($before.Count -ne 1) { throw 'PushLaunch disappeared during observation.' }
                $info = Get-ScheduledTaskInfo -InputObject $before[0] -ErrorAction Stop
                $after = @(Get-ScheduledTask -TaskPath $result.TaskPath -TaskName $result.TaskName -ErrorAction Stop)
                if ($after.Count -ne 1) { throw 'PushLaunch disappeared during observation.' }
                if ($null -eq $info -or $info.LastRunTime -isnot [DateTime] -or $null -eq $info.LastTaskResult) { throw 'PushLaunch run information is unavailable.' }
                $result.ObservedTaskState = [string]$after[0].State
                $result.ObservedLastRunTimeUtc = $info.LastRunTime.ToUniversalTime().ToString('o')
                if ($info.LastRunTime -gt $baselineTime) {
                    $result.NewRunObserved = $true
                    $result.LastTaskResult = [long]$info.LastTaskResult -band 0xFFFFFFFFL
                    $result.LastTaskResultHex = '0x{0:X8}' -f $result.LastTaskResult
                    if ([string]$before[0].State -eq 'Ready' -and $result.ObservedTaskState -eq 'Ready' -and $result.LastTaskResult -notin @(0x00041301, 0x00041303, 0x00041325)) {
                        if ($result.LastTaskResult -eq 0) { $result.Status = 'Completed' }
                        else {
                            $result.Status = 'Failed'
                            $result.Error = "PushLaunch finished with task result $($result.LastTaskResultHex)."
                        }
                        break
                    }
                }
                $remaining = ($WaitSeconds * 1000) - $clock.ElapsedMilliseconds
                if ($remaining -gt 0) { [Threading.Tasks.Task]::Delay([int][Math]::Min(250, $remaining)).Wait() }
            } while ($clock.Elapsed.TotalSeconds -lt $WaitSeconds)
            $clock.Stop()
            if ($result.Status -eq 'TimedOut') { $result.Error = "A newer completed PushLaunch run was not observed within $WaitSeconds seconds. The task was not stopped or retried." }
            Write-Verbose "PushLaunch observation: $($result.Status); task result: $($result.LastTaskResultHex). MDM sync remains unverified."
        }
    }
    catch {
        $result.Status = 'Failed'
        $failure = $_.Exception.GetBaseException()
        $result.HResult = '0x{0:X8}' -f $failure.HResult
        $result.Error = $failure.Message
    }
    $result.FinishedUtc = [DateTime]::UtcNow.ToString('o')
    [pscustomobject]$result
}

function Get-ClientValue {
    <#
    .SYNOPSIS
        Reads an optional object property without a StrictMode missing-member error.
    .PARAMETER InputObject
        Object to inspect; null is allowed.
    .PARAMETER Name
        Literal property name to read.
    .OUTPUTS
        The property value, or explicit null if the object/property is absent.
        Explicit null prevents AutomationNull from serializing as an empty object.
    #>
    param(
        $InputObject,
        [string]$Name
    )

    if ($null -ne $InputObject -and $InputObject.PSObject.Properties[$Name]) {
        return $InputObject.$Name
    }
    return $null
}

function Get-ClientExecutionContext {
    <#
    .SYNOPSIS
        Reports the current token, process architecture and PowerShell edition.
    .OUTPUTS
        PSCustomObject with IsSystem, IsAdministrator, Is64Bit and Edition.
        Performs no elevation or identity change; token-query errors propagate.
    #>
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    try {
        $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
        [pscustomobject]@{
            IsSystem = ($identity.User.Value -eq 'S-1-5-18')
            IsAdministrator = $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
            Is64Bit = [Environment]::Is64BitProcess
            Edition = $PSVersionTable.PSEdition
        }
    }
    finally { $identity.Dispose() }
}

function Invoke-ClientReadCommand {
    <#
    .SYNOPSIS
        Captures output from a read-only native command with a 15-second wait.
    .DESCRIPTION
        Used for dsregcmd /status and netsh winhttp show proxy. Drains stdout and
        stderr asynchronously to avoid full-pipe deadlocks. On timeout, terminates
        only the process it started. This helper does not enforce read-only use;
        callers must supply a diagnostic command.
    .PARAMETER FileName
        Full path to the diagnostic executable.
    .PARAMETER Arguments
        Native argument string passed directly to the executable, without a shell.
    .OUTPUTS
        System.String containing stdout. Throws on timeout, launch failure or
        nonzero exit; the failure message includes stderr when available.
    #>
    param(
        [string]$FileName,
        [string]$Arguments
    )

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo.FileName = $FileName
    $process.StartInfo.Arguments = $Arguments
    $process.StartInfo.UseShellExecute = $false
    $process.StartInfo.CreateNoWindow = $true
    $process.StartInfo.RedirectStandardOutput = $true
    $process.StartInfo.RedirectStandardError = $true
    try {
        [void]$process.Start()
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit(15000)) {
            $process.Kill()
            throw "Read-only command timed out: $FileName"
        }
        if ($process.ExitCode -ne 0) { throw "Read-only command failed: $FileName; exit $($process.ExitCode); $($stderr.Result)" }
        $stdout.Result
    }
    finally { $process.Dispose() }
}

#endregion Helpers

#region Collectors

function Get-ClientJoinEvidence {
    <#
    .SYNOPSIS
        Reads selected device and tenant fields from dsregcmd /status.
    .DESCRIPTION
        Uses the first occurrence of each field, avoiding later repeated fields
        from user/workplace sections. Missing fields remain null. Does not infer
        successful enrollment from MdmUrl or evaluate the user's PRT as SYSTEM.
    .OUTPUTS
        PSCustomObject with AzureAdJoined, DomainJoined, DeviceId, TenantId,
        DeviceAuthStatus and MdmUrl. Native command errors propagate.
    #>
    $text = Invoke-ClientReadCommand -FileName "$env:SystemRoot\System32\dsregcmd.exe" -Arguments '/status'
    $join = [ordered]@{}
    foreach ($field in @('AzureAdJoined', 'DomainJoined', 'DeviceId', 'TenantId', 'DeviceAuthStatus', 'MdmUrl')) {
        $fieldMatch = [regex]::Match($text, ('(?m)^\s*{0}\s*:\s*([^\r\n]*)' -f [regex]::Escape($field)))
        $join[$field] = $null
        if ($fieldMatch.Success) { $join[$field] = $fieldMatch.Groups[1].Value.Trim() }
    }
    [pscustomobject]$join
}

function Get-ClientCertificateEvidence {
    <#
    .SYNOPSIS
        Checks a referenced certificate in LocalMachine\My without exporting a key.
    .PARAMETER Thumbprint
        Normalized certificate thumbprint. Empty means no reference was found.
    .OUTPUTS
        PSCustomObject with Status, Thumbprint, Subject, NotBeforeUtc, NotAfterUtc
        and HasPrivateKey. Status is MissingReference, Missing, NoPrivateKey,
        NotYetValid, Expired or Present; date failures take priority over key state.
    .NOTES
        Present means dates and HasPrivateKey passed local checks only. Does not
        test signing, certificate trust/revocation, server acceptance or clock
        accuracy. Certificate-store access errors propagate rather than becoming
        a misleading Missing result.
    #>
    param([string]$Thumbprint)

    $result = [ordered]@{
        Status        = 'MissingReference'
        Thumbprint    = $Thumbprint
        Subject       = $null
        NotBeforeUtc  = $null
        NotAfterUtc   = $null
        HasPrivateKey = $null
    }
    if ($Thumbprint) {
        $certPath = "Cert:\LocalMachine\My\$Thumbprint"
        $result.Status = 'Missing'
        if (Test-Path -LiteralPath $certPath -ErrorAction Stop) {
            $certificate = Get-Item -LiteralPath $certPath -ErrorAction Stop
            $result.Subject = $certificate.Subject
            $result.NotBeforeUtc = $certificate.NotBefore.ToUniversalTime().ToString('o')
            $result.NotAfterUtc = $certificate.NotAfter.ToUniversalTime().ToString('o')
            $result.HasPrivateKey = $certificate.HasPrivateKey
            $result.Status = 'Present'
            if (-not $certificate.HasPrivateKey) { $result.Status = 'NoPrivateKey' }
            if ($certificate.NotBefore.ToUniversalTime() -gt [DateTime]::UtcNow) { $result.Status = 'NotYetValid' }
            if ($certificate.NotAfter.ToUniversalTime() -le [DateTime]::UtcNow) { $result.Status = 'Expired' }
        }
    }
    [pscustomobject]$result
}

function Get-ClientEnrollmentEvidence {
    <#
    .SYNOPSIS
        Collects primary Intune and linked Windows management enrollment evidence.
    .DESCRIPTION
        Reads GUID-named enrollment keys for MS DM Server and Microsoft Device
        Management, with matching OMA-DM accounts. Prefers DMPCertThumbPrint and
        falls back to a recognized MY;System certificate reference. Reports
        conflicting or unrecognized references without trying to repair them.
    .OUTPUTS
        Zero or more PSCustomObjects containing EnrollmentId, ProviderId, TenantId,
        OmaDmAccountExists, DiscoveryUrl, Certificate and certificate-reference
        flags. The caller must use @() when it needs a stable array.
        Access failures and malformed thumbprints throw, not imply no enrollment.
    .NOTES
        Registry layout is diagnostic implementation evidence, not a public API.
        A registry entry or account alone does not establish successful sync.
    #>
    $root = 'HKLM:\SOFTWARE\Microsoft\Enrollments'
    if (-not (Test-Path -LiteralPath $root -ErrorAction Stop)) { return }
    foreach ($key in @(Get-ChildItem -LiteralPath $root -ErrorAction Stop)) {
        if ($key.PSChildName -notmatch '^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$') { continue }
        $values = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction Stop
        $provider = [string](Get-ClientValue $values 'ProviderID')
        if ($provider -notin @('MS DM Server', 'Microsoft Device Management')) { continue }
        $accountPath = "HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\$($key.PSChildName)"
        $accountExists = Test-Path -LiteralPath $accountPath -ErrorAction Stop
        $account = $null
        if ($accountExists) { $account = Get-ItemProperty -LiteralPath $accountPath -ErrorAction Stop }
        $thumbprint = ([string](Get-ClientValue $values 'DMPCertThumbPrint')) -replace '\s', ''
        $reference = [string](Get-ClientValue $account 'SslClientCertReference')
        $referenceMatch = [regex]::Match($reference, '(?i)^MY;System;([a-f0-9\s]+)$')
        $referenceMismatch = $false
        if ($referenceMatch.Success) {
            $accountThumbprint = $referenceMatch.Groups[1].Value -replace '\s', ''
            if (-not $thumbprint) { $thumbprint = $accountThumbprint }
            elseif ($thumbprint -ne $accountThumbprint) { $referenceMismatch = $true }
        }
        if ($thumbprint -and $thumbprint -notmatch '^[a-f0-9]{40}$') { throw "Unexpected certificate thumbprint in enrollment $($key.PSChildName)" }
        [pscustomobject]@{
            EnrollmentId = $key.PSChildName
            ProviderId = $provider
            TenantId = [string](Get-ClientValue $values 'AADTenantID')
            OmaDmAccountExists = [bool]$accountExists
            DiscoveryUrl = [string](Get-ClientValue $values 'DiscoveryServiceFullURL')
            Certificate = Get-ClientCertificateEvidence -Thumbprint $thumbprint
            CertificateReferenceMismatch = $referenceMismatch
            UnrecognizedCertificateReference = [bool]($reference -and -not $referenceMatch.Success)
        }
    }
}

function Get-ClientServiceEvidence {
    <#
    .SYNOPSIS
        Reads MDM-related and IME service state without changing any service.
    .OUTPUTS
        One PSCustomObject per expected service: Name, Present, State, StartMode.
        Missing services have Present=false and null state/start mode. CIM errors
        propagate. Stopped and Disabled are deliberately kept distinct.
    #>
    $names = @('Schedule', 'dmwappushservice', 'DmEnrollmentSvc', 'IntuneManagementExtension')
    $filter = ($names | ForEach-Object { "Name='$_'" }) -join ' OR '
    $services = @(Get-CimInstance -ClassName Win32_Service -Filter $filter -OperationTimeoutSec 10 -ErrorAction Stop)
    foreach ($name in $names) {
        $service = $services | Where-Object { $_.Name -eq $name } | Select-Object -First 1
        [pscustomobject]@{
            Name = $name
            Present = [bool]$service
            State = Get-ClientValue $service 'State'
            StartMode = Get-ClientValue $service 'StartMode'
        }
    }
}

function Get-ClientTaskEvidence {
    <#
    .SYNOPSIS
        Reads EnterpriseMgmt task state and scheduler history without starting tasks.
    .OUTPUTS
        Zero or more PSCustomObjects with Path, Name, State, LastRunUtc and
        LastTaskResult. Never-run sentinel dates become null. Query errors throw.
    .NOTES
        A task run or scheduler result is activity evidence, not proof of a
        successful MDM exchange. Does not inspect EnterpriseMgmtNonCritical tasks.
    #>
    $tasks = @(Get-ScheduledTask -ErrorAction Stop | Where-Object { $_.TaskPath -like '\Microsoft\Windows\EnterpriseMgmt\*' })
    foreach ($task in $tasks) {
        $info = Get-ScheduledTaskInfo -InputObject $task -ErrorAction Stop
        $lastRun = $null
        if ($info.LastRunTime.Year -gt 1999) { $lastRun = $info.LastRunTime.ToUniversalTime().ToString('o') }
        [pscustomobject]@{
            Path = $task.TaskPath
            Name = $task.TaskName
            State = [string]$task.State
            LastRunUtc = $lastRun
            LastTaskResult = $info.LastTaskResult
        }
    }
}

function Get-ClientMdmEvents {
    <#
    .SYNOPSIS
        Reads up to 12 recent MDM Admin warnings/errors for diagnostic context.
    .PARAMETER Since
        Earliest event time to include. The entry point supplies three days ago.
    .OUTPUTS
        Zero or more PSCustomObjects with RecordId, EventId, TimeUtc and Message.
        Messages are limited to 1,200 characters. No matching events is an empty
        result; other event-log errors propagate to the caller.
    .NOTES
        Events are not correlated to the requested task and may predate it.
        An empty result is not proof of health or absence of an earlier failure.
    #>
    param([DateTime]$Since)

    try {
        Get-WinEvent -FilterHashtable @{
            LogName = 'Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin'
            StartTime = $Since
            Level = @(2, 3)
        } -MaxEvents 12 -ErrorAction Stop | ForEach-Object {
            $message = [string]$_.Message
            if ($message.Length -gt 1200) { $message = $message.Substring(0, 1200) }
            [pscustomobject]@{
                RecordId = $_.RecordId
                EventId  = $_.Id
                TimeUtc  = $_.TimeCreated.ToUniversalTime().ToString('o')
                Message  = $message
            }
        }
    }
    catch {
        if ($_.FullyQualifiedErrorId -notlike 'NoMatchingEventsFound*') { throw }
    }
}

function Get-ClientImeLogEvidence {
    <#
    .SYNOPSIS
        Reads existence and last-write times of three IME logs, not their contents.
    .OUTPUTS
        PSCustomObjects with Name, Present and LastWriteUtc for the main IME,
        AppWorkload and HealthScripts logs. Missing files have a null timestamp;
        access failures propagate. Log activity does not prove workload success.
    #>
    foreach ($name in @('IntuneManagementExtension.log', 'AppWorkload.log', 'HealthScripts.log')) {
        $path = Join-Path $env:ProgramData "Microsoft\IntuneManagementExtension\Logs\$name"
        $file = $null
        if (Test-Path -LiteralPath $path -ErrorAction Stop) { $file = Get-Item -LiteralPath $path -ErrorAction Stop }
        [pscustomobject]@{
            Name = $name
            Present = [bool]$file
            LastWriteUtc = $(if ($file) { $file.LastWriteTimeUtc.ToString('o') } else { $null })
        }
    }
}

function Test-ClientDiscoveryEndpoint {
    <#
    .SYNOPSIS
        Probes a discovery host using bounded DNS and direct TCP operations.
    .PARAMETER Uri
        Absolute HTTPS discovery URI selected by the caller. Only its host and
        port are used; no HTTP request, credentials or TLS handshake is sent.
    .OUTPUTS
        PSCustomObject with HostName, Port, Scope, Dns, Tcp and Error. DNS and TCP
        each receive a three-second observation wait. Exceptions/timeouts populate
        Error; Dns/Tcp retain the last observed state rather than a success guess.
    .NOTES
        Does not use or validate a proxy path. Direct failure can coexist with
        working proxied MDM. DNS/TCP success cannot establish MDM or IME health.
    #>
    param([uri]$Uri)

    $result = [ordered]@{
        HostName = $Uri.DnsSafeHost
        Port     = $Uri.Port
        Scope    = 'DirectDiscoveryOnly'
        Dns      = 'Unknown'
        Tcp      = 'NotTested'
        Error    = $null
    }
    $client = $null
    try {
        $dnsTask = [System.Net.Dns]::GetHostAddressesAsync($Uri.DnsSafeHost)
        if (-not $dnsTask.Wait(3000)) { throw 'DNS observation timed out.' }
        if ($dnsTask.Result.Count -eq 0) { throw 'DNS returned no addresses.' }
        $result.Dns = 'Resolved'
        $client = New-Object System.Net.Sockets.TcpClient
        $connectTask = $client.ConnectAsync($Uri.DnsSafeHost, $Uri.Port)
        if (-not $connectTask.Wait(3000)) { throw 'Direct TCP observation timed out.' }
        $result.Tcp = 'Connected'
    }
    catch { $result.Error = $_.Exception.GetBaseException().Message }
    finally { if ($client) { $client.Dispose() } }
    [pscustomobject]$result
}

#endregion Collectors

#region Assessment

function Get-ClientSyncAssessment {
    <#
    .SYNOPSIS
        Classifies collected prerequisites without querying or changing the device.
    .DESCRIPTION
        Issues and Unknowns prevent this tool from requesting sync. Observations
        remain advisory. Evaluates primary accounts separately from linked
        enrollments, and IME separately from MDM. Collector flags distinguish
        unavailable evidence from a successfully collected empty result.
    .PARAMETER Enrollments
        Enrollment objects from Get-ClientEnrollmentEvidence.
    .PARAMETER Services
        Service objects from Get-ClientServiceEvidence.
    .PARAMETER Join
        Optional dsregcmd fields from Get-ClientJoinEvidence.
    .PARAMETER EnrollmentRead
        True only when enrollment collection completed without throwing.
    .PARAMETER ServiceRead
        True only when service collection completed without throwing.
    .PARAMETER Tasks
        Task objects from Get-ClientTaskEvidence; used for advisory observations.
    .PARAMETER TaskRead
        True only when task collection completed without throwing.
    .OUTPUTS
        PSCustomObject with CanSync and string arrays Issues, Unknowns and
        Observations. CanSync is a local safety gate, not proof sync will succeed.
    #>
    param(
        [object[]]$Enrollments,
        [object[]]$Services,
        $Join,
        [bool]$EnrollmentRead,
        [bool]$ServiceRead,
        [object[]]$Tasks,
        [bool]$TaskRead
    )

    $issues = New-Object 'System.Collections.Generic.List[string]'
    $unknowns = New-Object 'System.Collections.Generic.List[string]'
    $observations = New-Object 'System.Collections.Generic.List[string]'
    $primary = @($Enrollments | Where-Object { $_.ProviderId -eq 'MS DM Server' -and $_.OmaDmAccountExists })
    $linked = @($Enrollments | Where-Object { $_.ProviderId -eq 'Microsoft Device Management' -and $_.OmaDmAccountExists })
    if (-not $EnrollmentRead) { [void]$unknowns.Add('EnrollmentEvidenceUnavailable') }
    elseif ($primary.Count -eq 0) {
        [void]$issues.Add('NoPrimaryIntuneOmaDmAccount')
        if ($linked.Count -gt 0) { [void]$issues.Add('LinkedEnrollmentWithoutPrimaryAccount') }
    }
    elseif ($primary.Count -gt 1) { [void]$unknowns.Add('MultiplePrimaryAccountsRequireInvestigation') }
    else {
        $enrollment = $primary[0]
        if ($enrollment.Certificate.Status -ne 'Present') { [void]$issues.Add("PrimaryCertificate:$($enrollment.Certificate.Status)") }
        if ($enrollment.CertificateReferenceMismatch -or $enrollment.UnrecognizedCertificateReference) {
            [void]$unknowns.Add('PrimaryCertificateReferenceRequiresInvestigation')
        }
        $joinedTenant = [string](Get-ClientValue $Join 'TenantId')
        if ($joinedTenant -and $enrollment.TenantId -and $joinedTenant -ne $enrollment.TenantId) { [void]$issues.Add('PrimaryEnrollmentTenantMismatch') }
        if ($TaskRead) {
            $taskPath = "\Microsoft\Windows\EnterpriseMgmt\$($enrollment.EnrollmentId)\"
            $primaryTasks = @($Tasks | Where-Object { $_.Path -like "$taskPath*" })
            if ($primaryTasks.Count -eq 0) { [void]$observations.Add('PrimaryEnrollmentTasksNotFound') }
            elseif (@($primaryTasks | Where-Object { $_.State -ne 'Disabled' }).Count -eq 0) { [void]$observations.Add('PrimaryEnrollmentTasksAllDisabled') }
        }
    }
    if ([string](Get-ClientValue $Join 'DeviceAuthStatus') -like 'FAILED*') { [void]$observations.Add('EntraDeviceAuthStatusFailed:InvestigateDeviceIdentity') }
    if (-not $ServiceRead) { [void]$unknowns.Add('ServiceEvidenceUnavailable') }
    else {
        foreach ($name in @('Schedule', 'dmwappushservice')) {
            $service = $Services | Where-Object { $_.Name -eq $name } | Select-Object -First 1
            if (-not $service -or -not $service.Present) { [void]$issues.Add("ServiceMissing:$name") }
            elseif ($service.StartMode -eq 'Disabled') { [void]$issues.Add("ServiceDisabled:$name") }
        }
        $ime = $Services | Where-Object { $_.Name -eq 'IntuneManagementExtension' } | Select-Object -First 1
        if (-not $ime -or -not $ime.Present) { [void]$observations.Add('ImeAbsent:WhetherRequiredDependsOnAssignments') }
        elseif ($ime.StartMode -eq 'Disabled') { [void]$observations.Add('ImeDisabled:InvestigateSeparatelyFromMdm') }
        elseif ($ime.State -ne 'Running') { [void]$observations.Add('ImeNotRunning:InvestigateSeparatelyFromMdm') }
    }
    foreach ($enrollment in $linked) {
        if ($enrollment.Certificate.Status -ne 'Present') { [void]$observations.Add("LinkedCertificate:$($enrollment.EnrollmentId):$($enrollment.Certificate.Status)") }
        if ($primary.Count -eq 1 -and $enrollment.TenantId -and $primary[0].TenantId -and $enrollment.TenantId -ne $primary[0].TenantId) {
            [void]$observations.Add("LinkedEnrollmentTenantMismatch:$($enrollment.EnrollmentId)")
        }
    }
    [pscustomobject]@{
        CanSync = ($issues.Count -eq 0 -and $unknowns.Count -eq 0)
        Issues = @($issues.ToArray())
        Unknowns = @($unknowns.ToArray())
        Observations = @($observations.ToArray())
    }
}

function Invoke-IntuneClientAssessment {
    <#
    .SYNOPSIS
        Collects client evidence, evaluates prerequisites and optionally requests sync.
    .DESCRIPTION
        Rejects unsupported execution contexts before collection. Collectors run
        independently so one failure does not hide other evidence. Discovery
        probes are deduplicated by host/port and their results remain advisory.
        ShouldProcess gates task submission, not the diagnostic collection.
    .PARAMETER RequestSync
        Submit one task if the local assessment permits it. Default: false.
    .PARAMETER WaitSeconds
        Observe a newer task run for up to this many seconds. Default: 120.
    .OUTPUTS
        PSCustomObject containing the full report, Assessment, Sync, Verdict and
        ExitCode. Expected collector failures are retained in CollectionErrors.
        Does not serialize, write reports or exit; the script entry point does.
    .NOTES
        An attempted task submission's verdict takes precedence over diagnostic collector
        errors; EvidenceComplete and CollectionErrors must be checked separately.
        CloudLastSyncVerified and PolicyApplicationVerified always remain false.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [switch]$RequestSync,
        [ValidateRange(1, 600)][int]$WaitSeconds = 120
    )

    $errors = New-Object 'System.Collections.Generic.List[string]'
    $report = [ordered]@{
        SchemaVersion = 3
        ComputerName = $env:COMPUTERNAME
        CollectedUtc = [DateTime]::UtcNow.ToString('o')
        DeviceResponsive = $true
        RunningAsSystem = $false
        PowerShellVersion = [string]$PSVersionTable.PSVersion
        ExecutionContext = $null
        OperatingSystem = $null
        Join = $null
        Enrollments = @()
        Services = @()
        Tasks = @()
        WinHttpProxy = $null
        DiscoveryConnectivity = @()
        ImeLogs = @()
        RecentMdmWarningsAndErrors = @()
        Assessment = $null
        CollectionErrors = @()
        EvidenceComplete = $false
        Sync = [pscustomobject]@{ Status = 'NotRequested' }
        CloudLastSyncVerified = $false
        PolicyApplicationVerified = $false
        Verdict = 'AssessmentIncomplete'
        ExitCode = 2
    }
    Write-Verbose "Starting client assessment on $($report.ComputerName); sync requested: $([bool]$RequestSync)."
    $context = Get-ClientExecutionContext
    $report.RunningAsSystem = $context.IsSystem
    $report.ExecutionContext = $context
    Write-Verbose "Execution context: PowerShell $($report.PowerShellVersion) $($context.Edition); 64-bit=$($context.Is64Bit); elevated=$($context.IsAdministrator); SYSTEM=$($context.IsSystem)."
    if (-not $context.Is64Bit) {
        [void]$errors.Add('The PowerShell process is 32-bit. Run in 64-bit PowerShell.')
    }
    if (-not $context.IsAdministrator) {
        [void]$errors.Add('The process is not elevated. Open PowerShell with Run as administrator, or deploy as SYSTEM.')
    }
    if ($errors.Count -gt 0) {
        $report.CollectionErrors = @($errors.ToArray())
        if ($RequestSync) { $report.Sync = [pscustomobject]@{ Status = 'SkippedExecutionContext' } }
        foreach ($message in $errors) { Write-Verbose "Collection skipped: $message" }
        return [pscustomobject]$report
    }
    $enrollmentRead = $false
    $serviceRead = $false
    $taskRead = $false
    Write-Verbose 'Collecting operating system and last-boot information.'
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -OperationTimeoutSec 10 -ErrorAction Stop
        $report.OperatingSystem = [pscustomobject]@{
            Version     = $os.Version
            Build       = $os.BuildNumber
            LastBootUtc = $os.LastBootUpTime.ToUniversalTime().ToString('o')
        }
    }
    catch { [void]$errors.Add("OperatingSystem: $($_.Exception.Message)"); Write-Verbose "OperatingSystem collection failed: $($_.Exception.Message)" }
    Write-Verbose 'Collecting Entra join evidence using dsregcmd /status.'
    try { $report.Join = Get-ClientJoinEvidence }
    catch { [void]$errors.Add("Join: $($_.Exception.Message)"); Write-Verbose "Join collection failed: $($_.Exception.Message)" }
    Write-Verbose 'Collecting primary/linked MDM enrollments and referenced certificates.'
    try { $report.Enrollments = @(Get-ClientEnrollmentEvidence); $enrollmentRead = $true }
    catch { [void]$errors.Add("Enrollments: $($_.Exception.Message)"); Write-Verbose "Enrollment collection failed: $($_.Exception.Message)" }
    if ($enrollmentRead) { Write-Verbose "Enrollment records collected: $($report.Enrollments.Count)." }
    Write-Verbose 'Collecting MDM-related and IME service state; no services will be changed.'
    try { $report.Services = @(Get-ClientServiceEvidence); $serviceRead = $true }
    catch { [void]$errors.Add("Services: $($_.Exception.Message)"); Write-Verbose "Service collection failed: $($_.Exception.Message)" }
    Write-Verbose 'Collecting EnterpriseMgmt task state and scheduler history.'
    try { $report.Tasks = @(Get-ClientTaskEvidence); $taskRead = $true }
    catch { [void]$errors.Add("Tasks: $($_.Exception.Message)"); Write-Verbose "Task collection failed: $($_.Exception.Message)" }
    if ($taskRead) { Write-Verbose "EnterpriseMgmt tasks collected: $($report.Tasks.Count). Task activity does not prove MDM sync success." }
    Write-Verbose 'Reading WinHTTP proxy configuration.'
    try { $report.WinHttpProxy = Invoke-ClientReadCommand -FileName "$env:SystemRoot\System32\netsh.exe" -Arguments 'winhttp show proxy' }
    catch { [void]$errors.Add("WinHttpProxy: $($_.Exception.Message)"); Write-Verbose "Proxy collection failed: $($_.Exception.Message)" }
    Write-Verbose 'Reading IME log existence and last-write times.'
    try { $report.ImeLogs = @(Get-ClientImeLogEvidence) }
    catch { [void]$errors.Add("ImeLogs: $($_.Exception.Message)"); Write-Verbose "IME log collection failed: $($_.Exception.Message)" }
    $urls = @($report.Enrollments | Select-Object -ExpandProperty DiscoveryUrl) + @([string](Get-ClientValue $report.Join 'MdmUrl'))
    $targets = @{}
    foreach ($url in $urls) {
        $uri = $null
        if ([uri]::TryCreate($url, [UriKind]::Absolute, [ref]$uri) -and $uri.Scheme -eq 'https') {
            $targets[$uri.Authority] = $uri
        }
    }
    Write-Verbose "Probing $($targets.Count) discovery host(s) using direct DNS/TCP; results do not validate TLS or the proxy path."
    $report.DiscoveryConnectivity = @(foreach ($uri in $targets.Values) {
        Write-Verbose "Probing $($uri.DnsSafeHost):$($uri.Port)."
        $probe = Test-ClientDiscoveryEndpoint -Uri $uri
        Write-Verbose "Discovery probe result: DNS=$($probe.Dns); TCP=$($probe.Tcp); Error=$($probe.Error)"
        $probe
    })
    Write-Verbose 'Evaluating local sync prerequisites.'
    $report.Assessment = Get-ClientSyncAssessment -Enrollments $report.Enrollments -Services $report.Services -Join $report.Join -EnrollmentRead $enrollmentRead -ServiceRead $serviceRead -Tasks $report.Tasks -TaskRead $taskRead
    Write-Verbose "Assessment: CanSync=$($report.Assessment.CanSync); issues=$($report.Assessment.Issues.Count); unknowns=$($report.Assessment.Unknowns.Count); observations=$($report.Assessment.Observations.Count)."
    foreach ($issue in $report.Assessment.Issues) { Write-Verbose "Prerequisite issue: $issue" }
    foreach ($unknown in $report.Assessment.Unknowns) { Write-Verbose "Unresolved prerequisite: $unknown" }
    foreach ($observation in $report.Assessment.Observations) { Write-Verbose "Observation: $observation" }
    $report.Verdict = 'AssessmentComplete'
    $report.ExitCode = 0
    if ($errors.Count -gt 0 -or $report.Assessment.Unknowns.Count -gt 0) { $report.Verdict = 'AssessmentIncomplete'; $report.ExitCode = 2 }
    if ($report.Assessment.Issues.Count -gt 0) { $report.Verdict = 'LocalPrerequisiteIssue'; $report.ExitCode = 1 }
    if ($RequestSync) {
        $report.Sync = [pscustomobject]@{ Status = 'SkippedPrerequisites' }
        if ($report.Assessment.CanSync) {
            if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, 'Submit the primary Intune enrollment PushLaunch task once')) {
                Write-Verbose 'Sync approved. Validating the enrollment task; no automatic retries.'
                $report.Sync = Invoke-ClientMdmTask -Enrollments $report.Enrollments -WaitSeconds $WaitSeconds
                Write-Verbose "Sync outcome: $($report.Sync.Status)."
                if ($report.Sync.Error) { Write-Verbose "Task failure at $($report.Sync.Stage): $($report.Sync.Error)" }
                switch ($report.Sync.Status) {
                    'Completed' { $report.Verdict = 'MdmTaskCompleted'; $report.ExitCode = 0 }
                    'TimedOut' { $report.Verdict = 'MdmTaskTimedOut'; $report.ExitCode = 2 }
                    'AlreadyRunning' { $report.Verdict = 'MdmTaskAlreadyRunning'; $report.ExitCode = 2 }
                    default { $report.Verdict = 'MdmTaskFailed'; $report.ExitCode = 1 }
                }
            }
            else {
                $report.Sync = [pscustomobject]@{ Status = 'NotRequestedByShouldProcess' }
                Write-Verbose 'Sync skipped by WhatIf or declined confirmation.'
            }
        }
        else { Write-Verbose 'Sync skipped because local prerequisites were not met or could not be established.' }
    }
    else { Write-Verbose 'Diagnostic-only run: no MDM sync requested.' }
    Write-Verbose 'Reading up to 12 MDM Admin warnings/errors from the last three days; these are not task-correlated.'
    try { $report.RecentMdmWarningsAndErrors = @(Get-ClientMdmEvents -Since ([DateTime]::Now.AddDays(-3))) }
    catch {
        [void]$errors.Add("MdmEvents: $($_.Exception.Message)")
        Write-Verbose "MDM event collection failed: $($_.Exception.Message)"
        if ($report.Verdict -eq 'AssessmentComplete') { $report.Verdict = 'AssessmentIncomplete'; $report.ExitCode = 2 }
    }
    $report.CollectionErrors = @($errors.ToArray())
    $report.EvidenceComplete = ($errors.Count -eq 0 -and $report.Assessment.Unknowns.Count -eq 0)
    Write-Verbose "Assessment finished: $($report.Verdict); exit code=$($report.ExitCode); collection errors=$($errors.Count)."
    [pscustomobject]$report
}

#endregion Assessment

#region Entry point

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
try {
    $script:ClientSyncLogWarningWritten = $false
    $script:ClientSyncLogPath = Join-Path $env:ProgramData 'Microsoft\IntuneManagementExtension\Logs\Invoke-IntuneClientSync.log'
    Write-ClientSyncLog "Starting client script; sync requested=$([bool]$Sync); output format=$OutputFormat."
    $report = Invoke-IntuneClientAssessment -RequestSync:$Sync -WaitSeconds $TimeoutSeconds -Verbose 4>&1 | ForEach-Object {
        if ($_ -is [System.Management.Automation.VerboseRecord]) {
            Write-ClientSyncLog $_.Message
            Write-Verbose $_.Message
        }
        else { $_ }
    }
    Write-ClientSyncLog $report
    $fullPath = $null
    if ($OutputPath) {
        $fullPath = [System.IO.Path]::GetFullPath($OutputPath)
        Write-ClientSyncLog "Writing full JSON report to $fullPath."
        Write-Verbose "Writing full JSON report to $fullPath."
        [void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($fullPath))
        [System.IO.File]::WriteAllText($fullPath, ($report | ConvertTo-Json -Depth 10), (New-Object System.Text.UTF8Encoding($true)))
        Write-ClientSyncLog 'Full JSON report written.'
        Write-Verbose 'Full JSON report written.'
    }
    Write-ClientSyncLog "Emitting $OutputFormat output; exit code=$($report.ExitCode)."
    Write-Verbose "Emitting $OutputFormat output; exit code=$($report.ExitCode)."
    if ($OutputFormat -eq 'Object') {
        $report
    }
    elseif ($OutputFormat -eq 'DetailedJson') {
        $report | ConvertTo-Json -Depth 10
    }
    else {
        [pscustomobject]@{
            ComputerName = $report.ComputerName
            CollectedUtc = $report.CollectedUtc
            DeviceResponsive = $report.DeviceResponsive
            Verdict = $report.Verdict
            ExitCode = $report.ExitCode
            PowerShellVersion = $report.PowerShellVersion
            ExecutionContext = $report.ExecutionContext
            EntraDeviceId = Get-ClientValue $report.Join 'DeviceId'
            Assessment = $report.Assessment
            Sync = $report.Sync
            CollectionErrorCount = $report.CollectionErrors.Count
            CollectionErrors = @($report.CollectionErrors | Select-Object -First 3 | ForEach-Object { $_.Substring(0, [Math]::Min(200, $_.Length)) })
            EvidenceComplete = $report.EvidenceComplete
            DiscoveryProbeWarningCount = @($report.DiscoveryConnectivity | Where-Object { $_.Error }).Count
            CloudLastSyncVerified = $report.CloudLastSyncVerified
            PolicyApplicationVerified = $report.PolicyApplicationVerified
            ReportPath = $fullPath
        } | ConvertTo-Json -Depth 8 -Compress
    }
    exit $report.ExitCode
}
catch {
    Write-Verbose "Script failed: $($_.Exception.Message)"
    $errorResult = [pscustomobject]@{
        ComputerName = $env:COMPUTERNAME
        CollectedUtc = [DateTime]::UtcNow.ToString('o')
        Verdict      = 'ScriptError'
        ExitCode     = 2
        Error        = $_.Exception.Message
    }
    Write-ClientSyncLog $errorResult
    if ($OutputFormat -eq 'Object') { $errorResult }
    else { $errorResult | ConvertTo-Json -Compress:($OutputFormat -eq 'SummaryJson') }
    exit 2
}

#endregion Entry point