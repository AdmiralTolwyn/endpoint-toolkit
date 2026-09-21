<#
.SYNOPSIS
    Repairs a disabled dmwappushservice and submits one enrollment-specific sync task.

.DESCRIPTION
    Implements Microsoft's fix for Intune clients with dmwappushservice disabled.
    Rechecks the service before making a change and verifies the startup mode
    afterwards. Manual and Automatic startup modes are unchanged.
    Once the service is confirmed enabled, selects the unique primary Intune
    enrollment with an OMA-DM account and submits its existing PushLaunch task.
    Validates the task's enrollment-specific action and SYSTEM principal before
    starting it. Missing, disabled or ambiguous targets fail without guessing.
    A direct run also submits the task if the service was already Auto/Manual.
    Does not start/restart services, restart IME, change enrollment, clear policy
    caches, create/modify tasks or retry. Observes a newer task run for up to
    TimeoutSeconds. Task completion does not prove sync completed, policies
    applied or the Intune console timestamp updated.

.PARAMETER TimeoutSeconds
    Task observation deadline, 1-600 seconds, default 120. A timeout never stops
    the task or retries. Individual scheduler calls are not bounded by this timer.

.INPUTS
    None.

.OUTPUTS
    System.String. Separate service and sync status messages, including the
    target enrollment/task and an error/HRESULT on failure.

.EXAMPLE
    .\Remediate-IntuneMdmSyncService.ps1

    Changes Disabled to Automatic when needed, verifies it, and submits PushLaunch.

.NOTES
    Run on Windows as SYSTEM or an elevated administrator in 64-bit PowerShell
    5.1 or 7. Uses ScheduledTasks directly; no WinRT or child-process bridge.
    Exit 0: service enabled and a newer task run finished with result zero.
    Exit 1: service/task operation failed or the observed task result was nonzero.
    Exit 2: task already running/queued, or observation timed out.
    Pair with Detect-IntuneMdmSyncService.ps1. No required parameters or dependencies.
    Intune invokes this only when detection finds Disabled. Each direct run can
    submit a task; one-time means at most once per invocation, not once forever.
    Do not repeatedly rerun. Service repair is not rolled back if task submission
    fails. The built-in task layout is Windows implementation detail; validate
    this targeting on your supported Windows builds before fleet deployment.
    Investigate any policy or script that disables the service again. Confirm
    policy application and console reporting separately after remediation.
    Appends UTC-stamped progress, task results and exit codes to
    C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\Remediate-IntuneMdmSyncService.log.
    Creates the folder if missing. Log-write failures warn without changing exit codes.

.LINK
    https://learn.microsoft.com/en-us/troubleshoot/mem/intune/device-management/cannot-sync-windows-10-devices

.LINK
    https://learn.microsoft.com/en-us/powershell/module/scheduledtasks/start-scheduledtask
#>
#Requires -Version 5.1

param([ValidateRange(1, 600)][int]$TimeoutSeconds = 120)

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

function Get-ServiceMdmEnrollments {
    <#
    .SYNOPSIS
        Reads primary Intune enrollments and their OMA-DM account presence.
    .OUTPUTS
        Enrollment objects for MS DM Server only. Registry access failures throw,
        rather than returning misleading missing-enrollment evidence. Registry
        footprint alone is not proof of a healthy enrollment or successful sync.
    #>
    $root = 'HKLM:\SOFTWARE\Microsoft\Enrollments'
    if (-not (Test-Path -LiteralPath $root -ErrorAction Stop)) { return }
    foreach ($key in @(Get-ChildItem -LiteralPath $root -ErrorAction Stop)) {
        if ($key.PSChildName -notmatch '^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$') { continue }
        $values = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction Stop
        if (-not $values.PSObject.Properties['ProviderID'] -or $values.ProviderID -ne 'MS DM Server') { continue }
        $tenantId = $null
        if ($values.PSObject.Properties['AADTenantID']) { $tenantId = $values.AADTenantID }
        [pscustomobject]@{
            EnrollmentId = $key.PSChildName
            ProviderId = $values.ProviderID
            TenantId = $tenantId
            OmaDmAccountExists = Test-Path -LiteralPath "HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\$($key.PSChildName)" -ErrorAction Stop
        }
    }
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

$ErrorActionPreference = 'Stop'
$serviceName = 'dmwappushservice'
$stage = 'Service remediation'
$script:ClientSyncLogPath = Join-Path $env:ProgramData 'Microsoft\IntuneManagementExtension\Logs\Remediate-IntuneMdmSyncService.log'
$script:ClientSyncLogWarningWritten = $false
Write-ClientSyncLog "Starting service remediation; task observation deadline=$TimeoutSeconds seconds."

try {
    $service = Get-CimInstance -ClassName Win32_Service -Filter "Name='$serviceName'" -OperationTimeoutSec 10 -ErrorAction Stop
    if ($null -eq $service) { throw "Service '$serviceName' was not found." }

    if ($service.StartMode -in @('Auto', 'Manual')) {
        Write-ClientSyncLog "$serviceName is not disabled (startup: $($service.StartMode)). No change needed." -OutputMessage
    }
    elseif ($service.StartMode -eq 'Disabled') {
        Set-Service -Name $serviceName -StartupType Automatic -ErrorAction Stop
        $service = Get-CimInstance -ClassName Win32_Service -Filter "Name='$serviceName'" -OperationTimeoutSec 10 -ErrorAction Stop
        if ($null -eq $service -or $service.StartMode -ne 'Auto') {
            throw "Startup mode for '$serviceName' could not be verified as Automatic."
        }

        Write-ClientSyncLog "$serviceName changed from Disabled to Automatic; verified." -OutputMessage
    }
    else {
        throw "Unexpected startup mode '$($service.StartMode)' for '$serviceName'."
    }

    $stage = 'MDM task operation'
    Write-ClientSyncLog 'Selecting and submitting the primary enrollment task, then observing its result.'
    $enrollments = @(Get-ServiceMdmEnrollments)
    $sync = Invoke-ClientMdmTask -Enrollments $enrollments -WaitSeconds $TimeoutSeconds
    Write-ClientSyncLog $sync
    Write-ClientSyncLog "MDM sync stage: $($sync.Stage)" -OutputMessage
    if ($sync.TaskPath) { Write-ClientSyncLog "MDM task: $($sync.TaskPath)$($sync.TaskName)" -OutputMessage }
    if ($sync.SubmissionAccepted) { Write-ClientSyncLog 'PushLaunch task submitted once.' -OutputMessage }
    if ($null -ne $sync.LastTaskResult) { Write-ClientSyncLog "Observed task result: $($sync.LastTaskResultHex) ($($sync.LastTaskResult)); state: $($sync.ObservedTaskState)." -OutputMessage }
    if ($sync.Status -eq 'AlreadyRunning') {
        Write-ClientSyncLog 'PushLaunch already running/queued; no new task submitted. Service configuration remains enabled.' -OutputMessage
        Write-ClientSyncLog 'Finished remediation; exit code=2.'
        exit 2
    }
    if ($sync.Status -eq 'TimedOut') {
        Write-ClientSyncLog "MDM task observation timed out: $($sync.Error)" -OutputMessage
        Write-ClientSyncLog 'Service configuration remains enabled. MDM sync remains unverified.' -OutputMessage
        Write-ClientSyncLog 'Finished remediation; exit code=2.'
        exit 2
    }
    if ($sync.Status -ne 'Completed') {
        Write-ClientSyncLog ("MDM task operation failed: {0}" -f $sync.Error) -OutputMessage
        if ($sync.HResult) { Write-ClientSyncLog "Operation HRESULT: $($sync.HResult)" -OutputMessage }
        Write-ClientSyncLog 'Service configuration remains enabled. No rollback or retry was attempted.' -OutputMessage
        Write-ClientSyncLog 'Finished remediation; exit code=1.'
        exit 1
    }

    Write-ClientSyncLog 'A newer PushLaunch run finished with task result 0x00000000. MDM sync, policy application and Intune console reporting are not verified.' -OutputMessage
    Write-ClientSyncLog 'Finished remediation; exit code=0.'
    exit 0
}
catch {
    $failure = $_.Exception.GetBaseException()
    Write-ClientSyncLog ('{0} failed: {1} (HRESULT 0x{2:X8})' -f $stage, $failure.Message, $failure.HResult) -OutputMessage
    if ($stage -eq 'MDM task operation') { Write-ClientSyncLog 'Service configuration remains enabled. No rollback or retry was attempted.' -OutputMessage }
    Write-ClientSyncLog 'Finished remediation; exit code=1.'
    exit 1
}