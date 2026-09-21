<#
.SYNOPSIS
    Tests service repair and enrollment-specific task submission without device changes.
.DESCRIPTION
    Runs the actual detection and remediation files in isolated runspaces with
    mocked service, registry and ScheduledTasks operations. Checks the task helper
    matches the diagnostic script, and rejects old WinRT/child-process code.
    Both scripts must have one BOM and a comment as the first parser token.
.EXAMPLE
    .\Test-IntuneMdmSyncService.ps1
.NOTES
    Run with PowerShell 5.1 or 7 on Windows. No elevation or enrollment required.
    No real registry/service/task access, child workers, sync or network probes.
#>
#Requires -Version 5.1

$ErrorActionPreference = 'Stop'

$detector = Join-Path $PSScriptRoot 'Detect-IntuneMdmSyncService.ps1'
$detectionTokens = $null
$detectionErrors = $null
$null = [System.Management.Automation.Language.Parser]::ParseFile($detector, [ref]$detectionTokens, [ref]$detectionErrors)
if ($detectionErrors.Count) { throw ($detectionErrors | Out-String) }
$detectionBytes = [System.IO.File]::ReadAllBytes($detector)
if ([BitConverter]::ToString($detectionBytes, 0, 5) -ne 'EF-BB-BF-3C-23' -or $detectionTokens[0].Kind -ne 'Comment') {
    throw 'Detection must begin with one UTF-8 BOM and a recognized help comment.'
}
$detectionHelp = Get-Help $detector -Full
if (-not $detectionHelp.examples -or -not $detectionHelp.returnValues) { throw 'Detection help missing.' }

$source = Join-Path $PSScriptRoot 'Remediate-IntuneMdmSyncService.ps1'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($source, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }
$bytes = [System.IO.File]::ReadAllBytes($source)
if ([BitConverter]::ToString($bytes, 0, 5) -ne 'EF-BB-BF-3C-23' -or $tokens[0].Kind -ne 'Comment') {
    throw 'Remediation must begin with one UTF-8 BOM and a recognized help comment.'
}
$help = Get-Help $source -Full
if (-not $help.examples -or -not $help.returnValues) { throw 'Script help missing.' }
$taskHelper = $ast.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -eq 'Invoke-ClientMdmTask' }
$diagnosticPath = Join-Path $PSScriptRoot 'Invoke-IntuneClientSync.ps1'
$diagnosticAst = [Management.Automation.Language.Parser]::ParseFile($diagnosticPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }
$diagnosticHelper = $diagnosticAst.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -eq 'Invoke-ClientMdmTask' }
if (-not $taskHelper -or $taskHelper.Extent.Text -cne $diagnosticHelper.Extent.Text) { throw 'Standalone task helper copies differ.' }
foreach ($scriptAst in @($ast, $diagnosticAst)) {
    if ($scriptAst.Extent.Text -match 'TryCreateSession|StartAsync|EncodedCommand|Invoke-MdmWindowsPowerShell') { throw 'Obsolete WinRT bridge remains.' }
}

$detectionCases = @(
    @{ Mode = 'Disabled'; Exit = 1; Text = 'Remediation is required' }
    @{ Mode = 'Auto'; Exit = 0; Text = 'is not disabled' }
    @{ Mode = 'Manual'; Exit = 0; Text = 'is not disabled' }
    @{ Mode = 'Missing'; Exit = 2; Text = 'Detection error' }
    @{ Mode = 'Denied'; Exit = 2; Text = 'Detection error' }
    @{ Mode = 'Unexpected'; Exit = 2; Text = 'Detection error' }
)

$cases = @(
    @{ Mode = 'Disabled'; Task = 'Ready'; Exit = 0; Writes = 1; Starts = 1; Text = 'task submitted once' }
    @{ Mode = 'Auto'; Task = 'Ready'; Exit = 0; Writes = 0; Starts = 1; Text = 'task submitted once' }
    @{ Mode = 'Manual'; Task = 'Ready'; Exit = 0; Writes = 0; Starts = 1; Text = 'task submitted once' }
    @{ Mode = 'Missing'; Task = 'Ready'; Exit = 1; Writes = 0; Starts = 0; Text = 'Service remediation failed' }
    @{ Mode = 'Denied'; Task = 'Ready'; Exit = 1; Writes = 0; Starts = 0; Text = 'Service remediation failed' }
    @{ Mode = 'Unexpected'; Task = 'Ready'; Exit = 1; Writes = 0; Starts = 0; Text = 'Service remediation failed' }
    @{ Mode = 'SetDenied'; Task = 'Ready'; Exit = 1; Writes = 1; Starts = 0; Text = 'Service remediation failed' }
    @{ Mode = 'VerifyFailed'; Task = 'Ready'; Exit = 1; Writes = 1; Starts = 0; Text = 'could not be verified' }
    @{ Mode = 'VerifyMissing'; Task = 'Ready'; Exit = 1; Writes = 1; Starts = 0; Text = 'could not be verified' }
    @{ Mode = 'VerifyThrows'; Task = 'Ready'; Exit = 1; Writes = 1; Starts = 0; Text = 'Service remediation failed' }
    @{ Mode = 'Disabled'; Task = 'SubmitThrows'; Exit = 1; Writes = 1; Starts = 1; Text = 'fixture submission denied' }
    @{ Mode = 'Disabled'; Task = 'Ready'; Exit = 2; Writes = 1; Starts = 1; Text = 'already running/queued'; Repeat = $true }
)
foreach ($taskMode in @('NoEnrollment', 'RegistryDenied', 'RegistryReadDenied', 'NoAccount', 'LinkedOnly', 'MultiplePrimary', 'MissingTask', 'MultipleTasks', 'TaskDenied', 'WrongPath', 'WrongName', 'Disabled', 'DisabledSetting', 'NoDemand', 'WrongPrincipal', 'WrongExecutable', 'WrongGuid', 'WrongArguments', 'ExtraAction', 'UnknownState')) {
    $cases += @{ Mode = 'Disabled'; Task = $taskMode; Exit = 1; Writes = 1; Starts = 0; Text = 'failed' }
}
foreach ($taskMode in @('Running', 'Queued')) {
    $cases += @{ Mode = 'Auto'; Task = $taskMode; Exit = 2; Writes = 0; Starts = 0; Text = 'already running/queued' }
}
foreach ($taskMode in @('SystemSid', 'QualifiedSystem', 'UnquotedGuid', 'FullExecutable', 'LinkedPresent', 'ResiduePresent')) {
    $cases += @{ Mode = 'Manual'; Task = $taskMode; Exit = 0; Writes = 0; Starts = 1; Text = 'task submitted once' }
}
$cases += @(
    @{ Mode = 'Disabled'; Task = 'NonzeroResult'; Exit = 1; Writes = 1; Starts = 1; Text = '0x82AA000B' }
    @{ Mode = 'Disabled'; Task = 'StaleZero'; Exit = 2; Writes = 1; Starts = 1; Text = 'observation timed out' }
    @{ Mode = 'Disabled'; Task = 'BaselineDenied'; Exit = 1; Writes = 1; Starts = 0; Text = 'fixture baseline denied' }
    @{ Mode = 'Disabled'; Task = 'ObservationDenied'; Exit = 1; Writes = 1; Starts = 1; Text = 'fixture observation denied' }
)
$observationCases = @(
    @{ Task = 'Ready'; Status = 'Completed'; NewRun = $true; ResultHex = '0x00000000'; Starts = 1 }
    @{ Task = 'Transition'; Status = 'Completed'; NewRun = $true; ResultHex = '0x00000000'; Starts = 1 }
    @{ Task = 'QueuedTransition'; Status = 'Completed'; NewRun = $true; ResultHex = '0x00000000'; Starts = 1 }
    @{ Task = 'NonzeroResult'; Status = 'Failed'; NewRun = $true; ResultHex = '0x82AA000B'; Starts = 1 }
    @{ Task = 'SignedResult'; Status = 'Failed'; NewRun = $true; ResultHex = '0x82AA000B'; Starts = 1 }
    @{ Task = 'StaleZero'; Status = 'TimedOut'; NewRun = $false; Starts = 1 }
    @{ Task = 'StaleFailure'; Status = 'TimedOut'; NewRun = $false; Starts = 1 }
    @{ Task = 'ClockRollback'; Status = 'TimedOut'; NewRun = $false; Starts = 1 }
    @{ Task = 'PersistentRunning'; Status = 'TimedOut'; NewRun = $true; ResultHex = '0x00000000'; Starts = 1 }
    @{ Task = 'PersistentQueued'; Status = 'TimedOut'; NewRun = $true; ResultHex = '0x00000000'; Starts = 1 }
    @{ Task = 'RunningCode'; Status = 'TimedOut'; NewRun = $true; ResultHex = '0x00041301'; Starts = 1 }
    @{ Task = 'NotRunCode'; Status = 'TimedOut'; NewRun = $true; ResultHex = '0x00041303'; Starts = 1 }
    @{ Task = 'QueuedCode'; Status = 'TimedOut'; NewRun = $true; ResultHex = '0x00041325'; Starts = 1 }
    @{ Task = 'BaselineDenied'; Status = 'Failed'; NewRun = $false; Starts = 0 }
    @{ Task = 'BaselineMissing'; Status = 'Failed'; NewRun = $false; Starts = 0 }
    @{ Task = 'ObservationDenied'; Status = 'Failed'; NewRun = $false; Starts = 1 }
    @{ Task = 'Disappeared'; Status = 'Failed'; NewRun = $false; Starts = 1 }
    @{ Task = 'MissingInfo'; Status = 'Failed'; NewRun = $false; Starts = 1 }
)

$runnerBody = {
    param($Path, $Case, $HelperText)
    $global:MdmFixTestState = @{ Mode = $Case.Mode; Task = $Case.Task; Writes = 0; Starts = 0; Queries = 0; TaskQueries = 0 }
    $global:MdmFixPrimaryId = '11111111-1111-1111-1111-111111111111'
    function Get-CimInstance {
        <# .SYNOPSIS
            Returns fixture service evidence without accessing CIM.
        #>
        param($ClassName, $Filter, $OperationTimeoutSec, $ErrorAction)
        if ($ClassName -ne 'Win32_Service' -or $Filter -ne "Name='dmwappushservice'" -or $OperationTimeoutSec -ne 10) { throw 'Unexpected query.' }
        $global:MdmFixTestState.Queries++
        $mode = $global:MdmFixTestState.Mode
        if ($mode -eq 'Missing') { return }
        if ($mode -eq 'Denied') { throw 'Fixture query denied.' }
        if ($global:MdmFixTestState.Writes -gt 0) {
            if ($mode -eq 'VerifyMissing') { return }
            if ($mode -eq 'VerifyThrows') { throw 'Fixture read-back error.' }
            if ($mode -ne 'VerifyFailed') { $mode = 'Auto' }
        }
        if ($mode -like 'Verify*' -or $mode -eq 'SetDenied') { $mode = 'Disabled' }
        [pscustomobject]@{ StartMode = $mode; State = 'Stopped' }
    }
    function Set-Service {
        <# .SYNOPSIS
            Records the requested startup mode without changing any service.
        #>
        param($Name, $StartupType, $ErrorAction)
        if ($Name -ne 'dmwappushservice' -or $StartupType -ne 'Automatic') { throw 'Unexpected change.' }
        $global:MdmFixTestState.Writes++
        if ($global:MdmFixTestState.Mode -eq 'SetDenied') { throw 'Fixture set denied.' }
    }
    function Test-Path {
        <# .SYNOPSIS
            Returns fixture registry existence; never queries the device.
        #>
        param($LiteralPath, $ErrorAction)
        $mode = $global:MdmFixTestState.Task
        if ($mode -eq 'RegistryDenied') { throw 'fixture registry access denied' }
        if ($LiteralPath -eq 'HKLM:\SOFTWARE\Microsoft\Enrollments') { return ($mode -ne 'NoEnrollment') }
        if ($LiteralPath -like 'HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\*') {
            if ($mode -eq 'NoAccount') { return $false }
            if ($mode -eq 'ResiduePresent' -and $LiteralPath -notlike "*\$global:MdmFixPrimaryId") { return $false }
            return $true
        }
        throw "Unexpected registry path: $LiteralPath"
    }
    function Get-ChildItem {
        <# .SYNOPSIS
            Supplies primary, linked and residual enrollment keys for each case.
        #>
        param($LiteralPath, $ErrorAction)
        if ($LiteralPath -ne 'HKLM:\SOFTWARE\Microsoft\Enrollments') { throw 'Unexpected registry enumeration.' }
        [pscustomobject]@{ PSChildName = $global:MdmFixPrimaryId; PSPath = 'fixture-primary' }
        if ($global:MdmFixTestState.Task -in @('LinkedPresent', 'MultiplePrimary', 'ResiduePresent')) {
            [pscustomobject]@{ PSChildName = '22222222-2222-2222-2222-222222222222'; PSPath = 'fixture-secondary' }
        }
        [pscustomobject]@{ PSChildName = 'Status'; PSPath = 'fixture-ignore' }
    }
    function Get-ItemProperty {
        <# .SYNOPSIS
            Supplies enrollment provider and tenant fields without registry access.
        #>
        param($LiteralPath, $ErrorAction)
        $mode = $global:MdmFixTestState.Task
        if ($mode -eq 'RegistryReadDenied') { throw 'fixture registry value read denied' }
        if ($LiteralPath -notin @('fixture-primary', 'fixture-secondary')) { throw 'Non-enrollment key was read.' }
        $provider = 'MS DM Server'
        if ($mode -eq 'LinkedOnly' -or ($mode -eq 'LinkedPresent' -and $LiteralPath -eq 'fixture-secondary')) { $provider = 'Microsoft Device Management' }
        [pscustomobject]@{ ProviderID = $provider; AADTenantID = 'fixture-tenant' }
    }
    function Get-ScheduledTask {
        <# .SYNOPSIS
            Supplies only the selected task and simulates invalid/busy definitions.
        #>
        param($TaskPath, $TaskName, $ErrorAction)
        $mode = $global:MdmFixTestState.Task
        $global:MdmFixTestState.TaskQueries++
        if ($TaskPath -ne "\Microsoft\Windows\EnterpriseMgmt\$global:MdmFixPrimaryId\" -or $TaskName -ne 'PushLaunch') { throw 'Wrong enrollment/task selected.' }
        if ($mode -eq 'MissingTask') { return }
        if ($mode -eq 'TaskDenied') { throw 'fixture task query denied' }
        $task = [pscustomobject]@{
            TaskPath = $TaskPath; TaskName = $TaskName; State = 'Ready'
            Settings = [pscustomobject]@{ Enabled = $true; AllowDemandStart = $true }
            Principal = [pscustomobject]@{ UserId = 'SYSTEM' }
            Actions = @([pscustomobject]@{ Execute = '%windir%\system32\deviceenroller.exe'; Arguments = "/o `"$global:MdmFixPrimaryId`" /c /z" })
        }
        if ($global:MdmFixTestState.Starts -gt 0) {
            if ($mode -eq 'Disappeared') { return }
            if ($mode -eq 'PersistentRunning' -or ($mode -eq 'Transition' -and $global:MdmFixTestState.TaskQueries -le 3)) { $task.State = 'Running' }
            if ($mode -eq 'PersistentQueued' -or ($mode -eq 'QueuedTransition' -and $global:MdmFixTestState.TaskQueries -le 3)) { $task.State = 'Queued' }
        }
        switch ($mode) {
            'WrongPath' { $task.TaskPath = '\Microsoft\Windows\EnterpriseMgmt\wrong\' }
            'WrongName' { $task.TaskName = 'PushRenewal' }
            'Disabled' { $task.State = 'Disabled' }
            'DisabledSetting' { $task.Settings.Enabled = $false }
            'NoDemand' { $task.Settings.AllowDemandStart = $false }
            'WrongPrincipal' { $task.Principal.UserId = 'someone-else' }
            'SystemSid' { $task.Principal.UserId = 'S-1-5-18' }
            'QualifiedSystem' { $task.Principal.UserId = 'NT AUTHORITY\SYSTEM' }
            'WrongExecutable' { $task.Actions[0].Execute = 'C:\Other\deviceenroller.exe' }
            'FullExecutable' { $task.Actions[0].Execute = Join-Path $env:SystemRoot 'System32\deviceenroller.exe' }
            'WrongGuid' { $task.Actions[0].Arguments = '/o "22222222-2222-2222-2222-222222222222" /c /z' }
            'WrongArguments' { $task.Actions[0].Arguments = "/o $global:MdmFixPrimaryId /c /r" }
            'UnquotedGuid' { $task.Actions[0].Arguments = "/o $global:MdmFixPrimaryId /c /z" }
            'ExtraAction' { $task.Actions += $task.Actions[0] }
            'UnknownState' { $task.State = 'Unknown' }
            'Running' { $task.State = 'Running' }
            'Queued' { $task.State = 'Queued' }
            'MultipleTasks' { $task }
        }
        $task
    }
    function Get-ScheduledTaskInfo {
        param($InputObject, $ErrorAction)
        if ($InputObject.TaskName -ne 'PushLaunch' -or $InputObject.TaskPath -ne "\Microsoft\Windows\EnterpriseMgmt\$global:MdmFixPrimaryId\") { throw 'Wrong task history requested.' }
        $mode = $global:MdmFixTestState.Task
        $when = [datetime]'2026-09-20T12:00:00'
        $code = 0
        if ($global:MdmFixTestState.Starts -eq 0) {
            if ($mode -eq 'BaselineDenied') { throw 'fixture baseline denied' }
            if ($mode -eq 'BaselineMissing') { return }
        }
        else {
            if ($mode -eq 'ObservationDenied') { throw 'fixture observation denied' }
            if ($mode -eq 'MissingInfo') { return [pscustomobject]@{ LastRunTime = $null; LastTaskResult = $null } }
            if ($mode -notin @('StaleZero','StaleFailure')) { $when = $when.AddMinutes(1) }
            switch ($mode) {
                'ClockRollback' { $when = $when.AddHours(-1) }
                'NonzeroResult' { $code = 2192179211L }
                'SignedResult' { $code = -2102788085L }
                'StaleFailure' { $code = 2192179211L }
                'RunningCode' { $code = 0x00041301 }
                'NotRunCode' { $code = 0x00041303 }
                'QueuedCode' { $code = 0x00041325 }
                'Transition' { if ($global:MdmFixTestState.TaskQueries -le 3) { $code = 0x00041301 } }
                'QueuedTransition' { if ($global:MdmFixTestState.TaskQueries -le 3) { $code = 0x00041325 } }
            }
        }
        [pscustomobject]@{ LastRunTime = $when; LastTaskResult = $code }
    }
    function Start-ScheduledTask {
        <# .SYNOPSIS
            Records a submission to the exact validated task; never starts a task.
        #>
        param($InputObject, $ErrorAction)
        if ($InputObject.TaskName -ne 'PushLaunch' -or $InputObject.TaskPath -ne "\Microsoft\Windows\EnterpriseMgmt\$global:MdmFixPrimaryId\") { throw 'Wrong task submitted.' }
        $global:MdmFixTestState.Starts++
        if ($global:MdmFixTestState.Task -eq 'SubmitThrows') { throw 'fixture submission denied' }
    }
    $observed = $null
    if ($HelperText) {
        . ([scriptblock]::Create($HelperText))
        $enrollment = [pscustomobject]@{ EnrollmentId = $global:MdmFixPrimaryId; ProviderId = 'MS DM Server'; TenantId = 'fixture'; OmaDmAccountExists = $true }
        $observed = Invoke-ClientMdmTask -Enrollments @($enrollment) -WaitSeconds 1
    }
    elseif ($Case.Task) {
        & $Path -TimeoutSeconds 1
        if ($Case.Repeat) { $global:MdmFixTestState.Task = 'Running'; & $Path -TimeoutSeconds 1 }
    }
    else { & $Path }
    [pscustomobject]@{ ExitCode = $LASTEXITCODE; Writes = $global:MdmFixTestState.Writes; Starts = $global:MdmFixTestState.Starts; TaskQueries = $global:MdmFixTestState.TaskQueries; Observation = $observed }
}

$originalProgramData = $env:ProgramData
$logRoot = Join-Path $env:TEMP ('MdmServiceLogTest-' + [guid]::NewGuid().ToString('N'))
$logDirectory = Join-Path $logRoot 'Microsoft\IntuneManagementExtension\Logs'
try {
    $env:ProgramData = $logRoot
    foreach ($case in $detectionCases) {
        $runner = [PowerShell]::Create()
        try {
            $null = $runner.AddScript($runnerBody.ToString()).AddArgument($detector).AddArgument($case)
            $output = @($runner.Invoke())
            if ($runner.Streams.Error.Count) { throw ($runner.Streams.Error | Out-String) }
            $result = $output[-1]
            if ($output.Count -ne 2 -or $result.ExitCode -ne $case.Exit -or $result.Writes -ne 0 -or $result.Starts -ne 0 -or $output[0] -notmatch [regex]::Escape($case.Text)) {
                throw ("FAIL: Detection/$($case.Mode)`n" + ($output | Out-String))
            }
        }
        finally { $runner.Dispose() }
    }
    foreach ($case in $cases) {
        $runner = [PowerShell]::Create()
        try {
            $null = $runner.AddScript($runnerBody.ToString()).AddArgument($source).AddArgument($case)
            $output = @($runner.Invoke())
            if ($runner.Streams.Error.Count) { throw ($runner.Streams.Error | Out-String) }
            $result = $output[-1]
            $messages = ($output | Select-Object -SkipLast 1) -join [Environment]::NewLine
            if ($result.ExitCode -ne $case.Exit -or $result.Writes -ne $case.Writes -or $result.Starts -ne $case.Starts -or $messages -notmatch [regex]::Escape($case.Text)) {
                throw ("FAIL: $($case.Mode)/$($case.Task)`n" + ($output | Out-String))
            }
            if ($messages -match 'sync completed|session completed') { throw 'Task launch falsely claimed sync completion.' }
            if ($case.Exit -eq 0 -and $messages -notlike '*newer PushLaunch run finished with task result 0x00000000*') { throw 'Remediation returned zero without observed task completion.' }
            if ($case.Task -ne 'Ready' -and $case.Exit -eq 1 -and $messages -notmatch 'Service configuration remains enabled') { throw "Repair result not retained: $($case.Task)" }
            if (($case.Text -like 'Service*' -or $case.Text -eq 'could not be verified' -or $case.Task -in @('NoEnrollment','NoAccount','MultiplePrimary','LinkedOnly','RegistryDenied','RegistryReadDenied')) -and $result.TaskQueries -ne 0) { throw 'Queried tasks despite a service or enrollment failure.' }
        }
        finally { $runner.Dispose() }
    }
    foreach ($case in $observationCases) {
        $runner = [PowerShell]::Create()
        try {
            $null = $runner.AddScript($runnerBody.ToString()).AddArgument($source).AddArgument($case).AddArgument($taskHelper.Extent.Text)
            $output = @($runner.Invoke())
            if ($runner.Streams.Error.Count) { throw ($runner.Streams.Error | Out-String) }
            $result = $output[-1]
            $observed = $result.Observation
            if ($output.Count -ne 1 -or $result.Starts -ne $case.Starts -or $observed.Status -ne $case.Status -or $observed.NewRunObserved -ne $case.NewRun -or $observed.LastTaskResultHex -ne $case.ResultHex -or $observed.SubmissionAccepted -ne ($case.Starts -eq 1)) { throw ("FAIL: Observation/$($case.Task)`n" + ($observed | ConvertTo-Json)) }
            if ($case.Status -eq 'TimedOut' -and $observed.Error -notlike '*not stopped or retried*') { throw 'Timeout safety message missing.' }
            if (-not $case.NewRun -and $null -ne $observed.LastTaskResult) { throw 'An old result was attributed to a new run.' }
            if ($case.Task -in @('NonzeroResult','SignedResult') -and ($observed.LastTaskResult -ne 2192179211L -or $null -ne $observed.HResult)) { throw 'Task result was confused with a local exception HRESULT.' }
        }
        finally { $runner.Dispose() }
    }
    $detectionLog = [IO.File]::ReadAllText((Join-Path $logDirectory 'Detect-IntuneMdmSyncService.log'))
    $remediationLog = [IO.File]::ReadAllText((Join-Path $logDirectory 'Remediate-IntuneMdmSyncService.log'))
    if ([regex]::Matches($detectionLog, 'Starting disabled-service detection').Count -ne $detectionCases.Count -or $detectionLog -notlike '*Detection error:*exit code=2*') { throw 'Detection log did not retain repeated runs/errors.' }
    if ($remediationLog -notlike '*"LastTaskResultHex":"0x82AA000B"*' -or $remediationLog -notlike '*observation timed out*' -or $remediationLog -notlike '*exit code=0*' -or $remediationLog -notlike '*exit code=1*' -or $remediationLog -notlike '*exit code=2*') { throw 'Remediation log omitted task results or exit codes.' }
    "PASS: $($detectionCases.Count) detection, $($cases.Count) remediation/task and $($observationCases.Count) observation cases in PowerShell $($PSVersionTable.PSVersion); task helpers, help, BOM and isolated append-only logs checked. No real registry, service or task operations."
}
finally {
    $env:ProgramData = $originalProgramData
    foreach ($name in @('Detect-IntuneMdmSyncService.log', 'Remediate-IntuneMdmSyncService.log')) {
        $path = Join-Path $logDirectory $name
        if ([IO.File]::Exists($path)) { [IO.File]::Delete($path) }
    }
    foreach ($directory in @($logDirectory, (Join-Path $logRoot 'Microsoft\IntuneManagementExtension'), (Join-Path $logRoot 'Microsoft'), $logRoot)) {
        if ([IO.Directory]::Exists($directory)) { [IO.Directory]::Delete($directory) }
    }
}