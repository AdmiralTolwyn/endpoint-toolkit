[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$testPowerShell = Join-Path $PSHOME 'powershell.exe'
if ($PSVersionTable.PSEdition -eq 'Core') { $testPowerShell = Join-Path $PSHOME 'pwsh.exe' }
$source = Join-Path $PSScriptRoot 'Invoke-IntuneClientSync.ps1'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($source, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }
foreach ($definition in $ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] }) {
    . ([scriptblock]::Create($definition.Extent.Text))
}
$script:assertions = 0
function Assert-ClientTest {
    param([bool]$Condition, [string]$Name)
    if (-not $Condition) { throw "FAIL: $Name" }
    $script:assertions++
}

$missingValues = [pscustomobject]@{
    MissingObject = Get-ClientValue $null 'DeviceId'
    MissingProperty = Get-ClientValue ([pscustomobject]@{}) 'DeviceId'
} | ConvertTo-Json -Compress
Assert-ClientTest ($missingValues -ceq '{"MissingObject":null,"MissingProperty":null}') 'Absent optional values serialize as null, not empty objects'
Assert-ClientTest ((Get-ClientValue ([pscustomobject]@{ DeviceId = 'device-a' }) 'DeviceId') -ceq 'device-a') 'Present optional values remain scalar'

function New-EnrollmentFixture {
    param([string]$Provider = 'MS DM Server', [string]$Id = '11111111-1111-1111-1111-111111111111', [string]$CertificateStatus = 'Present', [string]$Tenant = 'tenant-a')
    [pscustomobject]@{
        EnrollmentId = $Id; ProviderId = $Provider; OmaDmAccountExists = $true; TenantId = $Tenant
        Certificate = [pscustomobject]@{ Status = $CertificateStatus }
        CertificateReferenceMismatch = $false; UnrecognizedCertificateReference = $false; DiscoveryUrl = ''
    }
}
$services = @(
    [pscustomobject]@{ Name = 'Schedule'; Present = $true; StartMode = 'Auto'; State = 'Running' }
    [pscustomobject]@{ Name = 'dmwappushservice'; Present = $true; StartMode = 'Manual'; State = 'Stopped' }
    [pscustomobject]@{ Name = 'IntuneManagementExtension'; Present = $true; StartMode = 'Auto'; State = 'Running' }
)
$join = [pscustomobject]@{ TenantId = 'tenant-a'; DeviceId = 'device-a'; MdmUrl = ''; DeviceAuthStatus = 'SUCCESS' }
$primary = New-EnrollmentFixture
$linked = New-EnrollmentFixture -Provider 'Microsoft Device Management' -Id '22222222-2222-2222-2222-222222222222'
$params = @{ Services = $services; Join = $join; EnrollmentRead = $true; ServiceRead = $true }

$result = Get-ClientSyncAssessment @params -Enrollments @($primary)
Assert-ClientTest $result.CanSync 'Stopped Manual-start MDM service is not a failure'
$result = Get-ClientSyncAssessment @params -Enrollments @($primary, $linked)
Assert-ClientTest $result.CanSync 'Normal primary plus linked enrollment is accepted'
$result = Get-ClientSyncAssessment @params -Enrollments @($linked)
Assert-ClientTest (-not $result.CanSync -and 'LinkedEnrollmentWithoutPrimaryAccount' -in $result.Issues) 'Linked-only enrollment is not treated as healthy primary MDM'
$result = Get-ClientSyncAssessment @params -Enrollments @()
Assert-ClientTest ('NoPrimaryIntuneOmaDmAccount' -in $result.Issues) 'Absent enrollment is explicit'
$result = Get-ClientSyncAssessment @params -Enrollments @($primary, (New-EnrollmentFixture -Id '33333333-3333-3333-3333-333333333333'))
Assert-ClientTest (-not $result.CanSync -and $result.Unknowns.Count -eq 1) 'Multiple primary accounts remain ambiguous'
foreach ($status in @('Expired', 'NotYetValid', 'Missing', 'NoPrivateKey', 'MissingReference')) {
    $result = Get-ClientSyncAssessment @params -Enrollments @((New-EnrollmentFixture -CertificateStatus $status))
    Assert-ClientTest (-not $result.CanSync -and "PrimaryCertificate:$status" -in $result.Issues) "Certificate gate: $status"
}
$result = Get-ClientSyncAssessment @params -Enrollments @((New-EnrollmentFixture -Tenant 'tenant-b'))
Assert-ClientTest ('PrimaryEnrollmentTenantMismatch' -in $result.Issues) 'Cross-tenant primary enrollment blocks trigger'
$primary.CertificateReferenceMismatch = $true
$result = Get-ClientSyncAssessment @params -Enrollments @($primary)
Assert-ClientTest (-not $result.CanSync -and $result.Unknowns.Count -eq 1) 'Conflicting certificate references remain unknown'
$primary.CertificateReferenceMismatch = $false
$services[1].StartMode = 'Disabled'
$result = Get-ClientSyncAssessment @params -Enrollments @($primary)
Assert-ClientTest ('ServiceDisabled:dmwappushservice' -in $result.Issues) 'Disabled MDM push service is flagged'
$services[1].StartMode = 'Manual'
$services[2].StartMode = 'Disabled'
$result = Get-ClientSyncAssessment @params -Enrollments @($primary)
Assert-ClientTest ($result.CanSync -and $result.Observations.Count -gt 0) 'Disabled IME is distinguished from MDM'
$services[2].StartMode = 'Auto'
$params.EnrollmentRead = $false
$result = Get-ClientSyncAssessment @params -Enrollments @()
Assert-ClientTest ('EnrollmentEvidenceUnavailable' -in $result.Unknowns -and $result.Issues.Count -eq 0) 'Unreadable registry is not misreported as no enrollment'
$params.EnrollmentRead = $true
$taskRows = @([pscustomobject]@{ Path = "\Microsoft\Windows\EnterpriseMgmt\$($primary.EnrollmentId)\"; State = 'Disabled' })
$result = Get-ClientSyncAssessment @params -Enrollments @($primary) -Tasks $taskRows -TaskRead $true
Assert-ClientTest ('PrimaryEnrollmentTasksAllDisabled' -in $result.Observations) 'Disabled tasks are visible without inventing a sync result'

$script:fixtureEnrollments = @($primary, $linked)
$script:denyEnrollment = $false
$script:denyTasks = $false
$script:syncCalls = 0
$script:syncStatus = 'Completed'
$script:execution = [pscustomobject]@{ IsSystem = $true; IsAdministrator = $true; Is64Bit = $true; Edition = 'Desktop' }
function Get-ClientExecutionContext { $script:execution }
function Get-CimInstance { param($ClassName, $OperationTimeoutSec, $ErrorAction); [pscustomobject]@{ Version = '10.0'; BuildNumber = '26100'; LastBootUpTime = [DateTime]::Now.AddHours(-1) } }
function Get-ClientJoinEvidence { $join }
function Get-ClientEnrollmentEvidence { if ($script:denyEnrollment) { throw 'fixture access denied' }; $script:fixtureEnrollments }
function Get-ClientServiceEvidence { $services }
function Get-ClientTaskEvidence { if ($script:denyTasks) { throw 'fixture task query failure' } }
function Invoke-ClientReadCommand { param($FileName, $Arguments); 'fixture proxy configuration' }
function Get-ClientImeLogEvidence { }
function Get-ClientMdmEvents { param($Since) }
function Test-ClientDiscoveryEndpoint { param($Uri); throw 'Offline test unexpectedly attempted a network probe' }
function Invoke-ClientMdmTask {
    param($Enrollments, $WaitSeconds)
    $script:syncCalls++
    $script:observedWait = $WaitSeconds
    if (@($Enrollments).Count -ne 2 -or $Enrollments[0].EnrollmentId -ne $primary.EnrollmentId) { throw 'Checked enrollment evidence was not passed to task selection.' }
    [pscustomobject]@{ Status = $script:syncStatus; Stage = 'SubmitTask'; Error = $(if ($script:syncStatus -eq 'Failed') { 'fixture submit failure' } else { $null }) }
}

$result = Invoke-IntuneClientAssessment
Assert-ClientTest ($result.Verdict -eq 'AssessmentComplete' -and $result.ExitCode -eq 0 -and $script:syncCalls -eq 0) 'Diagnostic default does not sync'
Assert-ClientTest ($result.DeviceResponsive -and -not $result.CloudLastSyncVerified -and -not $result.PolicyApplicationVerified) 'Responsiveness does not claim policy or console success'
$result = Invoke-IntuneClientAssessment -RequestSync -WhatIf
Assert-ClientTest ($script:syncCalls -eq 0 -and $result.Sync.Status -eq 'NotRequestedByShouldProcess') 'WhatIf never requests sync'
$result = Invoke-IntuneClientAssessment -RequestSync -Confirm:$false
Assert-ClientTest ($script:syncCalls -eq 1 -and $result.ExitCode -eq 0 -and $result.Verdict -eq 'MdmTaskCompleted' -and -not $result.CloudLastSyncVerified -and -not $result.PolicyApplicationVerified) 'Task completion does not claim sync success'
Assert-ClientTest ($script:observedWait -eq 120 -and $result.SchemaVersion -eq 3) 'Default wait and schema version are explicit'
$script:syncStatus = 'TimedOut'
$result = Invoke-IntuneClientAssessment -RequestSync -WaitSeconds 17 -Confirm:$false
Assert-ClientTest ($result.ExitCode -eq 2 -and $result.Verdict -eq 'MdmTaskTimedOut' -and $script:observedWait -eq 17 -and -not $result.CloudLastSyncVerified -and -not $result.PolicyApplicationVerified) 'Custom observation deadline and timeout classification preserve uncertainty'
$script:syncStatus = 'AlreadyRunning'
$result = Invoke-IntuneClientAssessment -RequestSync -Confirm:$false
Assert-ClientTest ($result.ExitCode -eq 2 -and $result.Verdict -eq 'MdmTaskAlreadyRunning') 'Busy task exit classification'
$script:syncStatus = 'Failed'
$result = Invoke-IntuneClientAssessment -RequestSync -Confirm:$false
Assert-ClientTest ($result.ExitCode -eq 1 -and $result.Verdict -eq 'MdmTaskFailed') 'Failed submission exit classification'
$script:denyEnrollment = $true
$callsBefore = $script:syncCalls
$result = Invoke-IntuneClientAssessment -RequestSync -Confirm:$false
Assert-ClientTest ($result.ExitCode -eq 2 -and $script:syncCalls -eq $callsBefore -and -not $result.EvidenceComplete) 'Incomplete prerequisites prevent sync and preserve errors'
$script:denyEnrollment = $false
$script:fixtureEnrollments = @()
$result = Invoke-IntuneClientAssessment -RequestSync -Confirm:$false
Assert-ClientTest ($result.ExitCode -eq 1 -and $result.Sync.Status -eq 'SkippedPrerequisites') 'No primary account prevents sync'
$script:fixtureEnrollments = @($primary, $linked)
$script:denyTasks = $true
$result = Invoke-IntuneClientAssessment
Assert-ClientTest ($result.ExitCode -eq 2 -and $result.CollectionErrors.Count -eq 1) 'Optional collector failures remain visible'
$script:denyTasks = $false
$script:execution.IsAdministrator = $false
$result = Invoke-IntuneClientAssessment -RequestSync -Confirm:$false
Assert-ClientTest ($result.ExitCode -eq 2 -and $null -eq $result.Assessment) 'Non-elevated context rejected before collecting'
Assert-ClientTest ($result.Sync.Status -eq 'SkippedExecutionContext' -and $result.CollectionErrors.Count -eq 1 -and $result.CollectionErrors[0] -like '*not elevated*') 'Elevation failure and skipped sync are specific'
Assert-ClientTest (-not $result.ExecutionContext.IsAdministrator -and $result.ExecutionContext.Edition -eq 'Desktop') 'Actual token and edition are retained in the report'
$script:execution.IsAdministrator = $true
$script:execution.Is64Bit = $false
$result = Invoke-IntuneClientAssessment
Assert-ClientTest ($result.ExitCode -eq 2 -and $null -eq $result.Assessment) '32-bit agent host rejected'
Assert-ClientTest ($result.CollectionErrors.Count -eq 1 -and $result.CollectionErrors[0] -like '*32-bit*') 'Bitness failure is specific'

$script:execution.Is64Bit = $true
$script:execution.Edition = 'Core'
$callsBefore = $script:syncCalls
$result = Invoke-IntuneClientAssessment
Assert-ClientTest ($result.ExitCode -eq 0 -and $result.CollectionErrors.Count -eq 0 -and $null -ne $result.Assessment -and $null -ne $result.OperatingSystem -and $result.Services.Count -eq 3) 'PowerShell Core proceeds through diagnostic collection'
Assert-ClientTest ($script:syncCalls -eq $callsBefore -and $result.Sync.Status -eq 'NotRequested' -and $result.ExecutionContext.Edition -eq 'Core') 'PowerShell Core diagnostics do not submit a task'
$result = Invoke-IntuneClientAssessment -RequestSync -WhatIf
Assert-ClientTest ($script:syncCalls -eq $callsBefore -and $result.Sync.Status -eq 'NotRequestedByShouldProcess' -and $null -ne $result.Assessment) 'PowerShell Core WhatIf collects evidence without syncing'
$result = Invoke-IntuneClientAssessment -RequestSync -Confirm:$false
Assert-ClientTest ($result.ExitCode -eq 1 -and $result.Verdict -eq 'MdmTaskFailed' -and $result.CollectionErrors.Count -eq 0 -and $null -ne $result.OperatingSystem -and $result.Services.Count -eq 3) 'A task failure in Core does not discard collected evidence'
$script:execution.Is64Bit = $false
$script:execution.IsAdministrator = $false
$callsBefore = $script:syncCalls
$result = Invoke-IntuneClientAssessment -RequestSync -WhatIf
Assert-ClientTest ($result.CollectionErrors.Count -eq 2 -and $script:syncCalls -eq $callsBefore -and $result.Sync.Status -eq 'SkippedExecutionContext') 'Bitness and elevation failures are reported without an edition block'
$rejectedReport = $result

$script:execution.Is64Bit = $true
$script:execution.IsAdministrator = $true
$script:execution.Edition = 'Desktop'
$fixtureReport = Invoke-IntuneClientAssessment
$quietOutput = @(Invoke-IntuneClientAssessment 4>&1)
Assert-ClientTest ($quietOutput.Count -eq 1 -and $quietOutput[0] -isnot [System.Management.Automation.VerboseRecord]) 'Default assessment does not emit verbose records'
$verboseOutput = @(Invoke-IntuneClientAssessment -Verbose 4>&1)
$messages = @($verboseOutput | Where-Object { $_ -is [System.Management.Automation.VerboseRecord] })
$reports = @($verboseOutput | Where-Object { $_ -isnot [System.Management.Automation.VerboseRecord] })
Assert-ClientTest ($reports.Count -eq 1 -and $reports[0].Verdict -eq 'AssessmentComplete') 'Verbose assessment returns one unchanged report'
Assert-ClientTest (($messages -join ' ') -like '*Starting client assessment*Collecting operating system*Enrollment records collected: 2*Evaluating local sync prerequisites*Diagnostic-only run*Assessment finished*') 'Assessment trace covers context, collection, counts and outcome'
$callsBefore = $script:syncCalls
$verboseOutput = @(Invoke-IntuneClientAssessment -RequestSync -WhatIf -Verbose 4>&1)
Assert-ClientTest ($script:syncCalls -eq $callsBefore -and ($verboseOutput -join ' ') -like '*Sync skipped by WhatIf*') 'Verbose WhatIf explains why sync did not run'
$script:denyEnrollment = $true
$verboseOutput = @(Invoke-IntuneClientAssessment -RequestSync -Confirm:$false -Verbose 4>&1)
Assert-ClientTest ($script:syncCalls -eq $callsBefore -and ($verboseOutput -join ' ') -like '*Enrollment collection failed: fixture access denied*Sync skipped because local prerequisites*') 'Verbose collection failures and blocked sync are explicit'
$script:denyEnrollment = $false
$script:syncStatus = 'Completed'
$verboseOutput = @(Invoke-IntuneClientAssessment -RequestSync -Confirm:$false -Verbose 4>&1)
Assert-ClientTest ($script:syncCalls -eq ($callsBefore + 1) -and ($verboseOutput -join ' ') -like '*Sync approved*Sync outcome: Completed*') 'Verbose sync still submits exactly once'
$script:execution.IsAdministrator = $false
$verboseOutput = @(Invoke-IntuneClientAssessment -RequestSync -Verbose 4>&1)
Assert-ClientTest (($verboseOutput -join ' ') -like '*Collection skipped:*not elevated*' -and ($verboseOutput -join ' ') -notlike '*Collecting operating system*') 'Verbose rejected context explains skip before collecting'
$script:execution.IsAdministrator = $true

$fixtureReport.CollectionErrors = @(('x' * 260), 'second error', 'third error', 'fourth error')
$probeDirectory = Join-Path $env:TEMP ('IntuneSyncTest-' + [Guid]::NewGuid().ToString('N'))
$probeScript = Join-Path $probeDirectory 'entry.ps1'
$probeReport = Join-Path $probeDirectory 'report.json'
$valueHelper = $ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -eq 'Get-ClientValue' } | Select-Object -First 1
$logHelper = $ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -eq 'Write-ClientSyncLog' } | Select-Object -First 1
$originalProgramData = $env:ProgramData
$logDirectory = Join-Path $probeDirectory 'Microsoft\IntuneManagementExtension\Logs'
$logPath = Join-Path $logDirectory 'Invoke-IntuneClientSync.log'
[void][System.IO.Directory]::CreateDirectory($probeDirectory)
try {
    $env:ProgramData = $probeDirectory
    foreach ($expectedExit in @(0, 1, 2)) {
        $fixtureReport.ExitCode = $expectedExit
        $fixtureJson = ($fixtureReport | ConvertTo-Json -Depth 12 -Compress).Replace("'", "''")
        $probeText = @(
            '[CmdletBinding()]'
            'param([string]$OutputFormat = ''SummaryJson'')'
            '$ErrorActionPreference = ''Stop'''
            'Set-StrictMode -Version 2.0'
            '$Sync = $false; $TimeoutSeconds = 1'
            ('$OutputPath = ''' + $probeReport.Replace("'", "''") + "'")
            $valueHelper.Extent.Text
            $logHelper.Extent.Text
            ('function Invoke-IntuneClientAssessment { [CmdletBinding()] param([switch]$RequestSync, [int]$WaitSeconds); if ($WaitSeconds -ne 1) { throw ''Entry point did not forward TimeoutSeconds'' }; Write-Verbose ''Fixture assessment progress''; ConvertFrom-Json -InputObject ''' + $fixtureJson + "' }")
            $ast.EndBlock.Statements[-1].Extent.Text
        ) -join [Environment]::NewLine
        [System.IO.File]::WriteAllText($probeScript, $probeText, (New-Object System.Text.UTF8Encoding($true)))
        foreach ($format in @('SummaryJson', 'DetailedJson', 'Object')) {
            if ($format -eq 'Object') {
                $objectCommand = '& { $result = & ''' + $probeScript.Replace("'", "''") + ''' -OutputFormat Object; $resultExit = $LASTEXITCODE; if (@($result).Count -ne 1 -or $result -isnot [pscustomobject]) { throw ''Expected a single native report object'' }; $result | ConvertTo-Json -Depth 12 -Compress; exit $resultExit }'
                $outputLines = @(& $testPowerShell -NoProfile -NonInteractive -Command $objectCommand)
            }
            elseif ($format -eq 'SummaryJson') {
                $outputLines = @(& $testPowerShell -NoProfile -NonInteractive -File $probeScript)
            }
            else {
                $outputLines = @(& $testPowerShell -NoProfile -NonInteractive -File $probeScript -OutputFormat $format)
            }
            $actualExit = $LASTEXITCODE
            $result = ConvertFrom-Json -InputObject ($outputLines -join [Environment]::NewLine)
            $savedReport = ConvertFrom-Json -InputObject ([System.IO.File]::ReadAllText($probeReport))
            Assert-ClientTest ($actualExit -eq $expectedExit -and $result.ExitCode -eq $expectedExit) "$format preserves exit code $expectedExit"
            Assert-ClientTest (@($savedReport.Enrollments).Count -eq 2 -and $savedReport.CollectionErrors[0].Length -eq 260) "$format preserves full report on disk"
            Assert-ClientTest ($result.ExecutionContext.Edition -eq 'Desktop' -and $result.PowerShellVersion -and -not $result.CloudLastSyncVerified) "$format exposes context without false cloud confirmation"
            if ($format -eq 'SummaryJson') {
                Assert-ClientTest ($outputLines.Count -eq 1 -and $result.ReportPath -eq $probeReport -and $result.CollectionErrors.Count -eq 3 -and $result.CollectionErrors[0].Length -eq 200) 'Default summary keeps compact output contract'
            }
            else {
                Assert-ClientTest (@($result.Enrollments).Count -eq 2 -and $result.CollectionErrors.Count -eq 4 -and $result.CollectionErrors[0].Length -eq 260) "$format returns all evidence and untruncated errors"
                if ($format -eq 'DetailedJson') { Assert-ClientTest ($outputLines.Count -gt 1) 'DetailedJson is indented multiline output' }
            }
        }
    }
    foreach ($format in @('SummaryJson', 'DetailedJson', 'Object')) {
        $runner = [PowerShell]::Create()
        try {
            $null = $runner.AddCommand($probeScript).AddParameter('OutputFormat', $format).AddParameter('Verbose', $true)
            $output = @($runner.Invoke())
            if ($runner.Streams.Error.Count) { throw ($runner.Streams.Error | Out-String) }
            if ($format -eq 'Object') {
                Assert-ClientTest ($output.Count -eq 1 -and $output[0] -is [pscustomobject] -and $output[0].ExitCode -eq 2) 'Verbose entry point returns a single native object on the success stream'
            }
            else {
                $parsed = ConvertFrom-Json -InputObject ($output -join [Environment]::NewLine)
                Assert-ClientTest ($parsed.ExitCode -eq 2 -and $parsed.ComputerName) "$format stays valid JSON with verbose enabled"
            }
            $verboseText = $runner.Streams.Verbose -join ' '
            Assert-ClientTest ($verboseText -like '*Writing full JSON report*Full JSON report written*Emitting*output; exit code=2*') "$format traces report output only on the verbose stream"
        }
        finally { $runner.Dispose() }
    }
    $rejectedJson = ($rejectedReport | ConvertTo-Json -Depth 12 -Compress).Replace("'", "''")
    $rejectedProbe = $probeText.Replace($fixtureJson, $rejectedJson)
    [System.IO.File]::WriteAllText($probeScript, $rejectedProbe, (New-Object System.Text.UTF8Encoding($true)))
    $outputLines = @(& $testPowerShell -NoProfile -NonInteractive -File $probeScript)
    $actualExit = $LASTEXITCODE
    $result = ConvertFrom-Json -InputObject ($outputLines -join [Environment]::NewLine)
    Assert-ClientTest ($actualExit -eq 2 -and $null -eq $result.EntraDeviceId -and $result.CollectionErrors.Count -eq 2 -and $result.Sync.Status -eq 'SkippedExecutionContext') 'Rejected context has JSON null identity and actionable failures'

    foreach ($format in @('SummaryJson', 'DetailedJson', 'Object')) {
        $errorProbe = $probeText.Replace($probeReport.Replace("'", "''"), $probeDirectory.Replace("'", "''"))
        [System.IO.File]::WriteAllText($probeScript, $errorProbe, (New-Object System.Text.UTF8Encoding($true)))
        if ($format -eq 'Object') {
            $outputLines = @(& $testPowerShell -NoProfile -NonInteractive -Command $objectCommand)
        }
        else {
            $outputLines = @(& $testPowerShell -NoProfile -NonInteractive -File $probeScript -OutputFormat $format)
        }
        $actualExit = $LASTEXITCODE
        $result = ConvertFrom-Json -InputObject ($outputLines -join [Environment]::NewLine)
        Assert-ClientTest ($actualExit -eq 2 -and $result.Verdict -eq 'ScriptError' -and $result.Error) "$format preserves script-error output and exit code"
    }
    foreach ($path in @($source, $PSCommandPath, $probeReport)) {
        $bytes = [System.IO.File]::ReadAllBytes($path)
        Assert-ClientTest ($bytes[0] -eq 239 -and $bytes[1] -eq 187 -and $bytes[2] -eq 191) "UTF-8 BOM: $([System.IO.Path]::GetFileName($path))"
    }
    $logText = [IO.File]::ReadAllText($logPath)
    Assert-ClientTest ($logText -like '*Fixture assessment progress*' -and $logText -like '*"Verdict":"ScriptError"*' -and $logText -like '*"SchemaVersion":3*') 'File log contains progress, full evidence and script errors'
    Assert-ClientTest ([regex]::Matches($logText, 'Starting client script').Count -gt 1) 'Diagnostic logs append across invocations'
    foreach ($sibling in @('Detect-IntuneMdmSyncService.ps1', 'Remediate-IntuneMdmSyncService.ps1')) {
        $siblingAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot $sibling), [ref]$tokens, [ref]$parseErrors)
        $siblingHelper = $siblingAst.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -eq 'Write-ClientSyncLog' }
        Assert-ClientTest ($parseErrors.Count -eq 0 -and $siblingHelper.Extent.Text -ceq $logHelper.Extent.Text) "Standalone logger matches: $sibling"
    }
    $runner = [PowerShell]::Create()
    try {
        $null = $runner.AddScript({
            param($HelperText, $BlockedPath)
            $ErrorActionPreference = 'Stop'
            $WarningPreference = 'Stop'
            $script:ClientSyncLogPath = $BlockedPath
            $script:ClientSyncLogWarningWritten = $false
            . ([scriptblock]::Create($HelperText))
            Write-ClientSyncLog 'first status' -OutputMessage
            Write-ClientSyncLog 'second status' -OutputMessage
        }.ToString()).AddArgument($logHelper.Extent.Text).AddArgument($probeDirectory)
        $output = @($runner.Invoke())
        Assert-ClientTest ($runner.Streams.Error.Count -eq 0 -and $runner.Streams.Warning.Count -eq 1 -and $output.Count -eq 2 -and $output[1] -eq 'second status') 'Unwritable logs warn once without changing status output or throwing'
    }
    finally { $runner.Dispose() }
}
finally {
    $env:ProgramData = $originalProgramData
    if ([IO.File]::Exists($logPath)) { [IO.File]::Delete($logPath) }
    foreach ($directory in @($logDirectory, (Join-Path $probeDirectory 'Microsoft\IntuneManagementExtension'), (Join-Path $probeDirectory 'Microsoft'))) {
        if ([IO.Directory]::Exists($directory)) { [IO.Directory]::Delete($directory) }
    }
    if ([System.IO.File]::Exists($probeScript)) { [System.IO.File]::Delete($probeScript) }
    if ([System.IO.File]::Exists($probeReport)) { [System.IO.File]::Delete($probeReport) }
    [System.IO.Directory]::Delete($probeDirectory)
}

"PASS: $script:assertions offline assertions in PowerShell $($PSVersionTable.PSVersion). No real sync, service changes, registry changes or network probes."