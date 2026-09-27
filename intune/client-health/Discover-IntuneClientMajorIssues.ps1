<#
.SYNOPSIS
    Discovers major Intune client issues for a ConfigMgr Compliance Baseline.
.DESCRIPTION
    Import this entire file into a CI PowerShell discovery script with data type
    String and a compliance rule equal to Passed. No companion files needed.
    Returns one scalar for a first-pass major-fault screen. Unavailable core
    evidence throws without a compliance value. Stopped services and historical
    event/task errors are diagnostic context only. No sync or remediation.
    Deploy as SYSTEM using 64-bit Windows PowerShell 5.1. Interactive runs are
    not rejected merely for being non-SYSTEM. Collection requires elevation and
    a 64-bit process. Includes bounded direct DNS/TCP diagnostic probes.
    No task triggers, certificate/enrollment repairs or IME/process restarts.
.OUTPUTS
    System.String. Passed, or IssueDetected with major-fault reasons and log path. Expected: Passed.
    A matching value means only the specified local condition was observed,
    not successful Intune check-in or policy/application processing.
.NOTES
    Log: %ProgramData%\Microsoft\IntuneManagementExtension\Logs\Discover-IntuneClientMajorIssues.log.
    Logs append and contain detailed diagnostic evidence. Logging failures warn
    once without altering the discovery result. Configure retention separately.
    Do not configure a remediation script or enable automatic remediation.
    See the adjacent README for scope, settings, limitations and pilot checks.
#>
#Requires -Version 5.1
[CmdletBinding()]
param()

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

function Get-ClientServiceEvidence {
    <#
    .SYNOPSIS
        Reads MDM-related and IME service state without changing any service.
    .OUTPUTS
        One PSCustomObject per expected service: Name, Present, State, StartMode.
        Missing services have Present=false and null state/start mode. CIM errors
        propagate. Stopped and Disabled are deliberately kept distinct.
    #>
    $names = @('Schedule', 'dmwappushservice', 'DmEnrollmentSvc', 'IntuneManagementExtension', 'WpnService')
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
        PSCustomObject with CanSync, MdmPrerequisiteStatus, ImeServiceStatus and
        string arrays Issues, Unknowns and Observations. Prerequisites passing
        or an IME service running does not prove successful management check-in.
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
    $imeServiceStatus = 'Unknown'
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
        if (Get-ClientValue $enrollment 'TenantReferenceMismatch') { [void]$unknowns.Add('PrimaryTenantReferencesRequireInvestigation') }
        if (Get-ClientValue $enrollment.Certificate 'ExpiringSoon') { [void]$observations.Add('PrimaryCertificateExpiresWithin30Days') }
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
        if (-not $ime -or -not $ime.Present) {
            $imeServiceStatus = 'MissingVerifyRequired'
            [void]$observations.Add('ImeAbsent:WhetherRequiredDependsOnAssignments')
        }
        elseif ($ime.StartMode -eq 'Disabled') {
            $imeServiceStatus = 'Disabled'
            [void]$observations.Add('ImeDisabled:InvestigateSeparatelyFromMdm')
        }
        elseif ($ime.State -ne 'Running') {
            if ($ime.State) { $imeServiceStatus = 'NotRunning' }
            [void]$observations.Add('ImeNotRunning:InvestigateSeparatelyFromMdm')
        }
        else { $imeServiceStatus = 'Running' }
    }
    foreach ($enrollment in $linked) {
        if ($enrollment.Certificate.Status -ne 'Present') { [void]$observations.Add("LinkedCertificate:$($enrollment.EnrollmentId):$($enrollment.Certificate.Status)") }
        if (Get-ClientValue $enrollment.Certificate 'ExpiringSoon') { [void]$observations.Add("LinkedCertificateExpiresWithin30Days:$($enrollment.EnrollmentId)") }
        if ($primary.Count -eq 1 -and $enrollment.TenantId -and $primary[0].TenantId -and $enrollment.TenantId -ne $primary[0].TenantId) {
            [void]$observations.Add("LinkedEnrollmentTenantMismatch:$($enrollment.EnrollmentId)")
        }
    }
    $mdmPrerequisiteStatus = 'Passed'
    if ($unknowns.Count -gt 0) { $mdmPrerequisiteStatus = 'Unknown' }
    if ($issues.Count -gt 0) { $mdmPrerequisiteStatus = 'IssueDetected' }
    [pscustomobject]@{
        CanSync = ($issues.Count -eq 0 -and $unknowns.Count -eq 0)
        MdmPrerequisiteStatus = $mdmPrerequisiteStatus
        ImeServiceStatus = $imeServiceStatus
        Issues = @($issues.ToArray())
        Unknowns = @($unknowns.ToArray())
        Observations = @($observations.ToArray())
    }
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
        Issuer, DaysToExpiry, ExpiringSoon and HasPrivateKey. ExpiringSoon means
        valid dates with expiry within 30 days; it is advisory, not renewal failure.
        Status is MissingReference, Missing, NoPrivateKey,
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
        Issuer        = $null
        NotBeforeUtc  = $null
        NotAfterUtc   = $null
        DaysToExpiry  = $null
        ExpiringSoon  = $null
        HasPrivateKey = $null
    }
    if ($Thumbprint) {
        $certPath = "Cert:\LocalMachine\My\$Thumbprint"
        $result.Status = 'Missing'
        if (Test-Path -LiteralPath $certPath -ErrorAction Stop) {
            $certificate = Get-Item -LiteralPath $certPath -ErrorAction Stop
            $now = [DateTime]::UtcNow
            $notBefore = $certificate.NotBefore.ToUniversalTime()
            $notAfter = $certificate.NotAfter.ToUniversalTime()
            $result.Subject = $certificate.Subject
            $result.Issuer = $certificate.Issuer
            $result.NotBeforeUtc = $notBefore.ToString('o')
            $result.NotAfterUtc = $notAfter.ToString('o')
            $result.DaysToExpiry = [int][Math]::Floor(($notAfter - $now).TotalDays)
            $result.ExpiringSoon = ($notBefore -le $now -and $notAfter -gt $now -and $notAfter -le $now.AddDays(30))
            $result.HasPrivateKey = $certificate.HasPrivateKey
            $result.Status = 'Present'
            if (-not $certificate.HasPrivateKey) { $result.Status = 'NoPrivateKey' }
            if ($notBefore -gt $now) { $result.Status = 'NotYetValid' }
            if ($notAfter -le $now) { $result.Status = 'Expired' }
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
        both certificates when references conflict, without trying to repair them.
        Uses AADTenantID with a GUID suffix in UPN as fallback; preserves both
        tenant sources so conflicts are visible rather than guessed away.
    .OUTPUTS
        Zero or more PSCustomObjects containing EnrollmentId, ProviderId, TenantId,
        tenant sources/conflicts, raw EnrollmentState/Type, OmaDmAccountExists,
        DiscoveryUrl, both certificate references and their evidence. The caller
        must use @() when it needs a stable array. Access failures and malformed
        DMPCertThumbPrint values throw, not imply no enrollment.
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
        $dmpThumbprint = $thumbprint
        $reference = [string](Get-ClientValue $account 'SslClientCertReference')
        $referenceMatch = [regex]::Match($reference, '(?i)^MY;System;([a-f0-9\s]+)$')
        $referenceMismatch = $false
        $accountThumbprint = $null
        $recognizedReference = $false
        $certificateSource = 'Missing'
        if ($thumbprint) { $certificateSource = 'DMPCertThumbPrint' }
        if ($referenceMatch.Success -and ($referenceMatch.Groups[1].Value -replace '\s', '') -match '^[a-f0-9]{40}$') {
            $recognizedReference = $true
            $accountThumbprint = $referenceMatch.Groups[1].Value -replace '\s', ''
            if (-not $thumbprint) { $thumbprint = $accountThumbprint; $certificateSource = 'SslClientCertReference' }
            elseif ($thumbprint -ne $accountThumbprint) { $referenceMismatch = $true }
        }
        if ($thumbprint -and $thumbprint -notmatch '^[a-f0-9]{40}$') { throw "Unexpected certificate thumbprint in enrollment $($key.PSChildName)" }
        $certificate = Get-ClientCertificateEvidence -Thumbprint $thumbprint
        $accountCertificate = $null
        if ($recognizedReference) {
            $accountCertificate = $certificate
            if ($referenceMismatch) { $accountCertificate = Get-ClientCertificateEvidence -Thumbprint $accountThumbprint }
        }
        $registryTenant = [string](Get-ClientValue $values 'AADTenantID')
        $upnTenant = $null
        $upnTenantMatch = [regex]::Match([string](Get-ClientValue $values 'UPN'), '(?i)@([a-f0-9]{8}-(?:[a-f0-9]{4}-){3}[a-f0-9]{12})\s*$')
        if ($upnTenantMatch.Success) { $upnTenant = $upnTenantMatch.Groups[1].Value }
        $tenantId = $registryTenant
        $tenantSource = 'Missing'
        if ($registryTenant) { $tenantSource = 'AADTenantID' }
        elseif ($upnTenant) { $tenantId = $upnTenant; $tenantSource = 'UPNSuffix' }
        [pscustomobject]@{
            EnrollmentId = $key.PSChildName
            ProviderId = $provider
            TenantId = $tenantId
            TenantIdSource = $tenantSource
            RegistryTenantId = $registryTenant
            UpnTenantId = $upnTenant
            TenantReferenceMismatch = [bool]($registryTenant -and $upnTenant -and $registryTenant -ne $upnTenant)
            EnrollmentState = Get-ClientValue $values 'EnrollmentState'
            EnrollmentType = Get-ClientValue $values 'EnrollmentType'
            OmaDmAccountExists = [bool]$accountExists
            DiscoveryUrl = [string](Get-ClientValue $values 'DiscoveryServiceFullURL')
            DmpCertThumbprint = $dmpThumbprint
            OmaDmCertificateReference = $reference
            OmaDmCertificateThumbprint = $accountThumbprint
            CertificateReferenceSource = $certificateSource
            Certificate = $certificate
            OmaDmCertificate = $accountCertificate
            CertificateReferenceMismatch = $referenceMismatch
            UnrecognizedCertificateReference = [bool]($reference -and -not $recognizedReference)
        }
    }
}

function Get-ClientTaskEvidence {
    <#
    .SYNOPSIS
        Reads EnterpriseMgmt and NonCritical task history without starting tasks.
    .OUTPUTS
        Zero or more PSCustomObjects with Path, Name, State, LastRunUtc and
        LastTaskResult, LastTaskResultHex, NextRunUtc, action details and a renewal
        candidate flag based only on the task name. Sentinel dates become null.
        Query errors throw; unknown history is not treated as a missing task.
    .NOTES
        A task run or scheduler result is activity evidence, not proof of a
        successful MDM exchange. Renewal candidates are a name heuristic only,
        not proof of certificate renewal or failure. Nonzero can be scheduler state.
    #>
    $tasks = @(Get-ScheduledTask -ErrorAction Stop | Where-Object { $_.TaskPath -like '\Microsoft\Windows\EnterpriseMgmt\*' -or $_.TaskPath -like '\Microsoft\Windows\EnterpriseMgmtNonCritical\*' })
    foreach ($task in $tasks) {
        $info = Get-ScheduledTaskInfo -InputObject $task -ErrorAction Stop
        $lastRun = $null
        if ($info.LastRunTime.Year -gt 1999) { $lastRun = $info.LastRunTime.ToUniversalTime().ToString('o') }
        $nextRun = $null
        if ($info.NextRunTime -and $info.NextRunTime.Year -gt 1999) { $nextRun = $info.NextRunTime.ToUniversalTime().ToString('o') }
        $resultHex = $null
        if ($null -ne $info.LastTaskResult) { $resultHex = '0x{0:X8}' -f ([long]$info.LastTaskResult -band 0xFFFFFFFFL) }
        [pscustomobject]@{
            Path = $task.TaskPath
            Name = $task.TaskName
            State = [string]$task.State
            Enabled = Get-ClientValue $task.Settings 'Enabled'
            AllowDemandStart = Get-ClientValue $task.Settings 'AllowDemandStart'
            PrincipalUserId = Get-ClientValue $task.Principal 'UserId'
            LastRunUtc = $lastRun
            NextRunUtc = $nextRun
            LastTaskResult = $info.LastTaskResult
            LastTaskResultHex = $resultHex
            IsRenewalCandidate = [bool]($task.TaskName -match 'Renew')
            ActionExecutables = @($task.Actions | ForEach-Object { [string](Get-ClientValue $_ 'Execute') })
            ActionArguments = @($task.Actions | ForEach-Object { [string](Get-ClientValue $_ 'Arguments') })
        }
    }
}

function Get-ClientEnrollmentDiagnostics {
    <#
    .SYNOPSIS
        Summarizes collected evidence per enrollment with investigation guidance.
    .DESCRIPTION
        Does not query or modify the device. Findings are diagnostic conditions,
        not a causal diagnosis or proof of stale enrollment. Task history and
        renewal-name candidates do not certify check-in or certificate renewal.
    .PARAMETER Enrollments
        All collected primary and linked enrollment records, including residue.
    .PARAMETER Join
        Current dsregcmd evidence; missing tenant means comparison is unknown.
    .PARAMETER Tasks
        Collected critical and NonCritical tasks, including raw result codes.
    .PARAMETER TaskRead
        Whether task collection completed. False means unknown, not missing tasks.
    .OUTPUTS
        One diagnostic row per enrollment with certificate, tenant and task
        summaries, Findings and NextChecks. Never selects a best-scoring account.
    #>
    param([object[]]$Enrollments, $Join, [object[]]$Tasks, [bool]$TaskRead)

    $joinedTenant = [string](Get-ClientValue $Join 'TenantId')
    foreach ($enrollment in $Enrollments) {
        $findings = New-Object 'System.Collections.Generic.List[string]'
        $nextChecks = New-Object 'System.Collections.Generic.List[string]'
        $tenantComparison = 'Unknown'
        if (Get-ClientValue $enrollment 'TenantReferenceMismatch') {
            $tenantComparison = 'ConflictingReferences'
            [void]$findings.Add('EnrollmentTenantReferencesConflict')
            [void]$nextChecks.Add('Compare AADTenantID and the recorded UPN tenant suffix with the intended tenant; do not choose one silently.')
        }
        elseif ($joinedTenant -and $enrollment.TenantId) {
            $tenantComparison = 'Match'
            if ($joinedTenant -ne $enrollment.TenantId) {
                $tenantComparison = 'Mismatch'
                [void]$findings.Add('EnrollmentTenantDiffersFromJoinedTenant')
                [void]$nextChecks.Add('Confirm the intended tenant and migration history; a mismatch alone does not authorize enrollment deletion.')
            }
        }
        else {
            [void]$findings.Add('TenantComparisonUnavailable')
            [void]$nextChecks.Add('Collect the current joined tenant and enrollment tenant before deciding whether enrollment is cross-tenant.')
        }
        if (-not $enrollment.OmaDmAccountExists) {
            [void]$findings.Add('EnrollmentWithoutOmaDmAccount')
            [void]$nextChecks.Add('Review this registry footprint against tasks and enrollment history; it may be incomplete or residual, not an active account.')
        }
        $certificate = Get-ClientValue $enrollment 'Certificate'
        $certificateStatus = [string](Get-ClientValue $certificate 'Status')
        if ($certificateStatus -ne 'Present') {
            [void]$findings.Add("Certificate:$certificateStatus")
            [void]$nextChecks.Add('Inspect the referenced LocalMachine certificate, validity dates, device clock and private-key association; do not substitute the newest certificate.')
        }
        if (Get-ClientValue $certificate 'ExpiringSoon') {
            [void]$findings.Add('CertificateExpiresWithin30Days')
            [void]$nextChecks.Add('Review certificate-renewal history and management connectivity before expiry; this warning is not proof renewal failed.')
        }
        if ($enrollment.CertificateReferenceMismatch -or $enrollment.UnrecognizedCertificateReference) {
            [void]$findings.Add('CertificateReferenceRequiresInvestigation')
            [void]$nextChecks.Add('Compare DMPCertThumbPrint with OMADM Accounts SslClientCertReference and both referenced certificates.')
        }
        $criticalTasks = @()
        $nonCriticalTasks = @()
        $renewalTasks = @()
        $taskCount = $null
        $nonCriticalCount = $null
        $pushState = 'Unknown'
        if ($TaskRead) {
            $criticalPath = "\Microsoft\Windows\EnterpriseMgmt\$($enrollment.EnrollmentId)\"
            $nonCriticalPath = "\Microsoft\Windows\EnterpriseMgmtNonCritical\$($enrollment.EnrollmentId)\"
            $criticalTasks = @($Tasks | Where-Object { $_.Path -like "$criticalPath*" })
            $nonCriticalTasks = @($Tasks | Where-Object { $_.Path -like "$nonCriticalPath*" })
            $taskCount = $criticalTasks.Count
            $nonCriticalCount = $nonCriticalTasks.Count
            if ($taskCount -eq 0) { [void]$findings.Add('CriticalEnrollmentTasksMissing') }
            elseif (@($criticalTasks | Where-Object { $_.State -ne 'Disabled' }).Count -eq 0) { [void]$findings.Add('CriticalEnrollmentTasksAllDisabled') }
            $pushTasks = @($criticalTasks | Where-Object { (Get-ClientValue $_ 'Name') -eq 'PushLaunch' -and $_.Path -eq $criticalPath })
            $pushState = 'Missing'
            if ($pushTasks.Count -eq 1) { $pushState = [string]$pushTasks[0].State }
            elseif ($pushTasks.Count -gt 1) { $pushState = 'Ambiguous' }
            if ($pushState -in @('Missing', 'Disabled', 'Ambiguous', 'Unknown')) {
                [void]$findings.Add("PushLaunch:$pushState")
                [void]$nextChecks.Add('Review the selected enrollment task definition and permissions; this diagnostic does not recreate or enable tasks.')
            }
            $renewalTasks = @(($criticalTasks + $nonCriticalTasks) | Where-Object { Get-ClientValue $_ 'IsRenewalCandidate' })
            if ($renewalTasks.Count -eq 0) { [void]$findings.Add('RenewalTaskNotIdentifiedByName') }
            foreach ($task in $renewalTasks) {
                if ($task.State -eq 'Disabled') { [void]$findings.Add("RenewalTaskDisabled:$($task.Name)") }
                $code = Get-ClientValue $task 'LastTaskResult'
                if ($task.LastRunUtc -and $null -ne $code -and $task.State -notin @('Running', 'Queued') -and ([long]$code -band 0xFFFFFFFFL) -notin @(0, 0x41300, 0x41301, 0x41303, 0x41325)) {
                    [void]$findings.Add("RenewalTaskResultNeedsReview:$($task.Name):$($task.LastTaskResultHex)")
                }
            }
            if (@($findings | Where-Object { $_ -like 'RenewalTask*' }).Count -gt 0) {
                [void]$nextChecks.Add('Correlate renewal-candidate task names, run times and raw results with MDM events. Missing name matches or old codes do not establish current renewal failure.')
            }
        }
        else {
            [void]$findings.Add('TaskEvidenceUnavailable')
            [void]$nextChecks.Add('Resolve the task collection error before inferring missing or disabled enrollment tasks.')
        }
        [pscustomobject]@{
            EnrollmentId = $enrollment.EnrollmentId
            ProviderId = $enrollment.ProviderId
            OmaDmAccountExists = $enrollment.OmaDmAccountExists
            TenantId = $enrollment.TenantId
            TenantIdSource = Get-ClientValue $enrollment 'TenantIdSource'
            TenantComparison = $tenantComparison
            CertificateStatus = $certificateStatus
            CertificateThumbprint = Get-ClientValue $certificate 'Thumbprint'
            CertificateIssuer = Get-ClientValue $certificate 'Issuer'
            CertificateHasPrivateKey = Get-ClientValue $certificate 'HasPrivateKey'
            CertificateNotAfterUtc = Get-ClientValue $certificate 'NotAfterUtc'
            CertificateDaysToExpiry = Get-ClientValue $certificate 'DaysToExpiry'
            OmaDmCertificateStatus = Get-ClientValue (Get-ClientValue $enrollment 'OmaDmCertificate') 'Status'
            CriticalTaskCount = $taskCount
            NonCriticalTaskCount = $nonCriticalCount
            PushLaunchState = $pushState
            RenewalTaskCandidates = $renewalTasks
            Findings = @($findings.ToArray())
            NextChecks = @($nextChecks.ToArray())
        }
    }
}

function Get-ClientOmaDmProcessEvidence {
    <#
    .SYNOPSIS
        Reads an omadmclient process snapshot without stopping or starting processes.
    .OUTPUTS
        ProcessId, StartedUtc and lifetime CpuSeconds per process. CIM errors throw.
        CPU totals are not utilization samples; process presence/count does not
        prove a hang, a certificate fault or successful MDM communication.
    #>
    foreach ($process in @(Get-CimInstance -ClassName Win32_Process -Filter "Name='omadmclient.exe'" -OperationTimeoutSec 10 -ErrorAction Stop)) {
        $startedUtc = $null
        if ($process.CreationDate) { $startedUtc = $process.CreationDate.ToUniversalTime().ToString('o') }
        $cpuSeconds = $null
        if ($null -ne $process.KernelModeTime -and $null -ne $process.UserModeTime) {
            $cpuSeconds = [Math]::Round(([double]$process.KernelModeTime + [double]$process.UserModeTime) / 10000000, 2)
        }
        [pscustomobject]@{ ProcessId = $process.ProcessId; StartedUtc = $startedUtc; CpuSeconds = $cpuSeconds }
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

function Get-ClientEventEvidence {
    <#
    .SYNOPSIS
        Reads up to 12 recent warnings/errors from the specified diagnostic log.
    .PARAMETER LogName
        Windows event channel to query without enabling or clearing it.
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
    param([string]$LogName, [DateTime]$Since)

    try {
        Get-WinEvent -FilterHashtable @{
            LogName = $LogName
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

function Get-ClientMdmEvents {
    <#
    .SYNOPSIS
        Reads recent MDM Admin warnings/errors without assuming task correlation.
    .PARAMETER Since
        Earliest timestamp to include.
    #>
    param([DateTime]$Since)
    Get-ClientEventEvidence -LogName 'Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin' -Since $Since
}

function Get-ClientWpnEvents {
    <#
    .SYNOPSIS
        Reads WPN operational warnings/errors without enabling the event channel.
    .PARAMETER Since
        Earliest timestamp to include. Missing/inaccessible channels throw.
    #>
    param([DateTime]$Since)
    $channel = Get-WinEvent -ListLog 'Microsoft-Windows-PushNotification-Platform/Operational' -ErrorAction Stop
    if (-not $channel.IsEnabled) { throw 'WPN operational event channel is disabled; event coverage is unavailable. No channel settings were changed.' }
    Get-ClientEventEvidence -LogName 'Microsoft-Windows-PushNotification-Platform/Operational' -Since $Since
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

function Get-ClientHealthSummary {
    <#
    .SYNOPSIS
        Separates observed local issues, review findings and missing evidence.
    .PARAMETER Report
        Collected report including Coverage. This function performs no device I/O.
    .OUTPUTS
        Status, Issues, ReviewFindings, Unknowns and ServiceChecks. NoIssuesDetected
        means only the implemented checks found no issue; never confirmed sync.
        Unknown takes precedence but does not erase observed issues/findings.
    #>
    param($Report)
    $issues = New-Object 'System.Collections.Generic.List[string]'
    $review = New-Object 'System.Collections.Generic.List[string]'
    $unknowns = New-Object 'System.Collections.Generic.List[string]'
    foreach ($check in $Report.Coverage.Keys) {
        if ($Report.Coverage[$check] -notin @('Collected', 'NotApplicable')) { [void]$unknowns.Add("${check}:$($Report.Coverage[$check])") }
    }
    if ($Report.Assessment) {
        foreach ($issue in $Report.Assessment.Issues) { [void]$issues.Add($issue) }
        foreach ($unknown in $Report.Assessment.Unknowns) { [void]$unknowns.Add($unknown) }
        foreach ($observation in $Report.Assessment.Observations) { [void]$review.Add($observation) }
    }
    if ([string](Get-ClientValue $Report.Join 'DeviceAuthStatus') -like 'FAILED*') { [void]$issues.Add('EntraDeviceAuthenticationFailed') }
    $serviceChecks = @(foreach ($name in @('Schedule', 'dmwappushservice', 'DmEnrollmentSvc', 'IntuneManagementExtension', 'WpnService')) {
        $status = 'Unknown'
        $service = $Report.Services | Where-Object { $_.Name -eq $name } | Select-Object -First 1
        if ($Report.Coverage.Services -eq 'Collected') {
            if (-not $service -or -not $service.Present) {
                $status = 'Missing'
                if ($name -eq 'IntuneManagementExtension') { [void]$review.Add('ImeMissing:VerifyAssignments') }
                else { [void]$issues.Add("ServiceMissing:$name") }
            }
            elseif ($service.StartMode -notin @('Auto', 'Manual', 'Disabled') -or -not $service.State) {
                [void]$unknowns.Add("ServiceMetadataUnavailable:$name")
            }
            elseif ($service.StartMode -eq 'Disabled') { $status = 'Disabled'; [void]$issues.Add("ServiceDisabled:$name") }
            elseif ($service.State -eq 'Running') { $status = 'Running' }
            else {
                $status = 'NotRunning'
                if ($name -in @('Schedule', 'IntuneManagementExtension', 'WpnService') -or $service.StartMode -eq 'Auto') {
                    [void]$review.Add("ServiceNotRunning:${name}:$($service.State)")
                }
            }
        }
        [pscustomobject]@{ Name = $name; Status = $status; StartMode = Get-ClientValue $service 'StartMode'; State = Get-ClientValue $service 'State' }
    })
    foreach ($diagnostic in $Report.EnrollmentDiagnostics) {
        $prefix = $diagnostic.EnrollmentId
        foreach ($finding in $diagnostic.Findings) {
            if ($finding -match '^(CriticalEnrollmentTasksMissing|CriticalEnrollmentTasksAllDisabled|PushLaunch:(Missing|Disabled)|Certificate:(Expired|Missing|NoPrivateKey|NotYetValid|MissingReference)|EnrollmentTenantDiffersFromJoinedTenant)') {
                if ($diagnostic.OmaDmAccountExists) { [void]$issues.Add("${prefix}:$finding") }
                else { [void]$review.Add("${prefix}:$finding") }
            }
            elseif ($finding -in @('TenantComparisonUnavailable', 'EnrollmentTenantReferencesConflict', 'CertificateReferenceRequiresInvestigation', 'TaskEvidenceUnavailable', 'PushLaunch:Unknown', 'PushLaunch:Ambiguous')) {
                [void]$unknowns.Add("${prefix}:$finding")
            }
            else { [void]$review.Add("${prefix}:$finding") }
        }
        if ($Report.Coverage.Tasks -eq 'Collected' -and $diagnostic.OmaDmAccountExists) {
            $taskPath = "\Microsoft\Windows\EnterpriseMgmt\$prefix\"
            $pushTasks = @($Report.Tasks | Where-Object { $_.Path -eq $taskPath -and $_.Name -eq 'PushLaunch' })
            if ($pushTasks.Count -eq 1) {
                $task = $pushTasks[0]
                if ($null -eq (Get-ClientValue $task 'Enabled') -or $null -eq (Get-ClientValue $task 'AllowDemandStart') -or -not (Get-ClientValue $task 'PrincipalUserId')) {
                    [void]$unknowns.Add("${prefix}:PushLaunchDefinitionUnavailable")
                }
                else {
                    if (-not $task.Enabled) { [void]$issues.Add("${prefix}:PushLaunchDisabled") }
                    if (-not $task.AllowDemandStart) { [void]$issues.Add("${prefix}:PushLaunchDemandStartDisabled") }
                    if ($task.PrincipalUserId -notin @('SYSTEM', 'NT AUTHORITY\SYSTEM', 'S-1-5-18')) { [void]$issues.Add("${prefix}:PushLaunchUnexpectedPrincipal") }
                    $executables = @(Get-ClientValue $task 'ActionExecutables')
                    $arguments = @(Get-ClientValue $task 'ActionArguments')
                    $expectedExecutable = Join-Path $env:SystemRoot 'System32\deviceenroller.exe'
                    $argumentPattern = '^\s*/o\s+(?:"{0}"|{0})\s+/c\s+/z\s*$' -f [regex]::Escape($prefix)
                    if ($executables.Count -ne 1 -or $arguments.Count -ne 1 -or [Environment]::ExpandEnvironmentVariables([string]$executables[0]) -ine $expectedExecutable -or $arguments[0] -notmatch $argumentPattern) {
                        [void]$issues.Add("${prefix}:PushLaunchUnexpectedAction")
                    }
                }
            }
        }
    }
    foreach ($task in $Report.Tasks) {
        $code = Get-ClientValue $task 'LastTaskResult'
        if ($task.LastRunUtc -and $null -ne $code -and $task.State -notin @('Running', 'Queued') -and ([long]$code -band 0xFFFFFFFFL) -notin @(0, 0x41300, 0x41301, 0x41303, 0x41325)) {
            [void]$review.Add("TaskHistory:$($task.Path)$($task.Name):$($task.LastTaskResultHex)")
        }
    }
    foreach ($probe in $Report.DiscoveryConnectivity) {
        if ($probe.Error -or $probe.Dns -ne 'Resolved' -or $probe.Tcp -ne 'Connected') { [void]$review.Add("DirectConnectivity:$($probe.HostName):$($probe.Error)") }
    }
    foreach ($log in $Report.ImeLogs) { if (-not $log.Present) { [void]$review.Add("ImeLogMissing:$($log.Name)") } }
    if ($Report.OmaDmProcessCount -gt 1) { [void]$review.Add('MultipleOmaDmProcesses:SampleBeforeCallingHang') }
    if (@($Report.RecentMdmWarningsAndErrors).Count -gt 0) { [void]$review.Add('RecentMdmErrors:CorrelateWithFailedCheckIn') }
    if (@($Report.RecentWpnWarningsAndErrors).Count -gt 0) { [void]$review.Add('RecentWpnErrors:CorrelateWithPushDelivery') }
    $status = 'NoIssuesDetected'
    if ($review.Count) { $status = 'ReviewRequired' }
    if ($issues.Count) { $status = 'IssueDetected' }
    if ($unknowns.Count) { $status = 'Unknown' }
    [pscustomobject]@{
        Status = $status
        Issues = @($issues.ToArray() | Select-Object -Unique)
        ReviewFindings = @($review.ToArray() | Select-Object -Unique)
        Unknowns = @($unknowns.ToArray() | Select-Object -Unique)
        ServiceChecks = $serviceChecks
    }
}

function Invoke-ClientEvidenceCollection {
    <#
    .SYNOPSIS
        Collects all read-only client evidence except the final event snapshot.
    .DESCRIPTION
        Rejects unsupported execution contexts before collection. Collectors run
        independently so one failure does not hide other evidence. Discovery
        probes are deduplicated by host/port and their results remain advisory.
        Never requests sync. Shared by interactive diagnostics and baseline
        discovery so adapting output cannot silently remove collector coverage.
    .OUTPUTS
        PSCustomObject containing the full report, Assessment, Sync, Verdict and
        ExitCode. Expected collector failures are retained in CollectionErrors.
        Does not serialize, write reports or exit; the script entry point does.
    .NOTES
        An attempted task submission's verdict takes precedence over diagnostic collector
        errors; EvidenceComplete and CollectionErrors must be checked separately.
        CloudLastSyncVerified and PolicyApplicationVerified always remain false.
    #>
    [CmdletBinding()]
    param()

    $errors = New-Object 'System.Collections.Generic.List[string]'
    $report = [ordered]@{
        SchemaVersion = 3
        ComputerName = $env:COMPUTERNAME
        CollectedUtc = [DateTime]::UtcNow.ToString('o')
        DeviceResponsive = $true
        MdmPrerequisiteStatus = 'NotAssessed'
        ImeServiceStatus = 'NotAssessed'
        WpnServiceStatus = 'NotAssessed'
        LocalHealthStatus = 'NotAssessed'
        Health = $null
        Coverage = [ordered]@{ OperatingSystem = 'NotCollected'; Join = 'NotCollected'; Enrollments = 'NotCollected'; Services = 'NotCollected'; Tasks = 'NotCollected'; WinHttpProxy = 'NotCollected'; ImeLogs = 'NotCollected'; OmaDmProcesses = 'NotCollected'; DiscoveryConnectivity = 'NotCollected'; MdmEvents = 'NotCollected'; WpnEvents = 'NotCollected' }
        RunningAsSystem = $false
        PowerShellVersion = [string]$PSVersionTable.PSVersion
        ExecutionContext = $null
        OperatingSystem = $null
        Join = $null
        Enrollments = @()
        EnrollmentDiagnostics = @()
        Services = @()
        Tasks = @()
        WinHttpProxy = $null
        DiscoveryConnectivity = @()
        ImeLogs = @()
        OmaDmProcesses = @()
        OmaDmProcessCount = $null
        RecentMdmWarningsAndErrors = @()
        RecentWpnWarningsAndErrors = @()
        Assessment = $null
        CollectionErrors = @()
        EvidenceComplete = $false
        Sync = [pscustomobject]@{ Status = 'NotRequested' }
        CloudLastSyncVerified = $false
        PolicyApplicationVerified = $false
        Verdict = 'AssessmentIncomplete'
        ExitCode = 2
    }
    Write-Verbose "Starting client assessment on $($report.ComputerName)."
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
        $report.Coverage.OperatingSystem = 'Collected'
    }
    catch { $report.Coverage.OperatingSystem = 'Failed'; [void]$errors.Add("OperatingSystem: $($_.Exception.Message)"); Write-Verbose "OperatingSystem collection failed: $($_.Exception.Message)" }
    Write-Verbose 'Collecting Entra join evidence using dsregcmd /status.'
    try { $report.Join = Get-ClientJoinEvidence; $report.Coverage.Join = 'Collected' }
    catch { $report.Coverage.Join = 'Failed'; [void]$errors.Add("Join: $($_.Exception.Message)"); Write-Verbose "Join collection failed: $($_.Exception.Message)" }
    Write-Verbose 'Collecting primary/linked MDM enrollments and referenced certificates.'
    try { $report.Enrollments = @(Get-ClientEnrollmentEvidence); $enrollmentRead = $true; $report.Coverage.Enrollments = 'Collected' }
    catch { $report.Coverage.Enrollments = 'Failed'; [void]$errors.Add("Enrollments: $($_.Exception.Message)"); Write-Verbose "Enrollment collection failed: $($_.Exception.Message)" }
    if ($enrollmentRead) { Write-Verbose "Enrollment records collected: $($report.Enrollments.Count)." }
    Write-Verbose 'Collecting MDM-related and IME service state; no services will be changed.'
    try { $report.Services = @(Get-ClientServiceEvidence); $serviceRead = $true; $report.Coverage.Services = 'Collected' }
    catch { $report.Coverage.Services = 'Failed'; [void]$errors.Add("Services: $($_.Exception.Message)"); Write-Verbose "Service collection failed: $($_.Exception.Message)" }
    Write-Verbose 'Collecting EnterpriseMgmt task state and scheduler history.'
    try { $report.Tasks = @(Get-ClientTaskEvidence); $taskRead = $true; $report.Coverage.Tasks = 'Collected' }
    catch { $report.Coverage.Tasks = 'Failed'; [void]$errors.Add("Tasks: $($_.Exception.Message)"); Write-Verbose "Task collection failed: $($_.Exception.Message)" }
    if ($taskRead) { Write-Verbose "EnterpriseMgmt tasks collected: $($report.Tasks.Count). Task activity does not prove MDM sync success." }
    Write-Verbose 'Reading WinHTTP proxy configuration.'
    try { $report.WinHttpProxy = Invoke-ClientReadCommand -FileName "$env:SystemRoot\System32\netsh.exe" -Arguments 'winhttp show proxy'; $report.Coverage.WinHttpProxy = 'Collected' }
    catch { $report.Coverage.WinHttpProxy = 'Failed'; [void]$errors.Add("WinHttpProxy: $($_.Exception.Message)"); Write-Verbose "Proxy collection failed: $($_.Exception.Message)" }
    Write-Verbose 'Reading IME log existence and last-write times.'
    try { $report.ImeLogs = @(Get-ClientImeLogEvidence); $report.Coverage.ImeLogs = 'Collected' }
    catch { $report.Coverage.ImeLogs = 'Failed'; [void]$errors.Add("ImeLogs: $($_.Exception.Message)"); Write-Verbose "IME log collection failed: $($_.Exception.Message)" }
    Write-Verbose 'Reading omadmclient process snapshot; lifetime CPU is not current utilization or proof of a hang.'
    try {
        $report.OmaDmProcesses = @(Get-ClientOmaDmProcessEvidence)
        $report.OmaDmProcessCount = $report.OmaDmProcesses.Count
        $report.Coverage.OmaDmProcesses = 'Collected'
    }
    catch { $report.Coverage.OmaDmProcesses = 'Failed'; [void]$errors.Add("OmaDmProcesses: $($_.Exception.Message)"); Write-Verbose "OMA-DM process collection failed: $($_.Exception.Message)" }
    $urls = @($report.Enrollments | Select-Object -ExpandProperty DiscoveryUrl) + @([string](Get-ClientValue $report.Join 'MdmUrl'))
    $targets = @{}
    foreach ($url in $urls) {
        $uri = $null
        if ([uri]::TryCreate($url, [UriKind]::Absolute, [ref]$uri) -and $uri.Scheme -eq 'https') {
            $targets[$uri.Authority] = $uri
        }
    }
    Write-Verbose "Probing $($targets.Count) discovery host(s) using direct DNS/TCP; results do not validate TLS or the proxy path."
    try {
        $report.DiscoveryConnectivity = @(foreach ($uri in $targets.Values) {
            Write-Verbose "Probing $($uri.DnsSafeHost):$($uri.Port)."
            $probe = Test-ClientDiscoveryEndpoint -Uri $uri
            Write-Verbose "Discovery probe result: DNS=$($probe.Dns); TCP=$($probe.Tcp); Error=$($probe.Error)"
            $probe
        })
        $report.Coverage.DiscoveryConnectivity = 'Collected'
        if ($targets.Count -eq 0) {
            $report.Coverage.DiscoveryConnectivity = 'NotApplicable'
            if (@($report.Enrollments | Where-Object { $_.OmaDmAccountExists }).Count -gt 0) { $report.Coverage.DiscoveryConnectivity = 'NoTarget' }
        }
    }
    catch { $report.Coverage.DiscoveryConnectivity = 'Failed'; [void]$errors.Add("DiscoveryConnectivity: $($_.Exception.Message)") }
    Write-Verbose 'Evaluating local sync prerequisites.'
    $report.Assessment = Get-ClientSyncAssessment -Enrollments $report.Enrollments -Services $report.Services -Join $report.Join -EnrollmentRead $enrollmentRead -ServiceRead $serviceRead -Tasks $report.Tasks -TaskRead $taskRead
    if ($enrollmentRead) {
        $report.EnrollmentDiagnostics = @(Get-ClientEnrollmentDiagnostics -Enrollments $report.Enrollments -Join $report.Join -Tasks $report.Tasks -TaskRead $taskRead)
        foreach ($diagnostic in $report.EnrollmentDiagnostics) {
            Write-Verbose "Enrollment $($diagnostic.EnrollmentId): certificate=$($diagnostic.CertificateStatus); days to expiry=$($diagnostic.CertificateDaysToExpiry); tenant=$($diagnostic.TenantComparison); findings=$($diagnostic.Findings -join ', ')."
        }
    }
    $report.MdmPrerequisiteStatus = $report.Assessment.MdmPrerequisiteStatus
    $report.ImeServiceStatus = $report.Assessment.ImeServiceStatus
    Write-Verbose "MDM prerequisites: $($report.MdmPrerequisiteStatus); IME service: $($report.ImeServiceStatus). Check-in and workload success remain unverified."
    Write-Verbose "Assessment: CanSync=$($report.Assessment.CanSync); issues=$($report.Assessment.Issues.Count); unknowns=$($report.Assessment.Unknowns.Count); observations=$($report.Assessment.Observations.Count)."
    foreach ($issue in $report.Assessment.Issues) { Write-Verbose "Prerequisite issue: $issue" }
    foreach ($unknown in $report.Assessment.Unknowns) { Write-Verbose "Unresolved prerequisite: $unknown" }
    foreach ($observation in $report.Assessment.Observations) { Write-Verbose "Observation: $observation" }
    $report.Verdict = 'AssessmentComplete'
    $report.ExitCode = 0
    if ($errors.Count -gt 0 -or $report.Assessment.Unknowns.Count -gt 0) { $report.Verdict = 'AssessmentIncomplete'; $report.ExitCode = 2 }
    if ($report.Assessment.Issues.Count -gt 0) { $report.Verdict = 'LocalPrerequisiteIssue'; $report.ExitCode = 1 }
    $report.CollectionErrors = @($errors.ToArray())
    [pscustomobject]$report
}

function Complete-ClientEvidenceReport {
    <#
    .SYNOPSIS
        Adds final MDM/WPN event snapshots and comprehensive local findings.
    .PARAMETER Report
        Shared read-only collection report. Interactive mode calls this after
        any approved sync; baseline calls it without any sync operation.
    #>
    param($Report)
    if (-not $Report.Assessment) { return $Report }
    foreach ($channel in @('Mdm', 'Wpn')) {
        $coverageKey = $channel + 'Events'
        $property = 'Recent' + $channel + 'WarningsAndErrors'
        Write-Verbose "Reading up to 12 $channel warnings/errors from the last three days; these are not task-correlated."
        try {
            if ($channel -eq 'Mdm') { $Report.$property = @(Get-ClientMdmEvents -Since ([DateTime]::Now.AddDays(-3))) }
            else { $Report.$property = @(Get-ClientWpnEvents -Since ([DateTime]::Now.AddDays(-3))) }
            $Report.Coverage[$coverageKey] = 'Collected'
        }
        catch {
            $Report.Coverage[$coverageKey] = 'Failed'
            $Report.CollectionErrors += "${coverageKey}:$($_.Exception.Message)"
            Write-Verbose "$channel event collection failed: $($_.Exception.Message)"
            if ($Report.Verdict -eq 'AssessmentComplete') { $Report.Verdict = 'AssessmentIncomplete'; $Report.ExitCode = 2 }
        }
    }
    $Report.Health = Get-ClientHealthSummary -Report $Report
    $Report.LocalHealthStatus = $Report.Health.Status
    $Report.WpnServiceStatus = ($Report.Health.ServiceChecks | Where-Object { $_.Name -eq 'WpnService' } | Select-Object -First 1).Status
    $Report.EvidenceComplete = ($Report.CollectionErrors.Count -eq 0 -and $Report.Health.Unknowns.Count -eq 0)
    Write-Verbose "Local diagnostic status: $($Report.LocalHealthStatus); WPN service: $($Report.WpnServiceStatus). No cloud sync success is implied."
    Write-Verbose "Assessment finished: $($Report.Verdict); exit code=$($Report.ExitCode); collection errors=$($Report.CollectionErrors.Count)."
    $Report
}

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$script:ClientSyncLogWarningWritten = $false
$script:ClientSyncLogPath = Join-Path $env:ProgramData 'Microsoft\IntuneManagementExtension\Logs\Discover-IntuneClientMajorIssues.log'
$evidence = [ordered]@{
    ComputerName = $env:COMPUTERNAME
    CollectedUtc = [DateTime]::UtcNow.ToString('o')
    Setting = 'IntuneClientDiscovery'
    ExecutionContext = $null
    Join = $null
    Enrollments = @()
    Services = @()
    Tasks = @()
    EnrollmentDiagnostics = @()
    Assessment = $null
    DiscoveryValue = $null
    Error = $null
}
try {
    Write-ClientSyncLog 'Starting ConfigMgr IntuneClientDiscovery discovery.'
    $evidence.ExecutionContext = Get-ClientExecutionContext
    $report = Invoke-ClientEvidenceCollection
    $report = Complete-ClientEvidenceReport -Report $report
    foreach ($property in $report.PSObject.Properties) { $evidence[$property.Name] = $property.Value }
    $unknowns = New-Object 'System.Collections.Generic.List[string]'
    if (-not $report.Health) {
        foreach ($message in $report.CollectionErrors) { [void]$unknowns.Add($message) }
        if ($unknowns.Count -eq 0) { [void]$unknowns.Add('No local assessment was produced.') }
    }
    else {
        foreach ($check in @('Join', 'Enrollments', 'Services', 'Tasks')) {
            if ($report.Coverage[$check] -ne 'Collected') {
                [void]$unknowns.Add("${check}:$($report.Coverage[$check])")
                foreach ($message in @($report.CollectionErrors | Where-Object { $_ -like "${check}:*" })) { [void]$unknowns.Add($message) }
            }
        }
        foreach ($unknown in $report.Health.Unknowns) {
            if ($unknown -match '^(Primary|MultiplePrimary|EnrollmentEvidence|ServiceEvidence|[a-fA-F0-9]{8}-(?:[a-fA-F0-9]{4}-){3}[a-fA-F0-9]{12}:)') { [void]$unknowns.Add($unknown) }
        }
        foreach ($service in $report.Services) {
            if ($service.Present -and $service.StartMode -notin @('Auto', 'Manual', 'Disabled')) { [void]$unknowns.Add("ServiceStartupUnavailable:$($service.Name)") }
        }
    }
    $majorIssues = @()
    if ($report.Health) { $majorIssues = @($report.Health.Issues) }
    $screenStatus = 'Passed'
    if ($majorIssues.Count -gt 0) { $screenStatus = 'IssueDetected' }
    if ($unknowns.Count -gt 0) { $screenStatus = 'Unknown' }
    $evidence.BaselineAssessment = [pscustomobject]@{ Status = $screenStatus; Issues = $majorIssues; Unknowns = @($unknowns.ToArray() | Select-Object -Unique) }
    if ($screenStatus -eq 'Unknown') {
        $details = @($evidence.BaselineAssessment.Unknowns | Select-Object -First 5 | ForEach-Object {
            $message = [string]$_ -replace '[\r\n]+', ' '
            if ($message.Length -gt 400) { $message = $message.Substring(0, 400) + '...' }
            $message
        })
        throw ("Discovery evidence is unresolved: {0}. Full report: {1}" -f ($details -join '; '), $script:ClientSyncLogPath)
    }
    $value = 'Passed'
    if ($majorIssues.Count -gt 0) {
        $reasons = @($majorIssues | Select-Object -Unique -First 5 | ForEach-Object {
            $reason = ([string]$_ -replace '^[a-fA-F0-9]{8}-(?:[a-fA-F0-9]{4}-){3}[a-fA-F0-9]{12}:', '') -replace '[\r\n]+', ' '
            $reason = $reason -replace '^ServiceDisabled:(.+)$', '$1 disabled' -replace '^ServiceMissing:(.+)$', '$1 missing'
            if ($reason.Length -gt 180) { $reason = $reason.Substring(0, 177) + '...' }
            $reason
        })
        if ($majorIssues.Count -gt 5) { $reasons += '+ more issues in log' }
        $value = 'IssueDetected | {0} | Details: {1}' -f ($reasons -join '; '), $script:ClientSyncLogPath
    }
    $evidence.DiscoveryValue = $value
    Write-ClientSyncLog $evidence
    $evidence.DiscoveryValue
}
catch {
    $evidence.DiscoveryValue = $null
    $evidence.Error = $_.Exception.Message
    Write-ClientSyncLog $evidence
    throw
}
