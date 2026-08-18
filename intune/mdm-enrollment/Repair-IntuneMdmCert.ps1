#Requires -Version 5.1
<#
.SYNOPSIS
    Detects (and optionally repairs) AVD/cloned hosts whose Intune MDM device
    certificate has expired or lost its private key, which wedges omadmclient.exe
    in a CPU-spinning cert-health loop (CertificateManager::GetSslClientCertWithTipTest).

.DESCRIPTION
    Background: on a batch of cloned hosts the "Microsoft Intune MDM Device CA"
    client certificate aged past NotAfter (shared template enrollment -> shared
    expiry) and auto-renewal never succeeded. omadmclient.exe then loops in the
    cert-selection self-test, burning CPU.

    AUDIT mode (default) is READ-ONLY. It reports:
      - the Intune MDM device cert thumbprint, NotBefore/NotAfter, days-to-expiry
      - whether the cert is expired or expiring soon
      - the enrollment renewal scheduled task and its last result
      - whether the Intune enrollment endpoint is reachable (HTTPS)
      - omadmclient.exe process count + accumulated CPU
      - a Verdict and RecommendedAction

    REPAIR mode is DESTRUCTIVE and GATED. When the host is fixable it:
      1. Refuses to run if the Intune enrollment endpoint is unreachable
         (re-enroll would just fail) unless -Force is supplied.
      2. Backs up the enrollment registry keys and records cert thumbprints.
      3. Stops omadmclient.exe.
      4. Removes the enrollment artifacts (scheduled tasks in the critical +
         non-critical EnterpriseMgmt folders, plus registry GUID keys) for both
         the Intune (MS DM Server) and the MMP-C / declared-config
         (Microsoft Device Management) channels, and clears MmpcEnrollmentFlag
         (leaves the SCCM co-management bridge alone).
      5. Removes the expired Intune MDM device cert(s) from LocalMachine\My.
      6. Re-enrolls headlessly via deviceenroller.exe. -EnrollMode User (default)
         uses /c /AutoEnrollMDM (run in the logged-on user's session);
         -EnrollMode Device uses /c /AutoEnrollMDMUsingAADDeviceCredential.
      7. Polls the DeviceManagement event log (event 75 = success, 76 = failure)
         and the cert store for up to -WaitSeconds to confirm the new enrollment,
         then reports.

.PARAMETER Mode
    Audit (default, read-only) or Repair (destructive, gated).

.PARAMETER ExpiryWarningDays
    Days-to-expiry threshold that flags a cert as "expiring soon". Default 30.

.PARAMETER IssuerMatch
    Regex matched against the cert Issuer to locate the MDM device cert.
    Default 'Microsoft Intune.*MDM Device CA' (matches both the production
    'Microsoft Intune MDM Device CA' and the 'Microsoft Intune Beta MDM Device CA'
    issuers; does not match the unrelated '...Device Management' enrollment cert).

.PARAMETER LogPath
    Override the log file location. By default the script logs (CMTrace format)
    to %ProgramData%\Microsoft\IntuneManagementExtension\Logs\Repair-IntuneMdmCert.log,
    falling back to %TEMP% if that folder is not writable.

.PARAMETER Force
    Proceed even when the Intune enrollment endpoint is unreachable (re-enroll
    would otherwise be skipped). This does NOT suppress the confirmation prompt -
    use -Confirm:$false for unattended runs.

.PARAMETER RepairHealthy
    Allow Repair to run on a host whose cert is currently valid (Verdict Healthy
    or CertExpiringSoon) - a deliberate teardown + re-enroll / cert rotation.
    Without it, Repair only acts on CertExpired / CertMissing. It does not lower
    any other gate; the confirmation prompt still applies unless -Confirm:$false.

.PARAMETER WaitSeconds
    In Repair mode, how long to wait after triggering re-enrollment for it to
    complete, polling the DeviceManagement event log (75 = Auto MDM Enroll
    succeeded, 76 = failed) and the cert store. Default 180. 0 disables the wait
    (fire-and-forget). Events 77-80 (retry / DMGetAadDeviceToken 'Access is
    denied') are transient and are not treated as a verdict.

.PARAMETER EnrollMode
    Which credential the headless re-enroll uses:
      User (default) -> deviceenroller.exe /c /AutoEnrollMDM. Correct for a
        normal user-driven Entra-joined device (the expired enrollment carries a
        UPN). MUST run in the logged-on user's session - as SYSTEM it cannot get
        a user token and fails with an AAD 0xCAA8xxxx error.
      Device -> deviceenroller.exe /c /AutoEnrollMDMUsingAADDeviceCredential.
        Correct for co-managed / Autopilot device-prep / AVD multi-session.

.EXAMPLE
    .\Repair-IntuneMdmCert.ps1
    Read-only audit of the local host.

.EXAMPLE
    .\Repair-IntuneMdmCert.ps1 | Export-Csv .\mdm-cert-audit.csv -NoTypeInformation
    Audit the local host and export the result object to CSV.

.EXAMPLE
    .\Repair-IntuneMdmCert.ps1 -Mode Repair -WhatIf
    Show exactly what Repair WOULD do on the local host without changing anything.

.EXAMPLE
    .\Repair-IntuneMdmCert.ps1 -Mode Repair -Force -Confirm:$false
    Repair the local host unattended (-Confirm:$false skips the prompt, -Force
    proceeds even if the enrollment endpoint probe says unreachable).

.INPUTS
    None. This script does not accept pipeline input.

.OUTPUTS
    System.Management.Automation.PSCustomObject
    A single audit object is emitted to the pipeline (suitable for Export-Csv /
    Where-Object). Key properties:
      ComputerName       - host the audit ran on.
      Verdict            - one of: NotEnrolled, CertMissing, CertExpired,
                           CertExpiringSoon, Healthy.
      Fixable            - $true when Verdict is CertExpired or CertMissing.
      RecommendedAction  - human-readable next step for this host.
      EnrollmentId       - GUID of the primary (Intune 'MS DM Server') enrollment.
      ProviderID / UPN   - enrollment provider and user.
      EnrollmentCount    - total MDM enrollments discovered.
      MmpcEnrolled / MmpcEnrollmentId / MmpcEnrollmentFlag - MMP-C (declared-config)
                           co-enrollment state; a lingering one blocks re-enroll
                           with 0x8018000A, so Repair tears it down too.
      MdmCert*           - thumbprint, subject, NotBefore/NotAfter of the selected
                           Intune MDM device cert.
      DaysToExpiry / IsExpired / IsExpiringSoon - cert lifetime flags.
      HasPrivateKey      - whether the cert claims an associated private key.
      MdmCertCount       - number of certs matching IssuerMatch.
      RenewalTaskName / RenewalLastRun / RenewalLastResult / AnyRenewalFailed -
                           cert-renewal scheduled-task state.
      EndpointTested / EndpointReachable / EndpointDetail - enrollment endpoint probe.
      OmaDmClientCount / OmaDmCpuSeconds - omadmclient.exe pressure.
      AllEnrollments / AllMdmCerts / AllRenewalTasks - full diagnostic collections.
      CollectedAt        - timestamp of the audit.

.NOTES
    PowerShell 5.1. Run elevated for Repair.

    Author : Anton Romanyuk
    Version : 1.0

    SUPPORTABILITY: Microsoft does not document or support manually deleting the
    enrollment registry roots below. Their enrollment-diagnostics article names
    only the first key and gives a heuristic, never a list (PolicyManager,
    EnterpriseResourceManager and OMADM are never mentioned in a teardown
    context).

    The eight-key teardown list traces to a single 2020 community blog post -
    Maxime Rastello, "Manually re-enroll a co-managed or Hybrid Azure AD Join
    Windows 10 PC to Microsoft Intune without loosing current configuration" -
    which every later script/article copies.

    Sources for the enrollment-teardown + headless re-enroll approach:
      - call4cloud (Rudy Ooms), "Troubleshooting Intune MDM Device enrollment
        errors": https://call4cloud.nl/intune-device-enrollment-errors-mdm-enrollment/
          * Section 5.5 ("Device previously AADR enrolled") is the basis for the
            registry-key list + scheduled-task / EnterpriseMgmt task-folder cleanup.
          * Section 7 documents the headless re-enroll switch used here:
            deviceenroller.exe /c /AutoEnrollMDMUsingAADDeviceCredential.
    This script discovers the enrollment GUID from the 'MS DM Server' registry
    key and only ever iterates concrete GUIDs, avoiding the empty-$EnrollmentID
    bug in older copies of the call4cloud script (which could delete ALL tasks).

.LINK
    https://call4cloud.nl/intune-device-enrollment-errors-mdm-enrollment/
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter()]
    [ValidateSet('Audit', 'Repair')]
    [string]$Mode = 'Audit',

    [Parameter()]
    [int]$ExpiryWarningDays = 30,

    [Parameter()]
    [string]$IssuerMatch = 'Microsoft Intune.*MDM Device CA',

    [Parameter()]
    [string]$LogPath,

    [Parameter()]
    [switch]$Quiet,

    [Parameter()]
    [switch]$RepairHealthy,

    [Parameter()]
    [switch]$Force,

    [Parameter()]
    [int]$WaitSeconds = 180,

    [Parameter()]
    [ValidateSet('Device', 'User')]
    [string]$EnrollMode = 'User'
)

# --------------------------------------------------------------------------
# LOGGING  (CMTrace format -> Intune IME log folder; read with CMTrace)
# --------------------------------------------------------------------------
# Default to the Intune Management Extension log folder so these logs sit
# alongside the other Intune client logs and open cleanly in CMTrace. Falls
# back to %TEMP% if that folder can't be created/written (e.g. non-managed or
# non-elevated test box). -LogPath overrides the file location.
$Script:LogComponent = 'Repair-IntuneMdmCert'
if ($LogPath) {
    $Script:LogFile = $LogPath
}
else {
    $imeLogDir = Join-Path $env:ProgramData 'Microsoft\IntuneManagementExtension\Logs'
    $Script:LogFile = Join-Path $imeLogDir 'Repair-IntuneMdmCert.log'
}
$Script:LogReady = $false
# Echo every logged line to the console as it is written; -Quiet turns it off
# for pipeline-only use (the emitted object and the log file are unaffected).
$Script:LogToConsole = -not $Quiet

# Strict-mode-safe property read: registry enrollment keys don't all carry the
# same values, and Set-StrictMode turns a missing-property read into a throw.
function Get-Prop {
    param($InputObject, [string]$Name)
    if ($InputObject -and $InputObject.PSObject.Properties[$Name]) { $InputObject.$Name } else { $null }
}

# Decode the common MDM-enrollment HRESULTs/Win32 codes from deviceenroller and
# the DM event log so the operator doesn't hand-decode. 0xCAA8xxxx = AAD/WAM
# token-broker errors whose low word is the underlying WinHTTP/WinINet code.
function Get-EnrollErrorText {
    param($Code)
    if ($null -eq $Code) { return $null }
    if ($Code -is [string]) {
        $s = $Code -replace '^0x', ''
        try { $u = [uint32]([Convert]::ToInt64($s, 16)) } catch { return "$Code" }
    }
    else {
        # 0xFFFFFFFFL (long): a bare 0xFFFFFFFF parses as int32 -1 in PS 5.1,
        # which would make [uint32] of a negative exit code throw.
        try { $u = [uint32]([int64]$Code -band 0xFFFFFFFFL) } catch { return "$Code" }
    }
    $hex = ('0x{0:X8}' -f $u)
    if ($hex -eq '0x8018000A') { return "$hex (MENROLL_E_DEVICE_ALREADY_ENROLLED - another live MDM enrollment is still present)" }
    if ($hex -like '0xCAA8*') {
        $low = [Convert]::ToInt32($hex.Substring(6), 16)
        $txt = switch ($low) {
            12002 { 'WinHTTP timeout' }
            12007 { 'WinHTTP name not resolved' }
            12029 { 'WinHTTP cannot connect' }
            12175 { 'WinHTTP TLS/secure-channel failure' }
            default { "WinHTTP/WinINet $low" }
        }
        return "$hex (AAD/WAM token error; $txt)"
    }
    return $hex
}

function Write-Log {
    <#
        Append one CMTrace-format line to $Script:LogFile and mirror it to the
        verbose stream. Severity maps to CMTrace type: INFO=1, WARN=2, ERROR=3.
        Self-healing: creates the log dir, rotates at ~5 MB, and on any write
        failure falls back to %TEMP% once so logging never throws.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    if ($Script:LogToConsole) {
        $fg = switch ($Level) { 'ERROR' { 'Red' } 'WARN' { 'Yellow' } default { 'Gray' } }
        Write-Host $Message -ForegroundColor $fg
    }
    else {
        switch ($Level) {
            'ERROR' { Write-Verbose "[ERROR] $Message" }
            'WARN'  { Write-Verbose "[WARN] $Message" }
            default { Write-Verbose $Message }
        }
    }

    $type = switch ($Level) { 'WARN' { 2 } 'ERROR' { 3 } default { 1 } }
    $now  = Get-Date
    $tz   = [System.TimeZoneInfo]::Local.GetUtcOffset($now).TotalMinutes
    $cm = '<![LOG[{0}]LOG]!><time="{1:HH:mm:ss.fff}{2:+000;-000}" date="{1:MM-dd-yyyy}" component="{3}" context="" type="{4}" thread="{5}" file="">' -f `
        $Message, $now, $tz, $Script:LogComponent, $type, $PID

    $write = {
        param($path)
        $dir = Split-Path -Path $path -Parent
        if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null }
        if ((Test-Path $path) -and ((Get-Item $path).Length -gt 5MB)) {
            $bak = [System.IO.Path]::ChangeExtension($path, '.lo_')
            Remove-Item $bak -Force -ErrorAction SilentlyContinue
            Rename-Item $path $bak -Force -ErrorAction SilentlyContinue
        }
        [System.IO.File]::AppendAllText($path, $cm + "`r`n", [System.Text.UTF8Encoding]::new($false))
    }

    try {
        & $write $Script:LogFile
        $Script:LogReady = $true
    }
    catch {
        # Fall back to %TEMP% once, then keep using it for the rest of the run.
        $fallback = Join-Path $env:TEMP 'Repair-IntuneMdmCert.log'
        if ($Script:LogFile -ne $fallback) {
            $Script:LogFile = $fallback
            try { & $write $Script:LogFile; $Script:LogReady = $true }
            catch { Write-Verbose "Write-Log failed (fallback): $($_.Exception.Message)" }
        }
        else {
            Write-Verbose "Write-Log failed: $($_.Exception.Message)"
        }
    }
}

# --------------------------------------------------------------------------
# AUDIT WORKER  (self-contained)
# --------------------------------------------------------------------------
$AuditWorker = {
    param(
        [string]$IssuerMatch,
        [int]$ExpiryWarningDays
    )

    $now = Get-Date
    $guidRx = '^[0-9A-Fa-f]{8}-([0-9A-Fa-f]{4}-){3}[0-9A-Fa-f]{12}$'

    # --- 1. MDM enrollment(s) -------------------------------------------------
    $enrollments = @()
    $enrollRoot = 'HKLM:\SOFTWARE\Microsoft\Enrollments'
    if (Test-Path $enrollRoot) {
        Get-ChildItem $enrollRoot -ErrorAction SilentlyContinue |
            Where-Object { $_.PSChildName -match $guidRx } |
            ForEach-Object {
                $p = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
                $disco = Get-Prop $p 'DiscoveryServiceFullURL'
                $upnVal = Get-Prop $p 'UPN'
                $provId = Get-Prop $p 'ProviderID'
                if ($p -and ($disco -or $upnVal -or $provId)) {
                    $enrollments += [pscustomobject]@{
                        EnrollmentId           = $_.PSChildName
                        ProviderID             = $provId
                        UPN                    = $upnVal
                        EnrollmentState        = Get-Prop $p 'EnrollmentState'
                        DiscoveryUrl           = $disco
                        SslClientCertReference = Get-Prop $p 'SslClientCertReference'
                    }
                }
            }
    }
    # Prefer the Intune ('MS DM Server') enrollment as primary.
    $primaryEnroll = $enrollments | Where-Object { $_.ProviderID -eq 'MS DM Server' } | Select-Object -First 1
    if (-not $primaryEnroll) { $primaryEnroll = $enrollments | Select-Object -First 1 }
    Write-Log ("[1/5] Enrollments: found {0} (providers: {1})." -f $enrollments.Count, (($enrollments | ForEach-Object { $_.ProviderID }) -join ', '))
    if ($primaryEnroll) {
        Write-Verbose ("      Primary enrollment: {0} (ProviderID='{1}', UPN='{2}', State={3})." -f $primaryEnroll.EnrollmentId, $primaryEnroll.ProviderID, $primaryEnroll.UPN, $primaryEnroll.EnrollmentState)
    }
    else {
        Write-Log '      No MDM enrollment found.' 'WARN'
    }

    # MMP-C / declared-configuration ('Microsoft Device Management') co-enrollment.
    # A lingering one blocks a fresh AutoEnrollMDM with 0x8018000A; Repair tears it
    # down too. Surface it here so the operator sees it before repairing.
    $mmpcEnroll = $enrollments | Where-Object { $_.ProviderID -eq 'Microsoft Device Management' } | Select-Object -First 1
    $mmpcFlag = Get-Prop (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Enrollments' -ErrorAction SilentlyContinue) 'MmpcEnrollmentFlag'
    if ($mmpcEnroll) {
        Write-Log ("      Also enrolled in MMP-C / declared config ({0}) - Repair tears this down too." -f $mmpcEnroll.EnrollmentId) 'WARN'
    }
    if ($mmpcFlag -eq 2) {
        Write-Log '      MmpcEnrollmentFlag=2 present (blocks re-enroll) - Repair will clear it.' 'WARN'
    }

    # --- 2. MDM device cert(s) by Issuer -------------------------------------
    $mdmCerts = @()
    Get-ChildItem 'Cert:\LocalMachine\My' -ErrorAction SilentlyContinue |
        Where-Object { $_.Issuer -match $IssuerMatch } |
        ForEach-Object { $mdmCerts += $_ }
    # Current cert = the one with the latest NotAfter.
    $cert = $mdmCerts | Sort-Object NotAfter -Descending | Select-Object -First 1
    Write-Log ("[2/5] MDM certs matching /{0}/ in LocalMachine\My: {1}." -f $IssuerMatch, $mdmCerts.Count)
    foreach ($mc in $mdmCerts) {
        Write-Verbose ("      - {0}  NotAfter={1:yyyy-MM-dd}  Issuer='{2}'" -f $mc.Thumbprint, $mc.NotAfter, $mc.Issuer)
    }

    $certThumb = $null; $certSubject = $null; $notBefore = $null; $notAfter = $null
    $daysToExpiry = $null; $isExpired = $false; $isExpiringSoon = $false
    $hasPrivKey = $false
    if ($cert) {
        $certThumb   = $cert.Thumbprint
        $certSubject = $cert.Subject
        $notBefore   = $cert.NotBefore
        $notAfter    = $cert.NotAfter
        $daysToExpiry = [int][math]::Floor(($cert.NotAfter - $now).TotalDays)
        $isExpired   = ($cert.NotAfter -lt $now)
        $isExpiringSoon = (-not $isExpired) -and ($daysToExpiry -le $ExpiryWarningDays)
        $hasPrivKey  = $cert.HasPrivateKey
        Write-Log ("      Selected cert {0}: NotAfter={1:yyyy-MM-dd} ({2}d), Expired={3}, HasPrivateKey={4}." -f `
            $certThumb, $notAfter, $daysToExpiry, $isExpired, $hasPrivKey)
    }
    else {
        Write-Verbose '      No matching MDM device cert selected.'
    }

    # --- 3. Renewal scheduled task(s) ----------------------------------------
    $renewTasks = @()
    foreach ($en in $enrollments) {
        $tp = "\Microsoft\Windows\EnterpriseMgmt\$($en.EnrollmentId)\"
        Get-ScheduledTask -TaskPath $tp -ErrorAction SilentlyContinue | ForEach-Object {
            $t = $_
            $info = $t | Get-ScheduledTaskInfo -ErrorAction SilentlyContinue
            $hex = $null
            if ($info -and ($null -ne $info.LastTaskResult)) {
                $hex = ('0x{0:X8}' -f ([uint32]($info.LastTaskResult -band 0xFFFFFFFF)))
            }
            $taskArgs = $null
            if ($t.Actions -and $t.Actions.Count -gt 0) {
                $taskArgs = Get-Prop $t.Actions[0] 'Arguments'
            }
            # All EnterpriseMgmt tasks shell deviceenroller.exe, so match on the
            # name/arguments only to identify the cert-renewal task specifically.
            $isRenew = ($t.TaskName -match 'Renew') -or ($taskArgs -match 'Renew')
            $renewTasks += [pscustomobject]@{
                EnrollmentId   = $en.EnrollmentId
                TaskName       = $t.TaskName
                IsRenewal      = [bool]$isRenew
                LastRunTime    = if ($info) { $info.LastRunTime } else { $null }
                LastTaskResult = if ($info) { $info.LastTaskResult } else { $null }
                LastResultHex  = $hex
                NextRunTime    = if ($info) { $info.NextRunTime } else { $null }
            }
        }
    }
    $renewFailed = @($renewTasks | Where-Object { $_.IsRenewal -and $null -ne $_.LastTaskResult -and $_.LastTaskResult -ne 0 }).Count -gt 0
    $primaryRenew = $renewTasks | Where-Object { $_.IsRenewal } | Select-Object -First 1
    Write-Log ("[3/5] Scheduled tasks under EnterpriseMgmt: {0} (renewal-classified: {1})." -f $renewTasks.Count, @($renewTasks | Where-Object { $_.IsRenewal }).Count)
    foreach ($rt in ($renewTasks | Where-Object { $_.IsRenewal })) {
        Write-Verbose ("      - {0}: LastResult={1}, LastRun={2}, Next={3}" -f $rt.TaskName, $rt.LastResultHex, $rt.LastRunTime, $rt.NextRunTime)
    }
    if ($primaryRenew) { Write-Verbose ("      Primary renewal task: '{0}' (AnyRenewalFailed={1})." -f $primaryRenew.TaskName, $renewFailed) }

    # --- 4. Intune enrollment endpoint reachability --------------------------
    # Reachability is judged by a TCP-443 connect to the endpoint host (the
    # authoritative signal). The enrollment discovery endpoint is a WCF/SOAP
    # service that may not answer a bare HTTP HEAD/GET, so an HTTP timeout there
    # is a false negative. The HTTP probe below only enriches the detail; ANY
    # HTTP status (even 4xx/5xx) is also treated as reachable.
    $endpointUrl = if ($primaryEnroll -and $primaryEnroll.DiscoveryUrl) { $primaryEnroll.DiscoveryUrl }
                   else { 'https://enrollment.manage.microsoft.com/EnrollmentServer/Discovery.svc' }
    $endpointReachable = $null
    $endpointDetail = $null
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $uri = $null
    try { $uri = [Uri]$endpointUrl } catch { }
    $epHost = if ($uri) { $uri.Host } else { 'enrollment.manage.microsoft.com' }
    $epPort = if ($uri -and $uri.Port -gt 0) { $uri.Port } else { 443 }

    # 4a. TCP connect with a bounded timeout (BeginConnect + WaitOne; PS 5.1 has
    # no connect-timeout on the synchronous TcpClient.Connect).
    $tcp = $null
    $tcpOk = $false
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $iar = $tcp.BeginConnect($epHost, $epPort, $null, $null)
        if ($iar.AsyncWaitHandle.WaitOne(5000, $false) -and $tcp.Connected) {
            $tcp.EndConnect($iar)
            $tcpOk = $true
        }
    }
    catch { }
    finally { if ($tcp) { $tcp.Close() } }

    # 4b. HTTP probe (best-effort enrichment). PS 5.1 throws WebException, PS 7
    # HttpResponseException; both expose .Response.StatusCode, so any status back
    # still proves the server answered.
    $httpDetail = $null
    try {
        $resp = Invoke-WebRequest -Uri $endpointUrl -Method Get -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
        $httpDetail = "HTTP $([int]$resp.StatusCode)"
    }
    catch {
        $respObj = Get-Prop $_.Exception 'Response'
        $statusCode = Get-Prop $respObj 'StatusCode'
        if ($respObj -and $null -ne $statusCode) { $httpDetail = "HTTP $([int]$statusCode)" }
        else { $httpDetail = $_.Exception.Message }
    }

    if ($tcpOk) {
        $endpointReachable = $true
        $endpointDetail = "TCP ${epHost}:${epPort} OK; $httpDetail"
    }
    elseif ($httpDetail -match '^HTTP ') {
        # TCP probe failed/blocked but the server still returned an HTTP status.
        $endpointReachable = $true
        $endpointDetail = "TCP ${epHost}:${epPort} failed; $httpDetail"
    }
    else {
        $endpointReachable = $false
        $endpointDetail = "TCP ${epHost}:${epPort} failed; $httpDetail"
    }
    Write-Log ("[4/5] Intune endpoint {0}: Reachable={1} ({2})." -f $endpointUrl, $endpointReachable, $endpointDetail) $(if ($endpointReachable -eq $false) { 'WARN' } else { 'INFO' })

    # --- 5. omadmclient pressure ---------------------------------------------
    $omadm = @(Get-Process -Name 'omadmclient' -ErrorAction SilentlyContinue)
    $omadmCount = $omadm.Count
    $omadmCpu = 0
    if ($omadmCount -gt 0) {
        $omadmCpu = [math]::Round((($omadm | Measure-Object -Property CPU -Sum).Sum), 1)
    }
    Write-Log ("[5/5] omadmclient.exe processes: {0} (total CPU {1}s)." -f $omadmCount, $omadmCpu) $(if ($omadmCount -gt 1) { 'WARN' } else { 'INFO' })

    # --- 6. Verdict -----------------------------------------------------------
    $verdict = 'Unknown'
    $action  = ''
    if (-not $primaryEnroll) {
        $verdict = 'NotEnrolled'
        $action  = 'No MDM enrollment found. Nothing to repair here.'
    }
    elseif (-not $cert) {
        $verdict = 'CertMissing'
        $action  = "No cert issued by '$IssuerMatch' in LocalMachine\My. Re-enroll to mint one."
    }
    elseif ($isExpired) {
        $verdict = 'CertExpired'
        $action  = "MDM device cert expired $([math]::Abs($daysToExpiry)) day(s) ago. Re-enroll (cannot renew past expiry)."
    }
    elseif ($isExpiringSoon) {
        $verdict = 'CertExpiringSoon'
        $action  = "MDM device cert expires in $daysToExpiry day(s). Verify auto-renewal works (connectivity/identity) before it lapses."
    }
    else {
        $verdict = 'Healthy'
        $action  = 'MDM device cert valid and unexpired. No action.'
    }

    $fixable = ($verdict -eq 'CertExpired' -or $verdict -eq 'CertMissing')
    Write-Log ("Verdict: {0} (Fixable={1}). {2}" -f $verdict, $fixable, $action) $(if ($fixable) { 'WARN' } else { 'INFO' })

    [pscustomobject]@{
        ComputerName        = $env:COMPUTERNAME
        Verdict             = $verdict
        Fixable             = $fixable
        RecommendedAction   = $action
        EnrollmentId        = if ($primaryEnroll) { $primaryEnroll.EnrollmentId } else { $null }
        ProviderID          = if ($primaryEnroll) { $primaryEnroll.ProviderID } else { $null }
        UPN                 = if ($primaryEnroll) { $primaryEnroll.UPN } else { $null }
        EnrollmentCount     = $enrollments.Count
        MmpcEnrolled        = [bool]$mmpcEnroll
        MmpcEnrollmentId    = if ($mmpcEnroll) { $mmpcEnroll.EnrollmentId } else { $null }
        MmpcEnrollmentFlag  = $mmpcFlag
        MdmCertThumbprint   = $certThumb
        MdmCertSubject      = $certSubject
        MdmCertNotBefore    = $notBefore
        MdmCertNotAfter     = $notAfter
        DaysToExpiry        = $daysToExpiry
        IsExpired           = $isExpired
        IsExpiringSoon      = $isExpiringSoon
        HasPrivateKey       = $hasPrivKey
        MdmCertCount        = $mdmCerts.Count
        RenewalTaskName     = if ($primaryRenew) { $primaryRenew.TaskName } else { $null }
        RenewalLastRun      = if ($primaryRenew) { $primaryRenew.LastRunTime } else { $null }
        RenewalLastResult   = if ($primaryRenew) { $primaryRenew.LastResultHex } else { $null }
        AnyRenewalFailed    = $renewFailed
        EndpointTested      = $endpointUrl
        EndpointReachable   = $endpointReachable
        EndpointDetail      = $endpointDetail
        OmaDmClientCount    = $omadmCount
        OmaDmCpuSeconds     = $omadmCpu
        AllEnrollments      = $enrollments
        AllMdmCerts         = @($mdmCerts | ForEach-Object { '{0} NotAfter={1:yyyy-MM-dd}' -f $_.Thumbprint, $_.NotAfter })
        AllRenewalTasks     = $renewTasks
        CollectedAt         = $now
    }
}

# --------------------------------------------------------------------------
# REPAIR WORKER  (self-contained; DESTRUCTIVE; gated by the caller)
# --------------------------------------------------------------------------
$RepairWorker = {
    param(
        [string]$IssuerMatch,
        [int]$WaitSeconds,
        [string]$EnrollMode
    )

    $log = New-Object System.Collections.Generic.List[string]
    # Level-tagged logging. Counts are derived from the tags at the end, so we
    # never mutate a shared counter from the nested function (that would shadow
    # to local scope in PS 5.1).
    function Add-Log {
        param(
            [string]$Message,
            [ValidateSet('INFO', 'WARN', 'ERROR')]
            [string]$Level = 'INFO'
        )
        # In-memory copy drives the returned object, the error/warn counts, and
        # the colorized console replay. Write-Log persists each line (CMTrace) to
        # the Intune IME log folder.
        $log.Add(('[{0,-5}] {1}' -f $Level, $Message))
        Write-Log $Message $Level
        if ($Level -eq 'ERROR') { Write-Error $Message -ErrorAction SilentlyContinue }
    }

    $repairRoot = Join-Path $env:ProgramData 'MdmCertRepair'
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'

    $guidRx = '^[0-9A-Fa-f]{8}-([0-9A-Fa-f]{4}-){3}[0-9A-Fa-f]{12}$'
    $success = $false
    $triggerTime = $null
    $enrolled = $null
    $enrollDetail = $null
    $enrollElapsed = 0
    try {
        Add-Log "Repair started on $env:COMPUTERNAME (PID $PID)."
        # Discover MDM enrollment GUID(s) to tear down: the Intune classic channel
        # (MS DM Server) AND the MMP-C / declared-config channel (Microsoft Device
        # Management). A lingering MMP-C enrollment blocks re-enroll with 0x8018000A.
        # WMI_Bridge_SCCM_Server (co-management bridge) is deliberately left alone.
        $realProviders = @('MS DM Server', 'Microsoft Device Management')
        $enrollIds = @()
        $enrollRoot = 'HKLM:\SOFTWARE\Microsoft\Enrollments'
        if (Test-Path $enrollRoot) {
            Get-ChildItem $enrollRoot -ErrorAction SilentlyContinue |
                Where-Object { $_.PSChildName -match $guidRx } |
                ForEach-Object {
                    $p = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
                    $prov = Get-Prop $p 'ProviderID'
                    if ($prov -in $realProviders) {
                        $enrollIds += $_.PSChildName
                        Add-Log ("Target enrollment {0} (ProviderID='{1}')." -f $_.PSChildName, $prov)
                    }
                }
        }
        if ($enrollIds.Count -eq 0) {
            Add-Log 'No MDM enrollment (MS DM Server / Microsoft Device Management) found; will still re-enroll.' 'WARN'
        }

        # 1. Backup --------------------------------------------------------------
        $backupDir = Join-Path $repairRoot ("Backup_{0}" -f $stamp)
        try {
            New-Item -ItemType Directory -Path $backupDir -Force -ErrorAction Stop | Out-Null
            Add-Log "Backup dir: $backupDir"
        }
        catch {
            Add-Log "Could not create backup dir ${backupDir}: $($_.Exception.Message)" 'ERROR'
        }

        $regRoots = @(
            'HKLM\SOFTWARE\Microsoft\Enrollments',
            'HKLM\SOFTWARE\Microsoft\EnterpriseResourceManager\Tracked',
            'HKLM\SOFTWARE\Microsoft\PolicyManager\Providers',
            'HKLM\SOFTWARE\Microsoft\Provisioning\OMADM'
        )
        $exported = 0
        if (Test-Path $backupDir) {
            foreach ($rr in $regRoots) {
                $safe = ($rr -replace '[\\:]', '_')
                $dest = Join-Path $backupDir ("$safe.reg")
                $regOut = & reg.exe export $rr $dest /y 2>&1
                if ($LASTEXITCODE -eq 0) { $exported++ }
                else { Add-Log ("Registry backup failed for ${rr} (reg.exe exit $LASTEXITCODE): $regOut") 'ERROR' }
            }
            $mdmThumbs = @(Get-ChildItem 'Cert:\LocalMachine\My' -ErrorAction SilentlyContinue |
                    Where-Object { $_.Issuer -match $IssuerMatch } |
                    ForEach-Object { '{0}`tNotAfter={1:o}`tSubject={2}' -f $_.Thumbprint, $_.NotAfter, $_.Subject })
            try { $mdmThumbs | Set-Content -Path (Join-Path $backupDir 'mdm-certs.txt') -Encoding UTF8 -ErrorAction Stop }
            catch { Add-Log "Could not write cert manifest: $($_.Exception.Message)" 'WARN' }
            Add-Log ("Backed up {0}/{1} registry root(s) and {2} MDM cert record(s)." -f $exported, $regRoots.Count, $mdmThumbs.Count)
        }
        if ($exported -eq 0) {
            Add-Log 'No registry backup was captured; teardown will proceed WITHOUT a restore point.' 'WARN'
        }

        # 2. Stop omadmclient ----------------------------------------------------
        Get-Process -Name 'omadmclient' -ErrorAction SilentlyContinue | ForEach-Object {
            try { Stop-Process -Id $_.Id -Force -ErrorAction Stop; Add-Log "Stopped omadmclient PID $($_.Id)." }
            catch { Add-Log "Could not stop omadmclient PID $($_.Id): $($_.Exception.Message)" 'WARN' }
        }

        # 3. Remove enrollment artifacts ----------------------------------------
        foreach ($g in $enrollIds) {
            # Tasks + folder live under BOTH the critical (EnterpriseMgmt) and
            # non-critical (EnterpriseMgmtNonCritical) parents for the same
            # enrollment GUID; clean both so no stale tasks point at a dead enrollment.
            foreach ($parentPath in @('\Microsoft\Windows\EnterpriseMgmt', '\Microsoft\Windows\EnterpriseMgmtNonCritical')) {
                $tp = "$parentPath\$g\"
                Get-ScheduledTask -TaskPath $tp -ErrorAction SilentlyContinue | ForEach-Object {
                    try { Unregister-ScheduledTask -TaskName $_.TaskName -TaskPath $tp -Confirm:$false -ErrorAction Stop; Add-Log "Removed task $($_.TaskName)." }
                    catch { Add-Log "Task removal failed ($($_.TaskName)): $($_.Exception.Message)" 'ERROR' }
                }
                # Remove the now-empty task folder via the Schedule.Service COM object.
                try {
                    $svc = New-Object -ComObject Schedule.Service
                    $svc.Connect()
                    $parent = $svc.GetFolder($parentPath)
                    $parent.DeleteFolder($g, 0)
                    Add-Log ("Removed {0}\$g task folder." -f ($parentPath -replace '.*\\'))
                }
                catch { Add-Log "Task folder removal skipped/failed for $parentPath\${g}: $($_.Exception.Message)" 'WARN' }
            }

            # Registry GUID keys. All roots use the DASHED enrollment GUID
            # (verified on a live enrolled Win11 26100 host). The no-dash
            # OMADM\Accounts entry is a DEFENSIVE fallback only - no source
            # confirms any build ever used it - kept because it costs one
            # Test-Path and can't match a non-GUID key.
            $gNoDash = $g -replace '-', ''
            $keys = @(
                "HKLM:\SOFTWARE\Microsoft\Enrollments\$g",
                "HKLM:\SOFTWARE\Microsoft\Enrollments\Status\$g",
                "HKLM:\SOFTWARE\Microsoft\EnterpriseResourceManager\Tracked\$g",
                "HKLM:\SOFTWARE\Microsoft\PolicyManager\AdmxInstalled\$g",
                "HKLM:\SOFTWARE\Microsoft\PolicyManager\Providers\$g",
                "HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\$g",
                "HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\$gNoDash",
                "HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Logger\$g",
                "HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Sessions\$g"
            )
            foreach ($k in $keys) {
                if (Test-Path $k) {
                    try { Remove-Item -Path $k -Recurse -Force -ErrorAction Stop; Add-Log "Removed $k" }
                    catch { Add-Log "Remove failed $k : $($_.Exception.Message)" 'ERROR' }
                }
            }
        }

        # 3b. Clear the MMP-C enrollment block flag -----------------------------
        # A value of 2 independently blocks auto-enrollment (Bad request 400 /
        # 0x80190190) even after the enrollment keys are gone.
        $enrRootKey = 'HKLM:\SOFTWARE\Microsoft\Enrollments'
        $mmpcFlag = Get-Prop (Get-ItemProperty $enrRootKey -ErrorAction SilentlyContinue) 'MmpcEnrollmentFlag'
        if ($null -ne $mmpcFlag) {
            if ($mmpcFlag -ne 0) {
                try { Remove-ItemProperty -Path $enrRootKey -Name 'MmpcEnrollmentFlag' -Force -ErrorAction Stop; Add-Log "Cleared MmpcEnrollmentFlag (was $mmpcFlag)." }
                catch { Add-Log "Could not clear MmpcEnrollmentFlag (was $mmpcFlag): $($_.Exception.Message)" 'WARN' }
            }
            else {
                Add-Log 'MmpcEnrollmentFlag present and 0 (not blocking).'
            }
        }

        # 4. Remove expired Intune MDM device cert(s) ---------------------------
        $now = Get-Date
        Get-ChildItem 'Cert:\LocalMachine\My' -ErrorAction SilentlyContinue |
            Where-Object { $_.Issuer -match $IssuerMatch -and $_.NotAfter -lt $now } |
            ForEach-Object {
                try { Remove-Item -Path $_.PSPath -Force -ErrorAction Stop; Add-Log "Removed expired cert $($_.Thumbprint)." }
                catch { Add-Log "Cert removal failed $($_.Thumbprint): $($_.Exception.Message)" 'ERROR' }
            }

        # 5. Re-enroll headlessly -----------------------------------------------
        # EnrollMode Device = AAD device credential (co-mgmt / AVD multi-session);
        # User = AAD user credential (user-driven Entra-joined; needs a user token,
        # so must run in the user's session).
        $de = Join-Path $env:windir 'System32\deviceenroller.exe'
        if (Test-Path $de) {
            if ($EnrollMode -eq 'User') {
                $deArgs = @('/c', '/AutoEnrollMDM')
                if (([System.Security.Principal.WindowsIdentity]::GetCurrent()).IsSystem) {
                    Add-Log 'EnrollMode=User but running as SYSTEM - the user-credential path cannot obtain a user token here and will likely fail (0xCAA8xxxx). Run in the logged-on user session.' 'WARN'
                }
            }
            else {
                $deArgs = @('/c', '/AutoEnrollMDMUsingAADDeviceCredential')
            }
            Add-Log ("Launching (EnrollMode=$EnrollMode): deviceenroller.exe " + ($deArgs -join ' '))
            $triggerTime = Get-Date
            # deviceenroller.exe /c does not reliably exit once enrollment is
            # triggered - the user-credential path can block indefinitely - so a
            # blocking -Wait hangs here. Wait a bounded time; if it is still
            # running, treat the trigger as accepted and let the event-log poll
            # below report the real outcome (killing it could abort enrollment).
            $triggerWaitSec = 60
            $p = Start-Process -FilePath $de -ArgumentList $deArgs -PassThru -WindowStyle Hidden
            if ($p.WaitForExit($triggerWaitSec * 1000)) {
                if ($p.ExitCode -eq 0) {
                    Add-Log "deviceenroller exit code: 0 (re-enroll triggered)."
                    $success = $true
                }
                else {
                    Add-Log ("deviceenroller exit code: {0} - re-enroll trigger failed." -f (Get-EnrollErrorText $p.ExitCode)) 'ERROR'
                }
            }
            else {
                Add-Log ("deviceenroller.exe still running after {0}s - trigger accepted; continuing to the enrollment poll for the real result (not killing it)." -f $triggerWaitSec) 'WARN'
                $success = $true
            }
        }
        else {
            Add-Log "deviceenroller.exe not found at $de" 'ERROR'
        }

        # 6. Confirm enrollment completed (async) -------------------------------
        # deviceenroller returns immediately; enrollment runs in the background.
        # Authoritative signals from the DeviceManagement provider: event 75
        # (Auto MDM Enroll: Device Credential, Succeeded) and 76 (Failed). Events
        # 77-80 (retry / 'DMGetAadDeviceToken Failure (Access is denied.)') are
        # transient warnings, NOT a verdict. Also accept the Intune MDM device
        # cert reappearing as proof. Query by ProviderName so the channel
        # (Admin vs Enrollment) doesn't matter across builds.
        if ($success -and $WaitSeconds -gt 0 -and $triggerTime) {
            $provName = 'Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider'
            $deadline = $triggerTime.AddSeconds($WaitSeconds)
            Add-Log ("Waiting up to {0}s for enrollment to complete (DM event 75/76 + cert store)..." -f $WaitSeconds)
            while ((Get-Date) -lt $deadline) {
                Start-Sleep -Seconds 10
                $newCert = Get-ChildItem 'Cert:\LocalMachine\My' -ErrorAction SilentlyContinue |
                    Where-Object { $_.Issuer -match $IssuerMatch -and $_.NotAfter -gt (Get-Date) } |
                    Sort-Object NotBefore -Descending | Select-Object -First 1
                $ev = $null
                try {
                    $ev = Get-WinEvent -FilterHashtable @{ ProviderName = $provName; Id = 75, 76; StartTime = $triggerTime } -ErrorAction SilentlyContinue |
                        Sort-Object TimeCreated | Select-Object -Last 1
                }
                catch { }
                if ($ev -and $ev.Id -eq 76) {
                    $enrolled = $false
                    $line76 = ($ev.Message -split "`r?`n")[0]
                    if ($ev.Message -match '0x[0-9A-Fa-f]{8}') { $line76 += ' [' + (Get-EnrollErrorText $Matches[0]) + ']' }
                    $enrollDetail = 'Event 76 (Auto MDM Enroll Failed): ' + $line76
                    break
                }
                if (($ev -and $ev.Id -eq 75) -or $newCert) {
                    $enrolled = $true
                    if ($newCert) { $enrollDetail = "New MDM device cert $($newCert.Thumbprint) (NotAfter $($newCert.NotAfter.ToString('yyyy-MM-dd')))." }
                    else { $enrollDetail = "Event 75 (Auto MDM Enroll: Succeeded) at $($ev.TimeCreated)." }
                    break
                }
            }
            $enrollElapsed = [int]((Get-Date) - $triggerTime).TotalSeconds
            if ($enrolled -eq $true) { Add-Log ("Enrollment CONFIRMED after {0}s. {1}" -f $enrollElapsed, $enrollDetail) }
            elseif ($enrolled -eq $false) { Add-Log ("Enrollment FAILED after {0}s. {1}" -f $enrollElapsed, $enrollDetail) 'ERROR' }
            else { Add-Log ("Enrollment not confirmed within {0}s; it may still complete in the background - re-audit shortly. (Transient 77-80 warnings are not fatal.)" -f $WaitSeconds) 'WARN' }
        }
    }
    catch {
        Add-Log "FATAL: $($_.Exception.Message)" 'ERROR'
        Add-Log $_.ScriptStackTrace 'ERROR'
    }

    $errCount  = @($log | Where-Object { $_ -match '^\[ERROR\]' }).Count
    $warnCount = @($log | Where-Object { $_ -match '^\[WARN ' }).Count

    [pscustomobject]@{
        ComputerName     = $env:COMPUTERNAME
        Success          = $success
        Enrolled         = $enrolled
        EnrollDetail     = $enrollDetail
        EnrollElapsedSec = $enrollElapsed
        WaitSeconds      = $WaitSeconds
        ErrorCount       = $errCount
        WarnCount        = $warnCount
        LogPath          = $Script:LogFile
        Log              = $log.ToArray()
    }
}

# --------------------------------------------------------------------------
# DISPATCH 
# --------------------------------------------------------------------------
Write-Verbose "Auditing $env:COMPUTERNAME ..."
Write-Log ("=== Repair-IntuneMdmCert start (Mode={0}, User={1}\{2}) ===" -f $Mode, $env:USERDOMAIN, $env:USERNAME)
$result = & $AuditWorker $IssuerMatch $ExpiryWarningDays

# Human-readable summary to the host; the raw object still goes down the pipeline.
Write-Host ''
Write-Host '==== Intune MDM device-cert audit ====' -ForegroundColor Cyan
$result |
    Select-Object ComputerName, Verdict, DaysToExpiry, IsExpired, HasPrivateKey, EndpointReachable, OmaDmClientCount, MdmCertThumbprint |
    Format-Table -AutoSize | Out-Host

if ($result.RecommendedAction) {
    $color = 'Gray'
    if ($result.Verdict -eq 'CertExpired' -or $result.Verdict -eq 'CertMissing') { $color = 'Red' }
    elseif ($result.Verdict -eq 'CertExpiringSoon') { $color = 'Yellow' }
    elseif ($result.Verdict -eq 'Healthy') { $color = 'Green' }
    Write-Host ("  {0}: {1}" -f $result.ComputerName, $result.RecommendedAction) -ForegroundColor $color
}
Write-Host ''

# --------------------------------------------------------------------------
# REPAIR (gated)
# --------------------------------------------------------------------------
if ($Mode -eq 'Repair') {
    $t = $result.ComputerName
    # -RepairHealthy widens eligibility to a valid cert (deliberate rotation);
    # every other gate below is unchanged.
    $eligible = $result.Fixable -or ($RepairHealthy -and ($result.Verdict -eq 'Healthy' -or $result.Verdict -eq 'CertExpiringSoon'))
    if (-not $eligible) {
        Write-Host ("Skip: not fixable (Verdict={0})." -f $result.Verdict) -ForegroundColor DarkGray
    }
    elseif ($result.EndpointReachable -eq $false -and -not $Force) {
        Write-Warning ("Intune endpoint {0} is unreachable ({1}). Re-enrollment will likely fail (proxy / SSL inspection / network). Fix connectivity, or re-run with -Force to proceed anyway." -f $result.EndpointTested, $result.EndpointDetail)
    }
    else {
        if ($RepairHealthy -and -not $result.Fixable) {
            Write-Warning ("-RepairHealthy: forcing teardown + re-enroll on a host whose cert is currently VALID (Verdict={0})." -f $result.Verdict)
        }
        $what = "Tear down MDM enrollment + remove expired Intune cert + re-enroll ($EnrollMode credential)"
        # High-impact ShouldProcess prompts by default; suppress non-interactively
        # with -Confirm:$false. -Force does NOT affect the prompt (gate only).
        if ($PSCmdlet.ShouldProcess($t, $what)) {
            Write-Host "Repairing $t ..." -ForegroundColor Cyan
            $rr = & $RepairWorker $IssuerMatch $WaitSeconds $EnrollMode
            $teardownNote = if ($rr.ErrorCount -gt 0) { " ({0} teardown error(s))" -f $rr.ErrorCount } else { "" }
            if (-not $rr.Success) {
                Write-Host ("  Repair FAILED ({0} error(s), {1} warning(s)) - review the log." -f $rr.ErrorCount, $rr.WarnCount) -ForegroundColor Red
            }
            elseif ($rr.Enrolled -eq $true) {
                Write-Host ("  Re-enroll CONFIRMED in {0}s{1}. {2}" -f $rr.EnrollElapsedSec, $teardownNote, $rr.EnrollDetail) -ForegroundColor Green
            }
            elseif ($rr.Enrolled -eq $false) {
                Write-Host ("  Re-enroll triggered but FAILED{0}. {1}" -f $teardownNote, $rr.EnrollDetail) -ForegroundColor Red
            }
            elseif ($rr.WaitSeconds -le 0) {
                Write-Host ("  Re-enroll triggered{0} (no wait requested). Re-audit later to confirm the new cert." -f $teardownNote) -ForegroundColor Yellow
            }
            else {
                Write-Host ("  Re-enroll triggered{0} but NOT confirmed within {1}s. Re-audit shortly." -f $teardownNote, $rr.WaitSeconds) -ForegroundColor Yellow
            }
            if ($rr.LogPath) { Write-Host ("  Log: {0}" -f $rr.LogPath) -ForegroundColor DarkGray }
            Write-Host ''
            Write-Host 'NOTE: re-enrollment mints a NEW cert off the device Entra identity. If an SSL-inspecting proxy sits in the path for Intune/Entra in device context, the new cert will hit the same renewal failure and expire again.' -ForegroundColor Yellow
        }
    }
}

# Emit the raw audit object for downstream use (Export-Csv, Where-Object, etc.).
$result
