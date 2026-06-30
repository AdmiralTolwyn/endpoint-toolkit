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
      4. Removes the enrollment artifacts (scheduled tasks + registry GUID keys).
      5. Removes the expired Intune MDM device cert(s) from LocalMachine\My.
      6. Re-enrolls headlessly via:
            deviceenroller.exe /c /AutoEnrollMDMUsingAADDeviceCredential
      7. Re-audits and reports.

.PARAMETER Mode
    Audit (default, read-only) or Repair (destructive, gated).

.PARAMETER ExpiryWarningDays
    Days-to-expiry threshold that flags a cert as "expiring soon". Default 30.

.PARAMETER IssuerMatch
    Regex matched against the cert Issuer to locate the MDM device cert.
    Default 'Microsoft Intune.*MDM Device CA' (matches both the production
    'Microsoft Intune MDM Device CA' and the 'Microsoft Intune Beta MDM Device CA'
    issuers; does not match the unrelated '...Device Management' enrollment cert).

.PARAMETER Force
    Skip the interactive confirmation in Repair mode, and proceed even when the
    Intune enrollment endpoint is unreachable.

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
    .\Repair-IntuneMdmCert.ps1 -Mode Repair -Force
    Repair the local host (skips confirmation and the endpoint-reachability gate).

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
    [switch]$Force
)

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
                if ($p -and ($p.DiscoveryServiceFullURL -or $p.UPN -or $p.ProviderID)) {
                    $enrollments += [pscustomobject]@{
                        EnrollmentId           = $_.PSChildName
                        ProviderID             = $p.ProviderID
                        UPN                    = $p.UPN
                        EnrollmentState        = $p.EnrollmentState
                        DiscoveryUrl           = $p.DiscoveryServiceFullURL
                        SslClientCertReference = $p.SslClientCertReference
                    }
                }
            }
    }
    # Prefer the Intune ('MS DM Server') enrollment as primary.
    $primaryEnroll = $enrollments | Where-Object { $_.ProviderID -eq 'MS DM Server' } | Select-Object -First 1
    if (-not $primaryEnroll) { $primaryEnroll = $enrollments | Select-Object -First 1 }
    Write-Verbose ("[1/5] Enrollments: found {0} (providers: {1})." -f $enrollments.Count, (($enrollments | ForEach-Object { $_.ProviderID }) -join ', '))
    if ($primaryEnroll) {
        Write-Verbose ("      Primary enrollment: {0} (ProviderID='{1}', UPN='{2}', State={3})." -f $primaryEnroll.EnrollmentId, $primaryEnroll.ProviderID, $primaryEnroll.UPN, $primaryEnroll.EnrollmentState)
    }
    else {
        Write-Verbose '      No MDM enrollment found.'
    }

    # --- 2. MDM device cert(s) by Issuer -------------------------------------
    $mdmCerts = @()
    Get-ChildItem 'Cert:\LocalMachine\My' -ErrorAction SilentlyContinue |
        Where-Object { $_.Issuer -match $IssuerMatch } |
        ForEach-Object { $mdmCerts += $_ }
    # Current cert = the one with the latest NotAfter.
    $cert = $mdmCerts | Sort-Object NotAfter -Descending | Select-Object -First 1
    Write-Verbose ("[2/5] MDM certs matching /{0}/ in LocalMachine\My: {1}." -f $IssuerMatch, $mdmCerts.Count)
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
        Write-Verbose ("      Selected cert {0}: NotAfter={1:yyyy-MM-dd} ({2}d), Expired={3}, HasPrivateKey={4}." -f `
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
                $taskArgs = $t.Actions[0].Arguments
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
    Write-Verbose ("[3/5] Scheduled tasks under EnterpriseMgmt: {0} (renewal-classified: {1})." -f $renewTasks.Count, @($renewTasks | Where-Object { $_.IsRenewal }).Count)
    foreach ($rt in ($renewTasks | Where-Object { $_.IsRenewal })) {
        Write-Verbose ("      - {0}: LastResult={1}, LastRun={2}, Next={3}" -f $rt.TaskName, $rt.LastResultHex, $rt.LastRunTime, $rt.NextRunTime)
    }
    if ($primaryRenew) { Write-Verbose ("      Primary renewal task: '{0}' (AnyRenewalFailed={1})." -f $primaryRenew.TaskName, $renewFailed) }

    # --- 4. Intune enrollment endpoint reachability --------------------------
    $endpointUrl = if ($primaryEnroll -and $primaryEnroll.DiscoveryUrl) { $primaryEnroll.DiscoveryUrl }
                   else { 'https://enrollment.manage.microsoft.com/EnrollmentServer/Discovery.svc' }
    $endpointReachable = $null
    $endpointDetail = $null
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $resp = Invoke-WebRequest -Uri $endpointUrl -Method Head -UseBasicParsing -TimeoutSec 8 -ErrorAction Stop
        $endpointReachable = $true
        $endpointDetail = "HTTP $([int]$resp.StatusCode)"
    }
    catch {
        # A non-2xx HTTP status still proves we reached the server. PS 5.1 throws
        # System.Net.WebException, PS 7 throws HttpResponseException; both expose
        # a .Response with a .StatusCode, so inspect it generically.
        $respObj = $_.Exception.Response
        if ($respObj -and $null -ne $respObj.StatusCode) {
            $endpointReachable = $true
            $endpointDetail = "HTTP $([int]$respObj.StatusCode)"
        }
        else {
            $endpointReachable = $false
            $endpointDetail = $_.Exception.Message
        }
    }
    Write-Verbose ("[4/5] Intune endpoint {0}: Reachable={1} ({2})." -f $endpointUrl, $endpointReachable, $endpointDetail)

    # --- 5. omadmclient pressure ---------------------------------------------
    $omadm = @(Get-Process -Name 'omadmclient' -ErrorAction SilentlyContinue)
    $omadmCount = $omadm.Count
    $omadmCpu = 0
    if ($omadmCount -gt 0) {
        $omadmCpu = [math]::Round((($omadm | Measure-Object -Property CPU -Sum).Sum), 1)
    }
    Write-Verbose ("[5/5] omadmclient.exe processes: {0} (total CPU {1}s)." -f $omadmCount, $omadmCpu)

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
    Write-Verbose ("Verdict: {0} (Fixable={1}). {2}" -f $verdict, $fixable, $action)

    [pscustomobject]@{
        ComputerName        = $env:COMPUTERNAME
        Verdict             = $verdict
        Fixable             = $fixable
        RecommendedAction   = $action
        EnrollmentId        = if ($primaryEnroll) { $primaryEnroll.EnrollmentId } else { $null }
        ProviderID          = if ($primaryEnroll) { $primaryEnroll.ProviderID } else { $null }
        UPN                 = if ($primaryEnroll) { $primaryEnroll.UPN } else { $null }
        EnrollmentCount     = $enrollments.Count
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
        [string]$IssuerMatch
    )

    $log = New-Object System.Collections.Generic.List[string]
    function Add-Log { param([string]$m) $log.Add(('[{0:HH:mm:ss}] {1}' -f (Get-Date), $m)) }

    $guidRx = '^[0-9A-Fa-f]{8}-([0-9A-Fa-f]{4}-){3}[0-9A-Fa-f]{12}$'
    $success = $false
    try {
        # Discover Intune enrollment GUID(s).
        $enrollIds = @()
        $enrollRoot = 'HKLM:\SOFTWARE\Microsoft\Enrollments'
        if (Test-Path $enrollRoot) {
            Get-ChildItem $enrollRoot -ErrorAction SilentlyContinue |
                Where-Object { $_.PSChildName -match $guidRx } |
                ForEach-Object {
                    $p = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
                    if ($p -and $p.ProviderID -eq 'MS DM Server') { $enrollIds += $_.PSChildName }
                }
        }
        if ($enrollIds.Count -eq 0) {
            Add-Log 'No Intune (MS DM Server) enrollment found; will still re-enroll.'
        }
        else {
            Add-Log ("Intune enrollment GUID(s): " + ($enrollIds -join ', '))
        }

        # 1. Backup --------------------------------------------------------------
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $backupDir = Join-Path $env:ProgramData ("MdmCertRepair\Backup_{0}" -f $stamp)
        New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
        Add-Log "Backup dir: $backupDir"

        $regRoots = @(
            'HKLM\SOFTWARE\Microsoft\Enrollments',
            'HKLM\SOFTWARE\Microsoft\EnterpriseResourceManager\Tracked',
            'HKLM\SOFTWARE\Microsoft\PolicyManager\Providers',
            'HKLM\SOFTWARE\Microsoft\Provisioning\OMADM'
        )
        foreach ($rr in $regRoots) {
            $safe = ($rr -replace '[\\:]', '_')
            $dest = Join-Path $backupDir ("$safe.reg")
            & reg.exe export $rr $dest /y > $null 2>&1
        }
        $mdmThumbs = @(Get-ChildItem 'Cert:\LocalMachine\My' -ErrorAction SilentlyContinue |
                Where-Object { $_.Issuer -match $IssuerMatch } |
                ForEach-Object { '{0}`tNotAfter={1:o}`tSubject={2}' -f $_.Thumbprint, $_.NotAfter, $_.Subject })
        $mdmThumbs | Set-Content -Path (Join-Path $backupDir 'mdm-certs.txt') -Encoding UTF8
        Add-Log ("Backed up {0} registry root(s) and {1} MDM cert record(s)." -f $regRoots.Count, $mdmThumbs.Count)

        # 2. Stop omadmclient ----------------------------------------------------
        Get-Process -Name 'omadmclient' -ErrorAction SilentlyContinue | ForEach-Object {
            try { Stop-Process -Id $_.Id -Force -ErrorAction Stop; Add-Log "Stopped omadmclient PID $($_.Id)." }
            catch { Add-Log "Could not stop omadmclient PID $($_.Id): $($_.Exception.Message)" }
        }

        # 3. Remove enrollment artifacts ----------------------------------------
        foreach ($g in $enrollIds) {
            # Scheduled tasks under the enrollment folder.
            $tp = "\Microsoft\Windows\EnterpriseMgmt\$g\"
            Get-ScheduledTask -TaskPath $tp -ErrorAction SilentlyContinue | ForEach-Object {
                try { Unregister-ScheduledTask -TaskName $_.TaskName -TaskPath $tp -Confirm:$false -ErrorAction Stop; Add-Log "Removed task $($_.TaskName)." }
                catch { Add-Log "Task removal failed ($($_.TaskName)): $($_.Exception.Message)" }
            }
            # Remove the now-empty task folder via the Schedule.Service COM object.
            try {
                $svc = New-Object -ComObject Schedule.Service
                $svc.Connect()
                $parent = $svc.GetFolder('\Microsoft\Windows\EnterpriseMgmt')
                $parent.DeleteFolder($g, 0)
                Add-Log "Removed EnterpriseMgmt\$g task folder."
            }
            catch { Add-Log "Task folder removal skipped/failed for ${g}: $($_.Exception.Message)" }

            # Registry GUID keys. All roots use the DASHED enrollment GUID on
            # current builds (verified on an enrolled host) and match call4cloud's
            # dashed assumption. OMADM\Accounts on some older builds used a no-dash
            # GUID, so target both forms there (each is Test-Path-guarded below).
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
                    catch { Add-Log "Remove failed $k : $($_.Exception.Message)" }
                }
            }
        }

        # 4. Remove expired Intune MDM device cert(s) ---------------------------
        $now = Get-Date
        Get-ChildItem 'Cert:\LocalMachine\My' -ErrorAction SilentlyContinue |
            Where-Object { $_.Issuer -match $IssuerMatch -and $_.NotAfter -lt $now } |
            ForEach-Object {
                try { Remove-Item -Path $_.PSPath -Force -ErrorAction Stop; Add-Log "Removed expired cert $($_.Thumbprint)." }
                catch { Add-Log "Cert removal failed $($_.Thumbprint): $($_.Exception.Message)" }
            }

        # 5. Re-enroll headlessly via device credential -------------------------
        $de = Join-Path $env:windir 'System32\deviceenroller.exe'
        if (Test-Path $de) {
            Add-Log "Launching: deviceenroller.exe /c /AutoEnrollMDMUsingAADDeviceCredential"
            $p = Start-Process -FilePath $de -ArgumentList '/c', '/AutoEnrollMDMUsingAADDeviceCredential' -Wait -PassThru -WindowStyle Hidden
            Add-Log "deviceenroller exit code: $($p.ExitCode)"
            $success = ($p.ExitCode -eq 0)
        }
        else {
            Add-Log "deviceenroller.exe not found at $de"
        }
    }
    catch {
        Add-Log "FATAL: $($_.Exception.Message)"
    }

    [pscustomobject]@{
        ComputerName = $env:COMPUTERNAME
        Success      = $success
        Log          = $log.ToArray()
    }
}

# --------------------------------------------------------------------------
# DISPATCH 
# --------------------------------------------------------------------------
Write-Verbose "Auditing $env:COMPUTERNAME ..."
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
    if (-not $result.Fixable) {
        Write-Host ("Skip: not fixable (Verdict={0})." -f $result.Verdict) -ForegroundColor DarkGray
    }
    elseif ($result.EndpointReachable -eq $false -and -not $Force) {
        Write-Warning ("Intune endpoint {0} is unreachable ({1}). Re-enrollment will likely fail (proxy / SSL inspection / network). Fix connectivity, or re-run with -Force to proceed anyway." -f $result.EndpointTested, $result.EndpointDetail)
    }
    else {
        $what = "Tear down MDM enrollment + remove expired Intune cert + re-enroll via device credential"
        if ($PSCmdlet.ShouldProcess($t, $what)) {
            $proceed = $true
            if (-not $Force) {
                $ans = Read-Host "CONFIRM repair of this host? This removes the enrollment and re-enrolls. [y/N]"
                if ($ans -ne 'y' -and $ans -ne 'Y') {
                    Write-Host "Skipped (no confirmation)." -ForegroundColor DarkGray
                    $proceed = $false
                }
            }

            if ($proceed) {
                Write-Host "Repairing $t ..." -ForegroundColor Cyan
                $rr = & $RepairWorker $IssuerMatch
                $rr.Log | ForEach-Object { Write-Host ("    {0}" -f $_) -ForegroundColor DarkGray }
                if ($rr.Success) { Write-Host "  Re-enroll triggered OK. Re-check the cert in a few minutes / after a reboot." -ForegroundColor Green }
                else { Write-Host "  Repair completed with warnings (see log)." -ForegroundColor Yellow }
                Write-Host ''
                Write-Host 'NOTE: re-enrollment mints a NEW cert off the device Entra identity. If an SSL-inspecting proxy sits in the path for Intune/Entra in device context, the new cert will hit the same renewal failure and expire again.' -ForegroundColor Yellow
            }
        }
    }
}

# Emit the raw audit object for downstream use (Export-Csv, Where-Object, etc.).
$result
