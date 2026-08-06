<#
.SYNOPSIS
    Tears down a single stale or orphaned MDM enrollment (Intune 'MS DM Server'
    or Windows declared configuration 'Microsoft Device Management') so the
    device can re-enrol against the tenant it is actually joined to.

.DESCRIPTION
    Companion to Get-StaleMdmEnrollment.ps1. That script decides; this one acts.
    Run the detector first and feed it the EnrollmentId it reports.

    THIS IS AN UNSUPPORTED PROCEDURE. Neither the registry teardown below nor
    any part of it is documented or supported by Microsoft, and a device
    repaired this way is in a state Microsoft has not sanctioned. Use it when a
    reset is not viable, and be prepared to reset anyway if it does not take.

    THE SUPPORTED REMEDIATION IS A DEVICE RESET. Every documented way to remove
    an enrollment is server-initiated and needs a working enrollment to arrive
    through - Intune Retire/Wipe, and the DMClient CSP Unenroll and
    LinkedEnrollment/Unenroll Exec nodes. Once the primary enrollment is gone
    the server cannot reach the device at all, so the only supported route left
    is to wipe and rebuild: Autopilot Reset, Fresh Start, or a plain Windows
    reset / reimage, followed by a clean enrollment.

    Microsoft's supported teardown for a declared-configuration (WinDC / MMP-C)
    enrollment is specifically the DMClient CSP LinkedEnrollment/Unenroll Exec
    node, which is server-initiated only. It is
    not projected into the WMI Bridge - verified by enumerating all 465 classes
    in root\cimv2\mdm\dmmap as NT AUTHORITY\SYSTEM on Windows 11 26100: no
    MDM_DMClient* class exists, and nothing matching 'declared' or 'linked'.
    There is therefore no local, supported way to invoke it. When the primary
    MDM enrollment has already been retired away, the server can no longer
    reach the device to send the Exec either, so the enrollment cannot be
    removed by any documented route. Manual purge or reimage are the only
    remaining options; this script is the manual purge.

    For completeness: a local trigger for the MMP-C ENROLL direction does exist
    (deviceenroller.exe /c /EnrollMmpc, community-reported by call4cloud, not
    documented by Microsoft). There is no counterpart for unenroll. No
    Microsoft page documents any deviceenroller.exe switch at all, and a search
    for a '/DisenrollDevice' switch returns zero results in any source - it is
    not real, and a scan of the binary's strings on build 26100 finds no
    'Disenroll' token.

    NOT REMOVED BY THIS SCRIPT, but worth checking first:
    HKLM\SOFTWARE\Microsoft\Enrollments!MmpcEnrollmentFlag. A value of 2 is
    independently reported (Quest support KBs, a Microsoft techcommunity
    thread, the fleetdm issue tracker) to leave auto-enrollment failing with
    'Bad request (400)' / 0x80190190 until it is set to 0 or deleted. It is
    undocumented by Microsoft and is a machine-wide value rather than
    per-enrollment, so it is out of scope for a single-GUID purge. The detector
    reports it and raises the MmpcEnrollmentFlagBlocking flag; clear it as a
    separate, deliberate step.

    WHAT IT REMOVES, for one enrollment GUID:
      - Scheduled tasks under \Microsoft\Windows\EnterpriseMgmt\<guid>\ and the
        task folder itself (the folder needs the Schedule.Service COM object -
        there is no Unregister-ScheduledTaskFolder cmdlet)
      - Eight registry roots, all keyed by the DASHED GUID:
          SOFTWARE\Microsoft\Enrollments\<guid>
          SOFTWARE\Microsoft\Enrollments\Status\<guid>
          SOFTWARE\Microsoft\EnterpriseResourceManager\Tracked\<guid>
          SOFTWARE\Microsoft\PolicyManager\AdmxInstalled\<guid>
          SOFTWARE\Microsoft\PolicyManager\Providers\<guid>
          SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\<guid>
          SOFTWARE\Microsoft\Provisioning\OMADM\Logger\<guid>
          SOFTWARE\Microsoft\Provisioning\OMADM\Sessions\<guid>
        Verified on a live enrolled Windows 11 26100 host: ALL EIGHT roots use
        the dashed form, including OMADM\Accounts, for both the Intune and the
        WinDC enrollment. The undashed variant was probed on the same host and
        matched nothing anywhere.

        The undashed fallback in the code is DEFENSIVE ONLY. No source -
        Microsoft or community - was found asserting that any build ever used
        an undashed GUID here; an earlier version of this comment claimed older
        builds did, and that claim was unsourced and has been withdrawn. Keep
        the fallback because it costs one Test-Path and cannot delete anything
        that is not GUID-named, not because it is known to be needed.

        Provenance of the eight-key list: it traces to a single 2020 community
        blog post - Maxime Rastello, 'Manually re-enroll a co-managed or Hybrid
        Azure AD Join Windows 10 PC to Microsoft Intune without loosing current
        configuration', https://www.maximerastello.com/manually-re-enroll-a-co-managed-or-hybrid-azure-ad-join-windows-10-pc-to-microsoft-intune-without-loosing-current-configuration/
        - which every later script and article copies (steve-prentice's
        Remove-IntuneCurrentEnrollment.ps1 and ztrhgf's Reset-IntuneEnrollment.ps1
        both credit it by name). It is community consensus, not eight
        independent confirmations.

        MICROSOFT DOES NOT DOCUMENT OR SUPPORT THIS TEARDOWN. Their enrollment
        diagnostics article names only the first key and gives a heuristic
        rather than a list; PolicyManager, EnterpriseResourceManager and OMADM
        are never mentioned in a teardown context anywhere on Learn. A device
        repaired this way is in a state Microsoft has not sanctioned. The
        supported remediation is a device reset - see the .NOTES section.

        PolicyManager\AdmxInstalled and PolicyManager\Providers legitimately do
        not exist for a WinDC enrollment - absent is normal there, not an error.

        ORDERING WARNING: Enrollments\<guid> is deleted recursively, which takes
        the LinkedEnrollment subkey with it. That subkey is the only explicit
        pointer from the Intune enrollment to its MMP-C child. If both channels
        need purging, record LinkedEnrollmentId (or purge the MMP-C GUID) FIRST
        - once the Intune key is gone the child can only be re-identified by
        ProviderID, which is weaker.
      - The MDM client certificate in Cert:\LocalMachine\My named by the
        enrollment's DMPCertThumbPrint (skip with -SkipCertificate)

    SAFETY MODEL:
      - Dry run BY DEFAULT. Nothing is touched without -Execute.
      - Refuses to touch a key whose ProviderID is not one of the two real MDM
        channels, so the ~30 internal CSP-provider subkeys under Enrollments
        cannot be destroyed by a mistyped GUID.
      - Refuses to remove a HEALTHY enrollment - one whose AADTenantID matches
        the tenant the device is currently joined to - unless -Force is given.
        That is the guard against running this on the wrong machine.
      - Backs up every registry root (reg.exe export), every task definition
        (XML) and the certificate (public part, .cer) to -BackupPath before
        deleting anything, plus a manifest.json describing the enrollment.

    AFTERWARDS the device holds no MDM enrollment at all. It will not re-enrol
    on its own until either the Entra auto-enrollment MDM scope applies at the
    next PRT refresh, or -Reenroll triggers deviceenroller.exe directly.

    KNOWN FAILURE MODE worth checking before you run this: if the device's
    auto-enrollment path authenticates against a federated IdP belonging to the
    OLD tenant, re-enrollment will fail after a perfectly clean purge. Confirm
    the new tenant's enrollment path works on a comparable device first.

.PARAMETER EnrollmentId
    The enrollment GUID to remove, as reported by Get-StaleMdmEnrollment.ps1 in
    ActiveEnrollmentId or MmpcActiveId. Dashed form.

.PARAMETER BackupPath
    Directory to write the pre-deletion backup into. A per-run subfolder named
    <guid>_<timestamp> is created inside it. Defaults under ProgramData so it
    survives user profile removal.

.PARAMETER Execute
    Actually perform the removal. Without this the script only reports what it
    would do and exits - safe to hand to a customer for a first pass.

.PARAMETER SkipCertificate
    Leave the MDM client certificate in place. Use when the same certificate is
    referenced by a second enrollment you are keeping.

.PARAMETER Force
    Permit removal of an enrollment whose tenant matches the joined tenant, i.e.
    one that looks healthy. Required only for deliberate full re-enrollment.

.PARAMETER Reenroll
    After a successful purge, trigger deviceenroller.exe /c /AutoEnrollMDM.
    Ignored on a dry run. Verified switch; the enrollment attempt is still
    subject to the tenant's auto-enrollment configuration succeeding.

.OUTPUTS
    A [pscustomobject] summarising the run: the enrollment's identity, the
    backup location, and per-item Removed / NotPresent / Failed status for every
    registry root, the task folder and the certificate. Non-empty Failed means a
    partial purge - see NOTES.

.EXAMPLE
    .\Remove-StaleMdmEnrollment.ps1 -EnrollmentId 1B66378F-821F-4D1A-B842-6802BFAD3A85

    Dry run. Prints the full plan, changes nothing.

.EXAMPLE
    .\Remove-StaleMdmEnrollment.ps1 -EnrollmentId 1B66378F-821F-4D1A-B842-6802BFAD3A85 -Execute

    Backs up, then purges. Prompts before each destructive step.

.EXAMPLE
    .\Remove-StaleMdmEnrollment.ps1 -EnrollmentId 1B66378F-... -Execute -Confirm:$false -Reenroll

    Unattended purge and immediate re-enrollment attempt, for use from a
    ConfigMgr / RMM script once the procedure has been validated on a pilot.

.NOTES
    SUPPORTED ALTERNATIVE, prefer it where viable: reset the device. Autopilot
    Reset, Fresh Start, or a plain Windows reset / reimage clears all enrollment
    state and lets the device enrol cleanly. That is the only Microsoft-endorsed
    route once the primary enrollment is gone, because every documented removal
    mechanism is server-initiated and needs a working enrollment to arrive
    through. Use this script only when a reset is not viable.

    Requires: Windows PowerShell 5.1, elevated. SYSTEM is preferable - some
    subkeys under EnterpriseResourceManager\Tracked and PolicyManager carry ACLs
    that deny Administrators write access, and a delete under those will land in
    Failed. If that happens, re-run as SYSTEM (PsExec -s -i) before resorting to
    taking ownership; taking ownership of policy keys has its own side effects.

    A partial purge is worse than none - Windows can still consider itself
    enrolled off a single surviving root. Always check the returned Failed list
    is empty, and re-run Get-StaleMdmEnrollment.ps1 afterwards to confirm the
    verdict has moved to NotEnrolled.

    Reboot after a successful purge before evaluating re-enrollment. The OMA-DM
    client caches enrollment state in memory.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9A-Fa-f]{8}-([0-9A-Fa-f]{4}-){3}[0-9A-Fa-f]{12}$')]
    [string]$EnrollmentId,

    [string]$BackupPath = (Join-Path $env:ProgramData 'MdmEnrollmentPurge'),

    [switch]$Execute,
    [switch]$SkipCertificate,
    [switch]$Force,
    [switch]$Reenroll
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$TaskRoot     = '\Microsoft\Windows\EnterpriseMgmt'
$RealProviders = @('MS DM Server', 'Microsoft Device Management')

#region Preflight

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
if (-not (New-Object Security.Principal.WindowsPrincipal($identity)).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Must run elevated. The Enrollments hive is not writable by a standard user.'
}

$enrollKey = "HKLM:\SOFTWARE\Microsoft\Enrollments\$EnrollmentId"
if (-not (Test-Path $enrollKey)) {
    throw ("Enrollment {0} not found at {1}. Run Get-StaleMdmEnrollment.ps1 to list the real enrollment GUIDs." -f $EnrollmentId, $enrollKey)
}

$props      = Get-ItemProperty -Path $enrollKey
$providerId = $null
$tenantId   = $null
$upn        = $null
$certThumb  = $null
if ($props.PSObject.Properties['ProviderID'])        { $providerId = [string]$props.ProviderID }
if ($props.PSObject.Properties['AADTenantID'])       { $tenantId   = [string]$props.AADTenantID }
if ($props.PSObject.Properties['UPN'])               { $upn        = [string]$props.UPN }
if ($props.PSObject.Properties['DMPCertThumbPrint']) { $certThumb  = [string]$props.DMPCertThumbPrint }

# Guards against a mistyped GUID hitting an internal CSP-provider subkey, which
# shares the same GUID-shaped naming but carries no ProviderID.
if ($RealProviders -notcontains $providerId) {
    throw ("Key {0} has ProviderID '{1}', which is not a real MDM channel. Refusing to remove it." -f $EnrollmentId, $providerId)
}

$joinedTenant = $null
try {
    $dsreg = @(& dsregcmd.exe /status 2>$null)
    $m = $dsreg | Select-String -Pattern '^\s*TenantId\s*:\s*(\S+)' | Select-Object -First 1
    if ($m) { $joinedTenant = $m.Matches[0].Groups[1].Value }
}
catch { }

# String -eq is case-insensitive in PowerShell, which absorbs the GUID casing
# difference between dsregcmd's output and the registry. Do not tighten to -ceq.
$tenantMatches = [bool]($tenantId -and $joinedTenant -and $tenantId -eq $joinedTenant)
if ($tenantMatches -and -not $Force) {
    throw ("Enrollment {0} is bound to tenant {1}, which IS the tenant this device is joined to - it looks healthy. Re-run with -Force only if you intend a full re-enrollment." -f $EnrollmentId, $tenantId)
}

Write-Host ''
Write-Host ('Enrollment  : {0}' -f $EnrollmentId)
Write-Host ('ProviderID  : {0}' -f $providerId)
Write-Host ('Tenant      : {0}{1}' -f $tenantId, $(if ($tenantMatches) { '  (MATCHES joined tenant)' } else { '  (does NOT match joined tenant)' }))
Write-Host ('Joined      : {0}' -f $joinedTenant)
Write-Host ('UPN         : {0}' -f $upn)
Write-Host ('Cert        : {0}' -f $(if ($certThumb) { $certThumb } else { '<none>' }))
Write-Host ''

#endregion Preflight

#region Target enumeration

# Dashed GUID everywhere on current builds; the undashed OMADM\Accounts variant
# is a fallback for builds that predate the change.
$undashed = $EnrollmentId -replace '-', ''
$regTargets = @(
    "HKLM:\SOFTWARE\Microsoft\Enrollments\$EnrollmentId"
    "HKLM:\SOFTWARE\Microsoft\Enrollments\Status\$EnrollmentId"
    "HKLM:\SOFTWARE\Microsoft\EnterpriseResourceManager\Tracked\$EnrollmentId"
    "HKLM:\SOFTWARE\Microsoft\PolicyManager\AdmxInstalled\$EnrollmentId"
    "HKLM:\SOFTWARE\Microsoft\PolicyManager\Providers\$EnrollmentId"
    "HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\$EnrollmentId"
    "HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\$undashed"
    "HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Logger\$EnrollmentId"
    "HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Sessions\$EnrollmentId"
)

$tasks = @()
try {
    $tasks = @(Get-ScheduledTask -TaskPath ("{0}\{1}\" -f $TaskRoot, $EnrollmentId) -ErrorAction Stop)
}
catch { }

$cert = $null
if ($certThumb -and -not $SkipCertificate) {
    $cert = Get-ChildItem -Path 'Cert:\LocalMachine\My' -ErrorAction SilentlyContinue |
        Where-Object { $_.Thumbprint -eq $certThumb } | Select-Object -First 1
}

Write-Host 'Planned removals:'
foreach ($p in $regTargets) {
    Write-Host ('  [{0}] {1}' -f $(if (Test-Path $p) { 'x' } else { ' ' }), $p)
}
Write-Host ('  [{0}] {1}\{2}\  ({3} task(s))' -f $(if ($tasks.Count) { 'x' } else { ' ' }), $TaskRoot, $EnrollmentId, $tasks.Count)
Write-Host ('  [{0}] certificate {1}' -f $(if ($cert) { 'x' } else { ' ' }), $(if ($certThumb) { $certThumb } else { '<none>' }))
Write-Host ''

if (-not $Execute) {
    Write-Host 'DRY RUN - nothing was changed. Re-run with -Execute to apply.' -ForegroundColor Yellow
    return
}

#endregion Target enumeration

#region Backup

# Taken before anything is deleted and never cleaned up automatically. reg.exe
# export is used rather than a PowerShell serialisation because the resulting
# .reg files can be re-imported by double-click during an incident, without this
# script being present.
$stamp   = (Get-Date).ToString('yyyyMMdd-HHmmss')
$destDir = Join-Path $BackupPath ('{0}_{1}' -f $EnrollmentId, $stamp)
New-Item -Path $destDir -ItemType Directory -Force | Out-Null

@{
    EnrollmentId = $EnrollmentId
    ProviderId   = $providerId
    TenantId     = $tenantId
    JoinedTenant = $joinedTenant
    Upn          = $upn
    CertThumb    = $certThumb
    ComputerName = $env:COMPUTERNAME
    CapturedUtc  = (Get-Date).ToUniversalTime().ToString('o')
} | ConvertTo-Json | Set-Content -Path (Join-Path $destDir 'manifest.json') -Encoding UTF8

foreach ($p in $regTargets) {
    if (-not (Test-Path $p)) { continue }
    $native = $p -replace '^HKLM:\\', 'HKLM\'
    $file   = Join-Path $destDir (($native -replace '[\\:]', '_') + '.reg')
    & reg.exe export $native $file /y | Out-Null
}

foreach ($t in $tasks) {
    $xml = Export-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath
    Set-Content -Path (Join-Path $destDir ('task_' + $t.TaskName + '.xml')) -Value $xml -Encoding UTF8
}

if ($cert) {
    [System.IO.File]::WriteAllBytes(
        (Join-Path $destDir ('cert_' + $cert.Thumbprint + '.cer')),
        $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))
}

Write-Host ('Backup written to {0}' -f $destDir)
Write-Host ''

#endregion Backup

#region Removal

$results = New-Object System.Collections.ArrayList

function Add-Result {
    <#
        Records the outcome of one removal step. Every target produces exactly
        one row - including NotPresent ones - so the caller can prove a root was
        considered rather than silently skipped. [void] on .Add() suppresses the
        insertion index that ArrayList returns.
    #>
    param([string]$Item, [string]$Status, [string]$Detail)
    [void]$results.Add([pscustomobject]@{ Item = $Item; Status = $Status; Detail = $Detail })
}

# Tasks first: an OMA-DM sync firing mid-purge can rewrite keys just deleted.
if ($tasks.Count -gt 0) {
    if ($PSCmdlet.ShouldProcess(("{0}\{1}\ ({2} tasks)" -f $TaskRoot, $EnrollmentId, $tasks.Count), 'Unregister scheduled tasks')) {
        foreach ($t in $tasks) {
            try {
                Unregister-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath -Confirm:$false -ErrorAction Stop
                Add-Result -Item ('Task: ' + $t.TaskName) -Status 'Removed' -Detail ''
            }
            catch {
                Add-Result -Item ('Task: ' + $t.TaskName) -Status 'Failed' -Detail $_.Exception.Message
            }
        }
        # No cmdlet deletes a task FOLDER; the COM scheduler is the only route.
        try {
            $svc = New-Object -ComObject 'Schedule.Service'
            $svc.Connect()
            $svc.GetFolder($TaskRoot).DeleteFolder($EnrollmentId, 0)
            Add-Result -Item 'Task folder' -Status 'Removed' -Detail ($TaskRoot + '\' + $EnrollmentId)
        }
        catch {
            Add-Result -Item 'Task folder' -Status 'Failed' -Detail $_.Exception.Message
        }
    }
}
else {
    Add-Result -Item 'Task folder' -Status 'NotPresent' -Detail ($TaskRoot + '\' + $EnrollmentId)
}

foreach ($p in $regTargets) {
    if (-not (Test-Path $p)) {
        Add-Result -Item $p -Status 'NotPresent' -Detail ''
        continue
    }
    if ($PSCmdlet.ShouldProcess($p, 'Remove registry key')) {
        try {
            Remove-Item -Path $p -Recurse -Force -ErrorAction Stop
            Add-Result -Item $p -Status 'Removed' -Detail ''
        }
        catch {
            Add-Result -Item $p -Status 'Failed' -Detail $_.Exception.Message
        }
    }
}

if ($cert) {
    if ($PSCmdlet.ShouldProcess($cert.Thumbprint, 'Remove MDM client certificate')) {
        try {
            Remove-Item -Path ('Cert:\LocalMachine\My\' + $cert.Thumbprint) -Force -ErrorAction Stop
            Add-Result -Item ('Cert: ' + $cert.Thumbprint) -Status 'Removed' -Detail ''
        }
        catch {
            Add-Result -Item ('Cert: ' + $cert.Thumbprint) -Status 'Failed' -Detail $_.Exception.Message
        }
    }
}
else {
    Add-Result -Item 'Cert' -Status 'NotPresent' -Detail $(if ($SkipCertificate) { 'skipped by -SkipCertificate' } else { '' })
}

#endregion Removal

#region Re-enrollment

$reenrollStatus = 'NotRequested'
$failed = @($results | Where-Object { $_.Status -eq 'Failed' })

if ($Reenroll) {
    if ($failed.Count -gt 0) {
        # A surviving root can still make Windows believe it is enrolled, in
        # which case deviceenroller exits silently and hides the real problem.
        $reenrollStatus = 'SkippedDueToFailures'
    }
    elseif ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, 'Trigger deviceenroller.exe /c /AutoEnrollMDM')) {
        try {
            & "$env:SystemRoot\System32\deviceenroller.exe" /c /AutoEnrollMDM
            $reenrollStatus = 'Triggered'
        }
        catch {
            $reenrollStatus = 'Failed: ' + $_.Exception.Message
        }
    }
}

#endregion Re-enrollment

#region Output

$results | Format-Table -AutoSize Item, Status, Detail | Out-Host

if ($failed.Count -gt 0) {
    Write-Warning ("{0} item(s) could not be removed - this is a PARTIAL purge. Re-run as SYSTEM before doing anything else." -f $failed.Count)
}
else {
    Write-Host 'Purge complete. Reboot, then re-run Get-StaleMdmEnrollment.ps1 to confirm NotEnrolled.' -ForegroundColor Green
}

[pscustomobject]@{
    ComputerName    = $env:COMPUTERNAME
    EnrollmentId    = $EnrollmentId
    ProviderId      = $providerId
    EnrollmentTenant = $tenantId
    JoinedTenant    = $joinedTenant
    BackupPath      = $destDir
    RemovedCount    = @($results | Where-Object { $_.Status -eq 'Removed' }).Count
    FailedCount     = $failed.Count
    Results         = $results.ToArray()
    Reenrollment    = $reenrollStatus
    CompletedUtc    = (Get-Date).ToUniversalTime().ToString('o')
}

#endregion Output
