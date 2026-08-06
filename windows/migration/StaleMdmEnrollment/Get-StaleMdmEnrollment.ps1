<#
.SYNOPSIS
    Detects an MDM (Intune) enrollment that points at a different Entra tenant
    than the one the device is currently joined to.

.DESCRIPTION
    Written for tenant-to-tenant migrations (Continental AG -> Aumovio) where a
    hybrid-joined device ends up in a split state: the Entra device object and
    PRT belong to the NEW tenant, but the MDM enrollment, the MDM client
    certificate and the OMA-DM sync channel still belong to the OLD tenant. Such
    a device keeps checking in to the old tenant's Intune and remains fully
    manageable there, while appearing unmanaged in the new tenant.

    GROUND TRUTH (verified against a live enrolled Windows 11 host):

      Joined tenant   HKLM\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\
                      JoinInfo\<certThumbprint>\TenantId
                      (cross-checked against `dsregcmd /status` -> TenantId,
                      which wins because the registry can retain a leftover
                      JoinInfo key from the pre-migration tenant)

      Enrolled tenant HKLM\SOFTWARE\Microsoft\Enrollments\<enrollmentGuid>\
                      AADTenantID, on the subkey whose ProviderID is
                      'MS DM Server' (that ProviderID is what distinguishes a
                      real Intune enrollment from the ~30 other Enrollments
                      subkeys Windows creates for internal CSP providers)

    SECOND CHANNEL - 'Microsoft Device Management'. Windows can hold a separate,
    fully live enrollment on the MMP-C channel (config source
    MicrosoftManagementPlatformCloud, discovery discovery.dm.microsoft.com,
    check-in checkin.dm.microsoft.com, EnrollmentType 26 / 0x1A). It carries the
    same footprint as a classic enrollment - Enrollments key, OMA-DM account,
    client cert off the Intune Device Management Device CA, and a complete
    EnterpriseMgmt\<guid>\ task set.

      Sourcing: 'MicrosoftManagementPlatformCloud' and 'Enroll Type: (0x1A)'
      both appear verbatim in Microsoft's declared-configuration
      troubleshooting examples, so the string and the type code are documented.
      The dm.microsoft.com / checkin.dm.microsoft.com hostnames above are the
      PRODUCTION ring, observed on the affected device. Hosts on the preview
      ring carry discovery.dm-beta.microsoft.com and checkin.dm-beta.microsoft.com
      instead - match on ProviderID, not on the hostname.

    UNVERIFIED CLAIM, FLAGGED DELIBERATELY: that this enrollment survives an
    Intune Retire while the 'MS DM Server' enrollment is removed. That is our
    own field observation on one migrated device. No Microsoft documentation and
    no community write-up states it; Microsoft's Retire article says only that
    the device is unenrolled from Intune and says nothing about the linked
    enrollment either way. Treat it as a working hypothesis, not established
    behaviour. What IS documented is that the dual enrollment is only permitted
    while a primary MDM enrollment exists, so a WinDC enrollment outliving its
    parent is an unsupported state by design.

    A device left in that state appears unmanaged in Intune while Windows still
    believes it is MDM enrolled, so auto-enrollment and deviceenroller.exe exit
    silently with no events.

    CRITICALLY, the MMP-C enrollment carries its OWN AADTenantID, and it is not
    necessarily the same as the Intune channel's. Measured on a migrated device:
    the WinDC enrollment survived the tenant move still bound to the OLD tenant,
    with a UPN on the old domain and IsFederated = 1 pointing at the old ADFS,
    while the device was Entra joined to the new tenant. Once the Intune
    enrollment has been retired away there is no 'MS DM Server' key left to
    compare, so a check that looks only at that channel reports 'not enrolled'
    and misses a live cross-tenant enrollment completely. Both channels are
    therefore evaluated against the baseline independently.

    A channel's tenant differing from the joined tenant is the whole test.
    Everything else this script collects is corroboration or triage context.

    SECONDARY SIGNAL: the enrollment's UPN value is frequently stored as
    'user@domain.com@<tenantGuid>'. That appended GUID is an independent read on
    the enrolling tenant and is used as a fallback when AADTenantID is absent.

    LIVENESS SIGNALS - these separate an enrollment that is still actively
    syncing to the old tenant from inert registry litter left by a failed or
    partial cleanup:
      - Scheduled tasks under \Microsoft\Windows\EnterpriseMgmt\<enrollmentGuid>\
        and their most recent LastRunTime (SyncTaskCount / SyncTaskLastRun)
      - HKLM\SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\<enrollmentGuid>
      - Presence and expiry of the MDM client certificate named by
        DMPCertThumbPrint

    They are combined into LivenessScore, and the highest-scoring enrollment is
    marked IsActive. ONLY the active enrollment drives the verdict; the rest are
    counted as OrphanEnrollmentCount, so a half-finished cleanup cannot produce a
    false 'stale' verdict on an otherwise healthy device.

    WEAK SIGNAL - dsregcmd's MdmUrl. Every Intune tenant uses the same enrollment
    discovery URL, so DiscoveryServiceFullURL matching it cannot distinguish two
    Intune enrollments belonging to different tenants. It only rules out an
    enrollment pointing at a non-Intune MDM, so it carries the lowest weight in
    LivenessScore and is never used on its own.

    DELIBERATELY NOT USED - the MDM certificate's Subject CN. It is the Intune
    device ID, NOT the Entra device ID, and the two differ on perfectly healthy
    machines (measured: CN=1742e0d5-... vs dsregcmd DeviceId=4a71d882-...).
    Comparing them would flag the entire fleet as stale.

    CORRECTION (previous versions of this comment were wrong): the certificate
    is located via DMPCertThumbPrint under Enrollments\<guid>, NOT via
    SslClientCertReference. SslClientCertReference does not live under the
    Enrollments key at all - it lives under
    Provisioning\OMADM\Accounts\<guid> in the form 'MY;System;<thumbprint>',
    and on a healthy host it is populated and its thumbprint equals
    DMPCertThumbPrint. DMPCertThumbPrint is preferred simply because it is on
    the key already being read, not because the other value is unreliable.

    THIS SCRIPT IS STRICTLY READ-ONLY. It creates, modifies and deletes nothing.
    Remediation (enrollment purge + re-enrollment against the new tenant) is a
    separate, destructive operation - see scripts\EntraCutover\ and
    scripts\Repair-IntuneMdmCert.ps1.

.PARAMETER ExpectedTenantId
    The tenant this device SHOULD be managed by. When omitted, defaults to the
    tenant the device is actually Entra-joined to, which is the correct baseline
    after a migration and means you never have to hardcode the new tenant GUID.
    Supply it explicitly only to test a hypothetical, or to audit devices that
    are not joined anywhere.

.PARAMETER StaleTenantId
    Optional GUID of the known-bad (old) tenant, e.g. Continental AG's
    8d4b558f-7b2e-40ba-ad1f-e04d79e6265a. Purely a reporting refinement: when a
    stale enrollment matches it the verdict is 'StaleKnownTenant' instead of the
    generic 'StaleOtherTenant', which lets you separate migration fallout from
    unrelated cross-tenant enrollments in the same export.

.PARAMETER Json
    Emit the result as compressed JSON instead of a PowerShell object. Use this
    when collecting via ConfigMgr / a GPO startup script / an RMM, so the output
    can be dropped into a file share and aggregated later.

.OUTPUTS
    Default: a [pscustomobject] with these properties -

      ComputerName          NetBIOS name of the device
      Verdict               see the verdict table below
      Flags                 array of every condition that applied, independent of
                            the single Verdict string - IntuneCrossTenant,
                            MmpcCrossTenant, MmpcOrphan, MmpcUnlinked,
                            MmpcEnrollmentFlagBlocking, OrphanResidue,
                            MultipleJoinInfo. Several can be true at once, which
                            is why Verdict alone is not sufficient for triage.
      Reason                one-sentence human explanation of the verdict
      JoinedTenantId        tenant the device is Entra-joined to
      ExpectedTenantId      baseline actually used for the comparison
      JoinedUserEmail       UserEmail from CloudDomainJoin (the joining identity)
      EntraDeviceId         dsregcmd DeviceId in the joined tenant
      AzureAdJoined         dsregcmd AzureAdJoined (YES/NO)
      DomainJoined          dsregcmd DomainJoined (YES/NO) - both YES = hybrid
      MdmUrl                enrollment discovery URL advertised by dsregcmd
      JoinInfoKeyCount      >1 means a leftover JoinInfo key from the old tenant
      EnrollmentCount       number of 'MS DM Server' enrollments found
      ActiveEnrollmentId    GUID of the live enrollment, $null if all are inert
      ActiveTenantId        tenant that live Intune enrollment belongs to
      OrphanEnrollmentCount enrollments that are not the live one - cleanup
                            residue, visible but not verdict-driving
      StaleEnrollmentCount  1 when the ACTIVE enrollment is cross-tenant, else 0
      MmpcEnrollmentCount   number of 'Microsoft Device Management' enrollments
      MmpcActiveId          GUID of the live MMP-C enrollment, $null if none
      MmpcActiveTenantId    tenant that live MMP-C enrollment belongs to
      MmpcTenantMismatch    $true when that tenant is not the baseline
      MmpcActiveUpn         its UPN - the domain corroborates the tenant read
      MmpcActiveLastRun     most recent EnterpriseMgmt task run for that GUID
      MmpcEnrollmentFlag    raw HKLM\SOFTWARE\Microsoft\Enrollments!MmpcEnrollmentFlag.
                            2 is the reported enrollment-blocking state, 0 is
                            healthy. Undocumented by Microsoft - see the note in
                            the Evaluation region. $null when the value is absent.
      LinkedEnrollmentId    GUID the ACTIVE Intune enrollment explicitly points
                            at as its MMP-C child, from
                            Enrollments\<guid>\LinkedEnrollment. This is a hard
                            pointer, not an inference.
      LinkedEnrollStatus      raw EnrollStatus value (0-8)
      LinkedEnrollStatusText  that value decoded using Microsoft's published
                            DMClient CSP mapping (3 = Enrollment Failed,
                            4 = Enrollment Succeeded, 8 = UnEnrollment Succeeded)
      LinkedLastError       raw LastError from the same key, 0 when clean
      LinkedMmpcLocked      raw MMPCLocked value. Reported, NOT interpreted -
                            its meaning is undocumented and unknown.
      LinkedDiscoveryEndpoint  the MMP-C discovery URL the linked enrollment
                            targets. dm-beta hostnames indicate a preview ring.
      Enrollments           array of per-enrollment detail objects, each also
                            carrying ProviderId, Channel, LivenessScore (0-9),
                            IsActive and DiscoveryUrlMatch
      CollectedUtc          ISO-8601 collection timestamp

    Verdicts:
      Clean             active enrollment's tenant == joined tenant. Orphans may
                        still be present - check OrphanEnrollmentCount.
      StaleKnownTenant  active enrollment is cross-tenant AND matches
                        -StaleTenantId. Migration fallout.
      StaleOtherTenant  active enrollment is cross-tenant, but not the tenant you
                        named. Investigate.
      NotEnrolled       no live enrollment - either no 'MS DM Server' key at all,
                        or nothing but inert residue. Not stale, but not managed
                        anywhere either, so it still needs attention.
      MmpcOrphan        no live 'MS DM Server' enrollment, but a live MMP-C one
                        in the CORRECT tenant. Windows counts as enrolled, so
                        re-enrollment silently no-ops. The MMP-C enrollment has
                        to be torn down first.
      MmpcOrphanCrossTenant
                        as MmpcOrphan, but the surviving MMP-C enrollment belongs
                        to a DIFFERENT tenant than the device is joined to -
                        migration residue that was never torn down. Worse than
                        MmpcOrphan: renewal authenticates against the old
                        tenant's IdP and can never succeed.
      Unknown           no baseline to compare against (device not Entra joined
                        and no -ExpectedTenantId), or the active enrollment
                        exposes no tenant at all.

.EXAMPLE
    .\Get-StaleMdmEnrollment.ps1

    Interactive triage on a single machine. Baseline is inferred from the
    device's own join state.

.EXAMPLE
    .\Get-StaleMdmEnrollment.ps1 | Select-Object -ExpandProperty Enrollments |
        Format-List EnrollmentId, EnrollmentTenantId, Upn, SyncTaskLastRun

    Drill into the enrollment detail - most useful when EnrollmentCount > 1.

.EXAMPLE
    .\Get-StaleMdmEnrollment.ps1 -StaleTenantId 8d4b558f-7b2e-40ba-ad1f-e04d79e6265a

    Continental -> Aumovio triage. A StaleKnownTenant verdict means the device is
    still enrolled in Continental.

.EXAMPLE
    .\Get-StaleMdmEnrollment.ps1 -Json | Set-Content "\\fileserver\collect\$env:COMPUTERNAME.json" -Encoding UTF8

    Fleet collection via a GPO startup script or ConfigMgr package.

.NOTES
    Requires: Windows PowerShell 5.1. Must run as SYSTEM or elevated - the
    values under HKLM\SOFTWARE\Microsoft\Enrollments are not readable by a
    standard user, and without them the script cannot distinguish 'Clean' from
    'NotEnrolled'.

    PS 5.1 constraint worth remembering: registry keys returned by Get-ChildItem
    on a HKLM:\ path expose no LastWriteTime property, so JoinInfo subkeys cannot
    be ordered by recency. That is why dsregcmd is the authoritative source for
    the current join rather than the registry.

    Verified: the AADTenantID value name, the ProviderID = 'MS DM Server'
    discriminator, the 'user@domain@<tenantGuid>' UPN form, and the MDM cert CN
    NOT matching the Entra device ID - all confirmed on a live enrolled host.
    Also verified on a retired host: a ProviderID = 'Microsoft Device Management'
    enrollment surviving alone, syncing successfully, with its full task set.
    Inferred, not verified: that this MMP-C enrollment is what makes
    deviceenroller.exe /c /AutoEnrollMDM exit without logging. The silence plus
    the live enrollment is the measurement; the precondition check is not
    documented by Microsoft.
    Inferred, not verified: that a device can hold more than one 'MS DM Server'
    enrollment simultaneously. Windows keeps a single ACTIVE enrollment, but
    orphaned keys do survive a failed or partial unenrollment - that is why the
    teardown procedure has to clear eight registry roots by hand. The script
    therefore scores liveness and ranks, rather than assuming one enrollment.
#>
[CmdletBinding()]
param(
    [string]$ExpectedTenantId,
    [string]$StaleTenantId,
    [switch]$Json
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# Matches only well-formed GUID key names, so the ~30 non-GUID / partial subkeys
# under Enrollments are skipped before any property read is attempted.
$GuidRx = '^[0-9A-Fa-f]{8}-([0-9A-Fa-f]{4}-){3}[0-9A-Fa-f]{12}$'

#region Helpers

function Get-RegValue {
    <#
        Reads one registry value, returning $null for a missing key, a missing
        value, or an access denial. Get-ItemProperty -Name throws on a missing
        value even with -ErrorAction SilentlyContinue under StrictMode, hence the
        property-existence test instead.
    #>
    param([string]$Path, [string]$Name)
    try {
        $item = Get-ItemProperty -Path $Path -ErrorAction Stop
        if ($item.PSObject.Properties[$Name]) { return $item.$Name }
    }
    catch { }
    return $null
}

function Get-DsregField {
    <#
        Pulls a single 'Name : Value' field out of captured dsregcmd /status
        output. Only the first match is taken - dsregcmd repeats field names
        across its Device / Tenant / User / Workplace sections, and the first
        occurrence is always the device-state one we want.
    #>
    param([string[]]$Lines, [string]$Name)
    $m = $Lines | Select-String -Pattern ("^\s*{0}\s*:\s*(.+?)\s*$" -f [regex]::Escape($Name)) | Select-Object -First 1
    if ($m) { return $m.Matches[0].Groups[1].Value }
    return $null
}

#endregion Helpers

#region Collectors

function Get-JoinInfo {
    <#
        Establishes which tenant the device is CURRENTLY joined to - the left
        side of the comparison.

        dsregcmd is treated as authoritative and the registry only fills in the
        fields dsregcmd does not print (UserEmail, IdpDomain). A migrated device
        can retain a JoinInfo subkey from the old tenant, and because PS 5.1
        exposes no LastWriteTime on registry keys there is no way to tell which
        subkey is newer - so any subkey whose TenantId disagrees with dsregcmd is
        skipped. JoinInfoKeys is still reported: a count above 1 is itself a
        migration-residue indicator worth surfacing.
    #>
    $result = [ordered]@{
        TenantId      = $null
        UserEmail     = $null
        IdpDomain     = $null
        DeviceId      = $null
        AzureAdJoined = $null
        DomainJoined  = $null
        MdmUrl        = $null
        JoinInfoKeys  = 0
    }

    try {
        $dsreg = @(& dsregcmd.exe /status 2>$null)
        $result.TenantId      = Get-DsregField -Lines $dsreg -Name 'TenantId'
        $result.DeviceId      = Get-DsregField -Lines $dsreg -Name 'DeviceId'
        $result.AzureAdJoined = Get-DsregField -Lines $dsreg -Name 'AzureAdJoined'
        $result.DomainJoined  = Get-DsregField -Lines $dsreg -Name 'DomainJoined'
        $result.MdmUrl        = Get-DsregField -Lines $dsreg -Name 'MdmUrl'
    }
    catch { }

    $root = 'HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\JoinInfo'
    if (Test-Path $root) {
        $keys = @(Get-ChildItem -Path $root -ErrorAction SilentlyContinue)
        $result.JoinInfoKeys = $keys.Count
        foreach ($k in $keys) {
            $tid = [string](Get-RegValue -Path $k.PSPath -Name 'TenantId')
            if (-not $tid) { continue }
            # Ignore residue keys that disagree with the live dsregcmd join.
            if ($result.TenantId -and $tid -ne $result.TenantId) { continue }
            if (-not $result.TenantId) { $result.TenantId = $tid }
            $result.UserEmail = [string](Get-RegValue -Path $k.PSPath -Name 'UserEmail')
            $result.IdpDomain = [string](Get-RegValue -Path $k.PSPath -Name 'IdpDomain')
            break
        }
    }
    return $result
}

function Get-IntuneEnrollment {
    <#
        Establishes which tenant the device is ENROLLED in - the right side of
        the comparison - plus the liveness evidence needed to argue the old
        enrollment is still functioning.

        Windows creates roughly thirty subkeys under Enrollments for internal CSP
        providers; the ProviderID is what isolates a real cloud enrollment from
        them. Two are real: 'MS DM Server' (classic Intune) and
        'Microsoft Device Management' (MMP-C). Both are collected and tagged via
        Channel, because an orphaned MMP-C enrollment is precisely what makes a
        retired device look unenrolled while refusing to re-enroll.

        Returns an array so a device holding multiple enrollments is reported in
        full rather than silently truncated to the first hit, and scores each one
        so the caller can tell the live enrollment from cleanup residue.
    #>
    param([string]$MdmUrl)

    $found = @()
    $root = 'HKLM:\SOFTWARE\Microsoft\Enrollments'
    if (-not (Test-Path $root)) { return $found }

    foreach ($k in @(Get-ChildItem -Path $root -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -match $GuidRx })) {
        $provider = [string](Get-RegValue -Path $k.PSPath -Name 'ProviderID')
        switch ($provider) {
            'MS DM Server'                { $channel = 'IntuneMdm' }
            'Microsoft Device Management' { $channel = 'MmpC' }
            default                       { $channel = $null }
        }
        if (-not $channel) { continue }

        $id    = $k.PSChildName
        $upn   = [string](Get-RegValue -Path $k.PSPath -Name 'UPN')
        $thumb = [string](Get-RegValue -Path $k.PSPath -Name 'DMPCertThumbPrint')

        # UPN is often stored as user@domain@<tenantGuid>; that suffix is a
        # second, independent read on which tenant enrolled the device, used as
        # the fallback when AADTenantID is missing.
        $upnTenant = $null
        if ($upn -and $upn -match '@([0-9A-Fa-f]{8}-([0-9A-Fa-f]{4}-){3}[0-9A-Fa-f]{12})\s*$') {
            $upnTenant = $Matches[1]
        }

        # DMPCertThumbPrint is on the key already being read, so the cert is
        # resolved from it directly. (The equivalent OMADM\Accounts value,
        # SslClientCertReference = 'MY;System;<thumbprint>', carries the same
        # thumbprint but needs a second key read and a parse.)
        # An absent or expired cert means the old channel is degrading.
        $cert = $null
        if ($thumb) {
            $cert = Get-ChildItem -Path 'Cert:\LocalMachine\My' -ErrorAction SilentlyContinue |
                Where-Object { $_.Thumbprint -eq $thumb } | Select-Object -First 1
        }

        $omadmPath = "HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\$id"

        # Liveness: a recent LastRunTime here is direct evidence the device is
        # still syncing to whichever tenant owns this enrollment. Wrapped because
        # the task folder is absent on a partially torn-down enrollment.
        $taskLastRun = $null
        $taskCount   = 0
        try {
            $tasks = @(Get-ScheduledTask -TaskPath "\Microsoft\Windows\EnterpriseMgmt\$id\" -ErrorAction Stop)
            $taskCount = $tasks.Count
            # Year 1999 filter drops the 30/12/1899 sentinel for never-run tasks.
            $lastRun = $tasks | ForEach-Object { (Get-ScheduledTaskInfo -InputObject $_ -ErrorAction SilentlyContinue).LastRunTime } |
                Where-Object { $_ -and $_.Year -gt 1999 } | Sort-Object -Descending | Select-Object -First 1
            if ($lastRun) { $taskLastRun = $lastRun }
        }
        catch { }

        $discoveryUrl = [string](Get-RegValue -Path $k.PSPath -Name 'DiscoveryServiceFullURL')
        $urlMatch = $false
        if ($MdmUrl -and $discoveryUrl) {
            $urlMatch = ($discoveryUrl.TrimEnd('/') -eq $MdmUrl.TrimEnd('/'))
        }

        # OMA-DM account weighs most - it is what the sync engine actually binds
        # to. The URL match weighs least: it is identical across Intune tenants.
        $score = 0
        if (Test-Path $omadmPath) { $score += 3 }
        if ($taskCount -gt 0)     { $score += 2 }
        if ($cert)                { $score += 2 }
        if ($cert -and $cert.NotAfter -gt (Get-Date)) { $score += 1 }
        if ($urlMatch)            { $score += 1 }

        $found += [pscustomobject]@{
            EnrollmentId       = $id
            ProviderId         = $provider
            Channel            = $channel
            EnrollmentTenantId = [string](Get-RegValue -Path $k.PSPath -Name 'AADTenantID')
            UpnTenantId        = $upnTenant
            Upn                = $upn
            EnrollmentState    = Get-RegValue -Path $k.PSPath -Name 'EnrollmentState'
            EnrollmentType     = Get-RegValue -Path $k.PSPath -Name 'EnrollmentType'
            DiscoveryUrl       = $discoveryUrl
            DiscoveryUrlMatch  = $urlMatch
            AadResourceId      = [string](Get-RegValue -Path $k.PSPath -Name 'AADResourceID')
            CertThumbprint     = $thumb
            CertSubject        = $(if ($cert) { $cert.Subject } else { $null })
            CertIssuer         = $(if ($cert) { $cert.Issuer } else { $null })
            CertNotAfter       = $(if ($cert) { $cert.NotAfter } else { $null })
            CertPresent        = [bool]$cert
            OmaDmAccountExists = (Test-Path $omadmPath)
            SyncTaskCount      = $taskCount
            SyncTaskLastRun    = $taskLastRun
            LivenessScore      = $score
            IsActive           = $false
        }
    }

    # Ranked per channel: the two are independent, and an MMP-C enrollment must
    # never outrank and mask a live Intune one. Highest liveness wins, most
    # recent sync breaks the tie. A top score of 0 means all candidates are dead.
    foreach ($chan in 'IntuneMdm', 'MmpC') {
        $best = $found | Where-Object { $_.Channel -eq $chan -and $_.LivenessScore -gt 0 } |
            Sort-Object -Property @{ Expression = { $_.LivenessScore }; Descending = $true },
                                  @{ Expression = { $_.SyncTaskLastRun }; Descending = $true } |
            Select-Object -First 1
        if ($best) { $best.IsActive = $true }
    }

    return $found
}

function Get-EnrollmentTenant {
    <#
        Resolves the tenant of a single enrollment: AADTenantID first, the GUID
        suffix parsed out of UPN as fallback. $null means undeterminable, and the
        caller must not treat that as a mismatch.
    #>
    param([Parameter(Mandatory)]$Enrollment)
    if ($Enrollment.EnrollmentTenantId) { return $Enrollment.EnrollmentTenantId }
    return $Enrollment.UpnTenantId
}

#endregion Collectors

#region Evaluation

$join        = Get-JoinInfo
$allFound    = @(Get-IntuneEnrollment -MdmUrl $join.MdmUrl)
$enrollments = @($allFound | Where-Object { $_.Channel -eq 'IntuneMdm' })
$mmpc        = @($allFound | Where-Object { $_.Channel -eq 'MmpC' })
$active      = @($enrollments | Where-Object { $_.IsActive }) | Select-Object -First 1
$orphans     = @($enrollments | Where-Object { -not $_.IsActive })
$mmpcActive  = @($mmpc | Where-Object { $_.IsActive }) | Select-Object -First 1

# Baseline = what the device SHOULD be enrolled in. After a migration the
# device's own join state is the correct answer, so the parameter is an override
# rather than a requirement.
$baseline = $ExpectedTenantId
if (-not $baseline) { $baseline = $join.TenantId }

$verdict = 'Unknown'
$reason  = $null
$stale   = @()
$flags   = @()

# The MMP-C channel carries its OWN AADTenantID, and it is not necessarily the
# same as the Intune channel's. Measured on a migrated device: a WinDC
# enrollment survived the tenant move still bound to the old tenant and its
# federated IdP (UPN @old-domain, IsFederated 1) while the device was Entra
# joined to the new one. Evaluating only the Intune channel misses that entirely
# once the Intune enrollment has been retired away.
$mmpcTenant   = $null
if ($mmpcActive) { $mmpcTenant = Get-EnrollmentTenant -Enrollment $mmpcActive }
$mmpcMismatch = [bool]($mmpcTenant -and $baseline -and $mmpcTenant -ne $baseline)

if ($mmpcMismatch)            { $flags += 'MmpcCrossTenant' }
if ($orphans.Count -gt 0)     { $flags += 'OrphanResidue' }
if ($join.JoinInfoKeys -gt 1) { $flags += 'MultipleJoinInfo' }

# MmpcEnrollmentFlag - a REG_DWORD directly on the Enrollments root (not under
# any GUID). Undocumented by Microsoft, but independently reported by Quest
# (two support KBs), a Microsoft techcommunity thread, the fleetdm issue
# tracker and several comments on Rastello's teardown post: a value of 2 leaves
# auto-enrollment failing with 'Auto MDM Enroll: Device Credential (0x1),
# Failed (Bad request (400).)' / 0x80190190, and setting it to 0 or deleting it
# lets enrollment succeed immediately. Quest's own cutover tooling deletes it.
#
# NOTE THE POLARITY: 2 is the blocking state, 0 is the healthy state. Setting
# it to 0 does NOT suppress enrollment.
#
# The meaning of the values is not documented anywhere and the community does
# not know it either. The 'accepted answer' on the techcommunity thread claims
# 2 means successfully enrolled - it is contradicted by the thread's own OP and
# by every other report, so it is ignored here. fleetdm also reports devices
# where the value is already 0 or absent and enrollment still fails, so this is
# a signal, not a diagnosis.
#
# Reported, never acted on - this script is read-only.
$mmpcFlag = Get-RegValue -Path 'HKLM:\SOFTWARE\Microsoft\Enrollments' -Name 'MmpcEnrollmentFlag'
if ($null -ne $mmpcFlag -and [int]$mmpcFlag -eq 2) { $flags += 'MmpcEnrollmentFlagBlocking' }

# LinkedEnrollment - the EXPLICIT parent -> child pointer, read off the Intune
# enrollment key. This beats inferring the relationship from ProviderID.
#
# Microsoft documents the CSP node ./Device/Vendor/MSFT/DMClient/Provider/
# {ProviderID}/LinkedEnrollment with children DiscoveryEndpoint, Enroll,
# EnrollStatus, LastError and Unenroll, and publishes the EnrollStatus mapping
# below. DiscoveryEndpoint is listed as Windows Insider Preview. MMPCLocked is
# NOT a documented CSP node - it is client-internal state whose meaning is
# unknown (call4cloud guesses 'locked into MMP-C' and says so explicitly), so
# it is reported raw and nothing is inferred from it.
#
# Because LinkedEnrollment is a SUBKEY OF THE INTUNE ENROLLMENT, removing the
# Intune enrollment destroys the pointer to the MMP-C child along with it. That
# is the mechanical reason an orphaned MMP-C enrollment is hard to attribute
# after the fact, and it is why the absence of a LinkedEnrollmentId next to a
# live MMP-C enrollment is treated as an orphan signature.
$linked = @{ Id = $null; Status = $null; StatusText = $null; LastError = $null; Locked = $null; Endpoint = $null }
if ($active) {
    $lp = "HKLM:\SOFTWARE\Microsoft\Enrollments\$($active.EnrollmentId)\LinkedEnrollment"
    if (Test-Path $lp) {
        $linked.Id        = Get-RegValue -Path $lp -Name 'LinkedEnrollmentId'
        $linked.Status    = Get-RegValue -Path $lp -Name 'EnrollStatus'
        $linked.LastError = Get-RegValue -Path $lp -Name 'LastError'
        $linked.Locked    = Get-RegValue -Path $lp -Name 'MMPCLocked'
        $linked.Endpoint  = Get-RegValue -Path $lp -Name 'DiscoveryEndpoint'
    }
}

# Mapping published by Microsoft in the DMClient CSP reference.
$enrollStatusText = @{
    0 = 'Undefined'; 1 = 'Enrollment Not started'; 2 = 'Enrollment In Progress'
    3 = 'Enrollment Failed'; 4 = 'Enrollment Succeeded'
    5 = 'Unenrollment Not started'; 6 = 'UnEnrollment In Progress'
    7 = 'UnEnrollment Failed'; 8 = 'UnEnrollment Succeeded'
}
if ($null -ne $linked.Status -and $enrollStatusText.ContainsKey([int]$linked.Status)) {
    $linked.StatusText = $enrollStatusText[[int]$linked.Status]
}

# DO NOT USE EnrollStatus AS A LIVENESS SIGNAL. Measured on a healthy host:
# EnrollStatus = 3 ('Enrollment Failed') and LastError = 0, while the MMP-C
# enrollment it points at is fully live with a valid cert and a running task
# set. The value evidently records the outcome of some earlier attempt and is
# not refreshed on success. It is reported for triage only and deliberately
# carries no weight in LivenessScore or in the verdict.

# A live MMP-C enrollment that nothing points at. Strongest available evidence
# that the parent Intune enrollment was removed out from under it.
if ($mmpcActive -and -not $linked.Id) { $flags += 'MmpcUnlinked' }

if (-not $baseline) {
    $verdict = 'Unknown'
    $reason  = 'Device is not Entra joined (no CloudDomainJoin TenantId) and no -ExpectedTenantId supplied.'
}
elseif (-not $active) {
    # An orphaned MMP-C enrollment is a distinct failure mode from plain
    # unmanaged: Windows still counts as MDM enrolled, so re-enrollment no-ops.
    if ($mmpcActive) {
        $flags += 'MmpcOrphan'
        if ($mmpcMismatch) {
            $verdict = 'MmpcOrphanCrossTenant'
            $reason  = ("No live Intune (MS DM Server) enrollment. MMP-C enrollment {0} is live and bound to tenant {1}, not the joined tenant {2} - it predates the migration. Windows counts as enrolled, so auto-enrollment will not retry, and any renewal authenticates against the old tenant's IdP." -f $mmpcActive.EnrollmentId, $mmpcTenant, $baseline)
        }
        else {
            $verdict = 'MmpcOrphan'
            $reason  = ("No live Intune (MS DM Server) enrollment, but MMP-C enrollment {0} is live (ProviderID 'Microsoft Device Management'). Windows counts as enrolled, so auto-enrollment will not retry." -f $mmpcActive.EnrollmentId)
        }
    }
    elseif ($enrollments.Count -gt 0) {
        $verdict = 'NotEnrolled'
        $reason = ("No live MDM enrollment; {0} inert enrollment key(s) present." -f $enrollments.Count)
    }
    else {
        $verdict = 'NotEnrolled'
        $reason = 'No Intune (ProviderID = MS DM Server) enrollment present.'
    }
}
else {
    # THE TEST: the ACTIVE enrollment's tenant vs the joined tenant. Orphans are
    # reported via OrphanEnrollmentCount but never drive the verdict.
    $activeTenant = Get-EnrollmentTenant -Enrollment $active

    if (-not $activeTenant) {
        $verdict = 'Unknown'
        $reason  = 'Active enrollment exposes neither AADTenantID nor a tenant-suffixed UPN.'
    }
    elseif ($activeTenant -ne $baseline) {
        $stale = @($active)
        $flags += 'IntuneCrossTenant'
        if ($StaleTenantId -and $activeTenant -eq $StaleTenantId) { $verdict = 'StaleKnownTenant' }
        else { $verdict = 'StaleOtherTenant' }
        $reason = ("Active enrollment tenant {0} != joined tenant {1}." -f $activeTenant, $baseline)
    }
    else {
        $verdict = 'Clean'
        $reason  = ("Active enrollment tenant matches joined tenant {0}." -f $baseline)
        # A healthy Intune channel does not vouch for the MMP-C one; they are
        # enrolled and torn down independently.
        if ($mmpcMismatch) {
            $reason += (" MMP-C enrollment {0} is bound to tenant {1} instead - see Flags." -f $mmpcActive.EnrollmentId, $mmpcTenant)
        }
    }
}

if (-not $flags) { $flags = @() }

#endregion Evaluation

#region Output

$result = [pscustomobject]@{
    ComputerName          = $env:COMPUTERNAME
    Verdict               = $verdict
    Flags                 = @($flags)
    Reason                = $reason
    JoinedTenantId        = $join.TenantId
    ExpectedTenantId      = $baseline
    JoinedUserEmail       = $join.UserEmail
    EntraDeviceId         = $join.DeviceId
    AzureAdJoined         = $join.AzureAdJoined
    DomainJoined          = $join.DomainJoined
    MdmUrl                = $join.MdmUrl
    JoinInfoKeyCount      = $join.JoinInfoKeys
    EnrollmentCount       = $enrollments.Count
    ActiveEnrollmentId    = $(if ($active) { $active.EnrollmentId } else { $null })
    ActiveTenantId        = $(if ($active) { Get-EnrollmentTenant -Enrollment $active } else { $null })
    OrphanEnrollmentCount = $orphans.Count
    StaleEnrollmentCount  = $stale.Count
    MmpcEnrollmentCount   = $mmpc.Count
    MmpcActiveId          = $(if ($mmpcActive) { $mmpcActive.EnrollmentId } else { $null })
    MmpcActiveTenantId    = $mmpcTenant
    MmpcTenantMismatch    = $mmpcMismatch
    MmpcActiveUpn         = $(if ($mmpcActive) { $mmpcActive.Upn } else { $null })
    MmpcActiveLastRun     = $(if ($mmpcActive) { $mmpcActive.SyncTaskLastRun } else { $null })
    MmpcEnrollmentFlag    = $mmpcFlag
    LinkedEnrollmentId    = $linked.Id
    LinkedEnrollStatus    = $linked.Status
    LinkedEnrollStatusText = $linked.StatusText
    LinkedLastError       = $linked.LastError
    LinkedMmpcLocked      = $linked.Locked
    LinkedDiscoveryEndpoint = $linked.Endpoint
    Enrollments           = $allFound
    CollectedUtc          = (Get-Date).ToUniversalTime().ToString('o')
}

if ($Json) { $result | ConvertTo-Json -Depth 5 -Compress }
else { $result }

#endregion Output
