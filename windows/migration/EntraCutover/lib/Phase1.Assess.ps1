<#
    Phase 0/1 - Assess. Dot-sourced by Invoke-EntraCutover.ps1; runs under the
    caller's Set-StrictMode -Version 2.0. STRICTLY READ-ONLY: no Invoke-Step /
    Set-StepState, no files/dirs/registry writes, no device mutation. The only
    allowed state write is the single Save-CutoverState call at the very end,
    gated on $Ctx.ReadOnly (see Invoke-PhaseAssess).

    Private helpers: Get-EC<Noun>. Get-ECKfmStatus is SHARED - the Prepare
    phase (lib\Phase2.Prepare.ps1) calls it too, so its return shape
    (PolicyPresent/PolicyTenantId/ProfilesChecked/ProfilesHealthy/Healthy) is
    a contract, not just an implementation detail of this file.
#>

function Get-ECOptionValue {
    # $Options is a hashtable (fresh run) or an OrderedDictionary/PSCustomObject
    # (round-tripped through state.json) depending on caller - handle all three.
    param($Options, [Parameter(Mandatory)][string]$Name)
    if ($null -eq $Options) { return $null }
    if ($Options -is [System.Collections.IDictionary]) {
        if ($Options.Contains($Name)) { return $Options[$Name] }
        return $null
    }
    $prop = $Options.PSObject.Properties[$Name]
    if ($prop) { return $prop.Value }
    return $null
}

function Get-ECJoinState {
    param($Options)
    $blockers = @()
    $warnings = @()
    $dsreg = Get-DsregStatus

    $azureAdJoined = $null
    $domainJoined  = $null
    $tenantId      = $null
    $deviceId      = $null
    $domainName    = $null
    if ($dsreg.ContainsKey('AzureAdJoined')) { $azureAdJoined = $dsreg['AzureAdJoined'] }
    if ($dsreg.ContainsKey('DomainJoined'))  { $domainJoined  = $dsreg['DomainJoined'] }
    if ($dsreg.ContainsKey('TenantId'))      { $tenantId      = $dsreg['TenantId'] }
    if ($dsreg.ContainsKey('DeviceId'))      { $deviceId      = $dsreg['DeviceId'] }
    if ($dsreg.ContainsKey('DomainName'))    { $domainName    = $dsreg['DomainName'] }

    if ($azureAdJoined -ne 'YES' -or $domainJoined -ne 'YES') {
        $blockers += ("Device is not hybrid Entra joined (AzureAdJoined={0}, DomainJoined={1})." -f $azureAdJoined, $domainJoined)
    }

    $expectedTenant = Get-ECOptionValue -Options $Options -Name 'TenantId'
    if ($expectedTenant -and $tenantId -and ($expectedTenant -ne $tenantId)) {
        $blockers += ("TenantId mismatch: device is joined to {0}, expected {1}." -f $tenantId, $expectedTenant)
    }

    return [ordered]@{
        Blockers   = $blockers
        Warnings   = $warnings
        TenantId   = $tenantId
        DeviceId   = $deviceId
        DomainName = $domainName
    }
}

function Get-ECEnrollmentInventory {
    $blockers = @()
    $warnings = @()
    $guidRx   = '^[0-9A-Fa-f]{8}-([0-9A-Fa-f]{4}-){3}[0-9A-Fa-f]{12}$'
    $enrollmentId = $null
    $upn          = $null

    $enrollRoot = 'HKLM:\SOFTWARE\Microsoft\Enrollments'
    if (Test-Path $enrollRoot) {
        $subkeys = @(Get-ChildItem -Path $enrollRoot -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -match $guidRx })
        foreach ($k in $subkeys) {
            $p = Get-ItemProperty -Path $k.PSPath -ErrorAction SilentlyContinue
            if ($p -and $p.PSObject.Properties['ProviderID'] -and $p.ProviderID -eq 'MS DM Server') {
                $enrollmentId = $k.PSChildName
                if ($p.PSObject.Properties['UPN']) { $upn = $p.UPN }
                break
            }
        }
    }
    if (-not $enrollmentId) {
        $warnings += 'No Intune (MS DM Server) MDM enrollment found - nothing to purge in Teardown.'
    }

    $cert = Get-ChildItem -Path 'Cert:\LocalMachine\My' -ErrorAction SilentlyContinue |
        Where-Object { $_.Issuer -match 'Microsoft Intune.*MDM Device CA' } |
        Sort-Object NotAfter -Descending | Select-Object -First 1
    $certThumb = $null
    $certNotAfter = $null
    if ($cert) {
        $certThumb    = $cert.Thumbprint
        $certNotAfter = $cert.NotAfter
    }

    return [ordered]@{
        Blockers          = $blockers
        Warnings          = $warnings
        EnrollmentId      = $enrollmentId
        EnrollmentUpn     = $upn
        MdmCertThumbprint = $certThumb
        MdmCertNotAfter   = $certNotAfter
    }
}

function Get-ECPendingReboot {
    $reasons = @()
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') {
        $reasons += 'CBS RebootPending'
    }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') {
        $reasons += 'WindowsUpdate RebootRequired'
    }
    try {
        $pfro = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
        if ($pfro -and $pfro.PSObject.Properties['PendingFileRenameOperations'] -and @($pfro.PendingFileRenameOperations).Count -gt 0) {
            $reasons += 'PendingFileRenameOperations'
        }
    }
    catch { }

    $pending = (@($reasons).Count -gt 0)
    $blockers = @()
    if ($pending) {
        $blockers += ("Pending reboot detected ({0}); reboot before continuing." -f ($reasons -join ', '))
    }
    return [ordered]@{ Blockers = $blockers; Warnings = @(); PendingReboot = $pending }
}

function Get-ECBitLockerStatus {
    $warnings = @()
    $protectionStatus = $null
    $protectorTypes = @()
    try {
        $osVol = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop
        if ($osVol) {
            $protectionStatus = [string]$osVol.ProtectionStatus
            if ($osVol.PSObject.Properties['KeyProtector'] -and $osVol.KeyProtector) {
                $protectorTypes = @($osVol.KeyProtector | ForEach-Object { [string]$_.KeyProtectorType })
            }
            $hasPin = (@($protectorTypes | Where-Object { $_ -match 'Pin' }).Count -gt 0)
            if ($hasPin) {
                $warnings += 'BitLocker OS volume uses a TPM+PIN protector - interactive PIN entry across unattended reboots can stall the migration.'
            }
        }
    }
    catch {
        Write-Log "BitLocker check unavailable: $($_.Exception.Message)" 'INFO'
    }
    return [ordered]@{ Blockers = @(); Warnings = $warnings; ProtectionStatus = $protectionStatus; Protectors = $protectorTypes }
}

function Get-ECDcReachability {
    param([string]$DomainName, $Options)
    $blockers = @()
    $warnings = @()
    $reachable = $null

    if (-not $DomainName) {
        $warnings += 'Domain name unknown (join-state check failed) - cannot test DC reachability.'
        return [ordered]@{ Blockers = $blockers; Warnings = $warnings; Reachable = $reachable }
    }

    try {
        $result = Invoke-Exe -Path 'nltest.exe' -Arguments @("/dsgetdc:$DomainName") -TimeoutSec 30
        $reachable = ($result.ExitCode -eq 0)
    }
    catch {
        $reachable = $false
        $warnings += "nltest probe failed: $($_.Exception.Message)"
    }

    $offlineUnjoin = [bool](Get-ECOptionValue -Options $Options -Name 'OfflineUnjoin')
    if (-not $reachable) {
        if ($offlineUnjoin) {
            $warnings += "Domain controller for $DomainName is unreachable; proceeding relies on -OfflineUnjoin."
        }
        else {
            $blockers += "Domain controller for $DomainName is unreachable. Restore connectivity or re-run with -OfflineUnjoin."
        }
    }
    return [ordered]@{ Blockers = $blockers; Warnings = $warnings; Reachable = $reachable }
}

function Get-ECEndpointReachability {
    $urls = @('https://login.microsoftonline.com', 'https://enrollment.manage.microsoft.com')
    $details = @()
    $allReachable = $true

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    foreach ($url in $urls) {
        $reachable = $false
        $detail = $null
        try {
            $resp = Invoke-WebRequest -Uri $url -Method Head -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
            $reachable = $true
            $detail = "HTTP $([int]$resp.StatusCode)"
        }
        catch {
            # PS 5.1 throws WebException, PS7 throws HttpResponseException; both
            # expose .Response.StatusCode when the server actually answered.
            $respObj = $_.Exception.Response
            if ($respObj -and $null -ne $respObj.StatusCode) {
                $reachable = $true
                $detail = "HTTP $([int]$respObj.StatusCode)"
            }
            else {
                $reachable = $false
                $detail = $_.Exception.Message
            }
        }
        $details += ("{0}: reachable={1} ({2})" -f $url, $reachable, $detail)
        if (-not $reachable) { $allReachable = $false }
    }

    $blockers = @()
    if (-not $allReachable) {
        $blockers += ("Entra/Intune endpoint(s) unreachable - join and MDM enrollment will fail post-unjoin: {0}" -f ($details -join '; '))
    }
    return [ordered]@{ Blockers = $blockers; Warnings = @(); Reachable = $allReachable; Detail = $details }
}

function Get-ECPpkgReadiness {
    param($Options)
    $blockers = @()
    $warnings = @()

    $mode     = Get-ECOptionValue -Options $Options -Name 'Mode'
    $joinMode = Get-ECOptionValue -Options $Options -Name 'JoinMode'
    $ppkgPath = Get-ECOptionValue -Options $Options -Name 'PpkgPath'

    if ($joinMode -eq 'Ppkg' -and $mode -eq 'Migrate') {
        if (-not $ppkgPath -or -not (Test-Path $ppkgPath -ErrorAction SilentlyContinue)) {
            $blockers += "Provisioning package not found: $ppkgPath"
        }
        else {
            $fileName = Split-Path -Path $ppkgPath -Leaf
            if ($fileName -match 'Expires-(\d{4}-\d{2}-\d{2})') {
                $expiryText = $Matches[1]
                $expiryDate = $null
                if ([datetime]::TryParse($expiryText, [ref]$expiryDate) -and $expiryDate -lt (Get-Date).Date) {
                    $blockers += "Bulk token expired ($expiryText per ppkg filename) - rebuild the provisioning package."
                }
                else {
                    $warnings += "Ppkg filename indicates bulk-token expiry $expiryText - confirm it is still valid; the token's real expiry cannot be read locally."
                }
            }
            else {
                $warnings += 'Ppkg filename does not follow the Expires-YYYY-MM-DD convention - bulk-token expiry cannot be verified locally.'
            }
        }
    }
    return [ordered]@{ Blockers = $blockers; Warnings = $warnings }
}

function Get-ECKfmStatus {
    <#
        SHARED with Phase2.Prepare.ps1 - keep the return shape stable:
        PolicyPresent / PolicyTenantId / ProfilesChecked / ProfilesHealthy / Healthy.
    #>
    param($Options)
    $policyPresent  = $false
    $policyTenantId = $null
    try {
        $policyKey = 'HKLM:\SOFTWARE\Policies\Microsoft\OneDrive'
        if (Test-Path $policyKey) {
            $kfm = Get-ItemProperty -Path $policyKey -Name 'KFMSilentOptIn' -ErrorAction SilentlyContinue
            if ($kfm -and $kfm.PSObject.Properties['KFMSilentOptIn'] -and $kfm.KFMSilentOptIn) {
                $policyPresent  = $true
                $policyTenantId = $kfm.KFMSilentOptIn
            }
        }
    }
    catch { }

    $expectedTenant = Get-ECOptionValue -Options $Options -Name 'TenantId'
    $policyTenantOk = $true
    if ($expectedTenant -and $policyTenantId) {
        $policyTenantOk = ($policyTenantId -ieq $expectedTenant)
    }

    $hives = @()
    try {
        $hives = @(Get-ChildItem -Path 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue |
            Where-Object { $_.PSChildName -match '^S-1-5-21-\d+-\d+-\d+-\d+$' })
    }
    catch { $hives = @() }

    $profilesChecked = 0
    $profilesHealthy = 0
    foreach ($hive in $hives) {
        $profilesChecked++
        $shellFolders = "Registry::HKEY_USERS\$($hive.PSChildName)\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
        $healthy = $false
        if (Test-Path $shellFolders) {
            try {
                $folders = Get-ItemProperty -Path $shellFolders -ErrorAction Stop
                $watched = @()
                foreach ($name in @('Desktop', 'Personal', 'My Pictures')) {
                    if ($folders.PSObject.Properties[$name] -and $folders.$name) { $watched += $folders.$name }
                }
                if (@($watched).Count -gt 0 -and (@($watched | Where-Object { $_ -notmatch 'OneDrive' }).Count -eq 0)) {
                    $healthy = $true
                }
            }
            catch { }
        }
        if ($healthy) { $profilesHealthy++ }
    }

    $healthyOverall = $policyPresent -and $policyTenantOk -and ($profilesChecked -eq 0 -or $profilesHealthy -eq $profilesChecked)

    return [ordered]@{
        PolicyPresent   = $policyPresent
        PolicyTenantId  = $policyTenantId
        ProfilesChecked = $profilesChecked
        ProfilesHealthy = $profilesHealthy
        Healthy         = $healthyOverall
    }
}

function Get-ECLocalSidPrefix {
    # Best-effort: derive this machine's local-account SID prefix so a domain
    # profile can be told apart from a local S-1-5-21 profile.
    try {
        $lu = Get-LocalUser -ErrorAction Stop | Select-Object -First 1
        if (-not $lu) { return $null }
        $acct = New-Object System.Security.Principal.NTAccount("$env:COMPUTERNAME\$($lu.Name)")
        $sid = $acct.Translate([System.Security.Principal.SecurityIdentifier]).Value
        return $sid.Substring(0, $sid.LastIndexOf('-'))
    }
    catch { return $null }
}

function Get-ECProfileInventory {
    $localPrefix = Get-ECLocalSidPrefix
    $profiles = @()
    $wmiProfiles = @(Get-WmiObject -Class Win32_UserProfile -ErrorAction SilentlyContinue | Where-Object { -not $_.Special })

    foreach ($wp in $wmiProfiles) {
        $sid = $wp.SID
        $localPath = $wp.LocalPath
        $lastUse = $null
        if ($wp.LastUseTime) {
            try { $lastUse = [System.Management.ManagementDateTimeConverter]::ToDateTime($wp.LastUseTime) }
            catch { $lastUse = $null }
        }

        $isDomain = $false
        if ($sid -match '^S-1-5-21-') {
            $isDomain = $true
            if ($localPrefix -and $sid.StartsWith($localPrefix)) {
                $isDomain = $false
            }
            else {
                # Fallback when the local-prefix derivation failed: an account
                # that resolves back to THIS machine is local regardless of SID shape.
                try {
                    $secSid = New-Object System.Security.Principal.SecurityIdentifier($sid)
                    $acctName = $secSid.Translate([System.Security.Principal.NTAccount]).Value
                    if ($acctName -like "$env:COMPUTERNAME\*") { $isDomain = $false }
                }
                catch { }
            }
        }

        $profiles += [ordered]@{
            Path     = $localPath
            Sid      = $sid
            IsDomain = $isDomain
            LastUse  = $lastUse
        }
    }
    return [ordered]@{ Blockers = @(); Warnings = @(); Profiles = $profiles }
}

function Get-ECDataAtRisk {
    param([array]$Profiles)
    $blockers = @()
    $items = @()

    foreach ($p in $Profiles) {
        if (-not $p.IsDomain) { continue }
        if (-not $p.Path -or -not (Test-Path $p.Path -ErrorAction SilentlyContinue)) { continue }

        try {
            $pstFiles = @(Get-ChildItem -Path $p.Path -Recurse -Depth 4 -Include '*.pst', '*.ost' -Force -ErrorAction SilentlyContinue |
                Where-Object { -not $_.PSIsContainer })
            foreach ($f in $pstFiles) {
                if ($f.Extension -ieq '.pst') {
                    $items += ("PST at risk: {0} ({1:N0} MB)" -f $f.FullName, ($f.Length / 1MB))
                }
            }
        }
        catch { $items += "PST/OST scan failed for $($p.Path): $($_.Exception.Message)" }

        try {
            $downloads = Join-Path $p.Path 'Downloads'
            if (Test-Path $downloads -ErrorAction SilentlyContinue) {
                $sum = (Get-ChildItem -Path $downloads -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                if ($sum) {
                    $items += ("Downloads at risk: {0:N0} MB under {1}" -f ($sum / 1MB), $downloads)
                }
            }
        }
        catch { $items += "Downloads scan failed for $($p.Path): $($_.Exception.Message)" }

        try {
            $efsFiles = @(Get-ChildItem -Path $p.Path -Recurse -Depth 3 -Force -ErrorAction SilentlyContinue |
                Where-Object { -not $_.PSIsContainer -and ($_.Attributes -band [System.IO.FileAttributes]::Encrypted) })
            if (@($efsFiles).Count -gt 0) {
                $blockers += ("EFS-encrypted file(s) found under {0} ({1} file(s)) - data would be unreadable post-migration; decrypt first." -f $p.Path, @($efsFiles).Count)
            }
        }
        catch { $items += "EFS scan failed for $($p.Path): $($_.Exception.Message)" }
    }
    return [ordered]@{ Blockers = $blockers; Warnings = @(); Items = $items }
}

function Get-ECMiscSignals {
    $warnings = @()

    try {
        $dot3 = Get-Service -Name dot3svc -ErrorAction SilentlyContinue
        if ($dot3 -and $dot3.Status -eq 'Running') {
            $lan = Invoke-Exe -Path 'netsh.exe' -Arguments @('lan', 'show', 'profiles')
            if ($lan.Output -and ($lan.Output -match 'Profile')) {
                $warnings += 'Wired 802.1x profile(s) detected - machine-certificate authentication will break once the device leaves the domain.'
            }
        }
    }
    catch { }

    try {
        $wlan = Invoke-Exe -Path 'netsh.exe' -Arguments @('wlan', 'show', 'profiles')
        $wlanCount = @([regex]::Matches($wlan.Output, 'All User Profile\s*:')).Count
        if ($wlanCount -gt 0) {
            $warnings += ("{0} WLAN profile(s) found - per-user Wi-Fi credentials are lost after migration; redeploy via Intune." -f $wlanCount)
        }
    }
    catch { }

    try {
        $vpns = @(Get-VpnConnection -AllUserConnections -ErrorAction Stop)
        if (@($vpns).Count -gt 0) {
            $warnings += ("{0} all-user VPN connection(s) found - VPN client configuration may not survive migration." -f @($vpns).Count)
        }
    }
    catch { }

    try {
        $hives = @(Get-ChildItem -Path 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue |
            Where-Object { $_.PSChildName -match '^S-1-5-21-\d+-\d+-\d+-\d+$' })
        foreach ($hive in $hives) {
            $netPath = "Registry::HKEY_USERS\$($hive.PSChildName)\Network"
            if (Test-Path $netPath) {
                $drives = @(Get-ChildItem -Path $netPath -ErrorAction SilentlyContinue)
                foreach ($d in $drives) {
                    $props = Get-ItemProperty -Path $d.PSPath -ErrorAction SilentlyContinue
                    $remote = $null
                    if ($props -and $props.PSObject.Properties['RemotePath']) { $remote = $props.RemotePath }
                    Write-Log ("mapped drive (SID {0}): {1}: -> {2}" -f $hive.PSChildName, $d.PSChildName, $remote) 'INFO'
                }
            }
        }
    }
    catch { }

    try {
        $ngc = Join-Path $env:windir 'ServiceProfiles\LocalService\AppData\Local\Microsoft\NGC'
        if (Test-Path $ngc -ErrorAction SilentlyContinue) {
            $items = @(Get-ChildItem -Path $ngc -ErrorAction SilentlyContinue)
            if (@($items).Count -gt 0) {
                Write-Log 'Windows Hello for Business container present - Hello re-enrollment will be required post-migration.' 'INFO'
            }
        }
    }
    catch { }

    return [ordered]@{ Blockers = @(); Warnings = $warnings }
}

function Invoke-PhaseAssess {
    param([hashtable]$Ctx)

    Write-Log 'Phase Assess: read-only preflight starting.' 'STEP'
    $Options = $Ctx.Options

    $blockers = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]

    $oldDeviceId        = $null
    $tenantId           = $null
    $domain             = $null
    $enrollmentId       = $null
    $enrollmentUpn      = $null
    $mdmCertThumbprint  = $null
    $mdmCertNotAfter    = $null
    $bitlockerProtection = $null
    $bitlockerProtectors = @()
    $dcReachable        = $null
    $endpointsReachable = $null
    $kfm                = [ordered]@{ PolicyPresent = $false; PolicyTenantId = $null; ProfilesChecked = 0; ProfilesHealthy = 0; Healthy = $false }
    $profiles           = @()
    $dataAtRisk         = @()
    $pendingReboot      = $false

    # 1. Join state ------------------------------------------------------------
    try {
        $r = Get-ECJoinState -Options $Options
        foreach ($b in @($r.Blockers)) { $blockers.Add($b) }
        $tenantId = $r.TenantId; $oldDeviceId = $r.DeviceId; $domain = $r.DomainName
        Write-Log ("Join state: TenantId={0} DeviceId={1} Domain={2}" -f $tenantId, $oldDeviceId, $domain) 'INFO'
    }
    catch { $warnings.Add("check JoinState failed: $($_.Exception.Message)") }

    # 2. Enrollment / MDM cert --------------------------------------------------
    try {
        $r = Get-ECEnrollmentInventory
        foreach ($b in @($r.Blockers)) { $blockers.Add($b) }
        foreach ($w in @($r.Warnings)) { $warnings.Add($w) }
        $enrollmentId = $r.EnrollmentId; $enrollmentUpn = $r.EnrollmentUpn
        $mdmCertThumbprint = $r.MdmCertThumbprint; $mdmCertNotAfter = $r.MdmCertNotAfter
        Write-Log ("Enrollment: Id={0} Upn={1} MdmCert={2} NotAfter={3}" -f $enrollmentId, $enrollmentUpn, $mdmCertThumbprint, $mdmCertNotAfter) 'INFO'
    }
    catch { $warnings.Add("check Enrollment failed: $($_.Exception.Message)") }

    # 3. Pending reboot ----------------------------------------------------------
    try {
        $r = Get-ECPendingReboot
        foreach ($b in @($r.Blockers)) { $blockers.Add($b) }
        $pendingReboot = $r.PendingReboot
        Write-Log ("Pending reboot: {0}" -f $pendingReboot) 'INFO'
    }
    catch { $warnings.Add("check PendingReboot failed: $($_.Exception.Message)") }

    # 4. BitLocker ----------------------------------------------------------------
    try {
        $r = Get-ECBitLockerStatus
        foreach ($w in @($r.Warnings)) { $warnings.Add($w) }
        $bitlockerProtection = $r.ProtectionStatus; $bitlockerProtectors = $r.Protectors
        Write-Log ("BitLocker OS volume: Protection={0} Protectors={1}" -f $bitlockerProtection, ($bitlockerProtectors -join ',')) 'INFO'
    }
    catch { $warnings.Add("check BitLocker failed: $($_.Exception.Message)") }

    # 5. DC reachability -----------------------------------------------------------
    try {
        $r = Get-ECDcReachability -DomainName $domain -Options $Options
        foreach ($b in @($r.Blockers)) { $blockers.Add($b) }
        foreach ($w in @($r.Warnings)) { $warnings.Add($w) }
        $dcReachable = $r.Reachable
        Write-Log ("DC reachability ({0}): {1}" -f $domain, $dcReachable) 'INFO'
    }
    catch { $warnings.Add("check DcReachability failed: $($_.Exception.Message)") }

    # 6. Entra / Intune endpoint reachability --------------------------------------
    try {
        $r = Get-ECEndpointReachability
        foreach ($b in @($r.Blockers)) { $blockers.Add($b) }
        $endpointsReachable = $r.Reachable
        Write-Log ("Entra/Intune endpoints reachable: {0}" -f $endpointsReachable) 'INFO'
    }
    catch { $warnings.Add("check Endpoints failed: $($_.Exception.Message)") }

    # 7. Ppkg readiness -------------------------------------------------------------
    try {
        $r = Get-ECPpkgReadiness -Options $Options
        foreach ($b in @($r.Blockers)) { $blockers.Add($b) }
        foreach ($w in @($r.Warnings)) { $warnings.Add($w) }
    }
    catch { $warnings.Add("check Ppkg failed: $($_.Exception.Message)") }

    # 8. KFM ---------------------------------------------------------------------
    try {
        $kfm = Get-ECKfmStatus -Options $Options
        $skipKfmGate = [bool](Get-ECOptionValue -Options $Options -Name 'SkipKfmGate')
        if (-not $kfm.Healthy) {
            $msg = ("KFM not healthy (PolicyPresent={0}, ProfilesHealthy={1}/{2})." -f $kfm.PolicyPresent, $kfm.ProfilesHealthy, $kfm.ProfilesChecked)
            if ($skipKfmGate) { $warnings.Add($msg) } else { $blockers.Add($msg) }
        }
        Write-Log ("KFM: PolicyPresent={0} PolicyTenantId={1} ProfilesHealthy={2}/{3}" -f $kfm.PolicyPresent, $kfm.PolicyTenantId, $kfm.ProfilesHealthy, $kfm.ProfilesChecked) 'INFO'
    }
    catch { $warnings.Add("check Kfm failed: $($_.Exception.Message)") }

    # 9. Profile inventory ---------------------------------------------------------
    try {
        $r = Get-ECProfileInventory
        $profiles = $r.Profiles
        Write-Log ("Profiles discovered: {0} (domain: {1})" -f @($profiles).Count, @($profiles | Where-Object { $_.IsDomain }).Count) 'INFO'
    }
    catch { $warnings.Add("check ProfileInventory failed: $($_.Exception.Message)") }

    # 10. Data at risk (fresh-profile strategy) -------------------------------------
    try {
        $r = Get-ECDataAtRisk -Profiles $profiles
        foreach ($b in @($r.Blockers)) { $blockers.Add($b) }
        $dataAtRisk = $r.Items
        Write-Log ("Data-at-risk items: {0}" -f @($dataAtRisk).Count) 'INFO'
    }
    catch { $warnings.Add("check DataAtRisk failed: $($_.Exception.Message)") }

    # 11. Misc warnings (802.1x, WLAN, VPN, mapped drives, WHfB) --------------------
    try {
        $r = Get-ECMiscSignals
        foreach ($w in @($r.Warnings)) { $warnings.Add($w) }
    }
    catch { $warnings.Add("check MiscSignals failed: $($_.Exception.Message)") }

    $blockersArr = @($blockers.ToArray())
    $warningsArr = @($warnings.ToArray())
    $ready = (@($blockersArr).Count -eq 0)

    $verdict = [pscustomobject][ordered]@{
        ComputerName        = $env:COMPUTERNAME
        CollectedAt         = Get-Date
        Ready               = $ready
        Blockers            = $blockersArr
        Warnings            = $warningsArr
        OldDeviceId         = $oldDeviceId
        TenantId            = $tenantId
        Domain              = $domain
        EnrollmentId        = $enrollmentId
        EnrollmentUpn       = $enrollmentUpn
        MdmCertThumbprint   = $mdmCertThumbprint
        MdmCertNotAfter     = $mdmCertNotAfter
        BitLockerProtection = $bitlockerProtection
        BitLockerProtectors = $bitlockerProtectors
        DcReachable         = $dcReachable
        EndpointsReachable  = $endpointsReachable
        Kfm                 = $kfm
        Profiles            = $profiles
        DataAtRisk          = $dataAtRisk
        PendingReboot       = $pendingReboot
    }

    Write-Log ("Assess complete: Ready={0} Blockers={1} Warnings={2}" -f $ready, @($blockersArr).Count, @($warningsArr).Count) $(if ($ready) { 'SUCCESS' } else { 'ERROR' })

    # Console summary - the one place Write-Host is allowed in this file.
    Write-Host ''
    Write-Host '  ============================================================' -ForegroundColor DarkGray
    if ($ready) { Write-Host '   VERDICT: READY' -ForegroundColor Green }
    else { Write-Host '   VERDICT: BLOCKED' -ForegroundColor Red }
    Write-Host '  ============================================================' -ForegroundColor DarkGray
    foreach ($b in $blockersArr) { Write-Host ("   BLOCKER: {0}" -f $b) -ForegroundColor Red }
    foreach ($w in $warningsArr) { Write-Host ("   WARNING: {0}" -f $w) -ForegroundColor Yellow }
    Write-Host ''

    if (-not ($Ctx.ContainsKey('ReadOnly') -and $Ctx.ReadOnly)) {
        $Ctx.State.Device = @{
            OldDeviceId  = $oldDeviceId
            TenantId     = $tenantId
            Domain       = $domain
            EnrollmentId = $enrollmentId
        }
        Save-CutoverState -State $Ctx.State
    }

    return $verdict
}
