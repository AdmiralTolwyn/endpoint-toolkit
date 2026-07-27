#Requires -Version 5.1
<#
    Phase 3: Teardown. DESTRUCTIVE. Intune enrollment purge, dsregcmd /leave
    (Entra hybrid leave), and domain unjoin (point of no return), then a
    reboot request. Dot-sourced by Invoke-EntraCutover.ps1 - see that file's
    "PHASE IMPLEMENTATIONS" contract comment for the $Ctx shape and the
    Invoke-Step / Write-Log / Invoke-Exe / Get-DsregStatus / Set-MigrationNotice
    / Request-Reboot / Save-CutoverState helpers used below.

    Enrollment-teardown logic (GUID discovery, EnterpriseMgmt task + COM
    folder removal, registry key list incl. dashed/no-dash OMADM variants,
    cert removal) is adapted from Repair-IntuneMdmCert.ps1's $RepairWorker,
    the field-proven implementation used elsewhere in this repo. Guardrails
    preserved: only concrete GUID child keys are ever iterated; the top-level
    Enrollments\Context, Ownership, Status and ValidNodePaths keys are never
    touched (only Enrollments\Status\<GUID>, never Enrollments\Status itself).
#>

Set-StrictMode -Version 2.0

# --------------------------------------------------------------------------
# PRIVATE HELPERS
# --------------------------------------------------------------------------
function Get-ECEnrollmentGuid {
    <#
        Intune ('MS DM Server') enrollment GUIDs under
        HKLM:\SOFTWARE\Microsoft\Enrollments. Only GUID-shaped child key names
        are ever considered - never the top-level Enrollments key itself.
    #>
    $guidRx = '^[0-9A-Fa-f]{8}-([0-9A-Fa-f]{4}-){3}[0-9A-Fa-f]{12}$'
    $ids = @()
    $root = 'HKLM:\SOFTWARE\Microsoft\Enrollments'
    if (Test-Path $root) {
        Get-ChildItem -Path $root -ErrorAction SilentlyContinue |
            Where-Object { $_.PSChildName -match $guidRx } |
            ForEach-Object {
                $p = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
                if ($p -and $p.ProviderID -eq 'MS DM Server') { $ids += $_.PSChildName }
            }
    }
    return @($ids)
}

function Remove-ECEnrollmentTask {
    <# Scheduled tasks under \Microsoft\Windows\EnterpriseMgmt\<GUID>\, then the
       (now empty) task folder itself via the Schedule.Service COM object. #>
    param([Parameter(Mandatory)][string]$EnrollmentId)
    $removed = 0
    $taskPath = "\Microsoft\Windows\EnterpriseMgmt\$EnrollmentId\"
    Get-ScheduledTask -TaskPath $taskPath -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            Unregister-ScheduledTask -TaskName $_.TaskName -TaskPath $taskPath -Confirm:$false -ErrorAction Stop
            $removed++
        }
        catch { Write-Log "task removal failed ($($_.TaskName)): $($_.Exception.Message)" 'WARN' }
    }
    try {
        $svc = New-Object -ComObject Schedule.Service
        $svc.Connect()
        $parent = $svc.GetFolder('\Microsoft\Windows\EnterpriseMgmt')
        $parent.DeleteFolder($EnrollmentId, 0)
    }
    catch { Write-Log "task folder removal skipped/failed for ${EnrollmentId}: $($_.Exception.Message)" 'WARN' }
    return $removed
}

function Remove-ECEnrollmentKey {
    <#
        Per-GUID registry cleanup. All roots use the dashed GUID except
        OMADM\Accounts, which on some builds used a no-dash GUID - both forms
        are targeted there, each Test-Path guarded. Never touches the
        top-level Enrollments/Status/etc. keys, only their <GUID> children.
    #>
    param([Parameter(Mandatory)][string]$EnrollmentId)
    $removed = 0
    $gNoDash = $EnrollmentId -replace '-', ''
    $keys = @(
        "HKLM:\SOFTWARE\Microsoft\Enrollments\$EnrollmentId",
        "HKLM:\SOFTWARE\Microsoft\Enrollments\Status\$EnrollmentId",
        "HKLM:\SOFTWARE\Microsoft\EnterpriseResourceManager\Tracked\$EnrollmentId",
        "HKLM:\SOFTWARE\Microsoft\PolicyManager\AdmxInstalled\$EnrollmentId",
        "HKLM:\SOFTWARE\Microsoft\PolicyManager\Providers\$EnrollmentId",
        "HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\$EnrollmentId",
        "HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\$gNoDash",
        "HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Logger\$EnrollmentId",
        "HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Sessions\$EnrollmentId"
    )
    foreach ($k in $keys) {
        if (Test-Path $k) {
            try { Remove-Item -Path $k -Recurse -Force -ErrorAction Stop; $removed++ }
            catch { Write-Log "registry key removal failed ${k}: $($_.Exception.Message)" 'WARN' }
        }
    }
    return $removed
}

function Get-ECDeviceProperty {
    <# Read a property off $Ctx.State.Device regardless of whether it is a
       hashtable (fresh state) or a PSCustomObject (state reloaded from
       state.json). Missing property/key -> $null, never an exception. #>
    param($Device, [Parameter(Mandatory)][string]$Name)
    if ($null -eq $Device) { return $null }
    if ($Device -is [hashtable]) {
        if ($Device.ContainsKey($Name)) { return $Device[$Name] }
        return $null
    }
    $prop = $null
    if ($Device.PSObject -and $Device.PSObject.Properties) { $prop = $Device.PSObject.Properties[$Name] }
    if ($prop) { return $prop.Value }
    return $null
}

# --------------------------------------------------------------------------
# PHASE ENTRY POINT
# --------------------------------------------------------------------------
function Invoke-PhaseTeardown {
    param([Parameter(Mandatory)][hashtable]$Ctx)

    # ---- 1. Quiesce MDM clients (best-effort) -----------------------------
    Invoke-Step -Ctx $Ctx -Name 'Teardown.StopMdmClients' -AllowFail -Action {
        $svcPresent = $false
        $svc = Get-Service -Name 'IntuneManagementExtension' -ErrorAction SilentlyContinue
        if ($svc) {
            $svcPresent = $true
            try { Stop-Service -Name 'IntuneManagementExtension' -Force -ErrorAction Stop }
            catch { Write-Log "stop IntuneManagementExtension failed: $($_.Exception.Message)" 'WARN' }
        }
        $killed = 0
        Get-Process -Name 'omadmclient' -ErrorAction SilentlyContinue | ForEach-Object {
            try { Stop-Process -Id $_.Id -Force -ErrorAction Stop; $killed++ }
            catch { Write-Log "stop omadmclient PID $($_.Id) failed: $($_.Exception.Message)" 'WARN' }
        }
        @{ ServicePresent = $svcPresent; ProcessesKilled = $killed }
    }

    # ---- 2. Enrollment purge (proven Repair-IntuneMdmCert logic) ----------
    Invoke-Step -Ctx $Ctx -Name 'Teardown.EnrollmentPurge' -Action {
        $enrollIds = @(Get-ECEnrollmentGuid)
        if ($enrollIds.Count -eq 0) {
            Write-Log 'no Intune (MS DM Server) enrollment GUID found; continuing with cert cleanup only.' 'INFO'
        }
        else {
            Write-Log ('Intune enrollment GUID(s): ' + ($enrollIds -join ', '))
        }

        $tasksRemoved = 0
        $keysRemoved  = 0
        foreach ($g in $enrollIds) {
            $tasksRemoved += (Remove-ECEnrollmentTask -EnrollmentId $g)
            $keysRemoved  += (Remove-ECEnrollmentKey -EnrollmentId $g)
        }

        # Known re-enrollment blocker (Quest KB 4381245).
        $enrollRoot = 'HKLM:\SOFTWARE\Microsoft\Enrollments'
        if (Test-Path $enrollRoot) {
            $flagProp = Get-ItemProperty -Path $enrollRoot -Name 'MmpcEnrollmentFlag' -ErrorAction SilentlyContinue
            if ($flagProp -and ($flagProp.PSObject.Properties.Name -contains 'MmpcEnrollmentFlag')) {
                try {
                    Remove-ItemProperty -Path $enrollRoot -Name 'MmpcEnrollmentFlag' -Force -ErrorAction Stop
                    Write-Log 'removed MmpcEnrollmentFlag value.'
                }
                catch { Write-Log "MmpcEnrollmentFlag removal failed: $($_.Exception.Message)" 'WARN' }
            }
        }

        # All Intune MDM device-CA certs, expired or not - a fresh enrollment
        # follows, so nothing issued under the old identity should survive.
        $certsRemoved = 0
        Get-ChildItem -Path 'Cert:\LocalMachine\My' -ErrorAction SilentlyContinue |
            Where-Object { $_.Issuer -match 'Microsoft Intune.*MDM Device CA' } |
            ForEach-Object {
                try { Remove-Item -Path $_.PSPath -Force -ErrorAction Stop; $certsRemoved++ }
                catch { Write-Log "cert removal failed $($_.Thumbprint): $($_.Exception.Message)" 'WARN' }
            }

        @{
            EnrollmentIds = $enrollIds
            TasksRemoved  = $tasksRemoved
            KeysRemoved   = $keysRemoved
            CertsRemoved  = $certsRemoved
        }
    }

    # ---- 3. Company Portal (cosmetic) --------------------------------------
    Invoke-Step -Ctx $Ctx -Name 'Teardown.CompanyPortal' -AllowFail -Action {
        $removed = $false
        try {
            $pkg = Get-AppxPackage -AllUsers -Name 'Microsoft.CompanyPortal' -ErrorAction SilentlyContinue
            if ($pkg) {
                Remove-AppxPackage -AllUsers -Package $pkg.PackageFullName -ErrorAction Stop
                $removed = $true
            }
        }
        catch { Write-Log "Company Portal removal failed (cosmetic, non-fatal): $($_.Exception.Message)" 'WARN' }
        @{ Removed = $removed }
    }

    # ---- 4. Entra hybrid leave ---------------------------------------------
    Invoke-Step -Ctx $Ctx -Name 'Teardown.EntraLeave' -Action {
        Invoke-Exe -Path "$env:windir\System32\dsregcmd.exe" -Arguments @('/debug', '/leave') | Out-Null

        $joined = $null
        for ($attempt = 1; $attempt -le 6; $attempt++) {
            Start-Sleep -Seconds 10
            $status = Get-DsregStatus
            $joined = $null
            if ($status.ContainsKey('AzureAdJoined')) { $joined = $status['AzureAdJoined'] }
            Write-Log "dsregcmd poll $attempt/6: AzureAdJoined=$joined"
            if ($joined -eq 'NO') { break }
        }
        if ($joined -ne 'NO') { throw 'device still Entra-joined after dsregcmd /leave' }
        @{ AzureAdJoined = $joined }
    }

    # ---- 5. Domain unjoin - POINT OF NO RETURN -----------------------------
    Invoke-Step -Ctx $Ctx -Name 'Teardown.DomainUnjoin' -Action {
        $cs = Get-WmiObject -Class Win32_ComputerSystem
        $domain = Get-ECDeviceProperty -Device $Ctx.State.Device -Name 'Domain'
        if (-not $domain) { $domain = $cs.Domain }

        if (-not $cs.PartOfDomain) {
            Write-Log 'already unjoined' 'INFO'
            return @{ Method = 'AlreadyUnjoined'; Workgroup = $null }
        }

        $hasCred = [bool]$Ctx.DomainCredential
        $dcReachable = $false
        if ($hasCred) {
            try {
                $nltest = Invoke-Exe -Path "$env:windir\System32\nltest.exe" -Arguments @("/dsgetdc:$domain")
                $dcReachable = ($nltest.ExitCode -eq 0)
            }
            catch {
                Write-Log "nltest DC probe failed: $($_.Exception.Message)" 'WARN'
                $dcReachable = $false
            }
        }

        $method = $null

        # Attempt A: graceful unjoin with the operator-supplied credential.
        if ($hasCred -and $dcReachable) {
            try {
                Remove-Computer -UnjoinDomainCredential $Ctx.DomainCredential -WorkgroupName 'WORKGROUP' -Force -ErrorAction Stop
                $method = 'Graceful'
                Write-Log 'graceful domain unjoin succeeded.' 'SUCCESS'
            }
            catch { Write-Log "graceful domain unjoin failed: $($_.Exception.Message)" 'WARN' }
        }
        elseif ($hasCred -and -not $dcReachable) {
            Write-Log 'DC not reachable (nltest /dsgetdc failed); graceful unjoin not attempted.' 'WARN'
        }
        else {
            Write-Log 'no domain credential supplied; graceful unjoin not attempted.' 'INFO'
        }

        # Attempt B: offline unjoin (only if A was skipped/failed and permitted).
        if (-not $method) {
            $offlineAllowed = [bool]$Ctx.Options.OfflineUnjoin
            if (-not $offlineAllowed) {
                throw "graceful domain unjoin unavailable (credential supplied=$hasCred, DC reachable=$dcReachable) and -OfflineUnjoin was not set. Supply -DomainCredential with DC line-of-sight, or re-run with -OfflineUnjoin to permit a local-only workgroup swap."
            }

            $adapters = @(Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'Up' })
            $disabledNames = @()
            try {
                foreach ($a in $adapters) {
                    try { Disable-NetAdapter -Name $a.Name -Confirm:$false -ErrorAction Stop; $disabledNames += $a.Name }
                    catch { Write-Log "adapter disable failed ($($a.Name)): $($_.Exception.Message)" 'WARN' }
                }

                try {
                    Remove-Computer -WorkgroupName 'WORKGROUP' -Force -ErrorAction Stop
                    $method = 'Offline-RemoveComputer'
                    Write-Log 'offline domain unjoin (Remove-Computer) succeeded.' 'SUCCESS'
                }
                catch {
                    Write-Log "offline Remove-Computer failed: $($_.Exception.Message); falling back to WMI UnjoinDomainOrWorkgroup." 'WARN'
                    $csOffline = Get-WmiObject -Class Win32_ComputerSystem
                    $ret = $csOffline.UnjoinDomainOrWorkgroup($null, $null, 0)
                    if ($ret.ReturnValue -eq 0) {
                        $method = 'Offline-WMI'
                        Write-Log 'offline domain unjoin (WMI UnjoinDomainOrWorkgroup) succeeded.' 'SUCCESS'
                    }
                    else {
                        throw "offline unjoin failed: both Remove-Computer and WMI UnjoinDomainOrWorkgroup failed (WMI ReturnValue=$($ret.ReturnValue))."
                    }
                }
            }
            finally {
                foreach ($name in $disabledNames) {
                    try { Enable-NetAdapter -Name $name -Confirm:$false -ErrorAction Stop }
                    catch { Write-Log "adapter re-enable failed ($name) - MANUAL INTERVENTION MAY BE REQUIRED: $($_.Exception.Message)" 'ERROR' }
                }
            }
        }

        $csFinal = Get-WmiObject -Class Win32_ComputerSystem
        if ($csFinal.PartOfDomain) { throw 'domain unjoin verification failed: PartOfDomain still true after unjoin attempt.' }

        @{ Method = $method; Workgroup = 'WORKGROUP' }
    }

    # ---- 6. Mark point of no return -----------------------------------------
    Invoke-Step -Ctx $Ctx -Name 'Teardown.MarkPonr' -Action {
        $ponr = (Get-Date).ToUniversalTime().ToString('o')
        if ($null -eq $Ctx.State.Device) { $Ctx.State.Device = @{} }
        try { $Ctx.State.Device['PointOfNoReturnUtc'] = $ponr }
        catch { $Ctx.State.Device | Add-Member -NotePropertyName 'PointOfNoReturnUtc' -MemberType NoteProperty -Value $ponr -Force }
        Save-CutoverState -State $Ctx.State
        Set-MigrationNotice -Caption 'Device migration in progress' -Text 'Cloud join in progress. Do NOT sign in until this notice is gone. Phase: Teardown complete.'
        @{ PointOfNoReturnUtc = $ponr }
    }

    Request-Reboot -Ctx $Ctx -Reason 'domain unjoin committed; continuing with Entra join after restart'
}
