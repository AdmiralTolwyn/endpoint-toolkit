#Requires -Version 5.1
<#
    Phase 5 - Finalize. Runs headless as SYSTEM at startup after the Join
    reboot. Escrows BitLocker recovery keys to the NEW Entra device object,
    resumes BitLocker, clears stale domain GPO state, restores the legal
    notice, schedules break-glass retirement, removes the resume task and
    emits the final migration report. No reboot at the end.

    Entry point: Invoke-PhaseFinalize -Ctx <hashtable>.
    Private helpers are named <Verb>-EC<Noun> per the lib contract.
#>

Set-StrictMode -Version 2.0

# --------------------------------------------------------------------------
# State accessors (state.json round-trips to PSCustomObject; null-guard all)
# --------------------------------------------------------------------------
function Get-ECProp {
    # Safe property/key read from a PSCustomObject, hashtable, or $null.
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        return $null
    }
    $prop = $Object.PSObject.Properties[$Name]
    if ($null -eq $prop) { return $null }
    return $prop.Value
}

function Get-ECStepData {
    # Data blob recorded by a prior step, or $null.
    param($State, [string]$StepName)
    $steps = Get-ECProp $State 'Steps'
    $step  = Get-ECProp $steps $StepName
    return (Get-ECProp $step 'Data')
}

function Set-ECStateProp {
    # Write to a state object: indexer first (hashtable), Add-Member fallback
    # (round-tripped PSCustomObject). Caller Saves afterward.
    param($Object, [string]$Name, $Value)
    try { $Object[$Name] = $Value }
    catch { $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force }
}

# --------------------------------------------------------------------------
# BitLocker escrow to the NEW device object
# --------------------------------------------------------------------------
function Invoke-ECBitLockerEscrow {
    # Returns @{ Volumes=@(@{MountPoint;Escrowed;Event845}); AnyFailed; HadProtectors }.
    if (-not (Get-Command Get-BitLockerVolume -ErrorAction SilentlyContinue)) {
        Write-Log 'BitLocker cmdlets unavailable - no recovery keys to escrow.' 'INFO'
        return @{ Volumes = @(); AnyFailed = $false; HadProtectors = $false }
    }

    $blVolumes = @(Get-BitLockerVolume -ErrorAction SilentlyContinue)
    $volumesData = @()
    $anyFailed = $false
    $hadProtectors = $false

    foreach ($vol in $blVolumes) {
        $mp = "$($vol.MountPoint)"
        $recProtectors = @($vol.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' })
        if ($recProtectors.Count -eq 0) { continue }
        $hadProtectors = $true

        $escrowed = $false
        $evt845   = $false
        $before   = Get-Date
        foreach ($kp in $recProtectors) {
            try {
                BackupToAAD-BitLockerKeyProtector -MountPoint $mp -KeyProtectorId $kp.KeyProtectorId -ErrorAction Stop | Out-Null
                $escrowed = $true
                Write-Log ("BitLocker recovery key escrowed for {0}." -f $mp)
            }
            catch {
                $anyFailed = $true
                Write-Log ("BitLocker escrow FAILED for {0} ({1}): {2}" -f $mp, $kp.KeyProtectorId, $_.Exception.Message) 'ERROR'
            }
        }

        # Corroborate the escrow with Event 845 newer than the backup call (informational).
        if ($escrowed) {
            try {
                $e = @(Get-WinEvent -FilterHashtable @{ LogName = 'Microsoft-Windows-BitLocker/BitLocker Management'; Id = 845; StartTime = $before } -MaxEvents 1 -ErrorAction Stop)
                if ($e.Count -gt 0) { $evt845 = $true }
            }
            catch { Write-Log ("Event 845 corroboration unavailable for {0}." -f $mp) 'INFO' }
        }

        $volumesData += , @{ MountPoint = $mp; Escrowed = $escrowed; Event845 = $evt845 }
    }

    if (-not $hadProtectors) { Write-Log 'No BitLocker recovery-password protectors present - nothing to escrow.' 'INFO' }
    return @{ Volumes = $volumesData; AnyFailed = $anyFailed; HadProtectors = $hadProtectors }
}

# --------------------------------------------------------------------------
# Stale domain GPO cleanup (never touches HKLM\SOFTWARE\Policies)
# --------------------------------------------------------------------------
function Clear-ECStaleGpo {
    param([hashtable]$Ctx)

    $ts = (Get-Date).ToString('yyyyMMdd-HHmmss')
    $gpRoot  = Join-Path $env:WinDir 'System32\GroupPolicy'
    $gpUsers = Join-Path $env:WinDir 'System32\GroupPolicyUsers'
    $backupBase = Join-Path $Ctx.Paths.Backup ("GroupPolicy_{0}" -f $ts)

    # Back up the machine + user GPO folders before touching them.
    foreach ($src in @($gpRoot, $gpUsers)) {
        if (Test-Path $src) {
            if (-not (Test-Path $backupBase)) { New-Item -ItemType Directory -Path $backupBase -Force | Out-Null }
            Copy-Item -Path $src -Destination $backupBase -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    # Back up the GP History key.
    $histKey = 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\History'
    try { New-RegistryBackup -KeyPath $histKey | Out-Null }
    catch { Write-Log ("GP History backup failed: {0}" -f $_.Exception.Message) 'WARN' }

    # Clear the contents of both folders, keeping the folders themselves.
    $foldersCleared = @()
    foreach ($src in @($gpRoot, $gpUsers)) {
        if (Test-Path $src) {
            Get-ChildItem -Path $src -Force -ErrorAction SilentlyContinue | ForEach-Object {
                Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
            }
            $foldersCleared += $src
        }
    }

    # Remove the History key and the children of the State key.
    $historyRemoved = $false
    $histPs = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\History'
    if (Test-Path $histPs) {
        Remove-Item -Path $histPs -Recurse -Force -ErrorAction SilentlyContinue
        $historyRemoved = $true
    }
    $statePs = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State'
    if (Test-Path $statePs) {
        Get-ChildItem -Path $statePs -Force -ErrorAction SilentlyContinue | ForEach-Object {
            Remove-Item -Path $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    # Re-assert the hybrid re-registration block (Intune-delivered Policies hive left alone).
    $wpj = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WorkplaceJoin'
    if (-not (Test-Path $wpj)) { New-Item -Path $wpj -Force | Out-Null }
    New-ItemProperty -Path $wpj -Name 'BlockAADWorkplaceJoin' -Value 1 -PropertyType DWord -Force | Out-Null

    return @{ FoldersCleared = $foldersCleared; HistoryRemoved = $historyRemoved }
}

# --------------------------------------------------------------------------
# Final report console banner
# --------------------------------------------------------------------------
function Write-ECReportBanner {
    param([Parameter(Mandatory)]$Report)
    $line = '=' * 70
    Write-Host ''
    Write-Host "  $line" -ForegroundColor Green
    Write-Host '  MIGRATION COMPLETE' -ForegroundColor Green
    Write-Host "  $line" -ForegroundColor Green
    Write-Host ("  Computer          : {0}" -f $Report.Computer)
    Write-Host ("  Run               : {0}" -f $Report.RunId)
    Write-Host ("  Duration (min)    : {0}" -f $Report.DurationMinutes)
    Write-Host ("  Former domain     : {0}" -f $Report.Domain)
    Write-Host ("  Old device id     : {0}" -f $Report.OldDeviceId)
    Write-Host ("  New device id     : {0}" -f $Report.NewDeviceId)
    Write-Host ("  Tenant id         : {0}" -f $Report.TenantId)
    Write-Host ("  Entra join verified: {0}" -f $Report.JoinVerified)
    Write-Host ("  Enrollment confirmed: {0}" -f $Report.EnrollmentConfirmed)
    Write-Host ("  BitLocker escrowed : {0}" -f $Report.BitLockerEscrowed)
    Write-Host ("  Break-glass account: {0} (retires {1})" -f $Report.BreakGlassAccount, $Report.BreakGlassRetiresUtc)
    if ($Report.Warning) { Write-Host ("  WARNING           : {0}" -f $Report.Warning) -ForegroundColor Yellow }
    Write-Host ''
    Write-Host '  TENANT CLEANUP CHECKLIST' -ForegroundColor Cyan
    foreach ($item in @($Report.TenantCleanupChecklist)) { Write-Host "    $item" }
    Write-Host ''
    Write-Host '  FIRST SIGN-IN CHECKLIST' -ForegroundColor Cyan
    foreach ($item in @($Report.FirstSignInChecklist)) { Write-Host "    - $item" }
    Write-Host ''
}

# ==========================================================================
# ENTRY POINT
# ==========================================================================
function Invoke-PhaseFinalize {
    param([hashtable]$Ctx)

    # --- 1. Escrow BitLocker to the new device object (HARD FAIL) -----------
    Invoke-Step -Ctx $Ctx -Name 'Finalize.EscrowBitLocker' -Action {
        $r = Invoke-ECBitLockerEscrow
        if ($r.AnyFailed) {
            # Recovery-key lockout is the worst outcome; leave break-glass live and stop.
            throw 'One or more BitLocker recovery keys failed to escrow to the new Entra device object - operator intervention required.'
        }
        return @{ Volumes = @($r.Volumes) }
    }

    # --- 2. Resume BitLocker if Prepare suspended it ------------------------
    Invoke-Step -Ctx $Ctx -Name 'Finalize.ResumeBitLocker' -Action {
        $prep = Get-ECStepData $Ctx.State 'Prepare.SuspendBitLocker'
        $wasProtected = Get-ECProp $prep 'WasProtected'
        if (-not $wasProtected) {
            Write-Log 'BitLocker was not suspended by Prepare - nothing to resume.' 'INFO'
            return @{ Resumed = $false }
        }
        try {
            Resume-BitLocker -MountPoint $env:SystemDrive -ErrorAction Stop | Out-Null
            Write-Log 'BitLocker protection resumed.'
            return @{ Resumed = $true }
        }
        catch {
            # Rethrow only if the volume genuinely reports still-suspended protection.
            $vol = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction SilentlyContinue
            $stillSuspended = ($null -ne $vol) -and ("$($vol.ProtectionStatus)" -ne 'On')
            if ($stillSuspended) { throw }
            Write-Log ("Resume-BitLocker reported '{0}' but protection is On - continuing." -f $_.Exception.Message) 'WARN'
            return @{ Resumed = $true }
        }
    }

    # --- 3. Stale GPO cleanup (non-fatal) -----------------------------------
    Invoke-Step -Ctx $Ctx -Name 'Finalize.GpoCleanup' -AllowFail -Action {
        return (Clear-ECStaleGpo -Ctx $Ctx)
    }

    # --- 4. Restore the original legal notice -------------------------------
    Invoke-Step -Ctx $Ctx -Name 'Finalize.ClearNotice' -Action {
        $notice = Get-ECStepData $Ctx.State 'Prepare.LegalNotice'
        $cap = Get-ECProp $notice 'OriginalCaption'; if ($null -eq $cap) { $cap = '' }
        $txt = Get-ECProp $notice 'OriginalText';    if ($null -eq $txt) { $txt = '' }
        Set-MigrationNotice -Caption $cap -Text $txt
        return @{ CaptionRestored = $cap }
    }

    # --- 5. Break-glass retention -------------------------------------------
    Invoke-Step -Ctx $Ctx -Name 'Finalize.BreakGlassRetention' -Action {
        $retDays = Get-ECProp $Ctx.Options 'FallbackRetentionDays'
        if ($null -eq $retDays) { $retDays = 0 }
        if ($retDays -gt 0) {
            $retireAt = (Get-Date).AddDays($retDays)
            # Literal $false in the child command (backtick-escaped in this string).
            $inner  = "Remove-LocalUser -Name '$($Script:BreakGlassUser)' -ErrorAction SilentlyContinue; Unregister-ScheduledTask -TaskName '$($Script:RetireTask)' -Confirm:`$false"
            $action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument ('-NoProfile -Command "{0}"' -f $inner)
            $trigger   = New-ScheduledTaskTrigger -Once -At $retireAt
            $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
            Register-ScheduledTask -TaskName $Script:RetireTask -Action $action -Trigger $trigger -Principal $principal -Force | Out-Null
            $utc = $retireAt.ToUniversalTime().ToString('o')
            Write-Log ("break-glass '{0}' retires at {1} (task '{2}')." -f $Script:BreakGlassUser, $utc, $Script:RetireTask)
            return @{ RetireAtUtc = $utc }
        }
        Remove-LocalUser -Name $Script:BreakGlassUser -ErrorAction SilentlyContinue
        Write-Log ("break-glass '{0}' removed now (retention 0 days)." -f $Script:BreakGlassUser)
        return @{ RemovedNow = $true }
    }

    # --- 6. Remove the resume task ------------------------------------------
    Invoke-Step -Ctx $Ctx -Name 'Finalize.RemoveResumeTask' -Action {
        Unregister-ResumeTask
        return @{ Removed = $true }
    }

    # --- 7. Final report ----------------------------------------------------
    Invoke-Step -Ctx $Ctx -Name 'Finalize.Report' -Action {
        $state  = $Ctx.State
        $device = Get-ECProp $state 'Device'

        $runId       = Get-ECProp $state 'RunId'
        $computer    = Get-ECProp $state 'Computer'; if ($null -eq $computer) { $computer = $env:COMPUTERNAME }
        $startedUtc  = Get-ECProp $state 'StartedUtc'
        $finishedUtc = (Get-Date).ToUniversalTime().ToString('o')

        $durationMin = 0
        if ($startedUtc) {
            try {
                $span = ([datetime]$finishedUtc) - ([datetime]$startedUtc)
                $durationMin = [math]::Round($span.TotalMinutes, 1)
            }
            catch { $durationMin = 0 }
        }

        $oldDeviceId = Get-ECProp $device 'OldDeviceId'
        $newDeviceId = Get-ECProp $device 'NewDeviceId'
        $tenantId    = Get-ECProp $device 'TenantId'
        $domain      = Get-ECProp $device 'Domain'

        # Fresh join verification.
        $dsreg = Get-DsregStatus
        $aadJoined = Get-ECProp $dsreg 'AzureAdJoined'
        $domJoined = Get-ECProp $dsreg 'DomainJoined'
        $joinVerified = (("$aadJoined" -eq 'YES') -and ("$domJoined" -eq 'NO'))

        # Enrollment confirmation (Join.MdmEnrollment may be Skipped).
        $steps    = Get-ECProp $state 'Steps'
        $mdmStep  = Get-ECProp $steps 'Join.MdmEnrollment'
        $mdmStatus = Get-ECProp $mdmStep 'Status'
        $mdmData   = Get-ECProp $mdmStep 'Data'
        $enrollConfirmed = $false
        $warning = $null
        if ($mdmStatus -eq 'Skipped' -or $null -eq $mdmStatus) {
            $enrollConfirmed = $false
        }
        else {
            $flag = Get-ECProp $mdmData 'EnrollmentConfirmed'
            if ($null -eq $flag) { $flag = Get-ECProp $mdmData 'Confirmed' }
            if ($null -eq $flag) { $flag = Get-ECProp $mdmData 'Enrolled' }
            if ($null -ne $flag) { $enrollConfirmed = [bool]$flag }
            elseif ($mdmStatus -eq 'Completed') { $enrollConfirmed = $true }
        }
        if (-not $enrollConfirmed) {
            $warning = 'MDM enrollment unconfirmed - check Intune portal / re-run deviceenroller'
        }

        # BitLocker escrow outcome.
        $escrowData = Get-ECStepData $state 'Finalize.EscrowBitLocker'
        $escrowVols = @(Get-ECProp $escrowData 'Volumes')
        $blEscrowed = $false
        foreach ($v in $escrowVols) { if (Get-ECProp $v 'Escrowed') { $blEscrowed = $true } }

        # Break-glass retirement time.
        $retData   = Get-ECStepData $state 'Finalize.BreakGlassRetention'
        $retiresUtc = Get-ECProp $retData 'RetireAtUtc'

        $tenantChecklist = @(
            '1. Disable the on-prem AD computer account (do this FIRST or Entra Connect resurrects the old object)',
            ("2. Verify BitLocker recovery key visible on NEW device object {0} in Entra" -f $newDeviceId),
            ("3. Delete stale hybrid device object {0} in Entra (ONLY after step 2)" -f $oldDeviceId),
            '4. Assign primary user in Intune (bulk-token enrollments have none)',
            '5. Confirm device compliance + Conditional Access',
            '6. Switch Autopilot deployment profile to Entra-joined for this device',
            '7. Move/delete the AD computer account per AD hygiene'
        )
        $firstSignInChecklist = @(
            'Sign in with Entra ID (UPN)',
            'Re-enroll Windows Hello (PIN/biometrics)',
            'OneDrive: verify sign-in + KFM folders',
            'Re-auth Teams / Authenticator',
            'Saved browser/Wi-Fi/Credential Manager passwords are gone by design - see known-loss register',
            'Old profile data remains at C:\Users\<olduser> (read-only for the new account by default)'
        )

        $report = [pscustomobject]@{
            RunId                 = $runId
            Computer              = $computer
            StartedUtc            = $startedUtc
            FinishedUtc           = $finishedUtc
            DurationMinutes       = $durationMin
            OldDeviceId           = $oldDeviceId
            NewDeviceId           = $newDeviceId
            TenantId              = $tenantId
            Domain                = $domain
            JoinVerified          = $joinVerified
            EnrollmentConfirmed   = $enrollConfirmed
            Warning               = $warning
            BitLockerEscrowed     = $blEscrowed
            BreakGlassAccount     = $Script:BreakGlassUser
            BreakGlassRetiresUtc  = $retiresUtc
            TenantCleanupChecklist = $tenantChecklist
            FirstSignInChecklist  = $firstSignInChecklist
        }

        # Persist to state and to a standalone report.json.
        Set-ECStateProp $state 'Result' $report
        Save-CutoverState -State $state

        $reportPath = Join-Path $Ctx.Paths.Root 'report.json'
        try {
            $report | ConvertTo-Json -Depth 6 | Set-Content -Path $reportPath -Encoding UTF8 -ErrorAction Stop
            Write-Log ("final report written: {0}" -f $reportPath) 'SUCCESS'
        }
        catch { Write-Log ("report.json write failed: {0}" -f $_.Exception.Message) 'WARN' }

        if ($warning) { Write-Log $warning 'WARN' }
        Write-ECReportBanner -Report $report

        return @{ ReportPath = $reportPath; JoinVerified = $joinVerified }
    }

    # No reboot at end - migration is complete.
}
