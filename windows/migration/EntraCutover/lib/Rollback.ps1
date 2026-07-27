#Requires -Version 5.1
<#
    Rollback. Branches on live device state:

      REFUSE    - the new Entra join already succeeded: forward-only.
      PRE-PONR  - still domain-joined: undo Prepare in reverse.
      POST-PONR - stranded in a workgroup: offline-rejoin the domain from the
                  Prepare djoin blob (or emit manual guidance).

    Entry point: Invoke-CutoverRollback -Ctx <hashtable>.
    Private helpers are named <Verb>-EC<Noun> per the lib contract.
#>

Set-StrictMode -Version 2.0

# --------------------------------------------------------------------------
# State accessors (state.json round-trips to PSCustomObject; null-guard all)
# --------------------------------------------------------------------------
function Get-ECProp {
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
    param($State, [string]$StepName)
    $steps = Get-ECProp $State 'Steps'
    $step  = Get-ECProp $steps $StepName
    return (Get-ECProp $step 'Data')
}

function Set-ECStateProp {
    param($Object, [string]$Name, $Value)
    try { $Object[$Name] = $Value }
    catch { $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force }
}

# --------------------------------------------------------------------------
# Shared inverse actions (used by both PRE- and POST-PONR branches)
# --------------------------------------------------------------------------
function Restore-ECLegalNotice {
    param($State)
    $d = Get-ECStepData $State 'Prepare.LegalNotice'
    $cap = Get-ECProp $d 'OriginalCaption'; if ($null -eq $cap) { $cap = '' }
    $txt = Get-ECProp $d 'OriginalText';    if ($null -eq $txt) { $txt = '' }
    Set-MigrationNotice -Caption $cap -Text $txt
}

function Resume-ECBitLockerIfProtected {
    param($State)
    $d = Get-ECStepData $State 'Prepare.SuspendBitLocker'
    if (Get-ECProp $d 'WasProtected') {
        Resume-BitLocker -MountPoint $env:SystemDrive -ErrorAction Stop | Out-Null
        Write-Log 'BitLocker protection resumed.'
    }
    else { Write-Log 'BitLocker was not suspended by Prepare - nothing to resume.' 'INFO' }
}

function Restore-ECWorkplaceJoin {
    # Restore the WorkplaceJoin registry values captured by Prepare.
    # Prepare.BlockHybridRejoin records the PRIOR values as:
    #   PrevBlockAADWorkplaceJoin / PrevAutoWorkplaceJoin  ('absent' | dword string)
    # Map each back onto its real value name; 'absent' means "remove".
    param($BlockData)
    if ($null -eq $BlockData) { Write-Log 'No BlockHybridRejoin data recorded - skipping WorkplaceJoin restore.' 'WARN'; return }
    $wpjKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WorkplaceJoin'

    $map = @(
        @{ Name = 'BlockAADWorkplaceJoin'; Prev = (Get-ECProp $BlockData 'PrevBlockAADWorkplaceJoin') },
        @{ Name = 'autoWorkplaceJoin';     Prev = (Get-ECProp $BlockData 'PrevAutoWorkplaceJoin') }
    )

    foreach ($rec in $map) {
        $name = $rec.Name
        $val  = $rec.Prev
        try {
            if (($null -eq $val) -or ("$val" -eq 'absent')) {
                if (Test-Path $wpjKey) { Remove-ItemProperty -Path $wpjKey -Name $name -Force -ErrorAction SilentlyContinue }
                Write-Log ("WorkplaceJoin value removed (was absent pre-migration): {0}\{1}" -f $wpjKey, $name)
            }
            else {
                if (-not (Test-Path $wpjKey)) { New-Item -Path $wpjKey -Force | Out-Null }
                New-ItemProperty -Path $wpjKey -Name $name -Value ([int]$val) -PropertyType DWord -Force | Out-Null
                Write-Log ("WorkplaceJoin value restored: {0}\{1}={2}" -f $wpjKey, $name, $val)
            }
        }
        catch { Write-Log ("WorkplaceJoin restore failed for {0}: {1}" -f $name, $_.Exception.Message) 'WARN' }
    }
}

function Enable-ECAutoDeviceJoinTask {
    # Re-enable the Automatic-Device-Join task if Prepare recorded it enabled.
    # Prepare.BlockHybridRejoin records this as 'TaskWasEnabled'.
    param($BlockData)
    $wasEnabled = Get-ECProp $BlockData 'TaskWasEnabled'
    if (-not $wasEnabled) {
        Write-Log 'Automatic-Device-Join was not recorded as enabled - leaving as-is.' 'INFO'
        return
    }
    try {
        Enable-ScheduledTask -TaskName 'Automatic-Device-Join' -TaskPath '\Microsoft\Windows\Workplace Join\' -ErrorAction SilentlyContinue | Out-Null
        Write-Log 'Re-enabled Automatic-Device-Join scheduled task.'
    }
    catch { Write-Log ("Enable Automatic-Device-Join failed: {0}" -f $_.Exception.Message) 'WARN' }
}

function Invoke-ECOfflineDomainJoin {
    # Apply the Prepare djoin rollback blob. Returns the Invoke-Exe result.
    # $env:SystemRoot must be passed already-expanded: Invoke-Exe launches via
    # ProcessStartInfo, which (unlike cmd.exe) does NOT expand %SystemRoot%.
    param([string]$BlobPath)
    return (Invoke-Exe -Path "$env:windir\System32\djoin.exe" -Arguments @(
        '/requestODJ', '/loadfile', ('"{0}"' -f $BlobPath), '/windowspath', ('"{0}"' -f $env:SystemRoot), '/localos'
    ))
}

# ==========================================================================
# ENTRY POINT
# ==========================================================================
function Invoke-CutoverRollback {
    param([hashtable]$Ctx)

    $state = $Ctx.State
    $device = Get-ECProp $state 'Device'

    # Live domain membership.
    $cs = Get-WmiObject -Class Win32_ComputerSystem -ErrorAction SilentlyContinue
    $partOfDomain = $false
    if ($cs -and $cs.PartOfDomain) { $partOfDomain = [bool]$cs.PartOfDomain }

    $dsreg = Get-DsregStatus
    $aadJoined = Get-ECProp $dsreg 'AzureAdJoined'
    $newDeviceId = Get-ECProp $device 'NewDeviceId'

    # ----- REFUSE: the new Entra join already succeeded ---------------------
    if (("$aadJoined" -eq 'YES') -and (-not $partOfDomain) -and $newDeviceId) {
        Write-Log 'Migration has completed the Entra join - rollback is forward-only from here. Use tenant-side remediation.' 'ERROR'
        return
    }

    $nowUtc = (Get-Date).ToUniversalTime().ToString('o')
    $blob = Join-Path $Ctx.Paths.Rollback 'odj.blob'
    $blockData = Get-ECStepData $state 'Prepare.BlockHybridRejoin'

    if ($partOfDomain) {
        # ===== PRE-PONR: undo Prepare in reverse; continue on failure =======
        Write-Log 'Pre-PONR rollback: undoing Prepare (device still domain-joined).' 'STEP'

        try { Restore-ECLegalNotice -State $state }
        catch { Write-Log ("legal-notice restore failed: {0}" -f $_.Exception.Message) 'WARN' }

        try { Resume-ECBitLockerIfProtected -State $state }
        catch { Write-Log ("BitLocker resume failed: {0}" -f $_.Exception.Message) 'WARN' }

        try { Restore-ECWorkplaceJoin -BlockData $blockData }
        catch { Write-Log ("WorkplaceJoin restore failed: {0}" -f $_.Exception.Message) 'WARN' }

        try { Enable-ECAutoDeviceJoinTask -BlockData $blockData }
        catch { Write-Log ("Automatic-Device-Join re-enable failed: {0}" -f $_.Exception.Message) 'WARN' }

        # If Prepare's /reuse blob was staged, the AD machine password was reset -
        # machine trust is broken; re-apply the blob and reboot to restore it.
        $blobData = Get-ECStepData $state 'Prepare.RollbackBlob'
        $blobRan  = Get-ECProp $blobData 'BlobPresent'
        if ((Test-Path $blob) -and $blobRan) {
            try {
                $r = Invoke-ECOfflineDomainJoin -BlobPath $blob
                Write-Log ("djoin /requestODJ applied (exit {0}); a reboot restores the machine trust relationship." -f $r.ExitCode) 'WARN'
                Request-Reboot -Ctx $Ctx -Reason 'reapply offline-join blob to restore broken machine trust'
            }
            catch { Write-Log ("djoin /requestODJ failed: {0}" -f $_.Exception.Message) 'WARN' }
        }

        try { Remove-LocalUser -Name $Script:BreakGlassUser -ErrorAction SilentlyContinue }
        catch { Write-Log ("break-glass removal failed: {0}" -f $_.Exception.Message) 'WARN' }

        try { Unregister-ResumeTask }
        catch { Write-Log ("resume-task removal failed: {0}" -f $_.Exception.Message) 'WARN' }
        try { Unregister-ScheduledTask -TaskName $Script:RetireTask -Confirm:$false -ErrorAction SilentlyContinue }
        catch { Write-Log ("retire-task removal failed: {0}" -f $_.Exception.Message) 'WARN' }

        Set-ECStateProp $state 'Result' @{ Outcome = 'RolledBack-PrePonr'; WhenUtc = $nowUtc }
        Set-ECStateProp $state 'NextPhase' $null
        Save-CutoverState -State $state
        Write-Log 'Pre-PONR rollback complete - Prepare changes reverted.' 'SUCCESS'
        return
    }

    # ===== POST-PONR: stranded in a workgroup ===============================
    Write-Log 'Post-PONR rollback: device is in a workgroup (domain unjoined, Entra join not verified).' 'STEP'
    $domain = Get-ECProp $device 'Domain'

    if (-not (Test-Path $blob)) {
        Write-Log ("Rollback blob missing - offline domain rejoin unavailable. Rejoin manually: Add-Computer -DomainName {0} -Credential <cred> -Restart; break-glass account {1} remains available." -f $domain, $Script:BreakGlassUser) 'ERROR'
        Set-ECStateProp $state 'Result' @{ Outcome = 'RollbackUnavailable' }
        Set-ECStateProp $state 'NextPhase' $null
        Save-CutoverState -State $state
        return
    }

    $r = Invoke-ECOfflineDomainJoin -BlobPath $blob
    if ($r.ExitCode -ne 0) {
        Write-Log ("Offline domain rejoin failed - djoin /requestODJ exit {0}. Manual rejoin: Add-Computer -DomainName {1} -Credential <cred> -Restart; break-glass account {2} remains available." -f $r.ExitCode, $domain, $Script:BreakGlassUser) 'ERROR'
        Set-ECStateProp $state 'Result' @{ Outcome = 'RollbackUnavailable' }
        Set-ECStateProp $state 'NextPhase' $null
        Save-CutoverState -State $state
        return
    }

    # Success: restore hybrid-rejoin path so the device re-registers on next sync.
    try { Restore-ECWorkplaceJoin -BlockData $blockData }
    catch { Write-Log ("WorkplaceJoin restore failed: {0}" -f $_.Exception.Message) 'WARN' }
    try { Enable-ECAutoDeviceJoinTask -BlockData $blockData }
    catch { Write-Log ("Automatic-Device-Join re-enable failed: {0}" -f $_.Exception.Message) 'WARN' }
    try { Restore-ECLegalNotice -State $state }
    catch { Write-Log ("legal-notice restore failed: {0}" -f $_.Exception.Message) 'WARN' }
    try { Resume-ECBitLockerIfProtected -State $state }
    catch { Write-Log ("BitLocker resume failed: {0}" -f $_.Exception.Message) 'WARN' }

    # Keep the break-glass account as the safety net for the rejoined device.
    Write-Log ("Break-glass account {0} is retained for post-rollback access." -f $Script:BreakGlassUser) 'WARN'

    try { Unregister-ResumeTask }
    catch { Write-Log ("resume-task removal failed: {0}" -f $_.Exception.Message) 'WARN' }

    # Intune enrollment purge is intentionally NOT reverted here.
    Write-Log 'Follow-up: Intune enrollment purge is not reverted; the device re-registers hybrid via the re-enabled Automatic-Device-Join task on next sync and can be re-enrolled with the documented deviceenroller flow.' 'WARN'

    Set-ECStateProp $state 'Result' @{ Outcome = 'RolledBackToDomain'; WhenUtc = $nowUtc }
    Set-ECStateProp $state 'NextPhase' $null
    Save-CutoverState -State $state
    Request-Reboot -Ctx $Ctx -Reason 'offline domain rejoin applied'
}
