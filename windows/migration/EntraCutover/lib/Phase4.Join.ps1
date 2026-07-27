#Requires -Version 5.1
<#
    Phase 4 - JOIN (runs headless as SYSTEM from the resume startup task, after
    the device has been unjoined from its AD domain).

    Entra join via bulk-token provisioning package (or user-driven), MDM
    enrollment, and automatic djoin rollback-to-domain after a repeated join
    failure. Entry point: Invoke-PhaseJoin. Shared helpers (Write-Log,
    Invoke-Step, Set-StepState, Save-CutoverState, Invoke-Exe, Get-DsregStatus,
    Set-MigrationNotice, Request-Reboot) and Invoke-CutoverRollback are provided
    by the host script / sibling lib files at runtime.
#>

Set-StrictMode -Version 2.0

# --------------------------------------------------------------------------
# State-shape helpers  (State.Device / State.Options round-trip to
# PSCustomObject through JSON; hashtable on a fresh in-memory run).
# --------------------------------------------------------------------------
function Get-ECProp {
    # Null-guarded read that works for both hashtable and PSCustomObject.
    param($Obj, [Parameter(Mandatory)][string]$Name)
    if ($null -eq $Obj) { return $null }
    if ($Obj -is [System.Collections.IDictionary]) {
        if ($Obj.Contains($Name)) { return $Obj[$Name] }
        return $null
    }
    $prop = $Obj.PSObject.Properties[$Name]
    if ($null -ne $prop) { return $prop.Value }
    return $null
}

function Set-ECProp {
    # Try-indexer-then-Add-Member -Force: indexer for dictionaries, Add-Member
    # for PSCustomObject (round-tripped) shapes.
    param([Parameter(Mandatory)]$Obj, [Parameter(Mandatory)][string]$Name, $Value)
    try { $Obj[$Name] = $Value }
    catch { $Obj | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force }
}

function Get-ECDevice {
    # Return $Ctx.State.Device, creating it if absent so writes have a target.
    param([Parameter(Mandatory)][hashtable]$Ctx)
    $dev = Get-ECProp -Obj $Ctx.State -Name 'Device'
    if ($null -eq $dev) {
        $dev = New-Object -TypeName psobject
        Set-ECProp -Obj $Ctx.State -Name 'Device' -Value $dev
    }
    return $dev
}

function Get-ECStateOption {
    # The resume task re-invokes with only -Mode Resume -Force, so the live
    # $Ctx.Options loses JoinMode/PpkgPath/TenantId across reboots. Prefer the
    # persisted State.Options; fall back to the live hashtable.
    param([Parameter(Mandatory)][hashtable]$Ctx, [Parameter(Mandatory)][string]$Name)
    $val = $null
    $so = Get-ECProp -Obj $Ctx.State -Name 'Options'
    if ($null -ne $so) { $val = Get-ECProp -Obj $so -Name $Name }
    $empty = ($null -eq $val) -or (($val -is [string]) -and ($val -eq ''))
    if ($empty) {
        $opt = $Ctx.Options
        if (($null -ne $opt) -and ($opt -is [System.Collections.IDictionary]) -and $opt.Contains($Name)) {
            $val = $opt[$Name]
        }
    }
    return $val
}

function Get-ECDsregValue {
    # Read one key out of the Get-DsregStatus hashtable, null-guarded.
    param($Status, [Parameter(Mandatory)][string]$Key)
    if (($null -ne $Status) -and ($Status -is [System.Collections.IDictionary]) -and $Status.Contains($Key)) {
        return $Status[$Key]
    }
    return $null
}

# --------------------------------------------------------------------------
# Network / join polling
# --------------------------------------------------------------------------
function Wait-ECNetwork {
    # Poll HTTPS reachability of login.microsoftonline.com. Post-unjoin there is
    # no domain auth - plain internet egress is all we need. Any HTTP status
    # (incl. 4xx/5xx) counts as reachable; only connect/DNS/timeout is "down".
    $url = 'https://login.microsoftonline.com'
    $deadline = (Get-Date).AddMinutes(10)
    try {
        [Net.ServicePointManager]::SecurityProtocol = `
            [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    }
    catch { Write-Verbose "TLS1.2 enable failed: $($_.Exception.Message)" }

    while ((Get-Date) -lt $deadline) {
        $reachable = $false
        try {
            $null = Invoke-WebRequest -Uri $url -Method Head -TimeoutSec 10 -UseBasicParsing -ErrorAction Stop
            $reachable = $true
        }
        catch {
            # A returned HTTP response (any status) still proves reachability.
            $resp = $null
            try { $resp = $_.Exception.Response } catch { $resp = $null }
            if ($null -ne $resp) { $reachable = $true }
        }
        if ($reachable) {
            Write-Log 'network reachable (login.microsoftonline.com).' 'SUCCESS'
            return @{ Reachable = $true }
        }
        Write-Log 'network not yet reachable; retrying in 30s.' 'INFO'
        Start-Sleep -Seconds 30
    }
    throw 'network unreachable after 10 minutes (login.microsoftonline.com)'
}

function Wait-ECJoined {
    # Poll dsregcmd every 30s up to $TimeoutMin for AzureAdJoined = YES.
    param([int]$TimeoutMin = 15)
    $deadline = (Get-Date).AddMinutes($TimeoutMin)
    while ((Get-Date) -lt $deadline) {
        $ds = Get-DsregStatus
        if ((Get-ECDsregValue -Status $ds -Key 'AzureAdJoined') -eq 'YES') { return $true }
        Start-Sleep -Seconds 30
    }
    return $false
}

# --------------------------------------------------------------------------
# Ppkg join with retry budget  (own bookkeeping - NOT wrapped in a single
# Invoke-Step, so a failed attempt never poisons resume).
# Returns $true to proceed to verification, $false to end the phase (a
# retry-reboot is pending, or rollback was invoked).
# --------------------------------------------------------------------------
function Invoke-ECPpkgJoin {
    param([Parameter(Mandatory)][hashtable]$Ctx)

    # Already joined (e.g. package applied on a prior boot)? -> verify/capture.
    $ds0 = Get-DsregStatus
    if ((Get-ECDsregValue -Status $ds0 -Key 'AzureAdJoined') -eq 'YES') {
        Write-Log 'device already Entra joined; proceeding to verification.' 'INFO'
        return $true
    }

    # Attempt bookkeeping - persist BEFORE attempting so a crash still counts.
    $dev = Get-ECDevice -Ctx $Ctx
    $attempts = 0
    $prev = Get-ECProp -Obj $dev -Name 'JoinAttempts'
    if ($null -ne $prev) { try { $attempts = [int]$prev } catch { $attempts = 0 } }
    $attempts = $attempts + 1
    Set-ECProp -Obj $dev -Name 'JoinAttempts' -Value $attempts
    Save-CutoverState -State $Ctx.State

    # Locate the ppkg staged by Prepare (Root\join.ppkg).
    $ppkg = Get-ECStateOption -Ctx $Ctx -Name 'PpkgPath'
    if ((-not $ppkg) -or (-not (Test-Path -LiteralPath $ppkg))) {
        $staged = Join-Path $Ctx.Paths.Root 'join.ppkg'
        if (Test-Path -LiteralPath $staged) { $ppkg = $staged }
    }
    if ((-not $ppkg) -or (-not (Test-Path -LiteralPath $ppkg))) {
        throw 'staged ppkg not found'
    }

    # Apply the provisioning package. A non-zero exit is logged but not fatal -
    # the package can partially apply and the join can still be in flight.
    $dism = Join-Path $env:windir 'System32\dism.exe'
    $r = Invoke-Exe -Path $dism -TimeoutSec 600 -Arguments @(
        '/Online', '/Add-ProvisioningPackage', ('/PackagePath:"{0}"' -f $ppkg), '/Quiet'
    )
    if ($r.ExitCode -ne 0) {
        Write-Log ("dism /Add-ProvisioningPackage returned exit {0}; continuing to poll (join may be in flight)." -f $r.ExitCode) 'ERROR'
    }

    if (Wait-ECJoined -TimeoutMin 15) {
        Set-StepState -State $Ctx.State -Name 'Join.JoinDevice' -Status 'Completed' -Data @{ Attempts = $attempts }
        Write-Log ("Entra join succeeded on attempt {0}." -f $attempts) 'SUCCESS'
        return $true
    }

    # Timed out. Retry budget is 2 attempts.
    if ($attempts -lt 2) {
        Write-Log ("join attempt {0} failed; retrying after reboot." -f $attempts) 'ERROR'
        Request-Reboot -Ctx $Ctx -Reason ('Entra join attempt {0} did not complete; retrying after restart' -f $attempts)
        $Ctx.StayInPhase = $true
        return $false
    }

    # Budget exhausted -> rollback to the domain if we have the offline blob.
    $blob = Join-Path $Ctx.Paths.Rollback 'odj.blob'
    if (Test-Path -LiteralPath $blob) {
        Write-Log ("Entra join failed after {0} attempts - rolling back to domain." -f $attempts) 'ERROR'
        Invoke-CutoverRollback -Ctx $Ctx   # sets Result / NextPhase=$null / reboot
        return $false
    }
    throw ('Entra join failed after {0} attempts and no rollback blob exists - manual intervention required (break-glass account is available)' -f $attempts)
}

# --------------------------------------------------------------------------
# MDM enrollment  (registry key is the source of truth; Event 75 corroborates).
# --------------------------------------------------------------------------
function Get-ECMdmEnrollmentId {
    # GUID of an existing 'MS DM Server' enrollment under Enrollments, or $null.
    $root = 'HKLM:\SOFTWARE\Microsoft\Enrollments'
    if (-not (Test-Path -LiteralPath $root)) { return $null }
    $guidRe = '^[0-9A-Fa-f]{8}-([0-9A-Fa-f]{4}-){3}[0-9A-Fa-f]{12}$'
    $children = @(Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue)
    foreach ($child in $children) {
        if ($child.PSChildName -notmatch $guidRe) { continue }
        $provider = $null
        try {
            $item = Get-ItemProperty -LiteralPath $child.PSPath -Name 'ProviderID' -ErrorAction SilentlyContinue
            if ($null -ne $item) { $provider = $item.ProviderID }
        }
        catch { $provider = $null }
        if ($provider -eq 'MS DM Server') { return $child.PSChildName }
    }
    return $null
}

function Set-ECMdmDiscoveryUrls {
    # Ensure the Intune MDM discovery URLs on the tenant's CloudDomainJoin key.
    # Returns $true when written. Match Options.TenantId if set, else first.
    param([Parameter(Mandatory)][hashtable]$Ctx)
    $base = 'HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\TenantInfo'
    if (-not (Test-Path -LiteralPath $base)) { return $false }
    $guidRe = '^\{?[0-9A-Fa-f]{8}-([0-9A-Fa-f]{4}-){3}[0-9A-Fa-f]{12}\}?$'
    $wantTid = Get-ECStateOption -Ctx $Ctx -Name 'TenantId'
    $wantNorm = $null
    if ($wantTid) { $wantNorm = ($wantTid -replace '[{}]', '').ToLowerInvariant() }

    $target = $null
    $children = @(Get-ChildItem -LiteralPath $base -ErrorAction SilentlyContinue)
    foreach ($child in $children) {
        if ($child.PSChildName -notmatch $guidRe) { continue }
        if ($wantNorm) {
            if (($child.PSChildName -replace '[{}]', '').ToLowerInvariant() -eq $wantNorm) { $target = $child; break }
        }
        else { $target = $child; break }
    }
    if ($null -eq $target) { return $false }

    Set-ItemProperty -LiteralPath $target.PSPath -Name 'MdmEnrollmentUrl' -Type String -Force `
        -Value 'https://enrollment.manage.microsoft.com/enrollmentserver/discovery.svc'
    Set-ItemProperty -LiteralPath $target.PSPath -Name 'MdmTermsOfUseUrl' -Type String -Force `
        -Value 'https://portal.manage.microsoft.com/TermsofUse.aspx'
    Set-ItemProperty -LiteralPath $target.PSPath -Name 'MdmComplianceUrl' -Type String -Force `
        -Value 'https://portal.manage.microsoft.com/?portalAction=Compliance'
    Write-Log ("MDM discovery URLs set on TenantInfo\{0}." -f $target.PSChildName) 'INFO'
    return $true
}

function Test-ECEnrollEvent75 {
    # Informational corroboration only - the registry enrollment is authoritative.
    $found = $false
    try {
        $ev = Get-WinEvent -MaxEvents 1 -ErrorAction Stop -FilterHashtable @{
            LogName = 'Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin'
            Id      = 75
        }
        if ($null -ne $ev) { $found = $true }
    }
    catch { $found = $false }
    return $found
}

function Invoke-ECMdmEnroll {
    # Trigger + confirm auto-enrollment. Runs as SYSTEM (correct context).
    param([Parameter(Mandatory)][hashtable]$Ctx)

    $existing = Get-ECMdmEnrollmentId
    if ($existing) {
        Write-Log ("MDM enrollment already present (id {0})." -f $existing) 'INFO'
        return @{ Enrolled = $true; EnrollmentId = $existing; Event75 = (Test-ECEnrollEvent75) }
    }

    $null = Set-ECMdmDiscoveryUrls -Ctx $Ctx

    $enroller = Join-Path $env:windir 'System32\deviceenroller.exe'
    $null = Invoke-Exe -Path $enroller -TimeoutSec 300 -Arguments @('/c', '/AutoEnrollMDM')

    $deadline = (Get-Date).AddMinutes(10)
    $id = $null
    while ((Get-Date) -lt $deadline) {
        $id = Get-ECMdmEnrollmentId
        if ($id) { break }
        Start-Sleep -Seconds 30
    }

    $event75 = Test-ECEnrollEvent75
    if (-not $id) { throw 'MDM enrollment not confirmed' }

    Write-Log ("MDM enrollment confirmed (id {0}; event75={1})." -f $id, $event75) 'SUCCESS'
    return @{ Enrolled = $true; EnrollmentId = $id; Event75 = $event75 }
}

# --------------------------------------------------------------------------
# Phase entry point
# --------------------------------------------------------------------------
function Invoke-PhaseJoin {
    param([Parameter(Mandatory)][hashtable]$Ctx)

    # 1. Wait for plain internet egress before any join attempt.
    Invoke-Step -Ctx $Ctx -Name 'Join.WaitNetwork' -Action { Wait-ECNetwork }

    # 2. Join mechanism (persisted option - survives the resume re-invocation).
    $joinMode = Get-ECStateOption -Ctx $Ctx -Name 'JoinMode'
    if (-not $joinMode) { $joinMode = 'Ppkg' }

    if ($joinMode -eq 'UserDriven') {
        $ds = Get-DsregStatus
        if ((Get-ECDsregValue -Status $ds -Key 'AzureAdJoined') -ne 'YES') {
            Set-MigrationNotice -Caption 'Action required: join this device to Entra' `
                -Text 'Sign in with a local account, open Settings > Accounts > Access work or school > Connect, and join with your work account. The migration will continue automatically after the next restart.'
            Write-Log 'user-driven join pending: device is not yet Entra joined - staying in phase.' 'WARN'
            $Ctx.StayInPhase = $true
            return
        }
        Write-Log 'user-driven join detected as complete; proceeding to verification.' 'INFO'
    }
    else {
        # 3. Ppkg join with its own retry budget.
        $proceed = Invoke-ECPpkgJoin -Ctx $Ctx
        if (-not $proceed) { return }
    }

    # 4. Verify identity and capture the new device id.
    Invoke-Step -Ctx $Ctx -Name 'Join.VerifyIdentity' -Action {
        $ds = Get-DsregStatus
        $aadJoined  = Get-ECDsregValue -Status $ds -Key 'AzureAdJoined'
        $domJoined  = Get-ECDsregValue -Status $ds -Key 'DomainJoined'
        $tenantId   = Get-ECDsregValue -Status $ds -Key 'TenantId'
        $newDevice  = Get-ECDsregValue -Status $ds -Key 'DeviceId'

        if ($aadJoined -ne 'YES') { throw 'verify failed: AzureAdJoined is not YES' }
        if ($domJoined -eq 'YES') { throw 'verify failed: DomainJoined is still YES (expected NO after unjoin)' }

        $wantTid = Get-ECStateOption -Ctx $Ctx -Name 'TenantId'
        if ($wantTid -and ($tenantId -ne $wantTid)) {
            throw ("verify failed: joined tenant '{0}' does not match expected '{1}'" -f $tenantId, $wantTid)
        }

        Set-ECProp -Obj (Get-ECDevice -Ctx $Ctx) -Name 'NewDeviceId' -Value $newDevice
        Save-CutoverState -State $Ctx.State
        return @{ NewDeviceId = $newDevice; TenantId = $tenantId }
    }

    # 5. MDM enrollment - non-fatal (often completes minutes later).
    Invoke-Step -Ctx $Ctx -Name 'Join.MdmEnrollment' -AllowFail -Action { Invoke-ECMdmEnroll -Ctx $Ctx }

    # 6. Update the progress notice and reboot into Finalize.
    Set-MigrationNotice -Caption 'Device migration in progress' `
        -Text 'Almost done. Do NOT sign in until this notice is gone. Phase: Join complete.'
    Request-Reboot -Ctx $Ctx -Reason 'Entra join complete; finalizing after restart'
}
