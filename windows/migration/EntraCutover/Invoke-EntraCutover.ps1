#Requires -Version 5.1
<#
.SYNOPSIS
    [EXPERIMENTAL - NOT SUPPORTED BY MICROSOFT] In-place migration of a Windows
    device from Hybrid Entra Join (on-prem AD joined + Intune enrolled) to
    Entra-only join. No wipe, no Autopilot reprovision. CLI-only, CMTrace
    logging, resumable across reboots.

    This is community-built tooling that performs an operation Microsoft does
    NOT support in-place (the supported path is wipe + Autopilot). Treat it as
    experimental: validate on a throwaway test ring before any production use.

.DESCRIPTION
    UNSUPPORTED BY MICROSOFT. Microsoft's documented path is wipe + Autopilot
    ("Windows Autopilot for existing devices"). This tool implements the
    field-proven in-place technique used by commercial device-cutover products
    (ForensiT, Quest ODM, PowerSyncPro) and community frameworks (AADMigration).
    Design doc: docs/ENTRA-CUTOVER-TOOL-PLAN.md.

    Phase state machine (persisted to %ProgramData%\EntraCutover\state.json,
    resumed after each reboot by a SYSTEM startup scheduled task):

      Assess    Read-only preflight. Verdict: Ready / Blockers. Default mode.
      Prepare   Reversible staging: break-glass admin, djoin rollback blob,
                registry backups, block hybrid re-registration, suspend
                BitLocker, KFM gate, resume task, lock-screen notice.
      Teardown  DESTRUCTIVE: Intune enrollment purge, dsregcmd /leave, domain
                unjoin (point of no return), reboot.
      Join      SYSTEM: apply bulk-token provisioning package, verify Entra
                join, trigger/verify MDM enrollment, reboot. Automatic djoin
                rollback to domain after repeated join failure.
      Finalize  SYSTEM: BitLocker key escrow to the NEW device object, resume
                BitLocker, stale-GPO cleanup, clear notices, schedule
                break-glass retention removal, final report.

    User data strategy is fresh-profile + OneDrive Known Folder Move: the old
    domain profile is left on disk untouched; users sign in with their Entra ID
    and KFM re-hydrates Desktop/Documents/Pictures. KFM health is a Prepare
    gate. Guaranteed losses (DPAPI secrets, Windows Hello, per-app re-auth) are
    listed in the final report and the design doc's known-loss register.

    Tenant-side prerequisites (tool cannot do these; see design doc section 4):
    device scoped OUT of Entra Connect hybrid-join sync, bulk token ppkg with
    CA/MFA exclusion for the package_* account, licensing (Entra ID P1 +
    Intune, users in MDM scope), KFM policy deployed, Intune parity for GPO
    settings. Migrate mode refuses to start unless -AcknowledgePrereqs is set.

.PARAMETER Mode
    Assess   - read-only preflight audit (default).
    Migrate  - full migration (Assess gate first, then Prepare onward).
    Resume   - continue after reboot (used by the resume scheduled task).
    Status   - print current migration state from state.json.
    Rollback - revert. Pre-unjoin: undo Prepare. Post-unjoin: djoin blob
               rejoin to the domain.

.PARAMETER PpkgPath
    Path to the bulk-enrollment provisioning package (.ppkg) built with
    Windows Configuration Designer. Required for Migrate with -JoinMode Ppkg.

.PARAMETER JoinMode
    Ppkg (default) - silent Entra join via bulk token provisioning package.
    UserDriven     - stop after Teardown; user joins via Settings > Access
                     work or school. Immune to package_* CA issues and Token
                     Protection, but interactive.

.PARAMETER DomainCredential
    Credential for the graceful domain unjoin and the djoin rollback-blob
    provisioning. If omitted, Prepare skips the rollback blob (WARN) and
    Teardown attempts unjoin with the current user context, falling back to
    offline unjoin when -OfflineUnjoin is set.

.PARAMETER OfflineUnjoin
    Permit domain unjoin without DC line-of-sight (disables physical NICs,
    performs a local-only workgroup swap, re-enables NICs).

.PARAMETER BreakGlassCredential
    Password for the break-glass local admin the tool creates in Prepare. When
    supplied, the account is set to this KNOWN password - nothing is generated
    and nothing is echoed, so the operator already holds the credential (recover
    it from your own vault). Only the password is used; the account name is
    always the fixed local account (a supplied username is ignored with a
    warning). REQUIRED for unattended (-Force) Migrate runs: without it the tool
    would mint a random password recoverable only by an admin already on the box
    (a chicken-and-egg for the very lockout the account exists for). Attended
    Migrate runs may omit it and get the generate-and-show-once behavior.
    Day-2 backstop: hand this account to Windows LAPS (BackupDirectory=1) via
    Intune once the device is Entra-joined + enrolled; LAPS does not cover the
    migration window itself, which is why the specified password is still needed.

.PARAMETER TenantId
    Expected Entra tenant ID. Used to validate the KFM policy and the
    post-join device state.

.PARAMETER SkipKfmGate
    Proceed even when OneDrive KFM cannot be verified healthy. The operator
    accepts the data-loss risk for Desktop/Documents/Pictures.

.PARAMETER FallbackRetentionDays
    Days the break-glass local admin persists after Finalize. Default 7.

.PARAMETER AcknowledgePrereqs
    Required for Migrate. Asserts the tenant-side runbook (design doc section
    4) is complete: hybrid-join sync scoping, CA exclusion, licensing, KFM
    policy, Intune parity.

.PARAMETER LogPath
    Override log file location. Default: CMTrace-format log in
    %ProgramData%\Microsoft\IntuneManagementExtension\Logs\EntraCutover.log
    (falls back to %TEMP%).

.PARAMETER Force
    Skip interactive confirmations. Implied in Resume mode (headless).

.EXAMPLE
    .\Invoke-EntraCutover.ps1
    Read-only preflight. Emits verdict object; safe on any host.

.EXAMPLE
    .\Invoke-EntraCutover.ps1 -Mode Status
    Show migration progress (works mid-migration, after reboots).

.EXAMPLE
    $cred = Get-Credential CONTOSO\svc-unjoin
    .\Invoke-EntraCutover.ps1 -Mode Migrate -PpkgPath .\AADJ-Bulk.ppkg `
        -TenantId 11111111-2222-3333-4444-555555555555 `
        -DomainCredential $cred -AcknowledgePrereqs
    Full migration with graceful unjoin and djoin rollback insurance.

.EXAMPLE
    .\Invoke-EntraCutover.ps1 -Mode Rollback
    Revert: pre-unjoin undoes Prepare; post-unjoin rejoins the domain from
    the offline-join blob captured during Prepare.

.OUTPUTS
    PSCustomObject - verdict/result object (assess verdict or migration
    result), suitable for Export-Csv fleet reporting.

.NOTES
    PowerShell 5.1+, run elevated. Exit codes:
      0 success | 1 blocked in preflight | 2 failed before point of no return
      (reverted) | 3 failed after PONR (rollback executed) | 4 failed after
      PONR (manual intervention required; break-glass account live).

    Author : Anton Romanyuk
    Version: 0.1.0 (M2 skeleton)

    Sources: see docs/ENTRA-CUTOVER-TOOL-PLAN.md section 2 (Microsoft Learn
    dsregcmd/FAQ/bulk-enrollment; call4cloud enrollment teardown; AADMigration
    toolkit resume pattern; TechPress unjoin sequence; AADInternals BPRT).
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter()]
    [ValidateSet('Assess', 'Migrate', 'Resume', 'Status', 'Rollback')]
    [string]$Mode = 'Assess',

    [Parameter()]
    [string]$PpkgPath,

    [Parameter()]
    [ValidateSet('Ppkg', 'UserDriven')]
    [string]$JoinMode = 'Ppkg',

    [Parameter()]
    [System.Management.Automation.PSCredential]$DomainCredential,

    [Parameter()]
    [switch]$OfflineUnjoin,

    [Parameter()]
    [System.Management.Automation.PSCredential]$BreakGlassCredential,

    [Parameter()]
    [ValidatePattern('^$|^[0-9A-Fa-f]{8}-([0-9A-Fa-f]{4}-){3}[0-9A-Fa-f]{12}$')]
    [string]$TenantId,

    [Parameter()]
    [switch]$SkipKfmGate,

    [Parameter()]
    [ValidateRange(0, 90)]
    [int]$FallbackRetentionDays = 7,

    [Parameter()]
    [switch]$AcknowledgePrereqs,

    [Parameter()]
    [string]$LogPath,

    [Parameter()]
    [switch]$Force
)

Set-StrictMode -Version 2.0

# ==========================================================================
# CONSTANTS / PATHS
# ==========================================================================
$Script:ToolVersion  = '0.1.0'
$Script:CutoverRoot  = Join-Path $env:ProgramData 'EntraCutover'
$Script:BinDir       = Join-Path $Script:CutoverRoot 'Bin'
$Script:BackupDir    = Join-Path $Script:CutoverRoot 'Backup'
$Script:RollbackDir  = Join-Path $Script:CutoverRoot 'Rollback'
$Script:StateFile    = Join-Path $Script:CutoverRoot 'state.json'
$Script:ResumeTask   = 'EntraCutover-Resume'
$Script:RetireTask   = 'EntraCutover-BreakGlassRetire'
$Script:BreakGlassUser = 'ECFallback'
$Script:PhaseOrder   = @('Assess', 'Prepare', 'Teardown', 'Join', 'Finalize')
# Phases at/after which the point of no return has passed (domain unjoined).
$Script:PostPonrPhases = @('Join', 'Finalize')

# ==========================================================================
# LOGGING  (CMTrace file + colored console mirror)
# ==========================================================================
$Script:LogComponent = 'EntraCutover'
if ($LogPath) {
    $Script:LogFile = $LogPath
}
else {
    $imeLogDir = Join-Path $env:ProgramData 'Microsoft\IntuneManagementExtension\Logs'
    $Script:LogFile = Join-Path $imeLogDir 'EntraCutover.log'
}

function Write-Log {
    <#
        One CMTrace-format line to $Script:LogFile + colored console mirror.
        Levels: INFO(1) WARN(2) ERROR(3) SUCCESS(1) STEP(1).
        Self-healing: creates dir, rotates at ~5 MB, falls back to %TEMP% once.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS', 'STEP')]
        [string]$Level = 'INFO'
    )

    $color = switch ($Level) {
        'WARN'    { 'Yellow' }
        'ERROR'   { 'Red' }
        'SUCCESS' { 'Green' }
        'STEP'    { 'White' }
        default   { 'Gray' }
    }
    Write-Host ("  [{0,-7}] {1}" -f $Level, $Message) -ForegroundColor $color

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
    try { & $write $Script:LogFile }
    catch {
        $fallback = Join-Path $env:TEMP 'EntraCutover.log'
        if ($Script:LogFile -ne $fallback) {
            $Script:LogFile = $fallback
            try { & $write $Script:LogFile } catch { Write-Verbose "Write-Log failed (fallback): $($_.Exception.Message)" }
        }
    }
}

function Write-PhaseBanner {
    param([string]$Text)
    $line = '=' * 70
    Write-Host ''
    Write-Host "  $line" -ForegroundColor DarkCyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host "  $line" -ForegroundColor DarkCyan
    Write-Log $Text 'STEP'
}

# ==========================================================================
# STATE MANAGEMENT  (state.json - atomic writes, resume pointer, step ledger)
# ==========================================================================
function New-CutoverState {
    param([hashtable]$Options)
    [ordered]@{
        Version    = 1
        ToolVersion = $Script:ToolVersion
        RunId      = [guid]::NewGuid().ToString()
        StartedUtc = (Get-Date).ToUniversalTime().ToString('o')
        Computer   = $env:COMPUTERNAME
        NextPhase  = 'Assess'
        Options    = $Options
        Steps      = @{}          # "<Phase>.<Step>" -> @{Status;StartedUtc;FinishedUtc;Data}
        Device     = @{}          # captured facts: old device id, domain, tenant, ...
        Losses     = @()          # applied known-loss entries for the report
        Result     = $null
    }
}

function Get-CutoverState {
    if (-not (Test-Path $Script:StateFile)) { return $null }
    try {
        $json = Get-Content -Path $Script:StateFile -Raw -ErrorAction Stop
        return ($json | ConvertFrom-Json)
    }
    catch {
        Write-Log "State file unreadable: $($_.Exception.Message)" 'ERROR'
        return $null
    }
}

function Save-CutoverState {
    param([Parameter(Mandatory)]$State)
    if (-not (Test-Path $Script:CutoverRoot)) { New-Item -ItemType Directory -Path $Script:CutoverRoot -Force | Out-Null }
    $tmp = "$($Script:StateFile).tmp"
    $State | ConvertTo-Json -Depth 12 | Set-Content -Path $tmp -Encoding UTF8 -ErrorAction Stop
    Move-Item -Path $tmp -Destination $Script:StateFile -Force -ErrorAction Stop
}

function Test-StepCompleted {
    param([Parameter(Mandatory)]$State, [Parameter(Mandatory)][string]$Name)
    $step = $State.Steps.PSObject.Properties[$Name]
    if ($null -eq $step -and $State.Steps -is [hashtable]) { $step = if ($State.Steps.ContainsKey($Name)) { @{ Value = $State.Steps[$Name] } } else { $null } }
    if ($null -eq $step) { return $false }
    return ($step.Value.Status -eq 'Completed')
}

function Set-StepState {
    param(
        [Parameter(Mandatory)]$State,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][ValidateSet('Running', 'Completed', 'Failed', 'Skipped')][string]$Status,
        $Data
    )
    $entry = @{
        Status      = $Status
        StartedUtc  = $null
        FinishedUtc = $null
        Data        = $Data
    }
    # Preserve StartedUtc across Running -> terminal transitions.
    $existing = $null
    if ($State.Steps -is [hashtable]) {
        if ($State.Steps.ContainsKey($Name)) { $existing = $State.Steps[$Name] }
    }
    elseif ($State.Steps.PSObject.Properties[$Name]) { $existing = $State.Steps.$Name }

    if ($Status -eq 'Running') { $entry.StartedUtc = (Get-Date).ToUniversalTime().ToString('o') }
    else {
        if ($existing -and $existing.StartedUtc) { $entry.StartedUtc = $existing.StartedUtc }
        $entry.FinishedUtc = (Get-Date).ToUniversalTime().ToString('o')
    }

    if ($State.Steps -is [hashtable]) { $State.Steps[$Name] = $entry }
    else { $State.Steps | Add-Member -NotePropertyName $Name -NotePropertyValue $entry -Force }
    Save-CutoverState -State $State
}

function Invoke-Step {
    <#
        Idempotent step wrapper - THE contract for all phase implementations.
        Skips if the step already completed (resume safety), records
        Running/Completed/Failed with timestamps in state.json, logs, and
        rethrows on failure so the dispatcher can stop the phase.
        The scriptblock's return value is stored as the step's Data (must be
        JSON-serializable or $null).
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Ctx,
        [Parameter(Mandatory)][string]$Name,     # "<Phase>.<Step>", e.g. "Prepare.BreakGlass"
        [Parameter(Mandatory)][scriptblock]$Action,
        [switch]$AllowFail                        # WARN + Skipped instead of abort
    )
    if (Test-StepCompleted -State $Ctx.State -Name $Name) {
        Write-Log "step $Name already completed - skipping (resume)." 'INFO'
        return
    }
    Write-Log "step $Name ..." 'STEP'
    Set-StepState -State $Ctx.State -Name $Name -Status 'Running' -Data $null
    try {
        $data = & $Action
        Set-StepState -State $Ctx.State -Name $Name -Status 'Completed' -Data $data
        Write-Log "step $Name completed." 'SUCCESS'
    }
    catch {
        if ($AllowFail) {
            Set-StepState -State $Ctx.State -Name $Name -Status 'Skipped' -Data $_.Exception.Message
            Write-Log "step $Name failed (non-fatal): $($_.Exception.Message)" 'WARN'
        }
        else {
            Set-StepState -State $Ctx.State -Name $Name -Status 'Failed' -Data $_.Exception.Message
            Write-Log "step $Name FAILED: $($_.Exception.Message)" 'ERROR'
            throw
        }
    }
}

# ==========================================================================
# SHARED HELPERS  (available to all phase implementations)
# ==========================================================================
function Test-IsAdmin {
    $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]$identity
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Invoke-Exe {
    <#
        Run an external executable, capture stdout+stderr and exit code, log
        both. Returns @{ ExitCode; Output }. Throws only if the exe is missing.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [string[]]$Arguments = @(),
        [int]$TimeoutSec = 600
    )
    if (-not (Get-Command $Path -ErrorAction SilentlyContinue)) { throw "Executable not found: $Path" }
    Write-Log ("exec: {0} {1}" -f $Path, ($Arguments -join ' '))
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = (Get-Command $Path).Source
    $psi.Arguments = ($Arguments -join ' ')
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow  = $true
    $proc = [System.Diagnostics.Process]::Start($psi)
    # Drain both pipes concurrently via async readers BEFORE waiting. Reading
    # one stream to end before the other (or waiting before draining) deadlocks
    # once a child's combined output exceeds the OS pipe buffer (~4 KB) - e.g.
    # dsregcmd /status. The async tasks keep both buffers empty while we wait.
    $outTask = $proc.StandardOutput.ReadToEndAsync()
    $errTask = $proc.StandardError.ReadToEndAsync()
    if (-not $proc.WaitForExit($TimeoutSec * 1000)) {
        try { $proc.Kill() } catch { Write-Verbose "kill failed: $($_.Exception.Message)" }
        Write-Log ("exec TIMEOUT after {0}s: {1}" -f $TimeoutSec, $Path) 'ERROR'
        return @{ ExitCode = -1; Output = '<timeout>' }
    }
    # WaitForExit(timeout) can return before the async readers flush the final
    # buffer; the parameterless overload closes that race.
    $proc.WaitForExit()
    $stdout = ''; $stderr = ''
    try { $stdout = $outTask.Result } catch { Write-Verbose "stdout read failed: $($_.Exception.Message)" }
    try { $stderr = $errTask.Result } catch { Write-Verbose "stderr read failed: $($_.Exception.Message)" }
    $out = ("$stdout`n$stderr").Trim()
    Write-Log ("exec exit {0} (0x{0:X8})" -f $proc.ExitCode)
    if ($out) { Write-Verbose $out }
    return @{ ExitCode = $proc.ExitCode; Output = $out }
}

function Get-DsregStatus {
    <#
        Parse `dsregcmd /status` into a flat hashtable of Key -> Value
        (AzureAdJoined, DomainJoined, DomainName, TenantId, DeviceId,
        AzureAdPrt, ...). Values are raw strings ('YES'/'NO'/etc).
    #>
    $result = @{}
    $raw = Invoke-Exe -Path "$env:windir\System32\dsregcmd.exe" -Arguments '/status'
    foreach ($line in ($raw.Output -split "`r?`n")) {
        if ($line -match '^\s*([A-Za-z][A-Za-z0-9 ]+?)\s*:\s*(.+?)\s*$') {
            $key = $Matches[1] -replace '\s', ''
            if (-not $result.ContainsKey($key)) { $result[$key] = $Matches[2] }
        }
    }
    return $result
}

function New-RegistryBackup {
    <#
        reg.exe export of a key into $Script:BackupDir. Returns the file path
        or $null when the key doesn't exist. Throws on export failure.
    #>
    param([Parameter(Mandatory)][string]$KeyPath)   # e.g. 'HKLM\SOFTWARE\Microsoft\Enrollments'
    $psPath = $KeyPath -replace '^HKLM\\', 'HKLM:\' -replace '^HKCU\\', 'HKCU:\'
    if (-not (Test-Path $psPath)) { return $null }
    if (-not (Test-Path $Script:BackupDir)) { New-Item -ItemType Directory -Path $Script:BackupDir -Force | Out-Null }
    $safe = ($KeyPath -replace '[\\:]', '_')
    $dest = Join-Path $Script:BackupDir ("{0}_{1:yyyyMMdd-HHmmss}.reg" -f $safe, (Get-Date))
    $regOut = & reg.exe export $KeyPath $dest /y 2>&1
    if ($LASTEXITCODE -ne 0) { throw "reg export failed for ${KeyPath}: $regOut" }
    Write-Log "registry backup: $KeyPath -> $dest"
    return $dest
}

function Register-ResumeTask {
    <#
        SYSTEM 'At Startup' task that re-invokes this tool with -Mode Resume
        from the persisted Bin copy. Idempotent.
    #>
    $scriptPath = Join-Path $Script:BinDir 'Invoke-EntraCutover.ps1'
    $action  = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument `
        ("-NoProfile -ExecutionPolicy Bypass -File `"{0}`" -Mode Resume -Force" -f $scriptPath)
    $trigger = New-ScheduledTaskTrigger -AtStartup
    $trigger.Delay = 'PT30S'
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
        -ExecutionTimeLimit (New-TimeSpan -Hours 2) -MultipleInstances IgnoreNew
    Register-ScheduledTask -TaskName $Script:ResumeTask -Action $action -Trigger $trigger `
        -Principal $principal -Settings $settings -Force | Out-Null
    Write-Log "resume task '$Script:ResumeTask' registered (SYSTEM, at startup)."
}

function Unregister-ResumeTask {
    Unregister-ScheduledTask -TaskName $Script:ResumeTask -Confirm:$false -ErrorAction SilentlyContinue
    Write-Log "resume task '$Script:ResumeTask' removed."
}

function Set-MigrationNotice {
    <#
        Lock-screen legal notice used as headless progress display between
        reboots. Empty strings clear the notice. Original values are saved to
        state the first time (Prepare) and restored by Finalize/Rollback.
    #>
    param([string]$Caption, [string]$Text)
    $key = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
    Set-ItemProperty -Path $key -Name 'legalnoticecaption' -Value $Caption -Force
    Set-ItemProperty -Path $key -Name 'legalnoticetext' -Value $Text -Force
    if ($Caption) { Write-Log "lock-screen notice: $Caption" } else { Write-Log 'lock-screen notice cleared.' }
}

function Request-Reboot {
    <#
        Phases call this instead of rebooting directly. The dispatcher
        performs the actual reboot after saving state.
    #>
    param([Parameter(Mandatory)][hashtable]$Ctx, [string]$Reason = '')
    $Ctx.RebootRequired = $true
    Write-Log "reboot requested: $Reason" 'WARN'
}

function Protect-Secret {
    # DPAPI machine-scope protect -> base64 (break-glass password at-rest form).
    param([Parameter(Mandatory)][string]$Plain)
    Add-Type -AssemblyName System.Security -ErrorAction SilentlyContinue
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Plain)
    $prot  = [System.Security.Cryptography.ProtectedData]::Protect($bytes, $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
    return [Convert]::ToBase64String($prot)
}

function Unprotect-Secret {
    param([Parameter(Mandatory)][string]$Cipher)
    Add-Type -AssemblyName System.Security -ErrorAction SilentlyContinue
    $prot  = [Convert]::FromBase64String($Cipher)
    $bytes = [System.Security.Cryptography.ProtectedData]::Unprotect($prot, $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
    return [System.Text.Encoding]::UTF8.GetString($bytes)
}

# ==========================================================================
# PHASE IMPLEMENTATIONS  (dot-sourced; one Invoke-Phase<Name> per file)
# Contract for lib\Phase*.ps1:
#   - Define exactly one entry function: Invoke-PhaseAssess / Prepare /
#     Teardown / Join / Finalize / plus Invoke-CutoverRollback in Rollback.ps1.
#   - Signature: param([hashtable]$Ctx). $Ctx keys:
#       State          state object (persist via Invoke-Step / Set-StepState)
#       Options        resolved options hashtable (Mode/JoinMode/PpkgPath/...)
#       Paths          @{ Root; Bin; Backup; Rollback; StateFile }
#       RebootRequired set via Request-Reboot to end phase with a reboot
#       StayInPhase    set $true to exit WITHOUT advancing the phase pointer
#                      (phase re-runs at next resume; e.g. waiting on a
#                      user-driven join). Resume task must remain registered.
#       DomainCredential  PSCredential or $null (never persisted to state)
#   - Wrap EVERY unit of work in Invoke-Step -Name "<Phase>.<Step>".
#   - Assess: read-only; returns verdict PSCustomObject. Other phases: return
#     value ignored; throw to abort; PS 5.1-compatible code only.
#   - Private helpers must be named <Verb>-EC<Noun> to avoid collisions.
# ==========================================================================
$libDir = Join-Path $PSScriptRoot 'lib'
if (Test-Path $libDir) {
    Get-ChildItem -Path $libDir -Filter '*.ps1' | Sort-Object Name | ForEach-Object {
        . $_.FullName
    }
}

# ==========================================================================
# DISPATCH
# ==========================================================================
if (-not (Test-IsAdmin)) {
    Write-Log 'This tool must run elevated.' 'ERROR'
    exit 1
}

Write-PhaseBanner ("EntraCutover v{0} - Mode={1} - {2}" -f $Script:ToolVersion, $Mode, $env:COMPUTERNAME)
Write-Host '  [EXPERIMENTAL] Community-built; NOT supported by Microsoft. The supported' -ForegroundColor Yellow
Write-Host '                 hybrid-to-Entra path is wipe + Autopilot. Pilot on a test ring.' -ForegroundColor Yellow
Write-Log 'EXPERIMENTAL tool - not supported by Microsoft (supported path is wipe + Autopilot).' 'WARN'
Write-Log ("log file: {0}" -f $Script:LogFile)

# ---- Status: print state summary and exit ---------------------------------
if ($Mode -eq 'Status') {
    $state = Get-CutoverState
    if (-not $state) {
        Write-Host '  No migration state found on this host.' -ForegroundColor DarkGray
        exit 0
    }
    Write-Host ("  Run       : {0} (started {1})" -f $state.RunId, $state.StartedUtc)
    Write-Host ("  Next phase: {0}" -f $state.NextPhase) -ForegroundColor Cyan
    $steps = @($state.Steps.PSObject.Properties | Sort-Object { $_.Value.StartedUtc })
    foreach ($s in $steps) {
        $c = switch ($s.Value.Status) {
            'Completed' { 'Green' } 'Failed' { 'Red' } 'Running' { 'Yellow' } default { 'DarkGray' }
        }
        Write-Host ("    {0,-40} {1}" -f $s.Name, $s.Value.Status) -ForegroundColor $c
    }
    if ($state.Result) { Write-Host ("  Result: {0}" -f ($state.Result | ConvertTo-Json -Depth 4 -Compress)) }
    exit 0
}

# ---- Build context --------------------------------------------------------
$options = [ordered]@{
    Mode                  = $Mode
    JoinMode              = $JoinMode
    PpkgPath              = if ($PpkgPath) { (Resolve-Path -Path $PpkgPath -ErrorAction SilentlyContinue).Path } else { $null }
    TenantId              = $TenantId
    OfflineUnjoin         = [bool]$OfflineUnjoin
    SkipKfmGate           = [bool]$SkipKfmGate
    FallbackRetentionDays = $FallbackRetentionDays
    BreakGlassUser        = $Script:BreakGlassUser
    HasDomainCredential   = [bool]$DomainCredential
    HasBreakGlassCredential = [bool]$BreakGlassCredential
}

# The break-glass account name is fixed; only the supplied password is used.
# Warn (don't fail) if the operator passed a credential with a different name.
if ($BreakGlassCredential) {
    $suppliedName = ($BreakGlassCredential.UserName -split '\\')[-1]
    if ($suppliedName -and ($suppliedName -ne $Script:BreakGlassUser)) {
        Write-Log ("Supplied break-glass username '{0}' is ignored; the account is always '{1}'. Only the password is used." -f $suppliedName, $Script:BreakGlassUser) 'WARN'
    }
}

$state = Get-CutoverState

switch ($Mode) {
    'Assess' {
        # Standalone read-only run; never persists phase state.
        $ctx = @{
            State = New-CutoverState -Options $options
            Options = $options; DomainCredential = $DomainCredential
            Paths = @{ Root = $Script:CutoverRoot; Bin = $Script:BinDir; Backup = $Script:BackupDir; Rollback = $Script:RollbackDir; StateFile = $null }
            RebootRequired = $false
            ReadOnly = $true
        }
        $verdict = Invoke-PhaseAssess -Ctx $ctx
        $verdict
        if ($verdict.Ready) { exit 0 } else { exit 1 }
    }

    'Rollback' {
        if (-not $state) { Write-Log 'Nothing to roll back (no state).' 'WARN'; exit 0 }
        $ctx = @{
            State = $state; Options = $options; DomainCredential = $DomainCredential
            Paths = @{ Root = $Script:CutoverRoot; Bin = $Script:BinDir; Backup = $Script:BackupDir; Rollback = $Script:RollbackDir; StateFile = $Script:StateFile }
            RebootRequired = $false
        }
        if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, 'Roll back Entra cutover migration')) {
            Invoke-CutoverRollback -Ctx $ctx
            if ($ctx.RebootRequired) {
                Write-Log 'Rebooting to complete rollback.' 'WARN'
                Restart-Computer -Force
            }
        }
        exit 0
    }

    { $_ -in 'Migrate', 'Resume' } {
        if ($Mode -eq 'Migrate') {
            if ($state -and $state.NextPhase -and $state.NextPhase -ne 'Assess') {
                Write-Log "Migration already in progress (next phase: $($state.NextPhase)). Use -Mode Resume, Status, or Rollback." 'ERROR'
                exit 1
            }
            if (-not $AcknowledgePrereqs) {
                Write-Log 'Migrate requires -AcknowledgePrereqs (tenant-side runbook complete: sync scoping, CA exclusion, licensing, KFM policy, Intune parity). See docs/ENTRA-CUTOVER-TOOL-PLAN.md section 4.' 'ERROR'
                exit 1
            }
            if ($JoinMode -eq 'Ppkg' -and -not $options.PpkgPath) {
                Write-Log '-PpkgPath is required for -JoinMode Ppkg.' 'ERROR'
                exit 1
            }
            # Fail-closed: an unattended run must not mint a random break-glass
            # password that is only recoverable by an admin already on the box.
            if ($Force -and -not $BreakGlassCredential) {
                Write-Log 'Unattended Migrate (-Force) requires -BreakGlassCredential so the break-glass admin has a known, operator-held password. Without it the generated password is unrecoverable from a locked device. Re-run with -BreakGlassCredential, or run attended (without -Force) to use generate-and-show-once.' 'ERROR'
                exit 1
            }
            $stateObj = New-CutoverState -Options $options
            Save-CutoverState -State $stateObj
            $state = Get-CutoverState   # normalize to PSCustomObject form
        }
        else {
            if (-not $state) { Write-Log 'Resume: no state file - nothing to resume.' 'ERROR'; exit 1 }
            Write-Log ("Resuming run {0} at phase {1}." -f $state.RunId, $state.NextPhase)
        }

        $ctx = @{
            State = $state; Options = $options; DomainCredential = $DomainCredential
            BreakGlassCredential = $BreakGlassCredential   # never persisted to state
            Paths = @{ Root = $Script:CutoverRoot; Bin = $Script:BinDir; Backup = $Script:BackupDir; Rollback = $Script:RollbackDir; StateFile = $Script:StateFile }
            RebootRequired = $false
            StayInPhase = $false
        }

        $exitCode = 0
        while ($ctx.State.NextPhase) {
            $phase = $ctx.State.NextPhase
            $idx = [array]::IndexOf($Script:PhaseOrder, $phase)
            Write-PhaseBanner ("[Phase {0}/{1}] {2}" -f ($idx + 1), $Script:PhaseOrder.Count, $phase)

            # Destructive gate at the Teardown boundary (operator-attended only).
            if ($phase -eq 'Teardown' -and $Mode -eq 'Migrate' -and -not $Force) {
                if (-not $PSCmdlet.ShouldProcess($env:COMPUTERNAME, 'Tear down Intune enrollment, leave Entra hybrid state, unjoin domain (POINT OF NO RETURN)')) { break }
                $ans = Read-Host 'CONFIRM: unjoin this device from the domain? This is the point of no return. [y/N]'
                if ($ans -notin @('y', 'Y')) { Write-Log 'Aborted by operator before Teardown.' 'WARN'; break }
            }

            try {
                switch ($phase) {
                    'Assess'   { $v = Invoke-PhaseAssess -Ctx $ctx
                                 if (-not $v.Ready -and -not $Force) { Write-Log 'Preflight blockers present - aborting (use -Force to override warnings-only).' 'ERROR'; $exitCode = 1; break } }
                    'Prepare'  { Invoke-PhasePrepare  -Ctx $ctx }
                    'Teardown' { Invoke-PhaseTeardown -Ctx $ctx }
                    'Join'     { Invoke-PhaseJoin     -Ctx $ctx }
                    'Finalize' { Invoke-PhaseFinalize -Ctx $ctx }
                }
            }
            catch {
                Write-Log ("Phase {0} failed: {1}" -f $phase, $_.Exception.Message) 'ERROR'
                Write-Log $_.ScriptStackTrace 'ERROR'
                $postPonr = $Script:PostPonrPhases -contains $phase
                $exitCode = if ($postPonr) { 4 } else { 2 }
                break
            }
            if ($exitCode -ne 0) { break }

            if ($ctx.StayInPhase) {
                Write-Log ("Phase {0} is waiting on an external action - it will re-run at next resume/boot." -f $phase) 'WARN'
                Save-CutoverState -State $ctx.State
                if ($ctx.RebootRequired) { Restart-Computer -Force }
                exit 0
            }

            # Advance the resume pointer.
            $next = if ($idx + 1 -lt $Script:PhaseOrder.Count) { $Script:PhaseOrder[$idx + 1] } else { $null }
            $ctx.State.NextPhase = $next
            Save-CutoverState -State $ctx.State

            if ($ctx.RebootRequired) {
                Write-Log ("Rebooting to continue at phase: {0}" -f $next) 'WARN'
                Restart-Computer -Force
                exit 0
            }
        }

        if ($exitCode -eq 0 -and -not $ctx.State.NextPhase) {
            Write-PhaseBanner 'Migration complete.'
            if ($ctx.State.Result) { $ctx.State.Result }
        }
        exit $exitCode
    }
}
