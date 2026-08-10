<#
.SYNOPSIS
    Repairs Secure Boot *reporting* prerequisites only. Does NOT touch certificate state.

.DESCRIPTION
    Scope is deliberately narrow: the three documented conditions that stop a device from
    reporting Secure Boot status, leaving it in the "Unknown" bucket.

      1. \Microsoft\Windows\PI\Secure-Boot-Update scheduled task disabled -> re-enable.
      2. DisableOneSettingsDownloads policy enabled -> set to 0 (blocks OneSettings config retrieval).
      3. AllowTelemetry below Required(1) -> raise to 1. GATED OFF BY DEFAULT (see $SetTelemetry).

    Explicitly NOT done here, and must not be added:
      - AvailableUpdates is never written (that would force a certificate update).
      - MicrosoftUpdateManagedOptIn / HighConfidenceOptOut are never written.
      - No DBX, no boot manager, no boot order, no BitLocker change.
      - The scheduled task is not started; it runs at startup and every 12 hours on its own.

    Item 2 and 3 write to HKLM\SOFTWARE\Policies\..., which is normally owned by Group Policy
    or a Policy CSP. If those settings arrive from GPO/Intune, fix them at source - a local
    write here will be reverted at the next policy refresh and the device will regress.

.NOTES
    PowerShell 5.1. Runs as SYSTEM. Exit 0 = success, exit 1 = at least one action failed.
#>

[CmdletBinding()]
param()

# Set to $true only after the privacy/diagnostic-data change has been approved.
# Raising AllowTelemetry is a data-sharing decision, not a technical one.
$SetTelemetry = $false

$ErrorActionPreference = 'Stop'
$actions = New-Object System.Collections.ArrayList
$failed  = $false

function Add-Action {
    param([string]$Text)
    [void]$actions.Add($Text)
}

# --- 1. Secure-Boot-Update scheduled task -----------------------------------
try {
    $task = Get-ScheduledTask -TaskPath '\Microsoft\Windows\PI\' -TaskName 'Secure-Boot-Update' -ErrorAction SilentlyContinue
    if ($null -eq $task) {
        Add-Action 'task=NotFound (needs recreation - out of scope for this script)'
        $failed = $true
    }
    elseif ($task.State -eq 'Disabled') {
        Enable-ScheduledTask -TaskPath '\Microsoft\Windows\PI\' -TaskName 'Secure-Boot-Update' | Out-Null
        Add-Action 'task=Enabled'
    }
    else {
        Add-Action ('task=' + $task.State + ' (no change)')
    }
}
catch {
    Add-Action ('task=ERROR ' + $_.Exception.Message)
    $failed = $true
}

# --- 2. DisableOneSettingsDownloads -----------------------------------------
$dcPol = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection'
try {
    $oneS = $null
    if (Test-Path $dcPol) {
        $p = Get-ItemProperty -Path $dcPol -Name 'DisableOneSettingsDownloads' -ErrorAction SilentlyContinue
        if ($p) { $oneS = $p.DisableOneSettingsDownloads }
    }
    if ($oneS -eq 1) {
        Set-ItemProperty -Path $dcPol -Name 'DisableOneSettingsDownloads' -Value 0 -Type DWord
        Add-Action 'oneSettings=0 (was 1)'
    }
    else {
        Add-Action ('oneSettings=' + $(if ($null -eq $oneS) { 'NotSet' } else { $oneS }) + ' (no change)')
    }
}
catch {
    Add-Action ('oneSettings=ERROR ' + $_.Exception.Message)
    $failed = $true
}

# --- 3. AllowTelemetry ------------------------------------------------------
try {
    $tel = $null
    if (Test-Path $dcPol) {
        $p = Get-ItemProperty -Path $dcPol -Name 'AllowTelemetry' -ErrorAction SilentlyContinue
        if ($p) { $tel = $p.AllowTelemetry }
    }

    if ($null -ne $tel -and $tel -ge 1) {
        Add-Action ('telemetry=' + $tel + ' (no change)')
    }
    elseif (-not $SetTelemetry) {
        Add-Action ('telemetry=' + $(if ($null -eq $tel) { 'NotSet' } else { $tel }) + ' NEEDS-APPROVAL (not changed)')
    }
    else {
        if (-not (Test-Path $dcPol)) { New-Item -Path $dcPol -Force | Out-Null }
        Set-ItemProperty -Path $dcPol -Name 'AllowTelemetry' -Value 1 -Type DWord
        Add-Action 'telemetry=1 (raised to Required)'
    }
}
catch {
    Add-Action ('telemetry=ERROR ' + $_.Exception.Message)
    $failed = $true
}

Write-Output ($actions -join '; ')
if ($failed) { exit 1 }
exit 0
