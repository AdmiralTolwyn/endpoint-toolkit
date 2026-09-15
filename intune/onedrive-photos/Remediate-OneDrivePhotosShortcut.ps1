<#
.SYNOPSIS
    Removes Start Menu / Desktop shortcuts targeting OneDrive.App.exe ("OneDrive Photos").

.DESCRIPTION
    Deletes only .lnk files whose resolved TargetPath is OneDrive.App.exe.
    Never touches OneDrive.exe, OneDrive.App.exe itself, or anything under
    C:\Program Files\Microsoft OneDrive\ -- this is shortcut-only.

    Idempotent, so it is safe to run on a schedule via Intune Remediations,
    a ConfigMgr baseline or package, or a computer startup script. The shortcut
    returns after OneDrive client updates and for new user profiles, so a daily
    schedule is expected.

    Exit 0 = success, including the no-op and -WhatIf cases.
    Exit 1 = a deletion was attempted and failed.

    Intune: run as SYSTEM, 64-bit PowerShell, daily.

.PARAMETER WhatIf
    Preview only -- reports what would be removed and deletes nothing.

.PARAMETER LogPath
    Log file or folder. A value ending in .log is used as-is; anything else is
    treated as a folder and PR_OneDrivePhotosShortcut.log is appended. Defaults
    to %ProgramData%\Microsoft\IntuneManagementExtension\Logs. Logging is
    best-effort and never affects the exit code.

.EXAMPLE
    .\Remediate-OneDrivePhotosShortcut.ps1 -WhatIf

.EXAMPLE
    .\Remediate-OneDrivePhotosShortcut.ps1 -LogPath 'C:\Windows\CCM\Logs'

.NOTES
    Author:  Anton Romanyuk
    Version: 2.0
    Date:    2026-08-04

    TEMPORARY MITIGATION. No supported policy control suppresses this app
    surface today; removing the shortcut only hides the entry (it stays
    launchable from Search). Retire this once a supported control exists.

    THIS SCRIPT IS PROVIDED "AS-IS" WITHOUT WARRANTY OF ANY KIND.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateNotNullOrEmpty()]
    [string]$LogPath
)

$ErrorActionPreference = 'Stop'

$targetLeaf     = 'OneDrive.App.exe'
$profileListKey = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'

if ($LogPath -and [IO.Path]::GetExtension($LogPath) -eq '.log') { $logFile = $LogPath }
elseif ($LogPath) { $logFile = Join-Path $LogPath 'PR_OneDrivePhotosShortcut.log' }
else { $logFile = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\PR_OneDrivePhotosShortcut.log" }

function Write-Log {
    param([string]$Message)
    Write-Verbose $Message
    try {
        $dir = Split-Path $logFile -Parent
        # -WhatIf:$false -- logging is diagnostics, not a previewed state change.
        if (-not (Test-Path $dir)) { New-Item $dir -ItemType Directory -Force -WhatIf:$false | Out-Null }
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  [Remediate]  $Message" |
            Out-File $logFile -Append -Encoding utf8 -WhatIf:$false
    }
    # Best-effort: a bad log path must never fail the run, but -Verbose should
    # still surface why nothing is being written.
    catch { Write-Verbose "Log write failed: $($_.Exception.Message)" }
}

function Get-PhotosShortcut {
    # Enumerate real user profiles: local/domain (S-1-5-21-*) and Entra ID
    # (S-1-12-1-*) only. System and service profiles are ACL-protected and would
    # throw access-denied.
    $roots = @(
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs"
        "$env:PUBLIC\Desktop"
    )
    foreach ($key in Get-ChildItem $profileListKey) {
        if ($key.PSChildName -notmatch '^S-1-(5-21|12-1)-') { continue }
        $profilePath = (Get-ItemProperty $key.PSPath -ErrorAction SilentlyContinue).ProfileImagePath
        if ($profilePath) {
            $roots += "$profilePath\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"
            $roots += "$profilePath\Desktop"
        }
    }

    $shell = New-Object -ComObject WScript.Shell
    try {
        foreach ($root in $roots) {
            # Test-Path can throw UnauthorizedAccessException regardless of -ErrorAction.
            try { if (-not (Test-Path -LiteralPath $root)) { continue } }
            catch { Write-Log "Skipped inaccessible root: $root"; continue }

            # Start Menu folders are nested; Desktop is scanned flat.
            $recurse = $root -like '*\Start Menu\Programs'
            foreach ($lnk in Get-ChildItem -LiteralPath $root -Filter *.lnk -File -Recurse:$recurse -ErrorAction SilentlyContinue) {
                $target = $null
                try { $target = $shell.CreateShortcut($lnk.FullName).TargetPath }
                catch { Write-Log "Could not resolve: $($lnk.FullName)" }

                if ($target -and [IO.Path]::GetFileName($target) -eq $targetLeaf) {
                    Write-Log "Match: $($lnk.FullName) -> $target"
                    $lnk.FullName
                }
            }
        }
    }
    finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell) }
}

if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    Write-Output 'ERROR: Running 32-bit PowerShell on a 64-bit OS -- set "Run script in 64-bit PowerShell" to Yes.'
    exit 1
}

try { $found = @(Get-PhotosShortcut) }
catch {
    Write-Log "Discovery failed: $($_.Exception.Message)"
    Write-Output "ERROR: Discovery failed: $($_.Exception.Message)"
    exit 1
}

if ($found.Count -eq 0) {
    Write-Log 'Nothing to remove.'
    Write-Output "SUCCESS: No shortcuts targeting $targetLeaf found."
    exit 0
}

$removed = 0
$failed  = @()

foreach ($path in $found) {
    # ShouldProcess prints its own "What if:" line per item under -WhatIf and
    # returns $false, so nothing below runs in preview mode.
    if (-not $PSCmdlet.ShouldProcess($path, 'Remove shortcut')) { continue }
    try {
        Remove-Item -LiteralPath $path -Force
        Write-Log "Removed: $path"
        Write-Output "Removed: $path"
        $removed++
    }
    catch {
        # One locked file must not abort the run.
        Write-Log "Failed: $path -- $($_.Exception.Message)"
        $failed += $path
    }
}

if ($failed.Count -gt 0) {
    Write-Output "FAILED: $($failed.Count) of $($found.Count) could not be removed: $($failed -join '; ')"
    exit 1
}
if ($WhatIfPreference) {
    Write-Output "WHATIF: $($found.Count) shortcut(s) would be removed. No changes made."
    exit 0
}

Write-Log "Removed $removed shortcut(s)."
Write-Output "SUCCESS: Removed $removed shortcut(s) targeting $targetLeaf."
exit 0
