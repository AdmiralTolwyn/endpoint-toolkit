<#
.SYNOPSIS
    Detects Start Menu / Desktop shortcuts targeting OneDrive.App.exe ("OneDrive Photos").

.DESCRIPTION
    Scans the all-users Start Menu, the Public Desktop, and every real user
    profile's Start Menu and Desktop for .lnk files whose resolved TargetPath is
    OneDrive.App.exe.

    Matching is on the resolved target, never the shortcut's display name -- the
    name is localized and may change between client versions.

    Exit 1 = shortcut found, remediate.
    Exit 0 = none found, OR an unexpected error (fail-safe: never trigger blind
             remediation on an unknown state).

    Intune: run as SYSTEM, 64-bit PowerShell, daily.

.PARAMETER LogPath
    Log file or folder. A value ending in .log is used as-is; anything else is
    treated as a folder and PR_OneDrivePhotosShortcut.log is appended. Defaults
    to %ProgramData%\Microsoft\IntuneManagementExtension\Logs. Logging is
    best-effort and never affects the exit code.

.EXAMPLE
    .\Detect-OneDrivePhotosShortcut.ps1

.EXAMPLE
    .\Detect-OneDrivePhotosShortcut.ps1 -LogPath 'C:\Windows\CCM\Logs'

.NOTES
    Author:  Anton Romanyuk
    Version: 2.0
    Date:    2026-08-04

    TEMPORARY MITIGATION. No supported policy control suppresses this app
    surface today; removing the shortcut only hides the entry (it stays
    launchable from Search). Retire this once a supported control exists.
    Never touches OneDrive.exe, OneDrive.App.exe, or C:\Program Files\Microsoft OneDrive\.

    THIS SCRIPT IS PROVIDED "AS-IS" WITHOUT WARRANTY OF ANY KIND.
#>

[CmdletBinding()]
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
        if (-not (Test-Path $dir)) { New-Item $dir -ItemType Directory -Force | Out-Null }
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  [Detect]  $Message" | Out-File $logFile -Append -Encoding utf8
    }
    # Best-effort: a bad log path must never fail the run, but -Verbose should
    # still surface why nothing is being written.
    catch { Write-Verbose "Log write failed: $($_.Exception.Message)" }
}

function Get-PhotosShortcut {
    $roots = @(
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs"
        "$env:PUBLIC\Desktop"
    )
    foreach ($key in Get-ChildItem $profileListKey) {
        # Real user profiles only: local/domain (S-1-5-21-*) and Entra ID
        # (S-1-12-1-*). System and service profiles are ACL-protected and would
        # throw access-denied on every run.
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

try {
    if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
        throw 'Running 32-bit PowerShell on a 64-bit OS -- set "Run script in 64-bit PowerShell" to Yes.'
    }

    $found = @(Get-PhotosShortcut)

    if ($found.Count -eq 0) {
        Write-Log 'Compliant: no matching shortcuts.'
        Write-Output "COMPLIANT: No shortcuts targeting $targetLeaf found."
        exit 0
    }

    Write-Log "Non-compliant: $($found.Count) shortcut(s)."
    Write-Output "NON-COMPLIANT: $($found.Count) shortcut(s) targeting ${targetLeaf}: $($found -join '; ')"
    exit 1
}
catch {
    # Fail-safe: an unknown state must not trigger blind remediation.
    Write-Log "ERROR: $($_.Exception.Message)"
    [Console]::Error.WriteLine("ERROR: $($_.Exception.Message)")
    exit 0
}
