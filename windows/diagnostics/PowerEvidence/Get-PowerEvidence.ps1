#Requires -Version 5.1
<#
.SYNOPSIS
    Collects a portable evidence bundle for power, standby, and screen-on behaviour,
    then packages it as a single .zip for offline analysis.

.DESCRIPTION
    Read-only with respect to system configuration; the only writes are the artifact
    files it produces. Runs without administrator rights. When elevated it additionally
    captures the SleepStudy and System Power reports, which require elevation.

    Collected in one pass:
      1. Available sleep states                 powercfg /a
      2. Battery report (HTML + XML)            powercfg /batteryreport
      3. SleepStudy + System Power report       powercfg /sleepstudy, /systempowerreport   (elevated only)
      4. Wake diagnostics                       powercfg /lastwake, /waketimers, /requests  (/requests elevated)
      5. Power events from the System log       Kernel-General / Kernel-Power / Kernel-Boot
      6. Fast Startup / Hibernate configuration  HiberbootEnabled + sleep-state availability

    Every artifact is written into a timestamped folder, a plain-text summary.txt is added,
    and the whole folder is compressed to <folder>.zip so the user can return one file.

.PARAMETER OutputRoot
    Parent folder for the evidence bundle. Default: the current user's Desktop.

.PARAMETER Days
    How many days of history to request for the battery/sleep reports and the System log
    query. Default 7. Range 1-60.

.PARAMETER NoZip
    Keep the raw folder only; skip creating the .zip archive.

.PARAMETER NoLaunch
    Do not open the output folder in Explorer when finished.

.EXAMPLE
    .\Get-PowerEvidence.ps1
    Collect 7 days of evidence to the Desktop and produce a .zip. Run elevated for the
    SleepStudy and System Power reports.

.EXAMPLE
    .\Get-PowerEvidence.ps1 -Days 14 -OutputRoot C:\Temp
    Collect 14 days of evidence under C:\Temp.

.EXAMPLE
    .\Get-PowerEvidence.ps1 -NoZip -NoLaunch
    Collect the raw folder only, without archiving or opening Explorer (useful for automation).

.NOTES
    File:     windows/diagnostics/PowerEvidence/Get-PowerEvidence.ps1
    Author:   Anton Romanyuk
    Version:  1.0.0
    Requires: PowerShell 5.1+. Elevation is required only for the SleepStudy, System Power,
              and /requests sections; every other section works unelevated.

    Exit codes:
      0 - OK        all sections collected (run was elevated)
      3 - PARTIAL   ran unelevated; SleepStudy / System Power / /requests were skipped
      4 - ERROR     unexpected failure caught at top level

    References:
      https://learn.microsoft.com/windows-hardware/design/device-experiences/modern-standby
      https://learn.microsoft.com/windows/win32/power/system-power-states

.DISCLAIMER
    THIS SCRIPT IS PROVIDED "AS-IS" WITHOUT WARRANTY OF ANY KIND. Test against a
    representative device before deploying through an endpoint management platform.
#>

[CmdletBinding()]
param(
    [string]$OutputRoot = "$env:USERPROFILE\Desktop",

    [ValidateRange(1, 60)]
    [int]$Days = 7,

    [switch]$NoZip,      # keep the raw folder only; skip creating the .zip
    [switch]$NoLaunch    # do not open the output folder in Explorer when finished
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Continue'

$EXIT_OK      = 0
$EXIT_PARTIAL = 3
$EXIT_ERROR   = 4

# ── Console helpers (house style: Cyan headers, coloured status) ────────────
function Write-Section {
    param([string]$Title)
    Write-Host ''
    Write-Host ('=== ' + $Title + ' ===') -ForegroundColor Cyan
}

function Write-StepResult {
    param([string]$Status, [string]$Detail = '')
    $colour = switch ($Status) {
        'OK'      { 'Green' }
        'SKIPPED' { 'Yellow' }
        'ERROR'   { 'Red' }
        default   { 'Gray' }
    }
    Write-Host ('  [{0}] ' -f $Status) -ForegroundColor $colour -NoNewline
    Write-Host $Detail -ForegroundColor DarkGray
}

try {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
               ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    $stamp  = Get-Date -Format 'yyyyMMdd-HHmmss'
    $outDir = Join-Path $OutputRoot "PowerEvidence_$($env:COMPUTERNAME)_$stamp"
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null

    # Plain-text record that travels inside the bundle.
    $summary = New-Object System.Collections.Generic.List[string]
    function Add-Summary { param([string]$Text) $summary.Add($Text) }

    Add-Summary '=== Power Evidence Collection ==='
    Add-Summary ("Computer  : {0}" -f $env:COMPUTERNAME)
    Add-Summary ("User      : {0}" -f $env:USERNAME)
    Add-Summary ("Collected : {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz'))
    Add-Summary ("Elevated  : {0}" -f $isAdmin)
    Add-Summary ("Window    : last {0} day(s)" -f $Days)
    Add-Summary ("Output    : {0}" -f $outDir)
    Add-Summary ''

    Write-Host ''
    Write-Host '=== Power Evidence Collection ===' -ForegroundColor Cyan
    Write-Host ("Computer: {0}   User: {1}   Elevated: {2}   Window: {3}d" -f `
        $env:COMPUTERNAME, $env:USERNAME, $isAdmin, $Days) -ForegroundColor DarkGray
    if (-not $isAdmin) {
        Write-Host '  Not elevated - SleepStudy, System Power, and /requests will be skipped.' -ForegroundColor Yellow
    }

    # ── 1. Available sleep states ────────────────────────────────────────────
    Write-Section '[1/6] Available sleep states (powercfg /a)'
    $powercfgA = powercfg /a 2>&1
    $powercfgAText = ($powercfgA | Out-String)
    $powercfgA | Out-File (Join-Path $outDir 'powercfg-a.txt') -Encoding utf8
    Add-Summary '--- [1/6] Available sleep states -> powercfg-a.txt ---'
    Add-Summary $powercfgAText.TrimEnd()
    Add-Summary ''
    Write-StepResult 'OK' 'powercfg-a.txt'

    # ── 2. Battery report (no elevation required) ────────────────────────────
    Write-Section '[2/6] Battery report (powercfg /batteryreport)'
    $batteryHtml = Join-Path $outDir 'batteryreport.html'
    $batteryXml  = Join-Path $outDir 'batteryreport.xml'
    powercfg /batteryreport /duration $Days /output $batteryHtml 2>&1 | Out-Null
    powercfg /batteryreport /duration $Days /output $batteryXml /xml 2>&1 | Out-Null
    $batteryOk = (Test-Path $batteryHtml)
    Add-Summary '--- [2/6] Battery report -> batteryreport.html + .xml ---'
    if ($batteryOk) {
        Add-Summary 'Written: batteryreport.html, batteryreport.xml'
        Write-StepResult 'OK' 'batteryreport.html + .xml'
    } else {
        Add-Summary 'No battery report produced (likely a desktop with no battery).'
        Write-StepResult 'SKIPPED' 'no battery present'
    }
    Add-Summary ''

    # ── 3. SleepStudy + System Power report (elevated only) ──────────────────
    Write-Section '[3/6] SleepStudy + System Power report (requires elevation)'
    Add-Summary '--- [3/6] SleepStudy + System Power report (requires elevation) ---'
    if ($isAdmin) {
        powercfg /sleepstudy /duration $Days /output (Join-Path $outDir 'sleepstudy.html') 2>&1 | Out-Null
        powercfg /systempowerreport /output (Join-Path $outDir 'systempowerreport.html') 2>&1 | Out-Null
        Add-Summary 'Written: sleepstudy.html, systempowerreport.html'
        Write-StepResult 'OK' 'sleepstudy.html + systempowerreport.html'
    } else {
        Add-Summary 'SKIPPED - not elevated. Re-run in an elevated console for the full data set.'
        Write-StepResult 'SKIPPED' 're-run elevated for SleepStudy / System Power'
    }
    Add-Summary ''

    # ── 4. Wake diagnostics ──────────────────────────────────────────────────
    Write-Section '[4/6] Wake diagnostics (lastwake / waketimers / requests)'
    $wakeFile = Join-Path $outDir 'powercfg-wake.txt'
    $wake = New-Object System.Collections.Generic.List[string]
    $wake.Add('=== powercfg /lastwake ===')
    $wake.Add((powercfg /lastwake 2>&1 | Out-String).TrimEnd())
    $wake.Add('')
    $wake.Add('=== powercfg /waketimers ===')
    $wake.Add((powercfg /waketimers 2>&1 | Out-String).TrimEnd())
    $wake.Add('')
    $wake.Add('=== powercfg /requests (requires elevation) ===')
    if ($isAdmin) {
        $wake.Add((powercfg /requests 2>&1 | Out-String).TrimEnd())
    } else {
        $wake.Add('SKIPPED - requires elevation.')
    }
    Set-Content -Path $wakeFile -Value ($wake -join [Environment]::NewLine) -Encoding UTF8
    Add-Summary '--- [4/6] Wake diagnostics -> powercfg-wake.txt ---'
    Add-Summary 'Written: powercfg-wake.txt'
    Add-Summary ''
    Write-StepResult 'OK' ('powercfg-wake.txt' + $(if (-not $isAdmin) { ' (/requests skipped - not elevated)' } else { '' }))

    # ── 5. Power events from the System log ──────────────────────────────────
    Write-Section '[5/6] Power events from the System log'
    $bootType = @{ 0 = 'Cold boot (S5)'; 1 = 'Hibernation / Fast Startup'; 2 = 'Resume from sleep' }
    $idMeaning = @{
        12  = 'Kernel-General: OS started'
        13  = 'Kernel-General: OS shutting down'
        27  = 'Kernel-Boot: boot type'
        41  = 'Kernel-Power: unexpected shutdown (rebooted without a clean shutdown)'
        42  = 'Kernel-Power: entering sleep'
        107 = 'Kernel-Power: resume from sleep'
        109 = 'Kernel-Power: kernel preparing to shut down'
    }

    Add-Summary '--- [5/6] Power events -> power-events.csv / .txt ---'
    try {
        $events = Get-WinEvent -FilterHashtable @{
            LogName      = 'System'
            ProviderName = 'Microsoft-Windows-Kernel-General',
                           'Microsoft-Windows-Kernel-Power',
                           'Microsoft-Windows-Kernel-Boot'
            Id           = 12, 13, 27, 41, 42, 107, 109
            StartTime    = (Get-Date).AddDays(-$Days)
        } -ErrorAction Stop |
        Sort-Object TimeCreated |
        Select-Object @{n='TimeCreated'; e={ $_.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss') }},
                      Id,
                      @{n='Provider'; e={ $_.ProviderName -replace '^Microsoft-Windows-', '' }},
                      @{n='Meaning' ; e={ $idMeaning[[int]$_.Id] }},
                      @{n='Detail'  ; e={
                            if ($_.Id -eq 27 -and $_.Properties.Count -gt 0) {
                                $bt = [int]$_.Properties[0].Value
                                "BootType=$bt ($($bootType[$bt]))"
                            } else {
                                (($_.Message -replace '\s+', ' ').Trim())
                            }
                        }}

        $events | Export-Csv (Join-Path $outDir 'power-events.csv') -NoTypeInformation -Encoding UTF8
        $events | Format-Table TimeCreated, Id, Provider, Meaning -AutoSize |
            Out-File (Join-Path $outDir 'power-events.txt') -Encoding utf8 -Width 400

        $eventCount = @($events).Count
        Add-Summary ("{0} event(s) captured over the last {1} day(s)." -f $eventCount, $Days)
        Add-Summary ''
        Add-Summary 'Last 15 events:'
        $tail = (($events | Select-Object -Last 15 |
            Format-Table TimeCreated, Id, Meaning -AutoSize | Out-String -Width 400).TrimEnd())
        Add-Summary $tail
        Write-StepResult 'OK' ("{0} event(s) -> power-events.csv / .txt" -f $eventCount)
        if ($tail) {
            Write-Host $tail -ForegroundColor Gray
        }
    } catch {
        Add-Summary ("ERROR reading the System log: {0}" -f $_.Exception.Message)
        Write-StepResult 'ERROR' $_.Exception.Message
    }
    Add-Summary ''

    # ── 6. Fast Startup / Hibernate configuration ────────────────────────────
    Write-Section '[6/6] Fast Startup / Hibernate configuration'
    $pwrKey        = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power'
    $fastStartRaw  = Get-ItemProperty -Path $pwrKey -Name HiberbootEnabled -ErrorAction SilentlyContinue
    $fastStart     = if ($fastStartRaw) { $fastStartRaw.HiberbootEnabled } else { $null }

    $fastStartText = switch ($fastStart) {
        1       { '1 = enabled' }
        0       { '0 = disabled' }
        default { 'not set (default = enabled)' }
    }
    $hibernateText = if ($powercfgAText -match '(?m)^\s*Hibernate\s*$') { 'yes' } else { 'no' }
    $modernText    = if ($powercfgAText -match 'S0 Low Power Idle')     { 'yes' } else { 'no - classic S3 system' }

    Add-Summary '--- [6/6] Fast Startup / Hibernate configuration ---'
    Add-Summary ("Fast Startup (HiberbootEnabled) : {0}" -f $fastStartText)
    Add-Summary ("Hibernate available             : {0}" -f $hibernateText)
    Add-Summary ("Modern Standby (S0) available   : {0}" -f $modernText)
    Add-Summary ''

    Write-Host ('  {0,-32}: {1}' -f 'Fast Startup (HiberbootEnabled)', $fastStartText) -ForegroundColor Gray
    Write-Host ('  {0,-32}: {1}' -f 'Hibernate available', $hibernateText) -ForegroundColor Gray
    Write-Host ('  {0,-32}: {1}' -f 'Modern Standby (S0) available', $modernText) -ForegroundColor Gray

    # ── Write the summary, then package the bundle ───────────────────────────
    Add-Summary '=== Done ==='
    $summaryPath = Join-Path $outDir 'summary.txt'
    Set-Content -Path $summaryPath -Value ($summary -join [Environment]::NewLine) -Encoding UTF8

    $zipPath = $null
    if (-not $NoZip) {
        $zipPath = "$outDir.zip"
        if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::CreateFromDirectory($outDir, $zipPath)
    }

    Write-Section 'Done'
    Write-Host ('  Folder: {0}' -f $outDir) -ForegroundColor Green
    if ($zipPath) {
        Write-Host ('  Zip:    {0}' -f $zipPath) -ForegroundColor Green
        Write-Host '  Please return the .zip file for analysis.' -ForegroundColor DarkGray
    } else {
        Write-Host '  Please zip this folder and return it for analysis.' -ForegroundColor DarkGray
    }
    if (-not $isAdmin) {
        Write-Host '  NOTE: run elevated for SleepStudy / System Power / active power requests.' -ForegroundColor Yellow
    }

    if (-not $NoLaunch) {
        try { Start-Process explorer.exe $outDir } catch { }
    }

    exit $(if ($isAdmin) { $EXIT_OK } else { $EXIT_PARTIAL })
}
catch {
    Write-Host ''
    Write-Host ('ERROR: {0}' -f $_.Exception.Message) -ForegroundColor Red
    exit $EXIT_ERROR
}
