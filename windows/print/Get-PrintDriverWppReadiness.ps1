<#
.SYNOPSIS
    Flags machines that still have third-party v3/v4 print drivers (i.e. not yet
    Windows Protected Print (WPP) ready), so a fleet can be screened proactively
    before WPP is enforced.

.DESCRIPTION
    Windows Protected Print Mode (WPP) only allows printing through the inbox
    Microsoft IPP Class Driver and a small subset of inbox Microsoft print
    drivers. Any third-party v3 (kernel/user-mode classic) or v4 print driver is
    incompatible and will stop working once WPP is enforced.

    This script enumerates installed print drivers via Get-PrinterDriver, splits
    them into "WPP-safe" (Microsoft inbox drivers) and "blocking" (third-party
    v3/v4) buckets, maps which printers depend on the blocking drivers, and
    reports the machine's current WPP policy state.

    It works in two modes:
      1. Intune Proactive Remediation DETECTION script — emits a single-line JSON
         summary to STDOUT and sets the exit code:
            exit 0  = WPP-ready (no third-party v3/v4 drivers found)
            exit 1  = NOT WPP-ready (one or more blocking drivers found)
         This is the default behaviour, suitable for the "Detection" slot.
      2. Standalone inventory report — use -CsvPath / -JsonPath to write a
         per-driver report you can collect from a fleet (e.g. via a share,
         Log Analytics custom log, or a collected Intune script output).

    PowerShell 5.1 compatible. Read-only: it never installs, removes, or changes
    any driver or policy.

    This script is provided as-is, without warranty of any kind, express or implied,
    including but not limited to merchantability, fitness for a particular purpose,
    and noninfringement. Use at your own risk and validate behavior in a test
    environment before broad deployment.

.PARAMETER CsvPath
    Optional. Write the full per-driver classification to this CSV path.

.PARAMETER JsonPath
    Optional. Write the full structured result (drivers, mapped printers, WPP
    state) to this JSON path.

.PARAMETER IncludeSafeDrivers
    Include WPP-safe (Microsoft inbox) drivers in the CSV/JSON output. By default
    only the blocking third-party v3/v4 drivers are written to keep reports lean.

.PARAMETER Quiet
    Suppress the human-readable console summary. The single-line JSON summary is
    still emitted to STDOUT (so Intune log harvesting still works).

.EXAMPLE
    .\Get-PrintDriverWppReadiness.ps1
    Detection mode. Prints a one-line JSON summary and exits 1 if any third-party
    v3/v4 driver is present, otherwise 0. Drop this straight into an Intune
    Proactive Remediation "Detection script" slot.

.EXAMPLE
    .\Get-PrintDriverWppReadiness.ps1 -CsvPath C:\Temp\wpp-drivers.csv -IncludeSafeDrivers
    Writes a full per-driver report (safe + blocking) to CSV for fleet analysis.

.EXAMPLE
    Invoke-Command -ComputerName (Get-Content .\hosts.txt) -FilePath .\Get-PrintDriverWppReadiness.ps1
    Fan the detection out across a list of machines via remoting and collect the
    JSON summaries centrally.

.NOTES
    Uncertainty / scope caveats (read before enforcing on the back of this):
      * Microsoft has not published a complete, authoritative enumeration of the
        "subset of inbox Windows print drivers" WPP keeps. This script treats
        Microsoft-published inbox drivers (Manufacturer = Microsoft, plus an
        explicit allowlist of known inbox names) as WPP-safe and flags everything
        else. That matches the real risk signal (third-party drivers) but a
        third-party driver that happens to report "Microsoft" as manufacturer
        would be misclassified — spot-check the report before mass action.
      * A driver being present does not always mean a printer actively uses it;
        the report maps printers -> drivers so you can see actual usage.
      * v3 vs v4 is read from MajorVersion (3 / 4). Some inbox drivers report 0;
        those are evaluated by manufacturer/allowlist only.
#>

[CmdletBinding()]
param(
    [string]$CsvPath,

    [string]$JsonPath,

    [switch]$IncludeSafeDrivers,

    [switch]$Quiet
)

#region WPP-safe inbox driver allowlist

# Known inbox / WPP-compatible print driver names. Matched case-insensitively
# against the driver Name. Manufacturer = "Microsoft" is also treated as safe.
$Script:WppSafeDriverNames = @(
    'Microsoft IPP Class Driver'
    'Microsoft enhanced Point and Print compatibility driver'
    'Microsoft Print To PDF'
    'Microsoft XPS Document Writer'
    'Microsoft XPS Document Writer v4'
    'Microsoft Shared Fax Driver'
    'Microsoft Software Printer Driver'
    'Generic / Text Only'
    'Generic / Text'
    'Remote Desktop Easy Print'
)

#endregion

#region Helpers

function Test-WppSafeDriver {
    param(
        [Parameter(Mandatory)]
        [psobject]$Driver
    )

    $name = "$($Driver.Name)".Trim()
    $manufacturer = "$($Driver.Manufacturer)".Trim()

    if ($manufacturer -and $manufacturer -ieq 'Microsoft') { return $true }

    foreach ($safe in $Script:WppSafeDriverNames) {
        if ($name -ieq $safe) { return $true }
    }

    # IPP Class Driver / Microsoft inbox families sometimes carry a suffix.
    if ($name -imatch '^Microsoft\s.+(Class Driver|Document Writer|Print To PDF)') {
        return $true
    }

    return $false
}

function Get-WppPolicyState {
    # Reads the on-box WPP policy / state. Returns a small object; never throws.
    $result = [ordered]@{
        ProtectedPrintModeEnabled = $false
        GroupPolicyState          = $null
        Source                    = 'NotConfigured'
    }

    # Effective state (set by the feature itself)
    $statePaths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print'
    )
    foreach ($p in $statePaths) {
        try {
            $item = Get-ItemProperty -Path $p -ErrorAction Stop
            if ($null -ne $item.WindowsProtectedPrintMode) {
                $result.ProtectedPrintModeEnabled = [int]$item.WindowsProtectedPrintMode -eq 1
                if ($result.ProtectedPrintModeEnabled) { $result.Source = 'LocalState' }
            }
        }
        catch { }
    }

    # Group Policy / MDM-applied state
    $policyPaths = @(
        'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\WPP'
    )
    foreach ($p in $policyPaths) {
        try {
            $item = Get-ItemProperty -Path $p -ErrorAction Stop
            if ($null -ne $item.WindowsProtectedPrintGroupPolicyState) {
                $result.GroupPolicyState = [int]$item.WindowsProtectedPrintGroupPolicyState
                if ($result.GroupPolicyState -eq 1) {
                    $result.ProtectedPrintModeEnabled = $true
                    $result.Source = 'GroupPolicy'
                }
            }
        }
        catch { }
    }

    return [pscustomobject]$result
}

#endregion

#region Collect

$computerName = $env:COMPUTERNAME

# Map driver name -> printers using it (so we report actual dependency, not just presence)
$printersByDriver = @{}
try {
    foreach ($printer in (Get-Printer -ErrorAction Stop)) {
        $dn = "$($printer.DriverName)".Trim()
        if (-not $dn) { continue }
        if (-not $printersByDriver.ContainsKey($dn)) {
            $printersByDriver[$dn] = New-Object System.Collections.Generic.List[string]
        }
        [void]$printersByDriver[$dn].Add($printer.Name)
    }
}
catch {
    Write-Verbose "Get-Printer failed: $_"
}

$drivers = @()
try {
    $drivers = @(Get-PrinterDriver -ErrorAction Stop)
}
catch {
    # No PrintManagement module / Spooler off — report as an error result, exit 1
    $errSummary = [ordered]@{
        Computer    = $computerName
        WppReady    = $false
        Error       = "Get-PrinterDriver failed: $($_.Exception.Message)"
        Blocking    = -1
    }
    Write-Output ([pscustomobject]$errSummary | ConvertTo-Json -Compress)
    exit 1
}

$report = New-Object System.Collections.Generic.List[psobject]
foreach ($d in $drivers) {
    $isSafe = Test-WppSafeDriver -Driver $d
    $major = 0
    if ($null -ne $d.MajorVersion) { $major = [int]$d.MajorVersion }

    $driverKey = "$($d.Name)".Trim()
    $usedBy = @()
    if ($printersByDriver.ContainsKey($driverKey)) {
        $usedBy = $printersByDriver[$driverKey].ToArray()
    }

    $report.Add([pscustomobject]@{
        Computer       = $computerName
        DriverName     = $d.Name
        Manufacturer   = $d.Manufacturer
        DriverVersion  = $d.DriverVersion
        MajorVersion   = $major
        DriverModel    = if ($major -ge 4) { 'v4' } elseif ($major -eq 3) { 'v3' } else { "v$major" }
        Environment    = $d.PrinterEnvironment
        WppSafe        = $isSafe
        Blocking       = (-not $isSafe)
        InUse          = ($usedBy.Count -gt 0)
        UsedByPrinters = ($usedBy -join '; ')
    })
}

#endregion

#region Evaluate

$blocking = @($report | Where-Object { $_.Blocking })
$blockingV3 = @($blocking | Where-Object { $_.MajorVersion -eq 3 })
$blockingV4 = @($blocking | Where-Object { $_.MajorVersion -ge 4 })
$blockingInUse = @($blocking | Where-Object { $_.InUse })

$wppState = Get-WppPolicyState
$wppReady = ($blocking.Count -eq 0)

#endregion

#region Output (reports)

$outputSet = if ($IncludeSafeDrivers) { $report } else { $blocking }

if ($CsvPath) {
    try {
        $dir = Split-Path -Path $CsvPath -Parent
        if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $outputSet | Sort-Object Blocking -Descending |
            Select-Object Computer, DriverName, Manufacturer, DriverModel, MajorVersion,
                          DriverVersion, Environment, WppSafe, Blocking, InUse, UsedByPrinters |
            Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
        if (-not $Quiet) { Write-Host "CSV written: $CsvPath" -ForegroundColor Cyan }
    }
    catch { Write-Warning "Failed to write CSV: $_" }
}

if ($JsonPath) {
    try {
        $dir = Split-Path -Path $JsonPath -Parent
        if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $full = [ordered]@{
            Computer  = $computerName
            Timestamp = (Get-Date).ToString('o')
            WppReady  = $wppReady
            WppState  = $wppState
            Counts    = [ordered]@{
                TotalDrivers   = $report.Count
                Blocking       = $blocking.Count
                BlockingV3     = $blockingV3.Count
                BlockingV4     = $blockingV4.Count
                BlockingInUse  = $blockingInUse.Count
            }
            Drivers   = $outputSet
        }
        $full | ConvertTo-Json -Depth 6 | Out-File -FilePath $JsonPath -Encoding UTF8
        if (-not $Quiet) { Write-Host "JSON written: $JsonPath" -ForegroundColor Cyan }
    }
    catch { Write-Warning "Failed to write JSON: $_" }
}

#endregion

#region Console summary

if (-not $Quiet) {
    Write-Host ""
    Write-Host "  WPP readiness — $computerName" -ForegroundColor White
    Write-Host "  --------------------------------------------------" -ForegroundColor DarkGray
    Write-Host ("  WPP currently enforced : {0} ({1})" -f $wppState.ProtectedPrintModeEnabled, $wppState.Source) -ForegroundColor Gray
    Write-Host ("  Total print drivers    : {0}" -f $report.Count) -ForegroundColor Gray
    if ($wppReady) {
        Write-Host "  Result                 : WPP-READY (no third-party v3/v4 drivers)" -ForegroundColor Green
    }
    else {
        Write-Host ("  Result                 : NOT WPP-READY - {0} blocking driver(s)" -f $blocking.Count) -ForegroundColor Red
        Write-Host ("    v3 (classic)         : {0}" -f $blockingV3.Count) -ForegroundColor Yellow
        Write-Host ("    v4                   : {0}" -f $blockingV4.Count) -ForegroundColor Yellow
        Write-Host ("    actively in use      : {0}" -f $blockingInUse.Count) -ForegroundColor Yellow
        Write-Host ""
        foreach ($b in ($blocking | Sort-Object InUse -Descending)) {
            $usage = if ($b.InUse) { "in use by: $($b.UsedByPrinters)" } else { "not bound to a printer" }
            Write-Host ("    [{0}] {1}  ({2})  - {3}" -f $b.DriverModel, $b.DriverName, $b.Manufacturer, $usage) -ForegroundColor Yellow
        }
    }
    Write-Host ""
}

#endregion

#region Intune detection summary + exit code

$summary = [ordered]@{
    Computer      = $computerName
    WppReady      = $wppReady
    WppEnforced   = $wppState.ProtectedPrintModeEnabled
    Total         = $report.Count
    Blocking      = $blocking.Count
    BlockingV3    = $blockingV3.Count
    BlockingV4    = $blockingV4.Count
    BlockingInUse = $blockingInUse.Count
    Drivers       = @($blocking | ForEach-Object { "$($_.DriverModel):$($_.DriverName)" })
}
Write-Output ([pscustomobject]$summary | ConvertTo-Json -Compress)

if ($wppReady) { exit 0 } else { exit 1 }

#endregion
