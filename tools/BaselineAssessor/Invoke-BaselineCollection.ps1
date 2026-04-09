#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    BaselinePilot Data Collector — headless security baseline data collection for Windows clients.
.DESCRIPTION
    Collects security configuration data from a Windows 11 client machine for offline analysis
    by BaselinePilot. Gathers registry baselines, security policy exports, audit policy, Defender
    configuration, firewall state, BitLocker, Credential Guard, services, drivers, event logs,
    and more. Outputs a single JSON file with no external module dependencies.
.PARAMETER OutputPath
    Path for the output JSON file. Default: .\<hostname>_baseline_<timestamp>.json
.PARAMETER LookbackDays
    Event log query lookback window in days. Default: 30
.PARAMETER MaxEventsPerQuery
    Maximum events to retrieve per query group. Default: 10000
.PARAMETER SkipEventCollection
    Skip event log collection entirely for faster runs (~30s total).
.PARAMETER EventSummaryOnly
    Collect event counts and top-N stats only, not individual events.
.PARAMETER Quiet
    Suppress console output (for automation).
.NOTES
    Author : Anton Romanyuk
    Version: 1.0.0
    Date   : 2026-03-31
    Requires: PowerShell 5.1, Local Admin, No external modules
.EXAMPLE
    .\Invoke-BaselineCollection.ps1
    Runs full collection with defaults (30-day lookback, events included).
.EXAMPLE
    .\Invoke-BaselineCollection.ps1 -SkipEventCollection -OutputPath C:\Reports\baseline.json
    Fast run without event log collection, saving to a specific path.
.EXAMPLE
    .\Invoke-BaselineCollection.ps1 -EventSummaryOnly -LookbackDays 7 -Quiet
    Headless run with summary-only events and 7-day window.
#>
[CmdletBinding()]
param(
    [string]$OutputPath,
    [int]$LookbackDays       = 30,
    [int]$MaxEventsPerQuery  = 10000,
    [switch]$SkipEventCollection,
    [switch]$EventSummaryOnly,
    [switch]$Quiet
)

$ErrorActionPreference = 'Continue'
$Script:CollectorVersion = '0.1.0'
$Script:StartTime        = [DateTime]::Now
$Script:TotalAreas       = if ($SkipEventCollection) { 21 } else { 22 }
$Script:AreaResults      = @{}
$Script:Errors           = [System.Collections.ArrayList]::new()

# ═══════════════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════════════════
# PROGRESS & FORMATTING HELPERS
# ═══════════════════════════════════════════════════════════════════════

$Script:BW = 54  # Box width for console output formatting

# --- Console Box Drawing Helpers ---
# Render Unicode box-drawing characters for the startup banner.
# All suppress output when -Quiet is set.
function Write-BoxTop    { if ($Quiet) { return }; Write-Host "  $([char]0x2554)$([string]([char]0x2550) * $Script:BW)$([char]0x2557)" -ForegroundColor DarkCyan }
function Write-BoxMid    { if ($Quiet) { return }; Write-Host "  $([char]0x2560)$([string]([char]0x2550) * $Script:BW)$([char]0x2563)" -ForegroundColor DarkCyan }
function Write-BoxBottom { if ($Quiet) { return }; Write-Host "  $([char]0x255A)$([string]([char]0x2550) * $Script:BW)$([char]0x255D)" -ForegroundColor DarkCyan }

<#
.SYNOPSIS
    Renders a single line of text inside the console box.
.PARAMETER Text
    The text content to display. Truncated to box width if too long.
.PARAMETER Color
    Console foreground color for the text. Default: White.
#>
function Write-BoxLine {
    param([string]$Text, [string]$Color = 'White')
    if ($Quiet) { return }
    # Truncate if too long for the box
    if ($Text.Length -gt $Script:BW) { $Text = $Text.Substring(0, $Script:BW - 2) + '..' }
    $pad = $Script:BW - $Text.Length
    Write-Host "  $([char]0x2551)" -NoNewline -ForegroundColor DarkCyan
    Write-Host $Text -NoNewline -ForegroundColor $Color
    Write-Host "$(' ' * [math]::Max(0, $pad))$([char]0x2551)" -ForegroundColor DarkCyan
}

<#
.SYNOPSIS
    Renders a key-value pair inside the console box with aligned columns.
.PARAMETER Label
    The label text (left-aligned).
.PARAMETER Value
    The value text (right-aligned).
.PARAMETER LabelColor
    Console color for the label. Default: Gray.
.PARAMETER ValueColor
    Console color for the value. Default: White.
#>
function Write-BoxKV {
    param([string]$Label, [string]$Value, [string]$LabelColor = 'Gray', [string]$ValueColor = 'White')
    if ($Quiet) { return }
    $maxValLen = $Script:BW - 8 - $Label.Length
    if ($Value.Length -gt $maxValLen) { $Value = $Value.Substring(0, [math]::Max(3, $maxValLen - 2)) + '..' }
    $gap = $Script:BW - 6 - $Label.Length - $Value.Length
    Write-Host "  $([char]0x2551)" -NoNewline -ForegroundColor DarkCyan
    Write-Host "   " -NoNewline
    Write-Host $Label -NoNewline -ForegroundColor $LabelColor
    Write-Host "$(' ' * [math]::Max(1, $gap))" -NoNewline
    Write-Host $Value -NoNewline -ForegroundColor $ValueColor
    Write-Host "   $([char]0x2551)" -ForegroundColor DarkCyan
}

<#
.SYNOPSIS
    Writes a timestamped, color-coded status message to the console.
.PARAMETER Message
    The status message text.
.PARAMETER Level
    Severity level: INFO, ERROR, WARN, SUCCESS, CHECK, or SECTION.
    Controls the icon prefix and color. SECTION renders a sub-header bar.
#>
function Write-Status {
    param([string]$Message, [string]$Level = 'INFO')
    if ($Quiet) { return }
    $ts = (Get-Date).ToString('HH:mm:ss')
    $Icon = switch ($Level) {
        'ERROR'   { "$([char]0x2717)" }
        'WARN'    { "$([char]0x26A0)" }
        'SUCCESS' { "$([char]0x2713)" }
        'CHECK'   { "$([char]0x25BA)" }
        'SECTION' { "$([char]0x2500)" }
        default   { "$([char]0x00B7)" }
    }
    $Color = switch ($Level) {
        'ERROR'   { 'Red' }
        'WARN'    { 'Yellow' }
        'SUCCESS' { 'Green' }
        'CHECK'   { 'Cyan' }
        'SECTION' { 'DarkCyan' }
        default   { 'Gray' }
    }
    if ($Level -eq 'SECTION') {
        Write-Host ""
        Write-Host "  $([char]0x250C)$([char]0x2500)$([char]0x2500) " -NoNewline -ForegroundColor DarkCyan
        Write-Host $Message -NoNewline -ForegroundColor Cyan
        $barLen = [math]::Max(1, 48 - $Message.Length)
        Write-Host " $([string]([char]0x2500) * $barLen)$([char]0x2510)" -ForegroundColor DarkCyan
    } else {
        Write-Host "  " -NoNewline
        Write-Host $Icon -NoNewline -ForegroundColor $Color
        Write-Host " " -NoNewline
        Write-Host "[$ts] " -NoNewline -ForegroundColor DarkGray
        Write-Host $Message -ForegroundColor $(if ($Level -eq 'INFO') { 'White' } else { $Color })
    }
}

$Script:ProgressLinePending = $false

<#
.SYNOPSIS
    Displays structured progress output for each collection area step.
.DESCRIPTION
    Used by Invoke-CollectionArea and inline collection logic to show
    hierarchical progress: a 'start' line opens the step, 'info'/'warn'
    lines show sub-status, and 'done'/'error' lines close the step with
    elapsed time. Respects the -Quiet switch.
.PARAMETER Step
    The 1-based area step number (1-22).
.PARAMETER Total
    The total number of areas being collected.
.PARAMETER Name
    Display name for the current step or sub-step.
.PARAMETER Status
    Progress state: start, done, info, warn, or error.
.PARAMETER ElapsedSec
    Elapsed seconds (used with 'done' status).
.PARAMETER Detail
    Optional detail text appended to the status line.
#>
function Write-CollectorProgress {
    param(
        [int]$Step,
        [int]$Total,
        [string]$Name,
        [ValidateSet('start','done','info','warn','error')][string]$Status,
        [double]$ElapsedSec = 0,
        [string]$Detail = ''
    )
    if ($Quiet) { return }
    $ts  = Get-Date -Format 'HH:mm:ss'
    $num = "$Step".PadLeft(2)
    switch ($Status) {
        'start' {
            $dots = "$([char]0x2500)" * [math]::Max(1, 38 - $Name.Length)
            Write-Host "  $([char]0x25BA) " -NoNewline -ForegroundColor Cyan
            Write-Host "[$ts] " -NoNewline -ForegroundColor DarkGray
            Write-Host "[$num/$Total] " -NoNewline -ForegroundColor DarkCyan
            Write-Host "$Name " -NoNewline -ForegroundColor White
            Write-Host "$dots " -NoNewline -ForegroundColor DarkGray
            $Script:ProgressLinePending = $true
        }
        'done' {
            if (-not $Script:ProgressLinePending) {
                # Sub-lines were emitted; print done on a new indented line
                $dots2 = "$([char]0x2500)" * 20
                Write-Host "     $([char]0x2514)$([char]0x2500) " -NoNewline -ForegroundColor DarkGray
                Write-Host "done " -NoNewline -ForegroundColor DarkGray
                Write-Host "$dots2 " -NoNewline -ForegroundColor DarkGray
            }
            $secs = [math]::Round($ElapsedSec, 1)
            $d = if ($Detail) { "  [$Detail]" } else { '' }
            Write-Host "$([char]0x2713) " -NoNewline -ForegroundColor Green
            Write-Host "${secs}s" -NoNewline -ForegroundColor DarkGray
            if ($d) { Write-Host $d -ForegroundColor DarkGray } else { Write-Host '' }
            $Script:ProgressLinePending = $false
        }
        'info' {
            # Close pending start line first
            if ($Script:ProgressLinePending) { Write-Host ''; $Script:ProgressLinePending = $false }
            Write-Host "     $([char]0x2502)  " -NoNewline -ForegroundColor DarkGray
            Write-Host "$Name " -NoNewline -ForegroundColor Gray
            if ($Detail) { Write-Host "$Detail" -ForegroundColor DarkGray } else { Write-Host '' }
        }
        'warn' {
            if ($Script:ProgressLinePending) { Write-Host ''; $Script:ProgressLinePending = $false }
            Write-Host "     $([char]0x2502)  " -NoNewline -ForegroundColor DarkGray
            Write-Host "$([char]0x26A0) $Name " -NoNewline -ForegroundColor Yellow
            if ($Detail) { Write-Host "$Detail" -ForegroundColor Yellow } else { Write-Host '' }
        }
        'error' {
            if ($Script:ProgressLinePending) { Write-Host ''; $Script:ProgressLinePending = $false }
            Write-Host "     $([char]0x2514)  " -NoNewline -ForegroundColor DarkGray
            Write-Host "$([char]0x2717) $Name " -NoNewline -ForegroundColor Red
            Write-Host "FAILED" -NoNewline -ForegroundColor Red
            if ($Detail) { Write-Host ": $Detail" -ForegroundColor DarkRed } else { Write-Host '' }
        }
    }
}

<#
.SYNOPSIS
    Executes a collection area script block with timing, error handling, and progress output.
.DESCRIPTION
    Wraps each collection area (Areas 1-22) in a standard pattern: emits a 'start' progress
    line, executes the script block, catches and records any errors, and emits 'done' or 'error'
    with elapsed time. Returns the script block result or $null on failure.
.PARAMETER Step
    The 1-based area number.
.PARAMETER Name
    Display name for the collection area.
.PARAMETER Script
    ScriptBlock containing the collection logic. Should return a hashtable with collected data
    and an optional '_detail' key for the progress summary.
.OUTPUTS
    [hashtable] The collection result, or $null if the script block threw an exception.
#>
function Invoke-CollectionArea {
    param(
        [int]$Step,
        [string]$Name,
        [ScriptBlock]$Script
    )
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Write-CollectorProgress -Step $Step -Total $Script:TotalAreas -Name $Name -Status 'start'

    try {
        $result = & $Script
        $sw.Stop()
        $detail = if ($result -is [hashtable] -and $result._detail) { $result._detail } else { '' }
        Write-CollectorProgress -Step $Step -Total $Script:TotalAreas -Name $Name -Status 'done' `
            -ElapsedSec $sw.Elapsed.TotalSeconds -Detail $detail
        return $result
    } catch {
        $sw.Stop()
        Write-CollectorProgress -Step $Step -Total $Script:TotalAreas -Name $Name -Status 'error' `
            -ElapsedSec $sw.Elapsed.TotalSeconds -Detail $_.Exception.Message
        [void]$Script:Errors.Add(@{ Area = $Name; Error = $_.Exception.Message })
        return $null
    }
}


# ═══════════════════════════════════════════════════════════════════════
# SAFE REGISTRY READ HELPERS
# ═══════════════════════════════════════════════════════════════════════

<#
.SYNOPSIS
    Reads a single registry value, returning $null if the key or value doesn't exist.
.PARAMETER Path
    Registry path without the provider prefix (e.g. 'HKLM\SOFTWARE\...').
.PARAMETER Name
    The value name to read.
.OUTPUTS
    The registry value, or $null if not found.
#>
function Read-RegistryValue {
    param([string]$Path, [string]$Name)
    try {
        $val = Get-ItemProperty -Path "Registry::$Path" -Name $Name -ErrorAction Stop
        return $val.$Name
    } catch { return $null }
}

<#
.SYNOPSIS
    Reads all values from a registry key, returning an empty hashtable if the key doesn't exist.
.PARAMETER Path
    Registry path without the provider prefix (e.g. 'HKLM\SOFTWARE\...').
.OUTPUTS
    [hashtable] Name-value pairs for all values under the key.
#>
function Read-RegistryValues {
    param([string]$Path)
    try {
        $item = Get-Item -Path "Registry::$Path" -ErrorAction Stop
        $result = @{}
        foreach ($name in $item.GetValueNames()) {
            $result[$name] = $item.GetValue($name)
        }
        return $result
    } catch { return @{} }
}

# ═══════════════════════════════════════════════════════════════════════
# BANNER
# ═══════════════════════════════════════════════════════════════════════

$hostname = $env:COMPUTERNAME
$osInfo   = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
$osBuild  = if ($osInfo) { $osInfo.BuildNumber } else { 'Unknown' }
$osVer    = if ($osInfo) { $osInfo.Version } else { 'Unknown' }
$osEdition = if ($osInfo) { $osInfo.Caption } else { 'Windows' }
$dv = try { (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop).DisplayVersion } catch { '' }

if (-not $Quiet) {
    Write-Host ''
    Write-Host '   ___                _ _            ___ _ _     _   ' -ForegroundColor Cyan
    Write-Host '  | _ ) __ _ ___ ___ | (_)_ _  ___  | _ (_) |___| |_ ' -ForegroundColor Cyan
    Write-Host '  | _ \/ _` (_-</ -_)| | | '' \/ -_) |  _/ | / _ \  _|' -ForegroundColor Cyan
    Write-Host '  |___/\__,_/__/\___||_|_|_||_\___| |_| |_|_\___/\__|' -ForegroundColor Cyan
    Write-Host ''
    Write-Host '  v' -NoNewline -ForegroundColor DarkGray
    Write-Host $Script:CollectorVersion -NoNewline -ForegroundColor White
    Write-Host '  |  ' -NoNewline -ForegroundColor DarkGray
    Write-Host 'Data Collector' -NoNewline -ForegroundColor Cyan
    Write-Host '  |  ' -NoNewline -ForegroundColor DarkGray
    Write-Host 'Security Baseline' -ForegroundColor DarkCyan
    Write-Host ''

    Write-BoxTop
    Write-BoxLine ' '
    Write-BoxLine "  Host:      $hostname" 'White'
    Write-BoxLine "  OS:        $osEdition" 'White'
    Write-BoxLine "  Build:     $osBuild (v$osVer) $dv" 'White'
    Write-BoxLine "  Started:   $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" 'Gray'
    Write-BoxLine "  Areas:     $Script:TotalAreas collection targets" 'Gray'
    Write-BoxLine ' '
    Write-BoxBottom
    Write-Host ''
}

# ═══════════════════════════════════════════════════════════════════════
# AREA 1: SYSTEM INFORMATION
# ═══════════════════════════════════════════════════════════════════════

$systemInfo = Invoke-CollectionArea -Step 1 -Name 'System Information' -Script {
    $os  = Get-CimInstance Win32_OperatingSystem
    $cs  = Get-CimInstance Win32_ComputerSystem
    $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
    $tpm = try { Get-CimInstance -Namespace 'root\cimv2\Security\MicrosoftTpm' -ClassName Win32_Tpm -ErrorAction Stop } catch { $null }
    $sb  = try { Confirm-SecureBootUEFI -ErrorAction Stop } catch { $null }
    $dv  = try {
        $reg = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
        $reg.DisplayVersion
    } catch { '' }

    # Windows 11 build-to-version mapping with support status
    $buildNum = $os.BuildNumber.ToString()
    $versionName = $dv
    $isSupported = $null
    $endOfService = $null
    $isWindows11 = $os.Caption -match 'Windows 11'
    switch ($buildNum) {
        '22000' { $versionName = '21H2'; $endOfService = '2024-10-08'; $isSupported = $false }
        '22621' { $versionName = '22H2'; $endOfService = '2025-10-14'; $isSupported = $true }
        '22631' { $versionName = '23H2'; $endOfService = '2026-11-10'; $isSupported = $true }
        '26100' { $versionName = '24H2'; $endOfService = '2027-11-09'; $isSupported = $true }
        '26200' { $versionName = '25H2'; $endOfService = '2028-11-14'; $isSupported = $true }
    }
    if (-not $versionName) { $versionName = 'Unknown' }

    $result = @{
        ComputerName     = $env:COMPUTERNAME
        OSCaption        = $os.Caption
        osVersion        = $os.Version
        osBuild          = $os.BuildNumber
        osEdition        = $os.Caption
        displayVersion   = $dv
        versionName      = $versionName
        isWindows11      = $isWindows11
        isSupported      = $isSupported
        endOfService     = $endOfService
        hostname         = $env:COMPUTERNAME
        domain           = $cs.Domain
        ramGB            = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
        cpuName          = $cpu.Name
        cpuCores         = $cpu.NumberOfCores
        tpmVersion       = if ($tpm) { $tpm.SpecVersion -replace ',.*' } else { 'Not found' }
        secureBootEnabled = $sb
        installDate      = $os.InstallDate.ToString('o')
    }
    try { $result['freeSpaceGB'] = [math]::Round((Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" -EA Stop).FreeSpace / 1GB, 1) } catch { $result['freeSpaceGB'] = $null }
    try { $result['lastBootTime'] = $os.LastBootUpTime.ToString('o') } catch { $result['lastBootTime'] = $null }
    try { $result['uptimeDays'] = [math]::Round(((Get-Date) - $os.LastBootUpTime).TotalDays, 1) } catch { $result['uptimeDays'] = $null }

    $result['_detail'] = "$versionName (Build $buildNum)$(if (-not $isSupported) { ' UNSUPPORTED' })"
    $result
}

# ═══════════════════════════════════════════════════════════════════════
# AREA 2: JOIN TYPE DETECTION
# ═══════════════════════════════════════════════════════════════════════

$joinType = Invoke-CollectionArea -Step 2 -Name 'Join Type Detection' -Script {
    $dsreg = dsregcmd /status 2>&1
    $parse = { param($pattern) ($dsreg | Select-String $pattern) -match 'YES' }
    $aadJoined    = & $parse 'AzureAdJoined\s*:\s*YES'
    $domainJoined = & $parse 'DomainJoined\s*:\s*YES'

    @{
        azureAdJoined  = [bool]$aadJoined
        domainJoined   = [bool]$domainJoined
        hybridJoined   = [bool]($aadJoined -and $domainJoined)
        workgroup      = [bool](-not $aadJoined -and -not $domainJoined)
    }
}

# ═══════════════════════════════════════════════════════════════════════
# AREA 3: APPLIED POLICIES (GPResult)
# ═══════════════════════════════════════════════════════════════════════

$appliedGPOs = Invoke-CollectionArea -Step 3 -Name 'Applied Policies' -Script {
    # Check join type from Area 2 result to decide strategy
    $isDomainJoined = $joinType -and $joinType.domainJoined
    $isEntraOnly    = $joinType -and $joinType.azureAdJoined -and -not $joinType.domainJoined

    if (-not $isDomainJoined -and -not $isEntraOnly) {
        # Workgroup - no GPOs
        Write-CollectorProgress -Step 3 -Total $Script:TotalAreas -Name 'Workgroup device' -Status 'info' `
            -Detail 'skipping gpresult (no domain)'
        @{ appliedGPOs = @(); joinType = 'Workgroup'; _detail = 'Workgroup - no GPOs' }
        return
    }

    # Domain-joined, hybrid, or Entra ID - run gpresult /x (XML) with timeout
    # Entra-only devices still have local policy and MDM-applied settings visible in gpresult
    $gpTimeout = 120  # Sufficient for all join types including Entra (Start-Job isolates hangs)
    $gpXmlPath = Join-Path $env:TEMP "gpresult_$([guid]::NewGuid().ToString('N')).xml"
    Write-CollectorProgress -Step 3 -Total $Script:TotalAreas -Name 'Running gpresult /x' -Status 'info' `
        -Detail "timeout ${gpTimeout}s"
    $job = Start-Job -ScriptBlock {
        param($outPath)
        gpresult /scope computer /x $outPath /f 2>&1
    } -ArgumentList $gpXmlPath
    $finished = $job | Wait-Job -Timeout $gpTimeout
    if (-not $finished) {
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        Remove-Item $gpXmlPath -Force -ErrorAction SilentlyContinue
        Write-CollectorProgress -Step 3 -Total $Script:TotalAreas -Name 'gpresult timeout' -Status 'warn' `
            -Detail "timed out after ${gpTimeout}s - trying WMI fallback"

        # WMI fallback: RSOP namespace has GPO data from last policy processing
        $gpos = [System.Collections.ArrayList]::new()
        $denied = [System.Collections.ArrayList]::new()
        try {
            $rsopGpos = Get-CimInstance -Namespace 'root\RSOP\Computer' -ClassName 'RSOP_GPO' -ErrorAction Stop
            foreach ($g in $rsopGpos) {
                if (-not $g.Name) { continue }
                $gpoObj = @{ Name = $g.Name }
                if ($g.GuidName) { $gpoObj['Guid'] = $g.GuidName }
                if ($g.SOM)      { $gpoObj['LinkPath'] = $g.SOM }
                if ($g.AccessDenied) {
                    [void]$denied.Add($gpoObj)
                } elseif ($g.Enabled -and $g.FilterAllowed) {
                    [void]$gpos.Add($gpoObj)
                }
            }
            Write-CollectorProgress -Step 3 -Total $Script:TotalAreas -Name 'WMI fallback' -Status 'info' `
                -Detail "$($gpos.Count) GPOs from RSOP"
        } catch {
            Write-CollectorProgress -Step 3 -Total $Script:TotalAreas -Name 'WMI fallback failed' -Status 'warn' `
                -Detail $_.Exception.Message
        }

        # Also try GP History registry as a secondary fallback
        if ($gpos.Count -eq 0) {
            try {
                $histPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\History'
                if (Test-Path $histPath) {
                    $histKeys = Get-ChildItem $histPath -ErrorAction SilentlyContinue
                    foreach ($k in $histKeys) {
                        $displayName = try { (Get-ItemProperty $k.PSPath -Name 'DisplayName' -ErrorAction Stop).DisplayName } catch { $null }
                        $guidName    = try { (Get-ItemProperty $k.PSPath -Name 'GPOName' -ErrorAction Stop).GPOName } catch { $null }
                        $existing    = $gpos | ForEach-Object { $_.Name }
                        if ($displayName -and $displayName -notin $existing) {
                            $gpoObj = @{ Name = $displayName }
                            if ($guidName) { $gpoObj['Guid'] = $guidName }
                            [void]$gpos.Add($gpoObj)
                        }
                    }
                    if ($gpos.Count -gt 0) {
                        Write-CollectorProgress -Step 3 -Total $Script:TotalAreas -Name 'GP History fallback' -Status 'info' `
                            -Detail "$($gpos.Count) GPOs from registry"
                    }
                }
            } catch { }
        }

        $result = @{ appliedGPOs = @($gpos); timedOut = $true; fallback = 'WMI+Registry'; _detail = "gpresult timed out, $($gpos.Count) GPOs via fallback" }
        if ($denied.Count -gt 0) { $result['deniedGPOs'] = @($denied) }
        $result
        return
    }
    $null = Receive-Job $job
    Remove-Job $job -Force -ErrorAction SilentlyContinue

    if (-not (Test-Path $gpXmlPath)) {
        Write-CollectorProgress -Step 3 -Total $Script:TotalAreas -Name 'gpresult failed' -Status 'warn' `
            -Detail 'XML file not created'
        @{ appliedGPOs = @(); _detail = 'gpresult /x produced no output' }
        return
    }

    try {
        [xml]$doc = Get-Content $gpXmlPath -Raw -ErrorAction Stop
    } catch {
        @{ appliedGPOs = @(); _detail = "XML parse failed: $($_.Exception.Message)" }
        return
    } finally {
        Remove-Item $gpXmlPath -Force -ErrorAction SilentlyContinue
    }

    # Extract applied GPOs from XML - element names are always English
    # Structure: Rsop > ComputerResults > GPO nodes
    # FilterAllowed=true + AccessDenied=false = applied; AccessDenied=true = denied
    $gpos = [System.Collections.ArrayList]::new()
    $denied = [System.Collections.ArrayList]::new()

    # Use namespace manager for the RSOP namespace
    $nsMgr = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $rsopNs = $doc.DocumentElement.NamespaceURI
    if ($rsopNs) { $nsMgr.AddNamespace('rsop', $rsopNs) }

    # Try with namespace first, then without
    $gpoNodes = $null
    if ($rsopNs) {
        $gpoNodes = $doc.SelectNodes('//rsop:ComputerResults/rsop:GPO', $nsMgr)
    }
    if (-not $gpoNodes -or $gpoNodes.Count -eq 0) {
        $gpoNodes = $doc.SelectNodes('//ComputerResults/GPO')
    }

    if ($gpoNodes) {
        foreach ($gpo in $gpoNodes) {
            $name = $gpo.Name
            if (-not $name) { continue }
            $isAllowed  = $gpo.FilterAllowed -ne 'false'
            $isDenied   = $gpo.AccessDenied  -eq 'true'

            # Build a rich GPO object with available metadata
            $gpoObj = @{ Name = $name }
            if ($gpo.Path -and $gpo.Path.Identifier) {
                $gpoObj['Guid'] = $gpo.Path.Identifier.InnerText
            }
            if ($gpo.Link -and $gpo.Link.SOMPath) {
                $gpoObj['LinkPath'] = $gpo.Link.SOMPath
            }
            if ($gpo.VersionDirectory) { $gpoObj['VersionAD'] = $gpo.VersionDirectory }
            if ($gpo.VersionSysvol)    { $gpoObj['VersionSysvol'] = $gpo.VersionSysvol }
            if ($gpo.IsValid)          { $gpoObj['IsValid'] = $gpo.IsValid }

            if ($isDenied) {
                [void]$denied.Add($gpoObj)
            } elseif ($isAllowed) {
                [void]$gpos.Add($gpoObj)
            }
        }
    }

    $result = @{ appliedGPOs = @($gpos); _detail = "$($gpos.Count) GPOs" }
    if ($denied.Count -gt 0) { $result['deniedGPOs'] = @($denied) }
    $result
}

# ═══════════════════════════════════════════════════════════════════════
# AREA 4: MDM ENROLLMENT
# ═══════════════════════════════════════════════════════════════════════

$mdmEnrollment = Invoke-CollectionArea -Step 4 -Name 'MDM Enrollment' -Script {
    $enrolled = $false
    $provider = ''
    $regPath  = 'HKLM:\SOFTWARE\Microsoft\Enrollments'
    if (Test-Path $regPath) {
        $subs = Get-ChildItem $regPath -ErrorAction SilentlyContinue
        foreach ($sub in $subs) {
            $upn = (Get-ItemProperty $sub.PSPath -Name 'UPN' -ErrorAction SilentlyContinue).UPN
            $prov = (Get-ItemProperty $sub.PSPath -Name 'ProviderID' -ErrorAction SilentlyContinue).ProviderID
            if ($upn -or $prov) {
                $enrolled = $true
                $provider = if ($prov) { $prov } else { 'Unknown' }
                break
            }
        }
    }

    # MDMWinsOverGP — determines if MDM settings take precedence over GPO
    $mdmWinsOverGP = $null
    try {
        $mdmWinsReg = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\MDM' -Name 'MDMWinsOverGP' -ErrorAction SilentlyContinue
        if ($mdmWinsReg) { $mdmWinsOverGP = [int]$mdmWinsReg.MDMWinsOverGP }
    } catch {}

    # PolicyManager managed areas — shows which CSP areas have Intune policies applied
    $managedAreas = @{}
    $pmPath = 'HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device'
    if ($enrolled -and (Test-Path $pmPath)) {
        $areas = Get-ChildItem $pmPath -ErrorAction SilentlyContinue
        foreach ($area in $areas) {
            $areaName = $area.PSChildName
            # Check for _ProviderSet entries (indicates Intune is managing this area)
            $props = Get-ItemProperty $area.PSPath -ErrorAction SilentlyContinue
            $managedSettings = @()
            if ($props) {
                $providerSetKeys = $props.PSObject.Properties | Where-Object { $_.Name -match '_ProviderSet$' -and $_.Value -eq 1 }
                foreach ($pk in $providerSetKeys) {
                    $settingName = $pk.Name -replace '_ProviderSet$', ''
                    $managedSettings += $settingName
                }
            }
            if ($managedSettings.Count -gt 0) {
                $managedAreas[$areaName] = @{
                    settingCount = $managedSettings.Count
                    settings     = $managedSettings
                }
            }
        }
        Write-CollectorProgress -Step 4 -Total $Script:TotalAreas -Name 'PolicyManager scan' -Status 'info' `
            -Detail "$($managedAreas.Count) MDM-managed areas"
    }

    $detail = if ($enrolled) { "$provider$(if ($mdmWinsOverGP -eq 1) { ' (MDM wins over GPO)' } else { '' })" } else { 'Not enrolled' }
    @{
        mdmEnrolled    = $enrolled
        mdmProvider    = $provider
        mdmWinsOverGP  = $mdmWinsOverGP
        managedAreas   = $managedAreas
        _detail        = $detail
    }
}

# ═══════════════════════════════════════════════════════════════════════
# AREA 5: SECURITY POLICY EXPORT (secedit)
# ═══════════════════════════════════════════════════════════════════════

$securityPolicy = Invoke-CollectionArea -Step 5 -Name 'Security Policy Export' -Script {
    $tmpFile = Join-Path $env:TEMP "bp_secedit_$(Get-Random).inf"
    try {
        $null = secedit /export /cfg $tmpFile /quiet 2>&1
        $content = Get-Content $tmpFile -ErrorAction Stop
        $result  = @{}
        $section = ''
        foreach ($line in $content) {
            if ($line -match '^\[(.+)\]') { $section = $Matches[1]; continue }
            if ($line -match '^(.+?)\s*=\s*(.+)$') {
                $key = "$section`_$($Matches[1].Trim())"
                $val = $Matches[2].Trim()
                # Try to convert numeric values
                if ($val -match '^\d+$') { $val = [int]$val }
                $result[$key] = $val
            }
        }
        $result['_detail'] = "$($result.Count) settings"
        $result
    } finally {
        if (Test-Path $tmpFile) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }
    }
}

# ═══════════════════════════════════════════════════════════════════════
# AREA 6: AUDIT POLICY
# ═══════════════════════════════════════════════════════════════════════

$auditPolicy = Invoke-CollectionArea -Step 6 -Name 'Audit Policy' -Script {
    # Use /r (CSV) mode first - locale-independent column structure
    $rawCsv = auditpol /get /category:* /r 2>&1
    $result = @{}

    # Check if /r succeeded (requires admin; returns CSV with Machine Name,Policy Target,Subcategory,...)
    $csvLines = @($rawCsv | Where-Object { $_ -is [string] -and $_ -match ',' })
    if ($csvLines.Count -gt 1) {
        # CSV format: Machine Name,Policy Target,Subcategory,Subcategory GUID,Inclusion Setting,Exclusion Setting
        # Column 2 = Subcategory (display name, localized), Column 3 = GUID (locale-independent), Column 4 = Inclusion Setting
        foreach ($csvLine in $csvLines | Select-Object -Skip 1) {
            $cols = $csvLine -split ','
            if ($cols.Count -ge 5) {
                $subcat  = $cols[2].Trim()
                $guid    = $cols[3].Trim() -replace '[{}]', ''
                $setting = $cols[4].Trim()
                if ($subcat -and $setting) {
                    # Use GUID as key for locale-independent matching, store display name alongside
                    $key = if ($guid) { $guid } else { $subcat }
                    $result[$key] = @{
                        name    = $subcat
                        setting = $setting
                    }
                }
            }
        }
    } else {
        # /r failed (not admin or other error) - fall back to text mode with multi-locale regex
        $rawText = auditpol /get /category:* 2>&1
        foreach ($line in $rawText) {
            # Match subcategory lines: 2+ leading spaces, then name, then 2+ spaces, then setting
            # Settings: EN (Success and Failure|Success|Failure|No Auditing)
            #           DE (Erfolg und Fehler|Erfolg|Fehler|Keine Ueberwachung)
            #           FR (Reussite et echec|Reussite|Echec|Pas d'audit)
            #           ES (Correcto e incorrecto|Correcto|Incorrecto|Sin auditoria)
            if ($line -match '^\s{2}(\S.+?)\s{2,}(\S.+?)\s*$') {
                $subcat  = $Matches[1].Trim()
                $setting = $Matches[2].Trim()
                # Exclude category headers (no setting value, or suspiciously short)
                if ($setting.Length -gt 2) {
                    $result[$subcat] = $setting
                }
            }
        }
    }

    $result['_detail'] = "$($result.Count) subcategories"
    $result
}

# ═══════════════════════════════════════════════════════════════════════
# AREA 7: REGISTRY BASELINES
# ═══════════════════════════════════════════════════════════════════════

$registryBaselines = Invoke-CollectionArea -Step 7 -Name 'Registry Baselines' -Script {
    # Key registry paths from Intune + GPO security baselines
    # Both GPO (SOFTWARE\Policies\) and CSP/MDM (SOFTWARE\Microsoft\) paths are read
    # so the assessor can evaluate compliance regardless of delivery mechanism
    $paths = @(
        # Device Lock / Personalization
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\Personalization'; Values = @('NoLockScreenCamera','NoLockScreenSlideshow') }
        # Smart Screen (GPO path)
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\System'; Values = @('EnableSmartScreen','ShellSmartScreenLevel') }
        # Smart Screen (CSP/MDM path)
        @{ Path = 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer'; Values = @('SmartScreenEnabled') }
        # Defender (GPO paths)
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows Defender'; Values = @('DisableAntiSpyware','PUAProtection') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection'; Values = @('DisableBehaviorMonitoring','DisableRealtimeMonitoring','DisableOnAccessProtection','DisableScanOnRealtimeEnable','DisableIOAVProtection') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet'; Values = @('SpynetReporting','SubmitSamplesConsent') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\MpEngine'; Values = @('MpCloudBlockLevel') }
        # Defender (CSP/MDM - Intune writes here instead of Policies path)
        @{ Path = 'HKLM\SOFTWARE\Microsoft\Windows Defender'; Values = @('DisableAntiSpyware','PUAProtection') }
        @{ Path = 'HKLM\SOFTWARE\Microsoft\Windows Defender\Real-Time Protection'; Values = @('DisableRealtimeMonitoring','DisableBehaviorMonitoring','DisableOnAccessProtection','DisableScanOnRealtimeEnable','DisableIOAVProtection') }
        @{ Path = 'HKLM\SOFTWARE\Microsoft\Windows Defender\SpyNet'; Values = @('SpynetReporting','SubmitSamplesConsent') }
        @{ Path = 'HKLM\SOFTWARE\Microsoft\Windows Defender\MpEngine'; Values = @('MpCloudBlockLevel') }
        @{ Path = 'HKLM\SOFTWARE\Microsoft\Windows Defender\Features'; Values = @('TamperProtection') }
        # BitLocker
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\FVE'; Values = @('UseAdvancedStartup','EnableBDEWithNoTPM','UseTPM','UseTPMPIN','UseTPMKey','UseTPMKeyPIN','MinimumPIN','EncryptionMethodWithXtsOs','EncryptionMethodWithXtsFdv','EncryptionMethodWithXtsRdv') }
        # UAC
        @{ Path = 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'; Values = @('EnableLUA','ConsentPromptBehaviorAdmin','ConsentPromptBehaviorUser','PromptOnSecureDesktop','EnableInstallerDetection','EnableSecureUIAPaths','EnableVirtualization','FilterAdministratorToken','ValidateAdminCodeSignatures','LegalNoticeCaption','LegalNoticeText') }
        # Credential Guard / VBS
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Control\DeviceGuard'; Values = @('EnableVirtualizationBasedSecurity','RequirePlatformSecurityFeatures','Locked') }
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Control\Lsa'; Values = @('LmCompatibilityLevel','NoLMHash','RestrictAnonymousSAM','RestrictAnonymous','EveryoneIncludesAnonymous','RestrictRemoteSAM','DisableDomainCreds','LimitBlankPasswordUse','SCENoApplyLegacyAuditPolicy','RunAsPPL') }
        # Network Security
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters'; Values = @('RequireSecuritySignature','EnableSecuritySignature','RestrictNullSessAccess','NullSessionPipes','NullSessionShares','SMB1') }
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Services\LanManWorkstation\Parameters'; Values = @('RequireSecuritySignature','EnableSecuritySignature','EnablePlainTextPassword') }
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Services\LDAP'; Values = @('LDAPClientIntegrity') }
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters'; Values = @('RequireSignOrSeal','SealSecureChannel','SignSecureChannel','RequireStrongKey','DisablePasswordChange','MaximumPasswordAge') }
        # Remote Desktop
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services'; Values = @('fDenyTSConnections','UserAuthentication','MinEncryptionLevel','SecurityLayer','fPromptForPassword','DisablePasswordSaving','fDisableCdm','fEncryptRPCTraffic','DeleteTempDirsOnExit','PerSessionTempDir','MaxDisconnectionTime','MaxIdleTime') }
        # Windows Update
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'; Values = @('NoAutoUpdate','AUOptions','ScheduledInstallDay','ScheduledInstallTime','NoAutoRebootWithLoggedOnUsers') }
        # PowerShell
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging'; Values = @('EnableScriptBlockLogging','EnableScriptBlockInvocationLogging') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\PowerShell\Transcription'; Values = @('EnableTranscripting','OutputDirectory','EnableInvocationHeader') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ModuleLogging'; Values = @('EnableModuleLogging') }
        # WinRM
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\WinRM\Client'; Values = @('AllowBasic','AllowUnencryptedTraffic','AllowDigest') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\WinRM\Service'; Values = @('AllowBasic','AllowUnencryptedTraffic','AllowAutoConfig','DisableRunAs') }
        # AutoPlay
        @{ Path = 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer'; Values = @('NoDriveTypeAutoRun','NoAutorun') }
        # Encryption / TLS
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 2.0\Client'; Values = @('Enabled','DisabledByDefault') }
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 2.0\Server'; Values = @('Enabled','DisabledByDefault') }
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Client'; Values = @('Enabled','DisabledByDefault') }
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Server'; Values = @('Enabled','DisabledByDefault') }
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Client'; Values = @('Enabled','DisabledByDefault') }
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Server'; Values = @('Enabled','DisabledByDefault') }
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client'; Values = @('Enabled','DisabledByDefault') }
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Server'; Values = @('Enabled','DisabledByDefault') }
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client'; Values = @('Enabled','DisabledByDefault') }
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server'; Values = @('Enabled','DisabledByDefault') }
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.3\Client'; Values = @('Enabled','DisabledByDefault') }
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.3\Server'; Values = @('Enabled','DisabledByDefault') }
        # MSS (Legacy) Settings
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters'; Values = @('DisableIPSourceRouting','EnableICMPRedirect','KeepAliveTime','PerformRouterDiscovery','TcpMaxDataRetransmissions') }
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters'; Values = @('DisableIPSourceRouting','TcpMaxDataRetransmissions') }
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Services\Netbt\Parameters'; Values = @('NoNameReleaseOnDemand','NodeType') }
        # Spectre / Meltdown Mitigations
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management'; Values = @('FeatureSettingsOverride','FeatureSettingsOverrideMask') }
        # Screen Saver / Lock
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\Control Panel\Desktop'; Values = @('ScreenSaveActive','ScreenSaverIsSecure','ScreenSaveTimeOut') }
        # Windows Defender Firewall logging
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\WindowsFirewall\DomainProfile\Logging'; Values = @('LogFilePath','LogFileSize','LogDroppedPackets','LogSuccessfulConnections') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\WindowsFirewall\PrivateProfile\Logging'; Values = @('LogFilePath','LogFileSize','LogDroppedPackets','LogSuccessfulConnections') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\WindowsFirewall\PublicProfile\Logging'; Values = @('LogFilePath','LogFileSize','LogDroppedPackets','LogSuccessfulConnections') }
        # Attack Surface Reduction (GPO path)
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\ASR'; Values = @('ExploitGuard_ASR_Rules') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\ASR\Rules'; Values = $null }
        # Attack Surface Reduction (CSP/MDM - Intune writes directly here)
        @{ Path = 'HKLM\SOFTWARE\Microsoft\Windows Defender\Windows Defender Exploit Guard\ASR'; Values = @('ExploitGuard_ASR_Rules') }
        @{ Path = 'HKLM\SOFTWARE\Microsoft\Windows Defender\Windows Defender Exploit Guard\ASR\Rules'; Values = $null }
        # Internet Explorer / Edge legacy
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Internet Explorer\Main'; Values = @('DisableFirstRunCustomize') }
        # LAPS (GPO path)
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft Services\AdmPwd'; Values = @('AdmPwdEnabled') }
        # LAPS (CSP/MDM and Windows LAPS)
        @{ Path = 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\LAPS\Config'; Values = $null }
        @{ Path = 'HKLM\SOFTWARE\Microsoft\Policies\LAPS'; Values = @('BackupDirectory','PasswordAgeDays','PasswordLength','PasswordComplexity') }

        # ── Additional paths referenced by checks.json but not yet covered ──

        # Security: WDigest, SafeDllSearchMode, SEHOP, Driver Signing, Protocol Hardening
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest'; Values = @('UseLogonCredential') }
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Control\Session Manager'; Values = @('SafeDllSearchMode','ProtectionMode') }
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\kernel'; Values = @('DisableExceptionChainValidation','MitigationOptions') }
        @{ Path = 'HKLM\SOFTWARE\Microsoft\Driver Signing'; Values = @('Policy') }
        @{ Path = 'HKLM\SOFTWARE\Microsoft\PowerShell\1'; Values = @('PSEngineVersion2') }

        # Authentication: Winlogon, WHfB PIN, Biometrics
        @{ Path = 'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'; Values = @('AutoAdminLogon','CachedLogonsCount','ScRemoveOption','ScreenSaverGracePeriod') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\PassportForWork\PINComplexity'; Values = @('MaximumPINLength','MinimumPINLength') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Biometrics\FacialFeatures'; Values = @('EnhancedAntiSpoofing') }

        # Networking: DNS, mDNS, NetBridge, Shared Access, Hardened UNC, Wi-Fi, DoH
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows NT\DNSClient'; Values = @('EnableMulticast','EnableMDNS') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\Network Connections'; Values = @('NC_AllowNetBridge_NLA','NC_ShowSharedAccessUI') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\NetworkProvider'; Values = @('HardenedPaths') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\Wireless\Policy'; Values = $null }
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Services\Dnscache\Parameters'; Values = @('EnableAutoDoh') }
        @{ Path = 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections'; Values = @('WinHttpSettings') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Cryptography\Configuration\SSL\00010002'; Values = @('Functions') }

        # Data Protection: Telemetry, Advertising, Location, OneDrive, Cortana, Recall, WER, Removable Storage
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\DataCollection'; Values = @('AllowTelemetry') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo'; Values = @('DisabledByGroupPolicy') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors'; Values = @('DisableLocation') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive'; Values = @('DisableFileSyncNGSC') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search'; Values = @('AllowCortanaAboveLock') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsAI'; Values = @('DisableAIDataAnalysis') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting'; Values = @('DontSendAdditionalData') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\RemovableStorageDevices\{53f5630d-b6bf-11d0-94f2-00a0c91efb8b}'; Values = @('Deny_Write') }

        # Policy: Consumer Features, Installer, Printer, Attachments, Power/Hibernate
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\CloudContent'; Values = @('DisableWindowsConsumerFeatures') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\Installer'; Values = @('AlwaysInstallElevated') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Printers'; Values = @('RegisterSpoolerRemoteRpcEndPoint') }
        @{ Path = 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Attachments'; Values = @('SaveZoneInformation') }
        @{ Path = 'HKLM\SYSTEM\CurrentControlSet\Control\Power'; Values = @('HibernateEnabled') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Power\PowerSettings\29F6C1DB-86DA-48C5-9FDB-F2B67B1F44DA'; Values = @('ACSettingIndex') }
        @{ Path = 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update'; Values = @('RebootRequired') }

        # Edge extensions blocklist
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Edge'; Values = @('ExtensionInstallBlocklist') }

        # Defender CSP-mapped paths (not in Policies\)
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard'; Values = @('EnableVirtualizationBasedSecurity','RequirePlatformSecurityFeatures','ConfigureSystemGuardLaunch','LsaCfgFlags','HypervisorEnforcedCodeIntegrity') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Signature Updates'; Values = @('SignatureUpdateInterval') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\Controlled Folder Access'; Values = @('EnableControlledFolderAccess') }
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\Network Protection'; Values = @('EnableNetworkProtection') }
        # Corresponding CSP/MDM paths for the above
        @{ Path = 'HKLM\SOFTWARE\Microsoft\Windows Defender\Windows Defender Exploit Guard\Controlled Folder Access'; Values = @('EnableControlledFolderAccess') }
        @{ Path = 'HKLM\SOFTWARE\Microsoft\Windows Defender\Windows Defender Exploit Guard\Network Protection'; Values = @('EnableNetworkProtection') }

        # P2P Distribution
        @{ Path = 'HKLM\SOFTWARE\Policies\Microsoft\PeerDist\Service'; Values = @('HashPublicationForPeerCaching') }
    )

    $result = @{}
    $readCount = 0
    foreach ($entry in $paths) {
        if ($null -eq $entry.Values) {
            # Read all values from the key
            $allVals = Read-RegistryValues -Path $entry.Path
            foreach ($k in $allVals.Keys) {
                $result["$($entry.Path)\$k"] = $allVals[$k]
                $readCount++
            }
        } else {
            foreach ($valName in $entry.Values) {
                $val = Read-RegistryValue -Path $entry.Path -Name $valName
                $result["$($entry.Path)\$valName"] = $val
                $readCount++
            }
        }
    }
    $result['_detail'] = "$readCount keys read"
    $result
}

# ═══════════════════════════════════════════════════════════════════════
# AREA 8: DEFENDER CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════

$defenderConfig = Invoke-CollectionArea -Step 8 -Name 'Defender Configuration' -Script {
    $pref   = try { Get-MpPreference -ErrorAction Stop } catch { $null }
    $status = try { Get-MpComputerStatus -ErrorAction Stop } catch { $null }

    if (-not $pref -and -not $status) { return @{ available = $false } }

    $result = @{ available = $true }
    if ($pref) {
        $result['RealTimeProtectionEnabled']     = -not $pref.DisableRealtimeMonitoring
        $result['BehaviorMonitoringEnabled']      = -not $pref.DisableBehaviorMonitoring
        $result['IoavProtectionEnabled']          = -not $pref.DisableIOAVProtection
        $result['CloudBlockLevel']                = $pref.CloudBlockLevel
        $result['CloudExtendedTimeout']           = $pref.CloudExtendedTimeout
        $result['PUAProtection']                  = $pref.PUAProtection
        $result['SubmitSamplesConsent']            = $pref.SubmitSamplesConsent
        $result['MAPSReporting']                   = $pref.MAPSReporting
        $result['EnableNetworkProtection']         = $pref.EnableNetworkProtection
        $result['EnableControlledFolderAccess']    = $pref.EnableControlledFolderAccess
        $result['AttackSurfaceReductionRules_Ids']     = $pref.AttackSurfaceReductionRules_Ids
        $result['AttackSurfaceReductionRules_Actions'] = $pref.AttackSurfaceReductionRules_Actions
        $result['ScanScheduleDay']                = $pref.ScanScheduleDay
        $result['SignatureScheduleDay']            = $pref.SignatureScheduleDay
        $result['DisableArchiveScanning']          = $pref.DisableArchiveScanning
        $result['DisableRemovableDriveScanning']   = $pref.DisableRemovableDriveScanning
        $result['DisableEmailScanning']            = $pref.DisableEmailScanning
    }
    if ($status) {
        $result['AMServiceEnabled']                = $status.AMServiceEnabled
        $result['AntispywareEnabled']              = $status.AntispywareEnabled
        $result['AntivirusEnabled']                = $status.AntivirusEnabled
        $result['RealTimeProtectionRunning']       = $status.RealTimeProtectionEnabled
        $result['NISEnabled']                      = $status.NISEnabled
        $result['TamperProtectionSource']          = try { $status.IsTamperProtected } catch { $null }
        $result['IsTamperProtected']                = try { $status.IsTamperProtected } catch { $null }  # Alias for checks.json DEF-027
        $result['AntivirusSignatureAge']           = $status.AntivirusSignatureAge
        $result['AntivirusSignatureLastUpdated']   = try { $status.AntivirusSignatureLastUpdated.ToString('o') } catch { '' }
        $result['FullScanAge']                     = $status.FullScanAge
        $result['QuickScanAge']                    = $status.QuickScanAge
        $result['AMProductVersion']                = $status.AMProductVersion
        $result['AMEngineVersion']                 = $status.AMEngineVersion
    }
    $result
}

# ═══════════════════════════════════════════════════════════════════════
# AREA 9: FIREWALL PROFILES
# ═══════════════════════════════════════════════════════════════════════

$firewallProfiles = Invoke-CollectionArea -Step 9 -Name 'Firewall Profiles' -Script {
    $result = @{}
    try {
        $profiles = Get-NetFirewallProfile -ErrorAction Stop
        foreach ($p in $profiles) {
            $result[$p.Name] = @{
                Enabled              = [bool]($p.Enabled -eq 'True' -or $p.Enabled -eq 1 -or $p.Enabled -eq 'True')
                DefaultInboundAction = $p.DefaultInboundAction.ToString()
                DefaultOutboundAction = $p.DefaultOutboundAction.ToString()
                AllowLocalFirewallRules = $p.AllowLocalFirewallRules.ToString()
                LogAllowed           = [bool]($p.LogAllowed -eq 'True' -or $p.LogAllowed -eq 1)
                LogBlocked           = [bool]($p.LogBlocked -eq 'True' -or $p.LogBlocked -eq 1)
                LogFileName          = $p.LogFileName
                LogMaxSizeKilobytes  = $p.LogMaxSizeKilobytes
                LogMaxSizeKB         = $p.LogMaxSizeKilobytes  # Alias used by checks.json
                NotifyOnListen       = [bool]($p.NotifyOnListen -eq 'True' -or $p.NotifyOnListen -eq 1)
            }
        }
    } catch {
        $result['error'] = $_.Exception.Message
    }
    $result
}

# ═══════════════════════════════════════════════════════════════════════
# AREA 10: SERVICES
# ═══════════════════════════════════════════════════════════════════════

$services = Invoke-CollectionArea -Step 10 -Name 'Services' -Script {
    # Baseline-relevant services to check
    $targets = @(
        'XboxGipSvc', 'XblAuthManager', 'XblGameSave', 'XboxNetApiSvc',
        'WinRM', 'RemoteRegistry', 'SSDPSRV', 'upnphost', 'lltdsvc', 'IKEEXT',
        'SharedAccess', 'WMPNetworkSvc', 'RpcLocator', 'Spooler', 'LanmanServer',
        'iphlpsvc', 'WerSvc', 'WSearch', 'MpsSvc', 'WinDefend', 'SecurityHealthService',
        'wuauserv', 'BFE', 'EventLog', 'SamSs', 'Dhcp', 'Dnscache',
        'Schedule', 'gpsvc', 'CryptSvc', 'LanmanWorkstation', 'Browser',
        'W32Time', 'RpcSs', 'Winmgmt', 'BITS', 'PNRPsvc'
    )
    $result = @{}
    foreach ($svc in $targets) {
        $s = Get-Service -Name $svc -ErrorAction SilentlyContinue
        if ($s) {
            $result[$svc] = @{
                Status    = $s.Status.ToString()
                StartType = $s.StartType.ToString()
            }
        }
    }
    $result['_detail'] = "$($result.Count) services"
    $result
}

# ═══════════════════════════════════════════════════════════════════════
# AREA 11: BITLOCKER STATUS
# ═══════════════════════════════════════════════════════════════════════

$bitlocker = Invoke-CollectionArea -Step 11 -Name 'BitLocker Status' -Script {
    $result = @{}
    try {
        $volumes = Get-BitLockerVolume -ErrorAction Stop
        $osDrive = $volumes | Where-Object { $_.VolumeType -eq 'OperatingSystem' } | Select-Object -First 1
        foreach ($v in $volumes) {
            $result[$v.MountPoint] = @{
                ProtectionStatus  = $v.ProtectionStatus.ToString()
                EncryptionMethod  = $v.EncryptionMethod.ToString()
                VolumeStatus      = $v.VolumeStatus.ToString()
                EncryptionPercentage = $v.EncryptionPercentage
                KeyProtectors     = @($v.KeyProtector | ForEach-Object { $_.KeyProtectorType.ToString() })
            }
        }
        # Summary fields for auto-evaluation
        $result['osProtectionStatus']  = if ($osDrive) { $osDrive.ProtectionStatus.ToString() } else { 'Unknown' }
        $result['osVolumeStatus']      = if ($osDrive) { $osDrive.VolumeStatus.ToString() } else { 'Unknown' }
        $result['osEncryptionMethod']  = if ($osDrive) { $osDrive.EncryptionMethod.ToString() } else { 'Unknown' }
        $result['hasRecoveryKey']      = if ($osDrive) { [bool]($osDrive.KeyProtector.KeyProtectorType -contains 'RecoveryPassword') } else { $false }
        $result['allVolumesEncrypted'] = ($volumes | Where-Object { $_.VolumeStatus.ToString() -ne 'FullyEncrypted' -and $_.VolumeType -ne 'Unknown' }).Count -eq 0
    } catch {
        $result['error'] = $_.Exception.Message
    }
    $result
}

# ═══════════════════════════════════════════════════════════════════════
# AREA 12: CREDENTIAL GUARD / VBS
# ═══════════════════════════════════════════════════════════════════════

$credentialGuard = Invoke-CollectionArea -Step 12 -Name 'Credential Guard' -Script {
    $dg = try { Get-CimInstance -ClassName Win32_DeviceGuard -Namespace 'root\Microsoft\Windows\DeviceGuard' -ErrorAction Stop } catch { $null }

    # Human-readable labels for DeviceGuard enum values
    $secSvcLabels = @{ 1 = 'CredentialGuard'; 2 = 'MemoryIntegrity'; 3 = 'SystemGuard' }
    $secPropLabels = @{ 1 = 'BaseVirtualization'; 2 = 'SecureBoot'; 3 = 'DMAProtection'; 4 = 'SecureMemoryOverwrite'; 5 = 'NXProtections'; 6 = 'SMMMitigations'; 7 = 'MBECOrAPIC'; 8 = 'OSManagedSMMPageTables' }
    $vbsStatusLabels = @{ 0 = 'NotEnabled'; 1 = 'EnabledNotRunning'; 2 = 'Running' }

    @{
        VirtualizationBasedSecurityStatus    = if ($dg) { $dg.VirtualizationBasedSecurityStatus } else { $null }
        VbsStatusLabel                       = if ($dg) { $vbsStatusLabels[[int]$dg.VirtualizationBasedSecurityStatus] } else { $null }
        SecurityServicesRunning              = if ($dg) { @($dg.SecurityServicesRunning | ForEach-Object { if ($secSvcLabels.ContainsKey([int]$_)) { $secSvcLabels[[int]$_] } else { $_ } }) } else { @() }
        SecurityServicesConfigured           = if ($dg) { @($dg.SecurityServicesConfigured | ForEach-Object { if ($secSvcLabels.ContainsKey([int]$_)) { $secSvcLabels[[int]$_] } else { $_ } }) } else { @() }
        RequiredSecurityProperties           = if ($dg) { @($dg.RequiredSecurityProperties | ForEach-Object { if ($secPropLabels.ContainsKey([int]$_)) { $secPropLabels[[int]$_] } else { $_ } }) } else { @() }
        AvailableSecurityProperties          = if ($dg) { @($dg.AvailableSecurityProperties | ForEach-Object { if ($secPropLabels.ContainsKey([int]$_)) { $secPropLabels[[int]$_] } else { $_ } }) } else { @() }
        CredentialGuardConfigured            = Read-RegistryValue -Path 'HKLM\SYSTEM\CurrentControlSet\Control\Lsa' -Name 'LsaCfgFlags'
        HVCIEnabled                          = Read-RegistryValue -Path 'HKLM\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity' -Name 'Enabled'
        # Summary booleans derived from WMI (see MS docs: Win32_DeviceGuard)
        # VBS status: 0=Not enabled, 1=Enabled but not running, 2=Running
        VbsRunning                           = if ($dg -and $dg.VirtualizationBasedSecurityStatus -eq 2) { $true } else { $false }
        # SecurityServicesConfigured/Running: 1=Credential Guard, 2=Memory Integrity (HVCI), 3=System Guard
        CredentialGuardIsConfigured          = if ($dg -and $dg.SecurityServicesConfigured -contains 1) { $true } else { $false }
        CredentialGuardIsRunning             = if ($dg -and $dg.SecurityServicesRunning -contains 1) { $true } else { $false }
        MemoryIntegrityIsConfigured          = if ($dg -and $dg.SecurityServicesConfigured -contains 2) { $true } else { $false }
        MemoryIntegrityIsRunning             = if ($dg -and $dg.SecurityServicesRunning -contains 2) { $true } else { $false }
        HypervisorEnforcedCodeIntegrityEnabled = if ($dg -and $dg.SecurityServicesRunning -contains 2) { $true } else { $false }
    }
}

# ═══════════════════════════════════════════════════════════════════════
# AREA 13: WINDOWS UPDATE HISTORY
# ═══════════════════════════════════════════════════════════════════════

$windowsUpdate = Invoke-CollectionArea -Step 13 -Name 'Windows Update History' -Script {
    $hotfixes = @(Get-HotFix -ErrorAction SilentlyContinue | Sort-Object InstalledOn -Descending -ErrorAction SilentlyContinue)
    $latest   = if ($hotfixes.Count -gt 0 -and $hotfixes[0].InstalledOn) { $hotfixes[0].InstalledOn } else { $null }
    $daysSince = if ($latest) { [math]::Round(([DateTime]::Now - $latest).TotalDays, 0) } else { -1 }

    @{
        lastInstallDate = if ($latest) { $latest.ToString('yyyy-MM-dd') } else { 'Unknown' }
        daysUnpatched   = $daysSince
        hotfixCount     = $hotfixes.Count
        recentHotfixes  = @($hotfixes | Select-Object -First 10 | ForEach-Object {
            @{ HotFixID = $_.HotFixID; InstalledOn = if ($_.InstalledOn) { $_.InstalledOn.ToString('yyyy-MM-dd') } else { '' }; Description = $_.Description }
        })
    }
}

# ═══════════════════════════════════════════════════════════════════════
# AREA 14: DRIVER INVENTORY
# ═══════════════════════════════════════════════════════════════════════

$drivers = Invoke-CollectionArea -Step 14 -Name 'Driver Inventory' -Script {
    $allDrivers = Get-CimInstance Win32_PnPSignedDriver -ErrorAction SilentlyContinue |
        Where-Object { $_.DriverProviderName -and $_.DeviceName }
    $unsigned = @($allDrivers | Where-Object { $_.IsSigned -eq $false } | ForEach-Object {
        @{ DeviceName = $_.DeviceName; DriverVersion = $_.DriverVersion; Manufacturer = $_.Manufacturer }
    })
    $problematic = @(Get-CimInstance Win32_PnPEntity -ErrorAction SilentlyContinue |
        Where-Object { $_.ConfigManagerErrorCode -ne 0 -and $_.Name } | ForEach-Object {
        @{ Name = $_.Name; ErrorCode = $_.ConfigManagerErrorCode; Status = $_.Status }
    })
    @{
        totalDrivers = if ($allDrivers) { $allDrivers.Count } else { 0 }
        unsigned     = $unsigned
        problematic  = $problematic
        _detail      = "$($unsigned.Count) unsigned, $($problematic.Count) problematic"
    }
}

# ═══════════════════════════════════════════════════════════════════════
# AREA 15: STARTUP PERFORMANCE
# ═══════════════════════════════════════════════════════════════════════

$startupPerf = Invoke-CollectionArea -Step 15 -Name 'Startup Performance' -Script {
    $bootEvents = try {
        Get-WinEvent -FilterHashtable @{ LogName='Microsoft-Windows-Diagnostics-Performance/Operational'; Id=100 } -MaxEvents 5 -ErrorAction Stop |
            ForEach-Object {
                # Extract BootDuration from event XML (reliable across Windows versions)
                # The Properties[] index varies, but XML data name 'BootDuration' is consistent
                $durationMs = $null
                try {
                    $xml = [xml]$_.ToXml()
                    # BootDuration may be under EventData/Data with different casing
                    $dataNodes = $xml.Event.EventData.Data
                    $node = $dataNodes | Where-Object { $_.Name -eq 'BootDuration' }
                    # Fallback: try 'BootTime' or 'TotalBootTimeMs' (varies by Windows version)
                    if (-not $node) { $node = $dataNodes | Where-Object { $_.Name -match 'Boot(Time|Duration)' } | Select-Object -First 1 }
                    if ($node) { $durationMs = [int]$node.'#text' }
                } catch { }
                # Fallback: try Properties array (index 0 or 1 depending on event schema)
                if ($null -eq $durationMs) {
                    foreach ($idx in 0, 1, 2) {
                        try {
                            $val = [long]$_.Properties[$idx].Value
                            # Only accept if reasonable ms range (1s to 10min)
                            if ($val -gt 1000 -and $val -lt 600000) { $durationMs = [int]$val; break }
                        } catch { }
                    }
                }
                @{
                    TimeCreated    = $_.TimeCreated.ToString('o')
                    BootDurationMs = $durationMs
                }
            }
    } catch { @() }
    @{ recentBoots = @($bootEvents) }
}

# ═══════════════════════════════════════════════════════════════════════
# AREA 16: SCHEDULED TASKS
# ═══════════════════════════════════════════════════════════════════════

$scheduledTasks = Invoke-CollectionArea -Step 16 -Name 'Scheduled Tasks' -Script {
    # Single call to Get-ScheduledTask (slow COM init) — filter results in memory
    $allTasks = @(Get-ScheduledTask -ErrorAction SilentlyContinue)
    $result = @{
        totalTasks = $allTasks.Count
    }

    # High-privilege tasks running as SYSTEM (security review)
    $systemTasks = @($allTasks | Where-Object {
        $_.Principal.UserId -match 'SYSTEM|LocalSystem|S-1-5-18' -and
        $_.State -ne 'Disabled' -and
        $_.TaskPath -notmatch '\\Microsoft\\'  # Exclude built-in MS tasks
    })
    $result['highPrivilegeTasks'] = @($systemTasks | ForEach-Object {
        @{ Name = $_.TaskName; Path = $_.TaskPath; State = $_.State.ToString(); RunAs = $_.Principal.UserId }
    })

    # Failed/errored tasks (last result != 0 and != 0x41301 running)
    $failedTasks = @($allTasks | Where-Object {
        $lr = $_.LastTaskResult
        $_.State -ne 'Disabled' -and $null -ne $lr -and $lr -ne 0 -and $lr -ne 267009
    } | Select-Object -First 20)
    $result['failedTasks'] = @($failedTasks | ForEach-Object {
        $lr = $_.LastTaskResult
        $hex = if ($null -ne $lr -and $lr -is [int]) { '0x{0:X}' -f $lr } elseif ($null -ne $lr) { "0x$lr" } else { 'unknown' }
        @{ Name = $_.TaskName; Path = $_.TaskPath; LastResult = $hex; State = $_.State.ToString() }
    })

    # Known indicator tasks (Xbox, telemetry, etc.)
    $indicatorNames = @('XblGameSaveTask', 'Consolidator', 'UsbCeip', 'DmClient', 'DmClientOnScenarioDownload', 'MapsToastTask', 'MapsUpdateTask')
    foreach ($t in $indicatorNames) {
        $task = $allTasks | Where-Object { $_.TaskName -eq $t } | Select-Object -First 1
        if ($task) {
            $result[$t] = @{ State = $task.State.ToString(); Enabled = ($task.State -ne 'Disabled') }
        }
    }

    $result['_detail'] = "$($allTasks.Count) total, $($systemTasks.Count) high-priv, $($failedTasks.Count) failed"
    $result
}

# ═══════════════════════════════════════════════════════════════════════
# AREA 17: SMB CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════

$smbConfig = Invoke-CollectionArea -Step 17 -Name 'SMB Configuration' -Script {
    $server = try { Get-SmbServerConfiguration -ErrorAction Stop } catch { $null }
    $client = try { Get-SmbClientConfiguration -ErrorAction Stop } catch { $null }
    @{
        server = if ($server) {
            @{
                SMB1Protocol             = $server.EnableSMB1Protocol
                SMB2Protocol             = $server.EnableSMB2Protocol
                RequireSecuritySignature = $server.RequireSecuritySignature
                EnableSecuritySignature  = $server.EnableSecuritySignature
                EncryptData              = $server.EncryptData
                RejectUnencryptedAccess  = $server.RejectUnencryptedAccess
            }
        } else { @{ error = 'Not available' } }
        client = if ($client) {
            @{
                RequireSecuritySignature = $client.RequireSecuritySignature
                EnableSecuritySignature  = $client.EnableSecuritySignature
                EncryptionCiphers        = $client.EncryptionCiphers
            }
        } else { @{ error = 'Not available' } }
    }
}

# ═══════════════════════════════════════════════════════════════════════
# AREA 18: TLS CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════

$tlsConfig = Invoke-CollectionArea -Step 18 -Name 'TLS Configuration' -Script {
    $protocols = @('SSL 2.0', 'SSL 3.0', 'TLS 1.0', 'TLS 1.1', 'TLS 1.2', 'TLS 1.3')
    $result = @{}
    foreach ($proto in $protocols) {
        foreach ($side in @('Client', 'Server')) {
            $path = "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$proto\$side"
            $enabled  = Read-RegistryValue -Path $path -Name 'Enabled'
            $disabled = Read-RegistryValue -Path $path -Name 'DisabledByDefault'
            # Flat keys used by checks.json (e.g. "TLS 1.2.Server.Enabled")
            $result["$proto.$side.Enabled"]           = $enabled
            $result["$proto.$side.DisabledByDefault"] = $disabled
        }
    }
    $result
}

# ═══════════════════════════════════════════════════════════════════════
# AREA 19: POWERSHELL CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════

$powershellConfig = Invoke-CollectionArea -Step 19 -Name 'PowerShell Configuration' -Script {
    @{
        ScriptBlockLogging          = Read-RegistryValue -Path 'HKLM\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging' -Name 'EnableScriptBlockLogging'
        ScriptBlockInvocationLogging = Read-RegistryValue -Path 'HKLM\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging' -Name 'EnableScriptBlockInvocationLogging'
        Transcription               = Read-RegistryValue -Path 'HKLM\SOFTWARE\Policies\Microsoft\Windows\PowerShell\Transcription' -Name 'EnableTranscripting'
        TranscriptionPath           = Read-RegistryValue -Path 'HKLM\SOFTWARE\Policies\Microsoft\Windows\PowerShell\Transcription' -Name 'OutputDirectory'
        ModuleLogging               = Read-RegistryValue -Path 'HKLM\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ModuleLogging' -Name 'EnableModuleLogging'
        ConstrainedLanguageMode     = $ExecutionContext.SessionState.LanguageMode.ToString()
        ExecutionPolicy             = try { (Get-ExecutionPolicy -Scope LocalMachine).ToString() } catch { 'Unknown' }
    }
}

# ═══════════════════════════════════════════════════════════════════════
# AREA 20: WINRM CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════

$winrmConfig = Invoke-CollectionArea -Step 20 -Name 'WinRM Configuration' -Script {
    $svc = Get-Service WinRM -ErrorAction SilentlyContinue
    $result = @{
        ServiceStatus = if ($svc) { $svc.Status.ToString() } else { 'NotInstalled' }
        ServiceStartType = if ($svc) { $svc.StartType.ToString() } else { 'N/A' }
    }
    if ($svc -and $svc.Status -eq 'Running') {
        try {
            $raw = winrm get winrm/config 2>&1 | Out-String
            $result['AllowBasicClient']       = Read-RegistryValue -Path 'HKLM\SOFTWARE\Policies\Microsoft\Windows\WinRM\Client' -Name 'AllowBasic'
            $result['AllowUnencryptedClient']  = Read-RegistryValue -Path 'HKLM\SOFTWARE\Policies\Microsoft\Windows\WinRM\Client' -Name 'AllowUnencryptedTraffic'
            $result['AllowDigestClient']       = Read-RegistryValue -Path 'HKLM\SOFTWARE\Policies\Microsoft\Windows\WinRM\Client' -Name 'AllowDigest'
            $result['AllowBasicService']       = Read-RegistryValue -Path 'HKLM\SOFTWARE\Policies\Microsoft\Windows\WinRM\Service' -Name 'AllowBasic'
            $result['AllowUnencryptedService'] = Read-RegistryValue -Path 'HKLM\SOFTWARE\Policies\Microsoft\Windows\WinRM\Service' -Name 'AllowUnencryptedTraffic'
        } catch { }
    }
    $result
}

# ═══════════════════════════════════════════════════════════════════════
# AREA 21: EVENT LOG METADATA
# ═══════════════════════════════════════════════════════════════════════

$eventLogMetadata = Invoke-CollectionArea -Step 21 -Name 'Event Log Metadata' -Script {
    $logNames = @('Security', 'System', 'Application',
                  'Microsoft-Windows-PowerShell/Operational',
                  'Microsoft-Windows-Sysmon/Operational',
                  'Microsoft-Windows-Windows Defender/Operational',
                  'Microsoft-Windows-AppLocker/EXE and DLL',
                  'Microsoft-Windows-CodeIntegrity/Operational')
    $result = @{}
    foreach ($log in $logNames) {
        try {
            $info = Get-WinEvent -ListLog $log -ErrorAction Stop
            $result[$log] = @{
                MaxSizeKB      = [math]::Round($info.MaximumSizeInBytes / 1KB, 0)
                RecordCount    = $info.RecordCount
                IsEnabled      = $info.IsEnabled
                OverflowAction = $info.LogMode.ToString()
                FileSize       = if ($info.FileSize) { [math]::Round($info.FileSize / 1KB, 0) } else { 0 }
            }
        } catch {
            # Normalize error: distinguish "log not found" from real errors
            $errMsg = if ($_.Exception -is [System.Diagnostics.Eventing.Reader.EventLogNotFoundException]) {
                'Log not found'
            } else {
                $_.Exception.Message
            }
            $result[$log] = @{ IsEnabled = $false; Error = $errMsg }
        }
    }
    $result['_detail'] = "$($result.Count) logs"
    $result
}

# ═══════════════════════════════════════════════════════════════════════
# AREA 22: SECURITY EVENT COLLECTION
# ═══════════════════════════════════════════════════════════════════════

if (-not $SkipEventCollection) {
    $eventData = Invoke-CollectionArea -Step 22 -Name 'Security Event Collection' -Script {
        $cutoff  = (Get-Date).AddDays(-$LookbackDays)
        $result  = @{ _queryMeta = @{
            lookbackDays          = $LookbackDays
            queryTimestamp        = (Get-Date).ToString('o')
            totalEventsCollected  = 0
            queryDurationSec      = 0
        }}
        $totalEvents = 0
        $querySw     = [System.Diagnostics.Stopwatch]::StartNew()

        # Event query definitions
        $queries = @(
            @{ Name = 'Logon Events (4624/4625)';             Key = 'logonEvents';        Log = 'Security';    Ids = @(4624,4625);                              Props = @('TargetUserName','TargetDomainName','LogonType','IpAddress','WorkstationName','FailureReason','SubStatus') }
            @{ Name = 'Account Lockout (4740)';               Key = 'accountLockout';     Log = 'Security';    Ids = @(4740);                                   Props = @('TargetUserName','TargetDomainName','CallerComputerName') }
            @{ Name = 'Account Management (4720-4735)';       Key = 'accountManagement';  Log = 'Security';    Ids = @(4720,4722,4723,4724,4725,4726,4727,4728,4729,4730,4731,4732,4733,4734,4735); Props = @('TargetUserName','SubjectUserName','GroupName','MemberName') }
            @{ Name = 'Privilege Use (4672/4673)';            Key = 'privilegeUse';       Log = 'Security';    Ids = @(4672,4673);                              Props = @('SubjectUserName','SubjectDomainName','PrivilegeList') }
            @{ Name = 'Process Creation (4688)';              Key = 'processCreation';    Log = 'Security';    Ids = @(4688);                                   Props = @('NewProcessName','ParentProcessName','SubjectUserName','CommandLine') }
            @{ Name = 'Policy Change (4719/4739)';            Key = 'policyChange';       Log = 'Security';    Ids = @(4719,4739);                              Props = @('SubjectUserName','CategoryId','SubcategoryGuid') }
            @{ Name = 'System Integrity (4612/4615/4616)';    Key = 'systemIntegrity';    Log = 'Security';    Ids = @(4612,4615,4616);                         Props = @('SubjectUserName') }
            @{ Name = 'Audit Policy Change (4902/4906/4907)'; Key = 'auditPolicyChange';  Log = 'Security';    Ids = @(4902,4904,4905,4906,4907);               Props = @('SubjectUserName','SubcategoryGuid','AuditPolicyChanges') }
            @{ Name = 'Kerberos (4768/4769/4771)';           Key = 'kerberosAuth';       Log = 'Security';    Ids = @(4768,4769,4771);                         Props = @('TargetUserName','ServiceName','IpAddress','TicketOptions','Status') }
            @{ Name = 'Object Access (4663/4656)';           Key = 'objectAccess';       Log = 'Security';    Ids = @(4663,4656,4658);                         Props = @('SubjectUserName','ObjectName','ObjectType','AccessMask','ProcessName') }
            @{ Name = 'Application Crashes (1000/1001)';     Key = 'applicationCrashes'; Log = 'Application'; Ids = @(1000,1001);                              Props = @() }
            @{ Name = 'System Errors (41/6008/6013)';        Key = 'systemErrors';       Log = 'System';      Ids = @(41,6008,6013);                           Props = @() }
            @{ Name = 'WHEA/Hardware (17/18/19/20)';         Key = 'hardwareErrors';     Log = 'System';      Ids = @(17,18,19,20);                            Props = @() }
        )

        foreach ($q in $queries) {
            $qSw = [System.Diagnostics.Stopwatch]::StartNew()
            try {
                $filter = @{
                    LogName   = $q.Log
                    Id        = $q.Ids
                    StartTime = $cutoff
                }
                $events = @(Get-WinEvent -FilterHashtable $filter -MaxEvents $MaxEventsPerQuery -ErrorAction Stop)
                $count  = $events.Count

                if ($EventSummaryOnly) {
                    # Summary mode: counts + top users only
                    $summary = @{
                        count      = $count
                        firstEvent = if ($count -gt 0) { $events[-1].TimeCreated.ToString('o') } else { $null }
                        lastEvent  = if ($count -gt 0) { $events[0].TimeCreated.ToString('o') } else { $null }
                    }
                    # Top users for security events
                    if ($q.Log -eq 'Security' -and $count -gt 0 -and $q.Props -contains 'TargetUserName') {
                        $topUsers = $events | ForEach-Object {
                            try { $_.Properties[5].Value } catch { '' }
                        } | Where-Object { $_ } | Group-Object | Sort-Object Count -Descending |
                            Select-Object -First 5 | ForEach-Object { @{ user = $_.Name; count = $_.Count } }
                        $summary['topUsers'] = @($topUsers)
                    }
                    $result[$q.Key] = $summary
                } else {
                    # Full mode: extract normalized events
                    $extracted = [System.Collections.ArrayList]::new()
                    foreach ($evt in $events) {
                        $entry = @{
                            id   = $evt.Id
                            time = $evt.TimeCreated.ToString('o')
                        }
                        if ($q.Props.Count -gt 0) {
                            $props = @{}
                            foreach ($propName in $q.Props) {
                                try {
                                    # Use XPath property names mapped to indices
                                    $xml = [xml]$evt.ToXml()
                                    $node = $xml.Event.EventData.Data | Where-Object { $_.Name -eq $propName }
                                    if ($node) { $props[$propName] = $node.'#text' }
                                } catch { }
                            }
                            if ($props.Count -gt 0) { $entry['props'] = $props }
                        } else {
                            $entry['message'] = $evt.Message -replace '\r?\n',' ' | ForEach-Object { $_.Substring(0, [math]::Min($_.Length, 200)) }
                        }
                        [void]$extracted.Add($entry)
                    }
                    $result[$q.Key] = @($extracted)
                }
                $totalEvents += $count
            } catch [Exception] {
                # Match "no events found" errors across all locales (EN/DE/FR/ES/etc.)
                $isNoEvents = $_.Exception -is [System.Diagnostics.Eventing.Reader.EventLogNotFoundException] -or
                              $_.Exception.HResult -eq -2146233088 -or    # 0x80131500 - No events matching criteria
                              $_.Exception.Message -match 'No events were found|Es wurden keine Ereignisse|Aucun .v.nement|No se encontraron eventos'
                if ($isNoEvents) {
                    $count = 0
                    $result[$q.Key] = if ($EventSummaryOnly) { @{ count = 0 } } else { @() }
                } else {
                    $result[$q.Key] = @{ error = $_.Exception.Message }
                    $count = 0
                }
            }
            $qSw.Stop()
            $evtLabel = if ($count -eq 1) { 'event' } else { 'events' }
            Write-CollectorProgress -Step 22 -Total $Script:TotalAreas -Name $q.Name -Status 'info' `
                -Detail ("$count $evtLabel (" + [math]::Round($qSw.Elapsed.TotalSeconds, 1) + "s)")
        }

        $querySw.Stop()
        $result['_queryMeta']['totalEventsCollected'] = $totalEvents
        $result['_queryMeta']['queryDurationSec']     = [math]::Round($querySw.Elapsed.TotalSeconds, 1)
        $result['_detail'] = "$totalEvents total events across $($queries.Count) queries"

        # Print final event summary line (done is handled by Invoke-CollectionArea)
        $result
    }
} else {
    $eventData = @{ skipped = $true; reason = 'SkipEventCollection switch specified' }
}

# ═══════════════════════════════════════════════════════════════════════
# ASSEMBLE OUTPUT
# ═══════════════════════════════════════════════════════════════════════

$endTime = [DateTime]::Now
$durationMs = [math]::Round(($endTime - $Script:StartTime).TotalMilliseconds, 0)

# Clean _detail keys from results before JSON export
<#
.SYNOPSIS
    Removes the internal '_detail' key from a collection result hashtable.
.DESCRIPTION
    Each collection area stores a '_detail' key with a human-readable summary
    for progress output. This function strips it before JSON serialization.
.PARAMETER obj
    The hashtable to clean. No-op if not a hashtable.
#>
function Remove-DetailKeys {
    param($obj)
    if ($obj -is [hashtable]) {
        $obj.Remove('_detail')
    }
}
Remove-DetailKeys $systemInfo
Remove-DetailKeys $mdmEnrollment
Remove-DetailKeys $securityPolicy
Remove-DetailKeys $auditPolicy
Remove-DetailKeys $registryBaselines
Remove-DetailKeys $services
Remove-DetailKeys $scheduledTasks
Remove-DetailKeys $eventLogMetadata
Remove-DetailKeys $drivers
if ($eventData -is [hashtable]) { Remove-DetailKeys $eventData }
if ($appliedGPOs -is [hashtable]) { Remove-DetailKeys $appliedGPOs }

$output = [ordered]@{
    _metadata = [ordered]@{
        collectorVersion   = $Script:CollectorVersion
        timestamp          = $Script:StartTime.ToString('o')
        hostname           = $hostname
        collectionDurationMs = $durationMs
        areasCompleted     = $Script:TotalAreas - $Script:Errors.Count
        areasTotal         = $Script:TotalAreas
        errors             = @($Script:Errors)
        parameters         = [ordered]@{
            lookbackDays       = $LookbackDays
            maxEventsPerQuery  = $MaxEventsPerQuery
            skipEventCollection = [bool]$SkipEventCollection
            eventSummaryOnly   = [bool]$EventSummaryOnly
        }
    }
    systemInfo        = $systemInfo
    joinType          = $joinType
    appliedGPOs       = if ($appliedGPOs) { $appliedGPOs.appliedGPOs } else { @() }
    deniedGPOs        = if ($appliedGPOs -and $appliedGPOs.deniedGPOs) { $appliedGPOs.deniedGPOs } else { @() }
    mdmEnrollment     = $mdmEnrollment
    securityPolicy    = $securityPolicy
    auditPolicy       = $auditPolicy
    registryBaselines = $registryBaselines
    defender          = $defenderConfig
    firewall          = $firewallProfiles
    services          = $services
    bitlocker         = $bitlocker
    credentialGuard   = $credentialGuard
    windowsUpdate     = $windowsUpdate
    drivers           = $drivers
    startupPerf       = $startupPerf
    scheduledTasks    = $scheduledTasks
    smbConfig         = $smbConfig
    tlsConfig         = $tlsConfig
    powershellConfig  = $powershellConfig
    winrmConfig       = $winrmConfig
    eventLogMetadata  = $eventLogMetadata
    eventData         = $eventData
}

# ═══════════════════════════════════════════════════════════════════════
# WRITE OUTPUT
# ═══════════════════════════════════════════════════════════════════════

if (-not $OutputPath) {
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path $PWD.Path "${hostname}_baseline_${ts}.json"
}

$jsonStr  = $output | ConvertTo-Json -Depth 10
$sizeMB   = [math]::Round($jsonStr.Length / 1MB, 1)
[System.IO.File]::WriteAllText($OutputPath, $jsonStr, [System.Text.Encoding]::UTF8)

# ═══════════════════════════════════════════════════════════════════════
# SUMMARY BANNER
# ═══════════════════════════════════════════════════════════════════════

if (-not $Quiet) {
    $durationSec  = [math]::Round($durationMs / 1000, 1)
    $fileName     = Split-Path $OutputPath -Leaf
    $successAreas = $Script:TotalAreas - $Script:Errors.Count
    $totalEvents  = if (-not $SkipEventCollection -and $eventData._queryMeta) { $eventData._queryMeta.totalEventsCollected } else { 0 }

    Write-Host ''
    Write-Host ''
    Write-BoxTop
    # Title centered
    $title = 'Collection Complete'
    $titlePad = [math]::Floor(($Script:BW - $title.Length) / 2)
    $titleLine = "$(' ' * $titlePad)$title$(' ' * ($Script:BW - $titlePad - $title.Length))"
    Write-Host "  $([char]0x2551)" -NoNewline -ForegroundColor Green
    Write-Host $titleLine -NoNewline -ForegroundColor Green
    Write-Host "$([char]0x2551)" -ForegroundColor Green
    Write-BoxMid

    # Collection stats
    Write-BoxLine ' ' 'White'
    Write-BoxLine '  COLLECTION SUMMARY' 'White'
    Write-BoxLine ' ' 'White'
    Write-BoxKV 'Duration' "${durationSec}s"
    Write-BoxKV 'Output File' $fileName 'Gray' 'Cyan'
    Write-BoxKV 'File Size' "${sizeMB} MB"
    Write-BoxKV 'Areas Collected' "$successAreas / $Script:TotalAreas"
    if (-not $SkipEventCollection) {
        Write-BoxKV 'Events Collected' "$totalEvents"
    } else {
        Write-BoxKV 'Events' 'Skipped' 'Gray' 'Yellow'
    }
    Write-BoxLine ' ' 'White'

    # Score bar for area success rate
    Write-BoxMid
    Write-BoxLine ' ' 'White'
    $pct = if ($Script:TotalAreas -gt 0) { [math]::Round(($successAreas / $Script:TotalAreas) * 100) } else { 0 }
    $scoreStr = "$pct% success"
    $scoreLine = "  AREAS  $(' ' * ($Script:BW - 10 - $scoreStr.Length))$scoreStr"
    $scoreColor = if ($pct -ge 90) { 'Green' } elseif ($pct -ge 70) { 'Yellow' } else { 'Red' }
    Write-BoxLine $scoreLine $scoreColor

    # Progress bar
    $barLen  = $Script:BW - 8
    $filled  = [math]::Round($barLen * $pct / 100)
    $empty   = $barLen - $filled
    $barStr  = "$([string]([char]0x2588) * $filled)$([string]([char]0x2591) * $empty)"
    Write-BoxLine "    $barStr" $scoreColor
    Write-BoxLine ' ' 'White'

    # Status counts
    $succLine = "    $([char]0x2713) Successful"
    $errLine  = "    $([char]0x2717) Failed"
    Write-BoxKV "$([char]0x2713) Successful" "$successAreas" 'Green' 'Green'
    if ($Script:Errors.Count -gt 0) {
        Write-BoxKV "$([char]0x2717) Failed" "$($Script:Errors.Count)" 'Red' 'Red'
    }
    Write-BoxLine ' ' 'White'

    # Data collected breakdown
    Write-BoxMid
    Write-BoxLine ' ' 'White'
    Write-BoxLine '  DATA COLLECTED' 'White'
    Write-BoxLine ' ' 'White'
    $dataItems = @(
        @{ L = 'Registry Keys';    V = if ($registryBaselines) { "$($registryBaselines.Count)" } else { '0' } }
        @{ L = 'Security Policies'; V = if ($securityPolicy) { "$($securityPolicy.Count)" } else { '0' } }
        @{ L = 'Audit Subcategories'; V = if ($auditPolicy) { "$($auditPolicy.Count)" } else { '0' } }
        @{ L = 'Services Checked';  V = if ($services) { "$($services.Count)" } else { '0' } }
        @{ L = 'Drivers Scanned';   V = if ($drivers) { "$($drivers.total)" } else { '0' } }
    )
    foreach ($d in $dataItems) {
        Write-BoxKV $d.L $d.V
    }
    Write-BoxLine ' ' 'White'
    Write-BoxBottom
    Write-Host ''

    # Errors section
    if ($Script:Errors.Count -gt 0) {
        Write-Status "Errors ($($Script:Errors.Count)):" -Level 'WARN'
        foreach ($err in $Script:Errors) {
            Write-Status "  $($err.Area): $($err.Error)" -Level 'ERROR'
        }
        Write-Host ''
    }

    Write-Status "Output: $OutputPath" -Level 'SUCCESS'
    Write-Host ''
}

Write-Output $OutputPath
