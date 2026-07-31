#Requires -Version 5.1
<#
.SYNOPSIS
    Reports the effective Microsoft Defender Antivirus / Defender for Endpoint state and
    detects third-party antivirus and EDR products sharing the same endpoint.

.DESCRIPTION
    Read-only. No changes are made to the system.

    Answers the question "what security software is *actually* running on this machine, and
    is more than one product doing the same job?" - using the products' own effective state
    rather than an Intune or GPO policy export. Where the device disagrees with the exported
    policy, the device is authoritative: an export shows what was *intended*, this shows what
    is *running*.

    Collected in one pass:
      1.  Machine identity, OS build, elevation state
      2.  Security services (Microsoft + a known third-party AV/EDR service list)
      3.  Security processes currently running
      4.  Defender for Endpoint onboarding state (HKLM\...\Advanced Threat Protection\Status)
      5.  Defender for Endpoint policy key, including ForceDefenderPassiveMode
      6.  Defender Antivirus effective state (Get-MpComputerStatus)
      7.  Defender Antivirus preferences relevant to performance and cloud protection
      8.  Attack Surface Reduction rules actually applied
      9.  Effective exclusions, plus an automated hygiene review (see below)
      10. Loaded file-system minifilters, classified by altitude band and mapped to products
      11. Products registered with Windows Security Center
      12. Effective EDR configuration read back from the SENSE operational log
      13. SENSE cloud connectivity and recent sensor errors
      14. Installed Defender platform versions
      15. A coexistence verdict

    Exclusion hygiene review flags the following classes of broken or over-broad rule:
      * Environment variables in exclusion paths. Defender Antivirus runs as LocalSystem, so
        %USERPROFILE%, %APPDATA%, %LOCALAPPDATA%, %TEMP% and %TMP% resolve to the *system*
        profile, not the signed-in user. Rules written this way silently match nothing -
        a common cause of Outlook OST/PST files being scanned despite an apparent exclusion.
        See the Microsoft reference in .NOTES.
      * Bare filenames in the path exclusion list (a path exclusion requires a path).
      * Driver files (.sys) in the process exclusion list (a driver never runs as a process).
      * Non-executable files in the process exclusion list.
      * Auto-start locations (Startup / Start Menu\Programs) excluded from scanning.
      * Bare-name process exclusions, which match any file of that name from any location.
      * Very broad roots (drive root or single top-level folder).
      * Executable or script extensions (exe, dll, ps1, js, ...) excluded wholesale.
      * Paths placed in the extension exclusion list.

.PARAMETER AsObject
    Emit the structured result object instead of the formatted console report. Use this when
    you want to post-process, compare two machines, or pipe into Export-Csv.

.PARAMETER JsonPath
    Optional. Write the full structured result to this path as JSON (UTF-8). The console
    report is still produced unless -Quiet is also supplied.

.PARAMETER Quiet
    Suppress the human-readable console report. A single-line JSON summary is still written
    to STDOUT so Intune script-output harvesting and log scraping keep working.

.PARAMETER MaxEvents
    How many records to read from the SENSE operational log when looking for configuration,
    connectivity and error events. Default 400. Raise it on a busy or long-uptime machine
    if the configuration events are not found.

.PARAMETER SkipEventLog
    Skip all SENSE operational log queries. Use on machines where the log is very large or
    access is slow; sections 12 and 13 are then reported as not collected.

.EXAMPLE
    .\Get-MdeCoexistenceState.ps1
    Full console report. Run elevated for the minifilter section.

.EXAMPLE
    .\Get-MdeCoexistenceState.ps1 -JsonPath C:\Temp\mde-state.json
    Console report plus a structured JSON artifact suitable for diffing two machines.

.EXAMPLE
    $a = .\Get-MdeCoexistenceState.ps1 -AsObject
    Capture the result object for comparison, e.g. $a.Exclusions.Findings.

.EXAMPLE
    Invoke-Command -ComputerName (Get-Content .\hosts.txt) -FilePath .\Get-MdeCoexistenceState.ps1
    Fan out across a fleet and collect the summaries centrally.

.NOTES
    File:     windows/diagnostics/MdeCoexistenceState/Get-MdeCoexistenceState.ps1
    Author:   Anton Romanyuk
    Version:  1.1.0
    Requires: PowerShell 5.1+. Elevation is required only for the minifilter section
              (fltmc.exe); every other section works unelevated.

    Exit codes:
      0 - OK        single security stack, no hygiene findings
      1 - WARNING   coexistence detected, or exclusion/configuration findings present
      2 - CRITICAL  sensor onboarded but never logged a successful server contact (event 4),
                    or Tamper Protection is off while Defender AV is active (Normal mode)
      3 - PARTIAL   ran unelevated, or Defender cmdlets unavailable; result is incomplete
      4 - ERROR     unexpected failure caught at top level

    Confidence and scope caveats - read before quoting this in a report:
      * Vendor attribution from minifilter and process names is an INFERENCE drawn from a
        static lookup table in this script. A driver named 'CSAgent' is almost certainly
        CrowdStrike Falcon, but this script does not verify the signing certificate. Treat
        the Product column as a strong hint, not proof. Filters not in the table are
        reported with their raw name and altitude band so nothing is silently dropped.
      * Altitude band names come from Microsoft's published Filter Manager allocated
        altitude ranges. The band tells you what a filter *claims* to be, not what it does.
      * Microsoft in-box filters (UCPD, bfs, FileInfo, Wof and similar) are deliberately
        EXCLUDED from the security-product count. Counting them inflates the apparent
        security stack and is the first thing a vendor will pick apart.
      * An exclusion does NOT remove a filter driver from the I/O path. Exclusions govern
        what a product does after its filter is called; the filter is still attached and
        still invoked on every file operation. Do not use this script to argue that
        exclusion tuning will remove filter overhead - it will not.
      * ForceDefenderPassiveMode is reported, not judged. Defender running in Normal mode
        alongside a third-party EDR is a legitimate design (third-party EDR + Defender AV).
        Whether it is correct depends on the intended architecture, which this script
        cannot know.

    Reference for the LocalSystem environment-variable behaviour:
      https://learn.microsoft.com/en-us/defender-endpoint/microsoft-defender-antivirus-exclusions-overview

.DISCLAIMER
    THIS SCRIPT IS PROVIDED "AS-IS" WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
    INCLUDING BUT NOT LIMITED TO MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
    NONINFRINGEMENT. Use at your own risk and validate in a test environment first.
#>

[CmdletBinding()]
param(
    [switch] $AsObject,

    [string] $JsonPath,

    [switch] $Quiet,

    [ValidateRange(50, 5000)]
    [int]    $MaxEvents = 400,

    [switch] $SkipEventLog
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Lookup tables
# ---------------------------------------------------------------------------

# Minifilter name (lowercase) -> product attribution.
# Security = $true  : counts toward the security filter stack
# Security = $false : Microsoft in-box component, reported but NOT counted
$Script:KnownFilters = @{
    # --- Microsoft security ---
    'wdfilter'          = @{ Product = 'Microsoft Defender Antivirus';            Security = $true  }
    'mssecflt'          = @{ Product = 'Microsoft Defender for Endpoint EDR';     Security = $true  }
    'wdnisdrv'          = @{ Product = 'Microsoft Defender Network Inspection';   Security = $true  }
    'mpksldrv'          = @{ Product = 'Microsoft Defender (kernel support)';     Security = $true  }
    # --- Microsoft in-box, NOT security products ---
    'ucpd'              = @{ Product = 'Windows User Choice Protection Driver';   Security = $false }
    'bfs'               = @{ Product = 'Windows Brokering File System';           Security = $false }
    'fileinfo'          = @{ Product = 'Windows File Information (SuperFetch)';   Security = $false }
    'wof'               = @{ Product = 'Windows Overlay File System';             Security = $false }
    'wcifs'             = @{ Product = 'Windows Container Isolation';             Security = $false }
    'cldflt'            = @{ Product = 'Windows Cloud Files (OneDrive)';          Security = $false }
    'bindflt'           = @{ Product = 'Windows Bind Filter';                     Security = $false }
    'filecrypt'         = @{ Product = 'Windows FileCrypt';                       Security = $false }
    'luafv'             = @{ Product = 'Windows LUA File Virtualization';         Security = $false }
    'npsvctrig'         = @{ Product = 'Named Pipe Service Trigger';              Security = $false }
    'storqosflt'        = @{ Product = 'Storage QoS Filter';                      Security = $false }
    'iorate'            = @{ Product = 'Windows I/O Rate Control';                Security = $false }
    'applockerfltr'     = @{ Product = 'AppLocker Filter';                        Security = $false }
    'bfltr'             = @{ Product = 'Windows Bitlocker helper';                Security = $false }
    # --- Third-party antivirus / EDR ---
    'csagent'           = @{ Product = 'CrowdStrike Falcon';                      Security = $true  }
    'sentinelmonitor'   = @{ Product = 'SentinelOne';                             Security = $true  }
    'carbonblackk'      = @{ Product = 'VMware Carbon Black';                     Security = $true  }
    'parity'            = @{ Product = 'VMware Carbon Black App Control';         Security = $true  }
    'cylancedrv'        = @{ Product = 'BlackBerry Cylance';                      Security = $true  }
    'cyoptics'          = @{ Product = 'BlackBerry CylanceOPTICS';                Security = $true  }
    'eamonm'            = @{ Product = 'ESET';                                    Security = $true  }
    'ehdrv'             = @{ Product = 'ESET';                                    Security = $true  }
    'klif'              = @{ Product = 'Kaspersky';                               Security = $true  }
    'klam'              = @{ Product = 'Kaspersky';                               Security = $true  }
    'symefa'            = @{ Product = 'Symantec / Broadcom Endpoint Security';   Security = $true  }
    'srtsp'             = @{ Product = 'Symantec / Broadcom AutoProtect';         Security = $true  }
    'mfehidk'           = @{ Product = 'McAfee / Trellix';                        Security = $true  }
    'mfeaskm'           = @{ Product = 'McAfee / Trellix';                        Security = $true  }
    'hdlpflt'           = @{ Product = 'McAfee / Trellix host DLP';               Security = $true  }
    'tmxpflt'           = @{ Product = 'Trend Micro';                             Security = $true  }
    'tmprefilter'       = @{ Product = 'Trend Micro';                             Security = $true  }
    'sophosed'          = @{ Product = 'Sophos Endpoint Defense';                 Security = $true  }
    'savonaccess'       = @{ Product = 'Sophos On-Access';                        Security = $true  }
    'cybkerneltracker'  = @{ Product = 'Cybereason';                              Security = $true  }
    'psinproc'          = @{ Product = 'Panda Security';                          Security = $true  }
    'psinfile'          = @{ Product = 'Panda Security';                          Security = $true  }
    'gzflt'             = @{ Product = 'Bitdefender';                             Security = $true  }
    'bddevflt'          = @{ Product = 'Bitdefender';                             Security = $true  }
    'edrsensor'         = @{ Product = 'BlackBerry / Cylance EDR';                Security = $true  }
    'fekern'            = @{ Product = 'FireEye / Trellix Endpoint';              Security = $true  }
    'esensor'           = @{ Product = 'Endgame / Elastic';                       Security = $true  }
    'nxtrdrv'           = @{ Product = 'Nexthink (endpoint analytics)';           Security = $true  }
    'vfdrv'             = @{ Product = 'CyberArk Endpoint Privilege Manager';     Security = $true  }
    'stadrv6x64'        = @{ Product = 'Netskope';                                Security = $true  }
    'stadrv'            = @{ Product = 'Netskope';                                Security = $true  }
    'dlpflt'            = @{ Product = 'Generic host DLP';                        Security = $true  }
    'zscalerfilt'       = @{ Product = 'Zscaler';                                 Security = $true  }
}

# Microsoft Filter Manager allocated altitude ranges (load-order groups).
$Script:AltitudeBands = @(
    @{ Min = 420000; Max = 429999; Name = 'Filter'            }
    @{ Min = 400000; Max = 409999; Name = 'Top'               }
    @{ Min = 360000; Max = 389999; Name = 'Activity Monitor'  }
    @{ Min = 340000; Max = 349999; Name = 'Undelete'          }
    @{ Min = 320000; Max = 329999; Name = 'Anti-Virus'        }
    @{ Min = 300000; Max = 309999; Name = 'Replication'       }
    @{ Min = 280000; Max = 289999; Name = 'Continuous Backup' }
    @{ Min = 260000; Max = 269999; Name = 'Content Screener'  }
    @{ Min = 240000; Max = 249999; Name = 'Quota Management'  }
    @{ Min = 220000; Max = 229999; Name = 'System Recovery'   }
    @{ Min = 200000; Max = 209999; Name = 'Cluster FS'        }
    @{ Min = 180000; Max = 189999; Name = 'HSM'               }
    @{ Min = 170000; Max = 175000; Name = 'Imaging'           }
    @{ Min = 160000; Max = 169999; Name = 'Compression'       }
    @{ Min = 140000; Max = 149999; Name = 'Encryption'        }
    @{ Min = 130000; Max = 139999; Name = 'Virtualization'    }
    @{ Min = 120000; Max = 129999; Name = 'Physical Quota'    }
    @{ Min = 100000; Max = 109999; Name = 'Open File'         }
    @{ Min =  80000; Max =  89999; Name = 'Security Enhancer' }
    @{ Min =  60000; Max =  69999; Name = 'Copy Protection'   }
    @{ Min =  40000; Max =  49999; Name = 'Bottom'            }
    @{ Min =  20000; Max =  29999; Name = 'System'            }
)

# Service name -> product attribution. Microsoft entries have Product = $null so they are
# never counted as third-party. Used both to probe services and, when the script runs
# unelevated (no minifilter data), as the fallback evidence for product detection.
$Script:KnownServices = [ordered]@{
    'WinDefend'             = $null
    'WdNisSvc'              = $null
    'MDCoreSvc'             = $null
    'Sense'                 = $null
    'SecurityHealthService' = $null
    'wscsvc'                = $null
    'CSFalconService'       = 'CrowdStrike Falcon'
    'CSAgent'               = 'CrowdStrike Falcon'
    'SentinelAgent'         = 'SentinelOne'
    'CbDefense'             = 'VMware Carbon Black'
    'CylanceSvc'            = 'BlackBerry Cylance'
    'CyOpticsSvc'           = 'BlackBerry CylanceOPTICS'
    'ekrn'                  = 'ESET'
    'AVP'                   = 'Kaspersky'
    'SepMasterService'      = 'Symantec / Broadcom Endpoint Security'
    'masvc'                 = 'McAfee / Trellix'
    'mfemms'                = 'McAfee / Trellix'
    'ds_agent'              = 'Trend Micro Deep Security'
    'SophosEndpointDefense' = 'Sophos Endpoint Defense'
    'CybereasonAV'          = 'Cybereason'
    'PandaAgent'            = 'Panda Security'
    'VSSERV'                = 'Bitdefender'
    'xagt'                  = 'FireEye / Trellix Endpoint'
    'ZSATrayManager'        = 'Zscaler'
    'nxtsvc'                = 'Nexthink'
    'CyberArkEPMSvc'        = 'CyberArk Endpoint Privilege Manager'
}

# Some vendors version their service name (Kaspersky registers e.g. 'AVP21.16'), which an
# exact-name lookup misses. These regexes are checked against every installed service.
$Script:KnownServicePatterns = @(
    @{ Pattern = '^AVP[\d\.]+$'; Product = 'Kaspersky' }
)

# Process-name regex for the "what is running" section. Matched against Get-Process
# ProcessName, which carries NO .exe extension - anchor short names instead.
$Script:SecurityProcessPattern = 'MsSense|Sense[A-Z]|MsMpEng|MpDefenderCoreService|NisSrv|MpCmdRun|MpDlp|DlpUserAgent|' +
                                 'CSFalcon|SentinelAgent|SentinelUI|^cb$|RepMgr|RepUtils|CylanceSvc|CyOptics|' +
                                 'ekrn|egui|^avp$|ccSvcHst|mcshield|masvc|macmnsvc|TmListen|NTRtScan|' +
                                 'SophosFS|SSPService|CybereasonAV|PSANHost|vsserv|xagt|nxtsvc|vf_agent'

# Environment variables that resolve to the SYSTEM profile under the Defender service.
$Script:SystemScopedEnvVars = @('%USERPROFILE%', '%APPDATA%', '%LOCALAPPDATA%', '%TEMP%', '%TMP%', '%HOMEPATH%')

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Get-PropertyValue {
    <#
        Safe property read. Set-StrictMode 2.0 throws on non-existent properties, and the
        Defender cmdlets expose different property sets across platform versions, so every
        dynamic read goes through here.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowNull()] $InputObject,
        [Parameter(Mandatory)] [string]      $Name,
                               [object]      $Default = $null
    )

    if ($null -eq $InputObject) { return $Default }
    $prop = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $prop) { return $Default }
    if ($null -eq $prop.Value) { return $Default }
    return $prop.Value
}

function Test-Elevated {
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-RegistryValueMap {
    <#
        Returns every value under a registry key as an ordered hashtable, with binary blobs
        summarised and long strings truncated. Returns $null when the key is absent.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path,
                               [int]    $MaxStringLength = 90
    )

    if (-not (Test-Path -Path $Path)) { return $null }

    $map = [ordered]@{}
    try {
        $item = Get-ItemProperty -Path $Path -ErrorAction Stop
    } catch {
        return $null
    }

    $names = $item.PSObject.Properties.Name | Where-Object { $_ -notlike 'PS*' } | Sort-Object
    foreach ($name in $names) {
        $value = $item.PSObject.Properties[$name].Value
        if ($value -is [byte[]]) {
            $value = '<binary, {0} bytes>' -f $value.Length
        } elseif ($value -is [string] -and $value.Length -gt $MaxStringLength) {
            $value = $value.Substring(0, $MaxStringLength) + '...'
        }
        $map[$name] = $value
    }
    return $map
}

function Resolve-AltitudeBand {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)] [double] $Altitude
    )

    foreach ($band in $Script:AltitudeBands) {
        if ($Altitude -ge $band.Min -and $Altitude -le $band.Max) { return $band.Name }
    }
    return 'Unallocated'
}

function Get-MinifilterInventory {
    <#
        Parses `fltmc.exe filters`. Requires elevation. Returns an empty array when
        unelevated or when fltmc is unavailable, so callers can degrade gracefully.
    #>
    [CmdletBinding()]
    param()

    $filters = New-Object System.Collections.Generic.List[object]

    try {
        $raw = & fltmc.exe filters 2>$null
    } catch {
        return $filters.ToArray()
    }
    if ($null -eq $raw) { return $filters.ToArray() }

    foreach ($line in $raw) {
        $text = "$line".Trim()
        if ($text -eq '' -or $text -like 'Filter Name*' -or $text -like '---*') { continue }

        # Filter Name | Num Instances | Altitude | Frame
        $m = [regex]::Match($text, '^(?<name>\S+)\s+(?<inst>\d+)\s+(?<alt>[\d\.]+)\s+(?<frame>\d+)')
        if (-not $m.Success) { continue }

        $name     = $m.Groups['name'].Value
        $altitude = [double]$m.Groups['alt'].Value
        $key      = $name.ToLowerInvariant()

        $product    = 'Unrecognised'
        $isSecurity = $false
        $known      = $false
        if ($Script:KnownFilters.ContainsKey($key)) {
            $known      = $true
            $product    = $Script:KnownFilters[$key].Product
            $isSecurity = $Script:KnownFilters[$key].Security
        }

        [void]$filters.Add([pscustomobject]@{
            Name          = $name
            Instances     = [int]$m.Groups['inst'].Value
            Altitude      = $m.Groups['alt'].Value
            Band          = Resolve-AltitudeBand -Altitude $altitude
            Product       = $product
            IsSecurity    = $isSecurity
            Recognised    = $known
        })
    }

    return ($filters.ToArray() | Sort-Object { [double]$_.Altitude } -Descending)
}

function Test-ExclusionHygiene {
    <#
        Reviews the effective exclusion lists for rules that are broken, inert, or wider
        than intended. Returns a list of finding objects; an empty list means nothing found.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()] [string[]] $ExclusionPath,
        [AllowNull()] [string[]] $ExclusionProcess,
        [AllowNull()] [string[]] $ExclusionExtension
    )

    $findings = New-Object System.Collections.Generic.List[object]

    function Add-Finding {
        param([string]$Severity, [string]$Category, [string]$Entry, [string]$Detail)
        [void]$findings.Add([pscustomobject]@{
            Severity = $Severity
            Category = $Category
            Entry    = $Entry
            Detail   = $Detail
        })
    }

    foreach ($entry in @($ExclusionPath)) {
        if ([string]::IsNullOrWhiteSpace($entry)) { continue }
        $trimmed = $entry.Trim()
        $upper   = $trimmed.ToUpperInvariant()

        # 1. Environment variables that resolve under LocalSystem.
        foreach ($var in $Script:SystemScopedEnvVars) {
            if ($upper.Contains($var)) {
                Add-Finding 'High' 'EnvVarUnderLocalSystem' $trimmed (
                    "$var resolves to the SYSTEM profile because the Defender service runs as " +
                    'LocalSystem, so this rule matches nothing. Rewrite using C:\Users\*\.')
                break
            }
        }

        # 2. Bare filename with no path component.
        if ($trimmed -notmatch '[\\/]' -and $trimmed -notmatch '^[A-Za-z]:') {
            Add-Finding 'Medium' 'BareNameInPathList' $trimmed (
                'A path exclusion requires a path. This entry matches nothing; it was probably ' +
                'meant to be a process exclusion.')
        }

        # 3. Auto-start locations excluded from scanning.
        if ($upper.Contains('\STARTUP') -or $upper.Contains('\START MENU\PROGRAMS')) {
            Add-Finding 'High' 'AutoStartLocationExcluded' $trimmed (
                'Excluding a user-writable auto-start location removes scanning from a common ' +
                'persistence path. Exclude the signed binary in its installed location instead.')
        }

        # 4. Drive root or single top-level folder.
        if ($trimmed -match '^[A-Za-z]:\\?$') {
            Add-Finding 'High' 'DriveRootExcluded' $trimmed 'An entire drive is excluded from scanning.'
        } elseif ($trimmed -match '^[A-Za-z]:\\[^\\]+\\?\*?$') {
            Add-Finding 'Low' 'BroadRootExcluded' $trimmed (
                'A whole top-level folder is excluded. Confirm this is intentional and as narrow as possible.')
        }
    }

    foreach ($entry in @($ExclusionProcess)) {
        if ([string]::IsNullOrWhiteSpace($entry)) { continue }
        $trimmed = $entry.Trim()

        # 5. Kernel drivers listed as processes.
        if ($trimmed -match '\.sys$') {
            Add-Finding 'Low' 'DriverAsProcess' $trimmed (
                'A kernel driver never runs as a process, so this rule does nothing. Harmless if ' +
                'the containing folder is already path-excluded; otherwise it is a gap.')
        }
        # 6. Non-executables listed as processes.
        elseif ($trimmed -notmatch '\.(exe|com|scr|bat|cmd|ps1)$') {
            Add-Finding 'Low' 'NonExecutableAsProcess' $trimmed (
                'Only executables can be process exclusions. This rule does nothing.')
        }

        # 7. Bare-name process exclusions.
        if ($trimmed -notmatch '[\\/]' -and $trimmed -match '\.(exe|com|scr)$') {
            Add-Finding 'Medium' 'BareNameProcess' $trimmed (
                'This works, but matches any file of that name from any location. Use a full path.')
        }

        # 8. Environment variables here too.
        $upperProc = $trimmed.ToUpperInvariant()
        foreach ($var in $Script:SystemScopedEnvVars) {
            if ($upperProc.Contains($var)) {
                Add-Finding 'High' 'EnvVarUnderLocalSystem' $trimmed (
                    "$var resolves to the SYSTEM profile under the Defender service; this rule matches nothing.")
                break
            }
        }
    }

    # File types that should never be excluded wholesale: an extension exclusion applies to
    # every file of that type on every drive, wherever it came from.
    $execExtensions = @('exe', 'com', 'scr', 'pif', 'dll', 'sys', 'ocx', 'cpl', 'ps1', 'psm1',
                        'bat', 'cmd', 'vbs', 'vbe', 'js', 'jse', 'wsf', 'wsh', 'hta', 'msi', 'msp')

    foreach ($entry in @($ExclusionExtension)) {
        if ([string]::IsNullOrWhiteSpace($entry)) { continue }
        $trimmed = $entry.Trim()
        $ext = $trimmed.TrimStart('*').TrimStart('.').ToLowerInvariant()

        # 9. A path where an extension belongs.
        if ($ext -match '[\\/]') {
            Add-Finding 'Medium' 'PathAsExtension' $trimmed (
                'An extension exclusion must be a bare extension (e.g. log). This entry looks like ' +
                'a path and matches nothing; move it to the path exclusion list.')
            continue
        }

        # 10. Executable or script types excluded wholesale.
        if ($execExtensions -contains $ext) {
            Add-Finding 'High' 'ExecutableExtensionExcluded' $trimmed (
                "Every .$ext file on every drive is excluded from scanning, wherever it came " +
                'from. Exclude specific paths or processes instead.')
        }
    }

    return $findings.ToArray()
}

function Get-SenseEventSummary {
    <#
        Reads the SENSE operational log. Two things live here and nowhere else on the
        device: the sensor's connectivity history, and the CSP read-back events the
        management layer logs when it queries the sensor:
          1803 Last Connected   1804 Org ID          1805 Sense Is Running
          1806 Onboarding State 1807 Onboarding Blob 1809 Sample Sharing
          1823 Telemetry Reporting Frequency
        The CSP events appear only on MDM-managed devices; their absence is not a finding.
        Connectivity events: 4, 5, 6, 7, 20 (4 = service contacted the server successfully,
        per Microsoft's published event table for this log).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [int] $MaxEvents
    )

    $result = [ordered]@{
        Available        = $false
        Config           = [ordered]@{}
        Connectivity     = @()
        Errors           = @()
        SawServerContact = $false
        Note             = ''
    }

    try {
        $events = Get-WinEvent -LogName 'Microsoft-Windows-SENSE/Operational' -MaxEvents $MaxEvents -ErrorAction Stop
    } catch {
        $result.Note = 'SENSE log not readable: ' + $_.Exception.Message
        return [pscustomobject]$result
    }

    $result.Available = $true
    $configIds = 1803, 1804, 1805, 1806, 1807, 1809, 1823

    foreach ($id in $configIds) {
        $match = $events | Where-Object { $_.Id -eq $id } | Select-Object -First 1
        if ($null -ne $match) {
            $text = [regex]::Replace(($match.Message -replace "`r`n", ' '), '\s+', ' ')
            $result.Config["$id"] = $text.Trim()
        }
    }

    $connIds = 4, 5, 6, 7, 20
    $result.Connectivity = @(
        $events | Where-Object { $connIds -contains $_.Id } | Select-Object -First 6 | ForEach-Object {
            $text = [regex]::Replace(($_.Message -replace "`r`n", ' '), '\s+', ' ')
            [pscustomobject]@{
                Time = $_.TimeCreated
                Id   = $_.Id
                Text = $text.Trim()
            }
        }
    )
    $result.SawServerContact = @($events | Where-Object { $_.Id -eq 4 }).Count -gt 0

    $result.Errors = @(
        $events | Where-Object { @('Error', 'Warning') -contains $_.LevelDisplayName } |
            Select-Object -First 10 | ForEach-Object {
                $text = [regex]::Replace(($_.Message -replace "`r`n", ' '), '\s+', ' ')
                if ($text.Length -gt 160) { $text = $text.Substring(0, 160) + '...' }
                [pscustomobject]@{
                    Time  = $_.TimeCreated
                    Id    = $_.Id
                    Level = $_.LevelDisplayName
                    Text  = $text.Trim()
                }
            }
    )

    return [pscustomobject]$result
}

function Write-Section {
    param([string]$Title)
    Write-Host ''
    Write-Host ('  ' + $Title) -ForegroundColor White
    Write-Host ('  ' + ('-' * 68)) -ForegroundColor DarkGray
}

function Write-Pair {
    param([string]$Label, $Value, [string]$Colour = 'Gray')
    if ($null -eq $Value -or "$Value" -eq '') { $Value = '(not set)' }
    Write-Host ('  {0,-32}: {1}' -f $Label, $Value) -ForegroundColor $Colour
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

try {
    $isAdmin = Test-Elevated
    $partial = -not $isAdmin

    # --- 1. Machine ---------------------------------------------------------
    $osBuild = 'unknown'
    try {
        $cv = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
        $osBuild = '{0}.{1}' -f (Get-PropertyValue $cv 'CurrentBuild' '?'), (Get-PropertyValue $cv 'UBR' '?')
    } catch {
        Write-Verbose "Could not read CurrentVersion: $($_.Exception.Message)"
    }

    # --- 2. Services --------------------------------------------------------
    $allServices = @{}
    foreach ($svc in (Get-Service -ErrorAction SilentlyContinue)) { $allServices[$svc.Name] = $svc }

    $services = New-Object System.Collections.Generic.List[object]
    foreach ($name in $Script:KnownServices.Keys) {
        if ($allServices.ContainsKey($name)) {
            $s = $allServices[$name]
            [void]$services.Add([pscustomobject]@{
                Name      = $s.Name
                Status    = "$($s.Status)"
                StartType = "$($s.StartType)"
                Product   = $Script:KnownServices[$name]
                Present   = $true
            })
        }
    }
    foreach ($name in $allServices.Keys) {
        if ($Script:KnownServices.Contains($name)) { continue }
        foreach ($pat in $Script:KnownServicePatterns) {
            if ($name -match $pat.Pattern) {
                $s = $allServices[$name]
                [void]$services.Add([pscustomobject]@{
                    Name      = $s.Name
                    Status    = "$($s.Status)"
                    StartType = "$($s.StartType)"
                    Product   = $pat.Product
                    Present   = $true
                })
                break
            }
        }
    }

    # --- 3. Processes -------------------------------------------------------
    $processes = @(
        Get-Process -ErrorAction SilentlyContinue |
            Where-Object { $_.ProcessName -match $Script:SecurityProcessPattern } |
            Group-Object ProcessName | Sort-Object Name |
            ForEach-Object { [pscustomobject]@{ Name = $_.Name; Count = $_.Count } }
    )

    # --- 4/5. Defender for Endpoint onboarding + policy ---------------------
    $atpStatus = Get-RegistryValueMap -Path 'HKLM:\SOFTWARE\Microsoft\Windows Advanced Threat Protection\Status'
    $atpPolicy = Get-RegistryValueMap -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Advanced Threat Protection'

    $onboardingState = $null
    if ($null -ne $atpStatus -and $atpStatus.Contains('OnboardingState')) { $onboardingState = $atpStatus['OnboardingState'] }
    $forcePassive = $null
    if ($null -ne $atpPolicy -and $atpPolicy.Contains('ForceDefenderPassiveMode')) { $forcePassive = $atpPolicy['ForceDefenderPassiveMode'] }

    # --- 6/7. Defender AV state + preferences -------------------------------
    $mpStatus = $null
    $mpPref   = $null
    try { $mpStatus = Get-MpComputerStatus -ErrorAction Stop } catch { $partial = $true }
    try { $mpPref   = Get-MpPreference    -ErrorAction Stop } catch { $partial = $true }

    $avState = [ordered]@{
        RunningMode           = Get-PropertyValue $mpStatus 'AMRunningMode'             'unavailable'
        ServiceEnabled        = Get-PropertyValue $mpStatus 'AMServiceEnabled'          $null
        RealTimeProtection    = Get-PropertyValue $mpStatus 'RealTimeProtectionEnabled' $null
        BehaviorMonitor       = Get-PropertyValue $mpStatus 'BehaviorMonitorEnabled'    $null
        NisEnabled            = Get-PropertyValue $mpStatus 'NISEnabled'                $null
        TamperProtection      = Get-PropertyValue $mpStatus 'IsTamperProtected'         $null
        SignatureVersion      = Get-PropertyValue $mpStatus 'AntivirusSignatureVersion' $null
        EngineVersion         = Get-PropertyValue $mpStatus 'AMEngineVersion'           $null
        PlatformVersion       = Get-PropertyValue $mpStatus 'AMProductVersion'          $null
    }

    $avPrefs = [ordered]@{
        MapsReporting            = Get-PropertyValue $mpPref 'MAPSReporting'               $null
        SubmitSamplesConsent     = Get-PropertyValue $mpPref 'SubmitSamplesConsent'        $null
        EnableNetworkProtection  = Get-PropertyValue $mpPref 'EnableNetworkProtection'     $null
        PuaProtection            = Get-PropertyValue $mpPref 'PUAProtection'               $null
        ScanAvgCpuLoadFactor     = Get-PropertyValue $mpPref 'ScanAvgCPULoadFactor'        $null
        EnableLowCpuPriority     = Get-PropertyValue $mpPref 'EnableLowCPUPriority'        $null
        DisableCatchupFullScan   = Get-PropertyValue $mpPref 'DisableCatchupFullScan'      $null
        ControlledFolderAccess   = Get-PropertyValue $mpPref 'EnableControlledFolderAccess' $null
        DisableLocalAdminMerge   = Get-PropertyValue $mpPref 'DisableLocalAdminMerge'      $null
        DisableScanningNetwork   = Get-PropertyValue $mpPref 'DisableScanningNetworkFiles' $null
    }

    # --- 8. ASR rules -------------------------------------------------------
    $asrRules = @()
    $asrIds = Get-PropertyValue $mpPref 'AttackSurfaceReductionRules_Ids' @()
    $asrActions = Get-PropertyValue $mpPref 'AttackSurfaceReductionRules_Actions' @()
    if (@($asrIds).Count -gt 0) {
        $asrRules = @(
            for ($i = 0; $i -lt @($asrIds).Count; $i++) {
                $action = 'unknown'
                if ($i -lt @($asrActions).Count) { $action = "$($asrActions[$i])" }
                [pscustomobject]@{ RuleId = "$($asrIds[$i])"; Action = $action }
            }
        )
    }

    # --- 9. Exclusions ------------------------------------------------------
    # Unelevated, Get-MpPreference returns the literal sentinel
    # "N/A: Must be an administrator to view exclusions" in place of each list. Treat that
    # as "not collected" rather than feeding it into the hygiene review.
    $exPath      = @(Get-PropertyValue $mpPref 'ExclusionPath'      @())
    $exProcess   = @(Get-PropertyValue $mpPref 'ExclusionProcess'   @())
    $exExtension = @(Get-PropertyValue $mpPref 'ExclusionExtension' @())

    $exclusionsReadable = $true
    foreach ($candidate in (@($exPath) + @($exProcess) + @($exExtension))) {
        if ("$candidate" -like 'N/A:*') { $exclusionsReadable = $false; break }
    }

    if ($exclusionsReadable) {
        $exFindings = Test-ExclusionHygiene -ExclusionPath $exPath -ExclusionProcess $exProcess -ExclusionExtension $exExtension
    } else {
        $exPath = @(); $exProcess = @(); $exExtension = @()
        $exFindings = @()
        $partial = $true
    }

    # --- 10. Minifilters ----------------------------------------------------
    $filters = @()
    if ($isAdmin) { $filters = @(Get-MinifilterInventory) } else { $partial = $true }
    $securityFilters = @($filters | Where-Object { $_.IsSecurity })
    $unrecognised    = @($filters | Where-Object { -not $_.Recognised })

    # --- 11. Windows Security Center ---------------------------------------
    $wscProducts = @()
    try {
        $wscProducts = @(
            Get-CimInstance -Namespace 'root\SecurityCenter2' -ClassName AntiVirusProduct -ErrorAction Stop |
                ForEach-Object {
                    [pscustomobject]@{
                        DisplayName  = Get-PropertyValue $_ 'displayName' ''
                        ProductState = Get-PropertyValue $_ 'productState' ''
                        Path         = Get-PropertyValue $_ 'pathToSignedProductExe' ''
                    }
                }
        )
    } catch {
        Write-Verbose "SecurityCenter2 unavailable: $($_.Exception.Message)"
    }

    # --- 12/13. SENSE event log --------------------------------------------
    $sense = $null
    if (-not $SkipEventLog) { $sense = Get-SenseEventSummary -MaxEvents $MaxEvents }

    # --- 14. Defender platform versions ------------------------------------
    $platformVersions = @()
    $platformRoot = Join-Path $env:ProgramData 'Microsoft\Windows Defender\Platform'
    if (Test-Path $platformRoot) {
        $platformVersions = @(Get-ChildItem -Path $platformRoot -Directory -ErrorAction SilentlyContinue |
            Sort-Object Name | Select-Object -ExpandProperty Name)
    }

    # --- 15. Verdict --------------------------------------------------------
    # Minifilters are the strongest evidence, but they need elevation. Fall back to running
    # services and processes so an unelevated run still produces a usable verdict.
    $runningServiceNames = @($services | Where-Object { $_.Status -eq 'Running' } | Select-Object -ExpandProperty Name)

    $microsoftAv  = (@($securityFilters | Where-Object { $_.Name -match '^(?i)wdfilter$' }).Count -gt 0) -or
                    ($runningServiceNames -contains 'WinDefend')
    $microsoftEdr = (@($securityFilters | Where-Object { $_.Name -match '^(?i)mssecflt$' }).Count -gt 0) -or
                    ($runningServiceNames -contains 'Sense')

    $thirdPartyProducts = New-Object System.Collections.Generic.List[string]
    foreach ($f in $securityFilters) {
        if ($f.Product -ne 'Unrecognised' -and $f.Product -notlike 'Microsoft*' -and
            -not $thirdPartyProducts.Contains($f.Product)) {
            [void]$thirdPartyProducts.Add($f.Product)
        }
    }
    foreach ($s in $services) {
        if ($null -ne $s.Product -and $s.Status -eq 'Running' -and -not $thirdPartyProducts.Contains($s.Product)) {
            [void]$thirdPartyProducts.Add($s.Product)
        }
    }
    foreach ($p in $wscProducts) {
        if ($p.DisplayName -notlike '*Defender*' -and $p.DisplayName -ne '' -and
            -not $thirdPartyProducts.Contains($p.DisplayName)) {
            [void]$thirdPartyProducts.Add($p.DisplayName)
        }
    }

    $coexistence = $thirdPartyProducts.Count -gt 0 -and ($microsoftAv -or $microsoftEdr)

    # A sensor that has never reported a successful cloud contact (event 67) while the device
    # is onboarded is a real health failure. Error/Warning records on their own are noisy and
    # are reported for context only - they do not on their own make the sensor unhealthy.
    $senseUnhealthy = $false
    $senseErrorCount = 0
    if ($null -ne $sense -and $sense.Available) {
        $senseErrorCount = @($sense.Errors).Count
        if ("$onboardingState" -eq '1' -and -not $sense.SawServerContact) { $senseUnhealthy = $true }
    }
    $tamperOff = ($avState.TamperProtection -eq $false)
    # Tamper Protection off is CRITICAL only while Defender AV is the active engine. In
    # Passive / EDR Block mode a third-party AV is primary and IsTamperProtected is routinely
    # false; treating that as critical would fail every intentionally third-party-primary device.
    $tamperOffCritical = $tamperOff -and ("$($avState.RunningMode)" -eq 'Normal')

    $result = [pscustomobject]@{
        ComputerName      = $env:COMPUTERNAME
        CollectedUtc      = (Get-Date).ToUniversalTime().ToString('s') + 'Z'
        OsBuild           = $osBuild
        Elevated          = $isAdmin
        Partial           = $partial
        DefenderAv        = [pscustomobject]$avState
        DefenderAvPrefs   = [pscustomobject]$avPrefs
        Mde               = [pscustomobject][ordered]@{
            OnboardingState          = $onboardingState
            ForceDefenderPassiveMode = $forcePassive
            StatusKey                = $atpStatus
            PolicyKey                = $atpPolicy
            Sense                    = $sense
        }
        Services          = $services.ToArray()
        Processes         = $processes
        Filters           = $filters
        SecurityFilters   = $securityFilters
        UnrecognisedFilters = $unrecognised
        WscProducts       = $wscProducts
        AsrRules          = $asrRules
        Exclusions        = [pscustomobject][ordered]@{
            Readable       = $exclusionsReadable
            PathCount      = @($exPath).Count
            ProcessCount   = @($exProcess).Count
            ExtensionCount = @($exExtension).Count
            Path           = $exPath
            Process        = $exProcess
            Extension      = $exExtension
            Findings       = $exFindings
        }
        PlatformVersions  = $platformVersions
        Verdict           = [pscustomobject][ordered]@{
            MicrosoftAvActive     = $microsoftAv
            MicrosoftEdrActive    = $microsoftEdr
            ThirdPartyProducts    = $thirdPartyProducts.ToArray()
            SecurityFilterCount   = @($securityFilters).Count
            FiltersCollected      = $isAdmin
            Coexistence           = $coexistence
            SenseUnhealthy        = $senseUnhealthy
            SenseErrorCount       = $senseErrorCount
            TamperProtectionOff   = $tamperOff
            ExclusionFindingCount = @($exFindings).Count
        }
    }

    # --- Console report -----------------------------------------------------
    if (-not $Quiet) {
        Write-Host ''
        Write-Host "  Defender / EDR coexistence state - $($result.ComputerName)" -ForegroundColor Cyan
        Write-Host '  Read-only. No changes were made.' -ForegroundColor DarkGray

        Write-Section 'Machine'
        Write-Pair 'OS build'  $osBuild
        Write-Pair 'Elevated'  $isAdmin $(if ($isAdmin) { 'Gray' } else { 'Yellow' })
        if (-not $isAdmin) {
            Write-Host '  Not elevated - the minifilter section was skipped.' -ForegroundColor Yellow
        }

        Write-Section 'Defender Antivirus'
        Write-Pair 'Running mode'        $avState.RunningMode
        Write-Pair 'Real-time protection' $avState.RealTimeProtection
        Write-Pair 'Behaviour monitoring' $avState.BehaviorMonitor
        Write-Pair 'Tamper protection'   $avState.TamperProtection $(if ($tamperOff) { 'Red' } else { 'Gray' })
        Write-Pair 'Platform version'    $avState.PlatformVersion
        Write-Pair 'Cloud (MAPSReporting)' $avPrefs.MapsReporting
        Write-Pair 'Network protection'  $avPrefs.EnableNetworkProtection
        Write-Pair 'Low CPU priority'    $avPrefs.EnableLowCpuPriority

        Write-Section 'Defender for Endpoint'
        Write-Pair 'Onboarding state'    $onboardingState
        Write-Pair 'ForceDefenderPassiveMode' $forcePassive
        if ($null -ne $sense -and $sense.Available) {
            Write-Pair 'Server contact (event 4)' $sense.SawServerContact $(if ($sense.SawServerContact) { 'Gray' } else { 'Red' })
            Write-Pair 'Recent sensor errors/warnings' $senseErrorCount $(if ($senseErrorCount -gt 0) { 'Yellow' } else { 'Gray' })
        } elseif ($SkipEventLog) {
            Write-Host '  SENSE log skipped (-SkipEventLog).' -ForegroundColor DarkGray
        } else {
            Write-Host '  SENSE log not available.' -ForegroundColor DarkGray
        }

        if (@($filters).Count -gt 0) {
            Write-Section "Security filter drivers ($(@($securityFilters).Count) of $(@($filters).Count) loaded filters)"
            foreach ($f in $securityFilters) {
                Write-Host ('  {0,-20} {1,12}  {2,-18} {3}' -f $f.Name, $f.Altitude, $f.Band, $f.Product) -ForegroundColor Gray
            }
            if (@($unrecognised).Count -gt 0) {
                Write-Host ''
                Write-Host '  Unrecognised filters (review manually):' -ForegroundColor DarkGray
                foreach ($f in $unrecognised) {
                    Write-Host ('  {0,-20} {1,12}  {2}' -f $f.Name, $f.Altitude, $f.Band) -ForegroundColor DarkGray
                }
            }
        }

        Write-Section 'Windows Security Center'
        if (@($wscProducts).Count -eq 0) {
            Write-Host '  No registered antivirus providers returned.' -ForegroundColor DarkGray
        } else {
            foreach ($p in $wscProducts) { Write-Host ('  ' + $p.DisplayName) -ForegroundColor Gray }
        }

        Write-Section 'Exclusions'
        if (-not $exclusionsReadable) {
            Write-Host '  Exclusion lists require elevation - not collected.' -ForegroundColor Yellow
        }
        Write-Pair 'Path exclusions'      $result.Exclusions.PathCount
        Write-Pair 'Process exclusions'   $result.Exclusions.ProcessCount
        Write-Pair 'Extension exclusions' $result.Exclusions.ExtensionCount
        if (@($exFindings).Count -gt 0) {
            Write-Host ''
            Write-Host ('  {0} hygiene finding(s):' -f @($exFindings).Count) -ForegroundColor Yellow
            $sevRank = @{ High = 0; Medium = 1; Low = 2 }
            foreach ($f in ($exFindings | Sort-Object { $sevRank[$_.Severity] }, Category)) {
                $colour = 'Gray'
                if ($f.Severity -eq 'High')   { $colour = 'Red' }
                if ($f.Severity -eq 'Medium') { $colour = 'Yellow' }
                Write-Host ('  [{0,-6}] {1,-26} {2}' -f $f.Severity, $f.Category, $f.Entry) -ForegroundColor $colour
            }
        } elseif ($exclusionsReadable) {
            Write-Host '  No exclusion hygiene findings.' -ForegroundColor Green
        }

        Write-Section 'Verdict'
        Write-Pair 'Microsoft Defender AV active'  $microsoftAv
        Write-Pair 'Microsoft Defender EDR active' $microsoftEdr
        Write-Pair 'Third-party security products' $(
            if ($thirdPartyProducts.Count -gt 0) { $thirdPartyProducts -join ', ' } else { 'none detected' })
        if (-not $isAdmin) {
            Write-Host '  Filter-driver evidence was unavailable; the verdict rests on services,' -ForegroundColor DarkGray
            Write-Host '  processes and Security Center registration only. Re-run elevated.' -ForegroundColor DarkGray
        }
        if ($coexistence) {
            Write-Host '  Multiple security vendors are active on this endpoint.' -ForegroundColor Yellow
            Write-Host '  Whether that is correct depends on the intended architecture.' -ForegroundColor DarkGray
        }
        Write-Host ''
    }

    # --- Outputs ------------------------------------------------------------
    if ($JsonPath) {
        try {
            $result | ConvertTo-Json -Depth 8 | Out-File -FilePath $JsonPath -Encoding UTF8
            if (-not $Quiet) { Write-Host "  JSON written to $JsonPath" -ForegroundColor DarkGray }
        } catch {
            Write-Warning "Could not write JSON to '$JsonPath': $($_.Exception.Message)"
        }
    }

    if ($AsObject) {
        Write-Output $result
    } elseif ($Quiet) {
        $summary = [ordered]@{
            Computer            = $result.ComputerName
            OsBuild             = $osBuild
            AvRunningMode       = $avState.RunningMode
            MdeOnboarded        = $onboardingState
            Coexistence         = $coexistence
            ThirdParty          = $thirdPartyProducts.ToArray()
            SecurityFilters     = @($securityFilters).Count
            ExclusionFindings   = @($exFindings).Count
            TamperProtectionOff = $tamperOff
            SenseUnhealthy      = $senseUnhealthy
            Partial             = $partial
        }
        Write-Output ([pscustomobject]$summary | ConvertTo-Json -Compress)
    }

    # --- Exit code ----------------------------------------------------------
    if ($senseUnhealthy -or $tamperOffCritical) { exit 2 }
    if ($coexistence -or @($exFindings).Count -gt 0) { exit 1 }
    if ($partial) { exit 3 }
    exit 0

} catch {
    Write-Error ("Unexpected failure: {0}`n{1}" -f $_.Exception.Message, $_.ScriptStackTrace)
    exit 4
}
