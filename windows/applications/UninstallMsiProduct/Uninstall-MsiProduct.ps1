<#
.SYNOPSIS
    Uninstalls one or more MSI-based products by name pattern, publisher, version,
    or explicit ProductCode -- without needing the original .msi file.

.DESCRIPTION
    Generic, idempotent MSI uninstaller for fleet use (Intune Win32 app, ConfigMgr
    package, scheduled task, AIB customizer). Designed for vendor agents whose
    ProductCode (PackageCode GUID) changes with every release, so you can't pin a
    single MSI GUID into your deployment tooling -- e.g. the Quest / KACE Agent
    documented in:

        https://support.quest.com/de-de/kb/4269674

    Discovery uses the Uninstall registry hive (NOT Win32_Product -- that class
    triggers an MSI self-repair pass on every installed product as a side effect
    of being enumerated). Both 64-bit and 32-bit (WOW6432Node) hives are scanned;
    per-user installs are scanned when -IncludePerUser is supplied.

    A registry entry is treated as MSI-uninstallable only when ALL of these hold:
      * The subkey name is a {GUID} (the ProductCode).
      * WindowsInstaller = 1, OR the UninstallString begins with msiexec.
      * DisplayName is populated.
      * SystemComponent != 1 (skips servicing / KB-style entries).
      * ParentKeyName is absent (skips MSI patches/sub-components).

    Each match is removed with:
        msiexec.exe /x {GUID} /qn /norestart /l*v "<log>" <AdditionalArguments>

    Exit codes 0 and 3010 are treated as success; 1605 (product not installed --
    common when a parallel install has just removed the same GUID) is treated as
    a no-op success. Everything else is a failure.

    SAFETY GATES
      * SupportsShouldProcess + ConfirmImpact='High' -- without -Confirm:$false
        or an explicit -Force, every match prompts.
      * If discovery returns more than -MaxMatches products, the run aborts
        unless -Force is supplied. Default: 10. Prevents a stray wildcard
        ('*' or '*agent*') from nuking the whole machine.
      * -WhatIf prints the discovery list and exits without invoking msiexec.

.PARAMETER DisplayName
    One or more DisplayName patterns. Wildcards (*, ?) supported -- matching is
    case-insensitive via PowerShell -like. Multiple patterns are OR'd together.
    Example: 'Quest*Agent','KACE Agent'.

.PARAMETER Publisher
    Optional Publisher wildcard filter applied AFTER -DisplayName. Useful to
    disambiguate generically named agents.
    Example: '*Quest*'.

.PARAMETER Version
    Optional DisplayVersion wildcard filter (string match, not SemVer compare).
    Example: '10.*' or '11.0.123.0'.

.PARAMETER ProductCode
    One or more explicit ProductCode GUIDs (with or without braces). When
    supplied, the registry discovery filter is bypassed for these GUIDs --
    msiexec /x is invoked directly. Other filter parameters are ignored.
    Use this when you already know the GUID(s) you want gone.

.PARAMETER IncludePerUser
    Also scan the per-user Uninstall hive (HKCU:\...\Uninstall) of the user
    running this script. Per-user installs of OTHER profiles are NOT touched
    (would require loading each ntuser.dat -- out of scope).

.PARAMETER Architecture
    Restrict discovery to a single registry view:
        Both : HKLM 64-bit + WOW6432Node + (optional) HKCU. Default.
        X64  : HKLM 64-bit only (skips WOW6432Node).
        X86  : WOW6432Node only.

.PARAMETER AdditionalArguments
    Extra arguments appended to the msiexec command line, AFTER the default
    '/qn /norestart /l*v <log>'. Use to pass MSI properties such as
    REMOVE=ALL, REBOOT=ReallySuppress, or vendor-specific suppress switches.
    Example: 'REBOOT=ReallySuppress','MSIRESTARTMANAGERCONTROL=Disable'.

.PARAMETER MaxMatches
    Safety cap on how many products a single run may target. Default 10.
    If discovery returns more, the run aborts (exit 1604) unless -Force is
    supplied. Use a higher value or -Force for bulk cleanups.

.PARAMETER Force
    Suppress per-product confirmation prompts AND override the -MaxMatches
    safety gate. Equivalent to -Confirm:$false + ignoring MaxMatches. Use
    in non-interactive contexts (Intune, scheduled task).

.PARAMETER TimeoutSeconds
    Per-product msiexec timeout. If msiexec hasn't returned within the
    window, it is killed and the product is marked as failed (exit 1460).
    Default 900 (15 min). Set to 0 to wait forever.

.PARAMETER LogDirectory
    Folder for the script log AND per-product MSI verbose logs (one .log
    per ProductCode). Default: $env:TEMP. Folder is created if missing.

.PARAMETER ListOnly
    Discovery-only mode: print the list of matching products and exit
    without invoking msiexec. Equivalent to -WhatIf but without the
    'WhatIf:' prefix in output. Useful for scripted pre-checks.

.OUTPUTS
    A PSCustomObject for each attempt with:
        DisplayName, DisplayVersion, Publisher, ProductCode, Architecture,
        Scope (Machine / PerUser), ExitCode, Action (Removed / RebootPending /
        NotInstalled / Failed / Skipped / WhatIf), DurationSeconds, MsiLog.

.NOTES
    File:    windows/applications/UninstallMsiProduct/Uninstall-MsiProduct.ps1
    Author:  Anton Romanyuk
    Version: 1.0.0
    Requires: PowerShell 5.1+, elevated session for machine-wide uninstalls.

    Exit codes (script-level, separate from per-product results):
      0    - All matches uninstalled successfully (or no matches when -ListOnly)
      1602 - User cancelled at the confirmation prompt
      1603 - At least one product failed to uninstall
      1604 - Too many matches; -MaxMatches gate tripped (use -Force or refine filter)
      1605 - No matching products found (only when at least one filter was supplied)
      3010 - Success but a reboot is required

.DISCLAIMER
    This script is provided "AS IS" with no warranties and confers no rights.
    Test in a non-production environment first. Vendor agents may require
    additional cleanup (services, drivers, certificates, scheduled tasks)
    that this script does NOT perform -- it only invokes the MSI's documented
    uninstall sequence.

.EXAMPLE
    # List every MSI whose name starts with "Quest" -- no changes made
    .\Uninstall-MsiProduct.ps1 -DisplayName 'Quest*' -ListOnly

.EXAMPLE
    # Remove the Quest KACE Agent (any version), prompt-free, non-interactive
    .\Uninstall-MsiProduct.ps1 -DisplayName 'Quest*KACE Agent*' -Publisher '*Quest*' -Force

.EXAMPLE
    # Remove a specific MSI by ProductCode without enumerating the registry
    .\Uninstall-MsiProduct.ps1 -ProductCode '{12345678-90AB-CDEF-1234-567890ABCDEF}' -Force

.EXAMPLE
    # Bulk cleanup of all "Acme " agents older than 5.0 (raise the safety cap)
    .\Uninstall-MsiProduct.ps1 -DisplayName 'Acme *' -Version '4.*' -MaxMatches 50 -Force

.EXAMPLE
    # Pass a vendor-specific property to suppress vendor reboot logic
    .\Uninstall-MsiProduct.ps1 -DisplayName 'Foglight*' -Force `
        -AdditionalArguments 'REBOOT=ReallySuppress','MSIRESTARTMANAGERCONTROL=Disable'

.EXAMPLE
    # WhatIf preview: see exactly which products would be touched
    .\Uninstall-MsiProduct.ps1 -DisplayName '*agent*' -WhatIf
#>

#Requires -Version 5.1
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string[]]$DisplayName,
    [string]  $Publisher,
    [string]  $Version,
    [string[]]$ProductCode,
    [switch]  $IncludePerUser,
    [ValidateSet('Both', 'X64', 'X86')] [string]$Architecture = 'Both',
    [string[]]$AdditionalArguments = @(),
    [ValidateRange(1, 1000)] [int]$MaxMatches = 10,
    [switch]  $Force,
    [ValidateRange(0, 86400)] [int]$TimeoutSeconds = 900,
    [string]  $LogDirectory = $env:TEMP,
    [switch]  $ListOnly
)

$ErrorActionPreference = 'Stop'

# -----------------------------------------------------------------------------
# LOGGING
# -----------------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $LogDirectory)) {
    New-Item -ItemType Directory -Path $LogDirectory -Force -WhatIf:$false -Confirm:$false | Out-Null
}
$ScriptName = $MyInvocation.MyCommand.Name
$LogFile    = Join-Path $LogDirectory ("{0}_{1}.log" -f [IO.Path]::GetFileNameWithoutExtension($ScriptName), (Get-Date -Format 'yyyyMMdd_HHmmss'))

function Write-Log {
<#
.SYNOPSIS
    Writes a timestamped, level-tagged line to both the console and the log file.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','DEBUG')][string]$Level = 'INFO'
    )
    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    # Opt the logger out of -WhatIf/-Confirm propagated from the script's
    # SupportsShouldProcess binding. Logging is a side-effect we always want
    # to record, even during a WhatIf preview run.
    Add-Content -LiteralPath $LogFile -Value $line -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false
    $color = switch ($Level) {
        'WARN'    { 'Yellow' }
        'ERROR'   { 'Red' }
        'SUCCESS' { 'Green' }
        'DEBUG'   { 'DarkGray' }
        default   { 'Gray' }
    }
    Write-Host $line -ForegroundColor $color
}

# -----------------------------------------------------------------------------
# DISCOVERY
# -----------------------------------------------------------------------------
$GuidRegex = '^\{[0-9A-Fa-f]{8}-([0-9A-Fa-f]{4}-){3}[0-9A-Fa-f]{12}\}$'

function Get-InstalledMsiProduct {
<#
.SYNOPSIS
    Enumerates MSI products from the Uninstall registry hive(s).
.DESCRIPTION
    Reads HKLM (64-bit), HKLM\WOW6432Node (32-bit on x64 hosts), and -- when
    -IncludePerUser is supplied -- HKCU. Returns one object per MSI product
    whose subkey name is a {GUID}, with WindowsInstaller=1 (or msiexec-based
    UninstallString), non-empty DisplayName, SystemComponent != 1, and no
    ParentKeyName (excludes MSI patches and sub-components).
.PARAMETER Architecture
    Both | X64 | X86. Limits which HKLM views are scanned.
.PARAMETER IncludePerUser
    Also scan HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall.
.OUTPUTS
    PSCustomObject with DisplayName, DisplayVersion, Publisher, ProductCode,
    UninstallString, Architecture, Scope, RegistryPath.
#>
    [CmdletBinding()]
    param(
        [ValidateSet('Both','X64','X86')][string]$Architecture = 'Both',
        [switch]$IncludePerUser
    )

    # Build hive list based on architecture
    $hives = New-Object System.Collections.Generic.List[object]
    if ($Architecture -in @('Both','X64')) {
        $hives.Add([pscustomobject]@{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall';            Arch = 'X64';     Scope = 'Machine' })
    }
    if ($Architecture -in @('Both','X86')) {
        $hives.Add([pscustomobject]@{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'; Arch = 'X86';     Scope = 'Machine' })
    }
    if ($IncludePerUser) {
        # HKCU has only a single view (the OS resolves the WOW6432Node redirection per process)
        $hives.Add([pscustomobject]@{ Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall';             Arch = 'Native';  Scope = 'PerUser' })
    }

    foreach ($hive in $hives) {
        if (-not (Test-Path -LiteralPath $hive.Path)) { continue }
        $subkeys = Get-ChildItem -LiteralPath $hive.Path -ErrorAction SilentlyContinue
        foreach ($sk in $subkeys) {
            $keyName = $sk.PSChildName
            if ($keyName -notmatch $GuidRegex) { continue }   # MSI products are keyed by ProductCode GUID

            $props = $null
            try {
                $props = Get-ItemProperty -LiteralPath $sk.PSPath -ErrorAction Stop
            } catch {
                continue
            }

            # MSI gate: WindowsInstaller=1 OR UninstallString begins with msiexec
            $uninst = [string]$props.UninstallString
            $isMsi  = ($props.WindowsInstaller -eq 1) -or ($uninst -match '(?i)^\s*"?[^"]*msiexec(\.exe)?"?\s')
            if (-not $isMsi) { continue }

            # Skip patches / sub-components / hidden servicing rows
            if ($props.SystemComponent -eq 1) { continue }
            if ($props.PSObject.Properties.Name -contains 'ParentKeyName' -and $props.ParentKeyName) { continue }
            if ([string]::IsNullOrWhiteSpace($props.DisplayName)) { continue }

            [pscustomobject]@{
                DisplayName     = [string]$props.DisplayName
                DisplayVersion  = [string]$props.DisplayVersion
                Publisher       = [string]$props.Publisher
                ProductCode     = $keyName
                UninstallString = $uninst
                Architecture    = $hive.Arch
                Scope           = $hive.Scope
                RegistryPath    = $sk.PSPath
            }
        }
    }
}

function Test-WildcardMatch {
<#
.SYNOPSIS
    Returns $true when $Value matches any of the supplied wildcard patterns
    (case-insensitive). When -Patterns is null/empty, returns $true.
#>
    param([string]$Value, [string[]]$Patterns)
    if ($null -eq $Patterns -or $Patterns.Count -eq 0) { return $true }
    foreach ($p in $Patterns) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        if ($Value -like $p) { return $true }
    }
    return $false
}

function ConvertTo-NormalizedGuid {
<#
.SYNOPSIS
    Returns the input GUID string in '{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}'
    uppercase form, or $null when the input is not a valid GUID.
#>
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $trimmed = $Value.Trim().Trim('{','}')
    $g = [Guid]::Empty
    if ([Guid]::TryParse($trimmed, [ref]$g)) {
        return '{' + $g.ToString().ToUpperInvariant() + '}'
    }
    return $null
}

# -----------------------------------------------------------------------------
# EXECUTION
# -----------------------------------------------------------------------------
function Get-MsiUninstallArgList {
<#
.SYNOPSIS
    Builds the msiexec.exe argument array for an uninstall: /x {GUID} /qn
    /norestart /l*v "<log>" plus any caller-supplied AdditionalArguments
    (whitespace-only entries dropped). Returns a [string[]] suitable for
    Start-Process -ArgumentList.
#>
    param(
        [Parameter(Mandatory)][string]$ProductCode,
        [Parameter(Mandatory)][string]$MsiLogPath,
        [string[]]$AdditionalArguments = @()
    )
    $argList = @(
        '/x', $ProductCode,
        '/qn',
        '/norestart',
        '/l*v', ('"{0}"' -f $MsiLogPath)
    )
    if ($AdditionalArguments -and $AdditionalArguments.Count -gt 0) {
        $extra = $AdditionalArguments | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        if ($extra) { $argList += $extra }
    }
    return ,[string[]]$argList
}

function Invoke-MsiUninstall {
<#
.SYNOPSIS
    Invokes msiexec.exe with a pre-built argument array and waits up to
    -TimeoutSeconds for completion. Returns the process exit code (or 1460
    on timeout, 1603 on process-start failure).
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$ArgumentList,
        [int]$TimeoutSeconds = 900
    )

    Write-Log ("  cmd: msiexec.exe {0}" -f ($ArgumentList -join ' ')) 'DEBUG'

    try {
        $proc = Start-Process -FilePath 'msiexec.exe' -ArgumentList $ArgumentList -PassThru -WindowStyle Hidden -ErrorAction Stop
    } catch {
        Write-Log "  msiexec failed to start: $($_.Exception.Message)" 'ERROR'
        return 1603
    }

    if ($TimeoutSeconds -le 0) {
        $proc.WaitForExit()
    } else {
        if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
            Write-Log "  msiexec exceeded ${TimeoutSeconds}s timeout -- killing PID $($proc.Id)" 'WARN'
            try { $proc.Kill() } catch { }
            try { $proc.WaitForExit(5000) | Out-Null } catch { }
            return 1460   # ERROR_TIMEOUT
        }
    }
    return $proc.ExitCode
}

function Get-MsiExitCodeMeaning {
<#
.SYNOPSIS
    Returns a (Action, Friendly) pair for a given msiexec exit code.
    Action is one of: Removed | RebootPending | NotInstalled | Failed.
#>
    param([int]$ExitCode)
    switch ($ExitCode) {
        0     { return @{ Action = 'Removed';        Friendly = 'Success' } }
        1605  { return @{ Action = 'NotInstalled';   Friendly = 'Product not installed (already removed)' } }
        1641  { return @{ Action = 'RebootPending';  Friendly = 'Success - reboot initiated by installer' } }
        3010  { return @{ Action = 'RebootPending';  Friendly = 'Success - reboot required' } }
        1460  { return @{ Action = 'Failed';         Friendly = 'Timed out and was terminated' } }
        1618  { return @{ Action = 'Failed';         Friendly = 'Another install is already in progress' } }
        1619  { return @{ Action = 'Failed';         Friendly = 'MSI package could not be opened' } }
        1620  { return @{ Action = 'Failed';         Friendly = 'MSI package is invalid' } }
        1624  { return @{ Action = 'Failed';         Friendly = 'Transform application failed' } }
        1625  { return @{ Action = 'Failed';         Friendly = 'Uninstall forbidden by system policy' } }
        1633  { return @{ Action = 'Failed';         Friendly = 'Package not supported on this platform' } }
        default { return @{ Action = 'Failed';       Friendly = "msiexec returned $ExitCode" } }
    }
}

# -----------------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------------
function Format-FilterValue {
<#
.SYNOPSIS
    Renders a filter parameter value for the startup banner. Returns '(none)'
    for null / empty / whitespace input so empty filters don't render as a
    blank right-hand side.
#>
    param($Value, [string]$Separator = ', ')
    if ($null -eq $Value) { return '(none)' }
    if ($Value -is [System.Array]) {
        $clean = $Value | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }
        if (-not $clean -or $clean.Count -eq 0) { return '(none)' }
        return ($clean -join $Separator)
    }
    $s = [string]$Value
    if ([string]::IsNullOrWhiteSpace($s)) { return '(none)' }
    return $s
}

Write-Log "=============================================================" 'INFO'
Write-Log "Uninstall-MsiProduct starting"                                  'INFO'
Write-Log "  Host:                $env:COMPUTERNAME ($env:USERNAME)"       'INFO'
Write-Log "  OS:                  $([System.Environment]::OSVersion.VersionString)" 'INFO'
Write-Log "  PowerShell:          $($PSVersionTable.PSVersion)"            'INFO'
Write-Log "  LogFile:             $LogFile"                                'INFO'
Write-Log "  Architecture:        $Architecture (IncludePerUser=$IncludePerUser)" 'INFO'
Write-Log "  DisplayName:         $(Format-FilterValue $DisplayName)"      'INFO'
Write-Log "  Publisher:           $(Format-FilterValue $Publisher)"        'INFO'
Write-Log "  Version:             $(Format-FilterValue $Version)"          'INFO'
Write-Log "  ProductCode:         $(Format-FilterValue $ProductCode)"      'INFO'
Write-Log "  MaxMatches:          $MaxMatches (Force=$Force)"              'INFO'
Write-Log "  Timeout:             ${TimeoutSeconds}s"                      'INFO'
Write-Log "  AdditionalArguments: $(Format-FilterValue $AdditionalArguments ' ')" 'INFO'
Write-Log "============================================================="  'INFO'

# Validate that the caller supplied at least one filter (refuse to act on "everything")
$haveFilter = (
    ($DisplayName -and $DisplayName.Count -gt 0) -or
    -not [string]::IsNullOrWhiteSpace($Publisher) -or
    -not [string]::IsNullOrWhiteSpace($Version) -or
    ($ProductCode -and $ProductCode.Count -gt 0)
)
if (-not $haveFilter) {
    Write-Log "No filter supplied -- refusing to match every installed product. Use -DisplayName, -Publisher, -Version, or -ProductCode." 'ERROR'
    exit 1605
}

# ----- Build target list -----
$targets = New-Object System.Collections.Generic.List[object]

if ($ProductCode -and $ProductCode.Count -gt 0) {
    # Explicit GUID mode: bypass the filter, look the product up only for metadata
    $installed = @(Get-InstalledMsiProduct -Architecture $Architecture -IncludePerUser:$IncludePerUser)
    foreach ($raw in $ProductCode) {
        $norm = ConvertTo-NormalizedGuid -Value $raw
        if (-not $norm) {
            Write-Log "Skipping invalid ProductCode: '$raw'" 'WARN'
            continue
        }
        $match = $installed | Where-Object { $_.ProductCode.ToUpperInvariant() -eq $norm } | Select-Object -First 1
        if ($null -eq $match) {
            # Not in registry -- still allow msiexec /x to run (it will return 1605 = NotInstalled)
            $match = [pscustomobject]@{
                DisplayName     = '<unknown - not in registry>'
                DisplayVersion  = ''
                Publisher       = ''
                ProductCode     = $norm
                UninstallString = ''
                Architecture    = 'Unknown'
                Scope           = 'Machine'
                RegistryPath    = ''
            }
        }
        $targets.Add($match)
    }
} else {
    $installed = @(Get-InstalledMsiProduct -Architecture $Architecture -IncludePerUser:$IncludePerUser)
    Write-Log ("Discovered {0} MSI product(s) in scope" -f $installed.Count) 'INFO'

    foreach ($p in $installed) {
        if (-not (Test-WildcardMatch -Value $p.DisplayName    -Patterns $DisplayName))         { continue }
        if (-not [string]::IsNullOrWhiteSpace($Publisher) -and -not ($p.Publisher      -like $Publisher)) { continue }
        if (-not [string]::IsNullOrWhiteSpace($Version)   -and -not ($p.DisplayVersion -like $Version))   { continue }
        $targets.Add($p)
    }
}

# ----- Report what we found -----
if ($targets.Count -eq 0) {
    Write-Log "No matching products found. Nothing to do." 'WARN'
    exit 1605
}

Write-Log ("Matched {0} product(s):" -f $targets.Count) 'INFO'
$i = 0
foreach ($t in $targets) {
    $i++
    $pub = if ([string]::IsNullOrWhiteSpace($t.Publisher))      { '(no publisher)' } else { $t.Publisher }
    $ver = if ([string]::IsNullOrWhiteSpace($t.DisplayVersion)) { '(no version)'   } else { $t.DisplayVersion }
    # One field per line, fixed-width labels - wrap-proof on narrow consoles
    # and removes the ambiguity where a long single line would wrap and the
    # publisher could appear orphaned next to the next match's bullet.
    Write-Log ("  [{0}] {1}" -f $i, $t.DisplayName) 'INFO'
    Write-Log ("       Version     : $ver")         'INFO'
    Write-Log ("       Publisher   : $pub")         'INFO'
    Write-Log ("       ProductCode : $($t.ProductCode)") 'INFO'
    Write-Log ("       Scope       : $($t.Architecture)/$($t.Scope)") 'INFO'
}

# Safety gate: refuse big blast radius without -Force
if ($targets.Count -gt $MaxMatches -and -not $Force) {
    Write-Log "Match count ($($targets.Count)) exceeds -MaxMatches ($MaxMatches). Refine filter or pass -Force." 'ERROR'
    exit 1604
}

# ----- WhatIf / ListOnly short-circuit -----
if ($ListOnly) {
    Write-Log "ListOnly: returning discovery only -- no msiexec invoked." 'INFO'
    $targets | ForEach-Object {
        [pscustomobject]@{
            DisplayName     = $_.DisplayName
            DisplayVersion  = $_.DisplayVersion
            Publisher       = $_.Publisher
            ProductCode     = $_.ProductCode
            Architecture    = $_.Architecture
            Scope           = $_.Scope
            ExitCode        = $null
            Action          = 'ListOnly'
            DurationSeconds = 0
            MsiLog          = $null
        }
    }
    exit 0
}

# ----- Uninstall loop -----
$results = New-Object System.Collections.Generic.List[object]
$anyFailed         = $false
$anyRebootPending  = $false
$anyConfirmed      = $false

foreach ($t in $targets) {
    $label = "{0} {1} {2}" -f $t.DisplayName, $t.DisplayVersion, $t.ProductCode

    # Build msi log path + full msiexec command line BEFORE ShouldProcess so we
    # can surface the exact command in the WhatIf / Confirm prompts.
    $safeName = ($t.DisplayName -replace '[^A-Za-z0-9._-]+','_').Trim('_')
    if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = 'product' }
    $msiLog   = Join-Path $LogDirectory ("Uninstall-MsiProduct_{0}_{1}_{2}.msi.log" -f $safeName, $t.ProductCode.Trim('{','}'), (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $argList  = Get-MsiUninstallArgList -ProductCode $t.ProductCode -MsiLogPath $msiLog -AdditionalArguments $AdditionalArguments
    $cmdLine  = 'msiexec.exe ' + ($argList -join ' ')

    # 3-arg ShouldProcess overload:
    #   $verboseDescription -> WhatIf prints this verbatim ("What if: <description>")
    #   $verboseWarning     -> Shown for -Confirm prompts
    #   $caption            -> Title for the Confirm prompt
    $verboseDescription = "Uninstall '$($t.DisplayName)' [$($t.ProductCode), $($t.Architecture)/$($t.Scope)] -> $cmdLine"
    $verboseWarning     = "Run msiexec to uninstall '$($t.DisplayName) $($t.DisplayVersion)' ($($t.ProductCode))? Command: $cmdLine"
    $caption            = 'Confirm MSI uninstall'

    if (-not $PSCmdlet.ShouldProcess($verboseDescription, $verboseWarning, $caption)) {
        Write-Log "  SKIPPED (WhatIf / declined): $label" 'WARN'
        $results.Add([pscustomobject]@{
            DisplayName     = $t.DisplayName
            DisplayVersion  = $t.DisplayVersion
            Publisher       = $t.Publisher
            ProductCode     = $t.ProductCode
            Architecture    = $t.Architecture
            Scope           = $t.Scope
            ExitCode        = $null
            Action          = if ($WhatIfPreference) { 'WhatIf' } else { 'Skipped' }
            DurationSeconds = 0
            MsiLog          = $null
        })
        continue
    }
    $anyConfirmed = $true

    Write-Log "Uninstalling: $label" 'INFO'

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $rc = Invoke-MsiUninstall -ArgumentList $argList -TimeoutSeconds $TimeoutSeconds
    $sw.Stop()

    $meaning = Get-MsiExitCodeMeaning -ExitCode $rc
    $level   = switch ($meaning.Action) {
        'Removed'       { 'SUCCESS' }
        'RebootPending' { 'SUCCESS' }
        'NotInstalled'  { 'WARN' }
        default         { 'ERROR' }
    }
    Write-Log ("  -> exit {0} ({1}) in {2:N1}s" -f $rc, $meaning.Friendly, $sw.Elapsed.TotalSeconds) $level
    Write-Log ("  -> MSI log: $msiLog") 'DEBUG'

    if ($meaning.Action -eq 'Failed')        { $anyFailed        = $true }
    if ($meaning.Action -eq 'RebootPending') { $anyRebootPending = $true }

    $results.Add([pscustomobject]@{
        DisplayName     = $t.DisplayName
        DisplayVersion  = $t.DisplayVersion
        Publisher       = $t.Publisher
        ProductCode     = $t.ProductCode
        Architecture    = $t.Architecture
        Scope           = $t.Scope
        ExitCode        = $rc
        Action          = $meaning.Action
        DurationSeconds = [math]::Round($sw.Elapsed.TotalSeconds, 1)
        MsiLog          = $msiLog
    })
}

# Emit objects on the pipeline for callers
$results

# ----- Final exit code -----
if (-not $anyConfirmed -and $WhatIfPreference) {
    Write-Log "Run complete (WhatIf preview only)." 'INFO'
    exit 0
}
if (-not $anyConfirmed) {
    Write-Log "All targets declined at confirmation prompt." 'WARN'
    exit 1602
}
if ($anyFailed) {
    Write-Log "One or more uninstalls FAILED. See per-product MSI logs in $LogDirectory." 'ERROR'
    exit 1603
}
if ($anyRebootPending) {
    Write-Log "All uninstalls succeeded -- REBOOT REQUIRED." 'WARN'
    exit 3010
}
Write-Log "All uninstalls completed successfully." 'SUCCESS'
exit 0
