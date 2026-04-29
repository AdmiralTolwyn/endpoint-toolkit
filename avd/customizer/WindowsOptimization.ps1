<#
.SYNOPSIS
    Hardened AVD / Windows 365 image-bake optimization wrapper around the Virtual
    Desktop Optimization Tool (VDOT) configuration JSON files. Selectively disables
    tasks, services, autologgers, and applies registry / LGPO tweaks for VDI workloads
    with resilient access-denied handling and verbose logging.

.DESCRIPTION
    Inspired by The Virtual Desktop Team's Windows_VDOT.ps1:
        https://github.com/The-Virtual-Desktop-Team/Virtual-Desktop-Optimization-Tool

    This script downloads (or reads from -ConfigBasePath) the upstream VDOT JSON
    configuration files for the running Windows release and applies the categories
    selected via -Optimizations.

    Hardening over the legacy version:
      * `$ErrorActionPreference = 'Stop'` plus per-operation try/catch so a single
        access-denied no longer corrupts the stopwatch / logging tail.
      * Uniform Write-Log helper writes timestamped, level-tagged lines to
        Write-Host (AIB / Packer pickup) AND a file under -LogDirectory.
      * Disk cleanup uses targeted folder lists instead of `Get-ChildItem c:\ -Recurse`
        so it no longer floods the log with access-denied to System Volume
        Information / $Recycle.Bin / Config.Msi / per-user profiles it cannot
        traverse.
      * Service / scheduled-task disablers check for existence first and treat
        Win32 access-denied (5 / 1314) as a WARN, not a hard error.
      * Default-user hive load/unload is bracketed by reg.exe exit-code checks; on
        load failure the section is skipped instead of half-applied.
      * Working folder defaults to `$env:TEMP\WindowsOptimization_<ts>` (always
         writable for SYSTEM) instead of `$PSScriptRoot\<version>` which is often
         read-only under AIB.
      * Optional -ConfigBasePath supports fully offline / air-gapped bakes by
        reading the same JSON files from a pre-staged local folder.

    Categories:
      WindowsMediaPlayer    - Disable + remove Windows Media Player payload.
      ScheduledTasks        - Disable VDI-hostile scheduled tasks.
      DefaultUserSettings   - Apply registry tweaks to the Default user hive.
      AutoLoggers           - Disable Windows trace autologgers.
      Services              - Disable VDI-hostile services.
      NetworkOptimizations  - Apply LanManWorkstation tunings.
      LGPO                  - Apply local group policy registry settings.
      DiskCleanup           - Sweep TEMP / WER / BranchCache / Recycle Bin.
      Edge                  - Apply Edge browser policy registry settings.
      RemoveLegacyIE        - Remove the Internet Explorer Windows Capability.
      RemoveOneDrive        - Uninstall OneDrive (per-user + per-machine setup) and
                              prune residual shortcuts.
      All                   - Apply every category above.

.PARAMETER Optimizations
    One or more category names from the validated set. Use 'All' to apply every
    category.

.PARAMETER ConfigBasePath
    Optional local folder containing pre-staged VDOT JSON files. When supplied the
    script reads each ConfigurationFiles\*.json from this folder INSTEAD of
    downloading from raw.githubusercontent.com. Layout expected:

        <ConfigBasePath>\
            ScheduledTasks.json
            DefaultUserSettings.json
            Autologgers.json
            Services.json
            LanManWorkstation.json
            PolicyRegSettings.json
            EdgeSettings.json

    Use this for air-gapped image bakes.

.PARAMETER LogDirectory
    Directory for the rolling log file. Default: $env:TEMP. The log is named
    WindowsOptimization_yyyyMMdd_HHmmss.log.

.PARAMETER ContinueOnError
    Do not exit non-zero when one or more sections raise WARN/ERROR. Useful when
    the script is one of many AIB customizers and you do not want a single
    deprecated handler to abort the whole bake.

.NOTES
    File:    avd/customizer/WindowsOptimization.ps1
    Author:  Anton Romanyuk (wrapper); upstream VDOT authored by The Virtual
             Desktop Team (Microsoft community).
    Version: 3.0.0
    Context: Azure Image Builder / Packer customizer. Runs as SYSTEM. Internet
             egress to raw.githubusercontent.com required UNLESS -ConfigBasePath
             is supplied.
    Requires: Windows 10/11 / Server, PowerShell 5.1+, admin.

    Changes:
      3.0.0 - Full rewrite. Resilient access-denied handling, file logger,
              targeted disk-sweep paths, hive load/unload checked, -ConfigBasePath
              and -ContinueOnError added, fixed Write-Host -EventId garbage call,
              ReleaseId fallback to DisplayVersion / CurrentBuild.
      2.0.0 - Header refresh, #Requires, exit 1 on bail-out.

.DISCLAIMER
    This script is provided "AS IS" with no warranties and confers no rights.
    It is not supported under any Microsoft standard support program or service.
    Use of this script is entirely at your own risk. The customer is solely
    responsible for testing and validating this script in their environment
    before deploying to production. VDOT applies broad, irreversible tweaks -
    test in a non-production image FIRST.

.EXAMPLE
    .\WindowsOptimization.ps1 -Optimizations All

.EXAMPLE
    .\WindowsOptimization.ps1 -Optimizations ScheduledTasks,Services,DiskCleanup

.EXAMPLE
    # Air-gapped bake with pre-staged JSON
    .\WindowsOptimization.ps1 -Optimizations All -ConfigBasePath C:\BuildArtifacts\VDOT
#>

#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet('All','WindowsMediaPlayer','ScheduledTasks','DefaultUserSettings',
                 'Autologgers','Services','NetworkOptimizations','LGPO','DiskCleanup',
                 'Edge','RemoveLegacyIE','RemoveOneDrive')]
    [string[]]$Optimizations,

    [string]$ConfigBasePath,

    [string]$LogDirectory = $env:TEMP,

    [switch]$ContinueOnError
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

# -----------------------------------------------------------------------------
# Module-scope state
# -----------------------------------------------------------------------------
$Script:ScriptName = 'WindowsOptimization'
$Script:LogFile    = Join-Path $LogDirectory ("{0}_{1}.log" -f $Script:ScriptName, (Get-Date -Format 'yyyyMMdd_HHmmss'))
$Script:ErrorCount = 0
$Script:WarnCount  = 0

# Upstream VDOT raw URLs (kept centralised for easy mirror/swap).
$VdotBaseUrl = 'https://raw.githubusercontent.com/The-Virtual-Desktop-Team/Virtual-Desktop-Optimization-Tool/main/2009/ConfigurationFiles'
$VdotConfigs = @{
    ScheduledTasks       = 'ScheduledTasks.json'
    DefaultUserSettings  = 'DefaultUserSettings.json'
    AutoLoggers          = 'Autologgers.Json'
    Services             = 'Services.json'
    NetworkOptimizations = 'LanManWorkstation.json'
    LGPO                 = 'PolicyRegSettings.json'
    Edge                 = 'EdgeSettings.json'
}

# -----------------------------------------------------------------------------
# LOGGING
# -----------------------------------------------------------------------------
function Write-Log {
<#
.SYNOPSIS
    Writes a colour-coded, timestamped, level-tagged line to console + log file.
.DESCRIPTION
    Format: [yyyy-MM-dd HH:mm:ss] [LEVEL] [WindowsOptimization] message.
    File appends use SilentlyContinue so a transient lock never aborts the bake.
    Tracks WARN/ERROR counts in module scope so the final summary can flag the
    bake without scraping logs.
.PARAMETER Message
    Free-form text.
.PARAMETER Level
    INFO | WARN | ERROR | SUCCESS | HEADER. Default INFO.
#>
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','HEADER')][string]$Level = 'INFO'
    )
    $ts   = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[$ts] [$Level] [$($Script:ScriptName)] $Message"

    Add-Content -LiteralPath $Script:LogFile -Value $line -ErrorAction SilentlyContinue

    switch ($Level) {
        'WARN'    { $Script:WarnCount++  ; $color = 'Yellow' }
        'ERROR'   { $Script:ErrorCount++ ; $color = 'Red'    }
        'SUCCESS' { $color = 'Green'  }
        'HEADER'  { $color = 'Cyan'   }
        default   { $color = 'Gray'   }
    }
    Write-Host $line -ForegroundColor $color
}

function Invoke-Section {
<#
.SYNOPSIS
    Runs a named optimization category and converts unhandled exceptions into a single
    ERROR log line so one bad section never blows up the whole script.
.PARAMETER Name
    Friendly section name used in HEADER log lines.
.PARAMETER ScriptBlock
    The body to execute.
#>
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock
    )
    Write-Log "===== $Name =====" -Level HEADER
    try {
        & $ScriptBlock
    }
    catch {
        Write-Log "Section '$Name' raised an unhandled exception: $($_.Exception.Message)" -Level ERROR
        Write-Log $_.ScriptStackTrace -Level ERROR
    }
}

# -----------------------------------------------------------------------------
# CONFIG SOURCING (download or local)
# -----------------------------------------------------------------------------
function Get-VdotConfig {
<#
.SYNOPSIS
    Returns a parsed JSON config for one VDOT category, sourced either from
    -ConfigBasePath or from raw.githubusercontent.com.
.PARAMETER Key
    Key into the $VdotConfigs hashtable (ScheduledTasks, DefaultUserSettings, etc).
.OUTPUTS
    Parsed object array, or $null if the file is missing / unreadable.
#>
    param([Parameter(Mandatory)][string]$Key)

    $fileName = $VdotConfigs[$Key]
    if (-not $fileName) {
        Write-Log "Get-VdotConfig: unknown key '$Key'" -Level ERROR
        return $null
    }

    $localCopy = Join-Path $Script:WorkingFolder $fileName

    if ($ConfigBasePath) {
        $src = Join-Path $ConfigBasePath $fileName
        if (-not (Test-Path -LiteralPath $src)) {
            Write-Log "Local config '$src' not found (-ConfigBasePath supplied) - skipping $Key" -Level WARN
            return $null
        }
        try {
            Copy-Item -LiteralPath $src -Destination $localCopy -Force
        }
        catch {
            Write-Log "Failed to stage local config '$src': $($_.Exception.Message)" -Level WARN
            return $null
        }
    }
    else {
        $url = "$VdotBaseUrl/$fileName"
        try {
            Write-Log "Downloading $Key config: $url"
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
            Invoke-WebRequest -Uri $url -OutFile $localCopy -UseBasicParsing -ErrorAction Stop
        }
        catch {
            Write-Log "Failed to download '$url': $($_.Exception.Message)" -Level WARN
            return $null
        }
    }

    try {
        return (Get-Content -LiteralPath $localCopy -Raw -ErrorAction Stop | ConvertFrom-Json)
    }
    catch {
        Write-Log "Failed to parse '$localCopy': $($_.Exception.Message)" -Level WARN
        return $null
    }
}

# -----------------------------------------------------------------------------
# SAFE HELPERS
# -----------------------------------------------------------------------------
function Set-RegValueSafe {
<#
.SYNOPSIS
    Idempotently writes a registry value, creating the parent key tree if absent,
    converting access-denied / missing-path into a single WARN log line.
.PARAMETER Path
    Full registry path.
.PARAMETER Name
    Value name.
.PARAMETER Value
    Value to set.
.PARAMETER Type
    Registry value type (default DWord).
#>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)]$Value,
        [Microsoft.Win32.RegistryValueKind]$Type = [Microsoft.Win32.RegistryValueKind]::DWord
    )
    try {
        if (-not (Test-Path -LiteralPath $Path)) {
            New-Item -Path $Path -Force | Out-Null
        }
        New-ItemProperty -LiteralPath $Path -Name $Name -Value $Value -PropertyType $Type -Force | Out-Null
    }
    catch [System.UnauthorizedAccessException] {
        Write-Log "Access denied writing $Path\$Name (skipping)" -Level WARN
    }
    catch {
        Write-Log "Failed to write $Path\$Name : $($_.Exception.Message)" -Level WARN
    }
}

function Disable-ScheduledTaskSafe {
<#
.SYNOPSIS
    Disables a scheduled task by name, treating not-found / already-disabled / access
    denied as informational rather than as failures.
.PARAMETER TaskName
    Task identifier as seen by Get-ScheduledTask (full path or leaf name).
#>
    param([Parameter(Mandatory)][string]$TaskName)

    try {
        $leaf = Split-Path $TaskName -Leaf
        $task = Get-ScheduledTask -TaskName $leaf -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $task) {
            $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        if (-not $task) {
            Write-Log "Task not found: $TaskName"
            return
        }
        if ($task.State -eq 'Disabled') {
            Write-Log "Task already disabled: $($task.TaskPath)$($task.TaskName)"
            return
        }
        Disable-ScheduledTask -InputObject $task -ErrorAction Stop | Out-Null
        Write-Log "Disabled task: $($task.TaskPath)$($task.TaskName)" -Level SUCCESS
    }
    catch [Microsoft.Management.Infrastructure.CimException] {
        Write-Log "CIM error disabling '$TaskName': $($_.Exception.Message)" -Level WARN
    }
    catch {
        Write-Log "Failed to disable '$TaskName': $($_.Exception.Message)" -Level WARN
    }
}

function Set-ServiceStartupSafe {
<#
.SYNOPSIS
    Sets the startup type of a service if it exists; treats access-denied (e.g. SYSTEM
    -protected services) as WARN.
.PARAMETER Name
    Service name.
.PARAMETER StartupType
    Disabled | Manual | Automatic | AutomaticDelayedStart. Default Disabled.
#>
    param(
        [Parameter(Mandatory)][string]$Name,
        [ValidateSet('Disabled','Manual','Automatic','AutomaticDelayedStart')]
        [string]$StartupType = 'Disabled'
    )
    $svc = Get-Service -Name $Name -ErrorAction SilentlyContinue
    if (-not $svc) {
        Write-Log "Service not present: $Name"
        return
    }
    try {
        Set-Service -Name $Name -StartupType $StartupType -ErrorAction Stop
        Write-Log "Set service '$Name' startup -> $StartupType" -Level SUCCESS
    }
    catch [System.ComponentModel.Win32Exception] {
        # 5 = Access denied, 1060 = service not exist, 1072 = marked for delete
        Write-Log "Win32 error on '$Name' (code $($_.Exception.NativeErrorCode)): $($_.Exception.Message)" -Level WARN
    }
    catch {
        Write-Log "Failed to set startup on '$Name': $($_.Exception.Message)" -Level WARN
    }
}

function Remove-PathSafe {
<#
.SYNOPSIS
    Recursively removes the contents of $Path, swallowing per-file access-denied
    so a single locked handle never aborts the sweep.
.PARAMETER Path
    Folder whose contents should be wiped (the folder itself is preserved).
.PARAMETER Include
    Optional wildcard filter applied to file names.
#>
    param(
        [Parameter(Mandatory)][string]$Path,
        [string[]]$Include
    )
    if (-not (Test-Path -LiteralPath $Path)) { return }
    try {
        $items = if ($Include) {
            Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue -Include $Include
        } else {
            Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue
        }
        $count = 0
        foreach ($it in $items) {
            try {
                Remove-Item -LiteralPath $it.FullName -Recurse -Force -ErrorAction Stop
                $count++
            }
            catch [System.UnauthorizedAccessException] { }
            catch [System.IO.IOException]            { }   # in-use file
            catch                                    { }
        }
        if ($count -gt 0) { Write-Log "  cleaned $count item(s) from $Path" }
    }
    catch {
        Write-Log "Sweep '$Path' raised: $($_.Exception.Message)" -Level WARN
    }
}

function Convert-PropertyValue {
<#
.SYNOPSIS
    Converts a VDOT JSON PropertyValue to the right .NET type for New-/Set-ItemProperty.
.DESCRIPTION
    BINARY values arrive as comma-separated strings. Everything else passes through
    unchanged.
.PARAMETER Item
    A single VDOT setting object (must expose .PropertyType / .PropertyValue).
#>
    param([Parameter(Mandatory)]$Item)
    if ($Item.PropertyType -ieq 'BINARY' -and $Item.PropertyValue -is [string]) {
        return ,([byte[]]($Item.PropertyValue.Split(',') | ForEach-Object { [byte]$_ }))
    }
    return $Item.PropertyValue
}

# -----------------------------------------------------------------------------
# WORKING FOLDER (writable for SYSTEM)
# -----------------------------------------------------------------------------
$Script:WorkingFolder = Join-Path $env:TEMP ("WindowsOptimization_{0}" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
try {
    New-Item -Path $Script:WorkingFolder -ItemType Directory -Force | Out-Null
}
catch {
    Write-Log "Cannot create working folder '$Script:WorkingFolder': $($_.Exception.Message)" -Level ERROR
    exit 1
}

$AllStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
Write-Log "===== WindowsOptimization v3.0.0 starting =====" -Level HEADER
Write-Log "Categories      : $($Optimizations -join ', ')"
Write-Log "Working folder  : $Script:WorkingFolder"
Write-Log "Log file        : $Script:LogFile"
Write-Log "Config source   : $(if ($ConfigBasePath) { "local '$ConfigBasePath'" } else { 'GitHub raw' })"

# Detect Windows version (ReleaseId is missing on Win11 21H2+, so fall back).
try {
    $cv = Get-ItemProperty -LiteralPath 'HKLM:\Software\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
    $WinVersion = $cv.DisplayVersion
    if (-not $WinVersion) { $WinVersion = $cv.ReleaseId }
    if (-not $WinVersion) { $WinVersion = $cv.CurrentBuild }
    Write-Log "Windows version: $WinVersion (build $($cv.CurrentBuild))"
}
catch {
    Write-Log "Could not read Windows version metadata: $($_.Exception.Message)" -Level WARN
}

$All = $Optimizations -contains 'All'

# =============================================================================
# 1. WINDOWS MEDIA PLAYER
# =============================================================================
if ($All -or $Optimizations -contains 'WindowsMediaPlayer') {
    Invoke-Section 'Windows Media Player' {
        try {
            Disable-WindowsOptionalFeature -Online -FeatureName WindowsMediaPlayer -NoRestart -ErrorAction Stop | Out-Null
            Write-Log "Disabled WindowsMediaPlayer optional feature" -Level SUCCESS
        }
        catch {
            Write-Log "Disable-WindowsOptionalFeature WindowsMediaPlayer: $($_.Exception.Message)" -Level WARN
        }
        try {
            Get-WindowsPackage -Online -ErrorAction Stop | Where-Object PackageName -Like '*Windows-mediaplayer*' | ForEach-Object {
                try {
                    Remove-WindowsPackage -PackageName $_.PackageName -Online -NoRestart -ErrorAction Stop | Out-Null
                    Write-Log "Removed package $($_.PackageName)" -Level SUCCESS
                } catch {
                    Write-Log "Remove-WindowsPackage $($_.PackageName): $($_.Exception.Message)" -Level WARN
                }
            }
        }
        catch {
            Write-Log "Get-WindowsPackage failed: $($_.Exception.Message)" -Level WARN
        }
    }
}

# =============================================================================
# 2. SCHEDULED TASKS
# =============================================================================
if ($All -or $Optimizations -contains 'ScheduledTasks') {
    Invoke-Section 'Scheduled Tasks' {
        $cfg = Get-VdotConfig -Key 'ScheduledTasks'
        if (-not $cfg) { return }
        $list = @($cfg | Where-Object VDIState -EQ 'Disabled')
        Write-Log "Tasks selected for disable: $($list.Count)"
        foreach ($item in $list) {
            Disable-ScheduledTaskSafe -TaskName $item.ScheduledTask
        }
    }
}

# =============================================================================
# 3. DEFAULT USER SETTINGS  (mounts C:\Users\Default\NTUSER.DAT as HKLM\VDOT_TEMP)
# =============================================================================
if ($All -or $Optimizations -contains 'DefaultUserSettings') {
    Invoke-Section 'Default User Settings' {
        $cfg = Get-VdotConfig -Key 'DefaultUserSettings'
        if (-not $cfg) { return }
        $list = @($cfg | Where-Object SetProperty -EQ $true)
        if ($list.Count -eq 0) { Write-Log "No default-user settings flagged SetProperty"; return }

        $hivePath = 'C:\Users\Default\NTUSER.DAT'
        if (-not (Test-Path -LiteralPath $hivePath)) {
            Write-Log "Default user hive not found at $hivePath" -Level WARN
            return
        }

        $loadProc = Start-Process -FilePath reg.exe -ArgumentList @('LOAD','HKLM\VDOT_TEMP', $hivePath) -PassThru -Wait -NoNewWindow
        if ($loadProc.ExitCode -ne 0) {
            Write-Log "reg.exe LOAD HKLM\VDOT_TEMP failed (exit $($loadProc.ExitCode)) - skipping section" -Level ERROR
            return
        }
        try {
            foreach ($item in $list) {
                $value = Convert-PropertyValue -Item $item
                $type  = if ($item.PropertyType) { $item.PropertyType } else { 'DWord' }
                try {
                    if (-not (Test-Path -LiteralPath $item.HivePath)) {
                        New-Item -Path $item.HivePath -Force | Out-Null
                    }
                    New-ItemProperty -LiteralPath $item.HivePath -Name $item.KeyName -Value $value -PropertyType $type -Force -ErrorAction Stop | Out-Null
                    Write-Log "Set $($item.HivePath)\$($item.KeyName) = $value"
                }
                catch {
                    Write-Log "Failed to set $($item.HivePath)\$($item.KeyName): $($_.Exception.Message)" -Level WARN
                }
            }
        }
        finally {
            [gc]::Collect(); [gc]::WaitForPendingFinalizers()
            $unloadProc = Start-Process -FilePath reg.exe -ArgumentList @('UNLOAD','HKLM\VDOT_TEMP') -PassThru -Wait -NoNewWindow
            if ($unloadProc.ExitCode -ne 0) {
                Write-Log "reg.exe UNLOAD HKLM\VDOT_TEMP failed (exit $($unloadProc.ExitCode)) - default-user hive may be left mounted" -Level WARN
            }
        }
    }
}

# =============================================================================
# 4. AUTOLOGGERS
# =============================================================================
if ($All -or $Optimizations -contains 'Autologgers') {
    Invoke-Section 'AutoLoggers' {
        $cfg = Get-VdotConfig -Key 'AutoLoggers'
        if (-not $cfg) { return }
        $list = @($cfg | Where-Object Disabled -EQ 'True')
        Write-Log "AutoLoggers to disable: $($list.Count)"
        foreach ($item in $list) {
            Set-RegValueSafe -Path $item.KeyName -Name 'Start' -Value 0 -Type ([Microsoft.Win32.RegistryValueKind]::DWord)
        }
    }
}

# =============================================================================
# 5. SERVICES
# =============================================================================
if ($All -or $Optimizations -contains 'Services') {
    Invoke-Section 'Services' {
        $cfg = Get-VdotConfig -Key 'Services'
        if (-not $cfg) { return }
        $list = @($cfg | Where-Object VDIState -EQ 'Disabled')
        Write-Log "Services to disable: $($list.Count)"
        foreach ($item in $list) {
            Set-ServiceStartupSafe -Name $item.Name -StartupType Disabled
        }
    }
}

# =============================================================================
# 6. NETWORK OPTIMIZATIONS  (LanManWorkstation)
# =============================================================================
if ($All -or $Optimizations -contains 'NetworkOptimizations') {
    Invoke-Section 'Network Optimizations' {
        $cfg = Get-VdotConfig -Key 'NetworkOptimizations'
        if (-not $cfg) { return }
        foreach ($hive in @($cfg)) {
            if (-not (Test-Path -LiteralPath $hive.HivePath)) {
                Write-Log "Hive path not present: $($hive.HivePath)" -Level WARN
                continue
            }
            $keys = @($hive.Keys | Where-Object SetProperty -EQ $true)
            foreach ($key in $keys) {
                $kind = [Microsoft.Win32.RegistryValueKind]::DWord
                if ($key.PropertyType) {
                    try { $kind = [Microsoft.Win32.RegistryValueKind]::Parse([Microsoft.Win32.RegistryValueKind], $key.PropertyType, $true) } catch { }
                }
                Set-RegValueSafe -Path $hive.HivePath -Name $key.Name -Value $key.PropertyValue -Type $kind
            }
        }
    }
}

# =============================================================================
# 7. LGPO  (registry path of upstream PolicyRegSettings.json)
# =============================================================================
if ($All -or $Optimizations -contains 'LGPO') {
    Invoke-Section 'LGPO / Policy Registry Settings' {
        $cfg = Get-VdotConfig -Key 'LGPO'
        if (-not $cfg) { return }
        $list = @($cfg | Where-Object VDIState -EQ 'Enabled')
        Write-Log "Policy registry entries to apply: $($list.Count)"
        foreach ($key in $list) {
            $kind = [Microsoft.Win32.RegistryValueKind]::DWord
            if ($key.RegItemValueType) {
                try { $kind = [Microsoft.Win32.RegistryValueKind]::Parse([Microsoft.Win32.RegistryValueKind], $key.RegItemValueType, $true) } catch { }
            }
            Set-RegValueSafe -Path $key.RegItemPath -Name $key.RegItemValueName -Value $key.RegItemValue -Type $kind
        }
    }
}

# =============================================================================
# 8. EDGE
# =============================================================================
if ($All -or $Optimizations -contains 'Edge') {
    Invoke-Section 'Edge Policy Settings' {
        $cfg = Get-VdotConfig -Key 'Edge'
        if (-not $cfg) { return }
        $list = @($cfg | Where-Object VDIState -EQ 'Enabled')
        Write-Log "Edge policy entries to apply: $($list.Count)"
        foreach ($key in $list) {
            if ($key.RegItemValueName -eq 'DefaultAssociationsConfiguration') {
                # Original VDOT script copies a local XML asset here. We only support
                # this side-effect when the asset is co-located with this script.
                $assoc = Join-Path $PSScriptRoot 'ConfigurationFiles\DefaultAssociationsConfiguration.xml'
                if (Test-Path -LiteralPath $assoc) {
                    try {
                        Copy-Item -LiteralPath $assoc -Destination $key.RegItemValue -Force
                        Write-Log "Copied DefaultAssociationsConfiguration -> $($key.RegItemValue)"
                    } catch {
                        Write-Log "Copy DefaultAssociationsConfiguration: $($_.Exception.Message)" -Level WARN
                    }
                } else {
                    Write-Log "DefaultAssociationsConfiguration.xml not co-located - skipping copy"
                }
            }
            $kind = [Microsoft.Win32.RegistryValueKind]::DWord
            if ($key.RegItemValueType) {
                try { $kind = [Microsoft.Win32.RegistryValueKind]::Parse([Microsoft.Win32.RegistryValueKind], $key.RegItemValueType, $true) } catch { }
            }
            Set-RegValueSafe -Path $key.RegItemPath -Name $key.RegItemValueName -Value $key.RegItemValue -Type $kind
        }
    }
}

# =============================================================================
# 9. REMOVE LEGACY IE
# =============================================================================
if ($All -or $Optimizations -contains 'RemoveLegacyIE') {
    Invoke-Section 'Remove Legacy IE' {
        try {
            $caps = @(Get-WindowsCapability -Online -ErrorAction Stop | Where-Object Name -Like '*Browser.Internet*')
            foreach ($c in $caps) {
                try {
                    Remove-WindowsCapability -Online -Name $c.Name -ErrorAction Stop | Out-Null
                    Write-Log "Removed capability $($c.Name)" -Level SUCCESS
                } catch {
                    Write-Log "Remove-WindowsCapability $($c.Name): $($_.Exception.Message)" -Level WARN
                }
            }
            if ($caps.Count -eq 0) { Write-Log "No legacy IE capability present" }
        }
        catch {
            Write-Log "Get-WindowsCapability failed: $($_.Exception.Message)" -Level WARN
        }
    }
}

# =============================================================================
# 10. REMOVE ONEDRIVE
# =============================================================================
if ($All -or $Optimizations -contains 'RemoveOneDrive') {
    Invoke-Section 'Remove OneDrive' {
        $setups = @(
            'C:\Windows\System32\OneDriveSetup.exe'
            'C:\Windows\SysWOW64\OneDriveSetup.exe'
        )
        foreach ($setup in $setups) {
            if (-not (Test-Path -LiteralPath $setup)) { continue }
            try {
                $proc = Start-Process -FilePath $setup -ArgumentList '/uninstall' -Wait -NoNewWindow -PassThru -ErrorAction Stop
                Write-Log "$setup /uninstall exited $($proc.ExitCode)"
            }
            catch {
                Write-Log "OneDrive uninstall via $setup failed: $($_.Exception.Message)" -Level WARN
            }
        }

        # Targeted shortcut sweep (avoid Get-ChildItem c:\* -Recurse access-denied storm).
        $shortcutRoots = @(
            "$env:ProgramData\Microsoft\Windows\Start Menu\Programs"
            "$env:PUBLIC\Desktop"
            "C:\Users\Default\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"
            "C:\Users\Default\Desktop"
        )
        foreach ($root in $shortcutRoots) {
            if (-not (Test-Path -LiteralPath $root)) { continue }
            try {
                Get-ChildItem -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue -Include 'OneDrive*.lnk' |
                    ForEach-Object {
                        try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop; Write-Log "Removed shortcut $($_.FullName)" }
                        catch { Write-Log "Could not remove $($_.FullName): $($_.Exception.Message)" -Level WARN }
                    }
            }
            catch {
                Write-Log "Shortcut sweep '$root' raised: $($_.Exception.Message)" -Level WARN
            }
        }
    }
}

# =============================================================================
# 11. DISK CLEANUP  (targeted, no c:\ -Recurse)
# =============================================================================
if ($All -or $Optimizations -contains 'DiskCleanup') {
    Invoke-Section 'Disk Cleanup' {
        $sweepFolders = @(
            "$env:WinDir\Temp"
            "$env:WinDir\Logs\CBS"
            "$env:WinDir\Panther"
            "$env:WinDir\SoftwareDistribution\Download"
            "$env:WinDir\Prefetch"
            "$env:WinDir\System32\LogFiles"
            "$env:ProgramData\Microsoft\Windows\WER\Temp"
            "$env:ProgramData\Microsoft\Windows\WER\ReportArchive"
            "$env:ProgramData\Microsoft\Windows\WER\ReportQueue"
            "$env:ProgramData\Microsoft\Windows\RetailDemo"
            $env:TEMP
        )
        $extPattern = @('*.tmp','*.dmp','*.etl','*.evtx','*.log','thumbcache*.db')

        foreach ($f in $sweepFolders) {
            Remove-PathSafe -Path $f -Include $extPattern
        }

        # Empty WER + RetailDemo entirely (everything inside, not just by extension).
        Remove-PathSafe -Path "$env:ProgramData\Microsoft\Windows\WER\Temp"
        Remove-PathSafe -Path "$env:ProgramData\Microsoft\Windows\WER\ReportArchive"
        Remove-PathSafe -Path "$env:ProgramData\Microsoft\Windows\WER\ReportQueue"
        Remove-PathSafe -Path "$env:ProgramData\Microsoft\Windows\RetailDemo"

        try {
            Clear-RecycleBin -Force -ErrorAction Stop
            Write-Log "Recycle Bin cleared" -Level SUCCESS
        }
        catch {
            Write-Log "Clear-RecycleBin: $($_.Exception.Message)" -Level WARN
        }

        if (Get-Command Clear-BCCache -ErrorAction SilentlyContinue) {
            try {
                Clear-BCCache -Force -ErrorAction Stop
                Write-Log "BranchCache cache cleared" -Level SUCCESS
            }
            catch {
                Write-Log "Clear-BCCache: $($_.Exception.Message)" -Level WARN
            }
        } else {
            Write-Log "Clear-BCCache not available on this SKU - skipping"
        }
    }
}

# =============================================================================
# CLEANUP + SUMMARY
# =============================================================================
try {
    if (Test-Path -LiteralPath $Script:WorkingFolder) {
        Remove-Item -LiteralPath $Script:WorkingFolder -Recurse -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path -LiteralPath 'C:\AVDImage') {
        Remove-Item -LiteralPath 'C:\AVDImage' -Recurse -Force -ErrorAction SilentlyContinue
    }
}
catch {
    Write-Log "Final cleanup raised: $($_.Exception.Message)" -Level WARN
}

$AllStopwatch.Stop()
Write-Log "===== Summary =====" -Level HEADER
Write-Log ("Warnings: {0} | Errors: {1} | Elapsed: {2}" -f $Script:WarnCount, $Script:ErrorCount, $AllStopwatch.Elapsed) -Level $(if ($Script:ErrorCount -eq 0) { 'SUCCESS' } else { 'WARN' })

if ($Script:ErrorCount -gt 0 -and -not $ContinueOnError) {
    Write-Log "Exiting 1 due to $($Script:ErrorCount) error(s). Use -ContinueOnError to mask." -Level ERROR
    exit 1
}
exit 0
