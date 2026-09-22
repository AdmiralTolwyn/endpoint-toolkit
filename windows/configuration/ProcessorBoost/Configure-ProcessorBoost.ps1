#Requires -Version 5.1

<#
.SYNOPSIS
    Queries or changes processor performance boost on the active power plan.
.DESCRIPTION
    Queries PERFBOOSTMODE by default. Changes apply to both AC and battery
    power, require an administrator, and reapply the active plan. Other power
    plans and settings are not changed. Displays the settings after a change.
.PARAMETER Disable
    Set mode 0 (Disabled). Cannot be combined with -Enable or -Configure.
.PARAMETER Enable
    Set mode 1 (Enabled). Does not restore a previous mode such as Aggressive.
    Cannot be combined with -Disable or -Configure.
.PARAMETER Configure
    Select a mode by its index. Cannot be combined with -Disable or -Enable.
    0 = Disabled
    1 = Enabled
    2 = Aggressive
    3 = Efficient Enabled
    4 = Efficient Aggressive
    5 = Aggressive At Guaranteed
    6 = Efficient Aggressive At Guaranteed
.EXAMPLE
    .\Configure-ProcessorBoost.ps1

    Query current AC and battery settings without making changes.
.EXAMPLE
    .\Configure-ProcessorBoost.ps1 -Disable

    Disable boost for AC and battery power.
.EXAMPLE
    .\Configure-ProcessorBoost.ps1 -Enable

    Enable boost for AC and battery power using mode 1.
.EXAMPLE
    .\Configure-ProcessorBoost.ps1 -Configure 2

    Select Aggressive mode for AC and battery power.
.LINK
    https://learn.microsoft.com/en-us/windows-hardware/customize/power-settings/options-for-perf-state-engine-perfboostmode
#>
[CmdletBinding(DefaultParameterSetName = 'Query')]
param(
    [Parameter(ParameterSetName = 'Disable')]
    [switch]$Disable,

    [Parameter(ParameterSetName = 'Enable')]
    [switch]$Enable,

    [Parameter(ParameterSetName = 'Configure')]
    [ValidateRange(0, 6)]
    [int]$Configure
)

$ErrorActionPreference = 'Stop'

if ($Disable -or $Enable -or $PSBoundParameters.ContainsKey('Configure')) {
    $principal = New-Object -TypeName Security.Principal.WindowsPrincipal -ArgumentList ([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'Run PowerShell as administrator to change processor performance boost.'
    }

    $boostMode = 0
    if ($Enable) {
        $boostMode = 1
    }
    elseif ($PSBoundParameters.ContainsKey('Configure')) {
        $boostMode = $Configure
    }

    powercfg.exe /setacvalueindex SCHEME_CURRENT SUB_PROCESSOR PERFBOOSTMODE $boostMode
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to set processor boost on AC power (exit code $LASTEXITCODE)."
    }

    powercfg.exe /setdcvalueindex SCHEME_CURRENT SUB_PROCESSOR PERFBOOSTMODE $boostMode
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to set processor boost on battery power (exit code $LASTEXITCODE)."
    }

    powercfg.exe /setactive SCHEME_CURRENT
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to reapply the active power plan (exit code $LASTEXITCODE)."
    }
}

powercfg.exe /qh SCHEME_CURRENT SUB_PROCESSOR PERFBOOSTMODE
if ($LASTEXITCODE -ne 0) {
    throw "Failed to query processor boost (exit code $LASTEXITCODE)."
}