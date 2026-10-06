#Requires -Version 5.1
<#
.SYNOPSIS
    mdmresult - gpresult-style policy report for Group Policy, Intune/MDM and co-managed devices.
.DESCRIPTION
    Runs PolicyPilot's device scan without the GUI and writes the same HTML report as
    PolicyPilot's "Export HTML". The scan, ADMX/CSP enrichment, gap analysis, conflict
    detection and report code are loaded from PolicyPilot.ps1 at runtime, so the output
    stays identical to the GUI.
.PARAMETER Mode
    Local    - Group Policy result (gpresult /x) with ADMX enrichment.
    Intune   - local MDM policy state (PolicyManager, MDM diagnostics, IME).
    Combined - both; use for hybrid joined / co-managed devices. Default.
.PARAMETER Path
    HTML output file. Default: .\mdmresult_<COMPUTERNAME>_<timestamp>.html
.PARAMETER Force
    Overwrite an existing output file.
.PARAMETER Open
    Open the report when finished.
.EXAMPLE
    .\mdmresult.ps1
.EXAMPLE
    .\mdmresult.ps1 -Mode Intune -Path C:\Temp\policy.html -Force -Open
.NOTES
    Run elevated: computer-scope gpresult and Win32 app state need administrator rights.
    Requires PolicyPilot.ps1, admx_metadata.json and csp_metadata.json in the same folder.
#>
[CmdletBinding()]
param(
    [ValidateSet('Local', 'Intune', 'Combined')]
    [string]$Mode = 'Combined',
    [Alias('H')]
    [string]$Path,
    [Alias('F')]
    [switch]$Force,
    [switch]$Open
)

# Same as PolicyPilot: its scan/report code relies on non-terminating errors
$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.Web

$sourcePath = Join-Path $PSScriptRoot 'PolicyPilot.ps1'
if (-not (Test-Path -LiteralPath $sourcePath)) { throw "PolicyPilot.ps1 not found in $PSScriptRoot." }

if (-not $Path) { $Path = "mdmresult_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmmss').html" }
$Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
if ((Test-Path -LiteralPath $Path) -and -not $Force) { throw "Output file exists: $Path (use -Force to overwrite)." }

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) { Write-Warning 'Not elevated: computer-scope Group Policy and Win32 app state may be incomplete.' }

# --- Load PolicyPilot code by AST; any layout change fails loudly instead of drifting ---
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($sourcePath, [ref]$null, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw "PolicyPilot.ps1 has parse errors: $($parseErrors[0].Message)" }

function Get-LastMatch($Items, [string]$What) {
    $found = @($Items)
    if ($found.Count -eq 0) { throw "PolicyPilot.ps1 layout changed: $What not found." }
    $found[$found.Count - 1]
}

function Get-NamedArgument($Command, [string]$Name) {
    $elements = $Command.CommandElements
    for ($i = 0; $i -lt $elements.Count - 1; $i++) {
        if ($elements[$i] -is [System.Management.Automation.Language.CommandParameterAst] -and $elements[$i].ParameterName -eq $Name) {
            return $elements[$i + 1].ScriptBlock
        }
    }
    throw "PolicyPilot.ps1 layout changed: -$Name of the scan call not found."
}

function Write-DebugLog { param([string]$Message, [string]$Level = 'INFO') Write-Verbose "[$Level] $Message" }
function Show-Toast { param($Title, $Message, $Type) Write-Warning "$Title - $Message" }

# Last definition wins, matching how PowerShell resolves PolicyPilot's duplicate definitions
foreach ($name in 'Load-MetadataDatabases', 'Resolve-PolicyFromRegistry', 'Find-Conflicts', 'Build-HtmlReport') {
    $definition = Get-LastMatch ($ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -eq $name }) "function $name"
    . ([scriptblock]::Create($definition.Extent.Text))
}

$scanCall = Get-LastMatch ($ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.CommandAst] -and
        $node.GetCommandName() -eq 'Start-BackgroundWork' -and
        $node.Extent.Text -match 'ScanMode\s*=\s*\$scanMode'
    }, $true)) 'scan Start-BackgroundWork call'
$workBlock = Get-NamedArgument $scanCall 'Work'
$completeStatements = (Get-NamedArgument $scanCall 'OnComplete').EndBlock.Statements

$enrichStep = Get-LastMatch ($completeStatements | Where-Object { $_ -is [System.Management.Automation.Language.IfStatementAst] -and $_.Clauses[0].Item1.Extent.Text -match 'AdmxByReg\.Count' }) 'post-scan ADMX/CSP enrichment'
$groupFunction = Get-LastMatch ($completeStatements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -eq 'Get-IntuneGroup' }) 'Get-IntuneGroup'
$gapStep = Get-LastMatch ($completeStatements | Where-Object { $_ -is [System.Management.Automation.Language.IfStatementAst] -and $_.Extent.Text -match 'NotConfiguredSettings\.Add' }) 'not-configured gap analysis'

# --- Script state PolicyPilot's functions expect ---
$Script:AppVersion = [regex]::Match($ast.Extent.Text, "(?m)^\`$Script:AppVersion\s*=\s*'([^']+)'").Groups[1].Value
$Script:AdmxDbPath = Join-Path $PSScriptRoot 'admx_metadata.json'
$Script:CspDbPath  = Join-Path $PSScriptRoot 'csp_metadata.json'
$Script:AdmxDb = $null; $Script:AdmxByReg = @{}; $Script:AdmxDbAge = $null
$Script:CspDb = $null; $Script:CspByReg = @{}; $Script:CspByPath = @{}; $Script:CspDbAge = $null
$Script:CspMetaKeys = @{}; $Script:CspDbCount = 0
$Script:Prefs = @{ ScanMode = $Mode }
$Script:AllIntuneApps = [System.Collections.Generic.List[object]]::new()
$Script:NotConfiguredSettings = [System.Collections.Generic.List[PSCustomObject]]::new()

Write-Host "mdmresult $($Script:AppVersion) - $Mode scan of $env:COMPUTERNAME" -ForegroundColor Cyan
Load-MetadataDatabases

# --- Run the scan in a runspace, exactly as the GUI does ---
$sync = [hashtable]::Synchronized(@{ StatusQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new()) })
$iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$iss.ExecutionPolicy = [Microsoft.PowerShell.ExecutionPolicy]::Bypass
$runspace = [runspacefactory]::CreateRunspace($iss)
$runspace.ApartmentState = [System.Threading.ApartmentState]::STA
$runspace.Open()
$scanVariables = @{ ScanMode = $Mode; DomainOvr = ''; DcOvr = ''; OUScope = ''; ForceRefresh = $false; SyncH = $sync; ScriptRoot = $PSScriptRoot }
foreach ($key in $scanVariables.Keys) { $runspace.SessionStateProxy.SetVariable($key, $scanVariables[$key]) }
$shell = [powershell]::Create()
$shell.Runspace = $runspace
[void]$shell.AddScript($workBlock.GetScriptBlock().ToString())

$levelColors = @{ ERROR = 'Red'; WARN = 'Yellow'; SUCCESS = 'Green'; STEP = 'Cyan'; INFO = 'Gray'; DEBUG = 'DarkGray' }
$verbose = $VerbosePreference -ne 'SilentlyContinue'
$progress = @{ Activity = "PolicyPilot $Mode scan"; Status = 'Scanning...'; PercentComplete = 0 }
function Write-ScanStatus {
    while ($sync.StatusQueue.Count -gt 0) {
        $message = $sync.StatusQueue.Dequeue()
        switch ($message.Type) {
            'Log' {
                if ($verbose -or $message.Level -in 'STEP', 'SUCCESS', 'WARN', 'ERROR') {
                    $color = if ($levelColors.ContainsKey("$($message.Level)")) { $levelColors["$($message.Level)"] } else { 'Gray' }
                    Write-Host "  $($message.Text)" -ForegroundColor $color
                }
            }
            'Status' { $progress.Status = "$($message.Text)"; Write-Progress @progress }
            'Progress' { $progress.PercentComplete = [math]::Min(100, [math]::Max(0, [int]$message.Value)); Write-Progress @progress }
        }
    }
}

try {
    $async = $shell.BeginInvoke()
    while (-not $async.IsCompleted) { Write-ScanStatus; Start-Sleep -Milliseconds 200 }
    $results = $shell.EndInvoke($async)
    Write-ScanStatus
    $scanErrors = @($shell.Streams.Error)
} finally {
    Write-Progress -Activity $progress.Activity -Completed
    $shell.Dispose()
    $runspace.Dispose()
}

$scanResult = if ($results.Count -gt 0) { $results[$results.Count - 1].psobject.BaseObject } else { $null }
if (-not ($scanResult -is [hashtable]) -or $scanResult.Error) {
    foreach ($scanError in $scanErrors) { Write-Warning "$scanError" }
    $reason = if ($scanResult -is [hashtable] -and $scanResult.Error) { $scanResult.Error } else { 'scan returned no data' }
    throw "Scan failed: $reason"
}

# --- Post-scan steps from PolicyPilot's OnComplete handler ---
$Script:ScanData = $scanResult
if ($scanResult.CspMeta) { $Script:CspMetaKeys = $scanResult.CspMeta }
$Script:CspDbAge   = $scanResult.CspDbAge
$Script:CspDbCount = $scanResult.CspDbCount
$Script:AllSettings = $scanResult.Settings
foreach ($app in @($scanResult.Apps)) { if ($app) { [void]$Script:AllIntuneApps.Add($app) } }
. ([scriptblock]::Create($enrichStep.Extent.Text))
. ([scriptblock]::Create($groupFunction.Extent.Text))
. ([scriptblock]::Create($gapStep.Extent.Text))

$conflicts = @(Find-Conflicts $scanResult.Settings)
$html = Build-HtmlReport
$outDir = Split-Path -Parent $Path
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) { [void](New-Item -ItemType Directory -Path $outDir -Force) }
[System.IO.File]::WriteAllText($Path, $html, [System.Text.Encoding]::UTF8)

# --- gpresult /r-style summary ---
$compliance = $null
if ($scanResult.MdmInfo -and $scanResult.MdmInfo.MdmDiag -and $scanResult.MdmInfo.MdmDiag.Compliance) { $compliance = $scanResult.MdmInfo.MdmDiag.Compliance.Status }
$failedApps = @($Script:AllIntuneApps | Where-Object { $_.InstallState -eq 'Failed' }).Count
Write-Host ''
Write-Host "Policy result for $env:COMPUTERNAME ($($scanResult.Domain))" -ForegroundColor White
Write-Host ('  {0,-22} {1}' -f 'Policy sources', @($scanResult.GPOs).Count)
Write-Host ('  {0,-22} {1}' -f 'Settings applied', @($scanResult.Settings).Count)
Write-Host ('  {0,-22} {1} / {2}' -f 'Conflicts / redundant', @($conflicts | Where-Object Severity -eq 'Conflict').Count, @($conflicts | Where-Object Severity -eq 'Redundant').Count)
if ($Mode -ne 'Local') {
    Write-Host ('  {0,-22} {1}' -f 'CSP not configured', $Script:NotConfiguredSettings.Count)
    Write-Host ('  {0,-22} {1} ({2} failed)' -f 'Apps tracked', $Script:AllIntuneApps.Count, $failedApps)
    if ($compliance) { Write-Host ('  {0,-22} {1}' -f 'Compliance', $compliance) }
}
Write-Host "Report: $Path" -ForegroundColor Green

if ($Open) { Invoke-Item -LiteralPath $Path }
