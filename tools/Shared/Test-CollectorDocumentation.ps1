<#
.SYNOPSIS
    Validates comment-based help for the assessor collectors and their shared helpers.
.DESCRIPTION
    Parses production scripts without executing them. Requires script and function synopses,
    help for every declared parameter, and the repository disclaimer in each script's notes.
    GUIs, database builders, launchers and tests are outside this documentation scope.
.NOTES
    Disclaimer: This script is provided "AS IS" with no warranties and confers no rights.
#>
#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
$ToolsRoot = Split-Path $PSScriptRoot -Parent
$Scripts = @(
    'Shared/CollectorPrivacy.ps1',
    'AvdAssessor/CollectorPrivacy.ps1',
    'AvdAssessor/Invoke-AvdDiscovery.ps1',
    'BaselineAssessor/CollectorPrivacy.ps1',
    'BaselineAssessor/Invoke-BaselineCollection.ps1',
    'IntuneAssessor/CollectorPrivacy.ps1',
    'IntuneAssessor/Get-IntuneDefenderEvidence.ps1',
    'IntuneAssessor/Get-IntuneEndpointEvidence.ps1',
    'IntuneAssessor/IntuneEndpointTimestamps.ps1',
    'IntuneAssessor/IntuneEpmRules.ps1',
    'IntuneAssessor/IntuneExpansion.ps1',
    'IntuneAssessor/IntuneMamLaunch.ps1',
    'IntuneAssessor/IntunePolicyPayloads.ps1',
    'IntuneAssessor/IntuneServices.ps1',
    'IntuneAssessor/Invoke-IntuneDiscovery.ps1',
    'W365Assessor/CollectorPrivacy.ps1',
    'W365Assessor/Invoke-W365Discovery.ps1',
    'W365Assessor/W365ConditionalAccess.ps1',
    'W365Assessor/W365Monitoring.ps1',
    'W365Assessor/W365Reports.ps1',
    'W365Assessor/W365UserExperienceSync.ps1'
)
$Disclaimer = 'This script is provided "AS IS" with no warranties and confers no rights.'
$FunctionCount = 0
foreach ($RelativePath in $Scripts) {
    $Path = Join-Path $ToolsRoot $RelativePath
    $Tokens = $null
    $ParseErrors = $null
    $Ast = [Management.Automation.Language.Parser]::ParseInput([IO.File]::ReadAllText($Path), [ref]$Tokens, [ref]$ParseErrors)
    if ($ParseErrors.Count) { throw "${RelativePath}: parse errors: $($ParseErrors.Message -join '; ')" }
    $ScriptHelp = $Ast.GetHelpContent()
    if (-not $ScriptHelp -or [string]::IsNullOrWhiteSpace($ScriptHelp.Synopsis)) { throw "${RelativePath}: missing script synopsis" }
    if (-not ([string]$ScriptHelp.Notes).Contains($Disclaimer)) { throw "${RelativePath}: missing disclaimer in .NOTES" }
    foreach ($Parameter in $Ast.ParamBlock.Parameters) {
        $Name = $Parameter.Name.VariablePath.UserPath
        if ([string]::IsNullOrWhiteSpace($ScriptHelp.Parameters[$Name.ToUpperInvariant()])) { throw "${RelativePath}: missing script parameter help for $Name" }
    }
    foreach ($Definition in $Ast.FindAll({ param($Node) $Node -is [Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
        $FunctionCount++
        $Help = $Definition.GetHelpContent()
        if (-not $Help -or [string]::IsNullOrWhiteSpace($Help.Synopsis)) { throw "${RelativePath}: missing synopsis for $($Definition.Name)" }
        $Parameters = if ($Definition.Body.ParamBlock) { $Definition.Body.ParamBlock.Parameters } else { $Definition.Parameters }
        foreach ($Parameter in $Parameters) {
            $Name = $Parameter.Name.VariablePath.UserPath
            if ([string]::IsNullOrWhiteSpace($Help.Parameters[$Name.ToUpperInvariant()])) { throw "${RelativePath}: missing parameter help for $($Definition.Name).$Name" }
        }
    }
}
Write-Output "PASS: script headers, disclaimers and function parameter help in $($Scripts.Count) scripts ($FunctionCount functions)."