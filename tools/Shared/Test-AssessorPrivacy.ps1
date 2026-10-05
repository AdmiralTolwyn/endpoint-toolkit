#Requires -Version 5.1
# Verifies protected persistence and classification helpers in the WPF assessors without starting the UI.
$ErrorActionPreference = 'Stop'
function Assert-Privacy([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
$ToolsRoot = Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'CollectorPrivacy.ps1')
$Apps = @{
    AvdAssessor = 'AvdAssessor\AvdAssessor.ps1'
    W365Assessor = 'W365Assessor\W365Assessor.ps1'
    BaselinePilot = 'BaselineAssessor\BaselinePilot.ps1'
}
$Temp = Join-Path ([IO.Path]::GetTempPath()) ('assessor-privacy-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($Temp)
try {
    foreach ($App in $Apps.GetEnumerator()) {
        $Path = Join-Path $ToolsRoot $App.Value
        $Tokens = $null; $Errors = $null
        $Ast = [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$Tokens, [ref]$Errors)
        Assert-Privacy ($Errors.Count -eq 0) "$($App.Key) does not parse"
        foreach ($Name in @('Write-AssessorProtectedText', 'Get-AssessorDataHandling', 'Get-BaselinePilotPrivacyMode')) {
            $Definition = $Ast.Find({ param($Node) $Node -is [Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -ceq $Name }, $true)
            if ($Definition) { . ([scriptblock]::Create($Definition.Extent.Text)) }
        }
        $Text = Get-Content -LiteralPath $Path -Raw
        Assert-Privacy ($Text -notmatch 'ConvertTo-Json -Depth 10 \| Set-Content|WriteAllText\((\$dlg\.FileName|\$Global:AutoSaveFile|\$BackupPath)|Export-Csv -Path \$dlg') "$($App.Key) still writes assessment data without protection"
        Assert-Privacy ($Text.Contains("Data handling <span>`$(& `$enc (Get-AssessorDataHandling))</span>")) "$($App.Key) HTML header lacks the classification"

        $Target = Join-Path $Temp ($App.Key + '.json')
        Write-AssessorProtectedText -Path $Target -Text '{"v":1}'
        $Rejected = $false
        try { Write-AssessorProtectedText -Path $Target -Text '{"v":2}' } catch { $Rejected = $true }
        Assert-Privacy ($Rejected -and (Get-Content -LiteralPath $Target -Raw) -ceq '{"v":1}') "$($App.Key) overwrote without -Replace"
        Remove-Item -LiteralPath $Target -Force
        [IO.File]::WriteAllText($Target, 'legacy')
        Assert-Privacy (-not (Get-Acl -LiteralPath $Target).AreAccessRulesProtected) 'Test precondition: legacy file should inherit permissions'
        Write-AssessorProtectedText -Path $Target -Text '{"v":3}' -Replace
        Assert-Privacy ((Get-Content -LiteralPath $Target -Raw) -ceq '{"v":3}' -and (Get-Acl -LiteralPath $Target).AreAccessRulesProtected) "$($App.Key) replace kept inherited permissions"
        Assert-Privacy (@(Get-ChildItem -LiteralPath $Temp | Where-Object { $_.Name -match '\.(tmp|old)$' }).Count -eq 0) "$($App.Key) left temporary files"
        $Bytes = [IO.File]::ReadAllBytes($Target)
        Assert-Privacy ($Bytes[0] -ne 0xEF) "$($App.Key) wrote a BOM"

        $Global:Assessment = [pscustomobject]@{ Discovery = [pscustomobject]@{ Privacy = [pscustomobject]@{ Mode = 'Pseudonymous' } }; CollectionData = [pscustomobject]@{ Privacy = [pscustomobject]@{ Mode = 'Pseudonymous' } } }
        Assert-Privacy ((Get-AssessorDataHandling) -ceq 'Confidential - source collection: Pseudonymous') "$($App.Key) classification text incorrect"
        $Global:Assessment = [pscustomobject]@{ Discovery = [pscustomobject]@{ SchemaVersion = '1.0' }; CollectionData = [pscustomobject]@{ systemInfo = @{} } }
        Assert-Privacy ((Get-AssessorDataHandling) -ceq 'Confidential - source collection: Legacy, unclassified') "$($App.Key) legacy classification text incorrect"
    }
    $W365 = Get-Content -LiteralPath (Join-Path $ToolsRoot $Apps.W365Assessor) -Raw
    Assert-Privacy ($W365.Contains("if (`$C.PSObject.Properties['UserPrincipalName']) { `$C.UserPrincipalName } else { `$C.UserKey }")) 'W365 report does not fall back to the user pseudonym'
    $Pilot = Get-Content -LiteralPath (Join-Path $ToolsRoot $Apps.BaselinePilot) -Raw
    Assert-Privacy ($Pilot.Contains("`$ActualValue.topUsers -and (Get-BaselinePilotPrivacyMode) -ne 'Pseudonymous'")) 'BaselinePilot renders topUsers for pseudonymous collections'
} finally {
    Remove-Item -LiteralPath $Temp -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Variable -Name Assessment -Scope Global -ErrorAction SilentlyContinue
}
Write-Output 'PASS: WPF assessors use protected CreateNew/replace persistence, classification headers, pseudonym user columns and pseudonymous topUsers suppression.'
