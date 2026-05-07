# DeviceDecommissioner.Tests.ps1 — Pester 5 smoke tests
# Run:  Invoke-Pester .\DeviceDecommissioner.Tests.ps1 -Output Detailed

BeforeAll {
    $Script:ScriptRoot = $PSScriptRoot
    $Script:ScriptPath = Join-Path $Script:ScriptRoot 'DeviceDecommissioner.ps1'

    # Parse the script to extract functions without running it (no WPF dependency).
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Script:ScriptPath, [ref]$null, [ref]$null)
    $fns = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)

    # Dot-source individual function bodies into this scope.
    function Import-FunctionFromAst {
        param([string]$Name)
        $fn = $fns | Where-Object { $_.Name -eq $Name } | Select-Object -First 1
        if (-not $fn) { throw "Function '$Name' not found in AST" }
        return "function $Name $($fn.Body.Extent.Text)"
    }
}

Describe 'Script parse validation' {
    It 'PS1 parses without errors' {
        $tokens = $null; $errors = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile($Script:ScriptPath, [ref]$tokens, [ref]$errors)
        $errors | Should -BeNullOrEmpty
    }

    It 'XAML parses without errors' {
        $xamlPath = Join-Path $Script:ScriptRoot 'DeviceDecommissioner_UI.xaml'
        { Add-Type -AssemblyName PresentationFramework; [void][Windows.Markup.XamlReader]::Parse([IO.File]::ReadAllText($xamlPath)) } | Should -Not -Throw
    }
}

Describe 'AD filter quote-escape' {
    BeforeAll {
        # The AD lookup builds a filter string. Let's validate escaping directly.
        # Pattern from the script: $q = $Query.Replace("'", "''")
        function Escape-ADFilter { param([string]$q) $q.Replace("'", "''") }
    }

    It 'escapes a single quote in hostname' {
        Escape-ADFilter "TEST'HOST" | Should -Be "TEST''HOST"
    }

    It 'escapes multiple quotes' {
        Escape-ADFilter "a'b'c" | Should -Be "a''b''c"
    }

    It 'returns normal name unchanged' {
        Escape-ADFilter 'LAPTOP-ABC123' | Should -Be 'LAPTOP-ABC123'
    }

    It 'filter with escaped name is injection-safe' {
        $q = Escape-ADFilter "'; DROP TABLE --"
        $filter = "Name -eq '$q'"
        $filter | Should -Be "Name -eq '''; DROP TABLE --'"
        # The value inside the filter is treated as a literal string, not AD filter syntax.
    }
}

Describe 'SCCM site-code regex' {
    BeforeAll {
        # Pattern from Apply-SettingsFromDialog: ^[A-Z0-9]{1,3}$
        $Script:SiteCodePattern = '^[A-Z0-9]{1,3}$'
    }

    It 'accepts valid 3-char code' {
        'PR1' -match $Script:SiteCodePattern | Should -Be $true
    }

    It 'accepts single-char code' {
        'A' -match $Script:SiteCodePattern | Should -Be $true
    }

    It 'rejects lowercase' {
        # Site code validation in the script uses .ToUpper() before matching, but the regex itself
        # only accepts uppercase. Use -cmatch for case-sensitive validation.
        'pr1' -cmatch $Script:SiteCodePattern | Should -Be $false
    }

    It 'rejects too-long code' {
        'ABCD' -match $Script:SiteCodePattern | Should -Be $false
    }

    It 'rejects empty string' {
        '' -match $Script:SiteCodePattern | Should -Be $false
    }

    It 'rejects special characters' {
        'A-1' -match $Script:SiteCodePattern | Should -Be $false
    }
}

Describe 'Generation-counter staleness drop' {
    It 'stale callback is ignored when gen mismatches' {
        # Simulate: gen at lookup start = 5, global gen incremented to 6.
        $Global:LookupGen = 6
        $callbackGen = 5
        # The Save-LookupResult guard:
        $isStale = ($callbackGen -ge 0 -and $callbackGen -ne $Global:LookupGen)
        $isStale | Should -Be $true
    }

    It 'current callback is processed when gen matches' {
        $Global:LookupGen = 7
        $callbackGen = 7
        $isStale = ($callbackGen -ge 0 -and $callbackGen -ne $Global:LookupGen)
        $isStale | Should -Be $false
    }
}

Describe 'Get-DefaultSettings structure' {
    BeforeAll {
        Invoke-Expression (Import-FunctionFromAst 'Get-DefaultSettings')
    }

    It 'returns a hashtable' {
        $s = Get-DefaultSettings
        $s | Should -BeOfType [hashtable]
    }

    It 'has all required top-level keys' {
        $s = Get-DefaultSettings
        $s.Keys | Should -Contain 'Theme'
        $s.Keys | Should -Contain 'AD'
        $s.Keys | Should -Contain 'Entra'
        $s.Keys | Should -Contain 'Intune'
        $s.Keys | Should -Contain 'SCCM'
        $s.Keys | Should -Contain 'SidebarVisible'
        $s.Keys | Should -Contain 'LogPanelVisible'
        $s.Keys | Should -Contain 'RecentActivityDays'
    }

    It 'defaults to Dark theme' {
        (Get-DefaultSettings).Theme | Should -Be 'Dark'
    }

    It 'defaults RecentActivityDays to 7' {
        (Get-DefaultSettings).RecentActivityDays | Should -Be 7
    }

    It 'defaults SCCM.Enabled to false' {
        (Get-DefaultSettings).SCCM.Enabled | Should -Be $false
    }
}

Describe 'Safety warnings' {
    BeforeAll {
        # Provide required globals
        $Global:Settings = @{ RecentActivityDays = 7 }
        $Global:LookupState = @{ Results = @{} }
        Invoke-Expression (Import-FunctionFromAst 'Get-SafetyWarnings')
    }

    It 'warns when AD has BitLocker keys' {
        $Global:LookupState.Results.AD = @{
            Found = $true
            Raw = @{ BitLockerKeyCount = 2; HasLAPSPassword = $false; LastLogonDate = $null }
        }
        $w = Get-SafetyWarnings -Targets @('AD')
        $w | Should -Not -BeNullOrEmpty
        ($w -join ';') | Should -Match 'BitLocker'
    }

    It 'warns when AD has LAPS password' {
        $Global:LookupState.Results.AD = @{
            Found = $true
            Raw = @{ BitLockerKeyCount = 0; HasLAPSPassword = $true; LastLogonDate = $null }
        }
        $w = Get-SafetyWarnings -Targets @('AD')
        $w | Should -Not -BeNullOrEmpty
        ($w -join ';') | Should -Match 'LAPS'
    }

    It 'warns when Intune device synced recently' {
        $recentDate = (Get-Date).AddDays(-2).ToString('o')
        $Global:LookupState.Results.Intune = @{
            Found = $true
            Raw = @{ LastSyncDateTime = $recentDate }
        }
        $w = Get-SafetyWarnings -Targets @('Intune')
        $w | Should -Not -BeNullOrEmpty
        ($w -join ';') | Should -Match 'still be in use'
    }

    It 'does not warn when Intune device synced long ago' {
        $oldDate = (Get-Date).AddDays(-30).ToString('o')
        $Global:LookupState.Results = @{
            Intune = @{
                Found = $true
                Raw = @{ LastSyncDateTime = $oldDate }
            }
        }
        $w = Get-SafetyWarnings -Targets @('Intune')
        $w | Should -BeNullOrEmpty
    }

    It 'returns no warnings when no data' {
        $Global:LookupState.Results.AD = @{ Found = $false; Raw = $null }
        $w = Get-SafetyWarnings -Targets @('AD')
        $w | Should -BeNullOrEmpty
    }
}

Describe 'Audit entry structure' {
    BeforeAll {
        $Global:AuditFile = Join-Path $TestDrive 'test-audit.json'
        # Provide Write-DebugLog stub
        function Write-DebugLog { param($Message, $Level) }
        Invoke-Expression (Import-FunctionFromAst 'Save-AuditEntry')
    }

    It 'creates a JSON file with valid structure' {
        Save-AuditEntry -DeviceQuery 'TEST-PC' -Targets @('AD','Entra') -DryRun $true -StepResults @{
            AD    = @{ Success = $true;  Message = 'Would remove' }
            Entra = @{ Success = $false; Message = 'Scope missing' }
        }
        $Global:AuditFile | Should -Exist
        $json = Get-Content $Global:AuditFile -Raw | ConvertFrom-Json
        $json | Should -Not -BeNullOrEmpty
        $entry = if ($json -is [array]) { $json[0] } else { $json }
        $entry.Device    | Should -Be 'TEST-PC'
        $entry.DryRun    | Should -Be $true
        $entry.Operator  | Should -Not -BeNullOrEmpty
        $entry.Results.AD.Success | Should -Be $true
        $entry.Results.Entra.Success | Should -Be $false
    }

    It 'appends multiple entries' {
        Save-AuditEntry -DeviceQuery 'TEST-PC2' -Targets @('SCCM') -DryRun $false -StepResults @{
            SCCM = @{ Success = $true; Message = 'Removed' }
        }
        $json = Get-Content $Global:AuditFile -Raw | ConvertFrom-Json
        $json.Count | Should -Be 2
    }

    It 'records Autopilot results' {
        Save-AuditEntry -DeviceQuery 'TEST-AP' -Targets @('Autopilot') -DryRun $false -StepResults @{
            Autopilot = @{ Success = $true; Message = 'Removed Autopilot identity id=abc' }
        }
        $json = Get-Content $Global:AuditFile -Raw | ConvertFrom-Json
        $entry = $json[-1]
        $entry.Device | Should -Be 'TEST-AP'
        $entry.Targets | Should -Contain 'Autopilot'
        $entry.Results.Autopilot.Success | Should -Be $true
    }

    It 'tolerates empty StepResults gracefully' {
        { Save-AuditEntry -DeviceQuery 'TEST-EMPTY' -Targets @() -DryRun $true -StepResults @{} } | Should -Not -Throw
        $json = Get-Content $Global:AuditFile -Raw | ConvertFrom-Json
        $json[-1].Device | Should -Be 'TEST-EMPTY'
    }

    It 'preserves entry ordering across appends' {
        $count = (Get-Content $Global:AuditFile -Raw | ConvertFrom-Json).Count
        Save-AuditEntry -DeviceQuery 'ORDER-TEST' -Targets @('AD') -DryRun $false -StepResults @{
            AD = @{ Success = $true; Message = 'OK' }
        }
        $json = Get-Content $Global:AuditFile -Raw | ConvertFrom-Json
        $json.Count | Should -Be ($count + 1)
        $json[-1].Device | Should -Be 'ORDER-TEST'
    }
}

Describe 'Safety warnings — Autopilot edge cases' {
    BeforeAll {
        $Global:Settings    = @{ RecentActivityDays = 7 }
        $Global:LookupState = @{ Results = @{} }
        Invoke-Expression (Import-FunctionFromAst 'Get-SafetyWarnings')
    }

    It 'no warnings when only Autopilot is queried (no recent-activity field)' {
        $Global:LookupState.Results = @{
            Autopilot = @{
                Found = $true
                Raw = @{ Id='abc'; SerialNumber='SN123'; Model='Surface' }
            }
        }
        $w = Get-SafetyWarnings -Targets @('Autopilot')
        $w | Should -BeNullOrEmpty
    }

    It 'combines warnings from multiple systems' {
        $recent = (Get-Date).AddDays(-1).ToString('o')
        $Global:LookupState.Results = @{
            AD     = @{ Found=$true; Raw=@{ HasLAPSPassword=$true; BitLockerKeyCount=2; LastLogonDate=$recent } }
            Intune = @{ Found=$true; Raw=@{ LastSyncDateTime=$recent } }
        }
        $w = Get-SafetyWarnings -Targets @('AD','Intune')
        $w.Count | Should -BeGreaterOrEqual 3
        ($w -join ';') | Should -Match 'BitLocker'
        ($w -join ';') | Should -Match 'LAPS'
        ($w -join ';') | Should -Match 'still be in use'
    }

    It 'skipped systems contribute no warnings' {
        $Global:LookupState.Results = @{
            AD = @{ Found=$true; Raw=@{ HasLAPSPassword=$true; BitLockerKeyCount=5 } }
        }
        $w = Get-SafetyWarnings -Targets @('Entra')   # AD not in target list
        $w | Should -BeNullOrEmpty
    }

    It 'respects custom RecentActivityDays threshold' {
        $Global:Settings.RecentActivityDays = 1   # tighter window
        $threeDaysAgo = (Get-Date).AddDays(-3).ToString('o')
        $Global:LookupState.Results = @{
            Intune = @{ Found=$true; Raw=@{ LastSyncDateTime=$threeDaysAgo } }
        }
        $w = Get-SafetyWarnings -Targets @('Intune')
        $w | Should -BeNullOrEmpty
        $Global:Settings.RecentActivityDays = 7   # restore for later tests
    }
}

Describe 'History view — symbol conversion' {
    BeforeAll {
        Invoke-Expression (Import-FunctionFromAst 'ConvertTo-HistorySymbol')
    }

    It 'returns dash for null result' {
        ConvertTo-HistorySymbol $null | Should -Be '-'
    }

    It 'returns checkmark for success' {
        $r = ConvertTo-HistorySymbol @{ Success=$true; Message='OK' }
        [int][char]$r | Should -Be 0x2713
    }

    It 'returns cross for failure' {
        $r = ConvertTo-HistorySymbol @{ Success=$false; Message='Err' }
        [int][char]$r | Should -Be 0x2717
    }
}

Describe 'AllSystems constant' {
    It 'parses to a 5-element array' {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($Script:ScriptPath, [ref]$null, [ref]$null)
        $assignments = $ast.FindAll({
            $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $args[0].Left.Extent.Text -match 'AllSystems'
        }, $true)
        $assignments.Count | Should -BeGreaterOrEqual 1
        $rhs = $assignments[0].Right.Extent.Text
        $rhs | Should -Match 'AD'
        $rhs | Should -Match 'Entra'
        $rhs | Should -Match 'Intune'
        $rhs | Should -Match 'Autopilot'
        $rhs | Should -Match 'SCCM'
    }
}

Describe 'Achievement definitions' {
    It 'defines exactly 30 achievements' {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($Script:ScriptPath, [ref]$null, [ref]$null)
        $defAssignments = $ast.FindAll({
            $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $args[0].Left.Extent.Text -match 'AchievementDefs'
        }, $true)
        $defAssignments.Count | Should -BeGreaterOrEqual 1
        # Count hashtables inside the @( … ) array literal
        $defs = $defAssignments[0].Right.FindAll({
            $args[0] -is [System.Management.Automation.Language.HashtableAst]
        }, $true)
        $defs.Count | Should -Be 30
    }

    It 'every achievement has Id, Icon, Name, Desc keys' {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($Script:ScriptPath, [ref]$null, [ref]$null)
        $defAssignments = $ast.FindAll({
            $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $args[0].Left.Extent.Text -match 'AchievementDefs'
        }, $true)
        $defs = $defAssignments[0].Right.FindAll({
            $args[0] -is [System.Management.Automation.Language.HashtableAst]
        }, $true)
        foreach ($d in $defs) {
            $keys = $d.KeyValuePairs | ForEach-Object { $_.Item1.Extent.Text }
            $keys | Should -Contain 'Id'
            $keys | Should -Contain 'Icon'
            $keys | Should -Contain 'Name'
            $keys | Should -Contain 'Desc'
        }
    }

    It 'all achievement Ids are unique' {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($Script:ScriptPath, [ref]$null, [ref]$null)
        $defAssignments = $ast.FindAll({
            $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $args[0].Left.Extent.Text -match 'AchievementDefs'
        }, $true)
        $defs = $defAssignments[0].Right.FindAll({
            $args[0] -is [System.Management.Automation.Language.HashtableAst]
        }, $true)
        $ids = foreach ($d in $defs) {
            $idPair = $d.KeyValuePairs | Where-Object { $_.Item1.Extent.Text -eq 'Id' }
            $idPair.Item2.Extent.Text.Trim("'", '"')
        }
        ($ids | Sort-Object -Unique).Count | Should -Be $ids.Count
    }
}
