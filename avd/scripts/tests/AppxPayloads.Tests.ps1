#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    $script:InstallerPath = Join-Path $PSScriptRoot '..\Install-AppxPayloads.ps1'
    $script:DownloaderPath = Join-Path $PSScriptRoot '..\Get-StubAppPayloads.ps1'
    $tokens = $null
    $parseErrors = $null
    $installerAst = [System.Management.Automation.Language.Parser]::ParseFile(
        $InstallerPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors) { throw ($parseErrors -join [Environment]::NewLine) }
    $script:InstallerBody = Join-Path $TestDrive 'InstallerUnderTest.ps1'
    ($installerAst.ParamBlock.Extent.Text + [Environment]::NewLine +
        (($installerAst.EndBlock.Statements | ForEach-Object { $_.Extent.Text }) -join [Environment]::NewLine)) |
        Set-Content -LiteralPath $InstallerBody -Encoding UTF8
    foreach ($statement in $installerAst.EndBlock.Statements) {
        if ($statement -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $statement.Name -ne 'Write-Log') {
            . ([scriptblock]::Create($statement.Extent.Text))
        }
    }
    function Write-Log { param($Message, $Level) }
    function Add-AppxProvisionedPackage {
        [CmdletBinding()]
        param([switch]$Online, [string]$PackagePath, [string]$LicensePath,
            [string[]]$DependencyPackagePath, [switch]$SkipLicense, [string]$StubPackageOption)
        throw 'Unmocked DISM call in a test.'
    }
    function Get-AppxProvisionedPackage {
        param([switch]$Online)
        throw 'Unmocked DISM inventory call in a test.'
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    function New-TestPackage {
        param(
            [string]$Path,
            [string]$Name = 'Contoso.App',
            [switch]$Framework,
            [switch]$Bundle
        )
        $archive = [IO.Compression.ZipFile]::Open($Path, [IO.Compression.ZipArchiveMode]::Create)
        try {
            $entryPath = if ($Bundle) { 'AppxMetadata/AppxBundleManifest.xml' } else { 'AppxManifest.xml' }
            $entry = $archive.CreateEntry($entryPath)
            $writer = [IO.StreamWriter]::new($entry.Open())
            try {
                if ($Bundle) {
                    $writer.Write(('<Bundle xmlns="http://schemas.microsoft.com/appx/2013/bundle"><Identity Name="{0}" Publisher="CN=Contoso" Version="1.0.0.0"/></Bundle>' -f $Name))
                }
                else {
                    $writer.Write(('<Package xmlns="http://schemas.microsoft.com/appx/manifest/foundation/windows10"><Identity Name="{0}" Publisher="CN=Contoso" Version="1.0.0.0" ProcessorArchitecture="x64"/><Properties><Framework>{1}</Framework></Properties></Package>' -f $Name, $Framework.IsPresent.ToString().ToLowerInvariant()))
                }
            }
            finally { $writer.Dispose() }
        }
        finally { $archive.Dispose() }
        Get-Item -LiteralPath $Path
    }
}

Describe 'Payload discovery' {
    BeforeEach {
        $root = Join-Path $TestDrive ([guid]::NewGuid().ToString())
        New-Item -Path $root -ItemType Directory | Out-Null
    }

    It 'treats an unlicensed loose appx as a main package' {
        $package = New-TestPackage -Path (Join-Path $root 'Friendly App.appx')
        $payload = Get-Payload -Root $root
        $payload.Bundles.Count | Should -Be 1
        $payload.Bundles[0].FullName | Should -Be $package.FullName
        $payload.Dependencies.Count | Should -Be 0
    }

    It 'recognizes frameworks in appx and msix containers' {
        New-TestPackage -Path (Join-Path $root 'Framework.appx') -Framework | Out-Null
        New-TestPackage -Path (Join-Path $root 'OtherFramework.msix') -Framework | Out-Null
        $payload = Get-Payload -Root $root
        $payload.Bundles.Count | Should -Be 0
        $payload.Dependencies.Count | Should -Be 2
    }

    It 'returns empty collections for an empty tree' {
        $payload = Get-Payload -Root $root
        $payload.Bundles.Count | Should -Be 0
        $payload.Dependencies.Count | Should -Be 0
    }

    It 'rejects a nonexistent source directory' {
        { Get-Payload -Root (Join-Path $root 'missing') } | Should -Throw
    }

    It 'reads bundle identity from AppxMetadata' {
        $bundle = New-TestPackage -Path (Join-Path $root 'Friendly Bundle.msixbundle') -Bundle
        (Get-Payload -Root $root).Bundles.Count | Should -Be 1
        Test-ShouldUpdateProvisioned -Bundle $bundle -ProvisionedNames @('Contoso.App') | Should -BeTrue
    }

    It 'rejects invalid package archives before provisioning' {
        Set-Content -LiteralPath (Join-Path $root 'Broken.appx') -Value 'not a package'
        { Get-Payload -Root $root } | Should -Throw
    }
}

Describe 'Provisioning selection and arguments' {
    BeforeEach {
        $package = New-TestPackage -Path (Join-Path $TestDrive "$([guid]::NewGuid()).appx")
        Mock Add-AppxProvisionedPackage { }
    }

    It 'matches the embedded identity even when the file has a friendly name' {
        Test-ShouldUpdateProvisioned -Bundle $package -ProvisionedNames @('contoso.app') | Should -BeTrue
    }

    It 'does not confuse an identity prefix with a match' {
        Test-ShouldUpdateProvisioned -Bundle $package -ProvisionedNames @('Contoso') | Should -BeFalse
    }

    It 'accepts an image with no provisioned apps' {
        Test-ShouldUpdateProvisioned -Bundle $package -ProvisionedNames @() | Should -BeFalse
    }

    It 'requests full provisioning with a license and dependencies in one call' {
        $result = Install-Bundle -Bundle $package -LicensePath 'License.xml' -DependencyPackagePath @('Framework.appx')
        $result.Status | Should -Be 'Success'
        Should -Invoke Add-AppxProvisionedPackage -Times 1 -Exactly -ParameterFilter {
            $Online -and $StubPackageOption -eq 'InstallFull' -and
            $LicensePath -eq 'License.xml' -and -not $SkipLicense -and
            $DependencyPackagePath.Count -eq 1 -and $DependencyPackagePath[0] -eq 'Framework.appx'
        }
    }

    It 'uses SkipLicense only with explicit permission' {
        (Install-Bundle -Bundle $package -SkipLicense).Status | Should -Be 'Success'
        Should -Invoke Add-AppxProvisionedPackage -Times 1 -Exactly -ParameterFilter { $SkipLicense -and -not $LicensePath }
    }

    It 'turns a DISM error into a failed result' {
        Mock Add-AppxProvisionedPackage { throw 'Simulated DISM failure' }
        $result = Install-Bundle -Bundle $package -SkipLicense
        $result.Status | Should -Be 'Failed'
        $result.Error | Should -Be 'Simulated DISM failure'
    }

    It 'fails without a license unless omission is explicitly allowed' {
        (Install-Bundle -Bundle $package).Status | Should -Be 'Failed'
        Should -Invoke Add-AppxProvisionedPackage -Times 0 -Exactly
    }
}

Describe 'License resolution' {
    BeforeEach {
        $root = Join-Path $TestDrive ([guid]::NewGuid().ToString())
        New-Item -Path $root -ItemType Directory | Out-Null
        $package = New-TestPackage -Path (Join-Path $root 'Friendly App.appx')
    }

    It 'finds the actual winget Store ID license filename' {
        $license = New-Item -Path (Join-Path $root '9TEST1234567_License.xml') -ItemType File
        Resolve-LicensePath -Bundle $package | Should -Be $license.FullName
    }

    It 'prefers an exact FoD basename license' {
        New-Item -Path (Join-Path $root '9TEST1234567_License.xml') -ItemType File | Out-Null
        $license = New-Item -Path (Join-Path $root 'Friendly App.xml') -ItemType File
        Resolve-LicensePath -Bundle $package | Should -Be $license.FullName
    }

    It 'rejects ambiguous Store licenses' {
        New-Item -Path (Join-Path $root 'FIRST_License.xml') -ItemType File | Out-Null
        New-Item -Path (Join-Path $root 'SECOND_License.xml') -ItemType File | Out-Null
        { Resolve-LicensePath -Bundle $package } | Should -Throw '*Ambiguous*'
    }

    It 'does not mistake an unrelated XML file for a license' {
        New-Item -Path (Join-Path $root 'Settings.xml') -ItemType File | Out-Null
        Resolve-LicensePath -Bundle $package | Should -BeNullOrEmpty
    }
}

Describe 'Downloader orchestration' {
    BeforeAll {
        function winget.exe { throw 'Unmocked winget call in a test.' }
    }

    AfterAll {
        Remove-Variable AppxPayloadTestState -Scope Global -ErrorAction SilentlyContinue
    }

    BeforeEach {
        $global:AppxPayloadTestState = @{
            Calls = [System.Collections.Generic.List[object]]::new()
            ExitCode = 0
            ThrowOnDownload = $false
        }
        $manifestPath = Join-Path $TestDrive 'apps.json'
        $downloadPath = Join-Path $TestDrive 'Payload With Spaces'
        @{ apps = @(@{ Name = 'Windows Test App'; Id = '9TEST1234567' }) } |
            ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifestPath
        Mock Get-Command { [pscustomobject]@{ Source = 'winget.exe' } } -ParameterFilter { $Name -eq 'winget.exe' }
        Mock winget.exe {
            if ($args -contains '--help') {
                $global:LASTEXITCODE = 0
                '--download-directory --skip-license'
                return
            }
            $global:AppxPayloadTestState.Calls.Add(@($args))
            if ($global:AppxPayloadTestState.ThrowOnDownload) { throw 'Simulated launch failure' }
            $global:LASTEXITCODE = $global:AppxPayloadTestState.ExitCode
        }
    }

    It 'preserves spaces, uses exact IDs, and retrieves licenses by default' {
        $result = @(& $DownloaderPath -ManifestPath $manifestPath -DownloadPath $downloadPath)
        $LASTEXITCODE | Should -Be 0
        $result.Count | Should -Be 1
        $result[0].Status | Should -Be 'Success'
        $arguments = $global:AppxPayloadTestState.Calls[0]
        $arguments | Should -Contain (Join-Path $downloadPath 'Windows Test App')
        $arguments | Should -Contain '--exact'
        $arguments | Should -Not -Contain '--skip-license'
        $arguments[[array]::IndexOf($arguments, '--architecture') + 1] | Should -Be 'x64'
    }

    It 'allows explicit license omission and parameter overrides' {
        & $DownloaderPath -ManifestPath $manifestPath -DownloadPath $downloadPath -SkipLicense -Architecture arm64 -Source private | Out-Null
        $arguments = $global:AppxPayloadTestState.Calls[0]
        $arguments | Should -Contain '--skip-license'
        $arguments[[array]::IndexOf($arguments, '--architecture') + 1] | Should -Be 'arm64'
        $arguments[[array]::IndexOf($arguments, '--source') + 1] | Should -Be 'private'
    }

    It 'returns exit 1 and a failed result for a nonzero native exit' {
        $global:AppxPayloadTestState.ExitCode = -1978335189
        $result = & $DownloaderPath -ManifestPath $manifestPath -DownloadPath $downloadPath
        $LASTEXITCODE | Should -Be 1
        $result.Status | Should -Be 'Failed'
        $result.ExitCode | Should -Be $global:AppxPayloadTestState.ExitCode
    }

    It 'continues collecting results after process launch failures' {
        $global:AppxPayloadTestState.ThrowOnDownload = $true
        @{ apps = @(@{ Name = 'First App'; Id = 'FIRST' }, @{ Name = 'Second App'; Id = 'SECOND' }) } |
            ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifestPath
        $results = @(& $DownloaderPath -ManifestPath $manifestPath -DownloadPath $downloadPath)
        $LASTEXITCODE | Should -Be 1
        $results.Count | Should -Be 2
        @($results | Where-Object Status -eq 'Failed').Count | Should -Be 2
    }

    It 'rejects folder traversal before running winget' {
        @{ apps = @(@{ Name = '..\Escape'; Id = 'INVALID' }) } |
            ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifestPath
        { & $DownloaderPath -ManifestPath $manifestPath -DownloadPath $downloadPath } | Should -Throw '*folder name*'
        Should -Invoke winget.exe -Times 0 -Exactly
    }

    It 'rejects duplicate entries before running winget' {
        @{ apps = @(@{ Name = 'First App'; Id = 'SAME' }, @{ Name = 'Second App'; Id = 'same' }) } |
            ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifestPath
        { & $DownloaderPath -ManifestPath $manifestPath -DownloadPath $downloadPath } | Should -Throw '*Duplicate*'
        Should -Invoke winget.exe -Times 0 -Exactly
    }

    It 'uses manifest defaults when parameters are absent' {
        @{ defaults = @{ architecture = 'arm64'; source = 'private' }; apps = @(@{ Name = 'Test'; Id = 'TEST' }) } |
            ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifestPath
        & $DownloaderPath -ManifestPath $manifestPath -DownloadPath $downloadPath | Out-Null
        $arguments = $global:AppxPayloadTestState.Calls[0]
        $arguments[[array]::IndexOf($arguments, '--architecture') + 1] | Should -Be 'arm64'
        $arguments[[array]::IndexOf($arguments, '--source') + 1] | Should -Be 'private'
    }

    It 'rejects invalid JSON and invalid app collections' -ForEach @(
        @{ Content = '{' },
        @{ Content = '{"apps":[]}' },
        @{ Content = '{"apps":{"Name":"Test","Id":"TEST"}}' },
        @{ Content = '{"apps":[{"Name":"Test"}]}' }
    ) {
        Set-Content -LiteralPath $manifestPath -Value $Content
        { & $DownloaderPath -ManifestPath $manifestPath -DownloadPath $downloadPath } | Should -Throw
        Should -Invoke winget.exe -Times 0 -Exactly
    }

    It 'validates every entry in the shipped manifest without downloads' {
        $shippedManifest = Join-Path $PSScriptRoot '..\StubApps.json'
        $manifest = Get-Content -LiteralPath $shippedManifest -Raw | ConvertFrom-Json
        $manifest.apps.Count | Should -Be 19
        $manifest.apps.Id | Should -Not -Contain '9NV2L4XVMCXM'
        $manifest.apps.Id | Should -Not -Contain 'XP89DCGQ3K6VLD'
        $manifest.apps.Name | Should -Not -Contain 'Power Automate'
        foreach ($app in $manifest.apps) { $app.Id | Should -Match '^[A-Z0-9]{12}$' }
        $results = @(& $DownloaderPath -ManifestPath $shippedManifest -DownloadPath $downloadPath)
        $LASTEXITCODE | Should -Be 0
        $results.Count | Should -Be $manifest.apps.Count
        $global:AppxPayloadTestState.Calls.Count | Should -Be $manifest.apps.Count
    }
}

Describe 'Installer orchestration with mocked DISM' {
    BeforeEach {
        $root = Join-Path $TestDrive ([guid]::NewGuid().ToString())
        $firstDirectory = Join-Path $root 'First App'
        $secondDirectory = Join-Path $root 'Second App'
        $firstDependencies = Join-Path $firstDirectory 'Dependencies'
        $secondDependencies = Join-Path $secondDirectory 'Dependencies'
        New-Item -Path $firstDependencies, $secondDependencies -ItemType Directory -Force | Out-Null
        New-TestPackage -Path (Join-Path $firstDirectory 'First.appx') -Name 'Contoso.First' | Out-Null
        New-TestPackage -Path (Join-Path $secondDirectory 'Second.msixbundle') -Name 'Contoso.Second' -Bundle | Out-Null
        New-TestPackage -Path (Join-Path $firstDependencies 'FirstFramework.msix') -Framework | Out-Null
        New-TestPackage -Path (Join-Path $secondDependencies 'SecondFramework.appx') -Framework | Out-Null
        $logDirectory = Join-Path $TestDrive "logs\$([guid]::NewGuid())"
        Mock Add-AppxProvisionedPackage { }
        Mock Get-AppxProvisionedPackage { [pscustomobject]@{ DisplayName = 'Contoso.First' } }
    }

    It 'installs both apps with only their own dependencies and creates the log directory' {
        $results = @(& $InstallerBody -SourcePath $root -LogDirectory $logDirectory -SkipLicense)
        $LASTEXITCODE | Should -Be 0
        $results.Count | Should -Be 2
        Test-Path -LiteralPath $logDirectory -PathType Container | Should -BeTrue
        Should -Invoke Add-AppxProvisionedPackage -Times 2 -Exactly
        Should -Invoke Add-AppxProvisionedPackage -Times 1 -Exactly -ParameterFilter {
            $PackagePath -like '*First.appx' -and $DependencyPackagePath.Count -eq 1 -and
            $DependencyPackagePath[0] -like '*FirstFramework.msix' -and $StubPackageOption -eq 'InstallFull'
        }
        Should -Invoke Add-AppxProvisionedPackage -Times 1 -Exactly -ParameterFilter {
            $PackagePath -like '*Second.msixbundle' -and $DependencyPackagePath.Count -eq 1 -and
            $DependencyPackagePath[0] -like '*SecondFramework.appx'
        }
    }

    It 'skips unprovisioned apps and their dependencies in update mode' {
        $results = @(& $InstallerBody -SourcePath $root -LogDirectory $logDirectory -Mode UpdateProvisioned -SkipLicense)
        $LASTEXITCODE | Should -Be 0
        @($results | Where-Object Status -eq 'Success').Count | Should -Be 1
        @($results | Where-Object Status -eq 'Skipped').Count | Should -Be 1
        Should -Invoke Add-AppxProvisionedPackage -Times 1 -Exactly -ParameterFilter { $PackagePath -like '*First.appx' }
        Should -Invoke Add-AppxProvisionedPackage -Times 0 -Exactly -ParameterFilter { $PackagePath -notlike '*First.appx' }
    }

    It 'skips everything when the provisioned inventory is empty' {
        Mock Get-AppxProvisionedPackage { }
        $results = @(& $InstallerBody -SourcePath $root -LogDirectory $logDirectory -Mode UpdateProvisioned)
        $LASTEXITCODE | Should -Be 0
        @($results | Where-Object Status -eq 'Skipped').Count | Should -Be 2
        Should -Invoke Add-AppxProvisionedPackage -Times 0 -Exactly
    }

    It 'continues after one DISM failure and exits nonzero' {
        Mock Add-AppxProvisionedPackage { throw 'Simulated failure' } -ParameterFilter { $PackagePath -like '*First.appx' }
        $results = @(& $InstallerBody -SourcePath $root -LogDirectory $logDirectory -SkipLicense)
        $LASTEXITCODE | Should -Be 1
        @($results | Where-Object Status -eq 'Failed').Count | Should -Be 1
        @($results | Where-Object Status -eq 'Success').Count | Should -Be 1
    }

    It 'rejects an empty payload tree' {
        $emptyRoot = New-Item -Path (Join-Path $TestDrive 'Empty') -ItemType Directory
        & $InstallerBody -SourcePath $emptyRoot.FullName -LogDirectory $logDirectory
        $LASTEXITCODE | Should -Be 1
        Should -Invoke Add-AppxProvisionedPackage -Times 0 -Exactly
    }
}