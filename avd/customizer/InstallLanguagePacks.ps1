<#
.SYNOPSIS
    Installs one or more Windows Display Languages (and their dependent FoDs) on the
    image so the captured AVD / Windows 365 host supports multiple UI languages.

.DESCRIPTION
    Wraps Install-Language (LanguagePackManagement module, Windows 11 / Server 2022+)
    and resolves friendly language display names to BCP-47 culture codes.

    The script:
      1. Disables the LanguageComponentsInstaller scheduled tasks that race the install
         and routinely cause ERROR_SHARING_VIOLATION (see internal bug 45044965).
      2. Installs each language with up to 5 retries to ride out transient package
         download / staging errors.
      3. Re-enables the LanguageComponentsInstaller tasks.
      4. Cleans up the C:\AVDImage staging folder if present.

.PARAMETER LanguageList
    One or more language display names from the validated set. Each is resolved to a
    BCP-47 culture code via the internal dictionary and passed to Install-Language.

.NOTES
    File:    avd/customizer/InstallLanguagePacks.ps1
    Author:  Anton Romanyuk
    Version: 2.0.0
    Context: Azure Image Builder / Packer customizer. Runs as SYSTEM.
    Requires: Windows 11 / Server 2022+ (LanguagePackManagement), PowerShell 5.1+, admin.

    Reference:
      https://learn.microsoft.com/powershell/module/languagepackmanagement/install-language

.DISCLAIMER
    This script is provided "AS IS" with no warranties and confers no rights.
    It is not supported under any Microsoft standard support program or service.
    Use of this script is entirely at your own risk. The customer is solely
    responsible for testing and validating this script in their environment
    before deploying to production.

.EXAMPLE
    .\InstallLanguagePacks.ps1 -LanguageList 'German (Germany)','French (France)'

.EXAMPLE
    .\InstallLanguagePacks.ps1 -LanguageList 'English (Australia)'
#>

#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet(
        'Arabic (Saudi Arabia)','Bulgarian (Bulgaria)','Chinese (Simplified, China)',
        'Chinese (Traditional, Taiwan)','Croatian (Croatia)','Czech (Czech Republic)',
        'Danish (Denmark)','Dutch (Netherlands)','English (United Kingdom)',
        'Estonian (Estonia)','Finnish (Finland)','French (Canada)','French (France)',
        'German (Germany)','Greek (Greece)','Hebrew (Israel)','Hungarian (Hungary)',
        'Indonesian (Indonesia)','Italian (Italy)','Japanese (Japan)','Korean (Korea)',
        'Latvian (Latvia)','Lithuanian (Lithuania)','Norwegian, Bokmål (Norway)',
        'Polish (Poland)','Portuguese (Brazil)','Portuguese (Portugal)',
        'Romanian (Romania)','Russian (Russia)','Serbian (Latin, Serbia)',
        'Slovak (Slovakia)','Slovenian (Slovenia)','Spanish (Mexico)','Spanish (Spain)',
        'Swedish (Sweden)','Thai (Thailand)','Turkish (Turkey)','Ukrainian (Ukraine)',
        'English (Australia)','English (United States)'
    )]
    [string[]]$LanguageList
)

$ErrorActionPreference = 'Stop'

# -----------------------------------------------------------------------------
# Display-name -> BCP-47 culture code map
# -----------------------------------------------------------------------------
$LanguagesDictionary = @{
    'Arabic (Saudi Arabia)'             = 'ar-SA'
    'Bulgarian (Bulgaria)'              = 'bg-BG'
    'Chinese (Simplified, China)'       = 'zh-CN'
    'Chinese (Traditional, Taiwan)'     = 'zh-TW'
    'Croatian (Croatia)'                = 'hr-HR'
    'Czech (Czech Republic)'            = 'cs-CZ'
    'Danish (Denmark)'                  = 'da-DK'
    'Dutch (Netherlands)'               = 'nl-NL'
    'English (United States)'           = 'en-US'
    'English (United Kingdom)'          = 'en-GB'
    'Estonian (Estonia)'                = 'et-EE'
    'Finnish (Finland)'                 = 'fi-FI'
    'French (Canada)'                   = 'fr-CA'
    'French (France)'                   = 'fr-FR'
    'German (Germany)'                  = 'de-DE'
    'Greek (Greece)'                    = 'el-GR'
    'Hebrew (Israel)'                   = 'he-IL'
    'Hungarian (Hungary)'               = 'hu-HU'
    'Indonesian (Indonesia)'            = 'id-ID'
    'Italian (Italy)'                   = 'it-IT'
    'Japanese (Japan)'                  = 'ja-JP'
    'Korean (Korea)'                    = 'ko-KR'
    'Latvian (Latvia)'                  = 'lv-LV'
    'Lithuanian (Lithuania)'            = 'lt-LT'
    'Norwegian, Bokmål (Norway)'        = 'nb-NO'
    'Polish (Poland)'                   = 'pl-PL'
    'Portuguese (Brazil)'               = 'pt-BR'
    'Portuguese (Portugal)'             = 'pt-PT'
    'Romanian (Romania)'                = 'ro-RO'
    'Russian (Russia)'                  = 'ru-RU'
    'Serbian (Latin, Serbia)'           = 'sr-Latn-RS'
    'Slovak (Slovakia)'                 = 'sk-SK'
    'Slovenian (Slovenia)'              = 'sl-SI'
    'Spanish (Mexico)'                  = 'es-MX'
    'Spanish (Spain)'                   = 'es-ES'
    'Swedish (Sweden)'                  = 'sv-SE'
    'Thai (Thailand)'                   = 'th-TH'
    'Turkish (Turkey)'                  = 'tr-TR'
    'Ukrainian (Ukraine)'               = 'uk-UA'
    'English (Australia)'               = 'en-AU'
}

$TemplateFolder = 'C:\AVDImage'

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO'
    )
    $ts    = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
    $color = switch ($Level) { 'WARN' {'Yellow'} 'ERROR' {'Red'} 'SUCCESS' {'Green'} default {'Gray'} }
    Write-Host "[$ts] [$Level] [InstallLanguagePacks] $Message" -ForegroundColor $color
}

function Set-LangCompInstallerTasks {
<#
.SYNOPSIS
    Enables or disables both LanguageComponentsInstaller scheduled tasks.
.DESCRIPTION
    These tasks race Install-Language and trigger ERROR_SHARING_VIOLATION on staged
    .cab files (internal repro: bug 45044965). We disable them for the duration of
    the install and re-enable them in the END block.
.PARAMETER Enable
    $true to enable, $false to disable.
#>
    param([Parameter(Mandatory)][bool]$Enable)

    $tasks = @(
        '\Microsoft\Windows\LanguageComponentsInstaller\Installation'
        '\Microsoft\Windows\LanguageComponentsInstaller\ReconcileLanguageResources'
    )
    foreach ($t in $tasks) {
        try {
            if ($Enable) {
                Enable-ScheduledTask -TaskName $t -ErrorAction Stop | Out-Null
                Write-Log "Enabled task $t"
            } else {
                Disable-ScheduledTask -TaskName $t -ErrorAction Stop | Out-Null
                Write-Log "Disabled task $t"
            }
        }
        catch {
            Write-Log "Could not toggle $t : $($_.Exception.Message)" -Level WARN
        }
    }
}

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
Write-Log "Starting InstallLanguagePacks customizer phase ($($LanguageList.Count) language(s))" -Level SUCCESS

Set-LangCompInstallerTasks -Enable:$false

$failures = 0
try {
    foreach ($Language in $LanguageList) {
        $code = $LanguagesDictionary[$Language]
        if (-not $code) {
            Write-Log "Unknown language '$Language' (no BCP-47 mapping)" -Level WARN
            $failures++
            continue
        }

        $installed = $false
        for ($i = 1; $i -le 5; $i++) {
            try {
                Write-Log "Installing language $code ($Language) - attempt $i/5"
                Install-Language -Language $code -ErrorAction Stop | Out-Null
                Write-Log "Installed $code" -Level SUCCESS
                $installed = $true
                break
            }
            catch {
                Write-Log "Install-Language $code attempt $i failed: $($_.Exception.Message)" -Level WARN
                Start-Sleep -Seconds (5 * $i)
            }
        }
        if (-not $installed) {
            Write-Log "Giving up on $code after 5 attempts" -Level ERROR
            $failures++
        }
    }
}
finally {
    if (Test-Path -LiteralPath $TemplateFolder) {
        Remove-Item -LiteralPath $TemplateFolder -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "Removed staging folder $TemplateFolder"
    }
    Set-LangCompInstallerTasks -Enable:$true
}

$stopwatch.Stop()
if ($failures -gt 0) {
    Write-Log "InstallLanguagePacks completed with $failures failure(s) in $($stopwatch.Elapsed)" -Level ERROR
    exit 1
}
Write-Log "InstallLanguagePacks completed in $($stopwatch.Elapsed)" -Level SUCCESS
exit 0
