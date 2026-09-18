#Requires -Version 5.1

function Get-W365CaStringList {
    param([object]$Value)
    $Known = $Value -is [Collections.IList] -and @($Value).Count -le 1000
    if ($Known) {
        foreach ($Entry in $Value) {
            if ($Entry -isnot [string] -or $Entry.Length -gt 256 -or $Entry.Length -eq 0) { $Known = $false; break }
        }
    }
    return [pscustomobject]@{ Known = $Known; Values = $(if ($Known) { @($Value) } else { @() }) }
}

function Test-W365CaAppIdentifier {
    param([string]$Value)
    $Identifier = [guid]::Empty
    return [guid]::TryParse($Value, [ref]$Identifier)
}

function ConvertTo-W365ConditionalAccessEvidence {
    param([object[]]$Policies)
    if ($Policies.Count -gt 10000) { throw 'Conditional Access evidence row limit exceeded' }
    $Apps = [ordered]@{
        Windows365 = '0af06dc6-e4b5-4f28-818e-e78e62d137a5'
        AzureVirtualDesktop = '9cdead84-a844-4324-93f2-b2e6bb768d07'
        WindowsCloudLogin = '270efc09-cd0d-444b-a71f-39af4910ec45'
    }
    foreach ($Policy in $Policies) {
        $Applications = $Policy.conditions.applications
        $Included = Get-W365CaStringList $Applications.includeApplications
        $Excluded = Get-W365CaStringList $Applications.excludeApplications
        $KnownIncludes = $Included.Known -and @($Included.Values | Where-Object { $_ -cne 'All' -and -not (Test-W365CaAppIdentifier $_) }).Count -eq 0
        $KnownExcludes = $Excluded.Known -and @($Excluded.Values | Where-Object { -not (Test-W365CaAppIdentifier $_) }).Count -eq 0
        $Targets = [ordered]@{}
        foreach ($Name in $Apps.Keys) {
            $AppId = $Apps[$Name]
            $Targets[$Name] = if ($Excluded.Known -and $Excluded.Values -contains $AppId) { 'Excluded' }
                elseif (-not $KnownIncludes -or -not $KnownExcludes -or $null -ne $Applications.applicationFilter -or $Applications.includeUserActions) { 'Unknown' }
                elseif ($Included.Values -ccontains 'All' -or $Included.Values -contains $AppId) { 'Included' }
                else { 'NotDirectlyIncluded' }
        }
        $Grant = $Policy.grantControls
        $Controls = Get-W365CaStringList $Grant.builtInControls
        $Custom = Get-W365CaStringList $Grant.customAuthenticationFactors
        $Terms = Get-W365CaStringList $Grant.termsOfUse
        $KnownControls = $Controls.Known -and @($Controls.Values | Where-Object { $_ -cnotin @('block','mfa','compliantDevice','domainJoinedDevice','approvedApplication','compliantApplication','passwordChange','riskRemediation') }).Count -eq 0
        $Mfa = 'Unknown'
        if ($KnownControls -and $Grant.operator -cin @('AND','OR')) {
            if ($Controls.Values -ccontains 'block') { $Mfa = 'BlockControlPresent' }
            elseif ($Controls.Values -ccontains 'mfa') {
                if ($Grant.operator -ceq 'AND') { $Mfa = 'RequiredInPolicy' }
                elseif (@($Controls.Values).Count -eq 1 -and $Custom.Known -and @($Custom.Values).Count -eq 0 -and $Terms.Known -and @($Terms.Values).Count -eq 0 -and $null -eq $Grant.authenticationStrength) { $Mfa = 'RequiredInPolicy' }
                else { $Mfa = 'AlternativeOrUnresolved' }
            } elseif ($null -ne $Grant.authenticationStrength) { $Mfa = 'AuthenticationStrengthNotEvaluated' }
            else { $Mfa = 'NotDeclared' }
        }
        $Frequency = $Policy.sessionControls.signInFrequency
        $FrequencyEnabled = if ($Frequency.isEnabled -is [bool]) { $Frequency.isEnabled } else { $null }
        $FrequencyMode = 'Unknown'
        $FrequencyValue = $null
        $FrequencyUnit = $null
        if ($null -ne $FrequencyEnabled) {
            if (-not $FrequencyEnabled) { $FrequencyMode = 'Disabled' }
            elseif ($Frequency.frequencyInterval -ceq 'everyTime' -and $Frequency.authenticationType -cin @('primaryAndSecondaryAuthentication','secondaryAuthentication')) { $FrequencyMode = 'EveryTime' }
            elseif ($Frequency.frequencyInterval -ceq 'timeBased' -and $Frequency.type -cin @('hours','days') -and ($Frequency.value -is [int] -or $Frequency.value -is [long]) -and $Frequency.value -gt 0 -and $Frequency.value -le [int]::MaxValue) {
                $FrequencyMode = 'TimeBased'
                $FrequencyValue = $Frequency.value
                $FrequencyUnit = $Frequency.type
            }
        }
        [pscustomobject]@{
            PolicyId = $(if ($Policy.id -is [string] -and $Policy.id -cmatch '^[A-Za-z0-9-]{1,128}$') { $Policy.id } else { 'Unidentified' })
            State = $(if ($Policy.state -cin @('enabled','disabled','enabledForReportingButNotEnforced')) { $Policy.state } else { 'Unknown' })
            ApplicationTargets = $Targets
            MfaRequirement = $Mfa
            SignInFrequencyEnabled = $FrequencyEnabled
            SignInFrequencyMode = $FrequencyMode
            SignInFrequencyValue = $FrequencyValue
            SignInFrequencyUnit = $FrequencyUnit
            TokenProtection = 'NotCollectedByV1Contract'
            EffectiveScope = 'NotEvaluated'
        }
    }
}