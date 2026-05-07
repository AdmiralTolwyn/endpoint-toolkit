# PSScriptAnalyzer settings for DeviceDecommissioner
# All suppressions are intentional — see justification comments.
@{
    ExcludeRules = @(
        # -- Architectural / Design --
        'PSAvoidGlobalVars'                       # WPF GUI tool; $Global: is the only reliable scope for DispatcherTimer callbacks and background runspace results.
        'PSUseDeclaredVarsMoreThanAssignments'     # Many variables are set once and read by XAML bindings or background callbacks.
        'PSUseShouldProcessForStateChangingFunctions'  # Internal GUI helpers, not cmdlets. ShouldProcess would be ignored in a WPF click handler.
        'PSAvoidUsingWriteHost'                    # Write-Host is used intentionally for console-output fallback alongside the WPF log panel.
        'PSReviewUnusedParameter'                  # Parameters like $Detail are used conditionally; analyzer can't trace through switch blocks.
        'PSUseApprovedVerbs'                       # Internal functions (Load-, Apply-, Refresh-, Edit-) are not exported cmdlets. Approved verbs would hurt readability for zero benefit.
        'PSUseSingularNouns'                       # Same — internal functions named for clarity (Get-DefaultSettings, Reset-AllCards, etc.).
        'PSUseBOMForUnicodeEncodedFile'            # UTF-8 without BOM is standard for Git repos. BOM causes issues with some editors and CI tools.

        # -- Intentional patterns --
        'PSAvoidUsingEmptyCatchBlock'              # Used for fire-and-forget cleanup: DispatcherTimer.Stop(), Get-MgContext, etc. Logging the error in these paths would be noise.
        'PSPossibleIncorrectComparisonWithNull'    # $raw.Prop -ne $null is used in hashtable value checks where left-side null doesn't improve clarity.
    )
}
