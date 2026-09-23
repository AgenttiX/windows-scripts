# PSScriptAnalyzer configuration for this repository.
#
# Run the linter locally with:
#     pwsh ./Invoke-Lint.ps1
# and let it apply the automatic formatting fixes with:
#     pwsh ./Invoke-Lint.ps1 -Fix
#
# The same script is used by the CI workflow, so a clean local run means a clean CI run.
# Invoke-Lint.ps1 passes this file to PSScriptAnalyzer with -Settings.
#
# Failure threshold: by default only findings of the "Error" severity fail the run.
# Warnings are reported but non-fatal, because the existing scripts still have known
# warnings that cannot be fixed without behavioural changes. Tighten the threshold with
# "-FailOn Warning" once those have been cleaned up.
#
# Note: the formatting rules (PSPlaceOpenBrace, PSUseConsistentIndentation,
# PSUseConsistentWhitespace, ...) are deliberately NOT enabled here. They report ~290
# findings on the current code base, which would drown out the real problems.
# Invoke-Lint.ps1 -Fix applies them as automatic fixes instead of reporting them.

@{
    IncludeDefaultRules = $true

    ExcludeRules        = @(
        # These scripts are run interactively by an administrator who is already answering
        # prompts from the scripts themselves. Plumbing -WhatIf/-Confirm through them
        # would add noise without adding safety.
        'PSUseShouldProcessForStateChangingFunctions'

        # The repository idiom is "Show-Output "message"" rather than
        # "Show-Output -Message "message"".
        'PSAvoidUsingPositionalParameters'

        # The functions here are private helpers of standalone scripts, not exported
        # cmdlets of a published module. Several of the flagged names are product names
        # (for example Install-DigilentWaveforms) or genuinely operate on collections.
        # Utils.ps1 already suppresses this rule in three places, so this only makes the
        # existing decision repository-wide.
        'PSUseSingularNouns'

        # The compatibility rules below ship with hard-coded profiles of old Windows and
        # PowerShell builds. Measured on this repository they produce 58 (PSUseCompatibleCmdlets)
        # and 41 (PSUseCompatibleCommands + PSUseCompatibleTypes) findings, essentially all
        # of them false positives: the bundled profiles do not know about the Hyper-V, HGS,
        # DISM or Windows Forms assemblies that these scripts legitimately use on
        # Windows 11 / Windows Server 2025. PSUseCompatibleSyntax below covers the part
        # that can be checked reliably.
        'PSUseCompatibleCmdlets'
        'PSUseCompatibleCommands'
        'PSUseCompatibleTypes'
    )

    Rules               = @{
        # Language syntax must parse on both Windows PowerShell 5.1 and PowerShell 7,
        # as required by AGENTS.md. This rule is accurate (it only looks at syntax) and
        # currently reports zero findings, so it is a free regression guard.
        PSUseCompatibleSyntax = @{
            Enable         = $true
            TargetVersions = @(
                '5.1',
                '7.0'
            )
        }
    }
}
