<#
.SYNOPSIS
    Configure FIDO2 and SSH for a YubiKey
.PARAMETER Elevated
    This parameter is for internal use to check whether an UAC prompt has already been attempted.
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSReviewUnusedParameter", "Elevated", Justification="Used in utils")]
param(
    [switch]$Elevated
)

. "${PSScriptRoot}\Utils.ps1"
Elevate($myinvocation.MyCommand.Definition)
. ".\venv\Scripts\activate.ps1"

Show-Output -ForegroundColor Cyan "Changing the FIDO2 PIN"
ykman fido access change-pin
Show-Output -ForegroundColor Cyan "Creating the SSH key"
ssh-keygen -t ed25519-sk -O resident -O verify-required
