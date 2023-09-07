<#
.SYNOPSIS
    Fully reset a YubiKey
.PARAMETER Elevated
    This parameter is for internal use to check whether an UAC prompt has already been attempted.
#>

param(
    [switch]$Elevated
)

. ".\Utils.ps1"

Elevate($myinvocation.MyCommand.Definition)

. ".\venv\Scripts\activate.ps1"

Show-Output -ForegroundColor Cyan "Resetting FIDO"
ykman fido reset
Show-Output -ForegroundColor Cyan "Resetting OATH"
ykman oath reset
Show-Output -ForegroundColor Cyan "Resetting OpenPGP"
ykman openpgp reset
Show-Output -ForegroundColor Cyan "Resetting OTP slot 1"
ykman otp delete 1
Show-Output -ForegroundColor Cyan "Resetting OTP slot 2"
ykman otp delete 2
Show-Output -ForegroundColor Cyan "Resetting PIV"
ykman piv reset
Show-Output -ForegroundColor Green "Reset ready"
