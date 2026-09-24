<#
.SYNOPSIS
    Create a TPM virtual smart card for storing an SSH key
.DESCRIPTION
    The virtual smart card is created with a random administrator key,
    so the PIN cannot be reset without the optional PUK.
    If the PIN is lost, destroy the card with Remove-TPMVirtualSmartCard.ps1 and create a new one.
    After creating the card, run New-SmartCardSSHKey.ps1 as a normal user to create the SSH key.
    The existing cards can be listed with Get-TPMVirtualSmartCard.ps1.

    Despite the name, the PIN does not have to be numeric.
    It can contain uppercase and lowercase letters, digits and special characters,
    but only printable ASCII characters are allowed, so e.g. letters with diacritics cannot be used.
    The maximum length is 127 characters. The PUK has to be at least 8 characters long.

    Note that Microsoft has deprecated virtual smart cards in favor of Windows Hello for Business and FIDO2,
    but tpmvscmgr is still included in Windows 11.
.PARAMETER Name
    Name of the virtual smart card
.PARAMETER MinPinLength
    Minimum length of the PIN
.PARAMETER Puk
    Create the card with a PIN unlock key (PUK), which can be used to unblock the PIN.
.PARAMETER Elevated
    This parameter is for internal use to check whether an UAC prompt has already been attempted.
.EXAMPLE
    .\New-TPMVirtualSmartCard.ps1 -Puk
.LINK
    https://learn.microsoft.com/en-us/windows/security/identity-protection/virtual-smart-cards/virtual-smart-card-tpmvscmgr
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSReviewUnusedParameter", "Elevated", Justification="Used in utils")]
[CmdletBinding()]
param(
    # Quotes would break the command line of the elevated process.
    [ValidatePattern('^[^"'']+$')][string]$Name = "SSH",
    [ValidateRange(8, 127)][int]$MinPinLength = 8,
    [switch]$Puk,
    [switch]$Elevated
)

Set-StrictMode -Version 3.0
. "${PSScriptRoot}\..\Utils.ps1"
. "${PSScriptRoot}\SmartCardUtils.ps1"
Elevate(Get-ElevateCommand -ScriptPath $myinvocation.MyCommand.Definition -BoundParameters $PSBoundParameters)

$Tpm = Get-Tpm
if (-not ($Tpm.TpmPresent -and $Tpm.TpmReady)) {
    Show-Output "The TPM is not present or not ready." -ForegroundColor Red
    exit 1
}

$ReadersBefore = @(Get-SmartCardReader)

Show-Output "Creating the TPM virtual smart card `"${Name}`"." -ForegroundColor Cyan
Show-Output "Use a PIN of ${MinPinLength}-127 characters. It can contain letters, digits and special characters, but only printable ASCII characters are allowed."
Show-Output "The TPM locks out the card after too many wrong attempts."
$Arguments = @("create", "/name", "${Name}", "/AdminKey", "RANDOM", "/PIN", "PROMPT", "/pinpolicy", "minlen", "${MinPinLength}", "/generate")
if ($Puk) {
    $Arguments += @("/PUK", "PROMPT")
}
& tpmvscmgr.exe @Arguments
if ($LASTEXITCODE -ne 0) {
    Show-Output "Creating the virtual smart card failed with exit code ${LASTEXITCODE}." -ForegroundColor Red
    exit 1
}

# The reader of the new card may take a moment to appear.
$NewReaders = @()
for ($i = 0; $i -lt 30; $i++) {
    $NewReaders = @(Get-SmartCardReader | Where-Object { $ReadersBefore -notcontains $_ })
    if ($NewReaders.Count -gt 0) {
        break
    }
    Start-Sleep -Seconds 1
}
if ($NewReaders.Count -ne 1) {
    Show-Output "Could not identify the reader of the new card. The readers are:" -ForegroundColor Yellow
    Get-SmartCardReader | ForEach-Object { Show-Output "  $_" }
    exit 1
}
Show-Output "The virtual smart card was created in the reader `"$($NewReaders[0])`"." -ForegroundColor Green
Show-Output "Next, run the following as a normal user (not elevated):"
Show-Output ".\New-SmartCardSSHKey.ps1 -ReaderName `"$($NewReaders[0])`""
