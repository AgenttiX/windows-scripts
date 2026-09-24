<#
.SYNOPSIS
    List the TPM virtual smart cards
.DESCRIPTION
    Shows the name, instance ID, PC/SC reader name and status of each TPM virtual smart card.
    The instance ID is needed for Remove-TPMVirtualSmartCard.ps1,
    and the reader name for New-SmartCardSSHKey.ps1.
    This does not require admin access.
.EXAMPLE
    .\Get-TPMVirtualSmartCard.ps1
#>

param()

Set-StrictMode -Version 3.0
. "${PSScriptRoot}\..\Utils.ps1"
. "${PSScriptRoot}\SmartCardUtils.ps1"

$Cards = @(Get-TPMVirtualSmartCardDevice)
if ($Cards.Count -eq 0) {
    Show-Output "No TPM virtual smart cards were found."
} else {
    Show-Output ($Cards | Format-Table -AutoSize | Out-String).TrimEnd()
}
