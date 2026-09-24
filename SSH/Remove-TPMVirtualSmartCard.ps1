<#
.SYNOPSIS
    Destroy a TPM virtual smart card
.DESCRIPTION
    The keys on the card are lost permanently.
    The script asks for confirmation before destroying the card.
.PARAMETER InstanceId
    Instance ID of the card, e.g. "ROOT\SMARTCARDREADER\0000". The instance IDs are shown by Get-TPMVirtualSmartCard.ps1.
.PARAMETER Elevated
    This parameter is for internal use to check whether an UAC prompt has already been attempted.
.EXAMPLE
    .\Remove-TPMVirtualSmartCard.ps1 -InstanceId "ROOT\SMARTCARDREADER\0000"
.LINK
    https://learn.microsoft.com/en-us/windows/security/identity-protection/virtual-smart-cards/virtual-smart-card-tpmvscmgr
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSReviewUnusedParameter", "Elevated", Justification="Used in utils")]
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][ValidatePattern('^ROOT\\SMARTCARDREADER\\\d+$')][string]$InstanceId,
    [switch]$Elevated
)

Set-StrictMode -Version 3.0
. "${PSScriptRoot}\..\Utils.ps1"
. "${PSScriptRoot}\SmartCardUtils.ps1"

# Check the card before elevating to show the error in the current window.
$Card = @(Get-TPMVirtualSmartCardDevice | Where-Object { $_.InstanceId -eq $InstanceId })
if ($Card.Count -eq 0) {
    Show-Output "The TPM virtual smart card `"${InstanceId}`" was not found. Use Get-TPMVirtualSmartCard.ps1 to list the cards." -ForegroundColor Red
    exit 1
}

Elevate(Get-ElevateCommand -ScriptPath $myinvocation.MyCommand.Definition -BoundParameters $PSBoundParameters)

Show-Output "Destroying the TPM virtual smart card:" -ForegroundColor Cyan
Show-Output ($Card | Format-Table -AutoSize | Out-String).TrimEnd()
if (-not (Get-YesNo "The keys on the card will be lost permanently. Are you sure?")) {
    exit 0
}
& tpmvscmgr.exe destroy /instance "${InstanceId}"
if ($LASTEXITCODE -ne 0) {
    Show-Output "Destroying the virtual smart card failed with exit code ${LASTEXITCODE}." -ForegroundColor Red
    exit 1
}
Show-Output "The virtual smart card was destroyed." -ForegroundColor Green
Show-Output "If it had an SSH key, also remove its certificate from the personal certificate store (certmgr.msc),"
Show-Output "the public key from the Git server and ~/.ssh, and the settings from `"${SmartCardSSHDir}`"."
