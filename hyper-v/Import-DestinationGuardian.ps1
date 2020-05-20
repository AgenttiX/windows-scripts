<#
.SYNOPSIS
    Imports the Hyper-V host guardian configuration of another host.
    Run this on the source host for migrating virtual machines.
.PARAMETER Name
    The name of the destination guardian, e.g. DestinationGuardian
.PARAMETER Path
    The path to the .xml file of the destination guardian
.NOTES
    Based on
    https://threadfin.com/allowing-an-additional-host-to-run-a-vm-with-virtual-tpm/
    https://gist.github.com/larsiwer/0eef34728d6697f8d5b899b1d866e573#file-update-keyprotectorforallvms-ps1
#>

param (
    [Parameter(Mandatory=$true)][string]$Name,
    [Parameter(Mandatory=$true)][string]$Path
)

# Import the destination guardian on the source machine
Import-HgsGuardian -Path $Path -Name $Name -AllowExpired -AllowUntrustedRoot
