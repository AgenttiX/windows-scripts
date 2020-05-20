<#
.SYNOPSIS
    Adds the key protector of the destination host guardian for migrating Hyper-V virtual machines.
    Run this on the source host.
.PARAMETER Name
    The name of the destination guardian, e.g. DestinationGuardian
.NOTES
    Based on
    https://threadfin.com/allowing-an-additional-host-to-run-a-vm-with-virtual-tpm/
    https://gist.github.com/larsiwer/0eef34728d6697f8d5b899b1d866e573#file-update-keyprotectorforallvms-ps1
#>

param (
    [Parameter(Mandatory=$true)][string]$Name
)

$destinationguardianname = $Name

# Get destination guardian
$destinationguardian = Get-HgsGuardian -Name $destinationguardianname

# Check if system is running in HGS local mode
If ((Get-HgsClientConfiguration | select -ExpandProperty Mode) -ne "Local")
{
    throw "HGS local mode required to update the key protector"
}

# Loop through all VMs existing on the local system
foreach ($vm in (Get-VM))
{
    # If the VM has the vTPM enabled, update the key protector
    If ((Get-VMSecurity -VM $vm).TpmEnabled -eq $true)
    {
        # Retrieve the current key protector for the virtual machine
        $keyprotector = ConvertTo-HgsKeyProtector -Bytes (Get-VMKeyProtector -VM $vm)

        # Check if the current system has the right owner keys present
        If ($keyprotector.Owner.HasPrivateSigningKey)
        {
            # Add the destination UntrustedGuardian to the key protector
            $newkeyprotector = Grant-HgsKeyProtectorAccess -KeyProtector $keyprotector -Guardian $destinationguardian `
                                                           -AllowUntrustedRoot -AllowExpired

            Write-Output "Updating key protector for $($vm.Name)"
            # Apply the updated key protector to VM
            Set-VMKeyProtector -VM $vm -KeyProtector $newkeyprotector.RawData
        }
        else
        {
            # Owner key information is not present
            Write-Warning "Skipping $($vm.Name) - Owner key information is not present"
        }
    }
}
