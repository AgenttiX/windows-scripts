<#
.SYNOPSIS
    Optimize the VHDs of all VMs that are not currently running.
.NOTES
    Based on
    https://dscottraynsford.wordpress.com/2015/04/07/multiple-vhdvhdx-optimization-using-powershell-workflows/
#>

Get-VM | Where { $_.State -eq 'Off' } | Get-VMHardDiskDrive | Optimize-VHD -Mode Full
