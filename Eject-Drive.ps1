<#
.SYNOPSIS
    Eject a disk such as a USB drive safely.
.LINK
    https://serverfault.com/a/580298
    https://community.spiceworks.com/topic/2120667-how-to-safely-remove-usb-key-in-windows-server-core?page=1#entry-7617376
.PARAMETER Name
    Drive letter
#>

param(
    [Parameter(Mandatory=$true)][char]$DriveLetter
)

# This works on Windows 10 but seems to have no effect on Hyper-V Server 2019
# $driveEject = New-Object -comObject Shell.Application
# $driveEject.Namespace(17).ParseName("${DriveLetter}:").InvokeVerb("Eject")

$vol = Get-WmiObject -Class Win32_Volume | where{$_.Name -eq "${DriveLetter}:\"}
$vol.DriveLetter = $null
$vol.Put()
$vol.Dismount($false, $false)
