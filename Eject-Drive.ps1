<#
.SYNOPSIS
    Eject a disk such as a USB drive safely.
.LINK
    https://superuser.com/questions/1750399/eject-dismount-a-drive-using-powershell
    https://serverfault.com/a/580298
    https://community.spiceworks.com/topic/2120667-how-to-safely-remove-usb-key-in-windows-server-core?page=1#entry-7617376
.PARAMETER DriveLetter
    Drive letter, e.g. C
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingWMICmdlet", "", Justification="Used on purpose since the WMI object does not have the proper methods")]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSReviewUnusedParameter", "DriveLetter", Justification="Set to null on purpose")]
param(
    [Parameter(Mandatory=$true)][char]$DriveLetter
)

# This works on Windows 10 but seems to have no effect on Hyper-V Server 2019
# $driveEject = New-Object -comObject Shell.Application
# $driveEject.Namespace(17).ParseName("${DriveLetter}:").InvokeVerb("Eject")

# WMI should be replaced with CIM
# when a solution is found for calling the proper ejection methdods on the CIM object.
# $vol = Get-CimInstance Win32_Volume | Where-Object {$_.Name -eq "${DriveLetter}:\"}
$vol = Get-WmiObject -Class Win32_Volume | Where-Object {$_.Name -eq "${DriveLetter}:\"}
$vol.DriveLetter = $null
$vol.Put()
$vol.Dismount($false, $false)
