<#
.SYNOPSIS
    Eject a disk such as a USB drive safely.
.LINK
    https://serverfault.com/a/580298
.PARAMETER Name
    Drive letter
#>

param(
    [Parameter(Mandatory=$true)][char]$DriveLetter
)

$driveEject = New-Object -comObject Shell.Application
$driveEject.Namespace(17).ParseName("${DriveLetter}:").InvokeVerb("Eject")
