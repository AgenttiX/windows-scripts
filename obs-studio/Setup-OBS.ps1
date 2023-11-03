<#
.SYNOPSIS
    Configure Open Broadcaster Software (OBS)
.LINK
    https://obsproject.com/
#>

. "$(Join-Path $((Get-Item "${PSScriptRoot}").Parent.FullName) "Utils.ps1")"

$OBSDirectory = "${env:AppData}\obs-studio"

Show-Output "Creating junction to OBS config directory."
New-Junction -Path "${OBSDirectory}" -Target "${PSScriptRoot}"
