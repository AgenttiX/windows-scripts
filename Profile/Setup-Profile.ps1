Set-StrictMode -Latest
$ErrorActionPreference = "Stop"

. "$(Join-Path $((Get-Item "${PSScriptRoot}").Parent.FullName) "Utils.ps1")"

Show-Output "Configuring PowerShell profile."

$ProfileDir = Split-Path -Path "${profile}" -Parent

$CreateNew = $true
if (Test-Path -Path "${ProfileDir}") {
    if ((Get-Item -Path "${ProfileDir}" -Force).LinkType -eq "Junction") {
        $CreateNew = $false
    } else {
        Move-Item -Path "${ProfileDir}" "${ProfileDir}-old"
    }
}

if ($CreateNew) {
    New-item -ItemType "Junction" -Path "${ProfileDir}" -Target "${PSScriptRoot}"
}

Show-Output "PowerShell profile configuration is ready."
