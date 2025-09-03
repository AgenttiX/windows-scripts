[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSReviewUnusedParameter", "Elevated", Justification="Used in utils")]
param(
    [Parameter(mandatory=$true)][string]$RepoName,
    [Parameter(mandatory=$false)][switch]$Elevated
)

$RepoPath = (Get-Item "${PSScriptRoot}").Parent.FullName
$GitPath = (Get-Item "${PSScriptRoot}").Parent.Parent.FullName
$UtilsPath = "${RepoPath}\Utils.ps1"
if (-not (Test-Path "${UtilsPath}")) {
    Write-Host "Utils.ps1 was not found at ${UtilsPath}"
    return
}
. "${UtilsPath}"

$ConfigDir = "${GitPath}\${RepoName}\ssh"
if (-not (Test-Path "${ConfigDir}")) {
    Write-Host "The SSH configuration was not found at ${ConfigDir}"
    return
}

$ControlMasterDir = "${ConfigDir}\controlmasters"
if (-not (Test-Path "${ControlMasterDir}")) {
    Write-Host "The SSH ControlMasters directory was not found. Creating it at `"${ControlMasterDir}`"."
    New-Item -Path "${ControlMasterDir}" -ItemType "directory"
}

$SSHDir = "${HOME}\.ssh"

# This should be as early as possible to avoid loading the function definitions etc. twice.
# Elevate($myinvocation.MyCommand.Definition)

# New-Item -ItemType "directory" -Path "${SSHDir}"
# Write-Output "If you get permission errors, enable developer mode in Windows settings."
# New-Item -ItemType "SymbolicLink" -Path "${SSHDir}\authorized_keys" -Target "${PSScriptRoot}\authorized_keys"
# New-Item -ItemType "SymbolicLink" -Path "${SSHDir}\config" -Target "${PSScriptRoot}\config"

Show-Output "Creating junction to SSH config directory."
New-Junction -Path "${SSHDir}" -Target "${ConfigDir}"

# $Service = Get-Service -Name "ssh-agent" -ErrorAction SilentlyContinue
# if ($Service.Length -gt 0) {
#     Remove-Service -Name "ssh-agent"
# }
# New-Service -Name "ssh-agent" -BinaryPathName "${env:ProgramFiles}\Git\cmd\start-ssh-agent.cmd" -StartupType "Automatic"
# Start-Service -Name "ssh-agent"
