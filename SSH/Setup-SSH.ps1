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

Show-Output "Creating junction to SSH config.d directory."
New-Junction -Path "${ConfigDir}\config.d" -Target "${GitPath}\linux-scripts\ssh\config.d"

# Do not run this code. The ssh-agent of Git for Windows is started by Profile/Profile.ps1 instead.
# This creates a service called "ssh-agent",
# which runs "C:\Program Files\Git\cmd\start-ssh-agent.cmd" as LocalSystem with automatic startup.
# The service fails to start, since a .cmd file is not a service executable,
# and it has the same name as the agent service of Windows OpenSSH.
# If it has been created, check it and remove it in an elevated shell with:
# Get-CimInstance Win32_Service -Filter "Name='ssh-agent'" | Select-Object Name, State, StartMode, PathName
# sc.exe delete ssh-agent
#
# $Service = Get-Service -Name "ssh-agent" -ErrorAction SilentlyContinue
# if ($Service.Length -gt 0) {
#     Remove-Service -Name "ssh-agent"
# }
# New-Service -Name "ssh-agent" -BinaryPathName "${env:ProgramFiles}\Git\cmd\start-ssh-agent.cmd" -StartupType "Automatic"
# Start-Service -Name "ssh-agent"
