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

# Older OpenSSH versions refuse the entire configuration if it contains this file.
# See linux-scripts/ssh/openssh10.conf.
# Without it, RemoteCommand breaks "ssh host command" and the SSH remoting of JetBrains IDEs.
# On Windows, it must be visible only to the MSYS-based OpenSSH of Git for Windows,
# since Microsoft OpenSSH, e.g. 9.5 that is used by Claude Desktop, may be older than 10.0.
# Therefore, it is linked as "openssh10:msys.conf", which the SSH config includes.
# MSYS stores the ":" of file names as the private-use character U+F03A, and translates it back,
# but for Windows programs such as Microsoft OpenSSH, "openssh10:msys.conf" refers to
# an alternate data stream of the nonexistent file "openssh10", so they don't find it.
# Creating a file symlink requires administrator privileges or Developer Mode.
$OpenSSH10LegacyPath = "${ConfigDir}\openssh10.conf"
if ((Get-Item "${OpenSSH10LegacyPath}" -ErrorAction SilentlyContinue).LinkType -eq "SymbolicLink") {
    Show-Output "Removing the symlink openssh10.conf, since Microsoft OpenSSH older than 10.0 would refuse the configuration."
    Remove-Item "${OpenSSH10LegacyPath}"
}
$OpenSSH10ConfPath = "${ConfigDir}\openssh10$([char]0xF03A)msys.conf"
$SSHVersion = (& ssh -V 2>&1 | Out-String)
if (($SSHVersion -notmatch "for_Windows") -and ($SSHVersion -match "OpenSSH_(\d+)\.") -and ([int]$Matches[1] -ge 10)) {
    Show-Output "Creating symlink to the configuration for OpenSSH 10.0 or newer for the OpenSSH of Git for Windows."
    if (Test-Path -LiteralPath "${OpenSSH10ConfPath}") {
        Remove-Item -LiteralPath "${OpenSSH10ConfPath}"
    }
    New-Item -ItemType "SymbolicLink" -Path "${OpenSSH10ConfPath}" -Target "${GitPath}\linux-scripts\ssh\openssh10.conf"
} elseif ((Get-Item -LiteralPath "${OpenSSH10ConfPath}" -ErrorAction SilentlyContinue).LinkType -eq "SymbolicLink") {
    Show-Output "Removing symlink to the configuration for OpenSSH 10.0 or newer, since the ssh in PATH is not a supported version."
    Remove-Item -LiteralPath "${OpenSSH10ConfPath}"
}

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
