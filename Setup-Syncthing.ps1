<#
.SYNOPSIS
    Setup Syncthing using a standalone managed service account (sMSA)
.LINK
    https://docs.syncthing.net/users/autostart.html#run-as-a-service-independent-of-user-login
.LINK
    https://cybergladius.com/secure-windows-scheduled-tasks-with-managed-service-accounts/
#>

# It may be possible to use psexec to run a command prompt as a service account
# https://learn.microsoft.com/en-us/sysinternals/downloads/psexec

$ErrorActionPreference = "Stop"

. "${PSScriptRoot}\Utils.ps1"

Install-Chocolatey
choco install syncthing -y

$Domain = Get-ADDomain
$AccountName = "svc-$("${env:ComputerName}".ToLower())-sync"
$AccountFullName = "$($Domain.NetBIOSName)\${AccountName}"
Show-Output "Using account name: ${AccountName}"

try {
    # Get-LocalUser "Syncthing"
    $Account = Get-ADServiceAccount -Identity "${AccountName}"
} catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
    Show-Output "The Syncthing user was not found. Creating."

    # If you get the error "New-ADServiceAccount : Key does not exist"
    # Run this:
    # Add-KdsRootKey -EffectiveTime ((get-date).addhours(-10))
    # https://learn.microsoft.com/en-us/windows-server/security/group-managed-service-accounts/create-the-key-distribution-services-kds-root-key

    $Account = New-ADServiceAccount -Name "${AccountName}" -RestrictToSingleComputer
}
Show-Output "Installing the AD Service account"
Install-ADServiceAccount $Account
Test-ADServiceAccount $AccountName

$SyncthingRoot = "${env:HomeDrive}\Syncthing"
if (Test-Path "${SyncthingRoot}") {
    Show-Output "Syncthing folder exists at ${SyncthingRoot}"
} else {
    Show-Output "Creating Syncthing folder at ${SyncthingRoot}"
    New-Item -Path "${SyncthingRoot}" -ItemType Directory
}

# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-acl?view=powershell-7.4#example-5-grant-administrators-full-control-of-the-file
$NewAcl = Get-Acl -Path "${SyncthingRoot}"
$NTAccount = New-Object System.Security.Principal.NTAccount($Domain.NetBIOSName, "${AccountName}`$")
$NewAcl.SetOwner($NTAccount)
$AccessRule = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule("${AccountName}`$", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
$NewAcl.SetAccessRule($AccessRule)
Set-Acl -Path "${SyncthingRoot}" -AclObject $NewAcl

# $Credential = New-Object System.Management.Automation.PSCredential("${AccountName}`$", (New-Object System.Security.SecureString))
# $Credential = Get-Credential -UserName $AccountName
# echo $AccountFullName
# echo $Credential

# try {
#     Get-Service "Syncthing"
#     sc.exe stop "Syncthing"
#     sc.exe delete "Syncthing"
#     Start-Sleep -Seconds 5
# } catch {}

# This results in error 1053 when starting
# New-Service -Name "Syncthing" -BinaryPathName "${env:ProgramData}\chocolatey\bin\syncthing.exe --no-console --no-restart --no-browser --home=`"${SyncthingRoot}`""

# .\nssm.exe install "Syncthing"

# Show-Output "You have to configure the account for the service manually. Go to services.msc -> Syncthing -> Properties -> Log On and use the account ${AccountFullName}`$ with an empty password."

Get-ScheduledTask -TaskName "Syncthing" -ErrorAction SilentlyContinue -OutVariable Task
if ($Task) {
    Show-Output "Unregistering old scheduled task"
    Unregister-ScheduledTask -TaskName "Syncthing"
    Start-Sleep -Seconds 1
}

$Trigger = New-ScheduledTaskTrigger -AtStartup
$Action = New-ScheduledTaskAction -Execute "${env:ProgramData}\chocolatey\bin\syncthing.exe" -Argument "--no-console --no-browser --home=`"${SyncthingRoot}`""
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -DontStopOnIdleEnd -ExecutionTimeLimit "00:00:00"  # -Priority 8
$Principal = New-ScheduledTaskPrincipal -UserId "${AccountFullName}`$" -LogonType Password -RunLevel Highest
Register-ScheduledTask -TaskName "Syncthing" -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principal
