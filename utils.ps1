<#
.SYNOPSIS
    Utility methods for scripts
#>

# Configure PSScriptRoot variable for old PowerShell versions
# https://stackoverflow.com/a/8406946
if($PSVersionTable.PSVersion.Major -lt 3) {
    $invocation = (Get-Variable MyInvocation).Value
    $PSScriptRoot = $directorypath = Split-Path $invocation.MyCommand.Path
}

# Configuration to ensure that the script is run as an administrator
# https://superuser.com/questions/108207/how-to-run-a-powershell-script-as-administrator
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

function Elevate {
    Param ($command)
    if ((Test-Admin) -eq $false)  {
        if ($elevated)
        {
            # tried to elevate, did not work, aborting
        }
        else {
            Write-Host "This script requires admin access. Elevating."
            Write-Host "$command"
            Start-Process -FilePath powershell.exe -Verb RunAs -ArgumentList ('-NoProfile -NoExit -Command "cd {0}; {1}" -elevated' -f ($pwd, $command))
        }
    exit
    }
}

function Install-Chocolatey {
    param(
        [switch]$Force
    )
    if($Force -Or (-Not (Test-CommandExists "choco"))) {
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
    }
}

Function Test-CommandExists {
    Param ($command)
    $oldPreference = $ErrorActionPreference
    $ErrorActionPreference = "stop"
    try {
        if (Get-Command $command){RETURN $true}
    }
    catch {RETURN $false}
    finally {$ErrorActionPreference=$oldPreference}
}
