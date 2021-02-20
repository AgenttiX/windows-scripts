<#
.SYNOPSIS
    Utility methods for scripts
#>

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
