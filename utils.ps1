<#
.SYNOPSIS
    Utility methods for scripts
#>

# Compatibility for old PowerShell versions
if($PSVersionTable.PSVersion.Major -lt 3) {
    # Configure PSScriptRoot variable for old PowerShell versions
    # https://stackoverflow.com/a/8406946
    $invocation = (Get-Variable MyInvocation).Value
    $PSScriptRoot = $directorypath = Split-Path $invocation.MyCommand.Path

    # Implement missing features

    function Invoke-WebRequest {
        <#
        .SYNOPSIS
            Download a file from the web
        .LINK
            https://stackoverflow.com/a/43544903
        #>
        param (
            [Parameter(Mandatory=$true)] [string]$Uri,
            [Parameter(Mandatory=$true)] [string]$OutFile
        )
        $WebClient_Obj = New-Object System.Net.WebClient
        # TODO: this fails for some reason
        $WebClient_Obj.DownloadFile($Uri, $OutFile)
    }

    function Expand-Archive {
        <#
        .SYNOPSIS
            Zip file decompression
        .DESCRIPTION
            TODO: test that this works
        .LINK
            https://stackoverflow.com/a/54687028
        #>
        param(
            [Parameter(Mandatory=$true)] [string]$Path,
            [Parameter(Mandatory=$true)] [string]$DestinationPath
        )
        $shell_ComObject = New-Object -ComObject shell.application
        $zip_file = $shell_ComObject.namespace($Path)
        $folder = $shell_ComObject.namespace($DestinationPath)
        $folder.Copyhere($zip_file.items())
    }
}

# These have to be after the compatibility section so that $PSScriptRoot is defined
# $RepoPath = Split-Path $PSScriptRoot -Parent
$RepoPath = $PSSCriptRoot
$StartPath = Get-Location
New-Item -Path "$RepoPath" -Name "downloads" -ItemType "directory" -Force
$Downloads = "${RepoPath}\downloads"

function Elevate {
    param(
        [Parameter(Mandatory=$true)][string]$command
    )
    if ((Test-Admin) -eq $false)  {
        if ($elevated)
        {
            Write-Host "Elevation did not work."
        }
        else {
            Write-Host "This script requires admin access. Elevating."
            Write-Host "$command"
            # Use newer PowerShell if available.
            if (Test-CommandExists "pwsh") {$shell = "pwsh"} else {$shell = "powershell"}
            Start-Process -FilePath "$shell" -Verb RunAs -ArgumentList ('-NoProfile -NoExit -Command "cd {0}; {1}" -elevated' -f ($pwd, $command))
        }
    exit
    }
}

function GitPull {
    # Git pull should be run before elevating
    if (-Not ($elevated)) {
        git pull
    }
}

function Install-Chocolatey {
    param(
        [switch]$Force
    )
    if($Force -Or (-Not (Test-CommandExists "choco"))) {
        Write-Host "Installing the Chocolatey package manager."
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
    }
}

function Install-Geekbench {
    param(
        [string[]]$GeekbenchVersions = @("5.4.4", "4.4.4", "3.4.4", "2.4.3")
    )
    foreach ($Version in $GeekbenchVersions) {
        Write-Host "Downloading Geekbench ${Version}"
        $Filename = "Geekbench-$Version-WindowsSetup.exe"
        $Url = "https://cdn.geekbench.com/$Filename"
        Invoke-WebRequest -Uri "$Url" -OutFile "${Downloads}\${Filename}"
    }
    foreach ($Version in $GeekbenchVersions) {
        Write-Host "Installing Geekbench ${Version}"
        Start-Process -NoNewWindow -Wait "${Downloads}\Geekbench-${Version}-WindowsSetup.exe"
    }
}

function Install-PTS {
    param(
        [string]$PTS_version = "10.8.1"
    )
    $PTS = "${Env:SystemDrive}\phoronix-test-suite\phoronix-test-suite.bat"
    if (-Not (Test-Path "$PTS")) {
        Write-Host "Downloading Phoronix Test Suite (PTS)"
        Invoke-WebRequest -Uri "https://github.com/phoronix-test-suite/phoronix-test-suite/archive/v${PTS_version}.zip" -OutFile "${Downloads}\phoronix-test-suite-${PTS_version}.zip"
        Write-Host "Extracting Phoronix Test Suite (PTS)"
        Expand-Archive -LiteralPath "${Downloads}\phoronix-test-suite-${PTS_version}.zip" -DestinationPath "${Downloads}\phoronix-test-suite-${PTS_version}" -Force
        # The installation script needs to be executed in its directory
        Set-Location "${Downloads}\phoronix-test-suite-${PTS_version}\phoronix-test-suite-${PTS_version}"
        Write-Host "Installing Phoronix Test Suite (PTS)."
        Write-Host "You may get prompts asking whether you accept the EULA and whether to send anonymous usage statistics."
        Write-Host "Please select yes to the former and preferably to the latter as well."
        & ".\install.bat"
        Set-Location "$StartPath"
        if (-Not (Test-Path "$PTS")) {
            Write-Host -ForegroundColor Red "Phoronix Test Suite (PTS) installation failed."
            exit 1
        }
        Write-Host "Phoronix Test Suite (PTS) has been installed. It is highly recommended that you log in now so that you can manage the uploaded results."
        & "$PTS" openbenchmarking-login
        Write-Host "You must now configure the batch mode according to your preferences. Some of the settings may get overwritten later by this script."
        & "$PTS" batch-setup
    }
    & "$PTS" openbenchmarking-refresh
}

function Test-Admin {
    # Test whether the script is being run as an administrator
    # https://superuser.com/questions/108207/how-to-run-a-powershell-script-as-administrator
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
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
