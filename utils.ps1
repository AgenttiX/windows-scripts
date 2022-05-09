﻿<#
.SYNOPSIS
    Utility methods for scripts
#>

# Compatibility for old PowerShell versions
if($PSVersionTable.PSVersion.Major -lt 3) {
    # Configure PSScriptRoot variable for old PowerShell versions
    # https://stackoverflow.com/a/8406946/
    # https://stackoverflow.com/a/5466355/
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
        param(
            [Parameter(Mandatory=$true)][string]$Uri,
            [Parameter(Mandatory=$true)][string]$OutFile
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
            [Parameter(Mandatory=$true)][string]$Path,
            [Parameter(Mandatory=$true)][string]$DestinationPath
        )
        $shell_ComObject = New-Object -ComObject shell.application
        $zip_file = $shell_ComObject.namespace($Path)
        $folder = $shell_ComObject.namespace($DestinationPath)
        $folder.Copyhere($zip_file.items())
    }
}

# These have to be after the compatibility section so that $PSScriptRoot is defined
# $RepoPath = Split-Path $PSScriptRoot -Parent
$RepoPath = $PSScriptRoot
$StartPath = Get-Location
New-Item -Path "$RepoPath" -Name "downloads" -ItemType "directory" -Force | Out-Null
New-Item -Path "$RepoPath" -Name "logs" -ItemType "directory" -Force | Out-Null
$Downloads = "${RepoPath}\downloads"
$LogPath = "${RepoPath}\logs"

function Clear-Path {
    <#
    .SYNOPSIS
        Clear the path for e.g. unpacking a program.
    .DESCRIPTION
        Return values: 0 not cleared, 1 cleared, 2 no need
    #>
    # [Parameter(Mandatory=$true)]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory=$true)][string]$Path
    )
    if (Test-Path "${Path}") {
        $Decision = $Host.UI.PromptForChoice("The path `"${Path}`" already exists.", "Shall I clear and overwrite it?", ("y", "n"), 0)
        if ($Decision -eq 0) {
            Remove-Item -Path "${Path}" -Recurse
            return 1
        }
        return 0
    }
    return 2
}

function Elevate {
    <#
    .SYNOPSIS
        Elevate the current process to admin privileges.
    .LINK
        https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/start-process?view=powershell-7.2#example-5--start-powershell-as-an-administrator
    #>
    param(
        [Parameter(Mandatory=$true)][string]$command
    )
    if (! (Test-Admin))  {
        if ($elevated)
        {
            Show-Output "Elevation did not work."
        }
        else {
            Show-Output "This script requires admin access. Elevating."
            Show-Output "$command"
            # Use newer PowerShell if available.
            if (Test-CommandExists "pwsh") {$shell = "pwsh"} else {$shell = "powershell"}
            Start-Process -FilePath "$shell" -Verb RunAs -ArgumentList ('-NoProfile -NoExit -Command "cd {0}; {1}" -elevated' -f ($pwd, $command))
            Show-Output "The script has been started in another window. You can close this window now." -ForegroundColor Green
        }
    exit
    }
}

function Find-First {
    <#
    .SYNOPSIS
        Find the first file that matches the given filter
    .LINK
        https://stackoverflow.com/a/1500148/
    #>
    [OutputType([System.IO.FileInfo])]
    param(
        [string]$Filter,
        [string]$Path
    )
    return @(Get-ChildItem -Filter "${Filter}" -Path "${Path}")[0]
}

function Get-InstallBitness {
    [OutputType([string])]
    param(
        [string]$x86 = "x86",
        [string]$x86_64 = "x86_64",
        [switch]$Verbose = $true
    )
    if ([System.Environment]::Is64BitOperatingSystem) {
        if ($Verbose) {
            Show-Information "64-bit operating system detected. Installing 64-bit version."
        }
        return $x86_64
    }
    if ($Verbose) {
        Show-Information "32-bit operating system detected. Installing 32-bit version."
    }
    return $x86
}

function GitPull {
    # Git pull should be run before elevating
    if (! ($elevated)) {
        git pull
    }
}

function Install-Chocolatey {
    param(
        [switch]$Force
    )
    if($Force -Or (-Not (Test-CommandExists "choco"))) {
        Show-Output "Installing the Chocolatey package manager."
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
        Show-Output "Downloading Geekbench ${Version}"
        $Filename = "Geekbench-$Version-WindowsSetup.exe"
        $Url = "https://cdn.geekbench.com/$Filename"
        Invoke-WebRequest -Uri "$Url" -OutFile "${Downloads}\${Filename}"
    }
    foreach ($Version in $GeekbenchVersions) {
        Show-Output "Installing Geekbench ${Version}"
        Start-Process -NoNewWindow -Wait "${Downloads}\Geekbench-${Version}-WindowsSetup.exe"
    }
}

function Install-PTS {
    param(
        [string]$PTS_version = "10.8.1"
    )
    $PTS = "${Env:SystemDrive}\phoronix-test-suite\phoronix-test-suite.bat"
    if (-Not (Test-Path "$PTS")) {
        Show-Output "Downloading Phoronix Test Suite (PTS), as it seems not to be installed yet."
        Invoke-WebRequest -Uri "https://github.com/phoronix-test-suite/phoronix-test-suite/archive/v${PTS_version}.zip" -OutFile "${Downloads}\phoronix-test-suite-${PTS_version}.zip"
        Show-Output "Extracting Phoronix Test Suite (PTS)"
        Expand-Archive -LiteralPath "${Downloads}\phoronix-test-suite-${PTS_version}.zip" -DestinationPath "${Downloads}\phoronix-test-suite-${PTS_version}" -Force
        # The installation script needs to be executed in its directory
        Set-Location "${Downloads}\phoronix-test-suite-${PTS_version}\phoronix-test-suite-${PTS_version}"
        Show-Output "Installing Phoronix Test Suite (PTS)."
        Show-Output "You may get prompts asking whether you accept the EULA and whether to send anonymous usage statistics."
        Show-Output "Please select yes to the former and preferably to the latter as well."
        & ".\install.bat"
        Set-Location "$StartPath"
        if (-Not (Test-Path "$PTS")) {
            Show-Output -ForegroundColor Red "Phoronix Test Suite (PTS) installation failed."
            exit 1
        }
        Show-Output "Phoronix Test Suite (PTS) has been installed. It is highly recommended that you log in now so that you can manage the uploaded results."
        & "$PTS" openbenchmarking-login
        Show-Output "You must now configure the batch mode according to your preferences. Some of the settings may get overwritten later by this script."
        & "$PTS" batch-setup
    }
    & "$PTS" openbenchmarking-refresh
}

function Install-Winget {
    [OutputType([bool])]
    param()
    if (Test-CommandExists "winget") {
        return $true
    }
    if (Get-AppxPackage -Name "Microsoft.DesktopAppInstaller") {
        Show-Output "App Installer seems to be installed on your system, but Winget was not found."
        return $false
    }
    if (Get-AppxPackage -Name "Microsoft.WindowsStore") {
        Show-Output "App Installer appears not to be installed. Please close this window and install it from the Windows Store. Then restart this script."
        Start-Process -Path "https://www.microsoft.com/en-us/p/app-installer/9nblggh4nns1"
        $confirmation = Show-Output "If you know what you're doing, you may also continue by writing `"force`", but some features may be disabled.".
        while ($confirmation -ne "force") {
            $confirmation = Show-Output "Close this window or write `"force`" to continue."
        }
        return $false;
    }
    Show-Output "Cannot install App Installer, as Microsoft Store appears not to be installed. This is normal on servers. Winget will not be available."
    return $false
}

function New-Shortcut {
    param(
        [string]$SourceExe,
        [string]$DestinationPath
    )
    if (! (Test-Path "${SourceExe}" -PathType Leaf)) {
        throw [System.IO.FileNotFoundException] "Could not find ${SourceExe}"
    }
    # if (! (Test-Path "${DestinationPath}" -PathType Container)) {
    #     throw [System.IO.DirectoryNotFoundException] "Could not find ${DestinationPath}"
    # }
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($DestinationPath)
    $Shortcut.TargetPath = $SourceExe
    $Shortcut.Save()
}

function Request-DomainConnection {
    [OutputType([bool])]
    $IsJoined = (Get-CimInstance -ClassName Win32_ComputerSystem).PartOfDomain
    if ($IsConnected) {
        Show-Output "Your computer seems to be a domain member. Please connect it to the domain network now. A VPN is OK but a physical connection is better."
    }
    return $IsJoined
}

function Show-Information {
    <#
    .SYNOPSIS
        Write-Information with colors. Use this instead of Write-Host.
    #>
    param(
        [Parameter(Position=0, Mandatory=$true)]$MessageData,
        [System.ConsoleColor]$ForegroundColor,
        [System.ConsoleColor]$BackgroundColor
        # This is set globally
        # https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_commonparameters?view=powershell-7.2#-informationaction
        # [string]$InformationAction = "Continue"
    )
    if ($InformationPreference -eq "SilentlyContinue") {
        $CustomInformationPreference = "Continue"
    } else {
        $CustomInformationPreference = $InformationPreference
    }
    Show-Stream @PSBoundParameters -Stream "Write-Information" -InformationAction $CustomInformationPreference
}

function Show-Output {
    <#
    .SYNOPSIS
        Write-Output with colors.
    #>
    param(
        [Parameter(Position=0, Mandatory=$true)]$InputObject,
        [System.ConsoleColor]$ForegroundColor,
        [System.ConsoleColor]$BackgroundColor
    )
    Show-Stream @PSBoundParameters -Stream "Write-Output"
}

function Show-Stream {
     <#
    .SYNOPSIS
        Show text or an object on the console.
    .DESCRIPTION
        With color support but without using Write-Host.
    .LINK
        https://stackoverflow.com/a/4647985/
    #>
    # Argument passing
    # https://stackoverflow.com/a/62861781/
    param(
        [Parameter(Position=0, Mandatory=$true)][Alias("MessageData")]$InputObject,
        [System.ConsoleColor]$ForegroundColor,
        [System.ConsoleColor]$BackgroundColor,
        [string]$Stream = "Write-Output"
    )
    if ($PSBoundParameters.ContainsKey("ForegroundColor")) {
        $OldForegroundColor = [Console]::ForegroundColor
        [Console]::ForegroundColor = $ForegroundColor
    }
    if ($PSBoundParameters.ContainsKey("BackgroundColor")) {
        $OldBackgroundColor = [Console]::BackgroundColor
        [Console]::BackgroundColor = $BackgroundColor
    }
    if ($args) {
        & $Stream $InputObject $args
    } elseif ($Stream -eq "Write-Information") {
        # Write-Information does not support piping
        & $Stream $InputObject
    } else {
        $InputObject | & "${Stream}"
    }
    if ($PSBoundParameters.ContainsKey("ForegroundColor")) {
        [Console]::ForegroundColor = $OldForegroundColor
    }
    if ($PSBoundParameters.ContainsKey("BackgroundColor")) {
        [Console]::BackgroundColor = $OldBackgroundColor
    }
}

function Test-Admin {
    <#
    .SYNOPSIS
        Test whether the script is being run as an administrator
    .LINK
        https://superuser.com/questions/108207/how-to-run-a-powershell-script-as-administrator
    #>
    [OutputType([bool])]
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

Function Test-CommandExists {
    [OutputType([bool])]
    Param(
        [Parameter(Mandatory=$true)][string]$command
    )
    $oldPreference = $ErrorActionPreference
    $ErrorActionPreference = "stop"
    try {
        if (Get-Command $command){RETURN $true}
    }
    catch {RETURN $false}
    finally {$ErrorActionPreference=$oldPreference}
}
