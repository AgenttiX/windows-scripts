<#
.SYNOPSIS
    Zero free space on the disk.
.DESCRIPTION
    Indended for VirtualBox guests.
    https://superuser.com/questions/529149/how-to-compact-virtualboxs-vdi-file-size
.PARAMETER DriveLetter
    Name of the drive to be cleaned
#>

param(
    [Parameter(Mandatory=$true)] [string]$DriveLetter
)

. "./utils.ps1"
Elevate($myinvocation.MyCommand.Definition)

if (! (Test-Path "./lib/SDelete")) {
    New-Item -Path "." -Name "lib" -ItemType "directory" -Force

    # https://docs.microsoft.com/en-us/sysinternals/downloads/sdelete
    Invoke-WebRequest -Uri "https://download.sysinternals.com/files/SDelete.zip" -OutFile "./lib/SDelete.zip"

    Expand-Archive -Path "./lib/SDelete.zip" -DestinationPath "./lib/SDelete"
}

if ([Environment]::Is64BitOperatingSystem) {
    Write-Host "64-bit operating system detected. Using 64-bit SDelete."
    ./lib/SDelete/sdelete64.exe $DriveLetter -z
} else {
    Write-Host "32-bit operating system detected. Using 32-bit SDelete."
    ./lib/SDelete/sdelete.exe $DriveLetter -z
}
