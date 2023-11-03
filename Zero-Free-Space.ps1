<#
.SYNOPSIS
    Zero free space on the disk.
.DESCRIPTION
    Indended for VirtualBox guests as a preparation for compacting the disk image.
.LINK
    https://superuser.com/questions/529149/how-to-compact-virtualboxs-vdi-file-size
.PARAMETER DriveLetter
    Name of the drive to be cleaned
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSReviewUnusedParameter", "Elevated", Justification="Used in utils")]
param(
    [Parameter(Mandatory=$true)] [string]$DriveLetter,
    [switch]$Elevated
)

. "${PSScriptRoot}\Utils.ps1"
Elevate($myinvocation.MyCommand.Definition)

if (! (Test-Path "./lib/SDelete")) {
    New-Item -Path "." -Name "lib" -ItemType "directory" -Force

    $SDeleteUri = "https://download.sysinternals.com/files/SDelete.zip"
    $SDeleteZipPath = "./lib/SDelete.zip"

    # https://docs.microsoft.com/en-us/sysinternals/downloads/sdelete
    Invoke-WebRequest -Uri $SDeleteUri -OutFile $SDeleteZipPath
    Expand-Archive -Path $SDeleteZipPath -DestinationPath "./lib/SDelete"
}

if ([Environment]::Is64BitOperatingSystem) {
    Show-Output "64-bit operating system detected. Using 64-bit SDelete."
    ./lib/SDelete/sdelete64.exe "${DriveLetter}:\" -z
} else {
    Show-Output "32-bit operating system detected. Using 32-bit SDelete."
    ./lib/SDelete/sdelete.exe "${DriveLetter}:\" -z
}
