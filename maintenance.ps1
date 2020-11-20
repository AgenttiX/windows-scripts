<#
.SYNOPSIS
    Maintenance script for Windows-based computers
.PARAMETER Deep
    Run a deep cleanup. This will take longer.
.PARAMETER Docker
    Perform a deep clean of Docker. This will clean all unused images.
.PARAMETER Elevated
    This parameter is for internal use to check whether an UAC prompt has already been attempted
.PARAMETER Firefox
    Remove Firefox cookies etc.
.PARAMETER Thunderbird
    Remove local Thunderbird IMAP folders etc.
#>

param(
    [switch]$Deep,
    [switch]$Docker,
    [switch]$Elevated,
    [switch]$Firefox,
    [switch]$Thunderbird
)

# Configuration to ensure that the script is run as an administrator
# https://superuser.com/questions/108207/how-to-run-a-powershell-script-as-administrator
function Test-Admin {
  $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
  $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

if ((Test-Admin) -eq $false)  {
    if ($elevated)
    {
        # tried to elevate, did not work, aborting
    }
    else {
        Write-Host "This script requires admin access. Elevating."
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    }
exit
}

# ---
# Constants
# ---

# The list of cleaners can be obtained with the parameter --list-cleaners
$bleachbit_features = @(
    "adobe_reader.*",
    # "amsn.*",
    "amule.*",
    # "apt.*",
    # "audacious.*",
    # "bash.*",
    # "beagle.*",
    "chromium.*",
    # "chromium.cache",
    # "chromium.cookies",
    # "chromium.current_session",
    # "chromium.dom",
    # "chromium.form_history",
    # "chromium.history",
    # "chromium.passwords",
    # "chromium.search_engines",
    # "chromium.vacuum"
    # "d4x.*",
    # "deepscan.backup",
    # "deepscan.ds_store",
    # "deepscan.thumbs_db",
    # "deepscan.tmp",
    # "easytag.*",
    # "elinks.*",
    # "emesene.*",
    # "epiphany.*",
    # "evolution.*",
    # "exaile.*",
    "filezilla.*",
    # "firefox.*",
    "firefox.backup",
    "firefox.cache",
    # "firefox.cookies",
    "firefox.crash_reports",
    "firefox.dom",
    # "firefox.download_history",
    "firefox.forms",
    # "firefox.passwords",
    "firefox.session_restore",
    "firefox.site_preferences",
    "firefox.url_history",
    "firefox.vacuum",
    "flash.*",
    # "gedit.*",
    # "gftp.*",
    "gimp.*",
    # "gl-117.*",
    # "gnome.*",
    "google_chrome.*",
    "google_earth.*",
    "google_toolbar.*",
    "gpodder.*",
    # "gwenview.*",
    "hexchat.*",
    "hippo_opensim_viewer.*",
    "internet_explorer.*",
    "java.*",
    # "kde.*",
    # "konqueror.*",
    "libreoffice.*",
    # "liferea.*",
    # "links2.*",
    "microsoft_office.*",
    "midnightcommander.*",
    "miro.*",
    # "nautilus.*",
    # "nexuiz.*",
    "octave.*",
    "openofficeorg.*",
    "opera.*",
    "paint.*",
    "pidgin.*",
    "realplayer.*",
    # "recoll.*",
    # "rhythmbox.*",
    "safari.*",
    "screenlets.*",
    "seamonkey.*",
    "secondlife_viewer.*",
    "silverlight.*",
    "skype.*",
    "smartftp.*",
    # "sqlite3.*",
    # "system.cache",
    "system.clipboard",
    # "system.custom",
    # "system.desktop_entry",
    # "system.free_disk_space",
    # "system.localizations",
    "system.logs",
    # "system.memory",
    "system.memory_dump",
    "system.muicache",
    "system.prefetch",
    # "system.recent_documents",
    "system.recycle_bin",
    # "system.rotated_logs",
    "system.tmp",
    # "system.trash",
    "system.updates",
    "teamviewer.*",
    # "thumbnails.*",
    # "thunderbird.cache",
    "thunderbird.cookies",
    # "thunderbird.index",
    # "thunderbird.passwords",
    "thunderbird.vacuum",
    "tortoisesvn.*",
    # "transmission.blocklists",
    # "transmission.history",
    # "transmission.torrents",
    # "tremulous.*",
    "vim.*",
    "vlc.*",
    "vuze.*",
    "warzone2100.*",
    "waterfox.*",
    "winamp.*",
    "windows_defender.*",
    # "wine.*",
    # "winetricks.*",
    "winrar.*",
    "winzip.*",
    "wordpad.*",
    # "x11.*",
    # "xine.*",
    "yahoo_messenger.*"
    # "yum.*"
)

$bleachbit_features_deep = @(
    "deepscan.backup",
    "deepscan.ds_store",
    "deepscan.thumbs_db",
    "deepscan.tmp"
)

$bleachbit_features_firefox = @(
    "firefox.*"
)

$bleachbit_features_thunderbird = @(
    "thunderbird.cache",
    "thunderbird.index"
)


# ---
# Functions
# ---

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

# ---
# Script starts here
# ---

# Resynchronize time with domain controllers or other NTP server.
# This may be needed for gpupdate if the internal clock is out of sync with the domain.
w32tm /resync

# Update group policies
gpupdate /force

if (Test-CommandExists "Install-Module") {
    Install-Module PSWindowsUpdate -Force
} else {
    Write-Host "Windows Update PowerShell module could not be installed. Check Windows updates manually."
}
if (Test-CommandExists "Install-WindowsUpdate") {
    Install-WindowsUpdate -MicrosoftUpdate -IgnoreReboot
}

if (-Not (Test-CommandExists "choco")) {
    Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
}
if (Test-CommandExists "choco") {
    choco upgrade all -y
}

# BleachBit
$bleachbit_path_native = "C:\Program Files\BleachBit\bleachbit_console.exe"
$bleachbit_path_x86 = "C:\Program Files (x86)\BleachBit\bleachbit_console.exe"

if ((-not ((Test-Path $bleachbit_path_native) -or (Test-Path $bleachbit_path_x86))) -and (Test-CommandExists "choco")) {
    choco install bleachbit -y
}
if ((Test-Path $bleachbit_path_native) -or (Test-Path $bleachbit_path_x86)) {
    $bleachbit_cleaners = $bleachbit_features
    if ($Deep) {$bleachbit_cleaners += $bleachbit_features_deep}
    if ($Firefox) {$bleachbit_cleaners += $bleachbit_features_firefox}
    if ($Thunderbird) {$bleachbit_cleaners += $bleachbit_features_thunderbird}
    Write-Host "Running Bleachbit with the following cleaners:"
    Write-Host $bleachbit_cleaners
    if (Test-Path $bleachbit_path_native) {
        & $bleachbit_path_native --clean $bleachbit_cleaners
    } else {
        & $bleachbit_path_x86 --clean $bleachbit_cleaners
    }
} else {
    Write-Host BleachBit could not be installed
}

Write-Host "Running Windows disk cleanup"
# This command is non-blocking
cleanmgr /verylowdisk

if (Test-CommandExists "docker") {
    Write-Host "Cleaning Docker"
    if ($Docker) {docker system prune -f -a}
    else {docker system prune -f}
} else {
    Write-Host "Docker was not found"
}

if (Test-CommandExists "Update-Help") {
    Write-Host "Updating PowerShell help"
    Update-Help
} else {
    Write-Host "Help updates are not supported by this PowerShell version"
}

Write-Host "Optimizing drives"
Get-Volume | ForEach-Object {
    if ($_.DriveLetter) {
        Optimize-Volume -DriveLetter $_.DriveLetter -Verbose
    }
}

# Antivirus
if (Test-CommandExists "Update-MpSignature") {
    Write-Host "Updating Windows Defender definitions"
    Update-MpSignature
} else {
    Write-Host "Virus definition updates are not supported. Check them manually."
}
if (Test-CommandExists "Start-MpScan") {
    Write-Host "Running Windows Defender full scan"
    Start-MpScan -ScanType "FullScan"
} else {
    Write-Host "Virus scan is not supported. Run it manually."
}

Write-Host "The maintenance script is ready. You can close this window now."
