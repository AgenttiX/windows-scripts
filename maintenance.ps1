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
.PARAMETER Reboot
    Reboot the computer when the script is ready.
.PARAMETER Shutdown
    Shut the computer down when the script is ready.
.PARAMETER Thunderbird
    Remove local Thunderbird IMAP folders etc.
#>

param(
    [switch]$Deep,
    [switch]$Docker,
    [switch]$Elevated,
    [switch]$Firefox,
    [switch]$Reboot,
    [switch]$Shutdown,
    [switch]$Thunderbird
)

if ($Reboot -and $Shutdown) {
    Write-Host "Both reboot and shutdown switches cannot be enabled at the same time."
    Exit 1
}

. "./utils.ps1"
GitPull
Elevate($myinvocation.MyCommand.Definition)

$host.ui.RawUI.WindowTitle = “Mika's maintenance script"

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
# Script starts here
# ---
if ($Reboot) {
    Write-Host "The computer will be rebooted automatically after the script is complete due to a command-line argument."
} elseif ($Shutdown) {
    Write-Host "The computer will be shut down automatically after the script is complete due to a command-line argument."
}
# else {
#     Write-Host "Do you want to reboot or shut down automatically after the script is complete?"
#     Write-Host "Do not enable these if you have large game updates to download, as those may not finish."
#     $reply = Read-Host -Prompt "[r/s/n]?"
#     if ($reply -match "[rR]") {
#         $Reboot = $true
#         Write-Host "Automatic reboot has been enabled."
#     } elseif ($reply -match "[sS]" ) {
#         $Shutdown = $true
#         Write-Host "Automatic shutdown has been enabled."
#     } else {
#         $Reboot = $false
#         $Shutdown = $false
#         Write-Host "Automatic reboot and shutdown are disabled."
#     }
# }

Write-Host "Performing initial steps that have to be performed before Windows Update."
Write-Host "Do not write in the console or press enter unless requested." -ForegroundColor Red
Write-Host "After a moment you may be asked about Windows Updates, and writing in the console now may cause in the selection of updates you don't want."

# Resynchronize time with domain controllers or other NTP server.
# This may be needed for gpupdate if the internal clock is out of sync with the domain.
Write-Host "Synchronizing system clock."
w32tm /resync

if (Test-CommandExists "gpupdate") {
    Write-Host "Updating group policies."
    gpupdate /force
} else {
    Write-Host "Group policy updates are not supported on this system."
}

if (Test-CommandExists "Install-Module") {
    Write-Host "You may now be asked whether to install the NuGet package provider. Please select yes."
    Install-Module PSWindowsUpdate -Force
} else {
    Write-Host "Windows Update PowerShell module could not be installed. Check Windows updates manually."
}
if (Test-CommandExists "Install-WindowsUpdate") {
    Write-Host "You may now be asked whether to install some Windows Updates."
    Write-Host "It's recommended to answer yes EXCEPT for the following:"
    Write-Host "- Microsoft Silverlight"
    Write-Host "- Preview versions"
    Install-WindowsUpdate -MicrosoftUpdate -IgnoreReboot
}

Install-Chocolatey
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

if (Test-CommandExists "cleanmgr") {
    Write-Host "Running Windows disk cleanup."
    # This command is non-blocking
    cleanmgr /verylowdisk
} else {
    # Cleanmgr is not installed on Hyper-V Server
    Write-Host "Windows disk cleanup was not found."
}

# Game updates (non-blocking)
$steam_path="C:\Program Files (x86)\Steam\Steam.exe"
if (Test-Path $steam_path) {
    Write-Host "Starting Steam for updates."
    & $steam_path
}

$epic_games_path="C:\Program Files (x86)\Epic Games\Launcher\Portal\Binaries\Win32\EpicGamesLauncher.exe"
if (Test-Path $epic_games_path) {
    Write-Host "Staring Epic Games Launcher for updates."
    & $epic_games_path
}

# TODO: add Origin

$ubisoft_connect_path="C:\Program Files (x86)\Ubisoft\Ubisoft Game Launcher\UbisoftConnect.exe"
if (Test-Path $ubisoft_connect_path) {
    Write-Host "Starting Ubisoft Connect for updates."
    & $ubisoft_connect_path
}

$riot_client_path="C:\Riot Games\Riot Client\RiotClientServices.exe"
if (Test-Path $riot_client_path) {
    Write-Host "Starting League of Legends for updates."
    & $riot_client_path --launch-product=league_of_legends --launch-patchline=live
}

$minecraft_path="C:\Program Files (x86)\Minecraft Launcher\MinecraftLauncher.exe"
if (Test-Path $minecraft_path) {
    Write-Host "Starting Minecraft for updates."
    & $minecraft_path
}

# Misc tasks

# Lenovo Vantage (non-blocking)
if ($Reboot -or $Shutdown) {
    Write-Host "Lenovo Vantage will not be started, as automatic reboot or shutdown is enabled."
} elseif (Get-AppxPackage -Name "E046963F.LenovoCompanion") {
    Write-Host "Starting Lenovo Vantage for updates."
    start shell:appsFolder\E046963F.LenovoCompanion_k1h2ywk1493x8!App
} else {
    Write-Host "Lenovo Vantage was not found."
}

if (Test-CommandExists "docker") {
    Write-Host "Cleaning Docker"
    if ($Docker) {docker system prune -f -a}
    else {docker system prune -f}
} else {
    Write-Host "Docker was not found."
}

if (Test-CommandExists "Update-Help") {
    Write-Host "Updating PowerShell help"
    Update-Help
} else {
    Write-Host "Help updates are not supported by this PowerShell version."
}

Write-Host "Optimizing drives"
Get-Volume | ForEach-Object {
    if ($_.DriveLetter) {
        Optimize-Volume -DriveLetter $_.DriveLetter -Verbose
    }
}

# Antivirus
if (Test-CommandExists "Update-MpSignature") {
    Write-Host "Updating Windows Defender definitions."
    Update-MpSignature
} else {
    Write-Host "Virus definition updates are not supported. Check them manually."
}
if (Test-CommandExists "Start-MpScan") {
    Write-Host "Running Windows Defender full scan."
    Start-MpScan -ScanType "FullScan"
} else {
    Write-Host "Virus scan is not supported. Run it manually."
}

Write-Host "The maintenance script is ready." -ForegroundColor Green
if ($Reboot) {
    Write-Host "The computer will be rebooted in 10 seconds."
    # The /g switch will automatically login and lock the current user, if this feature is enabled in Windows settings."
    shutdown /g /t 10 /c "Mika's maintenance script is ready. Rebooting."
} elseif ($Shutdown) {
    Write-host "The computer will be shut down in 10 seconds."
    shutdown /s /t 10 /c "Mika's maintenance script is ready. Shutting down."
} else {
    Write-Host "You can close this window now." -ForegroundColor Green
}
