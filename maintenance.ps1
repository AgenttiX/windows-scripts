<#
.SYNOPSIS
    Maintenance script for Windows-based computers
.PARAMETER Clean
    Clean the system by removing old temporary files etc.
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
    [switch]$Clean,
    [switch]$Deep,
    [switch]$Docker,
    [switch]$Elevated,
    [switch]$Firefox,
    [switch]$Reboot,
    [switch]$Shutdown,
    [switch]$Thunderbird,
    [switch]$Zerofree
)

if ($Reboot -and $Shutdown) {
    Show-Output -ForegroundColor Red "Both reboot and shutdown switches cannot be enabled at the same time."
    Exit 1
}

# Load utility functions from another file.
. ".\utils.ps1"

# This git pull may cause the GitHub rate limits to be reached in an enterprise network.
# It may also cause problems if the script is already elevated, as the script file would be modified while executing.
# GitPull

Elevate($myinvocation.MyCommand.Definition)
Start-Transcript -Path "${LogPath}\maintenance_$(Get-Date -Format "yyyy-MM-dd_HH-mm").txt"

$host.ui.RawUI.WindowTitle = "Mika's maintenance script"
Show-Output -ForegroundColor Cyan "Starting Mika's maintenance script."
Request-DomainConnection

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

if ((-not $Zerofree) -and (Test-Path "${env:ProgramFiles}\Oracle\VirtualBox Guest Additions")) {
    Show-Output -ForegroundColor Cyan "This seems to be a VirtualBox machine."
    $Zerofree = Get-YesNo "Do you want to zero free space at the end of this script?"
}
if ($Zerofree) {
    Show-Output -ForegroundColor Cyan "Free space will be zeroed at the end of this script."
}

# If a Lenovo computer does not have Lenovo Vantage installed
if (((Get-ComputerInfo -Property "BiosManufacturer").BiosManufacturer.ToLower() -eq "lenovo") -and (! ((Get-AppxPackage -Name "E046963F.LenovoCompanion") -or (Get-AppxPackage -Name "E046963F.LenovoSettingsforEnterprise")))) {
    Show-Output -ForegroundColor Red "It appears that you have a Lenovo computer but don't have Lenovo Vantage or Lenovo Commercial Vantage installed."
    if (Get-IsDomainJoined) {
        Show-Output -ForegroundColor Red "Your computer appears to be part of a domain. Please install Lenovo Commercial Vantage to get driver and firmware updates."
        Start-Process "https://apps.microsoft.com/store/detail/lenovo-commercial-vantage/9NR5B8GVVM13"
    } else {
        Show-Output -ForegroundColor Red "Please install Lenovo Vantage from Microsoft Store to get driver and firmware updates."
        Start-Process "https://apps.microsoft.com/store/detail/lenovo-vantage/9WZDNCRFJ4MV"
    }
}

if ($Reboot) {
    Show-Output -ForegroundColor Cyan "The computer will be rebooted automatically after the script is complete due to a command-line argument."
} elseif ($Shutdown) {
    Show-Output -ForegroundColor Cyan "The computer will be shut down automatically after the script is complete due to a command-line argument."
}
# else {
#     Show-Output "Do you want to reboot or shut down automatically after the script is complete?"
#     Show-Output "Do not enable these if you have large game updates to download, as those may not finish."
#     $reply = Read-Host -Prompt "[r/s/n]?"
#     if ($reply -match "[rR]") {
#         $Reboot = $true
#         Show-Output "Automatic reboot has been enabled."
#     } elseif ($reply -match "[sS]" ) {
#         $Shutdown = $true
#         Show-Output "Automatic shutdown has been enabled."
#     } else {
#         $Reboot = $false
#         $Shutdown = $false
#         Show-Output "Automatic reboot and shutdown are disabled."
#     }
# }

Show-Output -ForegroundColor Cyan "Performing initial steps that have to be performed before Windows Update."
Show-Output -ForegroundColor Red "Do not write in the console or press enter unless requested."
Show-Output -ForegroundColor Cyan "After a moment you may be asked about Windows Updates, and writing in the console now may cause in the selection of updates you don't want."

# Resynchronize time with domain controllers or other NTP server.
# This may be needed for gpupdate if the internal clock is out of sync with the domain.
Show-Output -ForegroundColor Cyan "Synchronizing system clock. If your computer is part of a domain but not connected to the domain network (e.g. with a VPN), this may fail."
w32tm /resync

if (Test-CommandExists "gpupdate") {
    Show-Output -ForegroundColor Cyan "Updating group policies. If your computer is part of a domain but not connected to the domain network (e.g. with a VPN), this will fail."
    Show-Output -ForegroundColor Cyan "`"Failed to apply`" error messages are also quite common, and may be caused by reasons unrelated to your computer (synchronization problems for domain controller SYSVOL etc.)."
    gpupdate /force
} else {
    Show-Output -ForegroundColor Cyan "Group policy updates are not supported on this system."
}

if (Test-CommandExists "Install-Module") {
    Show-Output -ForegroundColor Cyan "Installing PowerShell bindings for Windows Update. You may now be asked whether to install the NuGet package provider. Please select yes."
    Install-Module PSWindowsUpdate -Force
} else {
    Show-Output -ForegroundColor Red "Windows Update PowerShell module could not be installed. Check Windows updates manually."
}
if (Test-CommandExists "Install-WindowsUpdate") {
    Show-Output -ForegroundColor Cyan "You may now be asked whether to install some Windows Updates."
    Show-Output -ForegroundColor Cyan "It's recommended to answer yes EXCEPT for the following:"
    Show-Output -ForegroundColor Cyan "- Microsoft Silverlight"
    Show-Output -ForegroundColor Cyan "- Preview versions"
    Install-WindowsUpdate -MicrosoftUpdate -IgnoreReboot
} else {
    Show-Output -ForegroundColor Red "Windows Update bindings were not found. You have to check for Windows updates manually."
}

Install-Chocolatey
if (Test-CommandExists "choco") {
    Show-Output -ForegroundColor Cyan "Installing updates with Chocolatey"
    choco upgrade all -y
}

if (Test-CommandExists "winget") {
    Show-Output -ForegroundColor Cyan "Installing updates with Winget. If you are asked to agree to source agreements terms, please select yes."
    winget upgrade --all
}

# BleachBit
$bleachbit_path_native = "${env:ProgramFiles}\BleachBit\bleachbit_console.exe"
$bleachbit_path_x86 = "${env:ProgramFiles(x86)}\BleachBit\bleachbit_console.exe"

if ($Clean -or $Deep) {
    if ((-not ((Test-Path $bleachbit_path_native) -or (Test-Path $bleachbit_path_x86))) -and (Test-CommandExists "choco")) {
        choco install bleachbit -y
    }
    if ((Test-Path $bleachbit_path_native) -or (Test-Path $bleachbit_path_x86)) {
        $bleachbit_cleaners = $bleachbit_features
        if ($Deep) {$bleachbit_cleaners += $bleachbit_features_deep}
        if ($Firefox) {$bleachbit_cleaners += $bleachbit_features_firefox}
        if ($Thunderbird) {$bleachbit_cleaners += $bleachbit_features_thunderbird}
        Show-Output -ForegroundColor Cyan "Running Bleachbit with the following cleaners:"
        Show-Output $bleachbit_cleaners
        if (Test-Path $bleachbit_path_native) {
            & $bleachbit_path_native --clean $bleachbit_cleaners
        } else {
            & $bleachbit_path_x86 --clean $bleachbit_cleaners
        }
    } else {
        Show-Output -ForegroundColor Red "BleachBit could not be installed."
    }
} else {
    Show-Output "Skipping BleachBit, as the parameters -Clean or -Deep have not been given."
}


# Game updates (non-blocking)
# Todo: Create a function for these, which would check for both Program Files (x86) and Program Files, as the former does not exist on 32-bit systems.
# https://stackoverflow.com/a/19015642/

Show-Output -ForegroundColor Cyan "Installing game updates. (If this is a work computer, probably no games will be found.)"

$steam_path="${env:ProgramFiles(x86)}\Steam\Steam.exe"
if (Test-Path $steam_path) {
    Show-Output "Starting Steam for updates."
    & $steam_path
} else {
    Show-Output "Steam was not found."
}

$battle_net_path="${env:ProgramFiles(x86)}\Battle.net\Battle.net Launcher.exe"
if (Test-Path $battle_net_path) {
    Show-Output "Starting Battle.net for updates."
    & $battle_net_path
} else {
    Show-Output "Battle.net was not found."
}

$epic_games_path="${env:ProgramFiles(x86)}\Epic Games\Launcher\Portal\Binaries\Win32\EpicGamesLauncher.exe"
if (Test-Path $epic_games_path) {
    Show-Output "Staring Epic Games Launcher for updates."
    & $epic_games_path
} else {
    Show-Output "Epic Games Launcher was not found."
}

$origin_path="${env:ProgramFiles(x86)}\Origin\Origin.exe"
if (Test-Path $origin_path) {
    Show-Output "Starting Origin for updates."
    & $origin_path
} else {
    Show-Output "Origin was not found."
}

$ubisoft_connect_path="${env:ProgramFiles(x86)}\Ubisoft\Ubisoft Game Launcher\UbisoftConnect.exe"
if (Test-Path $ubisoft_connect_path) {
    Show-Output "Starting Ubisoft Connect for updates."
    & $ubisoft_connect_path
} else {
    Show-Output "Ubisoft Connect was not found."
}

$riot_client_path="C:\Riot Games\Riot Client\RiotClientServices.exe"
if (Test-Path $riot_client_path) {
    Show-Output "Starting Riot Games client for League of Legends updates."
    & $riot_client_path --launch-product=league_of_legends --launch-patchline=live
} else {
    Show-Output "Riot Games client was not found."
}

$minecraft_path="${env:ProgramFiles(x86)}\Minecraft Launcher\MinecraftLauncher.exe"
if (Test-Path $minecraft_path) {
    Show-Output "Starting Minecraft for updates."
    & $minecraft_path
} else {
    Show-Output "Minecraft was not found."
}


# Misc non-blocking tasks

$kingston_ssd_manager_path = "${env:ProgramFiles(x86)}\Kingston_SSD_Manager\KSM.exe"
if ($Reboot -or $Shutdown) {
    Show-Output "Kingston SSD Manager will not be started, as automatic reboot or shutdown is enabled."
} elseif (Test-Path $kingston_ssd_manager_path) {
    Show-Output -ForegroundColor Cyan "Starting Kingston SSD Manager to check for updates. If there are any, plase wait that the maintenance script is ready before installing them to ensure that no other updates will interfere with them."
    & $kingston_ssd_manager_path
} else {
    Show-Output "Kingston SSD Manager was not found."
}

if ($Clean -or $Deep) {
    if (Test-CommandExists "cleanmgr") {
        Show-Output -ForegroundColor Cyan "Running Windows disk cleanup. This will open some windows about `"low disk space condition`". You can close them when they are ready."
        # This command is non-blocking
        cleanmgr /verylowdisk
    } else {
        # Cleanmgr is not installed on Hyper-V Server
        Show-Output "Windows disk cleanup was not found."
    }
} else {
    Show-Output "Skipping Windows disk cleanup, as the parameters -Clean or -Deep has not been specified."
}

# Windows Store app updates (partially blocking)
# May update Lenovo Vantage, and therefore needs to be before it.
Show-Output -ForegroundColor Cyan "Updating Windows Store apps."
Get-CimInstance -Namespace "Root\cimv2\mdm\dmmap" -ClassName "MDM_EnterpriseModernAppManagement_AppManagement01" | Invoke-CimMethod -MethodName UpdateScanMethod


# Misc blocking tasks

$NIPackageManagerUpdaterPath = "${env:ProgramFiles}\National Instruments\NI Package Manager\Updater\Install.exe"
if (Test-Path "$NIPackageManagerUpdaterPath") {
    Show-Output -ForegroundColor Cyan "Updating NI Package Manager. This will open a window where you have to install the update."
    Show-Output -ForegroundColor Cyan "If you get a message saying that the program cannot be upgraded, it means you already have the latest version."
    Show-Output -ForegroundColor Cyan "In this case please close the window so that the script may continue."
    Show-Output -ForegroundColor Cyan "Once this update is ready, the NI Package Manager itself may open. Please close it. This script will then continue automatically and reopen it with the correct options enabled."
    Start-Process -NoNewWindow -Wait "$NIPackageManagerUpdaterPath"
} else {
    Show-Output "NI Package Manager updater was not found."
}
$NIPackageManagerPath = "${env:ProgramFiles}\National Instruments\NI Package Manager\NIPackageManager.exe"
if (Test-Path $NIPackageManagerPath) {
    Show-Output -ForegroundColor Cyan "Running NI Package Manager to check for updates."
    Show-Output -ForegroundColor Cyan "Please install the suggested updates and then close the package manager. This script will then continue automatically."
    Start-Process -NoNewWindow -Wait "$NIPackageManagerPath" -ArgumentList "--initial-view=Updates","--output-transactions","--prevent-reboot"
} else {
    Show-Output "NI Package Manager was not found."
}

if (Test-CommandExists "docker") {
    Show-Output -ForegroundColor Cyan "Cleaning Docker"
    if ($Docker) {docker system prune -f -a}
    else {docker system prune -f}
} else {
    Show-Output "Docker was not found."
}

if (Test-CommandExists "Update-Help") {
    Show-Output -ForegroundColor Cyan "Updating PowerShell help. All modules don't have help info, and therefore this may produce errors, which is OK."
    Update-Help
} else {
    Show-Output "Help updates are not supported by this PowerShell version."
}

Show-Output -ForegroundColor Cyan "Optimizing drives. SSDs will be trimmed and HDDs defragmented."
Show-Output -ForegroundColor Cyan "If some of the connected drives (e.g. USB flash sticks and SD cards) don't support optimization, this will produce errors, which is OK."
Get-Volume | ForEach-Object {
    if ($_.DriveLetter) {
        Optimize-Volume -DriveLetter $_.DriveLetter -Verbose
    }
}

# Antivirus
if (Test-CommandExists "Update-MpSignature") {
    Show-Output -ForegroundColor Cyan "Updating Windows Defender definitions. If you have another antivirus program installed, Windows Defender may be disabled, causing this to fail."
    Update-MpSignature
} else {
    Show-Output -ForegroundColor Red "Virus definition updates are not supported. Check them manually."
}
if (Test-CommandExists "Start-MpScan") {
    Show-Output -ForegroundColor Cyan "Running Windows Defender full scan. If you have another antivirus program installed, Windows Defender may be disabled, causing this to fail."
    Start-MpScan -ScanType "FullScan"
} else {
    Show-Output -ForegroundColor Red "Virus scan is not supported. Run it manually."
}

# Lenovo Vantage (non-blocking)
# This should be the last step in the script so that its updates are not installed during other updates.
if ($Reboot -or $Shutdown) {
    Show-Output -ForegroundColor Cyan "Lenovo Vantage will not be started, as automatic reboot or shutdown is enabled."
} elseif (Get-AppxPackage -Name "E046963F.LenovoSettingsforEnterprise") {
    Show-Output -ForegroundColor Cyan "Starting Lenovo Commercial Vantage for updates. Please select `"Check for system updates`" and install the suggested updates."
    Start-Process shell:appsFolder\E046963F.LenovoSettingsforEnterprise_k1h2ywk1493x8!App
} elseif (Get-AppxPackage -Name "E046963F.LenovoCompanion") {
    Show-Output -ForegroundColor Cyan "Starting Lenovo Vantage for updates. Please select `"Check for system updates`" and install the suggested updates."
    Start-Process shell:appsFolder\E046963F.LenovoCompanion_k1h2ywk1493x8!App
} else {
    Show-Output "Lenovo Vantage was not found."
}

if ($Zerofree) {
    .\zero-free-space.ps1 -DriveLetter "C"
}

Show-Output -ForegroundColor Green "The maintenance script is ready."
if ($Reboot) {
    Show-Output "The computer will be rebooted in 10 seconds."
    # The /g switch will automatically login and lock the current user, if this feature is enabled in Windows settings."
    shutdown /g /t 10 /c "Mika's maintenance script is ready. Rebooting."
} elseif ($Shutdown) {
    Show-Output "The computer will be shut down in 10 seconds."
    shutdown /s /t 10 /c "Mika's maintenance script is ready. Shutting down."
} else {
    Show-Output -ForegroundColor Green "You can now close this window."
}

Stop-Transcript
