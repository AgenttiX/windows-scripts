<#
.SYNOPSIS
    Install all the software and configuration usually needed for a computer.
.PARAMETER Elevated
    This parameter is for internal use to check whether an UAC prompt has already been attempted.
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSReviewUnusedParameter", "Elevated", Justification="Used in utils")]
param(
    [switch]$Elevated
)

. "${PSScriptRoot}\Utils.ps1"

if ($RepoInUserDir) {
    Update-Repo
}
# This should be as early as possible to avoid loading the function definitions etc. twice.
Elevate($myinvocation.MyCommand.Definition)

Start-Transcript -Path "${LogPath}\install-software_$(Get-Date -Format "yyyy-MM-dd_HH-mm").txt"
if (! $RepoInUserDir) {
    Update-Repo
}

# Startup info
$host.ui.RawUI.WindowTitle = "Mika's computer installation script"
Show-Output -ForegroundColor Cyan "Starting Mika's computer installation script."
Show-Output -ForegroundColor Cyan "If some installer requests a reboot, select no, and only reboot the computer when the installation script is ready."
Request-DomainConnection
Show-Output -ForegroundColor Cyan "The graphical user interface (GUI) is a very preliminary version and will be improved in the future."
Show-Output -ForegroundColor Cyan "If it doesn't fit on your monitor, please reduce the display scaling at:"
Show-Output -ForegroundColor Cyan "`"Settings -> System -> Display -> Scale and layout -> Change the size of text, apps and other items`""

# Startup tasks
Add-ScriptShortcuts
Set-RepoPermissions
Test-PendingRebootAndExit

# Global variables
$GlobalHeight = 800;
$GlobalWidth = 700;
$SoftwareRepoPath = "V:\IT\Software"

# TODO: hide non-work-related apps on domain computers
$ChocoPrograms = [ordered]@{
    "7-Zip" = "7zip", "File compression utility";
    "ActivityWatch" = "activitywatch", "Time management utility";
    "Adobe Acrobat Reader DC" = "adobereader", "PDF reader. Not usually needed, as web browsers have good integrated pdf readers.";
    "AltSnap" = "altsnap", "For moving windows easily";
    "Anaconda 3 (NOTE!)" = "anaconda3", "NOTE! Comes with lots of libraries. Use Miniconda or regular Python instead, unless you absolutely need this. Installation with Chocolatey does not work with PyCharm without custom symlinks.";
    "Android Debug Bridge" = "adb", "For developing Android applications";
    "BleachBit" = "bleachbit", "Utility for freeing disk space";
    "BOINC" = "boinc", "Distributed computing platform";
    "Chocolatey GUI (RECOMMENDED!)" = "chocolateygui", "Graphical user interface for managing packages installed with this script and for installing additional software.";
    "CMake" = "cmake", "C/C++ make utility";
    "Discord" = "discord", "Chat and group call platform";
    "DisplayCal" = "displaycal", "Display calibration utility";
    "Docker Desktop (NOTE!)" = "docker-desktop", "Container platform. NOTE! Windows Subsystem for Linux 2 (WSL 2) has to be installed before installing this";
    "EA App" = "ea-app", "Game store";
    "eDrawings Viewer" = "edrawings-viewer", "2D & 3D CAD viewer";
    "Epic Games Launcher" = "epicgameslauncher", "Game store";
    "Firefox" = "firefox", "Web browser";
    "GIMP" = "gimp", "Image editor";
    # "Git" = "git", "Version control";
    "GNU Octave" = "octave", "Free MATLAB alternative";
    "ImageJ" = "imagej", "Image processing and analysis software";
    "Inkscape" = "inkscape", "Vector graphics editor";
    "Intel Driver & Support Assistant" = "intel-dsa", "Intel driver updater";
    "Jellyfin" = "jellyfin-media-player", "Jellyfin client for playing media from a self-hosted server";
    "KeePassXC" = "keepassxc", "Password manager";
    "KLayout" = "klayout", "Lithography mask editor";
    "Kingston SSD Manager" = "kingston-ssd-manager", "Management tool and firmware updater for Kingston SSDs";
    "LibreOffice" = "libreoffice-fresh", "Office suite";
    "Logitech Gaming Software" = "logitechgaming", "Driver for old Logitech input devices";
    "Logitech G Hub" = "lghub", "Driver for new Logitech input devices";
    "Mattermost" = "mattermost-desktop", "Messaging app for teams and organizations (open source Slack alternative)";
    "Microsoft Teams" = "microsoft-teams", "Instant messaging and video conferencing platform";
    "MiKTeX" = "miktex", "LaTeX environment";
    "Miniconda 3 (NOTE!)" = "miniconda3", "Anaconda package manager and Python 3 without the pre-installed libraries. NOTE! Installation with Chocolatey does not work with PyCharm without custom symlinks.";
    "Mumble" = "mumble", "Group call platform";
    "Notepad++" = "notepadplusplus", "Text editor";
    "NVM" = "nvm", "Node.js Version Manager";
    "Obsidian" = "obsidian", "A note-taking app";
    "OBS Studio" = "obs-studio", "Screen capture and broadcasting utility";
    "OpenVPN" = "openvpn", "VPN client";
    "PDFsam" = "pdfsam", "PDF Split & Merge utility";
    "PDF-XChange Editor" = "pdfxchangeeditor", "PDF editor";
    "PDF Arranger" = "pdfarranger", "PDF split & merge utility";
    "pgAdmin" = "pgadmin4", "Graphical user interface (GUI) for managing PostgreSQL";
    "Plex" = "plex", "Plex client for playing media from a self-hosted server";
    "Plexamp" = "plexamp", "Plex client for playing music from a self-hosted server";
    "PostgreSQL" = "postgresql", "Database for e.g. web app development";
    "PowerToys" = "powertoys", "Various utilities for Windows";
    "PyCharm Community" = "pycharm-community", "Python IDE";
    "PyCharm Professional" = "pycharm", "Professional Python IDE, requires a license";
    "Python (NOTE!)" = "python", "NOTE! This will be updated automatically in the future, which may break installed libraries and virtualenvs.";
    "qBittorrent" = "qbittorrent", "Torrent client";
    "RescueTime" = "rescuetime", "Time management utility, requires a license";
    "Rufus" = "rufus", "Creates USB installers for operating systems";
    "Samsung Magician" = "samsung-magician", "Samsung SSD management software";
    "Signal" = "signal", "Secure instant messenger";
    "Slack" = "slack", "Messaging app for teams and organizations";
    "SpaceSniffer" = "spacesniffer", "See what's consuming the hard disk space";
    "Speedtest CLI" = "speedtest", "Command-line utility for measuring internet speed";
    "Spotify" = "spotify", "Music streaming service client";
    "Steam" = "steam", "Game store";
    "Syncthing / SyncTrayzor" = "synctrayzor", "Utility for directly synchronizing files between devices";
    "TeamViewer" = "teamviewer", "Remote control utility, commercial use requires a license";
    "Telegram" = "telegram", "Instant messenger";
    "Texmaker" = "texmaker", "LaTeX editor";
    "Thunderbird" = "thunderbird", "Email client";
    "Tidal" = "tidal", "Music streaming service client";
    "TightVNC" = "tightvnc", "VNC server. Can be used with noVNC by using the configuration from Mika's GitHub.";
    "Ubisoft Connect" = "ubisoft-connect", "Game store";
    "VirtualBox (NOTE!)" = "virtualbox", "Virtualization platform. NOTE! Cannot be installed on the same computer as Hyper-V. Hardware virtualization should be enabled in BIOS/UEFI before installing.";
    "VirtualBox Guest Additions (NOTE!)" = "virtualbox-guest-additions-guest.install", "NOTE! This should be installed inside a virtual machine.";
    "VLC" = "vlc", "Video player";
    "Visual Studio Code (NOTE!)" = "vscode", "Text editor / IDE. Sends lots of tracking data to Microsoft. Use VSCodium instead.";
    "VSCodium" = "vscodium", "Text editor / IDE. Open source version of Visual Studio Code.";
    "Wacom drivers" = "wacom-drivers", "Drivers for Wacom drawing tablets";
    "Xournal++" = "xournalplusplus", "For taking handwritten notes with a drawing tablet";
    "Xerox Global Print Driver PCL/PS V3" = "xeroxupd", "Driver for Xerox printers";
    "yt-dlp" = "yt-dlp", "Video downloader for e.g. YouTube";
    "YubiKey Manager" = "yubikey-manager", "Management software for the YubiKey hardware security keys";
    "Zoom" = "zoom", "Video conferencing";
    "Zotero" = "zotero", "Reference and citation management software";
}
$WingetPrograms = [ordered]@{
    "PowerShell" = "Microsoft.PowerShell", "The new cross-platform PowerShell (>= 7)";
    # The PowerToys version available from WinGet is a preview.
    # https://github.com/microsoft/PowerToys#via-winget-preview
    # "PowerToys" = "Microsoft.PowerToys";
}
$WindowsCapabilities = [ordered]@{
    # "OpenSSH client" = "OpenSSH.Client~~~~0.0.1.0", "NOTE! This is an old version that does not support FIDO2. Install SSH from the other programs menu instead.";
    "RSAT AD LDS (Active Directory management tools)" = "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0", "Active Directory management tools";
    "RSAT BitLocker tools" = "Rsat.BitLocker.Recovery.Tools~~~~0.0.1.0", "Active Directory BitLocker management tools";
    "RSAT DHCP tools" = "Rsat.DHCP.Tools~~~~0.0.1.0", "Active Directory DHCP management tools";
    "RSAT DNS tools" = "Rsat.Dns.Tools~~~~0.0.1.0", "Active Directory DNS management tools";
    "RSAT cluster tools" = "Rsat.FailoverCluster.Management.Tools~~~~0.0.1.0", "Server cluster management tools";
    "RSAT file server tools" = "Rsat.FileServices.Tools~~~~0.0.1.0", "Active Directory file server management tools";
    "RSAT group policy tools" = "Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0", "Active Directory group policy management tools";
    # "RSAT network controller tools" = "Rsat.NetworkController.Tools~~~~0.0.1.0", "RSAT network controller tools";
    "RSAT Server Manager" = "Rsat.ServerManager.Tools~~~~0.0.1.0", "Remote server management tools";
    "RSAT Shielded VM tools" = "Rsat.Shielded.VM.Tools~~~~0.0.1.0", "Management tools for shielded virtual machines";
    "RSAT WSUS tools" = "Rsat.WSUS.Tools~~~~0.0.1.0", "Active Directory Windows Update management tools";
    "SNMP client" = "SNMP.Client~~~~0.0.1.0", "SNMP remote monitoring client";
}
$WindowsFeatures = [ordered]@{
    "Hyper-V (NOTE!)" = "Microsoft-Hyper-V-All", "Virtualization platform. NOTE! Cannot be installed on the same computer as VirtualBox. Hardware virtualization should be enabled in BIOS/UEFI before installing.";
    "Hyper-V management tools" = "Microsoft-Hyper-V-Tools-All", "Tools for managing Hyper-V servers, both local and remote";
}

# Installer functions

function Install-BaslerPylon([string]$Version = "7_4_0_14900") {
    $Filename = "basler_pylon_${Version}.exe"
    Install-FromUri -Name "Basler Pylon Camera Software Suite" -Uri "https://www2.baslerweb.com/media/downloads/software/pylon_software/${Filename}" -Filename "${Filename}"
}

function Install-CorelDRAW {
    Install-FromUri -Name "CorelDRAW" -Uri "https://www.corel.com/akdlm/6763/downloads/free/trials/GraphicsSuite/22H1/JL83s3fG/CDGS.exe" -Filename "CDGS.exe"
}

function Install-DigiSign([string]$Version = "4.3.0(8707)") {
    $Filename = "DigiSignClient_for_dvv_${Version}.exe"
    Install-FromUri -Name "Fujitsu mPollux DigiSign" -Uri "https://dvv.fi/documents/16079645/216375523/${Filename}" -Filename "${Filename}"
}

function Install-Eduroam {
    Install-FromUri -Name "Eduroam" -Uri "https://dl.eduroam.app/windows/x86_64/geteduroam.exe" -Filename "geteduroam.exe"
}

function Install-Git {
    <#
    .SYNOPSIS
        Install Git with Chocolatey and custom parameters
    .LINK
        https://github.com/chocolatey-community/chocolatey-packages/blob/master/automatic/git.install/ARGUMENTS.md
    #>
    # if (Get-WindowsCapability -Online -Name "OpenSSH.Client~~~~0.0.1.0") {
    #     Show-Output "Found Windows integrated OpenSSH client, which the SSH included in Git for Windows should override. Removing."
    Remove-WindowsCapability -Online -Name "OpenSSH.Client~~~~0.0.1.0"
    # }
    Show-Output "Installing Git with Chocolatey and custom parameters"
    choco upgrade git.install -y --force --params "/GitAndUnixToolsOnPath /WindowsTerminalProfile"
}

function Install-IDSPeak ([string]$Version = "2.3.0.0") {
    $Folder = "ids-peak-win-${Version}"
    $Filename = "${Folder}.zip"
    Install-FromUri -Name "IDS Peak" -Uri "https://en.ids-imaging.com/files/downloads/ids-peak/software/windows/${Filename}" -Filename "${Filename}" -UnzipFolderName "${Folder}" -UnzippedFilePath "ids_peak_${Version}.exe"
}

function Install-IDSSoftwareSuite ([string]$Version = "4.95.2", [string]$Version2 = "49520") {
    $Folder = "ids-software-suite-win-${Version}"
    $Filename = "${Folder}.zip"
    Install-FromUri -Name "IDS Software Suite (µEye)" -Uri "https://en.ids-imaging.com/files/downloads/ids-software-suite/software/windows/${Filename}" -Filename "${Filename}" -UnzipFolderName "${Folder}" -UnzippedFilePath "uEye_${Version2}.exe"
}

function Install-LabVIEWRuntime ([string]$Version = "24.1") {
    $Filename ="ni-labview-2024-runtime-engine_${Version}_online.exe"
    Install-FromUri -Name "LabVIEW Runtime" -Uri "https://download.ni.com/support/nipkg/products/ni-l/ni-labview-2024-runtime-engine/${Version}/online/${Filename}" -Filename "${Filename}"
}

function Install-LabVIEWRuntime2014SP1 {
    <#
    .SYNOPSIS
        Install LabVIEW Runtime 2014 SP1 32-bit
    .LINK
        https://www.ni.com/en/support/downloads/software-products/download.labview-runtime.html#306243
    .DESCRIPTION
        Required for SSMbe
    #>
    $Folder = "LVRTE2014SP1_f11Patchstd"
    $Filename = "${Folder}.zip"
    Install-FromUri -Name "LabVIEW Runtime 2014 SP1 32-bit" -Uri "https://download.ni.com/support/softlib/labview/labview_runtime/2014%20SP1/Windows/f11/${Filename}" -Filename "${Filename}" -UnzipFolderName "${Folder}" -UnzippedFilePath "setup.exe" -SHA256 "2c54ab5169dd0cc9f14a7b0057881207b6f76e065c7c78bb8c898ac9c5ca0831"
}

function Install-MeerstetterTEC {
    <#
    .SYNOPSIS
        Install Meerstetter TEC Software
    .LINK
        https://www.meerstetter.ch/customer-center/downloads/category/31-latest-software
    #>
    # Spaces are not allowed in msiexec filenames. Please also see this issue:
    # https://stackoverflow.com/questions/10108517/what-can-cause-msiexec-error-1619-this-installation-package-could-not-be-opened
    Install-FromUri -Name "Meerstetter TEC software" -Uri "https://www.meerstetter.ch/customer-center/downloads/category/31-latest-software?download=331:tec-family-tec-controllers-software" -Filename "TEC_Software.msi"
}

function Install-NI4882 ([string]$Version = "23.5") {
    <#
    .SYNOPSIS
        Install National Instruments NI-488.2
    .LINK
        https://www.ni.com/fi-fi/support/downloads/drivers/download.ni-488-2.html#467646
    #>
    $Filename = "ni-488.2_${Version}_online.exe"
    Install-FromUri -Name "NI 488.2 (GPIB) drivers" -Uri "https://download.ni.com/support/nipkg/products/ni-4/ni-488.2/${Version}/online/${Filename}" -Filename "${Filename}"
}

function Install-NI-VISA1401Runtime {
    <#
    .SYNOPSIS
        Install NI-VISA 14.0.1 Runtime
    .LINK
        https://www.ni.com/en/support/downloads/drivers/download.ni-visa.html#306102
    .DESCRIPTION
        Required for SSMbe.
        LabVIEW (runtime) should be installed first,
        but this is automatically the case when installing both at the same time with this script,
        since LabVIEW comes alphabetically first.
    #>
    $Folder = "NIVISA1401runtime"
    $Filename = "${Folder}.zip"
    Install-FromUri -Name "NI-VISA 14.0.1 Runtime" -Uri "https://download.ni.com/support/softlib/visa/VISA%20Run-Time%20Engine/14.0.1/${Filename}" -Filename "${Filename}" -UnzipFolderName "${Folder}" -UnzippedFilePath "setup.exe" -SHA256 "960e0f68ab7dbff286ba8ac2286d4e883f02f9390c2a033b433005e29fb93e72"
}

function Install-OpenVPN {
    <#
    .SYNOPSIS
        Install OpenVPN Community
    .LINK
        https://openvpn.net/community-downloads/
    #>
    [OutputType([int])]
    param(
        [string]$Version = "2.5.8",
        [string]$Version2 = "I604"
    )
    $Arch = Get-InstallBitness -x86 "x86" -x86_64 "amd64"
    $Filename = "OpenVPN-${Version}-${Version2}-${Arch}.msi"
    Install-FromUri -Name "OpenVPN" -Uri "https://swupdate.openvpn.org/community/releases/$Filename" -Filename "${Filename}"
}

function Install-OriginLab {
    <#
        Install the full version of OriginLab
    #>
    [OutputType([int])]
    param(
        [string]$Version = "Origin2022bSr1No_H"
    )
    return Install-Executable -Name "OriginLab" -Path "${SoftwareRepoPath}\Origin\${Version}\setup.exe"
}

function Install-OriginViewer {
    <#
    .SYNOPSIS
        Install Origin Viewer, the free viewer for Origin data visualization and analysis files
    .LINK
        https://www.originlab.com/viewer/
    #>
    [OutputType([bool])]
    param()
    Show-Output "Downloading Origin Viewer"
    $Arch = Get-InstallBitness -x86 "" -x86_64 "_64"
    $Filename = "OriginViewer${Arch}.zip"
    Invoke-WebRequestFast -Uri "https://www.originlab.com/ftp/${Filename}" -OutFile "${Downloads}\${Filename}"
    $DestinationPath = "${Home}\Downloads\Origin Viewer"
    if(-Not (Clear-Path "${DestinationPath}")) {
        return $false
    }
    Show-Output "Extracting Origin Viewer"
    Expand-Archive -Path "${Downloads}\$Filename" -DestinationPath "${DestinationPath}"
    $ExePath = Find-First -Filter "*.exe" -Path "${DestinationPath}"
    if ($null -eq $ExePath) {
        Show-Information "No exe file was found in the extracted directory." -ForegroundColor Red
        return $false
    }
    New-Shortcut -Path "${env:APPDATA}\Microsoft\Windows\Start Menu\Programs\Origin Viewer.lnk" -TargetPath "${ExePath}"
    return $true
}

function Install-Rezonator1([string]$Version = "1.7.116.375") {
    <#
    .SYNOPSIS
        Install the reZonator laser cavity simulator
    .LINK
        http://rezonator.orion-project.org/?page=dload
    #>
    $Filename = "rezonator-${Version}.exe"
    Install-FromUri -Name "reZonator 1" -Uri "http://rezonator.orion-project.org/files/${Filename}" -Filename "${Filename}"
}

function Install-Rezonator2 {
    <#
    .SYNOPSIS
        Install the reZonator 2 laser cavity simulator
    .LINK
        http://rezonator.orion-project.org/?page=dload
    #>
    [OutputType([bool])]
    param(
        [string]$Version = "2.0.13-beta9"
    )
    Show-Output "Downloading reZonator 2"
    $Bitness = Get-InstallBitness -x86 "x32" -x86_64 "x64"
    $Filename = "rezonator-${Version}-win-${Bitness}.zip"
    Invoke-WebRequestFast -Uri "https://github.com/orion-project/rezonator2/releases/download/v${Version}/${Filename}" -OutFile "${Downloads}\${Filename}"
    $DestinationPath = "${Home}\Downloads\reZonator"
    if(-Not (Clear-Path "${DestinationPath}")) {
        return $false
    }
    Show-Output "Extracting reZonator 2"
    Expand-Archive -Path "${Downloads}\$Filename" -DestinationPath "${DestinationPath}"
    New-Shortcut -Path "${env:APPDATA}\Microsoft\Windows\Start Menu\Programs\reZonator 2.lnk" -TargetPath "${DestinationPath}\rezonator.exe"
    return $true
}

function Install-SNLO ([string]$Version = "78") {
    $Filename = "SNLO-v${Version}.exe"
    Install-FromUri -Name "SNLO" -Uri "https://as-photonics.com/snlo_files/${Filename}" -Filename "${Filename}"
}

function Install-SSMbe ([string]$Version = "20160525") {
    $FolderName = "SSMbe_exe_20160525"
    $FilePath = "${SoftwareRepoPath}\SS10-1 MBE software\${FolderName}.zip"
    $DestinationPath = "${Downloads}\SS10-1 MBE software"
    if (-Not (Clear-Path "${DestinationPath}")) {
        return $false
    }
    if (Test-Path $FilePath) {
        Expand-Archive -Path "${FilePath}" -DestinationPath "${DestinationPath}"
    } else {
        Show-Output -ForegroundColor Red "SSMbe installer was not found."
    }
    New-Shortcut -Path "${env:APPDATA}\Microsoft\Windows\Start Menu\Programs\SSMbe.lnk" -TargetPath "${DestinationPath}\${FolderName}\SSMbe.exe"
    return $true
}

function Install-StarLab {
    <#
    .SYNOPSIS
        Install Ophir StarLab
    .LINK
        https://www.ophiropt.com/laser--measurement/software/starlab-for-usb
    #>
    $Filename="StarLab.zip"
    Install-FromUri -Name "Ophir StarLab" -Uri "https://www.ophiropt.com/mam/celum/celum_assets/op/resources/${Filename}" -Filename "${Filename}" -UnzipFolderName "StarLab" -UnzippedFilePath "StarLab_Setup.exe"
}

function Install-ThorCam ([string]$Version = "3.7.0.6") {
    <#
    .SYNOPSIS
        Install ThorLabs ThorCam
    .LINK
        https://www.thorlabs.com/software_pages/ViewSoftwarePage.cfm?Code=ThorCam
    #>
    Show-Output "Downloading Thorlabs ThorCam. The web server has strict bot detection and the download may therefore fail, producing an invalid file."
    $Arch = Get-InstallBitness -x86 "x86" -x86_64 "x64"
    $FilenameRemote = "Thorlabs%20Scientific%20Imaging%20Software%20${Arch}.exe"
    $FilenameLocal = "Thorlabs Scientific Imaging Software ${Arch}.exe"
    # The "FireFox" typo is by Microsoft itself.
    # $UserAgent = [Microsoft.PowerShell.Commands.PSUserAgent]::FireFox
    Invoke-WebRequestFast -Uri "https://www.thorlabs.com/software/THO/ThorCam/ThorCam_V${Version}/${FilenameRemote}" -OutFile "${Downloads}\${FilenameLocal}" # -UserAgent $UserAgent
    Show-Output "https://www.thorlabs.com/software/THO/ThorCam/ThorCam_V${Version}/${FilenameRemote}"
    Show-Output "Installing Thorlabs ThorCam"
    $Process = Start-Process -NoNewWindow -Wait -PassThru "${Downloads}\${FilenameLocal}"
    if ($Process.ExitCode -ne 0) {
        Show-Output -ForegroundColor Red "ThorCam installation seems to have failed. Probably the server detected that this is a script, resulting in a corrupted download. Please download ThorCam manually from the Thorlabs website."
    }
}

function Install-ThorlabsBeam ([string]$Version = "8.2.5232.395") {
    <#
    .SYNOPSIS
        Install Thorlabs Beam
    .LINK
        https://www.thorlabs.com/software_pages/ViewSoftwarePage.cfm?Code=Beam
    #>
    Show-Output "Downloading Thorlabs Beam. The web server has strict bot detection and the download may therefore fail, producing an invalid file."
    $Folder = "Thorlabs_Beam_${Version}"
    $Filename = "Thorlabs_Beam_${Version}.zip"
    Invoke-WebRequestFast -Uri "https://www.thorlabs.com/software/MUC/Beam/Software/Beam_${version}/${filename}" -OutFile "${Downloads}\${Filename}"
    Show-Output "Installing Thorlabs Beam"
    Expand-Archive -Path "${Downloads}\${Filename}" -DestinationPath "${Downloads}\${Folder}"
    $Process = Start-Process -NoNewWindow -Wait -PassThru "${Downloads}\${Folder}\Thorlabs Beam Setup.exe"
    if ($Process.ExitCode -ne 0) {
        Show-Output -ForegroundColor Red "Thorlabs Beam installation seems to have failed. Probably the server detected that this is a script, resulting in a corrupted download. Please download ThorCam manually from the Thorlabs website."
    }
}

function Install-ThorlabsKinesis {
    <#
    .SYNOPSIS
        Install Thorlabs Kinesis
    .LINK
        https://www.thorlabs.com/software_pages/ViewSoftwarePage.cfm?Code=Motion_Control&viewtab=0
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "", Justification="Thorlabs Kinesis is a singular noun")]
    param(
        [string]$Version = "1.14.36",
        [string]$Version2 = "20973"
    )
    Show-Output "Downloading Thorlabs Kinesis. The web server has strict bot detection and the download may therefore fail, producing an invalid file."
    $Arch = Get-InstallBitness -x86 "x86" -x86_64 "x64"
    $Filename = "kinesis_${Version2}_setup_${Arch}.exe"
    Invoke-WebRequestFast -Uri "https://www.thorlabs.com/Software/Motion%20Control/KINESIS/Application/v${Version}/${Filename}" -OutFile "${Downloads}\${Filename}"
    Show-Output "Installing Thorlabs Kinesis"
    $Process = Start-Process -NoNewWindow -Wait -PassThru "${Downloads}\${Filename}"
    if ($Process.ExitCode -ne 0) {
        Show-Output -ForegroundColor Red "Thorlabs Kinesis installation seems to have failed. Probably the server detected that this is a script, resulting in a corrupted download. Please download ThorCam manually from the Thorlabs website."
    }
}

function Install-VCU {
    [OutputType([int])]
    param(
        [string]$Version = "0.13.40"
    )
    $FilePath = "${SoftwareRepoPath}\VCU\VCU_GUI_Setup_${Version}.exe"
    if (Test-Path "$FilePath") {
        Show-Output "Installing VCU GUI"
        Start-Process -NoNewWindow -Wait "$FilePath"
    } else {
        Show-Output "The VCU GUI installer was not found. Is the network drive mounted?"
        return 1
    }
    return 0
}

function Install-VeecoVision {
    [OutputType([int])]
    param()

    Show-Output "Searching for Veeco (Wyko) Vision from the network drive."
    $FilePath = "${SoftwareRepoPath}\Veeco\VISION64_V5.51_Release.zip"
    if (Test-Path "$FilePath") {
        Expand-Archive -Path "$FilePath" -DestinationPath "$Downloads"
        Show-Output "Installing Veeco (Wyko) Vision"
        Start-Process -NoNewWindow -Wait "${Downloads}\CD 775-425 SOFTWARE VISION64 V5.51\Install.exe"
    } else {
        Show-Output "Veeco (Wyko) Vision was not found. Is the network drive mounted?"
        Show-Output "It could be that your computer does not have the necessary group policies applied. Applying. You will need to reboot for the changes to become effective."
        gpupdate /force
        return 1
    }
    Show-Output "Searching for Veeco (Wyko) Vision update from the network drive."
    $FilePath = "${SoftwareRepoPath}\Veeco\Vision64 5.51 Update 3.zip"
    if (Test-Path "$FilePath") {
        Expand-Archive -Path "$FilePath" -DestinationPath "$Downloads"
        Show-Output "Installing Veeco (Wyko) Vision update"
        Start-Process -NoNewWindow -Wait "${Downloads}\Vision64 5.51 Update 3\CD\Vision64_5.51_Update_3.EXE"
    } else {
        Show-Output -ForegroundColor Red "Veeco (Wyko) Vision update was not found. Has the file been moved?"
        return 2
    }
    return 0
}

function Install-Wavesquared ([string]$Version = "4.4.2.25284") {
    # https://jrsoftware.org/ishelp/index.php?topic=setupcmdline
    Start-Process -NoNewWindow -Wait "${SoftwareRepoPath}\Imagine Optic\wavesquared_${Version}\WaveSuite_setup.exe"
}

function Install-WithSecure {
    $FilePath = "${SoftwareRepoPath}\WithSecure\ElementsAgentInstaller*.exe"
    if (Test-Path "${FilePath}") {
        Show-Output "Installing WithSecure Elements Agent"
        Start-Process -NoNewWindow -Wait "${FilePath}"
    } else {
        Show-Output -ForegroundColor Red "WithSecure Elements Agent installer was not found. Has the file been moved?"
        return 1
    }
    return 0
}

function Install-WSL {
    if (Test-CommandExists "wsl") {
        Show-Output "Installing Windows Subsystem for Linux (WSL), version >= 2"
        wsl --install
    } else {
        Show-Output -ForegroundColor Red "The installer command for Windows Subsystem for Linux (WSL) was not found. Are you running an old version of Windows?"
    }
}

function Install-Xeneth {
    [OutputType([int])]
    param()
    $Bitness = Get-InstallBitness -x86 "" -x86_64 "64"
    $FilePath = Resolve-Path "${SoftwareRepoPath}\Xenics\BOBCAT*\Software\Xeneth-Setup${Bitness}.exe"
    if (Test-Path "${FilePath}") {
        return Install-Executable -Name "Xeneth" -Path "${FilePath}"
    } else {
        Show-Output -ForegroundColor Red "Xeneth installer was not found. Has the file been moved?"
    }
}

$OtherOperations = [ordered]@{
    "Basler Pylon" = ${function:Install-BaslerPylon}, "Driver for Basler cameras";
    "CorelDRAW" = ${function:Install-CorelDRAW}, "Graphic design, illustration and technical drawing software. Requires a license.";
    "Eduroam" = ${function:Install-Eduroam}, "University Wi-Fi";
    "Fujitsu mPollux DigiSign" = ${function:Install-DigiSign}, "Card reader software for Finnish identity cards";
    "Geekbench" = ${function:Install-Geekbench}, "Performance testing utility, versions 2-5. Commercial use requires a license.";
    "Git" = ${function:Install-Git}, "Git with custom arguments (SSH available from PATH etc.)";
    "IDS Peak" = ${function:Install-IDSPeak}, "Driver for IDS cameras and old Thorlabs cameras";
    "IDS Software Suite (µEye, NOTE!)" = ${function:Install-IDSSoftwareSuite}, "Driver for old IDS/Thorlabs cameras. NOTE! Use IDS Peak instad if your cameras are compatible with it.";
    # "LabVIEW Runtime" = ${function:Install-LabVIEWRuntime}, "Required for running LabVIEW-based applications";
    "LabVIEW Runtime 2014 SP1 32-bit" = ${function:Install-LabVIEWRuntime2014SP1}, "Required for SSMbe (it requires this specific older version instead of the latest)";
    "Meerstetter TEC Software" = ${function:Install-MeerstetterTEC}, "Driver for Meerstetter TEC controllers";
    "NI 488.2 (GPIB)" = ${function:Install-NI4882}, "National Instruments GPIB drivers. Includes NI-VISA.";
    "NI-VISA 14.0.1 Runtime" = ${function:Install-NI-VISA1401Runtime}, "Required for SSMbe (it requires this specific older version instead of the latest)";
    # OpenVPN is also available from Chocolatey.
    # Use this manual version only when the package version in Chocolatey is too old.
    # "OpenVPN" = ${function:Install-OpenVPN}, "VPN client";
    "Ophir StarLab" = ${function:Install-StarLab}, "Driver for Ophir power meters";
    "OriginLab" = ${function:Install-OriginLab}, "OriginLab data graphing and analysis software";
    "Origin Viewer" = ${function:Install-OriginViewer}, "Viewer for OriginLab data graphing and analysis files";
    "Phoronix Test Suite" = ${function:Install-PTS}, "Performance testing framework";
    "reZonator 1" = ${function:Install-Rezonator1}, "Simulator for optical cavities (old stable version)";
    "reZonator 2" = ${function:Install-Rezonator2}, "Simulator for optical cavities (new beta version)";
    "SNLO" = ${function:Install-SNLO}, "Crystal nonlinear optics simulator";
    "SSMbe (NOTE!)" = ${function:Install-SSMbe}, "Control software for the SS10-1 MBE reactor. NOTE! Also install the LabVIEW Runtime and NI-VISA dependencies.";
    "Thorlabs ThorCam (NOTE!)" = ${function:Install-ThorCam}, "Driver for Thorlabs cameras. NOTE! Use IDS Peak instead for old cameras.";
    "Thorlabs Beam" = ${function:Install-ThorlabsBeam}, "Driver for Thorlabs beam profilers and M2 measurement systems";
    "Thorlabs Kinesis" = ${function:Install-ThorlabsKinesis}, "Driver for Thorlabs motors and stages";
    "VCU" = ${function:Install-VCU}, "VCU GUI";
    "Veeco (Wyko) Vision" = ${function:Install-VeecoVision}, "Data analysis tool for Veeco/Wyko profilers";
    "Wavesquared" = ${function:Install-Wavesquared}, "M2 factor analysis software";
    "Windows Subsystem for Linux (WSL, NOTE!)" = ${function:Install-WSL}, "Compatibility layer for running Linux applications on Windows, version >= 2. Hardware virtualization should be enabled in BIOS/UEFI before installing.";
    "WithSecure Elements Agent" = ${function:Install-WithSecure}, "Anti-virus. Requires a license.";
    "Xeneth" = ${function:Install-Xeneth}, "Driver for Xenics cameras";
}

# Function definitions should be after the loading of utilities
function CreateList {
    <#
    .SYNOPSIS
        Create a GUI element for selecting options from a list with checkboxes
    .LINK
        https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.checkedlistbox
     #>
    [OutputType([system.Windows.Forms.CheckedListBox])]
    param(
        [Parameter(mandatory=$true)][System.Object]$Parent,
        [Parameter(mandatory=$true)][String]$Title,
        [Parameter(mandatory=$true)][String[]]$Options,
        [int]$Width = $GlobalWidth
    )
    # Title label
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = $Title;
    $Label.Width = $Width;
    $Parent.Controls.Add($Label);
    # Create a CheckedListBox
    $List = New-Object -TypeName System.Windows.Forms.CheckedListBox;
    $Parent.Controls.Add($List);
    $List.Items.AddRange($Options);
    $List.CheckOnClick = $true;
    $List.Width = $Width;
    $List.Height = $Options.Count * 17 + 18;
    return $List;
}
function CreateTable {
    <#
    .SYNOPSIS
        Create a GUI element for selecting items from a list with checboxes
    .LINK
        https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.datagridview
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSReviewUnusedParameter", "e", Justification="Probably used by library code")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSReviewUnusedParameter", "sender", Justification="Probably used by library code")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSReviewUnusedParameter", "Form", Justification="Reserved for future use")]
    [OutputType([system.Windows.Forms.DataGridView])]
    param(
        [Parameter(mandatory=$true)][System.Object]$Form,
        [Parameter(mandatory=$true)][System.Object]$Parent,
        [Parameter(mandatory=$true)][String]$Title,
        [Parameter(mandatory=$true)]$Data
        # [int]$Width = $GlobalWidth
    )
    # Title label
    $Label = New-Object System.Windows.Forms.Label;
    $Label.Text = $Title;
    # $Label.MinimumSize = New-Object System.Drawing.Size($Width, 0);
    $Label.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $Parent.Controls.Add($Label);
    # Create the DataTable
    $Table = New-Object system.Data.DataTable;
    $Col = New-Object system.Data.DataColumn "Selected", ([bool]);
    $Table.Columns.Add($Col);
    $Col = New-Object system.Data.DataColumn "Name", ([string]);
    $Table.Columns.Add($Col);
    $Col = New-Object system.Data.DataColumn "Description", ([string]);
    $Table.Columns.Add($Col);
    $Col = New-Object system.Data.DataColumn "Command", ([Object]);
    $Table.Columns.Add($Col);
    # Fill the DataTable
    foreach($element in $Data.GetEnumerator()) {
        $row = $Table.NewRow();
        $row.Selected = $false;
        $row.Name = $element.Name;
        $row.Description = $element.Value[1];
        $row.Command = $element.Value[0];
        $Table.Rows.Add($row);
    }
    # Create the DataGridView
    $View = New-Object system.Windows.Forms.DataGridView;
    $View.DataSource = $Table;

    $View.AllowUserToAddRows = $false;
    $View.AllowUserToDeleteRows = $false;
    $View.AllowUserToOrderColumns = $false;
    $View.AllowUserToResizeColumns = $false;
    $View.AllowUserToResizeRows = $false;
    $View.AutoSizeColumnsMode = "AllCells";
    $View.ShowEditingIcon = $false;

    # This enables the desired resizing behaviour, but does not work properly without AutoSizeMode or equivalent.
    # $View.AutoSize = $true;
    # This property does not exist for DataGridView.
    # $View.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink;

    # $View.Height = $Data.Count * 25 + 50;
    # $View.Width = $Width;

    $View.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

    # https://forums.powershell.org/t/datagridview-hide-column/16739
    # https://stackoverflow.com/a/23763025/
    $dataBindingComplete = {
        param (
            [object]$sender,
            [System.EventArgs]$e
        )
        Show-Output "Locking the UI from modifications and hiding unnecessary columns. (This does not work yet.)";
        # Show-Output $View.Columns;
        foreach($column in $View.Columns) {
            if ($column.Name -ne "Selected") {
                $column.ReadOnly = $true;
            }
        }
        # $View.Columns["Command"].Visible = $false;
    }
    # $Form.Add_load($dataBindingComplete);

    $Parent.Controls.Add($View);
    # https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.datagridview.databindingcomplete
    $View.Add_DataBindingComplete($dataBindingComplete);

    return $View;
}
function GetSelectedCommands {
    [OutputType([string[]])]
    param(
        [Parameter(mandatory=$true)][system.Windows.Forms.DataGridView]$View
    )
    $rows = $View.DataSource.Select("Selected = 1")
    $commands = @()
    foreach($row in $rows) {
        $commands += $row.Command;
    }
    return $commands;
}

# Script starts here

Install-Chocolatey
Install-Winget

# Import Windows Forms Assembly
Add-Type -AssemblyName System.Windows.Forms;

# Create the Form
$Form = New-Object -TypeName System.Windows.Forms.Form;
$Form.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
# $Form.AutoSize = $true;
# $Form.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink;
$Form.Text = "Mika's installer script"
$Form.MinimumSize = New-Object System.Drawing.Size($GlobalWidth, $GlobalHeight);

$Layout = New-Object System.Windows.Forms.TableLayoutPanel;
$Layout.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
# $Layout.AutoSize = $true;
# $Layout.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink;
# $Layout.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D;
$Layout.RowCount = 6;
$Form.Controls.Add($Layout);

$ChocoProgramsView = CreateTable -Form $Form -Parent $Layout -Title "Centrally updated programs (Chocolatey)" -Data $ChocoPrograms;
$WingetProgramsView = CreateTable -Form $Form -Parent $Layout -Title "Centrally updated programs (Winget)" -Data $WingetPrograms;
$WingetProgramsView.Height = 50;
$WindowsCapabilitiesView = CreateTable -Form $Form -Parent $Layout -Title "Windows capabilities" -Data $WindowsCapabilities;
$WindowsCapabilitiesView.Height = 150;
$WindowsFeaturesView = CreateTable -Form $Form -Parent $Layout -Title "Windows features" -Data $WindowsFeatures;
$WindowsFeaturesView.Height = 95;
$OtherOperationsView = CreateTable -Form $Form -Parent $Layout -Title "Other programs and operations. These you have to keep updated manually." -Data $OtherOperations
$OtherOperationsView.height = 150;

# Disable unsupported features
if (!(Test-CommandExists "choco")) {
    $ChocoProgramsView.Enabled = "$false";
}
if (!(Test-CommandExists "winget")) {
    $WingetProgramsView.Enabled = $false;
}
if (!(Test-CommandExists "Add-WindowsCapability")) {
    $WindowsCapabilitiesView.Enabled = $false;
}
if (!(Test-CommandExists "Enable-WindowsOptionalFeature")) {
    $WindowsFeaturesView.Enabled = $false;
}

# Add OK button
$Form | Add-Member -MemberType NoteProperty -Name Continue -Value $false;
$OKButton = New-Object System.Windows.Forms.Button;
$OKButton.Text = "OK";
$OKButton.Add_Click({
    $Form.Continue = $true;
    $Form.Close();
})
$Layout.Controls.Add($OKButton);

function ResizeLayout {
    $Layout.Width = $Form.Width - 10;
    $Layout.Height = $Form.Height - 10;
    $ChocoProgramsView.Height = $Form.Height - 660;
}
$Form.Add_Resize($function:ResizeLayout);
ResizeLayout;

# Show the form
$Form.ShowDialog();
if (! $Form.Continue) {
    return 0;
}

$ChocoSelected = GetSelectedCommands $ChocoProgramsView
if ($ChocoSelected.Count) {
    Show-Output "Installing programs with Chocolatey."
    choco upgrade -y $ChocoSelected
} else {
    Show-Output "No programs were selected to be installed with Chocolatey."
}

$WingetSelected = GetSelectedCommands $WingetProgramsView
if ($WingetSelected.Count) {
    Show-Output "Installing programs with Winget. If asked to accept the license of the package repository, please select yes."
    foreach($program in $WingetSelected) {
        winget install "${program}"
    }
} else {
    Show-Output "No programs were selected to be installed with Winget."
}

$WindowsCapabilitiesSelected = GetSelectedCommands $WindowsCapabilitiesView
if ($WindowsCapabilitiesSelected.Count) {
    Show-Output "Installing Windows capabilities."
    foreach($capability in $WindowsCapabilitiesSelected) {
        Show-Output "Installing ${capability}"
        Add-WindowsCapability -Name "$capability" -Online
    }
} else {
    Show-Output "No Windows capabilities were selected to be installed."
}

$WindowsFeaturesSelected = GetSelectedCommands $WindowsFeaturesView
if ($WindowsFeaturesSelected.Count) {
    Show-Output "Installing Windows features."
    foreach($feature in $WindowsFeaturesSelected) {
        Show-Output "Installing ${feature}"
        Enable-WindowsOptionalFeature -FeatureName "$feature" -Online
    }
} else {
    Show-Output "No Windows features were selected to be installed."
}

# These have to be after the package manager -based installations, as the package managers may install some Visual C++ runtimes etc., which we want to update automatically.
$OtherSelected = GetSelectedCommands $OtherOperationsView
if ($OtherSelected.Count) {
    Show-Output "Running other selected operations."
    foreach($command in $OtherSelected) {
        . $command
    }
} else {
    Show-Output "No other operations were selected."
}

if (Test-RebootPending) {
    Show-Output -ForegroundColor Cyan "The computer is pending a reboot. Please reboot the computer."
}
Show-Output -ForegroundColor Green "The installation script is ready. You can now close this window."
Stop-Transcript
