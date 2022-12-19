<#
.SYNOPSIS
    Install all the software and configuration usually needed for a computer.
.PARAMETER Elevated
    This parameter is for internal use to check whether an UAC prompt has already been attempted.
#>

param(
    [switch]$Elevated
)

. ".\utils.ps1"

# This should be as early as possible to avoid loading the function definitions etc. twice.
Elevate($myinvocation.MyCommand.Definition)

Start-Transcript -Path "${LogPath}\install-computer_$(Get-Date -Format "yyyy-MM-dd_HH-mm").txt"

$host.ui.RawUI.WindowTitle = "Mika's computer installation script"
Show-Output "Starting Mika's computer installation script."
Request-DomainConnection
Show-Output "The graphical user interface (GUI) is a very preliminary version and will be improved in the future."
Show-Output "If it doesn't fit on your monitor, please reduce the display scaling at:"
Show-Output "`"Settings -> System -> Display -> Scale and layout -> Change the size of text, apps and other items`""

# Global variables
$GlobalHeight = 800;
$GlobalWidth = 700;

# TODO: hide non-work-related apps on domain computers
$ChocoPrograms = [ordered]@{
    "7-Zip" = "7zip", "File compression utility";
    "Adobe Acrobat Reader DC" = "adobereader", "PDF reader. Not usually needed, as web browsers have good integrated pdf readers.";
    "AltDrag" = "altdrag", "For moving windows easily";
    "Anaconda 3 (NOTE!)" = "anaconda3", "NOTE! Comes with lots of libraries. Use Miniconda or regular Python instead, unless you absolutely need this. Installation with Chocolatey does not work with PyCharm without custom symlinks.";
    "Android Debug Bridge" = "adb", "For developing Android applications";
    "BleachBit" = "bleachbit", "Utility for freeing disk space";
    "BOINC" = "boinc", "Distributed computing platform";
    "Chocolatey GUI (RECOMMENDED!)" = "chocolateygui", "Graphical user interface for managing packages installed with this script and for installing additional software.";
    "CMake" = "cmake", "C/C++ make utility";
    "Discord" = "discord", "Chat and group call platform";
    "DisplayCal" = "displaycal", "Display calibration utility";
    "Docker Desktop (NOTE!)" = "docker-desktop", "Container platform. NOTE! Windows Subsystem for Linux 2 (WSL 2) has to be installed before installing this";
    "Epic Games Launcher" = "epicgameslauncher", "Game store";
    "Firefox" = "firefox", "Web browser";
    "GIMP" = "gimp", "Image editor";
    # "Git" = "git", "Version control";
    "Inkscape" = "inkscape", "Vector graphics editor";
    "KeePassXC" = "keepassxc", "Password manager";
    "Kingston SSD Manager" = "kingston-ssd-manager", "Management tool and firmware updater for Kingston SSDs";
    "LibreOffice" = "libreoffice-fresh", "Office suite";
    "Logitech Gaming Software" = "logitechgaming", "Driver for old Logitech input devices";
    "Logitech G Hub" = "lghub", "Driver for new Logitech input devices";
    "Microsoft Teams" = "microsoft-teams", "Instant messaging and video conferencing platform";
    "MiKTeX" = "miktex", "LaTeX environment";
    "Miniconda 3 (NOTE!)" = "miniconda3", "Anaconda package manager and Python 3 without the pre-installed libraries. NOTE! Installation with Chocolatey does not work with PyCharm without custom symlinks.";
    "Mumble" = "mumble", "Group call platform";
    "Notepad++" = "notepadplusplus", "Text editor";
    "OBS Studio" = "obs-studio", "Screen capture and broadcasting utility";
    # The OpenVPN version available from the Chocolatey repositories is old
    # and does not support all the config values in the the config packages created by pfSense
    # "OpenVPN" = "openvpn", "VPN client";
    "Origin (game store)" = "origin", "Game store";
    "PDF-XChange Editor" = "pdfxchangeeditor", "PDF editor";
    "PowerToys" = "powertoys", "Various utilities for Windows";
    "PyCharm Community" = "pycharm-community", "Python IDE";
    "PyCharm Professional" = "pycharm", "Professional Python IDE, requires a license";
    "Python (NOTE!)" = "python", "NOTE! This will be updated automatically in the future, which may break installed libraries and virtualenvs.";
    "RescueTime" = "rescuetime", "Time management utility, requires a license";
    "Rufus" = "rufus", "Creates USB installers for operating systems";
    "Signal" = "signal", "Secure instant messenger";
    "Slack" = "slack", "Instant messenger";
    "SpaceSniffer" = "spacesniffer", "See what's consuming the hard disk space";
    "Speedtest CLI" = "speedtest", "Command-line utility for measuring internet speed";
    "Steam" = "steam", "Game store";
    "Syncthing / SyncTrayzor" = "synctrayzor", "Utility for directly synchronizing files between devices";
    "TeamViewer" = "teamviewer", "Remote control utility, commercial use requires a license";
    "Telegram" = "telegram", "Instant messenger";
    "Texmaker" = "texmaker", "LaTeX editor";
    "Thunderbird" = "thunderbird", "Email client";
    "Ubisoft Connect" = "ubisoft-connect", "Game store";
    "VirtualBox (NOTE!)" = "virtualbox", "Virtualization platform. NOTE! Cannot be installed on the same computer as Hyper-V. Hardware virtualization should be enabled in BIOS/UEFI before installing.";
    "VirtualBox Guest Additions (NOTE!)" = "virtualbox-guest-additions-guest.install", "NOTE! This should be installed inside a virtual machine.";
    "VLC" = "vlc", "Video player";
    "Visual Studio Code" = "vscode", "Text editor / IDE";
    "Wacom drivers" = "wacom-drivers", "Drivers for Wacom drawing tablets";
    "Xournal++" = "xournalplusplus", "For taking handwritten notes with a drawing tablet";
    "Zoom" = "zoom", "Video conferencing";
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
    Show-Output "Downloading IDS Peak"
    $Folder = "ids-peak-win-${Version}"
    $Filename = "${Folder}.zip"
    Invoke-WebRequest -uri "https://en.ids-imaging.com/files/downloads/ids-peak/software/windows/${Filename}" -OutFile "${Downloads}\${Filename}"
    Show-Output "Extracting IDS Peak"
    Expand-Archive -Path "${Downloads}\${Filename}" -DestinationPath "${Downloads}\${Folder}"
    Show-Output "Installing IDS Peak"
    Start-Process -NoNewWindow -Wait "${Downloads}\${Folder}\ids_peak_${Version}.exe"
}

function Install-IDSSoftwareSuite ([string]$Version = "4.95.2", [string]$Version2 = "49520") {
    Show-Output "Downloading IDS Software Suite (µEye)"
    $Folder = "ids-software-suite-win-${Version}"
    $Filename = "${Folder}.zip"
    Invoke-WebRequest -Uri "https://en.ids-imaging.com/files/downloads/ids-software-suite/software/windows/${Filename}" -OutFile "${Downloads}\${Filename}"
    Show-Output "Extracting IDS Software Suite (µEye)"
    Expand-Archive -Path "${Downloads}\${Filename}" -DestinationPath "${Downloads}\${Folder}"
    Show-Output "Installing IDS Software Suite (µEye)"
    Start-Process -NoNewWindow -Wait "${Downloads}\${Folder}\uEye_${Version2}.exe"
}

function Install-NI4882 ([string]$Version = "22.8") {
    <#
    .SYNOPSIS
        Install National Instruments NI-488.2
    .LINK
        https://www.ni.com/fi-fi/support/downloads/drivers/download.ni-488-2.html#467646
    #>
    Show-Output "Downloading NI 488.2 (GPIB) drivers"
    $Filename = "ni-488.2_${Version}_online.exe"
    Invoke-WebRequest -Uri "https://download.ni.com/support/nipkg/products/ni-4/ni-488.2/${Version}/online/${Filename}" -OutFile "${Downloads}\${Filename}"
    Show-Output "Installing NI 488.2 (GPIB) drivers"
    Start-Process -NoNewWindow -Wait "${Downloads}\${Filename}"
}

function Install-OpenVPN ([string]$Version = "2.5.8", [string]$Version2 = "I603") {
    <#
    .SYNOPSIS
        Install OpenVPN Community
    .LINK
        https://openvpn.net/community-downloads/
    #>
    Show-Output "Downloading OpenVPN"
    $Arch = Get-InstallBitness -x86 "x86" -x86_64 "amd64"
    $Filename = "OpenVPN-${Version}-${Version2}-${Arch}.msi"
    Invoke-WebRequest -Uri "https://swupdate.openvpn.org/community/releases/$Filename" -OutFile "${Downloads}\${Filename}"
    Start-Process -NoNewWindow -Wait "msiexec" -ArgumentList "/i","${Downloads}\${Filename}"
}

function Install-OriginLab {
    <#
        Install the full version of OriginLab
    #>
    [OutputType([bool])]
    param(
        [string]$Version = "Origin2022bSr1No_H"
    )
    Show-Output "Searching for OriginLab from the network drive."
    $FilePath = "V:\Software\Origin\${Version}\setup.exe"
    if (-not (Test-Path "$FilePath")) {
        Show-Output "The OriginLab installer was not found. Do you have the network drive mounted?"
        Show-Output "It could be that your computer does not have the necessary group policies applied. Applying. You will need to reboot for the changes to become effective."
        gpupdate /force
        return $false
    }
    Show-Output "OriginLab installer found. Installing."
    Start-Process -NoNewWindow -Wait "$FilePath"
    return $true
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
    Invoke-WebRequest -Uri "https://www.originlab.com/ftp/${Filename}" -OutFile "${Downloads}\${Filename}"
    $DestinationPath = "${Home}\Downloads\Origin Viewer"
    if(-Not (Clear-Path "${DestinationPath}")) {
        return $false
    }
    Show-Output "Extracting Origin Viewer"
    Expand-Archive -Path "${Downloads}\$Filename" -DestinationPath "${DestinationPath}"
    $ExePath = Find-First -Filter "*.exe" -Path "${DestinationPath}"
    if ($ExePath -eq $null) {
        Show-Information "No exe file was found in the extracted directory." -ForegroundColor Red
        return $false
    }
    New-Shortcut -SourceExe "${ExePath}" -DestinationPath "${env:APPDATA}\Microsoft\Windows\Start Menu\Programs\Origin Viewer.lnk"
    return $true
}

function Install-Rezonator1([string]$Version = "1.7.116.375") {
    <#
    .SYNOPSIS
        Install the reZonator laser cavity simulator
    .LINK
        http://rezonator.orion-project.org/?page=dload
    #>
    Show-Output "Downloading reZonator 1"
    $Filename = "rezonator-${Version}.exe"
    Invoke-WebRequest -Uri "http://rezonator.orion-project.org/files/${Filename}" -OutFile "${Downloads}\${Filename}"
    Show-Output "Installing reZonator 1"
    Start-Process -NoNewWindow -Wait "${Downloads}\${Filename}"
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
        [string]$Version = "2.0.10-beta6"
    )
    Show-Output "Downloading reZonator 2"
    $Bitness = Get-InstallBitness -x86 "x32" -x86_64 "x64"
    $Filename = "rezonator-${Version}-win-${Bitness}.zip"
    Invoke-WebRequest -Uri "https://github.com/orion-project/rezonator2/releases/download/${Version}/${Filename}" -OutFile "${Downloads}\${Filename}"
    $DestinationPath = "${Home}\Downloads\reZonator"
    if(-Not (Clear-Path "${DestinationPath}")) {
        return $false
    }
    Show-Output "Extracting reZonator 2"
    Expand-Archive -Path "${Downloads}\$Filename" -DestinationPath "${DestinationPath}"
    New-Shortcut -SourceExe "${DestinationPath}\rezonator.exe" -DestinationPath "${env:APPDATA}\Microsoft\Windows\Start Menu\Programs\reZonator 2.lnk"
    return $true
}

function Install-SNLO ([string]$Version = "78") {
    Show-Output "Downloading SNLO"
    $Filename = "SNLO-v${Version}.exe"
    Invoke-WebRequest -Uri "https://as-photonics.com/snlo_files/${Filename}" -OutFile "${Downloads}\${Filename}"
    Show-Output "Installing SNLO"
    Start-Process -NoNewWindow -Wait "${Downloads}\${Filename}"
}

function Install-StarLab {
    <#
    .SYNOPSIS
        Install Ophir StarLab
    .LINK
        https://www.ophiropt.com/laser--measurement/software/starlab-for-usb
    #>
    Show-Output "Downloading Ophir StarLab"
    $Filename="StarLab_Setup.exe"
    Invoke-WebRequest -Uri "https://www.ophiropt.com/laser/register_files/${Filename}" -OutFile "${Downloads}\${Filename}"
    Show-Output "Installing Ophir StarLab"
    Start-Process -NoNewWindow -Wait "${Downloads}\${Filename}"
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
    Invoke-WebRequest -Uri "https://www.thorlabs.com/software/THO/ThorCam/ThorCam_V${Version}/${FilenameRemote}" -OutFile "${Downloads}\${FilenameLocal}" # -UserAgent $UserAgent
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
    Invoke-WebRequest -Uri "https://www.thorlabs.com/software/MUC/Beam/Software/Beam_${version}/${filename}" -OutFile "${Downloads}\${Filename}"
    Show-Output "Installing Thorlabs Beam"
    Expand-Archive -Path "${Downloads}\${Filename}" -DestinationPath "${Downloads}\${Folder}"
    $Process = Start-Process -NoNewWindow -Wait -PassThru "${Downloads}\${Folder}\Thorlabs Beam Setup.exe"
    if ($Process.ExitCode -ne 0) {
        Show-Output -ForegroundColor Red "Thorlabs Beam installation seems to have failed. Probably the server detected that this is a script, resulting in a corrupted download. Please download ThorCam manually from the Thorlabs website."
    }
}

function Install-ThorlabsKinesis ([string]$Version = "1.14.36", [string]$Version2 = "20973") {
    <#
    .SYNOPSIS
        Install Thorlabs Kinesis
    .LINK
        https://www.thorlabs.com/software_pages/ViewSoftwarePage.cfm?Code=Motion_Control&viewtab=0
    #>
    Show-Output "Downloading Thorlabs Kinesis. The web server has strict bot detection and the download may therefore fail, producing an invalid file."
    $Arch = Get-InstallBitness -x86 "x86" -x86_64 "x64"
    $Filename = "kinesis_${Version2}_setup_${Arch}.exe"
    Invoke-WebRequest -Uri "https://www.thorlabs.com/Software/Motion%20Control/KINESIS/Application/v${Version}/${Filename}" -OutFile "${Downloads}\${Filename}"
    Show-Output "Installing Thorlabs Kinesis"
    $Process = Start-Process -NoNewWindow -Wait -PassThru "${Downloads}\${Filename}"
    if ($Process.ExitCode -ne 0) {
        Show-Output -ForegroundColor Red "Thorlabs Kinesis installation seems to have failed. Probably the server detected that this is a script, resulting in a corrupted download. Please download ThorCam manually from the Thorlabs website."
    }
}

function Install-VeecoVision {
    [OutputType([int])]
    param()

    Show-Output "Searching for Veeco (Wyko) Vision from the network drive."
    $FilePath = "V:\Software\Veeco\VISION64_V5.51_Release.zip"
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
    $FilePath = "V:\Software\Veeco\Vision64 5.51 Update 3.zip"
    if (Test-Path "$FilePath") {
        Expand-Archive -Path "$FilePath" -DestinationPath "$Downloads"
        Show-Output "Installing Veeco (Wyko) Vision update"
        Start-Process -NoNewWindow -Wait "${Downloads}\Vision64 5.51 Update 3\CD\Vision64_5.51_Update_3.EXE"
    } else {
        Show-Output "Veeco (Wyko) Vision update was not found. Has the file been moved?"
        return 2
    }
    return 0
}
function Install-WSL {
    if (Test-CommandExists "wsl") {
        Show-Output "Installing Windows Subsystem for Linux (WSL), version >= 2"
        wsl --install
    } else {
        Show-Output "The installer command for Windows Subsystem for Linux (WSL) was not found. Are you running an old version of Windows?"
    }
}

$OtherOperations = [ordered]@{
    "Geekbench" = ${function:Install-Geekbench}, "Performance testing utility, versions 2-5. Commercial use requires a license.";
    "Git" = ${function:Install-Git}, "Git with custom arguments (SSH available from PATH etc.)";
    "IDS Peak" = ${function:Install-IDSPeak}, "Driver for IDS cameras and old Thorlabs cameras";
    "IDS Software Suite (µEye, NOTE!)" = ${function:Install-IDSSoftwareSuite}, "Driver for old IDS/Thorlabs cameras. NOTE! Use IDS Peak instad if your cameras are compatible with it.";
    "NI 488.2 (GPIB)" = ${function:Install-NI4882}, "National Instruments GPIB drivers";
    "OpenVPN" = ${function:Install-OpenVPN}, "VPN client";
    "Ophir StarLab" = ${function:Install-StarLab}, "Driver for Ophir power meters";
    "OriginLab" = ${function:Install-OriginLab}, "OriginLab data graphing and analysis software";
    "Origin Viewer" = ${function:Install-OriginViewer}, "Viewer for OriginLab data graphing and analysis files";
    "Phoronix Test Suite" = ${function:Install-PTS}, "Performance testing framework";
    "Rezonator 1" = ${function:Install-Rezonator1}, "Simulator for optical cavities (old stable version)";
    "Rezonator 2" = ${function:Install-Rezonator2}, "Simulator for optical cavities (new beta version)";
    "SNLO" = ${function:Install-SNLO}, "Crystal nonlinear optics simulator";
    "Thorlabs ThorCam (NOTE!)" = ${function:Install-ThorCam}, "Driver for Thorlabs cameras. NOTE! Use IDS Peak instead for old cameras.";
    "Thorlabs Beam" = ${function:Install-ThorlabsBeam}, "Driver for Thorlabs beam profilers and M2 measurement systems";
    "Thorlabs Kinesis" = ${function:Install-ThorlabsKinesis}, "Driver for Thorlabs motors and stages";
    "Veeco (Wyko) Vision" = ${function:Install-VeecoVision}, "Data analysis tool for Veeco/Wyko profilers";
    "Windows Subsystem for Linux (WSL, NOTE!)" = ${function:Install-WSL}, "Compatibility layer for running Linux applications on Windows, version >= 2. Hardware virtualization should be enabled in BIOS/UEFI before installing.";
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
        Enable-WindowsOptionalFeature -Feature "$feature" -Online
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

Show-Output -ForegroundColor Green "The installation script is ready. You can now close this window."
Stop-Transcript
