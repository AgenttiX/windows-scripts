﻿<#
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

$host.ui.RawUI.WindowTitle = "Mika's computer installation script"
Write-Host "Starting Mika's computer installation script."

# Global variables
$GlobalHeight = 500;
$GlobalWidth = 700;

$ChocoPrograms = [ordered]@{
    "7-Zip" = "7zip", "File compression utility";
    "AltDrag" = "altdrag", "For moving windows easily";
    "Android Debug Bridge" = "adb", "For developing Android applications";
    "BleachBit" = "bleachbit", "Utility for freeing disk space";
    "CMake" = "cmake", "C/C++ make utility";
    "Discord" = "discord", "Chat and group call platform";
    "DisplayCal" = "displaycal", "Display calibration utility";
    "Epic Games Launcher" = "epicgameslauncher", "Game store";
    "Firefox" = "firefox", "Web browser";
    "GIMP" = "gimp", "Image editor";
    "Git" = "git", "Version control";
    "Inkscape" = "inkscape", "Vector graphics editor";
    "KeePassXC" = "keepassxc", "Password manager";
    "Kingston SSD Manager" = "kingston-ssd-manager", "Management tool and firmware updater for Kingston SSDs";
    "LibreOffice" = "libreoffice-fresh", "Office suite";
    "Logitech Gaming Software" = "logitechgaming", "Driver for old Logitech input devices";
    "Logitech G Hub" = "lghub", "Driver for new Logitech input devices";
    "Microsoft Teams" = "microsoft-teams", "Instant messaging and video conferencing platform";
    "Mumble" = "mumble", "Group call platform";
    "Notepad++" = "notepadplusplus", "Text editor";
    "OBS Studio" = "obs-studio", "Screen capture and broadcasting utility";
    "OpenVPN" = "openvpn", "VPN client";
    "Origin (game store)" = "origin", "Game store";
    "PowerToys" = "powertoys", "Various utilities for Windows";
    "PyCharm Community" = "pycharm-community", "Python IDE";
    "PyCharm Professional" = "pycharm", "Professional Python IDE, requires a license";
    "RescueTime" = "rescuetime", "Time management utility, requires a license)";
    "Rufus" = "rufus", "Creates USB installers for operating systems";
    "Signal" = "signal", "Secure instant messenger";
    "Slack" = "slack", "Instant messenger";
    "SpaceSniffer" = "spacesniffer", "See what's consuming the hard disk space";
    "Speedtest CLI" = "speedtest", "Command-line utility for measuring internet speed";
    "Steam" = "steam", "Game store";
    "Syncthing / SyncTrayzor" = "synctrayzor", "Utility for directly synchronizing files between devices";
    "TeamViewer" = "teamviewer", "Remote control utility, commercial use requires a license";
    "Telegram" = "telegram", "Instant messenger";
    "Thunderbird" = "thunderbird", "Email client";
    "Ubisoft Connect" = "ubisoft-connect", "Game store";
    "VLC" = "vlc", "Video player";
    "Visual Studio Code" = "vscode", "Text editor / IDE";
    "Wacom drivers" = "wacom-drivers", "Drivers for Wacom drawing tablet";
    "Xournal++" = "xournalplusplus", "For taking handwritten notes with a drawing tablet";
}
$WingetPrograms = [ordered]@{
    "PowerShell" = "Microsoft.PowerShell", "Update to the pre-installed PowerShell";
    # The PowerToys version available from WinGet is a preview.
    # https://github.com/microsoft/PowerToys#via-winget-preview
    # "PowerToys" = "Microsoft.PowerToys";
}

# Function definitions should be after the loading of utilities
function CreateList {
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
        Write-Host "Locking columns (this does not work yet)";
        # Write-Host $View.Columns;
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

if (!(Get-AppxPackage -Name "Microsoft.DesktopAppInstaller")) {
    if (Get-AppxPackage -Name "Microsoft.WindowsStore") {
        Write-Host "App Installer apperas not to be installed. Please close this window and install it from the Windows Store. Then restart this script."
        Start-Process -Path "https://www.microsoft.com/en-us/p/app-installer/9nblggh4nns1";
        $confirmation = Write-Host "If you know what you're doing, you may also continue by writing `"force`", but some features may not work.".
        while ($confirmation -ne "f") {
            Write-Host "Close this window or write `"force`" to continue."
        }
    } else {
        Write-Host "Cannot install App Installer, as Microsoft Store appears not to be installed. This is normal on servers. Winget may not be available.";
    }
}

Install-Chocolatey

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
$Layout.RowCount = 4;
$Form.Controls.Add($Layout);

$ChocoProgramsView = CreateTable $Form $Layout "Centrally updated programs (Chocolatey)" $ChocoPrograms;
$WingetProgramsView = CreateTable $Form $Layout "Centrally updated programs (Winget)" $WingetPrograms;
$WingetProgramsView.Height = 50;

# Disable Winget if it's not found
if (!(Test-CommandExists "winget")) {
    $WingetProgramsListBox.Enabled = $false;
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
    $ChocoProgramsView.Height = $Form.Height - 200;
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
    Write-Host "Installing programs with Chocolatey."
    choco upgrade -y $ChocoSelected
} else {
    Write-Host "No programs were selected to be installed with Chocolatey."
}

$WingetSelected = GetSelectedCommands $WingetProgramsView
if ($WingetSelected.Count) {
    Write-Host "Installing programs with Winget."
    foreach($program in $WingetSelected) {
        winget install "${program}"
    }
} else {
    Write-Host "No programs were selected to be installed with Winget."
}

Write-Host -ForegroundColor Green "The installation script is ready. You can close this window now."
