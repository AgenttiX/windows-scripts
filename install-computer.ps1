<#
.SYNOPSIS
    Install all the software and configuration usually needed for a computer.
#>

param(
    [switch]$Elevated
)

. ".\utils.ps1"

# This should be as early as possible to avoid loading the function definitions etc. twice.
# Elevate($myinvocation.MyCommand.Definition)

$host.ui.RawUI.WindowTitle = "Mika's computer installation script"
Write-Host "Starting Mika's computer installation script."

# Global variables
$GlobalHeight = 500;
$GlobalWidth = 300;

$ChocolateyPrograms = [ordered]@{
    "7-Zip" = "7zip";
    "AltDrag" = "altdrag";
    "Android Debug Bridge" = "adb";
    "Firefox" = "firefox";
    "LibreOffice" = "libreoffice";
    "OBS Studio" = "obs-studio";
    "OpenVPN" = "openvpn";
    "Steam" = "steam";
    "VLC" = "vlc";
}
$WingetPrograms = [ordered]@{
    "PowerShell 7 (aka. PowerShell Core)" = "Microsoft.PowerShell";
    # The PowerToys version available from WinGet is a preview.
    # https://github.com/microsoft/PowerToys#via-winget-preview
    # "PowerToys" = "Microsoft.PowerToys";
}

# Function definitions should be after the loading of utilities
function CreateList {
    param(
        [Parameter(mandatory=$true)][System.Object]$Parent,
        [Parameter(mandatory=$true)][String]$Title,
        [Parameter(mandatory=$true)][String[]]$Options,
        [int]$Width = $GlobalWidth
    )
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
        Write-Host "Cannot install App Installer, as Microsoft Store appears not to be installed. This is normal on servers.";
    }
}

Install-Chocolatey

# Import Windows Forms Assembly
Add-Type -AssemblyName System.Windows.Forms;

# Create a Form
$Form = New-Object -TypeName System.Windows.Forms.Form;
$Form.Text = "Mika's installer script"
$Form.Width = $GlobalWidth + 30;
$Form.Height = $GlobalHeight + 30;

$Table = New-Object System.Windows.Forms.TableLayoutPanel
$Table.RowCount = 4;
$Table.Width = $GlobalWidth + 10;
$Table.Height = $GlobalHeight;
$Form.Controls.Add($Table);

$ChocolateyProgramsListBox = CreateList $Table "Centrally updated programs (Chocolatey)" $ChocolateyPrograms.keys;
$WingetProgramsListBox = CreateList $Table "Centrally updated programs (Winget)" $WingetPrograms.keys;
if (!(Test-CommandExists "winget")) {
    $WingetProgramsListBox.Enabled = $false;
}

# Show the form
$Form.ShowDialog();
# $Form.Close();


[string[]] $ChocolateyInstalling = @();
foreach($item in $ChocolateyProgramsListBox.CheckedItems){
    $ChocolateyInstalling += $ChocolateyPrograms[$item];
}
Write-Host "Selected programs for Chocolatey: ${ChocolateyInstalling}"
# choco install "$ChocolateyInstalling"

Write-Host "Installing Winget programs"
foreach($item in $WingetProgramsListBox.CheckedItems){
    Write-Host "- ${item}";
}
